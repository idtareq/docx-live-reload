import queue
import shutil
import threading
import argparse
import zipfile
import win32com.client as win32
import os
import sys
import time
from pathlib import Path
import pythoncom
from lxml import etree

if sys.version_info < (3, 8):
    import importlib_metadata
else:
    import importlib.metadata as importlib_metadata


class Paths:
    def __init__(self, path) -> Path:
        self.path = Path(path).absolute()

    @property
    def preview_copy_path(self) -> Path:
        return self.path.with_stem(self.path.stem + "__preview")

    @property
    def ext_dirpath(self) -> Path:
        return self.path.with_name(self.path.name + "__extracted")

    @property
    def ext_xmls(self) -> list[Path]:
        return [
            self.ext_dirpath / "word" / "document.xml",
            self.ext_dirpath / "word" / "styles.xml",
        ]


class FilesWatcher:
    def __init__(self, watch_paths: list[Path]):
        self.watch_paths = watch_paths
        self.modified = {}
        self.stopped = True

    @property
    def changed(self) -> bool:
        if self.stopped:
            return False

        is_changed = False

        for path, modified in self.modified.items():
            self.modified[path] = os.path.getmtime(path)
            if not is_changed and self.modified[path] != modified:
                is_changed = True

        return is_changed

    def update_modified(self):
        for path in self.watch_paths:
            self.modified[path] = os.path.getmtime(path)

    def stop(self):
        self.stopped = True

    def start(self):
        self.update_modified()
        self.stopped = False


class Shared:
    paths: Paths
    docx_watcher: FilesWatcher
    xmls_watcher: FilesWatcher


def watcher_thread(cmds: queue.Queue):
    while True:
        if Shared.docx_watcher.changed:
            print("Change in docx file detected!")
            cmds.put("reload")

        if Shared.xmls_watcher.changed:
            print("Change in extracted xmls detected!")
            cmds.put("update")

        time.sleep(0.1)


def preview_thread(cmds: queue.Queue):
    word_app = win32.Dispatch("Word.Application", pythoncom.CoInitialize())
    word_app.Visible = True

    parser = etree.XMLParser(remove_blank_text=True)
    doc = run_preview(word_app, None, parser)

    Shared.docx_watcher.start()
    Shared.xmls_watcher.start()

    while True:
        try:
            cmd = cmds.get(timeout=1)
        except queue.Empty:
            cmd = None

            try:
                word_app.Visible
            except pythoncom.com_error:
                print()
                print("Word was closed. exiting..")
                exit()

        if cmd == "reload":
            print("reloading..")
            doc = run_preview(word_app, doc, parser)

        elif cmd == "update":
            update(parser)

        elif cmd == "quit":
            print("exiting..")
            try:
                if doc in word_app.Documents:
                    doc.Close()
                word_app.Quit()
            except pythoncom.com_error:
                pass
            break


def input_thread(cmds: queue.Queue):
    try:
        print("Press 'r' to reload. 'q' to quit.")

        while True:
            cmd = input("> ")

            if cmd == "q":
                cmds.put("quit")
                break

            elif cmd == "r":
                cmds.put("reload")

    except KeyboardInterrupt:
        cmds.put("quit")
        pass


def run_preview(word_app, doc, parser):
    if doc is not None and doc in word_app.Documents:
        doc.Close()

    shutil.copyfile(Shared.paths.path, Shared.paths.preview_copy_path)
    doc = word_app.Documents.Open(str(Shared.paths.preview_copy_path))

    Shared.xmls_watcher.stop()
    with zipfile.ZipFile(Shared.paths.path) as file:
        for name in file.namelist():
            file.extract(name, Shared.paths.ext_dirpath)
        for path in Shared.paths.ext_xmls:
            tree = etree.parse(path, parser)
            tree.write(path, pretty_print=True, encoding="utf-8")
    Shared.xmls_watcher.start()

    return doc


def update(parser):
    try:
        for path in Shared.paths.ext_xmls:
            etree.parse(path, parser)
    except etree.XMLSyntaxError:
        pass

    print("updating..")

    with zipfile.ZipFile(Shared.paths.path, "w") as file:
        for path in Shared.paths.ext_dirpath.glob("**/*"):
            file.write(path, path.relative_to(Shared.paths.ext_dirpath))


def main():
    if os.name != "nt":
        print("Currently, this tool is Windows only.")
        return

    parser = argparse.ArgumentParser(
        prog="docx-live-reload",
        description="Preview a Docx file in MS Word. Modify the Docx file or the extracted document.xml or style.xml and the document will reload in MS Word and show the changes.",
        epilog="Created by idtareq@gmail.com",
    )

    def check_file(filename):
        if not Path(filename).suffix == ".docx":
            parser.error("File must be of the Docx format")
        elif not Path(filename).exists():
            parser.error("File does not exist")
        else:
            return filename

    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=f"%(prog)s {importlib_metadata.version('docx-live-reload')}",
    )
    parser.add_argument("docx_path", type=check_file)

    args = parser.parse_args()

    Shared.paths = Paths(args.docx_path)
    Shared.docx_watcher = FilesWatcher([Shared.paths.path])
    Shared.xmls_watcher = FilesWatcher(Shared.paths.ext_xmls)

    cmds = queue.Queue()

    threading.Thread(target=watcher_thread, daemon=True, args=(cmds,)).start()
    threading.Thread(target=input_thread, daemon=True, args=(cmds,)).start()
    threading.Thread(target=preview_thread, args=(cmds,)).start()


if __name__ == "__main__":
    main()
