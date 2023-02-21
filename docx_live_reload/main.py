from queue import Queue
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


def watcher(cmds: Queue):
    while True:
        if Shared.docx_watcher.changed:
            print("Change in docx file detected!")
            cmds.put("reload")

        if Shared.xmls_watcher.changed:
            print("Change in extracted xmls detected!")
            cmds.put("update")

        time.sleep(0.1)


def preview(cmds: Queue):
    word = win32.Dispatch("Word.Application", pythoncom.CoInitialize())
    word.Visible = True

    parser = etree.XMLParser(remove_blank_text=True)
    doc = run_preview(word, None, parser)

    Shared.docx_watcher.start()
    Shared.xmls_watcher.start()

    while (cmd := cmds.get()) != None:
        if cmd == "reload":
            print("reloading..")
            doc = run_preview(word, doc, parser)

        if cmd == "update":
            update(parser)

        elif cmd == "quit":
            print("quitting..")
            try:
                if doc in word.Documents:
                    doc.Close()
                word.Quit()
            except pythoncom.com_error:
                pass
            break


def run_preview(word, doc, parser):
    if doc is not None and doc in word.Documents:
        doc.Close()

    shutil.copyfile(Shared.paths.path, Shared.paths.preview_copy_path)
    doc = word.Documents.Open(str(Shared.paths.preview_copy_path))

    Shared.xmls_watcher.stop()
    with zipfile.ZipFile(Shared.paths.path) as file:
        for name in file.namelist():
            file.extract(name, Shared.paths.ext_dirpath)
        for path in Shared.paths.ext_xmls:
            tree = etree.parse(path, parser)
            tree.write(path, pretty_print=True)
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

    cmds = Queue()

    try:
        threading.Thread(target=watcher, daemon=True, args=(cmds,)).start()
        threading.Thread(target=preview, args=(cmds,)).start()

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


if __name__ == "__main__":
    main()
