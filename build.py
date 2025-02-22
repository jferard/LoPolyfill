import argparse
import os
import shutil
import subprocess
import zipfile
from pathlib import Path

pt = os.path.join

OFFICE_HOME = Path(os.getenv("OFFICE_HOME", "/usr/lib/libreoffice"))
OO_SDK_HOME = Path(os.getenv("OO_SDK_HOME", pt(OFFICE_HOME, "sdk")))

UNOIDL_WRITE = pt(OO_SDK_HOME, "bin/unoidl-write")

MAIN_DIR = Path(__file__).parent
SRC_DIR = MAIN_DIR / "src"
DEST_DIR = MAIN_DIR / "dest"


class LoPolyfillbuilder:
    def prepare(self):
        print("Create new dest dir...")
        shutil.rmtree(DEST_DIR, ignore_errors=True)
        DEST_DIR.mkdir(parents=True)

    def create_rdb(self):
        print("Create RDB file...")
        command = [UNOIDL_WRITE, str(OFFICE_HOME / "program/types.rdb"),
                   str(OFFICE_HOME / "program/types/offapi.rdb"),
                   str(SRC_DIR / "lopolyfill.idl"),
                   str(DEST_DIR / "lopolyfill.rdb")]
        _process = self._run_command(command)

    def copy(self):
        print("Copy files...")
        for path in SRC_DIR.glob("*"):
            if path.name in (
                    "lopolyfill.xcu", "description.xml", "LoPolyfill.py",
                    "LICENSE"):
                self._copy_file(path, DEST_DIR / path.name)
            elif path.name.startswith("package-description"):
                self._copy_file(path, DEST_DIR / path.name)
        for name in (
                "lopolyfill.xcu", "description.xml", "LoPolyfill.py",
                "LICENSE"):
            self._copy_file(SRC_DIR / name, DEST_DIR / name)
        self._copy_tree(SRC_DIR / "META-INF", DEST_DIR / "META-INF")
        self._copy_tree(SRC_DIR / "pythonpath", DEST_DIR / "pythonpath")

    def _copy_file(self, source_path: Path, dest_path: Path):
        print("  {} -> {}".format(source_path, dest_path))
        shutil.copy2(source_path, dest_path)

    def _copy_tree(self, source_dir: Path, dest_dir: Path):
        print("  {}/ -> {}/".format(source_dir, dest_dir))
        shutil.copytree(source_dir, dest_dir)

    def process(self, debug: bool, noprefix: bool):
        print("Process... (debug={}, noprefix={})".format(debug, noprefix))
        if noprefix:
            "  -> remove prefix"
            path = DEST_DIR / "lopolyfill.xcu"
            with path.open("r", encoding="utf-8") as s:
                lines = [
                    line.replace(">LOP.", ">")
                    for line in s
                ]

            with path.open("w", encoding="utf-8") as d:
                d.writelines(lines)
        if not debug:
            print("  -> remove debug")
            path1 = DEST_DIR / "LoPolyfill.py"
            path2 = DEST_DIR / "pythonpath/lopolyfill_funcs.py"
            for path in (path1, path2):
                with path.open("r", encoding="utf-8") as s:
                    lines = []
                    in_debug = False
                    for line in s:
                        if line.lstrip().startswith("# IF_DEBUG"):
                            in_debug = True
                        elif line.lstrip().startswith("# ENDIF_DEBUG"):
                            in_debug = False
                        elif not in_debug:
                            lines.append(line)

                with path.open("w", encoding="utf-8") as d:
                    d.writelines(lines)

    def _run_command(self, command, **kwargs):
        print("  > " + " ".join(command))
        process = subprocess.run(command, stdout=subprocess.PIPE,
                                 universal_newlines=True, **kwargs)
        if process.returncode != 0:
            raise Exception(f"{process.returncode} {process.stdout}")
        return process

    def make_oxt(self):
        print("Make OXT...")
        dest = zipfile.ZipFile(MAIN_DIR / "lopolyfill.oxt", 'w',
                               zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk(DEST_DIR):
            for file in files:
                name = os.path.join(root, file)
                dest.write(name, os.path.relpath(name, DEST_DIR))

    def install(self):
        print("Install OXT...")
        command = [str(OFFICE_HOME / "program/unopkg"), "add", "-f",
                   str(MAIN_DIR / "lopolyfill.oxt")]
        process = self._run_command(command)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog='LoPolyfill builder',
        description='Builds the LoPolyfill extension')
    parser.add_argument("-d", "--debug", action="store_true")
    parser.add_argument("--noprefix", action="store_true")
    args = parser.parse_args()

    builder = LoPolyfillbuilder()
    builder.prepare()
    builder.create_rdb()
    builder.copy()
    builder.process(args.debug, args.noprefix)
    builder.make_oxt()
    builder.install()
