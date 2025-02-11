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
        shutil.rmtree(DEST_DIR, ignore_errors=True)
        DEST_DIR.mkdir(parents=True)

    def create_rdb(self):
        command = [UNOIDL_WRITE, str(OFFICE_HOME / "program/types.rdb"),
                   str(OFFICE_HOME / "program/types/offapi.rdb"),
                   str(SRC_DIR / "lopolyfill.idl"),
                   str(DEST_DIR / "lopolyfill.rdb")]
        process = self._run_command(command)

    def copy(self):
        for name in ("lopolyfill.xcu", "description.xml", "LoPolyfill.py", "LICENSE"):
            shutil.copy2(SRC_DIR / name, DEST_DIR / name)
        shutil.copytree(SRC_DIR / "META-INF", DEST_DIR / "META-INF")
        shutil.copytree(SRC_DIR / "pythonpath", DEST_DIR / "pythonpath")

    def _run_command(self, command, **kwargs):
        print("> " + " ".join(command))
        process = subprocess.run(command, stdout=subprocess.PIPE,
                                 universal_newlines=True, **kwargs)
        if process.returncode != 0:
            raise Exception(f"{process.returncode} {process.stdout}")
        return process

    def make_oxt(self):
        dest = zipfile.ZipFile(MAIN_DIR / "lopolyfill.oxt", 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk(DEST_DIR):
            for file in files:
                name = os.path.join(root, file)
                dest.write(name, os.path.relpath(name, DEST_DIR))

    def install(self):
        command = [str(OFFICE_HOME / "program/unopkg"), "add", "-f", str(MAIN_DIR / "lopolyfill.oxt")]
        process = self._run_command(command)


if __name__ == "__main__":
    builder = LoPolyfillbuilder()
    builder.prepare()
    builder.create_rdb()
    builder.copy()
    builder.make_oxt()
    builder.install()
