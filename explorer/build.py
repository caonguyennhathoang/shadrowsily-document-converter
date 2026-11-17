"""Build the Windows 11 x64 package from a PyInstaller onedir build."""

from argparse import ArgumentParser
from pathlib import Path
from shutil import copyfile, copytree
from subprocess import run
from tempfile import TemporaryDirectory

from PIL import Image


def make_logos(destination: Path) -> None:
    destination.mkdir()
    logo = Image.open(Path(__file__).resolve().parent.parent / "assets" / "app-icon-512.png")
    for filename, size in (("StoreLogo.png", 50), ("Square150.png", 150), ("Square44.png", 44)):
        logo.resize((size, size), Image.Resampling.LANCZOS).save(destination / filename)
    logo.close()


def skip_unused_template(directory, names):
    # python-docx loads default.docx. The exploded copy contains a reserved MSIX filename.
    if Path(directory).parts[-2:] == ("docx", "templates") and "default.docx" in names:
        return {"default-docx-template"}
    return set()


def main() -> None:
    source = Path(__file__).resolve().parent
    dist = source.parent / "dist"
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--cxx", required=True, help="Path to the x64 LLVM MinGW clang++ executable")
    parser.add_argument("--makeappx", required=True, help="Path to MakeAppx.exe from the Windows SDK")
    parser.add_argument("--app-dir", type=Path, default=dist / "ChuyenDoiTaiLieu",
                        help="Directory containing ChuyenDoiTaiLieu.exe and its _internal runtime")
    options = parser.parse_args()
    with TemporaryDirectory(prefix=".msix-", dir=dist) as temporary:
        stage = Path(temporary) / "payload"
        stage.mkdir()
        copyfile(options.app_dir / "ChuyenDoiTaiLieu.exe", stage / "ChuyenDoiTaiLieu.exe")
        copytree(options.app_dir / "_internal", stage / "_internal", ignore=skip_unused_template)
        copyfile(source / "AppxManifest.xml", stage / "AppxManifest.xml")
        make_logos(stage / "Assets")
        run([
            options.cxx, "-std=c++17", "-O2", "-Wall", "-Wextra", "-shared", "-static",
            "-o", str(stage / "ExplorerCommand.dll"), str(source / "command.cpp"),
            str(source / "command.def"), "-lole32", "-lshlwapi", "-luuid",
        ], check=True)
        package = Path(temporary) / "ChuyenDoiTaiLieu.msix"
        run([
            options.makeappx, "pack", "/o", "/d", str(stage), "/p",
            str(package),
        ], check=True)
        package.replace(dist / package.name)
    copyfile(source / "CaiMenuWindows11.cmd", dist / "CaiMenuWindows11.cmd")


if __name__ == "__main__":
    main()
