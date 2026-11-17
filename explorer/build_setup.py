"""Wrap the existing MSIX in a single Windows installer executable."""

from argparse import ArgumentParser
from pathlib import Path
from subprocess import run
from tempfile import TemporaryDirectory


def main() -> None:
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--cxx", required=True, type=Path)
    options = parser.parse_args()
    cxx = options.cxx.resolve(strict=True)
    source = Path(__file__).resolve().parent
    dist = source.parent / "dist"
    package = (dist / "ChuyenDoiTaiLieu.msix").resolve(strict=True)
    with TemporaryDirectory(prefix=".setup-", dir=dist) as temporary:
        stage = Path(temporary)
        resource = stage / "setup.rc"
        resource.write_text(
            f'101 RCDATA "{package.as_posix()}"\n1 24 "{(source / "setup.manifest").as_posix()}"\n'
            f'1 ICON "{(source.parent / "assets" / "app.ico").as_posix()}"\n',
            encoding="utf-8",
        )
        compiled = stage / "setup.res"
        run([str(cxx.with_name("llvm-rc.exe")), "/fo", str(compiled), str(resource)], check=True)
        output = stage / "CaiDat-Shadrovert.exe"
        run([
            str(cxx), "-std=c++17", "-O2", "-Wall", "-Wextra", "-static", "-municode", "-mwindows",
            str(source / "setup.cpp"), str(compiled), "-o", str(output),
        ], check=True)
        output.replace(dist / output.name)


if __name__ == "__main__":
    main()
