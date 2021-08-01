@id# 2>nul& @"%~dp0\python.cmd" "%~dp0\%~n0.cmd" %* & @exit /b %errorlevel%
def _(): pass

import locate
import os
import subprocess
from pathlib import Path

install_script = locate.this_dir().joinpath(r"..\tools\deploy-scripts\templates\installer.bat")
os.chdir(install_script.parent)
if Path("installer.bat").is_file():
    subprocess.call(["git", "reset", "--hard"])

with open(install_script, "r") as f:
    txt = f.read()

with open(install_script, "w") as f:
    f.write(
        txt.replace(
            ":: ========== Choose Install Dir ===========",
            ":: ========== Choose Install Dir ===========\n"
            "goto :exitchoice\n")
    )