"""This script enables automation of the calibration process.

Instead of manually tracking the progress of the abm, and restarting once a
certain file has been written, this script will start transcad, initialize the
abm, run the model, and terminate when a desired output is written. This script
also enables looping, such that any number of calibration iterations can be run
in sequence.

"""

import subprocess
from time import sleep

from pykeyboard import PyKeyboard


def launch_transcad():
    """Startup TransCAD and set to fullscreen.

    Returns
    -------
    proc : subprocess.Popen
        Process where TransCAD is running.

    """
    board = PyKeyboard()
    proc = subprocess.Popen([r'C:/Program Files/TransCAD 6.0\Tcw.exe'])
    sleep(5)
    board.tap_key(board.alt_key)
    board.tap_key(board.space_key)
    board.tap_key('x')
    return proc
