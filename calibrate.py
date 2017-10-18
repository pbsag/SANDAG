"""This script enables automation of the calibration process.

Instead of manually tracking the progress of the abm, and restarting once a
certain file has been written, this script will start transcad, initialize the
abm, run the model, and terminate when a desired output is written. This script
also enables looping, such that any number of calibration iterations can be run
in sequence.

"""

import subprocess
from time import sleep

from pymouse import PyMouse
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


def setup_abm(working_directory, start_iter=1, sample_rate=None):
    """Setup the sandag_abm with given parameters.

    Parameters
    ----------
    working_directory : str
        The path to the directory containing the gisdk and uec directories.
    start_iter : int, default : 1
        The iteration to start the process from.
    sample_rate : float or str, default : None
        The sample rate for the given starting iteration, or a string
        representing all of the sample rates.

    """
    rates = ['0.2', '0.5', '1.0']
    if sample_rate:
        if isinstance(sample_rate, float):
            rates[start_iter - 1] = str(sample_rate)
            sample_rates = ','.join(rates)
        elif isinstance(sample_rate, str):
            sample_rates = sample_rate
        else:
            raise TypeError('Type of sample_rate must be float or string.')

    mouse = PyMouse()
    board = PyKeyboard()
    xdim, ydim = mouse.screen_size()
    mouse.click(int(round(xdim * 0.55595)), int(round(ydim * 0.04952)))
    sleep(0.5)
    board.type_string('sandag_abm.lst')
    board.tap_key(board.tab_key, n=4)
    board.tap_key(board.enter_key)
    board.type_string(working_directory + '/gisdk')
    sleep(0.2)
    board.tap_key(board.enter_key)
    board.tap_key(board.tab_key, n=8)
    sleep(0.5)
    board.tap_key(board.enter_key)
    mouse.click(int(round(xdim * 0.57083)), int(round(ydim * 0.04952)))
    sleep(0.2)
    board.press_keys([board.alt_key, 'd'])
    board.press_keys([board.alt_key, 'n'])
    board.type_string('Setup Scenario')
    board.tap_key(board.enter_key)
    sleep(0.2)
    board.tap_key(board.space_key)
    sleep(5)
    board.tap_key(board.tab_key, n=22)
    board.tap_key(board.space_key)
    if start_iter == 1:
        board.tap_key(board.tab_key, n=7)
    else:
        board.tap_key(board.tab_key, n=2 + start_iter)
        board.tap_key(board.space_key)
        board.tap_key(board.tab_key, n=5 - start_iter)
    if sample_rate:
        board.tap_key(board.backspace_key)
        board.type_string(sample_rates)
    board.tap_key(board.tab_key)
    board.tap_key(board.space_key)
    sleep(1)


def launch_abm(working_directory):
    """Launch the sandag_abm.

    Parameters
    ----------
    working_directory : str
        The path to the directory containing the gisdk and uec directories.

    """
    board = PyKeyboard()
    board.tap_key(board.down_key)
    board.tap_key(board.space_key)
    sleep(0.5)
    board.tap_key(board.tab_key, n=3)
    board.tap_key(board.enter_key)
    board.type_string(working_directory)
    board.tap_key(board.enter_key)
    sleep(0.5)
    board.tap_key(board.tab_key, n=6)
    board.tap_key(board.backspace_key)
    board.tap_key(board.tab_key)
    board.tap_key(board.enter_key)
