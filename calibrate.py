"""This script enables automation of the calibration process.

Instead of manually tracking the progress of the abm, and restarting once a
certain file has been written, this script will start transcad, initialize the
abm, run the model, and terminate when a desired output is written. This script
also enables looping, such that any number of calibration iterations can be run
in sequence.

"""

import os.path as osp
import subprocess
from time import sleep, time

import psutil
from pymouse import PyMouse
from pykeyboard import PyKeyboard

from update import update


FILES = {'AO': ['AutoOwnership', 'aoResults.csv', '1_AO'],
         'CDAP': ['CoordinatedDailyActivityPattern', 'personData_{}.csv',
                  '2_CDAP']}


def kill_proc_tree(pid, including_parent=True):
    parent = psutil.Process(pid)
    children = parent.children(recursive=True)
    for child in children:
        child.kill()
    _, _ = psutil.wait_procs(children, timeout=5)
    if including_parent:
        parent.kill()
        parent.wait(5)


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


def compile_abm(working_directory):
    """Compile the sandag_abm.

    Parameters
    ----------
    working_directory : str
        The path to the directory containing the gisdk and uec directories.

    """
    mouse = PyMouse()
    board = PyKeyboard()
    xdim, ydim = mouse.screen_size()
    mouse.click(int(round(xdim * 0.55595)), int(round(ydim * 0.04952)))
    sleep(.5)
    board.type_string('sandag_abm.lst')
    sleep(.5)
    board.tap_key(board.tab_key, n=4)
    sleep(.5)
    board.tap_key(board.enter_key)
    sleep(.5)
    board.type_string(working_directory + '/gisdk')
    sleep(.5)
    board.tap_key(board.enter_key)
    sleep(.5)
    board.tap_key(board.tab_key, n=8)
    sleep(.5)
    board.tap_key(board.enter_key)
    sleep(2)


def set_abm_params(start_iter, sample_rates):
    """Set the parameters for the sandag_abm.

    Parameters
    ----------
    start_iter : int
        The iteration to start the process from.
    sample_rates : str or None
        The sample rates for the abm in string form.

    """
    mouse = PyMouse()
    board = PyKeyboard()
    xdim, ydim = mouse.screen_size()
    mouse.click(int(round(xdim * 0.57083)), int(round(ydim * 0.04952)))
    sleep(1)
    board.press_keys([board.alt_key, 'd'])
    board.press_keys([board.alt_key, 'n'])
    board.type_string('Setup Scenario')
    board.tap_key(board.enter_key)
    sleep(0.2)
    board.tap_key(board.space_key)
    sleep(5)
    # Set database write to 'No'
    board.tap_key(board.tab_key, n=22)
    board.tap_key(board.space_key)
    if start_iter == 1:
        board.tap_key(board.tab_key, n=7)
    else:
        board.tap_key(board.tab_key, n=2 + start_iter)
        board.tap_key(board.space_key)
        board.tap_key(board.tab_key, n=5 - start_iter)
    if sample_rates:
        board.tap_key(board.backspace_key)
        board.type_string(sample_rates)
    board.tap_key(board.tab_key)
    board.tap_key(board.space_key)
    sleep(1)


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
    sample_rates = None
    if sample_rate:
        if isinstance(sample_rate, float):
            rates[start_iter - 1] = str(sample_rate)
            sample_rates = ','.join(rates)
        elif isinstance(sample_rate, str):
            sample_rates = sample_rate
        else:
            raise TypeError('Type of sample_rate must be float or string.')

    compile_abm(working_directory)
    set_abm_params(start_iter, sample_rates)


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


def calibrate(working_directory, start_iter=1, sample_rate=None, max_iters=3,
              input_path='.', output_path='../Model Calibration'):
    """Calibrate abm with given parameters.

    Parameters
    ----------
    working_directory : str
        The path to the directory containing the gisdk and uec directories.
    start_iter : int, default : 1
        The iteration to start the process from.
    sample_rate : float or str, default : None
        The sample rate for the given starting iteration, or a string
        representing all of the sample rates.
    max_iters : int, default : 3
        The maximum number of iterations to run for each step.
    iput_path : str, default : '.'
        The relative path to the output and uec directories.
    output_path : str, default : '../Model Calibration'
        The relative path to the directory containing the calibration
        directories.

    """
    steps = ['AO', 'CDAP', 'DC', 'MC', 'ML', 'TR', 'AS']
    for step in steps:
        cal_path = output_path + '/{}'.format(FILES[step][2])
        update(0, input_path, cal_path, method=step)
        for iter_ in range(max_iters):
            proc = launch_transcad()
            setup_abm(working_directory, start_iter=start_iter,
                      sample_rate=sample_rate)
            launch_abm(working_directory)
            start_time = time()
            sleep(18000)
            last_mod = osp.getmtime(FILES[step][1].format(start_iter))
            while last_mod < start_time:
                sleep(1800)
                last_mod = osp.getmtime(FILES[step][1].format(start_iter))
            kill_proc_tree(proc.pid, including_parent=True)
            update(iter_ + 1, input_path, cal_path, method=step)
            print('Completed Step {} iteration {}.\n'
                  .format(FILES[step][0], iter_ + 1))
