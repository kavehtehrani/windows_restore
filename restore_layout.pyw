"""
Script to restore layout once a remote session (e.g. VPN) ends on Windows.
Stores layout of windows (enabled and visible) and saves down the state.
Once the VPN ends it restores to the last layout saved before the VPN began.

I wrote this little script since I was tired of moving my windows around while
using VPN to my PC since the resolutions are usually not the same

Kaveh Tehrani
"""

import datetime
import win32gui
import ctypes
import time
import win32con
import pickle


SM_REMOTE_SESSION = 0x1000
STR_HWND_SAVE_FILE = 'hwnd_state'
SUSPECT_TIME = 5*60                     # suspend time
d_hwnd = {}                             # storing window state


def read_windows(hwnd, _):
    """
    Records state of the windows
    """

    window_placement = win32gui.GetWindowPlacement(hwnd)
    d_hwnd[hwnd] = { 'pos':         win32gui.GetWindowRect(hwnd),
                     'text':        win32gui.GetWindowText(hwnd),
                     'enabled':     win32gui.IsWindowEnabled(hwnd),
                     'visible':     win32gui.IsWindowVisible(hwnd),
                     'maximized':   window_placement[1] == win32con.SW_SHOWMAXIMIZED,
                     'minimized':   window_placement[1] == win32con.SW_SHOWMINIMIZED
                     }


def restore_windows():
    """
    Restores windows to the state stored in d_hwnd
    """

    for hwnd in d_hwnd.keys():
        try:
            if d_hwnd[hwnd]['enabled'] and d_hwnd[hwnd]['visible'] and d_hwnd[hwnd]['text']:
                pos = d_hwnd[hwnd]['pos']
                win32gui.SetWindowPos(hwnd,
                                      win32con.HWND_NOTOPMOST,
                                      pos[0], pos[1], pos[2] - pos[0], pos[3] - pos[1],
                                      win32con.SWP_SHOWWINDOW)

                if d_hwnd[hwnd]['maximized']:
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                elif d_hwnd[hwnd]['minimized']:
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

                print(f'Restored {hwnd} window with text: {d_hwnd[hwnd]["text"]}')
        except Exception as e:
            print(f'Could not restore {hwnd} window with text: {d_hwnd[hwnd]["text"]} | {str(e)}')


if __name__ == '__main__':
    n_state = 0
    prev_state = ctypes.windll.user32.GetSystemMetrics(SM_REMOTE_SESSION) == 1

    while True:
        if ctypes.windll.user32.GetSystemMetrics(SM_REMOTE_SESSION) == 0:
            # logged in person
            if prev_state == 0:
                d_hwnd = {}
                win32gui.EnumWindows(read_windows, None)

                with open(f'{STR_HWND_SAVE_FILE}_{n_state}', 'wb') as handle:
                    pickle.dump(d_hwnd, handle, protocol=pickle.HIGHEST_PROTOCOL)
                    print(f'Wrote at {datetime.datetime.now()} to {STR_HWND_SAVE_FILE}')

            if prev_state == 1:
                # logged in person after a remote session, restore to last state before vpn
                with open(f'{STR_HWND_SAVE_FILE}_{n_state}', 'rb') as handle:
                    d_hwnd = pickle.load(handle)

                print(f'Remote session ended at {datetime.datetime.now()}. Restoring from {STR_HWND_SAVE_FILE}')
                restore_windows()

                n_state += 1

            prev_state = 0

        if ctypes.windll.user32.GetSystemMetrics(SM_REMOTE_SESSION) == 1:
            # remote session detected, do not store state
            prev_state = 1
            print(f'Remote session detected at {datetime.datetime.now()}')

        time.sleep(SUSPECT_TIME)
