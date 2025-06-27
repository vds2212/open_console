import argparse
import configparser
import logging
import os
import re
import subprocess
import sys
import time
from typing import Tuple

import win32con
import win32console
import win32gui

# import re


# import msvcrt
# logging.warning("pause")
# msvcrt.getch()


def main():
    with open(r"C:\Softs\open_console.log", mode="at") as g:
        g.write(" ".join(['"%s"' % arg for arg in sys.argv]) + "\n")

    parser = argparse.ArgumentParser()

    parser.add_argument("-t", action="store_true", help="Output timing information", default=False)

    parser.add_argument("working_dir", nargs="?", help="Working Directory", default="")

    parser.add_argument("--command", action="store", help="Input", default="")

    try:
        args = parser.parse_args()

    except Exception as e:
        logging.error(str(e))
        return

    working_dir = args.working_dir
    m = re.match('([A-Z]:)"', working_dir)
    if m:
        working_dir = m.group(1) + "\\"
    logging.warning(f"working_dir: {working_dir}")
    if not working_dir:
        return

    working_dir = os.path.join(os.getcwd(), working_dir)

    config = load_config()
    console = config.get("console", "console", fallback=CONSOLE)
    if console.lower() == "conemu":
        switch_to_conemu_tab(config, working_dir, args.command)

    elif console.lower() == "console2":
        switch_to_console2_tab(config, working_dir, args.command)


def switch_to_conemu_tab(config, working_dir, command):
    if not os.path.isdir(working_dir):
        logging.warning("Dir: '%s' doesn't exist" % working_dir)
        return

    ret = run_script(config, "IsConEmu()")
    if ret == "Yes":
        norm_working_dir = os.path.normpath(working_dir)
        norm_working_dir = os.path.normcase(norm_working_dir)
        logging.warning("Working Dir: %s" % norm_working_dir)

        # Test first the current tab:
        current_dir, hwnd = run_script(config, 'GetInfo("CurDir","HWND");')
        current_dir = os.path.normpath(current_dir)
        current_dir = os.path.normcase(current_dir)
        if current_dir == norm_working_dir:
            if command:
                script = ""
                escaped_command = command
                script += r'Print("\e %s");' % escaped_command
                ret = run_script(config, script)

            focus_windows(hwnd)
            return

        # Otherwise loop on all the tabs to find the first match if any:
        tabs = run_script(config, "Tab(12);")
        tabs = [tab for tab in tabs if tab]
        for index, tab in enumerate(tabs):
            script = ""
            script += "Context(%d);" % (index + 1)
            script += 'GetInfo("CurDir","HWND");'
            current_dir, hwnd = run_script(config, script)
            current_dir = os.path.normpath(current_dir)
            current_dir = os.path.normcase(current_dir)
            logging.warning(f"Current Dir: '{current_dir}'")
            if current_dir != norm_working_dir:
                continue

            logging.warning("Tab found")
            script = ""
            script += "Tab(7,%d);" % (index + 1)
            if command:
                script += "Context(%d);" % (index + 1)
                escaped_command = command
                script += r'Print("\e %s");' % escaped_command
            ret = run_script(config, script)

            focus_windows(hwnd)
            break
        else:
            logging.warning("New tab")
            # No tab is matching create a new tab:
            run_script(config, "Recreate(0,0);")

            if True:
                time.sleep(0.4)
                escaped_working_dir = working_dir
            else:
                # It seems that when conemu postpone the execution of the script it mess up with the escaping
                escaped_working_dir = escape_directory(working_dir)

            escaped_working_dir = escape_string(escaped_working_dir)

            script = ""

            # Change the current directory:
            script += 'Print("cd /d \\"%s\\"\\ncls\n");' % escaped_working_dir
            if command:
                escaped_command = command
                script += 'Print("%s");' % escaped_command

            run_script(config, script)
            script = ""

            script += 'GetInfo("HWND");'

            hwnd = str(run_script(config, script))
            focus_windows(hwnd)

    else:
        # ConEmu is not yet started:
        logging.warning("New process")
        subprocess.Popen([config.get("path", "conemnugui", fallback=CONEMU_GUI), "-Dir", working_dir])
        if command:
            time.sleep(0.2)
            escaped_command = command
            run_script(config, 'Print("%s");' % escaped_command)


def focus_windows(hwnd):
    handle = int(hwnd, 16)
    placement = win32gui.GetWindowPlacement(handle)
    if placement[1] == win32con.SW_SHOWMINIMIZED:
        win32gui.ShowWindow(handle, win32con.SW_SHOWNORMAL)

    win32gui.SetForegroundWindow(handle)


def escape_string(command):
    escaped_command = command

    escaped_command = escaped_command.replace("\\", "\\\\")
    return escaped_command


def escape_directory(directory):
    escaped_dir = directory
    for key in ["r", "n", "t", "b", "e", "a"]:
        escaped_dir = escaped_dir.replace("\\%s" % key, "\\%s" % key.upper())
    return escaped_dir


def run_script(config, script) -> Tuple[str, str]:
    logging.warning("Script: %s", script)
    command = ["/GUIMACRO:0", script]
    command = " ".join(command)
    logging.warning(" " * 4 + command)
    process = subprocess.Popen(
        command,
        executable=config.get("paths", "conemuconsole", fallback=CONEMU_CONSOLE),
        stdout=subprocess.PIPE,
    )
    process.wait()
    result = process.communicate()[0]
    logging.warning(" " * 4 + "Result: %s" % repr(result))
    result = result.split(b";")
    result = result[-1]
    ret = result.split(b"\n")
    ret = [x.decode(config.get("miscellaneous", "codepage", fallback=CONSOLE_CODE_PAGE)) for x in ret]
    logging.warning(" " * 4 + "Result: %s" % repr(ret))

    if len(ret) == 1:
        ret = ret[0]

    if len(ret) == 0:
        ret = ["", "0"]

    logging.warning(" " * 4 + "Ret: %s" % repr(ret))
    return ret


def switch_to_console2_tab(config, working_dir, command):
    if not os.path.isdir(working_dir):
        logging.warning("Dir: '%s' doesn't exist" % working_dir)
        return

    args = ["-d", working_dir]
    if command:
        args += ["-p", command]

    subprocess.Popen([config.get("path", "console2gui", fallback=CONSOLE2_GUI)] + args)


CONSOLE = "ConEmu"
# CONSOLE="Console2"

CONEMU_ROOT = r"C:\Program Files\ConEmu"
CONEMU_CONSOLE = os.path.join(CONEMU_ROOT, r"ConEmu\ConEmuC64.exe")
CONEMU_GUI = os.path.join(CONEMU_ROOT, "ConEmu64.exe")

CONSOLE2_ROOT = r"C:\Program Files\Console"
CONSOLE2_GUI = os.path.join(CONSOLE2_ROOT, "Console.exe")

# Use chcp command to determine your console code page
# CONSOLE_CODE_PAGE = "cp437"
CONSOLE_CODE_PAGE = "cp%d" % win32console.GetConsoleCP()


def load_config():
    config = configparser.ConfigParser()
    config_path = get_config_path()
    if config_path and os.path.isfile(config_path):
        config.read(config_path)

    return config


def get_config_path():
    if "userprofile" in os.environ:
        return os.path.join(os.environ["userprofile"], ".open_console")

    return None


if __name__ == "__main__":
    main()
