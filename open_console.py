# import sys
import time
import os
import subprocess
import argparse
import configparser
import logging

# import re

import win32gui
import win32console


def main():
    # with open(r"C:\Softs\open_console.log", mode="at") as g:
    #     g.write(" ".join(['"%s"' % arg for arg in sys.argv]) + "\n")

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "-t", action="store_true", help="Output timing information", default=False
    )

    parser.add_argument("working_dir", nargs="?", help="Working Directory", default="")

    parser.add_argument("--command", action="store", help="Input", default="")

    try:
        args = parser.parse_args()

    except Exception as e:
        logging.error(str(e))
        return

    working_dir = args.working_dir
    if not working_dir:
        return

    working_dir = os.path.join(os.getcwd(), working_dir)

    switch_to_tab(working_dir, args.command)


def switch_to_tab(working_dir, command):
    print("hello")
    config = load_config()

    if not os.path.isdir(working_dir):
        print("Dir: '%s' doesn't exist" % working_dir)
        return

    ret = run_script(config, "IsConEmu()")
    if ret == "Yes":
        norm_working_dir = os.path.normpath(working_dir)
        norm_working_dir = os.path.normcase(norm_working_dir)
        print("Working Dir:", norm_working_dir)

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

            handle = int(hwnd, 16)
            win32gui.SetForegroundWindow(handle)
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
            print("Current Dir:", current_dir)
            if current_dir != norm_working_dir:
                continue

            print("Tab found")
            script = ""
            script += "Tab(7,%d);" % (index + 1)
            if command:
                script += "Context(%d);" % (index + 1)
                escaped_command = command
                script += r'Print("\e %s");' % escaped_command
            ret = run_script(config, script)

            handle = int(hwnd, 16)
            win32gui.SetForegroundWindow(handle)
            break
        else:
            print("New tab")
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
            script += 'Print("cd \\"%s\\"\\ncls\n");' % escaped_working_dir
            if command:
                escaped_command = command
                script += 'Print("%s");' % escaped_command

            run_script(config, script)
            script = ""

            script += 'GetInfo("HWND");'

            hwnd = run_script(config, script)
            handle = int(hwnd, 16)
            win32gui.SetForegroundWindow(handle)

    else:
        # ConEmu is not yet started:
        print("New process")
        subprocess.Popen(
            [config.get("path", "conemnugui", fallback=CONEMU_GUI), "-Dir", working_dir]
        )
        if command:
            time.sleep(0.2)
            escaped_command = command
            run_script(config, 'Print("%s");' % escaped_command)


def escape_string(command):
    escaped_command = command

    escaped_command = escaped_command.replace("\\", "\\\\")
    return escaped_command


def escape_directory(directory):
    escaped_dir = directory
    for key in ["r", "n", "t", "b", "e", "a"]:
        escaped_dir = escaped_dir.replace("\\%s" % key, "\\%s" % key.upper())
    return escaped_dir


def run_script(config, script):
    print("Script:", script)
    command = ["/GUIMACRO:0", script]
    command = " ".join(command)
    print(" " * 4 + command)
    process = subprocess.Popen(
        command,
        executable=config.get("paths", "conemuconsole", fallback=CONEMU_CONSOLE),
        stdout=subprocess.PIPE,
    )
    process.wait()
    result = process.communicate()[0]
    print(" " * 4 + "Result:", result)
    result = result.split(b";")
    result = result[-1]
    ret = result.split(b"\n")
    ret = [
        x.decode(config.get("miscellaneous", "codepage", fallback=CONSOLE_CODE_PAGE))
        for x in ret
    ]
    print(" " * 4 + "Result:", ret)

    if len(ret) == 1:
        ret = ret[0]

    if len(ret) == 0:
        ret = None

    print(" " * 4 + "Ret:", ret)
    return ret


CONEMU_ROOT = r"C:\Program Files\ConEmu"
CONEMU_CONSOLE = os.path.join(CONEMU_ROOT, r"ConEmu\ConEmuC64.exe")
CONEMU_GUI = os.path.join(CONEMU_ROOT, "ConEmu64.exe")

# Use chcp command to determine your console code page
# CONSOLE_CODE_PAGE = "cp437"
CONSOLE_CODE_PAGE = "cp%d" % win32console.GetConsoleCP()


def load_config():
    # defaults = {
    #     "Paths" : {
    #         "ConEmuConsole" : CONEMU_CONSOLE,
    #         "ConEmuGui" : CONEMU_GUI,
    #         },
    #     "Miscellaneous" : {
    #         "CodePage" : CONSOLE_CODE_PAGE,
    #         }
    #     }
    # config = configparser.ConfigParser(defaults=defaults)
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
