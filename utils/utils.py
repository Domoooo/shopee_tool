import sys
from time import sleep


def check_path(path, name):
    if path == '':
        sys.stderr.write(f"你沒有選擇「{name}」，工具將在5秒後自動退出")
        sys.stderr.flush()
        sleep(5)
        sys.exit()
    sys.stdout.write(f"{name}在：{path}\n")
