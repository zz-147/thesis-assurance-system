"""PyInstaller 打包入口：把 tas 打成 exe 后仍以 `tas-web web ...` 形式使用。"""

from tas.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
