import shutil
import sys
from pathlib import Path


DEFAULT_TEMPLATE = Path(
    r"E:/新2/赤兔聊天分析工作台/templates/excel/三SKU首次基准_卷发棒标准模板.xlsx"
)


def main():
    if len(sys.argv) not in (2, 3):
        raise SystemExit("usage: create_from_excel_template.py output.xlsx [template.xlsx]")

    output = Path(sys.argv[1])
    template = Path(sys.argv[2]) if len(sys.argv) == 3 else DEFAULT_TEMPLATE

    if not template.exists():
        raise FileNotFoundError(f"template not found: {template}")

    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)
    print(output)


if __name__ == "__main__":
    main()
