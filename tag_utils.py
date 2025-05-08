from openpyxl.styles import PatternFill
import pandas as pd
from datetime import datetime, timedelta

def detect_tag_from_cell(cell, user_mapping):
    """ユーザー指定の背景色 → タグ変換"""
    if not isinstance(cell.fill, PatternFill) or cell.fill.patternType != "solid":
        return "*入力 名前:'{}'".format(cell.value or "項目")

    rgb_full = cell.fill.fgColor.rgb  # ARGB
    rgb = rgb_full[-6:].upper() if rgb_full else "FFFFFF"
    tag = user_mapping.get(rgb, "*入力")
    return f"{tag} 名前:'{cell.value or '未定義'}'"

def month_to_daily_df(ws, user_mapping):
    """月単位→日単位の行展開とタグ生成"""
    item_cells = [cell for cell in ws['A'] if cell.value]
    items = [cell.value for cell in item_cells]

    year = datetime.today().year
    month = datetime.today().month
    days = (datetime(year, month + 1, 1) - timedelta(days=1)).day

    data = []
    for day in range(1, days + 1):
        row = {"点検日": f"*日付 名前:'点検日' 初期値:'{year}/{month:02}/{day:02}'"}
        for i, cell in enumerate(item_cells):
            row[cell.value] = detect_tag_from_cell(cell, user_mapping)
        data.append(row)

    return pd.DataFrame(data)
