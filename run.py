import locale
from pathlib import Path, PosixPath
from typing import List

import numpy as np
import openpyxl as xl
import pandas as pd
from openpyxl.chart import BarChart, Reference, ScatterChart, Series
from openpyxl.styles.alignment import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# 設定
output_path = Path("output")
output_path.mkdir(parents=True, exist_ok=True)

# dfの読み込み
df = pd.read_csv(
    "../results/experiment/2021-08-26_2021-10-01_20-02-47/buttan_results.csv"
)

# 任意の整形
df_formatter = DataFrameFormatter()
# df=df_formatter.convert_col_format(df=df,convert_contain_str="ape",convert_format="2%")
# df=df_formatter.convert_col_format(df=df,convert_contain_str="物単",convert_format="2f")

# excelファイルの出力
excel_output_path = output_path / "excel_output.xlsx"
excel_outputer = ExcelOutputer()
excel_outputer.df2excel(
    df_list=[df], sheetname_list=["test"], output_path=excel_output_path
)

# excelファイルの整形
ws = xl.load_workbook(filename=excel_output_path)
excel_formatter = ExcelFormater()

# excel_formatter.convert_format(
#     ws=ws, convert_contain_str="ape", convert_format="0.00%"
# )
# excel_formatter.convert_format(
#     ws=ws, convert_contain_str="物単", convert_format="0.00"
# )
excel_formatter.adjust_width(ws=ws)
excel_formatter.align_center_row(ws=ws)

#  excelでグラフをプロット
excel_plotter = ExcelPlotter()
ws = excel_plotter.plot(
    ws=ws,
    title="物量予測誤差の推移",
    plot_cols_lines=[8, 9],
    y_label="物量予測誤差",
    x_label="日付",
    place="Z10",
)

# 保存
ws.save(excel_output_path)
