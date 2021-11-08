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

CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrapText=False)
NUM_DIGIT = 2  # 有効桁数
DATE_FORMAT = "%Y-%m-%d（%a）"


class DataFrameFormatter:
    def __init__(self):
        pass

    @staticmethod
    def convert_col_format(
        df: pd.DataFrame, convert_contain_str: str, convert_format: str = "2%"
    ) -> pd.DataFrame:
        """
        任意の名前が含まれたカラムの数値のフォーマットを変換する

        Args:
            df (pd.DataFrame): 変換対象のカラム
            convert_contain_str (str): どのような文字が含まれたカラムを対象とするか
            convert_format (str, optional):どのようなフォーマットにするか（2% or 2f。必要であれば、convert_typeに追加）

        Returns:
            pd.DataFrame:数値のフォーマットを変換したデータフレーム
        """
        transform_type_df = df.copy()
        cols = transform_type_df.columns

        for col in cols:
            if convert_contain_str in col:
                transform_type_df[col] = convert_type(
                    transform_type_df, col, convert_format
                )
        return transform_type_df

    @staticmethod
    def convert_type(df: pd.DataFrame, col: str, convert_name: str) -> pd.DataFrame:
        """
        指定したフォーマットタイプに応じて、数値を変換する

        Args:
            df (pd.DataFrame):変換対象のデータフレーム
            col (str): 変換対象のカラム名
            convert_name (str):変換する方法

        Raises:
            ValueError: [description]

        Returns:
            pd.DataFrame: 数値のフォーマットを変換したデータフレーム
        """
        if convert_name == "2%":
            df[col] = df[col].apply(lambda x: "{:.2%}".format(x))
        elif convert_name == "2f":
            df[col] = df[col].apply(lambda x: "{:.2f}".format(x))
        else:
            raise ValueError(f"no such as **{convert_name}** convert name")

        return df[col]

    @staticmethod
    def convert_date_format(
        df: pd.DataFrame, date_col: str, input_date_format: str
    ) -> pd.Series:
        """日付のフォーマットを揃える関数

        Args:
            df (pd.DataFrame): 対象のDataFrame
            date_col (str): 日付の入っているdf内のカラム
            input_date_format (str): 変換したい日付のフォーマット

        Returns:
            pd.Series: 日付のフォーマットをinput_date_formatに変換したdf[date_col]
        """

        # 曜日を日本語で扱うためにロケールを設定
        locale.setlocale(locale.LC_TIME, "ja_JP.UTF-8")
        formatted_series = pd.to_datetime(
            df[date_col], format=input_date_format
        ).dt.strftime(DATE_FORMAT)
        return formatted_series

    @staticmethod
    def ceil_num(df: pd.DataFrame, num_col: str, num_digit: int) -> pd.Series:
        """小数点以下の桁数をnum_digitで与えられた有効桁数に切り上げ処理をする関数

        Args:
            df (pd.DataFrame): 対象のDataFrame
            num_col (str): 有効桁数を揃えたいdf内のカラム。
            num_digit (int): 揃えたい有効桁数

        Returns:
            pd.Series: 有効桁数を揃えたdf[num_col]
        """
        if len(df[num_col]) == df[num_col].isna().sum():
            return df[num_col]

        formatted_series = np.ceil(df[num_col] * 10 ** num_digit) / 10 ** num_digit
        return formatted_series

    @staticmethod
    def translate_colname(df: pd.DataFrame, to: str) -> pd.DataFrame:
        """カラム名を日本語←→英語に変換する関数

        Args:
            df (pd.DataFrame): 対象のDataFrame
            to (str): 日→英 or 英→日を指定する引数。jpかenのみ指定可能。

        Raises:
            ValueError: toにjp, en以外が指定された場合、エラー

        Returns:
            pd.DataFrame: カラム名を変更したDataFrame
        """
        # カラム名を日本語←→英語に変換するための辞書
        EN2JP_COLNAME = {
            "predict_exec_date": "予測実施日",
            "predicted_date": "予測対象日（指示日）",
            "predict_buttan": "予測物量",
            "center_modified": "センター修正値",
            "predict_num_tokusya": "予測台数",
        }
        JP2EN_COLNAME = {v: k for k, v in EN2JP_COLNAME.items()}
        assert len(EN2JP_COLNAME) == len(JP2EN_COLNAME), "日本語←→英語変換が正しく行われません。"

        if to == "jp":
            df = df.rename(columns=EN2JP_COLNAME)
        elif to == "en":
            df = df.rename(columns=JP2EN_COLNAME)
        else:
            raise ValueError(f"不正な引数値です。 to : {to}")

        return df

