"""
Scripts to create an Excel report from a template using dataframes and pre-defined headers
Idea is to use the WorkSheetTemplate as a superclass, then set headers and row calculations
using inherited subclasses for a unique report

"""

import pandas
from pathlib import Path
import pandas as pd
import numpy as np
import math
import openpyxl as pxl
from openpyxl.styles import PatternFill
import json
from helperfunctions import remove_null, sheet_sort_rows, check_and_create_path
from openpyxl import Workbook


class WorkSheetTemplate:
    def __init__(self, headers: dict, data_sources: list):
        home_dir = str(Path.home())
        report_dir = Path(home_dir + '\\OpenPyReportTemplates')
        header_dir = Path(home_dir + '\\OpenPyReportTemplates' + '\\headers.json')
        check_and_create_path(report_dir)
        report_dir = Path(home_dir + '\\OpenPyReportTemplates')
        header_dir = Path(home_dir + '\\OpenPyReportTemplates' + '\\headers.json')
        if header_dir.exists() == False:
            with open(header_dir, 'w') as header_file:
                header_file.write('{}')


        """
        :param headers: dict object, IE: {'A1': df.iloc[foo]["Foo"]}
        :param data_sources: list of dataframes
        """
        self.headers = headers
        self.data_sources = data_sources

    def assign_hdrs(self, report: Path.__class__):
        """
        :param report:
        :return:
        sample header: {"A1": "Tracking", "B1": "Order", "C1": "Total UPS Incentive"}
        """
        with open(report, 'r') as f:
            headers = json.load(f)
        for hdr_index, hdr_val in headers:
            self.headers[hdr_index] = hdr_val


    def add_hdr(self, hdr: str):
        pass

    def remove_hdr(self, hdr: str):
        pass

    def swap_hdrs(self, hdr1: str, hdr2: str):
        pass

    def del_hdr(self, hdr: str):
        pass

    def save_hdrs(self):
        pass

    def change_row_calculation(self, hdr: str, datasource: pandas.DataFrame):
        pass



