from datetime import datetime, date, timedelta
import calendar
from tokenize import group
import requests
import xlwings as xw
import pandas as pd
from pandas import DataFrame
import json
import os


class Generator:
    def __init__(self, year) -> None:
        self.start_year = year

    def get_holidays(self, year):
        # reference:
        # https://blog.csdn.net/u012981882/article/details/112552450
        # http://www.apihubs.cn/#/holiday

        """
        workday : 1-yes,2-no
        weekend : 1-yes,2-no
        """
        url = "https://api.apihubs.cn/holiday/get"
        params = {
            "year": year,
            # 'month':'202205',
            "size": 400,
            "order_by": 1,  # asc
        }
        url += self.format_param_to_str(params)
        holidays = requests.get(url)
        return holidays.text

    def format_param_to_str(self, params):
        pam = ""
        if params:
            pam += "?"
            for idx, key in enumerate(params):
                if idx != 0:
                    pam += "&"
                pam += str(key) + "=" + str(params[key])
        return pam

    def write_to_excel(self, df):
        # reference :
        # https://zhuanlan.zhihu.com/p/353669230
        # https://docs.xlwings.org/en/stable/api.html

        # app=xw.App(visible=False,add_book=False)

        file_name = f"WorkCalendar_{self.start_year}.xlsx"

        wb = xw.Book()
        sheet = wb.sheets[0]

        self.global_format_before(sheet)
        self.write_top_actions(sheet, (0, 0))
        self.write_hearders(sheet, (7, 0))
        self.value_prepare(df)
        self.write_calendar(sheet, df, (7, 3))
        self.write_task_sample(sheet, (13, 0))
        self.global_format_after(sheet)
        self.freeze(wb, (13, 3))

        if os.path.exists(file_name):
            os.remove(file_name)
        wb.save(file_name)
        wb.close()
        # app.quit()

    def freeze(self, workbook, start_cell):

        start_row = start_cell[0]
        start_column = start_cell[1]

        active_window = workbook.app.api.ActiveWindow
        active_window.FreezePanes = False
        active_window.SplitColumn = start_column
        active_window.SplitRow = start_row
        active_window.FreezePanes = True

    def write_task_sample(self, sheet, start_cell):

        start_row = start_cell[0]
        start_column = start_cell[1]

        first_font_color = (131, 60, 12)  # 0x833C0C
        second_font_color = (109, 158, 235)  # 0x6D9EEB
        third_font_color = (191, 143, 0)  # 0xBF8F00

        border_line_style = 1

        sheet[start_row, start_column].value = "Agile"
        sheet[start_row, start_column + 1].value = "Specific Agile"
        sheet[start_row, start_column + 2].value = "Collect Data"
        sheet[start_row, start_column].font.bold = True
        sheet[start_row, start_column + 1].font.bold = True
        sheet[start_row + 1, start_column + 2].value = "Extract Information"
        sheet[start_row + 2, start_column + 2].value = "Practise"
        sheet[start_row + 3, start_column + 2].value = "Markdown"
        sheet[start_row, start_column].font.color = first_font_color
        sheet[start_row, start_column + 1].font.color = second_font_color
        sheet[start_row : start_row + 4, start_column + 2].font.color = third_font_color
        sheet[start_row, start_column].font.size = 12
        sheet[start_row : start_row + 4, start_column].api.Merge()
        sheet[start_row : start_row + 4, start_column + 1].api.Merge()

        for border in [7, 8, 9, 10, 11, 12]:
            sheet[
                start_row : start_row + 4, start_column : start_column + 3
            ].api.Borders(border).LineStyle = border_line_style

    def write_calendar(self, sheet, df: DataFrame, start_cell):

        # format prepare
        start_row = start_cell[0]
        start_column = start_cell[1]

        year_row = start_row
        quarter_row = start_row + 1
        month_desc_row = start_row + 2
        iso_week_row = start_row + 3
        month_day_row = start_row + 4
        week_day_row = start_row + 5

        ##################
        #     Days       #
        ##################
        for idx, row in df.iterrows():

            column = start_column + idx
            sheet[week_day_row, column].value = row["week_day"]
            sheet[month_day_row, column].value = row["month_day"]

            font_color_white = 0x000000
            font_color_black = 0xFFFFFF

            # set up font color & color for day's row
            if row["workday"] == 1 and row["weekend"] == 1:
                sheet[week_day_row, column].color = (91, 155, 213)
                sheet[week_day_row, column].api.Font.Color = font_color_white
                sheet[month_day_row, column].color = (91, 155, 213)
                sheet[month_day_row, column].api.Font.Color = font_color_white
            elif row["workday"] == 1 and row["weekend"] == 2:
                sheet[week_day_row, column].color = (109, 158, 235)
                sheet[week_day_row, column].api.Font.Color = font_color_black
                sheet[month_day_row, column].color = (109, 158, 235)
                sheet[month_day_row, column].api.Font.Color = font_color_black
            elif row["workday"] == 2 and row["weekend"] == 2:
                sheet[week_day_row, column].color = (255, 192, 0)
                sheet[week_day_row, column].api.Font.Color = font_color_white
                sheet[month_day_row, column].color = (255, 192, 0)
                sheet[month_day_row, column].api.Font.Color = font_color_white
            elif row["workday"] == 2 and row["weekend"] == 1:
                sheet[week_day_row, column].color = (255, 255, 0)
                sheet[week_day_row, column].api.Font.Color = font_color_white
                sheet[month_day_row, column].color = (255, 255, 0)
                sheet[month_day_row, column].api.Font.Color = font_color_white

        ##################
        #     Weeks      #
        ##################
        # The days in Jan. but iso_week far beyond Jan. are included in the week of last year
        df_week_out_year = df.loc[(df["month_num"] == 1) & (df["iso_week"] > 10)]
        df_week_in_year = df.drop(df_week_out_year.index)
        df_week_out_year.reset_index(inplace=True)
        df_week_in_year.reset_index(inplace=True)

        iso_week_out_gp = df_week_out_year.groupby(by=["iso_week"])
        iso_week_in_gp = df_week_in_year.groupby(by=["iso_week"])
        iso_week_end_column = start_column
        iso_week_start_column = 0

        for iw, group in iso_week_out_gp:

            # set up format of week's row
            iso_week_start_column = iso_week_end_column
            iso_week_end_column = iso_week_start_column + group.shape[0]

            week = group.iloc[0]["iso_week_desc"]

            self.set_week_and_day_row_format(
                sheet,
                week=week,
                iso_week_row=iso_week_row,
                month_day_row=month_day_row,
                week_day_row=week_day_row,
                iso_week_start_column=iso_week_start_column,
                iso_week_end_column=iso_week_end_column,
            )
            self.set_weekly_sample_border_format(
                sheet,
                start_row=week_day_row + 1,
                end_row=week_day_row + 1 + 4,
                start_column=iso_week_start_column,
                end_column=iso_week_end_column,
            )

        for iw, group in iso_week_in_gp:

            # set up format of week's row
            iso_week_start_column = iso_week_end_column
            iso_week_end_column = iso_week_start_column + group.shape[0]

            week = group.iloc[0]["iso_week_desc"]

            self.set_week_and_day_row_format(
                sheet,
                week=week,
                iso_week_row=iso_week_row,
                month_day_row=month_day_row,
                week_day_row=week_day_row,
                iso_week_start_column=iso_week_start_column,
                iso_week_end_column=iso_week_end_column,
            )
            self.set_weekly_sample_border_format(
                sheet,
                start_row=week_day_row + 1,
                end_row=week_day_row + 1 + 4,
                start_column=iso_week_start_column,
                end_column=iso_week_end_column,
            )

        ##################
        #     Months     #
        ##################
        month_gp = df.groupby(by=["month_num"])

        month_end_column = start_column
        month_start_column = 0

        for iw, group in month_gp:

            month_start_column = month_end_column
            month_end_column = month_start_column + group.shape[0]

            month_desc = group.iloc[0]["month_desc"]

            self.set_month_row_format(
                sheet,
                desc=month_desc,
                desc_row=month_desc_row,
                start_column=month_start_column,
                end_column=month_end_column,
            )

        ##################
        #     Quarter    #
        ##################
        quarter_gp = df.groupby(by=["quarter"])

        quarter_end_column = start_column
        quarter_start_column = 0

        for iw, group in quarter_gp:

            quarter_start_column = quarter_end_column
            quarter_end_column = quarter_start_column + group.shape[0]

            quarter_desc = group.iloc[0]["quarter"]

            self.set_quarter_row_format(
                sheet,
                desc=quarter_desc,
                desc_row=quarter_row,
                start_column=quarter_start_column,
                end_column=quarter_end_column,
            )

        ##################
        #     Years      #
        ##################
        year_gp = df.groupby(by=["year"])

        year_end_column = start_column
        year_start_column = 0

        for iw, group in year_gp:

            year_start_column = year_end_column
            year_end_column = year_start_column + group.shape[0]

            year_desc = group.iloc[0]["year"]

            self.set_year_row_format(
                sheet,
                desc=year_desc,
                desc_row=year_row,
                start_column=year_start_column,
                end_column=year_end_column,
            )

    def set_year_row_format(self, sheet, **kwargs):

        weight = 4

        desc = kwargs.get("desc")
        row = kwargs.get("desc_row")
        start_column = kwargs.get("start_column")
        end_column = kwargs.get("end_column")

        sheet[row, start_column].value = desc
        sheet[row, start_column].font.bold = True
        sheet[row, start_column:end_column].color = (174, 170, 170)
        sheet[row, start_column:end_column].api.Merge()
        sheet[row, start_column:end_column].api.Borders(7).Weight = weight
        sheet[row, start_column:end_column].api.Borders(10).Weight = weight
        sheet[row, start_column:end_column].column_width = 3

    def set_quarter_row_format(self, sheet, **kwargs):

        weight = 4

        desc = kwargs.get("desc")
        row = kwargs.get("desc_row")
        start_column = kwargs.get("start_column")
        end_column = kwargs.get("end_column")

        sheet[row, start_column].value = desc
        sheet[row, start_column].font.bold = True
        sheet[row, start_column:end_column].color = (198, 224, 180)
        sheet[row, start_column:end_column].api.Merge()
        sheet[row, start_column:end_column].api.Borders(7).Weight = weight
        sheet[row, start_column:end_column].api.Borders(10).Weight = weight

    def set_month_row_format(self, sheet, **kwargs):

        weight = 4

        desc = kwargs.get("desc")
        row = kwargs.get("desc_row")
        start_column = kwargs.get("start_column")
        end_column = kwargs.get("end_column")

        sheet[row, start_column].value = desc
        sheet[row, start_column].font.bold = True
        sheet[row, start_column:end_column].color = (221, 235, 247)
        sheet[row, start_column:end_column].api.Merge()
        sheet[row, start_column:end_column].api.Borders(7).Weight = weight
        sheet[row, start_column:end_column].api.Borders(10).Weight = weight

    def set_week_and_day_row_format(self, sheet, **kwargs):
        weight = -4138  # middle
        middle_border_color = 0xFFFFFF  # while
        middle_border_linestyle = 1  # xlHairline

        week = kwargs.get("week")
        iso_week_row = kwargs.get("iso_week_row")
        month_day_row = kwargs.get("month_day_row")
        week_day_row = kwargs.get("week_day_row")
        iso_week_start_column = kwargs.get("iso_week_start_column")
        iso_week_end_column = kwargs.get("iso_week_end_column")

        sheet[iso_week_row, iso_week_start_column].value = week
        sheet[iso_week_row, iso_week_start_column:iso_week_end_column].color = (
            60,
            120,
            216,
        )
        sheet[iso_week_row, iso_week_start_column:iso_week_end_column].api.Merge()
        sheet[iso_week_row, iso_week_start_column:iso_week_end_column].api.Borders(
            7
        ).Weight = weight
        sheet[iso_week_row, iso_week_start_column:iso_week_end_column].api.Borders(
            10
        ).Weight = weight

        # set up border of day's row
        sheet[month_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            11
        ).LineStyle = middle_border_linestyle
        sheet[month_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            11
        ).Color = middle_border_color
        sheet[month_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            7
        ).Weight = weight
        sheet[month_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            9
        ).Weight = weight
        sheet[month_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            10
        ).Weight = weight

        sheet[week_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            11
        ).LineStyle = middle_border_linestyle
        sheet[week_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            11
        ).Color = middle_border_color
        sheet[week_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            7
        ).Weight = weight
        sheet[week_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            9
        ).Weight = weight
        sheet[week_day_row, iso_week_start_column:iso_week_end_column].api.Borders(
            10
        ).Weight = weight

    def set_weekly_sample_border_format(self, sheet, **kwargs):

        weight = -4138  # middle

        start_row = kwargs.get("start_row")
        end_row = kwargs.get("end_row")
        start_column = kwargs.get("start_column")
        end_column = kwargs.get("end_column")

        for border in [7, 10, 8, 9]:
            sheet[start_row:end_row, start_column:end_column].api.Borders(
                border
            ).Weight = weight

    def value_prepare(self, df):
        """
        Useful fields:
            workday
            weekend
            python_datetime
            week_day
            month_day
            week
            iso_week
            iso_week_desc
            month_desc
            month_num
            quarter
            year
        """
        df["python_datetime"] = df["date"].apply(
            lambda x: datetime.strptime(str(x), "%Y%m%d")
        )
        df.sort_values(by=["python_datetime"], inplace=True)
        df.reset_index(inplace=True)
        df["week_day"] = df["week"].apply(lambda x: "D" + str(x))
        df["month_day"] = df["python_datetime"].apply(
            lambda x: datetime.strftime(x, "%d").lstrip("0")
        )
        df["iso_week"] = df["python_datetime"].apply(lambda x: x.isocalendar()[1])
        df["iso_week_desc"] = df["python_datetime"].apply(
            lambda x: "W" + str(x.isocalendar()[1])
        )
        df["month_desc"] = df["python_datetime"].apply(
            lambda x: datetime.strftime(x, "%B")
        )
        df["month_num"] = df["python_datetime"].apply(
            lambda x: int(datetime.strftime(x, "%m"))
        )
        df.loc[(df["month_num"] <= 3), "quarter"] = "Q1"
        df.loc[(df["month_num"] >= 4) & (df["month_num"] <= 6), "quarter"] = "Q2"
        df.loc[(df["month_num"] >= 7) & (df["month_num"] <= 9), "quarter"] = "Q3"
        df.loc[(df["month_num"] >= 10) & (df["month_num"] <= 12), "quarter"] = "Q4"

    def write_hearders(self, sheet, start_cell):

        start_row = start_cell[0]
        start_column = start_cell[1]

        font_color = 0xFFFFFF
        background_color = 0x757171

        border_line_style = 1

        for border in [11]:
            sheet[
                start_row : start_row + 6, start_column : start_column + 3
            ].api.Borders(border).LineStyle = border_line_style

        sheet[start_row, start_column].value = "Topic"
        sheet[start_row, start_column + 1].value = "Sub Topic"
        sheet[start_row, start_column + 2].value = "Tasks"
        sheet[start_row, start_column].font.bold = True
        sheet[start_row, start_column + 1].font.bold = True
        sheet[start_row, start_column + 2].font.bold = True
        sheet[start_row, start_column].color = background_color
        sheet[start_row, start_column + 1].color = background_color
        sheet[start_row, start_column + 2].color = background_color
        sheet[start_row, start_column].api.Font.Color = font_color
        sheet[start_row, start_column + 1].api.Font.Color = font_color
        sheet[start_row, start_column + 2].api.Font.Color = font_color
        sheet[start_row, start_column].font.size = 16
        sheet[start_row, start_column + 1].font.size = 16
        sheet[start_row, start_column + 2].font.size = 16
        sheet[start_row : start_row + 6, start_column].api.Merge()
        sheet[start_row : start_row + 6, start_column + 1].api.Merge()
        sheet[start_row : start_row + 6, start_column + 2].api.Merge()

    def write_top_actions(self, sheet, start_cell):

        start_row = start_cell[0]
        start_column = start_cell[1]

        sheet[start_row, start_column + 1].value = "Main"
        sheet[start_row, start_column + 2].value = "Secondary"
        sheet[start_row + 1, start_column].value = "Collect Data"
        sheet[start_row + 2, start_column].value = "Extract Information"
        sheet[start_row + 3, start_column].value = "Pracise"
        sheet[start_row + 4, start_column].value = "Markdown"
        sheet[start_row + 5, start_column].value = "Others"
        sheet[start_row + 6, start_column].value = "Out of Time"
        sheet[start_row + 1, start_column + 1].color = (198, 224, 180)
        sheet[start_row + 2, start_column + 1].color = (0, 176, 249)
        sheet[start_row + 3, start_column + 1].color = (47, 117, 181)
        sheet[start_row + 4, start_column + 1].color = (102, 102, 51)
        sheet[start_row + 5, start_column + 1].color = (255, 153, 102)
        sheet[start_row + 6, start_column + 1].color = (153, 0, 0)
        sheet[start_row + 1 : start_row + 7, start_column + 2].color = (247, 235, 221)

        for i in range(start_row + 1, start_row + 7):
            self.set_pure_color_cell_border(sheet[i, start_column + 1])

        for i in range(start_row + 1, start_row + 7):
            self.set_pure_color_cell_border(sheet[i, start_column + 2])

    def set_pure_color_cell_border(self, cell):
        # Borders: 9-bottom;7-left;8-top;10-right
        for b in [7, 8, 9, 10]:
            cell.api.Borders(b).LineStyle = 1
            cell.api.Borders(b).Color = 0xFFFFFF  # white
            cell.api.Borders(b).Weight = 3

    def global_format_before(self, sheet):
        sheet[:, :].api.HorizontalAlignment = -4108

    def global_format_after(self, sheet):
        sheet[0:17, 0:4].autofit()

    def generate(self):
        raw_data = self.get_holidays(self.start_year)
        sp = json.loads(raw_data)
        # sp = pd.read_json("date/api_sample.json")
        data = sp["data"]["list"]
        df = DataFrame.from_dict(data, orient="columns")
        self.write_to_excel(df)


if "__main__" == __name__:
    Generator(2022).generate()
