# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows,
# actions, and settings.

import copy
import os
import time
from datetime import date

import pandas as pd
import xlsxwriter
from natsort import natsorted

font_size = 15


class ExcelReport:

    def __init__(self):
        self.manual_tc = None
        self.auto_tc = None
        self.all = None

        self.op_list_percentage = []
        self.manual_new = 0
        self.columns = []
        self.operator_sprint = {}
        self.status_tcs = []
        self.manual = False
        self.status_based_dic = {}
        self.cap_format = None
        self.tc_maual_ids = []
        self.tc_auto_ids = []
        self.tc_types = ['Completed',
                         'In review',
                         'In Progress',
                         'New',
                         'Total testcase']
        self.not_eligible = {}
        self.invalid_tcs = {}
        self.excel_path = ''
        self.each_operator_all_sprint_auto = {}
        self.manual_data = {}

        self.op_df = {
            'Type of Testcases': self.tc_types,
            'Automation testcases': [],
            'Automation percentage': [],
            'Manual testcases': [],
            'Manual percentage': [],
            'Overall testcases': [],
            'Overall percentage': []
        }
        self.overall_summary = {'Testcases': ['Automation completed',
                                              'Automation in progress',
                                              'Automation in review',
                                              'Automation backlog',
                                              'Manual completed',
                                              'Total'],
                                'values (%)': []}

        self.sprint_bvt_fvt_svt = {}
        self.sprint_bvt_fvt_svt_auto = {}
        self.sumry_bvt_fvt_svt = {'BVT': [0, 0], 'FVT': [0, 0], 'SVT': [0, 0]}

        # operator data for all operator in each sprint
        self.each_operator_sprint_values = {}

        # data of all operators automation for all sprints (automation sheet)
        self.each_operator_all_sprint_data = {}
        self.manual_data = {
            'Type of Testcases': self.op_df['Type of Testcases']}
        # summary of all sprint data (in overall summary tab)
        self.all_operator_sprint_data = {}
        self.all_operator_sprint_auto = {}
        self.each_operator_all_sprint_auto = {
            'Type of Testcases': self.op_df['Type of Testcases']}

        self.each_operator_sprint_values_auto = {}
        self.worksheet = None
        self.workbook = None

    # ************************************************************************
    def generat_detailed_report(self,
                                excel_file, columns=[],
                                operator_sprint={},
                                test_ids=None):
        try:
            self.columns = columns
            self.operator_sprint = operator_sprint

            df1 = pd.read_excel(
                excel_file,
                engine='openpyxl', header=None,
                sheet_name='Test cases - sheet 1')
            if test_ids:
                df1 = df1[df1[self.columns.index('ID')].isin(test_ids)]

            self.excel_path = excel_file
            df1.to_csv('updated_TCs', sep='\t')

            self.not_eligible = self.get_table_json(df1[df1[self.columns.index(
                'AUTOMATIONSTATUS')] == 'Not Eligible'])

            # df1 = df1[df1[self.columns.index(
            #     'AUTOMATIONSTATUS')] != 'Not Eligible']

            search_str = 'BVT|FVT|SVT'
            filter = df1[self.columns.index('MODULE')].str.contains(
                search_str)
            self.invalid_tcs = self.get_table_json(df1[~filter])

            df1 = df1[filter]
            self.all = df1
            self.generate_report(df1)

            # sheets_df = {'Overall_summary': self.op_df}
            self.generate_overall_summary()

            auto_tcs = self.get_table_json(self.auto_tc)
            manual_tcs = self.get_table_json(self.manual_tc)

            sheets_data = {
                'Overall_summary': self.overall_summary,
                'Automation_stats': self.each_operator_all_sprint_auto,
                'Manual_stats': self.manual_data,
                'Not_eligible_TCs': self.not_eligible,
                'Invalid_format_TCs': self.invalid_tcs,
                'Automatable_TCs': auto_tcs,
                'Manual_TCs': manual_tcs}
            return sheets_data
            # self.data_in_excel(excel_file, sheets_data)
        except Exception as e:
            print('Exception found %s', str(e))

    def generate_report(self, df):
        try:
            print('====================Report Data=====================')
            # print('Total No. of testcases: ', len(df))
            self.manual_tc = df[df[self.columns.index('AUTOMATION')] == 'No']
            self.auto_tc = df[df[self.columns.index('AUTOMATION')] == 'Yes']

            self.get_all_executed_tc()

            self.generate_overall_data(df)

            self.generate_automation_data()

            self.generate_manual_data()

            self.generate_sprint_data(self.auto_tc, True)
            self.each_operator_sprint_values_auto = \
                self.each_operator_sprint_values
            self.each_operator_sprint_values = {}
            self.all_operator_sprint_auto = self.all_operator_sprint_data
            sum_list, per_list = self.sum_dict_lists(
                self.each_operator_all_sprint_data)
            self.each_operator_all_sprint_auto['Testcases (abs)'] = sum_list
            self.each_operator_all_sprint_auto['Testcases (%)'] = per_list
            self.sprint_bvt_fvt_svt_auto = self.sprint_bvt_fvt_svt
            self.sumry_bvt_fvt_svt = {
                'BVT': [
                    0, 0], 'FVT': [
                    0, 0], 'SVT': [
                    0, 0]}

            self.each_operator_all_sprint_auto.update(
                self.each_operator_all_sprint_data)

            self.generate_sprint_data(self.manual_tc)
            sum_list, per_list = self.sum_dict_lists(
                self.each_operator_all_sprint_data)
            self.manual_data['Testcases (abs)'] = sum_list
            self.manual_data['Testcases (%)'] = per_list
            self.manual_data.update(self.each_operator_all_sprint_data)

        except Exception as e:
            print(str(e))

    # *************************************************************************
    def data_in_excel(self, excel_file='', sheets_df={}, merge_data_obj=None):
        try:

            file_name, ext = os.path.splitext(excel_file)
            op_file = file_name + '_Summary' + ext

            if os.path.exists(op_file):
                os.remove(op_file)

            # Create a Pandas Excel writer using XlsxWriter as the engine.
            # writer = pd.ExcelWriter(op_file, engine='xlsxwriter')
            self.workbook = xlsxwriter.Workbook(op_file)
            self.cap_format = self.workbook.add_format(
                {'bold': True, 'font_size': font_size,
                 'font_color': '#33001a'})
            overview_sheet = None
            header = []
            # list of dic of headers in config.ini
            for col in self.columns:
                header.append({'header': col})

            # add sheets and data init in workbook
            for sheet, data in sheets_df.items():
                # Create a Pandas dataframe from some data.

                # Write the dataframe data to XlsxWriter.
                # Turn off the default header and
                # index and skip one row to allow us to insert
                # a user defined header.
                self.worksheet = self.workbook.add_worksheet(sheet)

                if 'Overall_summary' in sheet:
                    overview_sheet = self.worksheet
                elif 'Automation_stats' in sheet:
                    self.other_tabs(data,
                                    self.each_operator_sprint_values_auto,
                                    True)
                elif 'Manual_stats' in sheet:
                    self.other_tabs(data, self.each_operator_sprint_values)
                else:
                    self.add_table_by_json(data, 1, 1, False,
                                           header)

            # writer.save()
            self.worksheet = overview_sheet

            self.generate_sprint_data(self.auto_tc, True)

            #getting added at last moment to merge report
            if merge_data_obj:
                merge_data_obj.generate_sprint_data(merge_data_obj.auto_tc, True)
                self.sumry_bvt_fvt_svt["BVT"] = self._list_sum(self.sumry_bvt_fvt_svt["BVT"], merge_data_obj.sumry_bvt_fvt_svt["BVT"])
                self.sumry_bvt_fvt_svt["FVT"] = self._list_sum(
                    self.sumry_bvt_fvt_svt["FVT"],
                    merge_data_obj.sumry_bvt_fvt_svt["FVT"])
                self.sumry_bvt_fvt_svt["SVT"] = self._list_sum(
                    self.sumry_bvt_fvt_svt["SVT"],
                    merge_data_obj.sumry_bvt_fvt_svt["SVT"])

            self.overview_tab('Overall_summary')

            self.workbook.close()
        except Exception as e:
            print('Exception found %s', str(e))

    # *************************************************************************
    def overview_tab(self, sheet_name):

        table_start_col = 1
        table_start_row = 45
        chart_row = 3
        chart_col = 1

        heading_format = self.workbook.add_format({'bold': True,
                                                   'border': 1,
                                                   'align': 'center',
                                                   'font_size': 22,
                                                   'bg_color': '#000066',
                                                   'font_color': '#cceeff'})
        product = self.excel_path.rsplit('/', 1)[1]

        if product:
            if 'common' in product.lower():
                product = 'Common'
            elif 'sds' in product.lower():
                product = 'SDS'
            elif 'hci' in product.lower():
                product = 'HCI'
            else:
                product = ''
        title = 'IBM Spectrum Fusion Automation ({}) report'.format(product)
        curr_date = str(date.today())
        product = self.excel_path.rsplit('/',1)[1]
        self.worksheet.merge_range('F1:O1', title + curr_date,
                                   heading_format)

        self.write_caption(self.worksheet, table_start_row - 1,
                           table_start_col,
                           'Summary of Automatble/non-automatable',
                           self.cap_format)

        tbl_header = [{'header': 'TC Type\\value'}, {'header': 'value(%)'}]
        manual_total = self.manual_data['Testcases (abs)'][4]
        auto_total = self.each_operator_all_sprint_auto['Testcases (abs)'][4]

        auto = round(auto_total / (manual_total + auto_total), 2)
        auto_manual = {'Automatable': [auto],
                       'Not Automatable': [1 - auto]}
        self.add_table_by_json(auto_manual, table_start_row,
                               table_start_col, True, tbl_header, True)

        cell_index = 1
        chart_values = self.get_chart_excel_column(
                        sheet_name, table_start_col + cell_index,
                        table_start_row + 1, len(auto_manual))

        chart_category = self.get_chart_excel_column(sheet_name,
                                                     table_start_col,
                                                     table_start_row + 1,
                                                     len(auto_manual))
        char_cell = self.get_excel_cell(chart_col, chart_row)
        chart_color = [
            {'fill': {'color': '#ffff99'}},
            {'fill': {'color': '#66ccff'}},
        ]
        title = 'Automatable({}) vs Not automatble({}) TCs'.format(
            self.each_operator_all_sprint_auto['Testcases (abs)'][4],
            self.manual_data['Testcases (abs)'][4])
        self.add_pie_chart(chart_category, chart_values,
                           title, char_cell, chart_color)
        chart_col += 6

        # overall summary =======================================
        table_start_col += len(tbl_header) + 1
        self.write_caption(self.worksheet, table_start_row - 1,
                           table_start_col,
                           'Overall TCs status', self.cap_format)
        tbl_header, summary_dic = self.get_header_json(
            self.overall_summary)
        self.add_table_by_json(summary_dic, table_start_row,
                               table_start_col, True,
                               tbl_header, True)

        # chart =======================================
        auto_start_col = 1
        auto_start_row = 2
        chart_values, chart_category = self.get_pie_chart_value_category(
            'Automation_stats', auto_start_col,
            auto_start_row, len(summary_dic) - 2,
            cell_index)

        char_cell = self.get_excel_cell(chart_col, chart_row)
        chart_color = [
            {'fill': {'color': '#80ff80'}},
            {'fill': {'color': '#ccff33'}},
            {'fill': {'color': '#ffff33'}},
            {'fill': {'color': '#ff9980'}}
        ]
        title = 'Automation status (completed, backlog, etc)'
        self.add_pie_chart(chart_category, chart_values,
                           title, char_cell, chart_color)
        chart_col += 7
        # bvt/fvt/svt overall summary =======================================
        table_start_col += len(tbl_header) + 1
        tbl_header = [{'header': 'TC Type\\value'},
                      {'header': 'Completed'}, {'header': 'Total'}]
        self.write_caption(self.worksheet, table_start_row - 1,
                           table_start_col,
                           "BVT/FVT/SVT Summary",
                           self.cap_format)
        self.add_table_by_json(self.sumry_bvt_fvt_svt, table_start_row,
                               table_start_col, True, tbl_header)
        # column chart
        chart_category, names_values = \
            self.get_column_chart_category_values_names(
                sheet_name, table_start_col, table_start_row,
                len(self.sumry_bvt_fvt_svt), [1, 2])
        names_values['fill_color'] = [{'color': '#00334d'},
                                      {'color': '#80d4ff'}]
        char_cell = self.get_excel_cell(chart_col, chart_row)
        self.add_column_chart(chart_category, names_values, 'BVT/FVT/SVT '
                                                            'Summary',
                              char_cell, 'Test category', 'No. of '
                                                          'testcases')
        # sprint level summary =======================================
        table_start_col += len(tbl_header) + 1
        self.write_caption(self.worksheet, table_start_row - 1,
                           table_start_col,
                           "Sprint specific automation summary",
                           self.cap_format)

        self.add_table_by_json(self.all_operator_sprint_auto,
                               table_start_row,
                               table_start_col)

        chart_col = 1
        chart_row += 15
        # column chart
        chart_category, column_names_values = \
            self.get_column_chart_category_values_names(
                sheet_name, table_start_col, table_start_row,
                len(self.all_operator_sprint_auto), [5])
        chart_category, line_names_values = \
            self.get_column_chart_category_values_names(
                sheet_name, table_start_col, table_start_row,
                len(self.all_operator_sprint_auto), [1])
        column_names_values['fill_color'] = [{'color': '#3399ff'}]
        line_names_values['fill_color'] = [{'color': '#b30000'}]
        char_cell = self.get_excel_cell(chart_col, chart_row)
        size = {'width': 550, 'height': 300}
        self.add_column_line_chart(
            chart_category, column_names_values,
            line_names_values, 'sprint wise automation progress',
            char_cell,
            'sprints...', 'No. of testcases', size)

        # manual sprint summary
        table_start_col += len(tbl_header) + 4
        self.write_caption(self.worksheet, table_start_row - 1,
                           table_start_col,
                           "Sprint specific manual summary",
                           self.cap_format)

        self.add_table_by_json(
            self.all_operator_sprint_data,
            table_start_row,
            table_start_col)

        # column chart for all operators
        chart_col = 8

        chart_category = ['Automation_stats', 1, 4, 1, 15]
        names_values = {}
        names_values['name'] = [['Automation_stats', 2, 1],
                                ['Automation_stats', 6, 1]]
        names_values['value'] = [['Automation_stats', 2, 4, 2, 15],
                                 ['Automation_stats', 6, 4, 6, 15]]
        names_values['fill_color'] = [{'color': '#006600'},
                                      {'color': '#e62e00'}]
        size = {'width': 900, 'height': 300}
        char_cell = self.get_excel_cell(chart_col, chart_row)
        self.add_column_chart(chart_category, names_values, 'Operators: '
                                                            'completed vs all',
                              char_cell,
                              'Name of operators', 'No. of testcases', size)

    def other_tabs(self, data, operator_data, is_automation_tab=False):
        table_start_col = 1
        table_start_row = 2
        # chart_row = 3
        # chart_col = 1
        self.write_caption(self.worksheet, table_start_row - 1,
                           table_start_col,
                           'Operator level automation summary', self.cap_format)
        tbl_header, summary_dic = self.get_header_json(data)
        self.add_table_by_json(summary_dic, table_start_row,
                               table_start_col, True,
                               tbl_header)

        table_start_row += len(summary_dic) + 2

        row = table_start_row + 2
        col = 4

        self.write_caption(self.worksheet, table_start_row,
                           col,
                           'Operator specific automation TCs summary',
                           self.cap_format)

        custom_header = [{'header': 'sprint\\type of TCs'}]
        for colmn in self.tc_types:
            custom_header.append({'header': colmn})
        for colmn in self.tc_types[:-1]:
            colmn += '(TC id)'
            custom_header.append({'header': colmn})

        for key, value in operator_data.items():
            count_dic = {}
            tc_dic = {}

            format_color = None
            if is_automation_tab:
                if value['BVT']['status']:
                    format_color = self.workbook.add_format(
                        {'bold': True,
                         'font_size': font_size,
                         'font_color': '#00cc00'})
                else:
                    format_color = self.workbook.add_format(
                        {'bold': True,
                         'font_size': font_size,
                         'font_color': '#ff0000'})
                self.write_caption(self.worksheet, row - 1, col + 3, 'BVT',
                                   format_color)

            for sprint, tc in value.items():
                if sprint == 'BVT':
                    continue
                count_dic[sprint] = tc['TC count'] + tc['TC IDs']
                tc_dic[sprint] = tc['TC IDs']

                tc_list = [item for item in tc['TC IDs'] if item != '']
                if not len(tc_list):
                    continue
                tc_str = ' '.join(tc_list)
                tc_list = tc_str.split(' ')
                if not is_automation_tab:
                    self.tc_maual_ids.extend(tc_list)
                else:
                    self.tc_auto_ids.extend(tc_list)

            self.write_caption(self.worksheet, row - 1, col, key,
                               self.cap_format)

            self.add_table_by_json(count_dic, row, col, True, custom_header)

            # self.add_table_by_json(tc_dic, row, col + 7, False)
            row = row + len(count_dic) + 3
        #     # return

    # *************************************************************************
    def add_pie_chart(self, category, values, title, char_cell,
                      chart_format=[]):
        pie_chart = self.workbook.add_chart({'type': 'pie'})
        pie_chart.add_series({
            'categories': category,
            'values': values,
            'points': chart_format
        })
        pie_chart.set_title({'name': title})
        # Set an Excel chart style. Colors with white outline and shadow.
        pie_chart.set_style(10)
        # Insert the chart into the worksheet (with an offset).
        self.worksheet.insert_chart(char_cell, pie_chart,
                                    {'x_offset': 0, 'y_offset': 0})

    def _add_column_chart(self, category, name_values,
                          title, x_name, y_name, size={}):
        column_chart = self.workbook.add_chart({'type': 'column'})
        for index in range(len(name_values['name'])):
            column_chart.add_series({
                'name': name_values['name'][index],
                'categories': category,
                'values': name_values['value'][index],
                'fill': name_values['fill_color'][index]
            })
        column_chart.set_title({'name': title})
        column_chart.set_x_axis({'name': x_name})
        column_chart.set_y_axis({'name': y_name})

        if size:
            column_chart.set_size(size)
        # Set an Excel chart style. Colors with white outline and shadow.
        column_chart.set_style(11)
        return column_chart

    def _add_line_chart(self, category, name_values):
        line_chart_list = []
        for index in range(len(name_values['name'])):
            line_chart = self.workbook.add_chart({'type': 'line'})
            line_chart.add_series({
                'name': name_values['name'][index],
                'categories': category,
                'values': name_values['value'][index],
                'line': name_values['fill_color'][index]

            })
            line_chart_list.append(line_chart)

        return line_chart_list

    def add_column_chart(self, category, name_values, title, char_cell,
                         x_name='', y_name='', size={}):
        column_chart = self._add_column_chart(category, name_values, title,
                                              x_name, y_name, size)

        # Insert the chart into the worksheet (with an offset).
        self.worksheet.insert_chart(char_cell, column_chart,
                                    {'x_offset': 0, 'y_offset': 0})

    def add_column_line_chart(self, category, bar_name_values,
                              line_name_value, title, char_cell,
                              x_name='', y_name='', size={}):
        column_chart = self._add_column_chart(category, bar_name_values, title,
                                              x_name, y_name, size)
        line_chart_list = self._add_line_chart(category, line_name_value)

        for line_chart in line_chart_list:
            column_chart.combine(line_chart)

        # Insert the chart into the worksheet (with an offset).
        self.worksheet.insert_chart(char_cell, column_chart,
                                    {'x_offset': 0, 'y_offset': 0})

    def add_pie_chart1(self, start_row, start_col, max_row, percent_columns={},
                       sheet_name=''):
        # start pie chart calculation
        cat_col = chr(start_col + 65)
        category = '=' + sheet_name + '!$' + cat_col + '$' + str(
            start_row + 2) + ':$' + cat_col + '$' + str(start_row + max_row)

        for col_char, col in percent_columns.items():
            # Configure the series. Note the use of the list syntax to
            # define ranges:
            title = 'Chart of ' + col
            values = '=' + sheet_name + '!$' + col_char + '$' + str(
                start_row + 2) + ':$' + col_char + '$' + str(
                start_row + max_row)

        pie_chart = self.workbook.add_chart({'type': 'pie'})
        pie_chart.add_series({
            'name': title,
            'categories': category,
            'values': values,
        })

        # Set an Excel chart style. Colors with white outline and shadow.
        pie_chart.set_style(10)

        char_row = 'B' + str(max_row + start_row + 5)
        # Insert the chart into the worksheet (with an offset).
        self.worksheet.insert_chart(char_row, pie_chart,
                                    {'x_offset': 0, 'y_offset': 0})

    # *************************************************************************
    def get_header_json(self, df_json={}):
        header = []
        dict = {}
        ret_header = []
        for key, value in df_json.items():
            ret_header.append({'header': key})
            header.append(key)
        index = 0
        for item in df_json[header[0]]:
            row = []
            for head in header[1:]:
                row.append(df_json[head][index])
            index += 1
            dict[item] = row
        return ret_header, dict

    def write_caption(self, worksheet, row, col, comment, format):
        cell_id = self.get_excel_cell(col, row)
        worksheet.write(cell_id, comment, format)

    def get_excel_cell(self, col, row):
        return (str(chr(65 + col) + str(row)))

    def get_chart_excel_column(self, sheet_name, col, start_row, end_row=0):
        output = ''
        if end_row:
            output = ('=' + sheet_name + '!$' + chr(65 + col) + '$' + str(
                start_row) + ':' + '$' + chr(65 + col) + '$' + str(
                start_row + end_row - 1))
        else:
            output = ('=' + sheet_name + '!$' + chr(65 + col) + '$' + str(
                start_row))
        return output

    def get_pie_chart_value_category(self, sheet_name, col, start_row,
                                     num_rows,
                                     column_index):
        chart_values = self.get_chart_excel_column(sheet_name,
                                                   col + column_index,
                                                   start_row + 1, num_rows)
        chart_category = self.get_chart_excel_column(sheet_name, col,
                                                     start_row + 1, num_rows)
        return chart_values, chart_category

    def get_column_chart_category_values_names(self, sheet_name, col,
                                               start_row,
                                               num_rows, column_indexes=[]):
        chart_category = self.get_chart_excel_column(sheet_name, col,
                                                     start_row + 1, num_rows)
        chart_values = []
        chart_names = []
        for column_index in column_indexes:
            chart_values.append(self.get_chart_excel_column(sheet_name,
                                                            col + column_index,
                                                            start_row + 1,
                                                            num_rows))
            chart_names.append(self.get_chart_excel_column(sheet_name,
                                                           col + column_index,
                                                           start_row))
        return chart_category, {'name': chart_names, 'value': chart_values}

    def add_table_by_json(self, json_data, table_start_row, table_start_col,
                          add_fist_last_row=True, header=[],
                          percetage=False):

        # Add the Excel table structure. Pandas will add the data.
        json_data = dict(natsorted(json_data.items()))
        columns = header
        data = []
        if not header:
            if add_fist_last_row:
                columns = [{'header': 'sprint\\type of TCs'}]
            data = []
            for col in self.tc_types:
                if add_fist_last_row is False and col == 'Total testcase':
                    continue
                columns.append({'header': col})

        for key, value in json_data.items():
            if add_fist_last_row:
                row_list = [key] + value
            else:
                row_list = value
            data.append(row_list)

        # Get the dimensions of the dataframe.
        column = chr(65 + table_start_col)
        start_row_col = column + str(table_start_row)
        column = chr(65 + table_start_col + len(columns) - 1)
        end_row_col = column + str(len(json_data) + table_start_row)

        start_index = table_start_col + 65
        index = 0
        format_prcntg = self.workbook.add_format({'num_format': '0%'})
        format_numbr = self.workbook.add_format({'num_format': '00'})

        for each_header in header:
            n = index + start_index
            index += 1
            excel_col = chr(n) + ':' + chr(n)
            col = each_header['header']
            if 'percent' in col or '%' in col:
                # format_prcntg.set_text_wrap(True)
                format_prcntg.set_font_size(font_size)
                self.worksheet.set_column(excel_col, 12, format_prcntg)
            else:
                # format_numbr.set_text_wrap(True)
                format_numbr.set_font_size(font_size)
                self.worksheet.set_column(excel_col, 12, format_numbr)

        excel_col = start_row_col + ':' + end_row_col

        # self.worksheet.set_column(excel_col, 12, format_col)

        self.worksheet.add_table(excel_col, {'data': data, 'columns': columns,
                                             'autofilter': 0})

        return None

    def sum_dict_lists(self, dic):
        sum_list = [0, 0, 0, 0, 0]
        for op_stats in dic.values():
            sum_list = [sum(i) for i in zip(sum_list, op_stats)]
        percent_list = []
        n = len(sum_list)
        for all_op_stat in sum_list[:n]:
            if sum_list[n - 1]:
                percent_list.append(all_op_stat / sum_list[n - 1])
            else:
                percent_list.append(0)

        return sum_list, percent_list

    def generate_sprint_data(self, df, automation_only=False):
        print('========================sprint data=========================')
        self.each_operator_all_sprint_data = {}
        self.all_operator_sprint_data = {}
        for key, value in self.operator_sprint.items():
            sprint_dic = {}
            all_sprint_count = [0, 0, 0, 0, 0]
            bvt_status = 0

            for sprint, lables in value.items():
                self.status_based_dic['TC count'] = [0, 0, 0, 0, 0]
                self.status_based_dic['TC IDs'] = ['', '', '', '', '']

                str_lable = self.get_search_string(lables)
                if not len(str_lable):
                    sprint_dic[sprint] = copy.deepcopy(self.status_based_dic)
                    continue
                temp_df = df[df[self.columns.index('MODULE')].str.contains(
                    str_lable)]
                self.generate_status_based_data(temp_df)
                temp_dic = copy.deepcopy(self.status_based_dic)
                sprint_dic[sprint] = temp_dic
                print('key {} sprint {}, lables: {} returned df {}'.format(
                    key, sprint, lables, df))
                all_sprint_count = [sum(i) for i in zip(all_sprint_count,
                                                        temp_dic['TC count'])]
                # if 'sprintn' in sprint.lower() or not automation_only:
                #     continue
                if not automation_only:
                    continue
                # bvt check
                if not bvt_status:
                    bvt_status = self.is_bvt_done(temp_df)

                run_dic = {'BVT': [0, 0], 'FVT': [0, 0], 'SVT': [0, 0]}
                if sprint in self.all_operator_sprint_data:
                    self.all_operator_sprint_data[sprint] = \
                        [sum(i) for i in
                         zip(self.all_operator_sprint_data[sprint],
                             temp_dic['TC count'])]
                else:
                    self.all_operator_sprint_data[sprint] = temp_dic[
                        'TC count']
                    self.sprint_bvt_fvt_svt[sprint] = run_dic

                for k, val in self.sprint_bvt_fvt_svt[sprint].items():
                    run_cnt = self.get_test_type_count(temp_df, k)
                    run_dic[k] = [sum(i) for i in zip(val, run_cnt)]
                    self.sumry_bvt_fvt_svt[k] = [sum(i) for i in zip(
                        run_cnt, self.sumry_bvt_fvt_svt[k])]
                    print('')

                self.sprint_bvt_fvt_svt[sprint] = run_dic

            bvt_status = True if bvt_status else False
            sprint_dic['BVT'] = {'status': bvt_status}
            self.each_operator_sprint_values[key] = sprint_dic
            self.each_operator_all_sprint_data[
                key.split(' ', 1)[1]] = all_sprint_count

    def get_search_string(self, lables=[]):
        op_str = ''
        for lable in lables:
            op_str += lable + ' |'
        op_str = op_str[:-1]
        return op_str

    def generate_overall_summary(self):
        '''       self.overall_summary = {'Testcases':['Automation completed',
                                             'Automation in progress',
                                             'Automation in review',
                                             'Automation backlog',
                                             'Manual completed'],
        '''
        print('============overall Summary=============')
        prcnt_list = []
        total_tc = self.each_operator_all_sprint_auto['Testcases (abs)'][
                       4] + self.manual_data['Testcases (abs)'][0]
        i = 0
        while i < 4:
            prcnt_list.append(self.each_operator_all_sprint_auto[
                                  'Testcases (abs)'][i] / total_tc)
            i += 1
        prcnt_list.append(self.manual_data['Testcases (abs)'][0] / total_tc)
        prcnt_list.append(total_tc / total_tc)

        self.overall_summary['values (%)'] = prcnt_list

    def generate_overall_data(self, df):
        print('====================Overall Data=====================')
        abs_list = self.generate_status_based_data(df)

        abs_list[3] = abs_list[3] - self.manual_new
        self.op_list_percentage[3] = abs_list[3] / self.all_exec
        abs_list[4] = abs_list[
                          4] - self.manual_new - 1  # -1 to avoid the header
        self.op_list_percentage[4] = abs_list[4] / self.all_exec

        self.op_df['Overall percentage'] = self.op_list_percentage
        self.op_df['Overall testcases'] = abs_list

    def generate_automation_data(self):
        print('====================Automation Data=====================')
        abs_list = self.generate_status_based_data(self.auto_tc)
        self.op_df['Automation testcases'] = abs_list
        self.op_df['Automation percentage'] = self.op_list_percentage
        # self.each_operator_all_sprint_auto['Testcases (abs)'] = abs_list
        #
        # self.each_operator_all_sprint_auto[
        #     'Testcases (%)'] = self.get_percent_list(
        #     abs_list)

    def get_percent_list(self, abs_list):
        total = abs_list[len(abs_list) - 1]
        percent_list = []
        for ele in abs_list:
            percent_list.append(ele / total)
        return percent_list

    def generate_manual_data(self):
        print('====================Manual Data=====================')
        abs_list = self.generate_status_based_data(self.manual_tc)

        # abs_list[4] = abs_list[4] - self.manual_new
        # self.op_list_percentage[4] = abs_list[4] / self.all_exec

        self.op_df['Manual testcases'] = abs_list
        self.op_df['Manual percentage'] = self.op_list_percentage
        # self.manual_data['Testcase (abs)'] = abs_list
        # self.manual_data['Testcase (%)'] = self.get_percent_list(abs_list)

    def generate_status_based_data(self, df):
        try:
            op_list = []
            self.op_list_percentage = []
            self.status_tcs = []
            num = self.get_count_by_column_value(df, self.columns.index(
                'STATUS'), 'Baselined')
            op_list.append(num)
            self.op_list_percentage.append((num / self.all_exec))
            # self.op_list_percentage.append(str(
            #     '{}%'.format(round(num*100/self.all_exec),4)))

            num = self.get_count_by_column_value(df, self.columns.index(
                'STATUS'), 'Ready For Baseline')
            op_list.append(num)
            self.op_list_percentage.append((num / self.all_exec))
            # self.op_list_percentage.append(str(
            #     '{}%'.format(round(num*100/self.all_exec),4)))

            num = self.get_count_by_column_value(df, self.columns.index(
                'STATUS'), 'In Progress')
            op_list.append(num)
            self.op_list_percentage.append((num / self.all_exec))
            # self.op_list_percentage.append(str(
            #     '{}%'.format(round(num*100/self.all_exec),4)))

            num = self.get_count_by_column_value(df, self.columns.index(
                'STATUS'), 'New')
            op_list.append(num)
            self.op_list_percentage.append((num / self.all_exec))
            # self.op_list_percentage.append(str(
            #     '{}%'.format(round(num*100/self.all_exec),4)))

            num = len(df)
            op_list.append(num)
            self.op_list_percentage.append((num / self.all_exec))
            # self.op_list_percentage.append(str(
            #     '{}%'.format((num*100/self.all_exec),4)))
            print('Total No. of testcase: {}'.format(len(df)))

            self.status_based_dic['TC count'] = op_list
            self.status_tcs.append('')  # appending dummy for total
            self.status_based_dic['TC IDs'] = self.status_tcs

            return op_list
        except Exception as e:
            print(str(e))

    def get_percentage_value(self):
        return self.op_list_percentage

    def get_all_executed_tc(self):
        self.all_exec = self.get_count_by_column_value(
            self.manual_tc,
            self.columns.index('STATUS'), 'Baselined') + \
                        len(self.auto_tc)

        print('Number of executed TCs:', self.all_exec)
        return self.all_exec

    def get_count_by_column_value(self, df, columen=1, value=''):
        temp_df = df[df[columen] == value]
        count = len(temp_df)
        tc_ids = ' '.join(temp_df[self.columns.index(
            'ID')].to_list())
        self.status_tcs.append(tc_ids)
        print('Total No. of {} testcase: {}'.format(value, count))
        return count

    def is_bvt_done(self, df):
        temp_df = df[(df[self.columns.index('STATUS')] == 'Baselined') &
                     (df[self.columns.index('TESTPHASE')] == 'Unit Test')]
        return len(temp_df)

    def get_test_type_count(self, df, type='BVT'):
        # total count, baselined count
        count = []
        temp_df = df[(df[self.columns.index('STATUS')] == 'Baselined') &
                     (df[self.columns.index('MODULE')].str.contains(type))]
        count.append(len(temp_df))
        temp_df = df[(df[self.columns.index('MODULE')].str.contains(type))]
        count.append(len(temp_df))
        return count

    def get_bvt_counts(self, df):
        return self.get_test_type_count(df, 'BVT')

    def get_fvt_counts(self, df):
        return self.get_test_type_count(df, 'FVT')

    def get_svt_counts(self, df):
        return self.get_test_type_count(df, 'SVT')

    def get_table_json(self, df):
        # df.dropna(inplace=True)
        df.fillna(True, inplace=True)
        df.reset_index(drop=True, inplace=True)
        data = df.transpose()
        data = data.to_dict('list')
        return data

    # added at last moment where we need to merge two report's data
    def merged_data_in_excel(self, excel_file='', sheets_df={}, other_report_obj=None):
        try:

            file_name, ext = os.path.splitext(excel_file)
            op_file = file_name + '_Summary' + ext

            if os.path.exists(op_file):
                os.remove(op_file)

            # Create a Pandas Excel writer using XlsxWriter as the engine.
            # writer = pd.ExcelWriter(op_file, engine='xlsxwriter')
            self.workbook = xlsxwriter.Workbook(op_file)
            self.cap_format = self.workbook.add_format(
                {'bold': True, 'font_size': font_size,
                 'font_color': '#33001a'})
            overview_sheet = None
            header = []
            # list of dic of headers in config.ini
            for col in self.columns:
                header.append({'header': col})

            # add sheets and data init in workbook
            for sheet, data in sheets_df.items():
                # Create a Pandas dataframe from some data.

                # Write the dataframe data to XlsxWriter.
                # Turn off the default header and
                # index and skip one row to allow us to insert
                # a user defined header.
                self.worksheet = self.workbook.add_worksheet(sheet)

                if 'Overall_summary' in sheet:
                    overview_sheet = self.worksheet
                elif 'Automation_stats' in sheet:
                    self.other_tabs(data,
                                    self.each_operator_sprint_values_auto,
                                    True)
                elif 'Manual_stats' in sheet:
                    self.other_tabs(data, self.each_operator_sprint_values)
                else:
                    self.add_table_by_json(data, 1, 1, False,
                                           header)

            # writer.save()
            self.worksheet = overview_sheet
            self.generate_sprint_data(self.auto_tc, True)
            self.overview_tab('Overall_summary')

            self.workbook.close()
        except Exception as e:
            print('Exception found %s', str(e))

    def _list_sum(self, list1, list2):
        ret_list = []
        for i, (val1, val2) in enumerate(zip(list1, list2)):
            ret_list.append((val1+val2))
        return ret_list


#============================================================================
import numpy as np
class combined_reports:
    def __init__(self):
        self.reports = []

    def generate_report_to_excel(self, product, data):
        obj_report = ExcelReport()
        other_product = ""
        if product == "HCI":
            other_product = "SDS"
        elif product == "SDS":
            other_product = "HCI"
        else:
            print("product value is not correct")
            return

        final_result = {}
        anayzed_data = obj_report.generat_detailed_report(
                                        data[product]['Excel_path'],
                                        data['Columns'],
                                        data[product]['Operators'],
                                        data[product].get('TestIds'))
        if data[product]['Common_TC_path']:
            obj_report1 = ExcelReport()
            anayzed_common = obj_report1.generat_detailed_report(
                                        data[product]['Common_TC_path'],
                                        data['Columns'],
                                        data[other_product]['Operators'],
                                        data[other_product].get('TestIds'))
            final_result = self.merge_reports(anayzed_data, anayzed_common)
            obj_report.each_operator_all_sprint_auto = final_result["Automation_stats"]
            obj_report.manual_data = final_result["Manual_stats"]
            obj_report.data_in_excel(data[product]['Excel_path'], final_result, obj_report1)
        else:
            final_result = anayzed_data
            obj_report.data_in_excel(data[product]['Excel_path'], final_result)
        print ("common TCs are there")

    def merge_reports(self, product_report, common_report):
        merged_dict = copy.deepcopy(product_report)
        for key, value in product_report.items():
            if key == "Overall_summary":
                # lists = np.array([value['values (%)'], common_report[key]['values (%)']])
                merged_dict[key]['values (%)'] = \
                    self._list_average(value['values (%)'], common_report[key]['values (%)'])
            elif key == "Automation_stats" or key == "Manual_stats":
                for val_key, val_value in value.items():
                    if val_key == 'Type of Testcases':
                        continue

                    if val_key.find('%') >= 0:
                        # lists = np.array([val_value, common_report[key][val_key]])
                        merged_dict[key][val_key] = self._list_average(val_value, common_report[key][val_key])
                    elif val_key in common_report[key]:
                        merged_dict[key][val_key] = self._list_sum(val_value, common_report[key][val_key])
        return merged_dict

    def _list_average(self, list1, list2):
        ret_list = []
        for i, (val1, val2) in enumerate(zip(list1, list2)):
            ret_list.append((val1+val2)/2)
        return ret_list

    def _list_sum(self, list1, list2):
        ret_list = []
        for i, (val1, val2) in enumerate(zip(list1, list2)):
            ret_list.append((val1+val2))
        return ret_list

    # def udpate_data(self, obj1):
    #     obj1.



    # 'Overall_summary': self.overall_summary,
                # 'Automation_stats': self.each_operator_all_sprint_auto,
                # 'Manual_stats': self.manual_data,
                # 'Not_eligible_TCs': self.not_eligible,
                # 'Invalid_format_TCs': self.invalid_tcs,
                # 'Automatable_TCs': auto_tcs,
                # 'Manual_TCs': manual_tcs}