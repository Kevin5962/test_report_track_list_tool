'''
Description: get track list information from test case report
author: fengkai.xiao@cerence.com
date: 2021-05-05
'''

import os
import sys
import xlrd
import xlsxwriter

class TrackList(object):
    def __init__(self,tr_name):
        # self.test_report_root = outputPath
        self.current_root = os.path.dirname(os.path.abspath(__file__))
        self.test_report_name = tr_name
        # BMW_IDC_Function_Test_Summary_MNC_2022_05_05_14_04_58_易用性.xls
        # self.track_template_name = track_template_name
        # 0427版本-IDC23_Tickets_Tracking_易用性.xls
        self.test_report_path =  self.current_root + '\\' + self.test_report_name
        # self.track_template_path =  self.current_root + '\\' + self.track_template_name
        # self.track_list_root = 'C:\Project\IDC_Files\\track_list'
        self.track_list_file =  self.current_root + '\\' + self.test_report_name
        self.track_list_path =  self.test_report_path.replace('.xls','_track_list.xls')
        # print(self.test_report_path)
        # print(self.track_list_path)
        self.width_dict = {'A:A': 9, 'B:B': 9.5, 'C:C': 4.25, 'D:D': 6.25, 'E:E': 15.13, 'F:F': 6.63, 'G:G': 4.5, 'H:H': 20.63,
                           'I:I': 18.25,'J:J': 12.5, 'K:K': 14, 'L:L': 3.13, 'M:M': 3.13, 'N:N': 3.13, 'O:O': 3.13, 'P:P': 3.13, 'Q:Q': 18.25,
                           'R:R': 8.63, 'S:S': 8.63, 'T:T': 8.63, 'U:U': 8.63, 'V:V': 3.25,'W:W': 3.25, 'X:X': 3.25, 'Y:Y': 3.25, 'Z:Z': 18.25, 
                           'AA:AA': 8.63, 'AB:AB': 8.63, 'AC:AC': 8.63, 'AD:AD': 4.38,'AE:AE': 4.38, 'AF:AF': 4.38, 'AG:AG': 4.38, 'AH:AH': 4.38,
                            'AI:AI': 4.38,'AJ:AJ': 17.63, 'AK:AK': 17.63, 'AL:AL': 5, 'AM:AM': 5, 'AN:AN': 5, 'AO:AO': 5, 'AP:AP': 5,
                           'AQ:AQ': 5, 'AR:AR': 5, 'AS:AS': 5, 'AT:AT': 5, 'AU:AU': 15, 'AV:AV': 15}
        
        self.tr_workbook = xlrd.open_workbook(self.test_report_path,formatting_info = True)
        self.tr_all_sheet = self.tr_workbook.sheet_by_name("All")
        # self.track_template_workbook = xlrd.open_workbook(self.track_template_path,formatting_info = True)

    def create_title_format(self,wb):
        total_title_color_index_dict = {'ASR-Embedded': {'008000': [0, 1], 'FF6600': [2, 3, 4, 5], '333399': []}, 'ASR-Cloud': {'008000': [0, 1], 'FF6600': [2, 3, 4, 5], '333399': []}, 
        'onenlu-Embedded': {'008000': [0, 1, 2, 3, 6, 7, 8, 9], 'FF6600': [10, 11, 12, 13], '333399': [4, 5]}, 
        'onenlu-Cloud': {'008000': [0, 1, 2, 3, 6, 7, 8, 9], 'FF6600': [10, 11, 12, 13, 14], '333399': [4, 5]}, 
        'mapping': {'008000': [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27], 'FF6600': [], '333399': []}, 
        'ASR is empty': {'008000': [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28], 'FF6600': [], '333399': []}}

        title_dict_0 = {'border': True, 'bold': True, 'align': 'left', 'valign': 'vcenter',"bg_color": '008000', "font_color": 'FFFFFF',
                    'font_size': '11', 'font_name': 'Arial'}
        title_dict_1 = {'border': True, 'bold': True, 'align': 'left', 'valign': 'vcenter',"bg_color": 'FF6600', "font_color": 'FFFFFF',
                    'font_size': '11', 'font_name': 'Arial'}
        title_dict_2 = {'border': True, 'bold': True, 'align': 'left', 'valign': 'vcenter',"bg_color": '333399', "font_color": 'FFFFFF',
                    'font_size': '11', 'font_name': 'Arial'}
        title_color_dict = {}
        title_format_0 = wb.add_format(title_dict_0)
        title_format_1 = wb.add_format(title_dict_1)
        title_format_2 = wb.add_format(title_dict_2)
        title_color_dict['008000'] = title_format_0
        title_color_dict['FF6600'] = title_format_1
        title_color_dict['333399'] = title_format_2
        
        return total_title_color_index_dict,title_color_dict

    def get_cell_bg_color(self, wr, sheet, row_index, col_index): 
        """ 获取某一个单元格的背景颜色 :param wb: :param sheet: :param row_index: :param col_index: :return: """
        sheet_hex_color_dict = {'black':'000000','white':'FFFFFF','red':'FF0000','yellow':'FFFF00','lime':'00FF00'} 
        hex_font_color, hex_bg_color = 'FFFFFF','000000'

        #RGB2Hex
        def rgb2hex(rgb):
            r, g, b = rgb
            return ('{:02X}'*3).format(r, g, b)
        xfx = sheet.cell_xf_index(row_index, col_index) 
        xf = wr.xf_list[xfx] 
        # 字体颜色索引 
        font_color_index = wr.font_list[xf.font_index].colour_index 
        # 背景颜色索引 
        bg_color_index = xf.background.pattern_colour_index 
        # 索引映射到RGB 
        font_rgb_color = wr.colour_map.get(font_color_index)
        bg_rgb_color = wr.colour_map.get(bg_color_index)
        if font_rgb_color != None:
            hex_font_color = rgb2hex(font_rgb_color)
        if bg_rgb_color != None:
            hex_bg_color = rgb2hex(bg_rgb_color)

        return [(hex_font_color, hex_bg_color),(font_rgb_color,bg_rgb_color)] 
        # return font_color_index, bg_color_index
        # return font_rgb_color, bg_rgb_color

    def get_error_sheet_title_format(self,wb,workbook):  
        title_width_dict = {}
        title_color_list = []
        workbook = self.track_template_workbook
        workbook_sheet = workbook.sheet_by_name('onenlu-Cloud')
        for i in range(workbook_sheet.ncols):
            i_font_bg_colors = self.get_cell_bg_color(workbook,workbook_sheet, 0, i)
            title_color_list.append(i_font_bg_colors[0][1])
            # print(i_font_bg_colors[0][1],i_font_bg_colors[1][1])
        title_color_list = list(set(title_color_list))
        # print(title_color_list)

        total_title_color_index_dict = {}
        title_dict_0 = title_dict_1 = title_dict_2 = {}
        for i_sheet in workbook.sheet_names():
            if 'Summary' in i_sheet:
                continue
            else:
                # print(i_sheet)
                workbook_sheet = workbook.sheet_by_name(i_sheet)
                title_color_index_dict = {}
                title_index_0 = []
                title_index_1 = []
                title_index_2 = []
                title_width_list = []
                for i in range(workbook_sheet.ncols):
                    i_font_bg_colors = self.get_cell_bg_color(workbook,workbook_sheet, 0, i)
                    # print('i_font_bg_colors:%s' %(i_font_bg_colors))
                    if i_font_bg_colors[0][1] == title_color_list[0]:
                        title_index_0.append(i)
                    elif i_font_bg_colors[0][1] == title_color_list[1]:
                        title_index_1.append(i)
                    elif i_font_bg_colors[0][1] == title_color_list[2]:
                        title_index_2.append(i)

                title_color_index_dict[title_color_list[0]] = title_index_0
                title_color_index_dict[title_color_list[1]] = title_index_1
                title_color_index_dict[title_color_list[2]] = title_index_2
        
            total_title_color_index_dict[i_sheet] = title_color_index_dict
            title_width_dict[i_sheet] = title_width_list


        title_dict_0 = {'border': True, 'bold': True, 'align': 'left', 'valign': 'vcenter',"bg_color": title_color_list[0], "font_color": 'FFFFFF',
                    'font_size': '11', 'font_name': 'Arial'}
        title_dict_1 = {'border': True, 'bold': True, 'align': 'left', 'valign': 'vcenter',"bg_color": title_color_list[1], "font_color": 'FFFFFF',
                    'font_size': '11', 'font_name': 'Arial'}
        title_dict_2 = {'border': True, 'bold': True, 'align': 'left', 'valign': 'vcenter',"bg_color": title_color_list[2], "font_color": 'FFFFFF',
                    'font_size': '11', 'font_name': 'Arial'}
        title_color_dict = {}
        title_format_0 = wb.add_format(title_dict_0)
        title_format_1 = wb.add_format(title_dict_1)
        title_format_2 = wb.add_format(title_dict_2)
        title_color_dict[title_color_list[0]] = title_format_0
        title_color_dict[title_color_list[1]] = title_format_1
        title_color_dict[title_color_list[2]] = title_format_2

        return total_title_color_index_dict,title_color_dict,title_width_dict

    def get_cell_color_index_list(self,value,cell_list,color_format):
        if cell_list == None or len(cell_list)<2:
            segments_list = []
        else:
            color_list = []
            for i in range(len(cell_list)-1):
                color_list.append([cell_list[i][1],(cell_list[i][0],cell_list[i+1][0])])
            color_list = [[0,(0,cell_list[0][0])]] + color_list + [[0,(cell_list[-1][0],len(value))]]
            # print(color_list)
            
            segments_list = []
            for color in color_list:
                b,a = color[1][1],color[1][0]
                if color[0] == 0:
                    segments_list.append(value[a:b])
                else:
                    segments_list.append(color_format)
                    # segments_list.append('red')
                    segments_list.append(value[a:b])
        # print(segments_list)
        return segments_list

            
    def get_cell_color_distribution(self,i,rowx,colx,coly,workbook_sheet,wb,worksheet):
        red_str_color = wb.add_format({'bold':True,'color':'red'})
        top_format = wb.add_format({'align': 'top','valign': 'top'})
        segments = []
        if coly == 16:
            value11 = workbook_sheet.cell(i,11).value
            value16 = workbook_sheet.cell(i,16).value
            # print('i=%d,j=16' %i)
            # print('value11:%s' %value11)
            # print('value16:%s' %value16)
            if value11 == "F":
                # print('enter the 11F branch')
                cell_list1 = workbook_sheet.rich_text_runlist_map.get((i, 16)) # 注意是双层括号
                # print(cell_list1)
                segments = self.get_cell_color_index_list(value16,cell_list1,red_str_color)
                if len(segments)>0:
                    worksheet.write_rich_string(rowx, colx, *segments, top_format)
        if coly == 25: 
            value21 = workbook_sheet.cell(i,21).value
            value25 = workbook_sheet.cell(i,25).value
            # print('i=%d,j=25' %i)
            # print('value21:%s' %value21)       
            # print('value25:%s' %value25)
            if value21 == "F":
                # print('enter the 21F branch')
                cell_list2 = workbook_sheet.rich_text_runlist_map.get((i, 25)) # 注意是双层括号
                # print(cell_list2)
                segments = self.get_cell_color_index_list(value25,cell_list2,red_str_color)
                if len(segments)>0:
                    worksheet.write_rich_string(rowx, colx, *segments, top_format)
                    

    def writeSummarySheet(self, wb):
        print('begin to write summary sheet')
        summary_head_format = wb.add_format(
            {'border': True, 'bold': True, 'align': 'center', 'font_color': 'FFFFFF', 'bg_color': '000000',
             'font_size': '11', 'font_name': 'Arial'})
        summary_title_format = wb.add_format(
            {'border': True, 'bold': True, 'align': 'left', 'font_color': 'FFFFFF', 'bg_color': '238E23',
             'font_size': '11', 'font_name': 'Arial'})
        summary_value_format = wb.add_format(
            {'border': True, 'bold': False, 'align': 'left', 'font_color': '000000', 'bg_color': 'FFFFFF',
             'font_size': '11', 'font_name': 'Arial'})
        number_format = wb.add_format(
            {'num_format': '0.00%', 'border': True, 'bold': False, 'align': 'left', 'font_color': '000000',
             'bg_color': 'FFFFFF', 'font_size': '11', 'font_name': 'Arial'})
        not_test_format = wb.add_format(
            {'border': True, 'bold': False, 'align': 'top', "font_color": '000000', "bg_color": 'FFFF00',
             'font_size': '11', 'font_name': 'Arial'})
        not_test_format.set_border_color("FFFFFF")

        ##write data to summary sheet

        width_Sum_dict = {'B:B': 14, 'C:C': 18, 'D:D': 14, 'E:E': 14,'F:F': 14}
        worksheet = wb.add_worksheet('Summary')
        worksheet.set_tab_color('green')

        for w in width_Sum_dict:
            worksheet.set_column(w, width_Sum_dict[w])
        ###############  test result analysis of failed cases ###############
        # table name
        row = 1
        worksheet.merge_range(row, 1, row, 5, "test result analysis of failed cases", summary_head_format)
        # title name of the table
        title = ["owner", "type", "xxxx-version", "xxxx-version", "xxxx-version"]
        row = 2
        col = 1
        for item in title:
            worksheet.write(row, col, item, summary_title_format)
            col += 1

        owner_list = ['xin.zhang','yuanzhi.zhao','yu.dai','tao.yu','jian.zou','jian.zou']
        type_list = ['ASR-Embedded','ASR-Cloud','onenlu-Embedded','onenlu-Cloud','mapping',	'ASR is empty']

        row = 3
        col = 1
        for item in owner_list:
            worksheet.write(row, col, item, summary_value_format)
            row += 1
        row = 3
        col = 2
        for item in type_list:
            worksheet.write(row, col, item, summary_value_format)
            row += 1     
        row = 3
        col = 5
        # worksheet.write(row, col, "=COUNTIF('ASR-Embedded'!A:A,\"<>\")-1",summary_value_format)
        # worksheet.write(row+1, col, "=COUNTIF('ASR-Cloud'!A:A,\"<>\")-1",summary_value_format)
        # worksheet.write(row+2, col, "=COUNTIF('onenlu-Embedded'!A:A,\"<>\")-1",summary_value_format)
        # worksheet.write(row+3, col, "=COUNTIF('onenlu-Cloud'!A:A,\"<>\")-1",summary_value_format)
        # worksheet.write(row+4, col, "=COUNTIF('mapping'!A:A,\"<>\")-1",summary_value_format)
        # worksheet.write(row+5, col, "=COUNTIF('ASR is empty'!A:A,\"<>\")-1",summary_value_format)
        worksheet.write(row, col, "=COUNTIF('ASR-Embedded'!A2:A10000,\"<>\")",summary_value_format)
        worksheet.write(row+1, col, "=COUNTIF('ASR-Cloud'!A2:A10000,\"<>\")",summary_value_format)
        worksheet.write(row+2, col, "=COUNTIF('onenlu-Embedded'!A2:A10000,\"<>\")",summary_value_format)
        worksheet.write(row+3, col, "=COUNTIF('onenlu-Cloud'!A2:A10000,\"<>\")",summary_value_format)
        worksheet.write(row+4, col, "=COUNTIF('mapping'!A2:A10000,\"<>\")",summary_value_format)
        worksheet.write(row+5, col, "=COUNTIF('ASR is empty'!A2:A10000,\"<>\")",summary_value_format)

        print('finish to write summary sheet')
    def writeErrorSheet(self, wb, P_format, F_format, not_test_format):
        print('begin to write error sheet')
        tr_all_sheet = self.tr_all_sheet

        out_row1,out_row2 = 1,1
        out_row3,out_row4 = 1,1
        out_row5,out_row6 = 1,1

        worksheet01 = wb.add_worksheet('ASR-Embedded')
        worksheet02 = wb.add_worksheet('ASR-Cloud')
        worksheet03 = wb.add_worksheet('onenlu-Embedded')
        worksheet04 = wb.add_worksheet('onenlu-Cloud')
        worksheet05 = wb.add_worksheet('mapping')
        worksheet06 = wb.add_worksheet('ASR is empty')

        title_format_all_sheets = self.create_title_format(wb)
        # print(title_format_all_sheets[0])
        title_format_index_1 = title_format_all_sheets[0]['ASR-Embedded']
        title_format_index_2 = title_format_all_sheets[0]['ASR-Cloud']
        title_format_index_3 = title_format_all_sheets[0]['onenlu-Embedded']
        title_format_index_4 = title_format_all_sheets[0]['onenlu-Cloud']
        title_format_index_5 = title_format_all_sheets[0]['mapping']
        title_format_index_6 = title_format_all_sheets[0]['ASR is empty']
        title_format_cololr_list = title_format_all_sheets[1]

        width_dict_01,width_dict_02 = {},{}
        width_dict_03,width_dict_04 = {},{}
        width_dict_05,width_dict_06 = {},{}

        key_list = self.width_dict.keys()
        for key in key_list:
            width_dict_01[key] = 18
            width_dict_02[key] = 18
            width_dict_03[key] = 18
            width_dict_04[key] = 18
            width_dict_05[key] = 18
            width_dict_06[key] = 18

        reset_dict_01 = {'C:C': 40, 'E:E': 12, 'F:F': 40}
        for key,value in reset_dict_01.items():
            width_dict_01[key] = value
        for w in width_dict_01:
            worksheet01.set_column(w, width_dict_01[w])
        worksheet01.set_row(0,15)

        reset_dict_02 = {'C:C': 40, 'E:E': 12, 'F:F': 40}
        for key,value in reset_dict_02.items():
            width_dict_02[key] = value
        for w in width_dict_02:
            worksheet02.set_column(w, width_dict_02[w])
            worksheet02.set_row(0,15)
        
        reset_dict_03 = {'G:G': 12,'H:H': 12, 'I:I': 12, 'J:J': 12, 'K:K': 40}
        for key,value in reset_dict_03.items():
            width_dict_03[key] = value
        for w in width_dict_03:
            worksheet03.set_column(w, width_dict_03[w])
        worksheet03.set_row(0,15)

        reset_dict_04 = {'G:G': 12,'H:H': 12, 'I:I': 12, 'J:J': 12, 'K:K': 40, 'O:O': 40}
        for key,value in reset_dict_04.items():
            width_dict_04[key] = value
            worksheet04.set_column(w, width_dict_04[w])
        worksheet04.set_row(0,15)
        for w in self.width_dict:
            worksheet04.set_column(w, width_dict_04[w])
        worksheet04.set_row(0,15)

        reset_dict_05 = {'G:G': 12, 'H:H': 12, 'I:I': 12, 'J:J': 3.13, 'K:K': 12, 'L:L': 40, 'P:P': 10, 'Q:Q': 12, 'R:R': 12, 'S:S': 12, 'T:T': 12, 'U:U': 40, 'AB:AB': 40}
        for key,value in reset_dict_05.items():
            width_dict_05[key] = value
        for w in width_dict_05:
            worksheet05.set_column(w, width_dict_05[w])
        worksheet05.set_row(0,15)

        reset_dict_06 = {'G:G': 12,'H:H': 12, 'I:I': 12, 'J:J': 12, 'K:K': 12,'L:L': 12, 'M:M': 40, 'Q:Q': 12, 'R:R': 12, 'S:S': 12, 'T:T': 12,'U:U': 12, 'V:V': 40, 'AC:AC': 40}
        for key,value in reset_dict_06.items():
            width_dict_06[key] = value
        for w in width_dict_06:
            worksheet06.set_column(w, width_dict_06[w])
        worksheet06.set_row(0,15)

        wr_all_sheet = self.tr_all_sheet

        for tr_row in range(2, wr_all_sheet.nrows):

            if tr_all_sheet.cell(tr_row, 11).value == 'F' and tr_all_sheet.cell(tr_row, 20).value == 'onboard' and not (tr_all_sheet.cell(tr_row, 26).value == '' and tr_all_sheet.cell(tr_row, 29).value == ''):

                Col_lists = ['CaseID', 'utterance', 'Final Fail Reason', 'Final ASR', 'Final Origin','Off_Sessionid','Comment']
                if out_row1 == 1:
                    for key,value in title_format_index_1.items():
                        if len(value) > 0:
                            for i in value:
                                worksheet01.write(0, i, Col_lists[i], title_format_cololr_list[key])

                # print('out_row1:%d' %out_row1)

                tr_col_list = [1,4,16,17,20,32] 
                col_dict = {}
                for i in range(len(Col_lists)-1):
                    col_dict[i] = tr_col_list[i]

                for col in range(len(Col_lists)-1):
                    value = tr_all_sheet.cell(tr_row, col_dict[col]).value
                    if value == 'P':
                        worksheet01.write(out_row1, col, value, P_format)
                    elif value == 'F':
                        worksheet01.write(out_row1, col, value, F_format)
                    elif value == 'None' or value == 'Not Run':
                        worksheet01.write(out_row1, col, value, not_test_format)
                    else:
                        worksheet01.write(out_row1, col, value)
                    self.get_cell_color_distribution(tr_row,out_row1,col,col_dict[col],tr_all_sheet,wb,worksheet01)
                out_row1 = out_row1 + 1
            
            elif tr_all_sheet.cell(tr_row, 11).value == 'F' and tr_all_sheet.cell(tr_row, 20).value == 'offboard' and not (tr_all_sheet.cell(tr_row, 26).value == '' and tr_all_sheet.cell(tr_row, 29).value == ''):
                
                Col_lists = ['CaseID','utterance','Final Fail Reason','Final ASR','Final Origin','Off_Sessionid','Comment']
                
                if out_row2 == 1:
                    for key,value in title_format_index_2.items():
                        if len(value) > 0:
                            for i in value:
                                worksheet02.write(0, i, Col_lists[i], title_format_cololr_list[key])

                # print('out_row2:%d' %out_row2)

                tr_col_list = [1, 4, 16, 17, 20, 32]
                col_dict = {}
                for i in range(len(Col_lists)-1):
                    col_dict[i] = tr_col_list[i]

                for col in range(len(Col_lists)-1):
                    value = tr_all_sheet.cell(tr_row, col_dict[col]).value
                    if value == 'P':
                        worksheet02.write(out_row2, col, value, P_format)
                    elif value == 'F':
                        worksheet02.write(out_row2, col, value, F_format)
                    elif value == 'None' or value == 'Not Run':
                        worksheet02.write(out_row2, col, value, not_test_format)
                    else:
                        worksheet02.write(out_row2, col, value)
                    self.get_cell_color_distribution(tr_row,out_row2,col,col_dict[col],tr_all_sheet,wb,worksheet02)
                out_row2 = out_row2 + 1

            # U V Y 20 21 24
            elif tr_all_sheet.cell(tr_row, 20).value == 'onboard' and tr_all_sheet.cell(tr_row, 21).value == 'P' and tr_all_sheet.cell(tr_row, 24).value == 'F':
                                
                Col_lists = ['CaseID', 'utterance', 'Intent', 'Slot', 'Onenlu_intent', 'Onenlu_slot',
                        'Final Origin','OneNLU A-res', 'OneNLU T-res','OneNLU S-res', 'OneNLU Fail Reason','On_ASR', 'On_topic', 'On_Slots']
                if out_row3 == 1:
                    for key,value in title_format_index_3.items():
                        if len(value) > 0:
                            for i in value:
                                worksheet03.write(0, i, Col_lists[i], title_format_cololr_list[key])

                # print('out_row3:%d' %out_row3)

                tr_col_list = [1, 4, 7, 8, 9, 10, 20, 21, 22, 23, 25, 26, 27, 28]
                col_dict = {}
                for i in range(len(Col_lists)):
                    col_dict[i] = tr_col_list[i]

                for col in range(len(Col_lists)):
                    value = tr_all_sheet.cell(tr_row, col_dict[col]).value
                    worksheet03.write(out_row3, col, value)

                    if value == 'P':
                        worksheet03.write(out_row3, col, value, P_format)
                    elif value == 'F':
                        worksheet03.write(out_row3, col, value, F_format)
                    elif value == 'None' or value == 'Not Run':
                        worksheet03.write(out_row3, col, value, not_test_format)
                    else:
                        worksheet03.write(out_row3, col, value)
                    self.get_cell_color_distribution(tr_row,out_row3,col,col_dict[col],tr_all_sheet,wb,worksheet03)
                out_row3 = out_row3 + 1

            # U V Y 20 21 24
            elif tr_all_sheet.cell(tr_row, 20).value == 'offboard' and tr_all_sheet.cell(tr_row, 21).value == 'P' and tr_all_sheet.cell(tr_row, 24).value == 'F':
               
                Col_lists = ['CaseID', 'utterance', 'Intent', 'Slot', 'Onenlu_intent', 'Onenlu_slot',
                        'Final Origin','OneNLU A-res', 'OneNLU T-res','OneNLU S-res', 'OneNLU Fail Reason','Off_ASR', 'Off_topic', 'Off_Slots','Off_Sessionid']
                if out_row4 == 1:
                    for key,value in title_format_index_4.items():
                        if len(value) > 0:
                            for i in value:
                                worksheet04.write(0, i, Col_lists[i], title_format_cololr_list[key])

                # print('out_row4:%d' %out_row4)

                tr_col_list = [1, 4, 7, 8, 9, 10, 20, 21, 22, 23, 25, 29, 30, 31, 32]
                col_dict = {}
                for i in range(len(Col_lists)):
                    col_dict[i] = tr_col_list[i]

                for col in range(len(Col_lists)):
                    value = tr_all_sheet.cell(tr_row, col_dict[col]).value
                    if value == 'P':
                        worksheet04.write(out_row4, col, value, P_format)
                    elif value == 'F':
                        worksheet04.write(out_row4, col, value, F_format)
                    elif value == 'None' or value == 'Not Run':
                        worksheet04.write(out_row4, col, value, not_test_format)
                    else:
                        worksheet04.write(out_row4, col, value)
                    self.get_cell_color_distribution(tr_row,out_row4,col,col_dict[col],tr_all_sheet,wb,worksheet04)
                out_row4 = out_row4 + 1

            # L P V Y 12 16 22 25
            elif tr_all_sheet.cell(tr_row, 11).value == 'P' and tr_all_sheet.cell(tr_row, 15).value == 'F' and tr_all_sheet.cell(tr_row, 21).value == 'P' and tr_all_sheet.cell(tr_row, 24).value == 'P':
                Col_lists = ['CaseID',  'utterance', 'Intent', 'Slot', 'Onenlu_intent', 'Onenlu_slot',
                        'Final A-res', 'Final T-res', 'Final S-res', 'Final O-res','Final F-res', 'Final Fail Reason', 'Final ASR', 'Final Topic', 'Final Slots','Final Origin',
                        'OneNLU A-res', 'OneNLU T-res','OneNLU S-res', 'OneNLU F-res', 'OneNLU Fail Reason','On_ASR', 'On_topic', 'On_Slots',
                        'Off_ASR','Off_Topic','Off_Slot','Off_Sessionid']
                if out_row5 == 1:
                    for key,value in title_format_index_5.items():
                        if len(value) > 0:
                            for i in value:
                                worksheet05.write(0, i, Col_lists[i], title_format_cololr_list[key])

                # print('out_row5:%d' %out_row5)

                tr_col_list = [1, 4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32]
                col_dict = {}
                for i in range(len(Col_lists)):
                    col_dict[i] = tr_col_list[i]

                for col in range(len(Col_lists)):
                    value = tr_all_sheet.cell(tr_row, col_dict[col]).value
                    if value == 'P':
                        worksheet05.write(out_row5, col, value, P_format)
                    elif value == 'F':
                        worksheet05.write(out_row5, col, value, F_format)
                    elif value == 'None' or value == 'Not Run':
                        worksheet05.write(out_row5, col, value, not_test_format)
                    else:
                        worksheet05.write(out_row5, col, value)
                    self.get_cell_color_distribution(tr_row,out_row5,col,col_dict[col],tr_all_sheet,wb,worksheet05)
                
                out_row5 = out_row5 + 1

            elif tr_all_sheet.cell(tr_row, 26).value == '' and tr_all_sheet.cell(tr_row, 29).value == '':
                      
                Col_lists = ['Status', 'CaseID', 'utterance', 'Intent', 'Slot', 'Onenlu_intent', 'Onenlu_slot',
                'A-res', 'T-res', 'S-res', 'O-res','F-res', 'Fail Reason', 'Final ASR', 'Final Topic', 'Final Slots','Final Origin',
                'A-res', 'T-res','S-res', 'F-res', 'Fail Reason','On_ASR', 'On_topic', 'On_Slots','Off_ASR','Off_Topic','Off_Slot','Off_Sessionid']
                if out_row6 == 1:
                    for key,value in title_format_index_6.items():
                        if len(value) > 0:
                            for i in value:
                                worksheet06.write(0, i, Col_lists[i], title_format_cololr_list[key])

                # print('out_row6:%d' %out_row6)

                tr_col_list = [0, 1, 4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 11, 12, 13, 15, 16, 26, 27, 28, 29, 30, 31, 32]
                col_dict = {}
                for i in range(len(Col_lists)):
                    col_dict[i] = tr_col_list[i]

                for col in range(len(Col_lists)):
                    value = tr_all_sheet.cell(tr_row, col_dict[col]).value
                    if value == 'P':
                        worksheet06.write(out_row6, col, value, P_format)
                    elif value == 'F':
                        worksheet06.write(out_row6, col, value, F_format)
                    elif value == 'None' or value == 'Not Run':
                        worksheet06.write(out_row6, col, value, not_test_format)
                    else:
                        worksheet06.write(out_row6, col, value)
                    self.get_cell_color_distribution(tr_row,out_row6,col,col_dict[col],tr_all_sheet,wb,worksheet06)
                
                out_row6 = out_row6 + 1
        
        print('finish to write error sheet')

    def write2excel(self):
        wb = xlsxwriter.Workbook(self.track_list_path)
        # print(self.track_list_path)
        ###### error sheet format   #######
        # title_format = wb.add_format(
        #     {'border': True, 'bold': True, 'align': 'center', 'valign': 'vcenter',"bg_color": '209505', "font_color": 'FFFFFF',
        #         'font_size': '11', 'font_name': 'Arial'})
        not_test_format = wb.add_format(
            {'bold': False, 'align': 'top', "font_color": '000000', "bg_color": 'FFFF00', 'border': True,
                'font_size': '11', 'font_name': 'Arial'})
        not_test_format.set_border_color("FFFFFF")
        F_format = wb.add_format({'border': True, 'bold': False, 'align': 'top', "bg_color": 'FF0000'})
        P_format = wb.add_format({'border': True, 'bold': False, 'align': 'top', "bg_color": '00FF00'})

        self.writeSummarySheet(wb)
        self.writeErrorSheet(wb, P_format, F_format, not_test_format)
        self.tr_workbook.release_resources()
        del self.tr_workbook
        # self.track_template_workbook.release_resources()
        # del self.track_template_workbook
        wb.close()


    def run(self):
        self.write2excel()

if __name__ == '__main__':
    # required_sheets = ['ASR-Embedded','ASR-Cloud','onenlu-Embedded','onenlu-Cloud','mapping','ASR is empty']
    tr_name = sys.argv[1].replace(' ', '')
    track_list = TrackList(tr_name)
    track_list.run()
