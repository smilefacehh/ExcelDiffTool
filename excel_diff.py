# -*- coding:utf-8 -*-
# @Date: 2020-11-03
# @Autor: rafaellu
from enum import Enum
import xlrd
import xlsxwriter
import os
import re
import operator
import sys
import copy
from MainWindow import Ui_MainWindow

COMPARE_LINE_DIS = 100   # 行相似度计算，只比较相邻的行，全部行比较比较慢。全部行比较，这里写-1
COMPARE_GAMMA = 0.8     # 行相似度 = 0.8 * 一行中格子相同占比 + 0.2 * 行距


class MODIFY(Enum):
    """
    修改，0不变，1修改，2新增，3删除
    """
    UNCHANGED = 0
    MOD = 1
    ADD = 2
    DEL = 3
    CHANGE_ROW = 4 # 行变化了

class COMPARE(Enum):
    """
    比较，0表示共有，1、2表示独有
    """
    BOTH = 0
    OTHER = 1
    ME = 2


class Sheet:
    """
    Excel页签类
    """
    def __init__(self, excel_path, sheet_name):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.title = []
        self.content = []

    def set_title(self, title):
        for idx in range(len(title)):
            if title[idx] == '':
                if idx == 0:
                    title[idx] = '#'
                else:
                    title[idx] = str(title[idx-1]) + '#' # fix bug: 可能存在表头内容空的情况，默认给它一个名字，用前面的名字加上'#'
            title[idx] = str(title[idx])
            if title[idx] in self.title:
                c = 0
                new_title = title[idx] + '@'
                # 可能有多个相同的表头，@往后加
                for t in self.title:
                    if title[idx] in t:
                        if c < t.count('@'):
                            c = t.count('@')
                            new_title = t + '@'
                self.title.append(new_title) # fix bug: 可能存在表头内容相同的情况，名字加@
            else:
                self.title.append(title[idx])

    def set_content(self, content):
        self.content = content
        for idx in range(len(self.content[0])):
            self.content[0][idx] = self.title[idx]

    def compare_title(self, other):
        """
        @brief  比较两个sheet的列头
        @return 返回格式：{title:COMPARE,} ，0表示共有，1表示other独有，2表示自己独有
        """
        assert isinstance(other, Sheet), 'invalid type of other obj, should be Sheet'
        l = list(set(self.title + other.title))
        r = {}
        for t in l:
            if t in self.title:
                if t in other.title:
                    r[t] = COMPARE.BOTH
                else:
                    r[t] = COMPARE.ME
            else:
                r[t] = COMPARE.OTHER
        return r


class ExcelDiff:
    """
    Excel对比
    """
    def __init__(self, excel1, excel2, window=None):
        self.excel_path1 = excel1
        self.excel_path2 = excel2
        self.excel_path_diff = excel2.split('/')[-1].split('.')[0] + '_diff.xlsx' # 差异文件写回路径，当前目录
        self.color_table = {MODIFY.UNCHANGED:'FFFFFF', MODIFY.MOD:'FFFF00', MODIFY.ADD:'6B84EF', MODIFY.DEL:'FF4F4F', MODIFY.CHANGE_ROW:'90EE90'} # 白色、黄色、蓝色、红色、粉色
        self.log = {} # 输出日志，{sheetname: {titles: [], modify: [], add: [], del: [], changerow: []}}
        self.window = window # 用于打印日志

    def read_excel(self, path):
        """
        @brief  读Excel，多个页签
        @return 返回格式：{sheet_name:Sheet}
        """
        r = {}
        xl = xlrd.open_workbook(path)
        n = xl.nsheets
        for idx_sheet in range(n):
            sh = xl.sheets()[idx_sheet]
            content = []
            if sh.visibility == 1 or sh.nrows == 0:
                continue
            for idx_row in range(sh.nrows):
                row = []
                for idx_col in range(sh.ncols):
                    ctype = sh.cell(idx_row, idx_col).ctype # 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
                    cell = sh.cell_value(idx_row, idx_col)
                    if ctype == 2: 
                        if cell % 1 == 0.0: # 整数
                            cell = int(cell)
                        else:
                            cell = round(cell, 7) # 最多保留7位小数
                    row.append(cell)
                content.append(row)

            sheet = Sheet(path, sh.name)
            sheet.set_title(sh.row_values(0))
            sheet.set_content(content)
            r[sh.name] = sheet
        return r


    def write_excel(self, path, ex_content, ex_modify):
        """
        @brief 写回Excel
        """
        assert isinstance(ex_content, dict) and isinstance(ex_content, dict), 'invalid type.'
        xl = xlsxwriter.Workbook(path)
        for sheet_name, content_tup in ex_content.items():
            content, modify = content_tup
            sheet_modify = ex_modify[sheet_name]
            sh = xl.add_worksheet(sheet_name)

            for i in range(len(content)):
                for j in range(len(content[0])):
                    cell_format = xl.add_format({'bg_color': self.color_table[modify[i][j]], 'align': 'left', 'border': 1, 'border_color': '#CCCCCC','bold': i==0})
                    sh.write(i, j, content[i][j], cell_format)

            # 下方页签颜色
            if sheet_modify == MODIFY.MOD:
                sh.set_tab_color('yellow')
            elif sheet_modify == MODIFY.DEL:
                sh.set_tab_color('red')
            elif sheet_modify == MODIFY.ADD:
                sh.set_tab_color('green')
        xl.close()


    def line_similarity(self, line1, line2, idx_row1, idx_row2, max_row):
        """
        @brief 计算单元格相似度, [0,1]
        """
        assert isinstance(line1, list) and isinstance(line2, list), 'invalid type.'
        gamma = COMPARE_GAMMA
        same = 0
        total = len(line1)
        for idx in range(total):
            if line1[idx] == line2[idx]:
                same += 1
        return gamma * (same / total) + (1 - gamma) * (1 - abs(idx_row1 - idx_row2) / max(1, max_row))


    def union_sheet_name(self, ex1, ex2):
        """
        @brief 两个excel表格的页签集合，标记哪边独有、共有。0共有，1为ex1独有，2为ex2独有
        """
        r = {}
        l = list(set(list(ex1) + list(ex2)))
        for t in l:
            if t in ex1:
                if t in ex2:
                    r[t] = COMPARE.BOTH
                else:
                    r[t] = COMPARE.OTHER
            else:
                r[t] = COMPARE.ME
        return r

    def diff_sheet(self, sh1, sh2):
        """
        @brief 对比两个页签
        @return content：二维list，包含表头； modify：二维list，单元格修改情况
        """
        print('对比页签：%s' % sh1.sheet_name)
        # if self.window != None:
        #     self.window.log('[日志] 对比页签：%s' % sh1.sheet_name)
        assert isinstance(sh1, Sheet) and isinstance(sh2, Sheet), 'invalid type, should be Sheet.'
        title_union = sh2.compare_title(sh1) # 列头集合

        # Step1.这部分是公共列的内容，用来做对比
        sh_inter1 = copy.deepcopy(sh1.content)
        sh_inter2 = copy.deepcopy(sh2.content)

        # 遍历公共列
        for k,v in title_union.items():
            if v == COMPARE.OTHER:
                idx = -1
                for i in range(len(sh_inter1[0])):
                    if sh_inter1[0][i] == k:
                        idx = i
                        break
                if idx != -1:
                    for row in sh_inter1:
                        del row[idx]
            elif v == COMPARE.ME:
                idx = -1
                for i in range(len(sh_inter2[0])):
                    if sh_inter2[0][i] == k:
                        idx = i
                        break
                if idx != -1:
                    for row in sh_inter2:
                        del row[idx]

        # 调整列顺序，防止列顺序不对应
        for idx_col in range(len(sh_inter1[0])):
            if sh_inter1[0][idx_col] != sh_inter2[0][idx_col]:
                for i in range(len(sh_inter2[0])):
                    if sh_inter2[0][i] == sh_inter1[0][idx_col]:
                        for j in range(len(sh_inter2)):
                            sh_inter2[j][i], sh_inter2[j][idx_col] = sh_inter2[j][idx_col], sh_inter2[j][i]
                        break
        
        # Step2.遍历所有行，计算行相似度
        n1 = len(sh_inter1)
        n2 = len(sh_inter2)
        max_n = max(n1, n2)
        sh_match1 = [0] * n1  # sh1对应的sh2匹配行
        sh_match2 = [0] * n2  # sh2对应的sh1匹配行
        match_sim = {}        # 匹配，相似度，用于排序

        for i in range(1, n1):
            min_idx = 1
            max_idx = n2
            if COMPARE_LINE_DIS != -1:
                min_idx = max(1, i - COMPARE_LINE_DIS)
                max_idx = min(i + COMPARE_LINE_DIS, n2)
            for j in range(min_idx, max_idx):
                sim = self.line_similarity(sh_inter1[i], sh_inter2[j], i, j, max_n - 2)
                match_sim[(i,j)] = sim

        # Step3.相似度排序，行匹配
        match_sim_sorted = sorted(match_sim.items(), key=lambda x: x[1], reverse=True)
        for match in match_sim_sorted:
            idx1 = match[0][0]
            idx2 = match[0][1]
            if sh_match1[idx1] == 0 and sh_match2[idx2] == 0 and match[1] > 0:
                sh_match1[idx1] = idx2
                sh_match2[idx2] = idx1
        
        # Step4.生成合并内容
        content = [] # 差异内容
        modify = []  # 修改情况

        # 标题在原始数据中对应的列index，便于查找
        title_idx1 = {}
        title_idx2 = {}

        # 首行标题合并，按照表2的顺序
        title = []
        title_modify = []
        
        # 遍历表2的表头
        for i in range(len(sh2.title)):
            title_idx2[sh2.title[i]] = i
            title.append(sh2.title[i])
            if sh2.title[i] in sh_inter2[0]:
                title_modify.append(MODIFY.UNCHANGED)
            else:
                title_modify.append(MODIFY.ADD)

        for i in range(len(sh1.title)):
            title_idx1[sh1.title[i]] = i
            if sh1.title[i] not in title:
                title.append(sh1.title[i])
                title_modify.append(MODIFY.DEL)

        # 待写回的表头内容，恢复原来的内容
        title_writeback = []
        for t in title:
            if '#' in t:
                title_writeback.append('')
            elif '@' in t:
                title_writeback.append(t.replace('@', ''))
            else:
                title_writeback.append(t)
        content.append(title_writeback)
        modify.append(title_modify)
        ncols = len(title)

        # 遍历sh2行，修改、新增
        for idx in range(1, n2):
            line = [''] * ncols
            line_modify = [MODIFY.UNCHANGED] * ncols
            # 匹配行
            if sh_match2[idx] != 0:
                # 如果行位置变化了，先赋值一遍
                if idx != sh_match2[idx]:
                    line_modify = [MODIFY.CHANGE_ROW] * ncols

                    if sh2.sheet_name not in self.log:
                        self.log[sh2.sheet_name] = {
                            '标题': sh2.title,
                            '修改': [],
                            '新增': [],
                            '删除': [],
                            '行变化': []
                        }
                    self.log[sh2.sheet_name]['行变化'].append('表一行%s -> 表二行%s' % (sh_match2[idx] + 1, idx + 1))

                line1 = sh1.content[sh_match2[idx]]
                line2 = sh2.content[idx]
                # 遍历单元格
                for idx_title in range(len(title)):
                    if title[idx_title] in title_idx1 and title[idx_title] in title_idx2:
                        idx1 = title_idx1[title[idx_title]]
                        idx2 = title_idx2[title[idx_title]]
                        # 不变
                        if line1[idx1] == line2[idx2]:
                            line[idx_title] = line2[idx2]
                        # 修改
                        else:
                            line[idx_title] = line2[idx2]
                            line_modify[idx_title] = MODIFY.MOD

                            if sh2.sheet_name not in self.log:
                                self.log[sh2.sheet_name] = {
                                    '标题': sh2.title,
                                    '修改': [],
                                    '新增': [],
                                    '删除': [],
                                    '行变化': []
                                }
                            self.log[sh2.sheet_name]['修改'].append('(行%s,%s) %s -> %s' % (idx + 1, title[idx_title], line1[idx1], line2[idx2]))
                    # 删除
                    elif title[idx_title] in title_idx1:
                        idx1 = title_idx1[title[idx_title]]
                        line[idx_title] = line1[idx1]
                        line_modify[idx_title] = MODIFY.DEL

                        if sh2.sheet_name not in self.log:
                            self.log[sh2.sheet_name] = {
                                '标题': sh2.title,
                                '修改': [],
                                '新增': [],
                                '删除': [],
                                '行变化': []
                            }
                        self.log[sh2.sheet_name]['删除'].append('(行%s,%s) %s' % (sh_match2[idx] + 1, title[idx_title], line1[idx1]))
                    # 新增
                    elif title[idx_title] in title_idx2:
                        idx2 = title_idx2[title[idx_title]]
                        line[idx_title] = line2[idx2]
                        line_modify[idx_title] = MODIFY.ADD

                        if sh2.sheet_name not in self.log:
                            self.log[sh2.sheet_name] = {
                                '标题': sh2.title,
                                '修改': [],
                                '新增': [],
                                '删除': [],
                                '行变化': []
                            }
                        self.log[sh2.sheet_name]['新增'].append('(行%s,%s) %s' % (idx + 1, title[idx_title], line2[idx2]))
            # 未匹配行，新增
            else:
                line2 = sh2.content[idx]
                for idx_title in range(len(title)):
                    if title[idx_title] in title_idx2:
                        line[idx_title] = line2[title_idx2[title[idx_title]]]
                        line_modify[idx_title] = MODIFY.ADD

                if sh2.sheet_name not in self.log:
                    self.log[sh2.sheet_name] = {
                        '标题': sh2.title,
                        '修改': [],
                        '新增': [],
                        '删除': [],
                        '行变化': []
                    }
                line_str = [str(i) for i in line]
                self.log[sh2.sheet_name]['新增'].append('新增整行(行%s): %s' % (idx + 1, ', '.join(line_str)))
            content.append(line)
            modify.append(line_modify)
        
        # sh1独有的行，删除
        for idx in range(1, n1):
            line = [''] * ncols
            line_modify = [MODIFY.DEL] * ncols
            if sh_match1[idx] == 0:
                line1 = sh1.content[idx]
                for idx_title in range(len(title)):
                    if title[idx_title] in title_idx1:
                        line[idx_title] = line1[title_idx1[title[idx_title]]]
                        line_modify[idx_title] = MODIFY.DEL
                content.append(line)
                modify.append(line_modify)

                if sh2.sheet_name not in self.log:
                    self.log[sh2.sheet_name] = {
                        '标题': sh2.title,
                        '修改': [],
                        '新增': [],
                        '删除': [],
                        '行变化': []
                    }
                line_str = [str(i) for i in line]
                self.log[sh2.sheet_name]['删除'].append('删除整行(行%s): %s' % (idx + 1, ', '.join(line_str)))

        return content, modify

    def run(self):
        print('读表')
        ex1 = self.read_excel(self.excel_path1)
        ex2 = self.read_excel(self.excel_path2)
        ex_diff = {}    # 差异，包含所有页签
        ex_modify = {}  # 页签修改情况

        # Step1.页签对比
        sheet_name_dict = self.union_sheet_name(ex1.keys(), ex2.keys())

        # 页签排序
        sheet_name_sorted = []
        for sh in ex2.keys():
            sheet_name_sorted.append(sh)
        for sh in ex1.keys():
            if sh not in sheet_name_sorted:
                sheet_name_sorted.append(sh)

        # 遍历页签
        for sh in sheet_name_sorted:
            k = sh
            v = sheet_name_dict[k]
            if v == COMPARE.BOTH:
                content, modify = self.diff_sheet(ex1[k], ex2[k])
                ex_diff[k] = (content, modify)
                ex_modify[k] = MODIFY.UNCHANGED
                for a in modify:
                    for b in a:
                        if b != MODIFY.UNCHANGED:
                            ex_modify[k] = MODIFY.MOD
                            break
            elif v == COMPARE.OTHER:
                nrows = len(ex1[k].content)
                ncols = len(ex1[k].title)
                ex_diff[k] = (ex1[k].content, [[MODIFY.DEL] * ncols] * nrows)
                ex_modify[k] = MODIFY.DEL
            elif v == COMPARE.ME:
                nrows = len(ex2[k].content)
                ncols = len(ex2[k].title)
                ex_diff[k] = (ex2[k].content, [[MODIFY.ADD] * ncols] * nrows)
                ex_modify[k] = MODIFY.ADD

        # Step2.写回
        print('写回')
        self.write_excel(self.excel_path_diff, ex_diff, ex_modify)

if __name__ == '__main__':
    ed = ExcelDiff('1.xlsx', '2.xlsx')
    ed.run()
