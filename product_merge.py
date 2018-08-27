# -*- coding: utf-8 -*-
import pandas as pd
import os
import re

re_ma = re.compile('[SMLX均]')


def color_extract(row):
    try:
        color = re_ma.split(row['规格'])[0]
    except:
        color = ''
    return color


def size_extract(row):
    try:
        size = ''.join(re_ma.findall(row['规格']))
    except:
        size = ''
    return size


def color_split(row):
    color = row['商品属性'].split(';')[0].split('：')[-1]
    return color.strip()


def size_split(row):
    size = row['商品属性'].split(';')[-1].split('：')[-1]
    return size.strip()

'''
# print(os.listdir('.'))
xls1 = pd.read_excel('表一：ERP库存数据以此表为基础.xlsx')
print(xls1.columns.tolist())

print(xls1['规格'])
xls1['color'] = xls1.apply(lambda row: color_extract(row), axis=1)
xls1['size'] = xls1.apply(lambda row: size_extract(row), axis=1)
print(xls1[['color', 'size']])

xls2 = pd.read_excel('表二：退款中宝贝报表里面（货品编码）计数每款退款数据存入表一.xlsx')
# print(xls2)
xls2 = xls2.dropna()
# print(xls2)
xls2['货品编号'] = xls2['商家编码'].str.split(pat='_', expand=True)[0]
xls2['color'] = xls2.apply(lambda row: color_split(row), axis=1)
xls2['size'] = xls2.apply(lambda row: size_split(row), axis=1)
# print(xls2[['货品编号', 'color', 'size']])

# ke = xls2[['货品编号', 'color', 'size', '购买数量']].groupby(['货品编号', 'color', 'size'])
# print(ke.count())
kd = xls2['购买数量'].groupby([xls2['货品编号'], xls2['color'], xls2['size']])
# print(kd.count())

xls2_new = kd.count().reset_index()

# print(ke.count().reset_index())
# for i in ke:
    # print(i)
xls2_new.rename(columns={'购买数量': '退货数量'}, inplace=True)
print(xls2_new)
print(xls1.columns.tolist())
m = xls1.merge(xls2_new, on=['货品编号', 'color', 'size'], how='left', indicator=True)
print(m)

# xls1 = pd.read_excel('表一：ERP库存数据以此表为基础.xlsx')
# print(xls1.columns.tolist())
#
# print(xls1['规格'])
# xls1['color'] = xls1.apply(lambda row: color_extract(row), axis=1)
# xls1['size'] = xls1.apply(lambda row: size_extract(row), axis=1)
# print(xls1[['color', 'size']])
# xls1['品名'] = xls1['品名'].str.replace(' ', '')
# print(xls1['品名'])
# xls3 = pd.read_excel('表三：供应商AW订货表需要（品名）匹配表一获取每款订货数据.xlsx')
# xls3['品名'] = xls3.apply(lambda row: row['商品代码'] + row['商品名称'], axis=1)
# print(xls3)
# m2 = xls1.merge(xls3[['品名', '颜色名称', '尺码名称', '未完工数']], left_on=['品名', 'color', 'size'], right_on=['品名', '颜色名称', '尺码名称'], how='left')
# print(m2)

xls1 = pd.read_excel('表一：ERP库存数据以此表为基础.xlsx')
print(xls1.columns.tolist())

print(xls1['规格'])
xls1['color'] = xls1.apply(lambda row: color_extract(row), axis=1)
xls1['size'] = xls1.apply(lambda row: size_extract(row), axis=1)
# print(xls1[['color', 'size']])
xls1['品名'] = xls1['品名'].str.replace(' ', '')

xls4 = pd.read_excel('表四：获取周销量数据.xlsx')
# print(xls4.columns.tolist())
# print(xls4[['商家编码', '交易笔数']])
k = xls4['交易笔数'].groupby(xls4['商家编码']).sum()
k = k.reset_index()
# print(k.drop_duplicates('商家编码'))
xls1_for_4 = xls1.drop_duplicates('货品编号')
# print(xls1_for_4)
m = xls1_for_4.merge(k, left_on='货品编号', right_on='商家编码', how='left')
print(m['交易笔数'])
fi = xls1.merge(m[['货品编号', 'color', 'size', '交易笔数']], on=['货品编号', 'color', 'size'], how='left')
print(fi[['货品编号', 'color', 'size', '交易笔数']])
'''


class TableMix(object):
    def __init__(self, table1='', table2='', table3='', table4='', table5='', table6=''):
        self.org = ['货品编号', 'size', 'color']
        if table1:
            try:
                self.table1 = pd.read_excel(table1)
            except:
                self.table1 = pd.read_csv(table1)
        if table2:
            try:
                self.table2 = pd.read_excel(table2)
            except:
                self.table2 = pd.read_csv(table2)
        if table3:
            try:
                self.table3 = pd.read_excel(table3)
            except:
                self.table3 = pd.read_csv(table3)
        if table4:
            try:
                self.table4 = pd.read_excel(table4)
            except:
                self.table4 = pd.read_csv(table4)
        if table5:
            try:
                self.table5 = pd.read_excel(table5)
            except:
                self.table5 = pd.read_csv(table5)
        if table6:
            try:
                self.table6 = pd.read_excel(table6)
            except:
                self.table6 = pd.read_csv(table6)

    def __getattr__(self, item):
        return item

    def table1_rp(self):
        if not type(self.table1) == pd.DataFrame:
            return
        # 分离颜色
        table1 = self.table1
        table1['color'] = table1.apply(lambda row: color_extract(row), axis=1)
        # 分离尺码
        table1['size'] = table1.apply(lambda row: size_extract(row), axis=1)
        # 修改品名
        table1['品名'] = table1['品名'].str.replace(' ', '')
        # print(self.table1[['color', 'size']])
        return table1

    def table2_rp(self):
        if not type(self.table2) == pd.DataFrame:
            return
        table2 = self.table2
        # 去掉空值
        table2 = table2[self.table2['商家编码'] != 'null']

        # 分离货品编号
        # table2['货品编号'] = table2['商家编码'].str.split(pat='_', expand=True)[0]
        # 分离颜色
        # table2['color'] = table2.apply(lambda row: color_split(row), axis=1)
        # 分离尺码
        # table2['size'] = table2.apply(lambda row: size_split(row), axis=1)
        # 统计购买数量
        table2_groupby = table2['购买数量'].groupby(table2['商家编码'])
        table2_groupby_count = table2_groupby.sum().reset_index()
        table2_groupby_count.rename(columns={'购买数量': '退货数量'}, inplace=True)
        print(table2_groupby_count)
        return table2_groupby_count

    def table3_rp(self):
        if not type(self.table3) == pd.DataFrame:
            return
        # xls3['品名'] = xls3.apply(lambda row: row['商品代码'] + row['商品名称'], axis=1)
        table3 = self.table3
        table3['品名'] = table3.apply(lambda row: row['商品代码'] + row['商品名称'], axis=1)
        # print(table3)
        return table3

    def table4_rp(self):
        if not type(self.table4) == pd.DataFrame:
            return
        table4 = self.table4
        table4_sum = table4['交易笔数'].groupby(table4['商家编码']).sum()
        table4_sum = table4_sum.reset_index()
        table4_sum.rename(columns={'商家编码': '货品编号'}, inplace=True)
        # print(table4_sum)
        return table4_sum

    def table6_rp(self):
        if not type(self.table6) == pd.DataFrame:
            return
        table6 = self.table6
        # 去掉空值
        table6 = table6[self.table6['商家编码'] != 'null']
        table6 = table6[self.table6['订单状态'] != '交易关闭']

        table6_groupby = table6['购买数量'].groupby(table6['商家编码'])
        table6_groupby_count = table6_groupby.sum().reset_index()
        # table2_groupby_count.rename(columns={'购买数量': '退货数量'}, inplace=True)
        print(table6_groupby_count)
        return table6_groupby_count

    @staticmethod
    def get_fujiama(table1, table5):
        mix = table1.merge(table5[['货品编号', '规格', '条码+附加码']], on=['货品编号', '规格'], how='left')
        mix.rename(columns={'条码+附加码': '商家编码'}, inplace=True)
        return mix

    @staticmethod
    def merge_1(table1, table2):
        m1 = table1.merge(table2, on=['货品编号', 'color', 'size'], how='left', indicator=False)
        m1['退货数量'] = m1['退货数量'].fillna(0)
        return m1

    @staticmethod
    def merge_2(m1, table3):
        m2 = m1.merge(table3[['品名', '颜色名称', '尺码名称', '未完工数']], left_on=['品名', 'color', 'size'], right_on=['品名', '颜色名称', '尺码名称'], how='left')
        m2 = m2.drop(['颜色名称', '尺码名称'], axis=1)
        m2['未完工数'] = m2['未完工数'].fillna(0)
        return m2

    @staticmethod
    def merge_3(m2, table4):
        m2_for_table4 = m2.drop_duplicates('货品编号')

        temp = m2_for_table4.merge(table4[['货品编号', '交易笔数']], on='货品编号', how='left')

        m3 = m2.merge(temp[['货品编号', 'color', 'size', '交易笔数']], on=['货品编号', 'color', 'size'], how='left')
        m3['交易笔数'] = m3['交易笔数'].fillna(0)
        return m3

    @staticmethod
    def merge_process(m3):
        m3.rename(columns={'交易笔数': '周销量'}, inplace=True)
        m3['可订购'] = m3['库存量'] - m3['订购量'] - m3['待发量']
        # m3['周销量预期订货'] = m3['周销量'] - m3['未完工数'] - m3['退货数量'] - m3['可订购']
        return m3

    def table_merge(self):
        table1 = self.table1_rp()
        table2 = self.table2_rp()
        table3 = self.table3_rp()
        table4 = self.table4_rp()
        table6 = self.table6_rp()

        table1 = self.get_fujiama(table1, self.table5)

        m1 = table1.merge(table2, on=['商家编码'], how='left', indicator=False)
        # m1['退货数量'] = m1['退货数量'].fillna(0)
        m2 = m1.merge(table3[['品名', '颜色名称', '尺码名称', '未完工数']], left_on=['品名', 'color', 'size'], right_on=['品名', '颜色名称', '尺码名称'], how='left')
        m2 = m2.drop(['颜色名称', '尺码名称'], axis=1)
        # m2['未完工数'] = m2['未完工数'].fillna(0)


        m2_for_table4 = m2.drop_duplicates('货品编号')

        temp = m2_for_table4.merge(table4[['货品编号', '交易笔数']], on='货品编号', how='left')

        m3 = m2.merge(temp[['货品编号', 'color', 'size', '交易笔数']], on=['货品编号', 'color', 'size'], how='left')
        # m3['交易笔数'] = m3['交易笔数'].fillna(0)
        m3.rename(columns={'交易笔数': '周销量'}, inplace=True)

        m4 = m3.merge(table6, on=['商家编码'], how='left', indicator=False)
        m4['可订购'] = m4['库存量'] - m4['订购量'] - m4['待发量']
        # m3['周销量预期订货'] = m3['周销量'] - m3['未完工数'] - m3['退货数量'] - m3['可订购']
        print(m4)
        self.xlsx_maker(m4)

    @staticmethod
    def xlsx_maker(dataframe):
        writer = pd.ExcelWriter('output.xlsx')
        dataframe.to_excel(writer, 'Sheet1')
        writer.save()


if __name__ == '__main__':
    process = TableMix(table1='表1：ERP仓储数据.xlsx',
                       table2='表2：退款订单.xlsx',
                       table3='表3：供应商欠货.xls',
                       table4='表4生意经周销量数据.xlsx',
                       table5='表五：用来补齐表一缺的货品编码下的SKU编码.xlsx',
                       table6='表六：一周销售订单样本用于统计每个SKU一周销售数.xlsx')

    process.table_merge()
