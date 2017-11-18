# -*- coding:utf-8 -*-
import os
import re
from functools import reduce

import numpy as np
import pandas as pd

from setting import PROPERTIES, PROJECT_INTER_OUTPUT_DIR, EXCEL_SUFFIX

date_pattern = re.compile(r'原始客服表-(\d{4})_(\d{1,2})_(\d{1,2}).[Xx][Ll][Ss][Xx]*')


def process(file_name):
    match = date_pattern.findall(file_name)
    curr_date = '_'.join(match[0])
    df = pd.read_excel(PROPERTIES['origin_customer_table_input_dir'] + '\\' + file_name)
    origin_columns = list(df.columns.values)
    output_path = PROPERTIES['processed_customer_table_output_dir'] + '\\'
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    if not os.path.exists(PROJECT_INTER_OUTPUT_DIR + '\\customer'):
        os.makedirs(PROJECT_INTER_OUTPUT_DIR + '\\customer')
    # 1.预处理
    # start
    df['订单编号'] = df['订单编号'].astype(np.str)
    filtered_df = df.iloc[:, :origin_columns.index('备注')]  # 删掉备注后的列
    del filtered_df['推广费']  # 删掉推广费一列
    filtered_df.to_excel(output_path + '有效客服表-' +
                         curr_date + EXCEL_SUFFIX, index=False)
    filtered_columns = list(filtered_df.columns.values)
    # end
    # 讲退款中的订单提取出来
    # start
    # 筛选出退款中的记录
    sale_return = filtered_df[filtered_df.订单状态 == '退款中']
    # 导出表
    sale_return.sort_values(by='付款时间') \
        .to_excel(output_path + '退款客服表-' +
                  curr_date + EXCEL_SUFFIX, sheet_name='退款中',
                  index=False)
    # end
    # 2 start
    item_classification = dict()
    data = []
    item_pattern = re.compile(r'(.*?)\s*(（.*?）)*\s*\[数量:(\d+)\]')

    for i in filtered_df.index:
        column = filtered_df.iloc[i].values
        order_desc = str(column[origin_columns.index('订单描述')]).split(';')
        item_price = str(column[origin_columns.index('商品价格')]).split(';')
        by_num = str(column[origin_columns.index('购买数量')]).split(';')
        total_num = reduce(lambda x, y: x + y, map(lambda x: int(x), by_num))

        for desc, price, num in zip(order_desc, item_price, by_num):
            # 订单拆分
            match_result = item_pattern.findall(desc)
            item_name = match_result[0][0]
            d = column.copy()
            d[origin_columns.index('订单描述')] = item_name
            d[origin_columns.index('商品价格')] = price
            d[origin_columns.index('购买数量')] = num
            # 处理运费
            # print(d[origin_columns.index('运费')])
            d[origin_columns.index('运费')] = '0.0' if str(d[origin_columns.index('运费')]) == '0.0' \
                else str(d[origin_columns.index('运费')]) + '/' + str(total_num)
            data.append(d)
            items = item_classification.get(item_name, [])
            items.append(d)
            item_classification[item_name] = items
    result_df = pd.DataFrame(data=data, columns=filtered_columns)
    # 运费排序
    result_df = result_df.sort_values(by='运费')
    # 输出
    result_df.to_excel(output_path + '拆分客服表-' +
                       curr_date + EXCEL_SUFFIX, index=False)
    result_df.to_excel(PROJECT_INTER_OUTPUT_DIR + '\\customer\\拆分客服表-' +
                       curr_date + EXCEL_SUFFIX, index=False)

    # 2 end
    # 3 商品分类输出
    classify_path = output_path + '分类客户表-' + curr_date
    if not os.path.exists(classify_path):
        os.mkdir(classify_path)

    for item_name, items in item_classification.items():
        print(item_name)
        classify_df = pd.DataFrame(data=items, columns=filtered_columns).sort_values(by='运费')
        classify_df['购买数量'] = classify_df['购买数量'].astype(np.int)
        total = classify_df['购买数量'].sum()
        total_data = ['' for _ in range(filtered_columns.__len__())]
        total_data[0] = total
        total_df = pd.DataFrame(data=[total_data], columns=filtered_columns)
        classify_df = classify_df.append(total_df, ignore_index=True)
        classify_df.to_excel(classify_path + '/分类客户表-' + curr_date
                             + '-' + item_name.replace(':', '') + EXCEL_SUFFIX, index=False)


for file_or_dir in os.listdir(PROPERTIES['origin_customer_table_input_dir']):

    if PROPERTIES['origin_customer_table_input_dir'] + '\\' + file_or_dir \
            and date_pattern.match(file_or_dir):
        process(file_or_dir)
os.system("explorer " + PROPERTIES['processed_customer_table_output_dir'])