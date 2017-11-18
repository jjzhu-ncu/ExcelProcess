# -*- coding:utf-8 -*-
import os
import re
from functools import reduce
import datetime
import pandas as pd
import xlwt

from setting import PROPERTIES, PROJECT_INTER_OUTPUT_DIR, EXCEL_SUFFIX
from logger import LOGGER

pd.options.mode.chained_assignment = None

if not os.path.exists(PROPERTIES['processed_finance_table_output_dir']):
    LOGGER.warn('processed_finance_table_output_dir:%s do not exist, creating it.'
                % PROPERTIES['processed_finance_table_output_dir'])
    os.makedirs(PROPERTIES['processed_finance_table_output_dir'])
    LOGGER.info('create dir %s successful' % PROPERTIES['processed_finance_table_output_dir'])
origin_customer_dfs = []
# 加载所有拆分客服表
normal_customer_file_pattern = re.compile(r'拆分客服表-\d{4}_\d{1,2}_\d{1,2}.[Xx][Ll][Ss]')
normal_finance_file_pattern = re.compile(r'原始财务表-(\d{4})_(\d{1,2})_(\d{1,2}).[Xx][Ll][Ss]')
LOGGER.info('开始加载拆分客服表')
for file_or_dir in os.listdir(PROJECT_INTER_OUTPUT_DIR + '\\customer'):
    if not os.path.isdir(PROJECT_INTER_OUTPUT_DIR + '\\customer\\' + file_or_dir):
        if normal_customer_file_pattern.match(file_or_dir):
            LOGGER.info('loading customer table %s' % PROJECT_INTER_OUTPUT_DIR + '\\customer\\' + file_or_dir)
            tmp_df = pd.read_excel(PROJECT_INTER_OUTPUT_DIR + '\\customer\\' + file_or_dir)
            origin_customer_dfs.append(tmp_df)
LOGGER.info('加载完毕')
origin_customer_df = reduce(lambda x, y: x.append(y, ignore_index=True), origin_customer_dfs)
origin_customer_df = origin_customer_df[['订单编号', '订单金额', '订单描述', '购买数量', '商品价格', '运费']]
origin_customer_df['订单编号'] = origin_customer_df['订单编号'].apply(lambda x: '{:.0f}'.format(x))
origin_customer_df = origin_customer_df.rename(columns={'订单编号': '订单号'})


def process(file_name):
    LOGGER.info('开始处理 %s' % PROPERTIES['origin_finance_table_input_dir'] + '\\' + file_name)
    origin_finance_df = pd.read_excel(PROPERTIES['origin_finance_table_input_dir'] + '\\' + file_name)
    match = normal_finance_file_pattern.findall(file_name)
    curr_date = '_'.join(match[0])
    origin_finance_df['订单号'] = origin_finance_df['订单号'].apply(lambda x: '{:.0f}'.format(x))
    income_df = origin_finance_df[origin_finance_df.账单类型 == '货款收入']
    other_df = origin_finance_df[origin_finance_df.账单类型 != '货款收入']
    # 保存
    LOGGER.info('保存收入财务表到：%s' % PROPERTIES['processed_finance_table_output_dir'] + '\\收入财务表-' +
                curr_date + EXCEL_SUFFIX)
    income_df.sort_values(by='时间', ascending=False) \
        .to_excel(PROPERTIES['processed_finance_table_output_dir'] + '\\收入财务表-' +
                  curr_date + EXCEL_SUFFIX,
                  index=False)
    other_df = other_df.sort_values(by=['时间', '收支类型'], ascending=[False, True])
    credit_card_pay_df = other_df[other_df.账单类型 == '信用卡手续费']
    total_pay = credit_card_pay_df['收入(元)'].sum()
    other_df_columns = list(other_df.columns.values)
    extra_data = ['' for _ in range(len(other_df_columns))]
    extra_data[other_df_columns.index('收入(元)')] = total_pay
    extra_data[other_df_columns.index('收入(元)') - 1] = '总计手续费'
    extra_df = pd.DataFrame(data=[extra_data], columns=other_df_columns)
    LOGGER.info('保存其他财务表到：%s' % PROPERTIES['processed_finance_table_output_dir'] + '\\其他财务表-' +
                curr_date + EXCEL_SUFFIX)
    other_df.append(extra_df, ignore_index=True). \
        to_excel(PROPERTIES['processed_finance_table_output_dir'] + '\\其他财务表-' +
                 curr_date + EXCEL_SUFFIX,
                 index=False)
    print(income_df['订单号'].dtype)
    income_df = income_df[['时间', '收入(元)', '订单号']]
    merge_df = pd.merge(income_df, origin_customer_df, how='left', on='订单号')

    merge_df['总金额'] = merge_df[['购买数量', '商品价格']].apply(lambda x: x.prod(), axis=1)
    merge_columns = list(merge_df.columns.values)

    no_match_finance_record_nos = list(merge_df[pd.isnull(merge_df['订单金额'])]['订单号'].values)

    match_record = merge_df[pd.notnull(merge_df['订单金额'])]

    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 2
    red_background_style = xlwt.XFStyle()
    # style0.font = font0
    red_background_style.pattern = pattern
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Sheet1')
    for i in range(len(merge_columns)):
        ws.write(0, i, str(merge_columns[i]))
    LOGGER.info('开始核对财务表和客户订单金额')
    match_record = match_record.reset_index(drop=True)
    for index in match_record.index:
        print(index)
        error_flag = False

        column = match_record.iloc[index].values
        if column[merge_columns.index('收入(元)')] != column[merge_columns.index('订单金额')]:
            LOGGER.warn('订单号【%s】有误' % column[merge_columns.index('订单号')])
            error_flag = True
        for c_i in range(len(column)):
            if error_flag:
                ws.write(int(index) + 1, c_i, str(column[c_i]), red_background_style)
            else:
                ws.write(int(index) + 1, c_i, str(column[c_i]))
    wb.save(PROPERTIES['processed_finance_table_output_dir'] + '\\有效财务表-' +
            curr_date + EXCEL_SUFFIX)
    return origin_finance_df[origin_finance_df['订单号'].isin(no_match_finance_record_nos)]
    # merge_df.to_excel(output_path + '收入财务表2-' +
    #                   datetime.datetime.now().__format__('%Y_%m_%d') + '.xlsx',
    #                   index=False)


LOGGER.info('开始处理财务表')
if PROPERTIES['process_all']:
    no_match_dfs = []
    for file_or_dir in os.listdir(PROPERTIES['origin_finance_table_input_dir']):

        if os.path.isfile(PROPERTIES['origin_finance_table_input_dir'] + '\\' + file_or_dir) \
                and normal_finance_file_pattern.match(file_or_dir):
            # print(file_or_dir)
            no_match_dfs.append(process(file_or_dir))
        else:
            LOGGER.warn('无效的文件%s' % file_or_dir)
    no_match_df = reduce(lambda x, y: x.append(y, ignore_index=True), no_match_dfs)
    no_match_df.to_excel(PROPERTIES['processed_finance_table_output_dir'] + '\\原始未处理财务表【汇总】.xls',
                         index=False)
    os.system("explorer " + PROPERTIES['processed_finance_table_output_dir'])
else:
    today = datetime.datetime.now().strftime('%Y_%m_%d')
    process_file_name = "原始财务表-" + today + EXCEL_SUFFIX
    process_file_name_path = PROPERTIES['origin_finance_table_input_dir'] + "\\" + process_file_name
    if os.path.exists(process_file_name_path):
        process(process_file_name)
        os.system("explorer " + PROPERTIES['processed_finance_table_output_dir'])
    else:
        LOGGER.warn('文件夹下没有今天的数据【%s】' % process_file_name_path)
LOGGER.info('处理完毕')
