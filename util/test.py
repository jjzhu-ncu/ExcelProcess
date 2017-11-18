
import pandas as pd
import numpy as np
origin_finance_df = pd.read_excel(r'C:\workspace\space4py\ExcelProcess\input\原始财务表-2017_11_14.xls')
origin_customer_df = pd.read_excel(r'C:\workspace\space4py\ExcelProcess\input\原始客服表-2017_11_17.xls')
origin_customer_df = origin_customer_df[['订单编号', '订单金额', '订单描述', '购买数量', '商品价格', '运费']]
origin_customer_df['订单编号'] = origin_customer_df['订单编号'].apply(lambda x: '{:.0f}'.format(x))
origin_customer_df = origin_customer_df.rename(columns={'订单编号': '订单号'})
income_df = origin_finance_df[origin_finance_df.账单类型 == '货款收入']
income_df = income_df[['时间', '收入(元)', '订单号']]
income_df['订单号'] = income_df['订单号'].apply(lambda x: '{:.0f}'.format(x))
merge_df = pd.merge(income_df, origin_customer_df, how='left', on='订单号')
print(merge_df)
no_match_finance_records = list(merge_df[pd.isnull(merge_df['订单金额'])]['订单号'].values)
print(no_match_finance_records)
print(origin_finance_df[origin_finance_df['订单号'].isin(no_match_finance_records)])
