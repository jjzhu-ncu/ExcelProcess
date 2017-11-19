- conf目录
1. process.yaml目录

origin_customer_table_input_dir 配置原始客服表文件存放路径

processed_customer_table_output_dir 配置处理过后的客服表文件存放路径

origin_finance_table_input_dir 配置原始财务表文件存放路径

processed_finance_table_output_dir 配置处理后的客服表文件路径

process_all true or false，如果为true会处理配置的原始表下所有符合命名规范的文件， false则只处理运行程序当天的文件

- output目录

/output/customer目录下存放了历史的拆分客服表，不要删...
- log 

存放系统日志文件，太大可以删掉..