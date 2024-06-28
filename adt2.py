import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def descr_key1(x):  # 关键字提取，处理前后左右内饰特征
    if x.find('左前') >=0 :
        return '2左前'
    if x.find('右前') >=0  :
        return '2右前'
    if x.find('左后') >=0  :
        return '1左后'
    if x.find('右后') >=0  :
        return '1右后'
    return 'NA'

def to_excel(df: pd.DataFrame):
    in_memory_fp = BytesIO() #内存中读写二进制数据
    df.to_excel(in_memory_fp, index=False) #导出excel
    # Write the file out to disk to demonstrate that it worked.
    in_memory_fp.seek(0, 0) #移动读取指针的位置
    return in_memory_fp.read() #读取数据

st.title('细品计划分解系统v1.0.1')
st.header('恒展迈思软件')
st.text('该系统会根据一汽红旗计划与BOM系统自动分解计划与系统配置文件')

uf_bom = st.file_uploader("BOM",['xlsx', 'xls'])
if uf_bom is not None:                                   # 处理BOM文件
    dfb = pd.read_excel(uf_bom,sheet_name='BOM清单')      # 处理《BOM清单》excel文件
    dfb
    dfb = dfb.rename(columns={"去-零件号": "零件号"})       # 规范EXCEL文件列名
    # Pandas行列转换的4大技巧 - 尤而小屋的文章 - 知乎
    # https://zhuanlan.zhihu.com/p/445847116              # BOM数据行列转换处理
    bom = pd.melt(dfb,id_vars=['组件','零件号','颜色特性描述'],value_vars=dfb.columns[5:dfb.shape[1]-2].values)
    bom = bom.rename(columns={'variable': 'Category','value': 'Qty'})
    bom.dropna(subset=['Qty'], inplace = True)            # 空值数据排除

uf_item = st.file_uploader("零件名称对照表",['xlsx', 'xls'])    #处理物料主文件
if uf_item is not None:
    dfi = pd.read_excel(uf_item)                          # 读取《物料》EXCEL文件
    dfi

uf_plan = st.file_uploader("细品计划表",['xlsx', 'xls'])   # 读取《红旗生产计划文件》
if uf_plan is not None:
    dfp = pd.read_excel(uf_plan,sheet_name='细品计划-原表',skiprows=1,usecols='B:AK')   # 读取红旗生产计划文件，并排出空行空列数据
    dfp = dfp[~dfp['车型系列'].str.contains('小计')]
    dfp = dfp[~dfp['车型系列'].str.contains('合计')]
    dfp = dfp.rename(columns={'车型系列': 'Category'})
    dfp = dfp.rename(columns={'销售编码': 'VMC'})
    dfp = dfp.rename(columns={'内饰颜色': '颜色特性描述'})
    dfp = dfp.fillna(0)
    dfp


if st.button("生成计划", type="primary"):
    r1 = pd.merge(bom, dfi, on='零件号')
    r1['Key1'] = r1['描述'].apply(descr_key1).astype(str)
    exp1 = pd.merge(dfp, r1, on=['Category','颜色特性描述'])
    cols = ['Category','VMC','颜色特性描述','组件','描述']
    cols = np.append(cols,dfp.columns[5:46].values)
    exp1[cols]
    excel_data = to_excel(exp1[cols])                                                             # 导出EXCEL文件
    file_name = "结果_细品生产计划.xlsx"                                                                # 命名导出文件：结果_生产计划.xlsx
    st.download_button(f"下载文件 {file_name}", excel_data, file_name, f"text/{file_name}", key=file_name) # 提供文件下载


