import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def descr_key1(x):
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
    in_memory_fp = BytesIO()
    df.to_excel(in_memory_fp, index=False)
    # Write the file out to disk to demonstrate that it worked.
    in_memory_fp.seek(0, 0)
    return in_memory_fp.read()

st.title('计划分解系统v1.0.1')
st.header('恒展迈思软件')
st.text('该系统会根据一汽红旗计划与BOM系统自动分解计划与系统配置文件')

uf_bom = st.file_uploader("BOM",['xlsx', 'xls'])
if uf_bom is not None:
    dfb = pd.read_excel(uf_bom,sheet_name='BOM清单')
    dfb
    dfb = dfb.rename(columns={"去-零件号": "零件号"})
    # Pandas行列转换的4大技巧 - 尤而小屋的文章 - 知乎
    # https://zhuanlan.zhihu.com/p/445847116
    bom = pd.melt(dfb,id_vars=['组件','零件号','颜色特性描述'],value_vars=dfb.columns[5:dfb.shape[1]-2].values)
    bom = bom.rename(columns={'variable': 'Category','value': 'Qty'})
    bom.dropna(subset=['Qty'], inplace = True)

uf_item = st.file_uploader("物料主文件",['xlsx', 'xls'])
if uf_item is not None:
    dfi = pd.read_excel(uf_item)
    dfi

uf_plan = st.file_uploader("红旗生产计划",['xlsx', 'xls'])
if uf_plan is not None:
    dfp = pd.read_excel(uf_plan,sheet_name='M+6-原表',skiprows=5,usecols='C:L')
    dfp = dfp.rename(columns={'Unnamed: 2': 'Category'})
    dfp = dfp.rename(columns={'Unnamed: 3': 'VMC'})
    dfp = dfp.rename(columns={'Unnamed: 4': 'Color'})
    dfp = dfp.rename(columns={'Unnamed: 5': '颜色特性描述'})
    dfp.dropna(subset=['Category'], inplace = True)
    dfp

if st.button('生成配置'):
    r1 = pd.merge(bom, dfi, on='零件号')
    r1['Key1'] = r1['描述'].apply(descr_key1).astype(str)
    exp1 = pd.merge(dfp, r1, on=['Category','颜色特性描述'])
    #exp1.to_excel('exp1.xlsx',columns=['VMC','零件号','描述','Qty','Key1'],index=False)  
    exp1[['VMC','零件号','描述','Qty','Key1']]
    excel_data = to_excel(exp1[['VMC','零件号','描述','Qty','Key1']])
    file_name = "结果_配置文件.xlsx"
    st.download_button(f"下载文件 {file_name}", excel_data, file_name, f"text/{file_name}", key=file_name)

if st.button("生成计划", type="primary"):
    r1 = pd.merge(bom, dfi, on='零件号')
    r1['Key1'] = r1['描述'].apply(descr_key1).astype(str)
    exp1 = pd.merge(dfp, r1, on=['Category','颜色特性描述'])
    cols = ['Category','VMC','组件','描述']
    cols = np.append(cols,exp1.columns[3:10].values)
    #exp1.to_excel('exp2.xlsx',columns=cols)
    exp1[cols]
    excel_data = to_excel(exp1[cols])
    file_name = "结果_生产计划.xlsx"
    st.download_button(f"下载文件 {file_name}", excel_data, file_name, f"text/{file_name}", key=file_name)
