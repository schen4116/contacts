import pandas as pd
import streamlit as st
import os
import numpy as np
import time

st.set_page_config(page_title="Contact", page_icon=":shark:", layout="wide") #"random"
#For Excel File
sh_name = None #"职能", None
pd.set_option('display.max_columns', None)
fn = os.path.dirname(os.path.realpath(__file__))+f"/2022.01.xlsx"

@st.cache(persist=True)
def get_data():
    df = pd.read_excel(fn, skiprows = [0], header = 0,  keep_default_na=False, dtype={'姓名':'category'}, sheet_name=sh_name) #usecols = "C:I",
    if sh_name is None:
        for key in df.keys():
            cols = df[key].columns.tolist()
            rename_dict = {}
            for col in cols:
                rename_dict[col]=col.replace(" ", "")
            df[key].rename(columns=rename_dict,inplace=True)
            df[key].rename(columns={"办公室":"办公室/卡位"},inplace=True)
        df = pd.concat(df)
    df = df.iloc[:,[i for i in range(min(len(df.columns),9))]]
    # df[0].fillna(method='pad',inplace=True)
    cols = df.columns.tolist()
    print(f"{cols}")
    cols = cols[2:]+cols[:2]
    df = df[cols]
    df = df.drop("序号",1)
    df[df[["单元"]]==""] = np.NaN
    df[['单元']]=df[['单元']].ffill()
    # df[['单元']] = df[['单元']].fillna(method='ffill',inplace=True)
    
    df.fillna('', inplace=True)
    return df

df = get_data()
print(f"{df.columns}")
# df = df[['工号', '姓名', "办公室/卡位","办公电话","移动电话","邮箱","岗位"]]
fn_time = time.strftime('%Y-%m', time.gmtime(os.path.getmtime(fn)))
st.title(f"Contacts (updated on: {fn_time})")
key = st.text_input("") #placeholder="输入关键字以查找..."
st.text("* 用法：（1）关键词：查询记录; （2）@+研究所：查询成员；（3）空白：查询统计")
# col1, col2 = st.columns([1,5])
# with col1:
#     opt = st.selectbox("选择查询项", tuple(df.columns), index=1)
# with col2:
#     name = st.text_input("",placeholder="输入关键字以查询")
# print(f"{df.index.get_level_values(0).drop_duplicates().values}")
def highlight(s):
    if s.中心 == "":
        return ['background-color: yellow'] * len(s)
    else:
        return ['background-color: white'] * len(s)

# col1, col2 = st.columns(2)

# with col1:
#     key = st.text_input("") #placeholder="输入关键字以查找..."
# with col2:
#     st.text("关键词->查询记录; @+所部->查询成员; 空白->查询统计")

if key=="":
    df_sel = df
    st.info(f'共找到{len(df_sel.index)}条记录，按研究单元统计情况：')
    depts = df.index.get_level_values(0).drop_duplicates().values
    dat = []
    for dept in depts:
        
        dfi = df.loc[(dept,)]
        dat.append([f"{dept} ({dfi['姓名'].values[0]})","",len(dfi.index)])
        vals = set(dfi['单元'].tolist())
        datj=[]
        for val in vals:
            df_ij = dfi[dfi['单元']==val]
            szi = len(df_ij.index)
            df_k = df_ij[(df_ij['岗位'].str.contains("处长",na=False) & ~df_ij['岗位'].str.contains("副处长",na=False)) | df_ij['岗位'].str.contains("中心主任",na=False) \
                | df_ij['岗位'].str.contains("主持工作",na=False) | df_ij['岗位'].str.contains("执行主任",na=False) | df_ij['岗位'].str.contains("院长",na=False) \
                | df_ij['岗位'].str.contains("负责人",na=False) | df_ij['岗位'].str.contains("部长",na=False)]
            val = val.replace("\n","")
            if len(val)<50:
                li = df_k['姓名'].tolist()
                brch = f"{val} (负责人: {li[0] })" if (len(li)>0 and val!="所部") else val
                datj.append(["",brch,szi])
                # datj.append(["",val,szi])
        datj.sort(key=lambda x: x[2],reverse=True)
        for dj in datj:
            dat.append(dj)
    dfj = pd.DataFrame(dat, columns = ['研究所', '中心','人数'])
    dfj = dfj.style.apply(highlight, axis=1)
    # CSS to inject contained in a string
    hide_table_row_index = """
                <style>
                tbody th {display:none}
                .blank {display:none}
                </style>
                """
    # Inject CSS with Markdown
    st.markdown(hide_table_row_index, unsafe_allow_html=True)
    st.table(dfj)
    st.info(f'详细通讯录：')
elif key.find("@")!=-1:
    df_sel = df.loc[(key[1:],)]
    st.info(f'找到{len(df_sel.index)}条记录：')
else:
    cond = None
    for opt in list(df.columns):
        print(f"{opt}")
        if cond is None:
            cond = df[opt].str.contains(key, na=False)
        else:
            cond = cond | df[opt].str.contains(key, na=False)
    df_sel = df[cond]
    st.info(f'找到{len(df_sel.index)}条记录：')
    
st.table(df_sel.astype(str)) #.astype(str)