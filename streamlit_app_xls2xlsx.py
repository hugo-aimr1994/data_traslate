import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import itertools
import tkinter as tk
from tkinter import filedialog

#mpl.font_manager.fontManager.addfont('./SimHei.ttf') #临时注册新的全局字体
plt.rcParams['font.sans-serif'] = ['SimHei'] # 步骤一（替换sans-serif字体）
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 18  #设置字体大小，全局有效

#设置页面标题
st.title("格式转换")
#添加文件上传功能
#uploaded_datafile = st.file_uploader("🟦上传原始数据文件",type=["xlsx","csv"])

root = tk.Tk()
root.withdraw()
 
# Make folder picker dialog appear on top of other windows
root.wm_attributes('-topmost', 1)

st.write('请选择文件夹:')
clicked = st.button('Folder Picker')
if clicked:
    dirname = st.text_input('🟦选择文件夹:', filedialog.askdirectory(master=root))

type_option = st.selectbox('✅输入文件格式',('csv','xls'))
st.write('🟦文件格式:', type_option)

combin_option = st.selectbox('✅输出结果是否合并',('是','否'))
st.write('🟦合并:', combin_option)

def xls_to_xlsx(rootdir):
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    for parent, dirnames, filenames in os.walk(rootdir):
        for fn in filenames:
            filedir = os.path.join(parent, fn)
            #print(filedir)

            excel = win32.Dispatch('Excel.Application')
            wb = excel.Workbooks.Open(filedir)
            # xlsx: FileFormat=51
            # xls:  FileFormat=56,
            # 后缀名的大小写不通配，需按实际修改：xls，或XLS
            wb.SaveAs(filedir.replace('xls', 'xlsx').replace('XLS', 'xlsx'), FileFormat=51)  # 我这里原文件是大写
            wb.Close()                                 
            excel.Application.Quit()

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

#文件导入
#如果有文件上传，显示数据表格
if dirname is not None:
    #file_name = uploaded_datafile.name
    files = glob.glob(dirname + "*." + type_option)
    file_names = [file[0:-4] for file in files]
    if file_type == 'csv':
        #dfa = pd.read_csv(files[0], low_memory=False,encoding = 'utf-8',encoding_errors='ignore')
        df_list = ['df' + str(i) for i in range(len(files))]

        for i in range(len(files)):    
            #df_list[i] = pd.read_csv(files[i], low_memory=False,encoding = 'gbk')
            df_list[i] = pd.read_csv( files[i], low_memory=False,encoding = 'utf-8',encoding_errors='ignore')
            df_list[i].to_excel(file_names[i] + '.xlsx')

    elif file_type == 'xls':
        xls_to_xlsx(dirname)

    if combin_option == '是':        
        files = glob.glob(dirname + "*.xlsx")
        st.write('🟦合并结果：')
        df_list = ['df' + str(i) for i in range(len(files))]
        for i in range(len(files)):    
            df_list[i] = pd.read_excel(files[i])
        dfa_result = pd.concat(df_list,keys = files)    
        df_xlsx = to_excel(dfa_result)
        st.download_button(label='📥 下载合并结果',
                                    data=df_xlsx ,
                                    file_name= '合并结果.xlsx')
    
    st.set_option('deprecation.showPyplotGlobalUse', False)#屏蔽警告