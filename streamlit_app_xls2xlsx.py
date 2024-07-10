import streamlit as st
from streamlit import session_state as ss
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import itertools
import glob
import os
import os.path           
# 将DataFrame压缩成一个zip文件
def dataframe_to_zip(df, filename):
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, 'w') as zip_file:
        zip_file.writestr(f"{filename}.csv", df.to_csv(index=False))
    zip_file_bytes = buffer.getvalue()
    buffer.close()
    return zip_file_bytes

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

st.title("格式转换")
#添加文件上传功能
uploaded_files = st.file_uploader("🟦上传原始数据文件",type=["xls","csv"], accept_multiple_files=True)

type_option = st.selectbox('✅需转换文件类型',('xls','csv'))
df_list = ['df' + str(i) for i in range(len(uploaded_files))]
if uploaded_files is not None:

    
    for i in range(len(uploaded_files)):
        file=uploaded_files[i]
        file_path = os.path.abspath(file.name)
        fordle_path = os.path.dirname(file_path)
        st.write("file_path",fordle_path)
        st.write(f"File name: {file.name}")
        type_option = file.name[-3:]
        if type_option.lower()=='csv':
            df_list[i] = pd.read_csv( file, low_memory=False,encoding = 'utf-8',encoding_errors='ignore')
            df_list[i].to_excel(fordle_path + '\\' + file.name[0:-3] + '.xlsx')

        if type_option.lower()=='xls':

            df_list[i] = pd.read_excel( file)
            df_list[i].to_excel(fordle_path + '\\' +file.name[0:-3] + '.xlsx')
    
    st.write("⚠️如果显示'TypeError: This COM object ... process...'，关闭进程中的excel重试 ")
        
    if st.checkbox('数据合并'):        
        #files = glob.glob(dirname + "*.xlsx")
        st.write('🟦合并结果：')
        #df_list = ['df' + str(i) for i in range(len(files))]
        #for i in range(len(files)):    
            #df_list[i] = pd.read_excel(files[i])
        dfa_result = pd.concat(df_list,keys = files)    
        df_xlsx = to_excel(dfa_result)
        st.download_button(label='📥 下载合并结果',
                                    data=df_xlsx ,
                                    file_name= '合并结果.xlsx')

    st.set_option('deprecation.showPyplotGlobalUse', False)#屏蔽警告