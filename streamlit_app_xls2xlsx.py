import streamlit as st
from streamlit import session_state as ss
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import itertools
import glob
import os
import os.path

#mpl.font_manager.fontManager.addfont('./SimHei.ttf') #临时注册新的全局字体
#plt.rcParams['font.sans-serif'] = ['SimHei'] # 步骤一（替换sans-serif字体）
#plt.rcParams['axes.unicode_minus'] = False
#plt.rcParams['font.size'] = 18  #设置字体大小，全局有效
def xls_to_xlsx(rootdir):
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    for parent, dirnames, filenames in os.walk(rootdir):
        for fn in filenames:
            filedir = os.path.join(parent, fn)
            print(filedir)

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(filedir)
            # xlsx: FileFormat=51
            # xls:  FileFormat=56,
            # 后缀名的大小写不通配，需按实际修改：xls，或XLS            
            wb.SaveAs(filedir.replace('xls', 'xlsx').replace('XLS', 'xlsx').replace('/', '\\'), FileFormat=51)            
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

st.title("格式转换")
#添加文件上传功能
#uploaded_file = st.file_uploader("🟦上传原始数据文件",type=["xls","csv"])
#st.write('🟦文件路径:', uploaded_file.name)
# 用户输入文件路径
file_path = st.text_input('请输入文件路径，如D:\python:')
#file_path = file_path.replace("\","//")
type_option = st.selectbox('✅需转换文件类型',('xls','csv'))
# 检查文件是否存在
if file_path and os.path.exists(file_path):
#if uploaded_file is not None:
    # 获取文件路径
    #file_path = os.path.abspath(os.path.join(uploaded_file.name))
    
    #st.write('🟦文件路径:', file_path) 
    # 获取文件夹路径
    dirname = file_path.replace('\\','/') + '/'
    st.write('🟦文件夹路径:', dirname)    
    #type_option = file_path[-3:]
    st.write('🟦文件格式:', type_option)
    #file_name = uploaded_datafile.name
    files = glob.glob(dirname + "*." + type_option)
    st.write('🟦导入文件：',files)
    file_names = [file[0:-4] for file in files]
    df_list = ['df' + str(i) for i in range(len(files))]
    if type_option.lower()=='csv':
        for i in range(len(files)):    
            #df_list[i] = pd.read_csv(files[i], low_memory=False,encoding = 'gbk')
            df_list[i] = pd.read_csv( files[i], low_memory=False,encoding = 'utf-8',encoding_errors='ignore')
            df_list[i].to_excel(file_names[i] + '.xlsx')

    if type_option.lower()=='xls':
        for i in range(len(files)):   
            df_list[i] = pd.read_excel( files[i])
            df_list[i].to_excel(file_names[i] + '.xlsx')
    st.write("⚠️如果显示'TypeError: This COM object ... process...'，关闭进程中的excel重试 ")
    if st.checkbox('数据合并'):        
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