import openpyxl
import pandas as pd

import streamlit as st
from io import BytesIO

st.title('Execl Automater')

data=st.file_uploader('Upload your execl file')

buffer = BytesIO()

if data!=None:
    df=pd.read_excel(data)
    st.write(df)
    index=st.multiselect('Select Pivot Table Index',list(df.columns))
    columns=st.multiselect('Select Column for Pivot Table',list(df.columns))
    values=st.selectbox('Select column you want to aggergate ',list(df.columns))
    filename=st.text_input('Enter execl file name')
    
    indices=[]
    indices.extend(index)
    indices.extend(columns)
    indices.append(values)

    st.write(index)

    df_new=df[indices]

    aggfunction=st.selectbox('Selcet aggeration Function',['sum','mean','max','min'])

    button=st.button('Generate Pivot Table')   
    if button:
        pivot_table=df_new.pivot_table(index=index,columns=columns,values=values,aggfunc=aggfunction)
        st.write(pivot_table)
    
    
  



        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            
            
    # Write each dataframe to a different worksheet.
                pivot_table.to_excel(writer, sheet_name='Report',startrow=3)

    # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.save()
                st.download_button(
                    label="Download Excel worksheets",
                    data=buffer,
                    file_name=f"{filename}.xlsx",
                    mime="application/vnd.ms-excel"
                     )












    







#st.write(pivot_table)