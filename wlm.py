"""Performance comparison with python, 1.0
   Yangmeichen 2020/6/19"""
import pandas as pd
from pandas import Series, DataFrame
from pandas import to_datetime
from datetime import datetime
import xlsxwriter



def wlm():
    filenamebefore = input('Enter wlmfilebefore：')
    filenameafter = input('Enter wlmfileafter：')
    # read data
    df1= pd.read_excel(filenamebefore,encoding='utf8')
    df2 = pd.read_excel(filenameafter, encoding='utf8')
    # 重新设置索引
    df1.reset_index()
    df2.reset_index()
    # 设置TIME为索引
    df1 = df1.set_index('TIME')
    df2 = df2.set_index('TIME')
    # 提取联机数据
    dfindex1 = pd.DataFrame(df1['09.00.00': '17.00.00'])
    dfindex2 = pd.DataFrame(df2['09.00.00': '17.00.00'])
    writer = pd.ExcelWriter('WLM.xlsx', engine='xlsxwriter',datetime_format='YYYY/MM/DD')
    dfindex1.to_excel(writer,sheet_name='before')
    dfindex2.to_excel(writer,sheet_name='after')
    workbook  = writer.book
    worksheet = writer.sheets['before']
    #RATE
    chartrate = workbook.add_chart({'type': 'line'})
    chartrate.height=350
    chartrate.width=500
    chartrate.add_series({
        'categories': '=before!$A$2:$A$34',
        'values': '=before!$E$2:$E$34',
        'name': '=before!$B$3',
    })
    chartrate.add_series({
    'values': '=after!$E$2:$E$34',
    'name': '=after!$B$3',
})
    chartrate.set_legend({'position': 'bottom'})
    chartrate.set_title({'name': 'TRAX RATE'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B2', chartrate)

    #RESP
    chartresp = workbook.add_chart({'type': 'line'})
    chartresp.height=350
    chartresp.width=500
    chartresp.add_series({
        'categories': '=before!$A$2:$A$34',
        'values': '=before!$F$2:$F$34',
        'name': '=before!$B$3',
    })
    chartresp.add_series({
        'values': '=after!$F$2:$F$34',
        'name': '=after!$B$3',
    })
    chartresp.set_legend({'position': 'bottom'})
    chartresp.set_title({'name': 'RESP TIME'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K2', chartresp)

    #MIPS
    chartmips = workbook.add_chart({'type': 'line'})
    chartmips.height=350
    chartmips.width=500
    chartmips.add_series({
        'categories': '=before!$A$2:$A$34',
        'values': '=before!$O$2:$O$34',
        'name': '=before!$B$3',
    })
    chartmips.add_series({
        'values': '=after!$O$2:$O$34',
        'name': '=after!$B$3',
    })
    chartmips.set_legend({'position': 'bottom'})
    chartmips.set_title({'name': 'MIPS/TRX'})
      # Insert the chart into the worksheet.
    worksheet.insert_chart('B22', chartmips)

    #MIPS/TRX CICS
    chartcics = workbook.add_chart({'type': 'line'})
    chartcics.height=350
    chartcics.width=500
    chartcics.add_series({
        'categories': '=before!$A$2:$A$34',
        'values': '=before!$L$2:$L$34',
        'name': '=before!$B$3',
    })
    chartcics.add_series({
        'values': '=after!$L$2:$L$34',
        'name': '=after!$B$3',
    })
    chartcics.set_legend({'position': 'bottom'})
    chartcics.set_title({'name': 'MIPS/TRX CICS'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K22', chartcics)

    ##MIPS/TRX DB2
    chartdb2 = workbook.add_chart({'type': 'line'})
    chartdb2.height=350
    chartdb2.width=500
    chartdb2.add_series({
        'categories': '=before!$A$2:$A$34',
        'values': '=before!$K$2:$K$34',
        'name': '=before!$B$3',
    })
    chartdb2.add_series({
        'values': '=after!$K$2:$K$34',
        'name': '=after!$B$3',
    })
    chartdb2.set_legend({'position': 'bottom'})
    chartdb2.set_title({'name': 'MIPS/TRX DB2'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B42', chartdb2)

    #MIPS/TRX SYSTEM
    chartsys = workbook.add_chart({'type': 'line'})
    chartsys.height=350
    chartsys.width=500
    chartsys.add_series({
        'categories': '=before!$A$2:$A$34',
        'values': '=before!$I$2:$I$34',
        'name': '=before!$B$3',
    })
    chartsys.add_series({
        'values': '=after!$I$2:$I$34',
        'name': '=after!$B$3',
    })
    chartsys.set_legend({'position': 'bottom'})
    chartsys.set_title({'name': 'MIPS/TRX SYSTEM'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K42', chartsys)

    #MIPS/TRX STC
    chartstc = workbook.add_chart({'type': 'line'})
    chartstc.height=350
    chartstc.width=500
    chartstc.add_series({
        'categories': '=before!$A$2:$A$34',
        'values': '=before!$J$2:$J$34',
        'name': '=before!$B$3',
    })
    chartstc.add_series({
        'values': '=after!$J$2:$J$34',
        'name': '=after!$B$3',
    })
    chartstc.set_legend({'position': 'bottom'})
    chartstc.set_title({'name': 'MIPS/TRX STC'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B62', chartstc)

    writer.save()







