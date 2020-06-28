
"""Performance comparison with python, 1.0
   Yangmeichen 2020/6/22"""
import pandas as pd
from pandas import Series, DataFrame
from pandas import to_datetime
from datetime import datetime
import xlsxwriter

def cf():
    filenamebefore=input('Enter cfutilbefore name:')
    filenameafter=input('Enter cfutilafter name:')
    filenamecfbefore=input('Enter cflinkbefore name:')
    filenamecfafter=input('Enter cflinkafter name:')
    # read data
    df1= pd.read_excel(filenamebefore,encoding='utf8')
    df2 = pd.read_excel(filenameafter, encoding='utf8')
    df1.reset_index()
    df2.reset_index()
        # 设置TIME为索引
    df1 = df1.set_index('TIME')
    df2 = df2.set_index('TIME')
        # 提取数据
    df1 = pd.DataFrame(df1['00.00.00': '17.00.00'])
    df2 = pd.DataFrame(df2['00.00.00': '17.00.00'])
    dfspc1c2b=df1[(df1.CF_NAME == 'SPC1C2')]
    dfspc2c2b=df1[(df1.CF_NAME == 'SPC2C2')]
    dfspc1c2a=df2[(df2.CF_NAME == 'SPC1C2')]
    dfspc2c2a=df2[(df2.CF_NAME == 'SPC2C2')]

    dfcflinkb= pd.read_excel(filenamecfbefore,encoding='utf8')
    dfcflinka = pd.read_excel(filenamecfafter, encoding='utf8')
    dfcflinkb.reset_index()
    dfcflinka.reset_index()
        # 设置TIME为索引
    dfcflinkb =dfcflinkb.set_index('TIME')
    dfcflinka =dfcflinka.set_index('TIME')
        # 提取数据
    dfcflinkb = pd.DataFrame(dfcflinkb['00.00.00': '17.00.00'])
    dfcflinka = pd.DataFrame(dfcflinka['00.00.00': '17.00.00'])

    dfpp1ac1c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1A')&(dfcflinkb.CF_NAME == 'SPC1C2')]
    dfpp1bc1c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1B')&(dfcflinkb.CF_NAME == 'SPC1C2')]
    dfpp1cc1c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1C')&(dfcflinkb.CF_NAME == 'SPC1C2')]
    dfpp1dc1c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1D')&(dfcflinkb.CF_NAME == 'SPC1C2')]
    dfpp1ec1c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1E')&(dfcflinkb.CF_NAME == 'SPC1C2')]
    dfpp1fc1c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1F')&(dfcflinkb.CF_NAME == 'SPC1C2')]
    dfpp1gc1c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1G')&(dfcflinkb.CF_NAME == 'SPC1C2')]
    dfpp1hc1c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1H')&(dfcflinkb.CF_NAME == 'SPC1C2')]

    dfpp1ac1c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1A')&(dfcflinka.CF_NAME == 'SPC1C2')]
    dfpp1bc1c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1B')&(dfcflinka.CF_NAME == 'SPC1C2')]
    dfpp1cc1c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1C')&(dfcflinka.CF_NAME == 'SPC1C2')]
    dfpp1dc1c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1D')&(dfcflinka.CF_NAME == 'SPC1C2')]
    dfpp1ec1c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1E')&(dfcflinka.CF_NAME == 'SPC1C2')]
    dfpp1fc1c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1F')&(dfcflinka.CF_NAME == 'SPC1C2')]
    dfpp1gc1c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1G')&(dfcflinka.CF_NAME == 'SPC1C2')]
    dfpp1hc1c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1H')&(dfcflinka.CF_NAME == 'SPC1C2')]

    dfpp1ac2c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1A')&(dfcflinkb.CF_NAME == 'SPC2C2')]
    dfpp1bc2c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1B')&(dfcflinkb.CF_NAME == 'SPC2C2')]
    dfpp1cc2c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1C')&(dfcflinkb.CF_NAME == 'SPC2C2')]
    dfpp1dc2c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1D')&(dfcflinkb.CF_NAME == 'SPC2C2')]
    dfpp1ec2c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1E')&(dfcflinkb.CF_NAME == 'SPC2C2')]
    dfpp1fc2c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1F')&(dfcflinkb.CF_NAME == 'SPC2C2')]
    dfpp1gc2c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1G')&(dfcflinkb.CF_NAME == 'SPC2C2')]
    dfpp1hc2c2b=dfcflinkb[(dfcflinkb.SYSNAME == 'PP1H')&(dfcflinkb.CF_NAME == 'SPC2C2')]

    dfpp1ac2c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1A')&(dfcflinka.CF_NAME == 'SPC2C2')]
    dfpp1bc2c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1B')&(dfcflinka.CF_NAME == 'SPC2C2')]
    dfpp1cc2c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1C')&(dfcflinka.CF_NAME == 'SPC2C2')]
    dfpp1dc2c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1D')&(dfcflinka.CF_NAME == 'SPC2C2')]
    dfpp1ec2c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1E')&(dfcflinka.CF_NAME == 'SPC2C2')]
    dfpp1fc2c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1F')&(dfcflinka.CF_NAME == 'SPC2C2')]
    dfpp1gc2c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1G')&(dfcflinka.CF_NAME == 'SPC2C2')]
    dfpp1hc2c2a=dfcflinka[(dfcflinka.SYSNAME == 'PP1H')&(dfcflinka.CF_NAME == 'SPC2C2')]
    writer = pd.ExcelWriter('CFCPU.xlsx', engine='xlsxwriter', datetime_format='YYYY/MM/DD')
    dfspc1c2b.to_excel(writer, sheet_name='SPC1C2BEFORE')
    dfspc1c2a.to_excel(writer, sheet_name='SPC1C2AFTER')
    dfspc2c2b.to_excel(writer, sheet_name='SPC2C2BEFORE')
    dfspc2c2a.to_excel(writer, sheet_name='SPC2C2AFTER')

    dfpp1ac1c2b.to_excel(writer, sheet_name='PP1A')
    dfpp1bc1c2b.to_excel(writer, sheet_name='PP1B')
    dfpp1cc1c2b.to_excel(writer, sheet_name='PP1C')
    dfpp1dc1c2b.to_excel(writer, sheet_name='PP1D')
    dfpp1ec1c2b.to_excel(writer, sheet_name='PP1E')
    dfpp1fc1c2b.to_excel(writer, sheet_name='PP1F')
    dfpp1gc1c2b.to_excel(writer, sheet_name='PP1G')
    dfpp1hc1c2b.to_excel(writer, sheet_name='PP1H')

    dfpp1ac1c2a.to_excel(writer, sheet_name='PP1A',startrow = 71)
    dfpp1bc1c2a.to_excel(writer, sheet_name='PP1B',startrow = 71)
    dfpp1cc1c2a.to_excel(writer, sheet_name='PP1C',startrow = 71)
    dfpp1dc1c2a.to_excel(writer, sheet_name='PP1D',startrow = 71)
    dfpp1ec1c2a.to_excel(writer, sheet_name='PP1E',startrow = 71)
    dfpp1fc1c2a.to_excel(writer, sheet_name='PP1F',startrow = 71)
    dfpp1gc1c2a.to_excel(writer, sheet_name='PP1G',startrow = 71)
    dfpp1hc1c2a.to_excel(writer, sheet_name='PP1H',startrow = 71)

    dfpp1ac2c2b.to_excel(writer, sheet_name='PP1A',startrow = 142)
    dfpp1bc2c2b.to_excel(writer, sheet_name='PP1B',startrow = 142)
    dfpp1cc2c2b.to_excel(writer, sheet_name='PP1C',startrow = 142)
    dfpp1dc2c2b.to_excel(writer, sheet_name='PP1D',startrow = 142)
    dfpp1ec2c2b.to_excel(writer, sheet_name='PP1E',startrow = 142)
    dfpp1fc2c2b.to_excel(writer, sheet_name='PP1F',startrow = 142)
    dfpp1gc2c2b.to_excel(writer, sheet_name='PP1G',startrow = 142)
    dfpp1hc2c2b.to_excel(writer, sheet_name='PP1H',startrow = 142)

    dfpp1ac2c2a.to_excel(writer, sheet_name='PP1A',startrow = 213)
    dfpp1bc2c2a.to_excel(writer, sheet_name='PP1B',startrow = 213)
    dfpp1cc2c2a.to_excel(writer, sheet_name='PP1C',startrow = 213)
    dfpp1dc2c2a.to_excel(writer, sheet_name='PP1D',startrow = 213)
    dfpp1ec2c2a.to_excel(writer, sheet_name='PP1E',startrow = 213)
    dfpp1fc2c2a.to_excel(writer, sheet_name='PP1F',startrow = 213)
    dfpp1gc2c2a.to_excel(writer, sheet_name='PP1G',startrow = 213)
    dfpp1hc2c2a.to_excel(writer, sheet_name='PP1H',startrow = 213)
    workbook  = writer.book
    worksheet = writer.sheets['SPC1C2BEFORE']
    #spc1c2
    chartspc1c2 = workbook.add_chart({'type': 'line'})
    chartspc1c2.height=350
    chartspc1c2.width=500
    chartspc1c2.add_series({
        'categories': '=SPC1C2BEFORE!$A$2:$A$70',
        'values': '=SPC1C2BEFORE!$J$2:$J$70',
        'name': '=SPC1C2BEFORE!$B$2',
    })
    chartspc1c2.add_series({
        'values': '=SPC1C2AFTER!$J$2:$J$70',
        'name': '=SPC1C2AFTER!$B$2',
    })
    chartspc1c2.set_legend({'position': 'bottom'})
    chartspc1c2.set_title({'name': 'SPC1C2-CF CPU%'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B2', chartspc1c2)

    #spc2c2
    chartspc2c2 = workbook.add_chart({'type': 'line'})
    chartspc2c2.height=350
    chartspc2c2.width=500
    chartspc2c2.add_series({
        'categories': '=SPC2C2BEFORE!$A$2:$A$70',
        'values': '=SPC2C2BEFORE!$J$2:$J$70',
        'name': '=SPC2C2BEFORE!$B$2',
    })
    chartspc2c2.add_series({
        'values': '=SPC2C2AFTER!$J$2:$J$70',
        'name': '=SPC2C2AFTER!$B$2',
    })
    chartspc2c2.set_legend({'position': 'bottom'})
    chartspc2c2.set_title({'name': 'SPC2C2-CF CPU%'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K2', chartspc2c2)

    #SPC1C2 SYNA REQ
    chartc1c2syncreq = workbook.add_chart({'type': 'line'})
    chartc1c2syncreq.height=350
    chartc1c2syncreq.width=500
    chartc1c2syncreq.add_series({
        'categories': '=(PP1A!$C$2:$C$70,PP1A!$C$73:$C$141)',
        'values': '=(PP1A!$L$2:$L$70,PP1A!$L$73:$L$141)',
        'name': '=PP1A!$E$2',
    })
    chartc1c2syncreq.add_series({
        'values': '=(PP1B!$L$2:$L$70,PP1B!$L$73:$L$141)',
        'name': '=PP1B!$E$3',
    })
    chartc1c2syncreq.add_series({
        'values': '=(PP1C!$L$2:$L$70,PP1C!$L$73:$L$141)',
        'name': '=PP1C!$E$3',
    })
    chartc1c2syncreq.add_series({
        'values': '=(PP1D!$L$2:$L$70,PP1D!$L$73:$L$141)',
        'name': '=PP1D!$E$3',
    })
    chartc1c2syncreq.add_series({
        'values': '=(PP1E!$L$2:$L$70,PP1E!$L$73:$L$141)',
        'name': '=PP1E!$E$3',
    })
    chartc1c2syncreq.add_series({
        'values': '=(PP1F!$L$2:$L$70,PP1F!$L$73:$L$141)',
        'name': '=PP1F!$E$3',
    })
    chartc1c2syncreq.add_series({
        'values': '=(PP1G!$L$2:$L$70,PP1G!$L$73:$L$141)',
        'name': '=PP1G!$E$3',
    })
    chartc1c2syncreq.add_series({
        'values': '=(PP1H!$L$2:$L$70,PP1H!$L$73:$L$141)',
        'name': '=PP1H!$E$3',
    })
    chartc1c2syncreq.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartc1c2syncreq.set_legend({'position': 'bottom'})
    chartc1c2syncreq.set_title({'name': 'SPC1C2-SYNC-REQ NUMBER'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B22', chartc1c2syncreq)

    #SPC1C2 SYNC RESP
    chartc1c2syncresp = workbook.add_chart({'type': 'line'})
    chartc1c2syncresp.height=350
    chartc1c2syncresp.width=500
    chartc1c2syncresp.add_series({
        'categories': '=(PP1A!$C$2:$C$70,PP1A!$C$73:$C$141)',
        'values': '=(PP1A!$M$2:$M$70,PP1A!$M$73:$M$141)',
        'name': '=PP1A!$E$2',
    })
    chartc1c2syncresp.add_series({
        'values': '=(PP1B!$M$2:$M$70,PP1B!$M$73:$M$141)',
        'name': '=PP1B!$E$3',
    })
    chartc1c2syncresp.add_series({
        'values': '=(PP1C!$M$2:$M$70,PP1C!$M$73:$M$141)',
        'name': '=PP1C!$E$3',
    })
    chartc1c2syncresp.add_series({
        'values': '=(PP1D!$M$2:$M$70,PP1D!$M$73:$M$141)',
        'name': '=PP1D!$E$3',
    })
    chartc1c2syncresp.add_series({
        'values': '=(PP1E!$M$2:$M$70,PP1E!$M$73:$M$141)',
        'name': '=PP1E!$E$3',
    })
    chartc1c2syncresp.add_series({
        'values': '=(PP1F!$M$2:$M$70,PP1F!$M$73:$M$141)',
        'name': '=PP1F!$E$3',
    })
    chartc1c2syncresp.add_series({
        'values': '=(PP1G!$M$2:$M$70,PP1G!$M$73:$M$141)',
        'name': '=PP1G!$E$3',
    })
    chartc1c2syncresp.add_series({
        'values': '=(PP1H!$M$2:$M$70,2PP1H!$M$73:$M$141)',
        'name': '=PP1H!$E$3',
    })
    chartc1c2syncresp.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartc1c2syncresp.set_legend({'position': 'bottom'})
    chartc1c2syncresp.set_title({'name': 'SPC1C2-SYNC RESP TIME'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K22', chartc1c2syncresp)

    #SPC1C2 ASYNA REQ
    chartc1c2asyncreq = workbook.add_chart({'type': 'line'})
    chartc1c2asyncreq.height=350
    chartc1c2asyncreq.width=500
    chartc1c2asyncreq.add_series({
        'categories': '=(PP1A!$C$2:$C$70,PP1A!$C$73:$C$141)',
        'values': '=(PP1A!$N$2:$N$70,PP1A!$N$73:$N$141)',
        'name': '=PP1A!$E$2',
    })
    chartc1c2asyncreq.add_series({
        'values': '=(PP1B!$N$2:$N$70,PP1B!$N$73:$N$141)',
        'name': '=PP1B!$E$3',
    })
    chartc1c2asyncreq.add_series({
        'values': '=(PP1C!$N$2:$N$70,PP1C!$N$73:$N$141)',
        'name': '=PP1C!$E$3',
    })
    chartc1c2asyncreq.add_series({
        'values': '=(PP1D!$N$2:$N$70,PP1D!$N$73:$N$141)',
        'name': '=PP1D!$E$3',
    })
    chartc1c2asyncreq.add_series({
        'values': '=(PP1E!$N$2:$N$70,PP1E!$N$73:$N$141)',
        'name': '=PP1E!$E$3',
    })
    chartc1c2asyncreq.add_series({
        'values': '=(PP1F!$N$2:$N$70,PP1F!$N$73:$N$141)',
        'name': '=PP1F!$E$3',
    })
    chartc1c2asyncreq.add_series({
        'values': '=(PP1G!$N$2:$N$70,PP1G!$N$73:$N$141)',
        'name': '=PP1G!$E$3',
    })
    chartc1c2asyncreq.add_series({
        'values': '=(PP1H!$N$2:$N$70,PP1H!$N$73:$N$141)',
        'name': '=PP1H!$E$3',
    })
    chartc1c2asyncreq.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartc1c2asyncreq.set_legend({'position': 'bottom'})
    chartc1c2asyncreq.set_title({'name': 'SPC1C2-ASYNC-REQ NUMBER'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B42', chartc1c2asyncreq)

    #SPC1C2 ASYNC RESP
    chartc1c2asyncresp = workbook.add_chart({'type': 'line'})
    chartc1c2asyncresp.height=350
    chartc1c2asyncresp.width=500
    chartc1c2asyncresp.add_series({
        'categories': '=(PP1A!$C$2:$C$70,PP1A!$C$73:$C$141)',
        'values': '=(PP1A!$O$2:$O$70,PP1A!$O$73:$O$141)',
        'name': '=PP1A!$E$2',
    })
    chartc1c2asyncresp.add_series({
        'values': '=(PP1B!$O$2:$O$70,PP1B!$O$73:$O$141)',
        'name': '=PP1B!$E$3',
    })
    chartc1c2asyncresp.add_series({
        'values': '=(PP1C!$O$2:$O$70,PP1C!$O$73:$O$141)',
        'name': '=PP1C!$E$3',
    })
    chartc1c2asyncresp.add_series({
        'values': '=(PP1D!$O$2:$O$70,PP1D!$O$73:$O$141)',
        'name': '=PP1D!$E$3',
    })
    chartc1c2asyncresp.add_series({
        'values': '=(PP1E!$O$2:$O$70,PP1E!$O$73:$O$141)',
        'name': '=PP1E!$E$3',
    })
    chartc1c2asyncresp.add_series({
        'values': '=(PP1F!$O$2:$O$70,PP1F!$O$73:$O$141)',
        'name': '=PP1F!$E$3',
    })
    chartc1c2asyncresp.add_series({
        'values': '=(PP1G!$O$2:$O$70,PP1G!$O$73:$O$141)',
        'name': '=PP1G!$E$3',
    })
    chartc1c2asyncresp.add_series({
        'values': '=(PP1H!$O$2:$O$70,PP1H!$O$73:$O$141)',
        'name': '=PP1H!$E$3',
    })
    chartc1c2asyncresp.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartc1c2asyncresp.set_legend({'position': 'bottom'})
    chartc1c2asyncresp.set_title({'name': 'SPC1C2-ASYNC RESP TIME'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K42', chartc1c2asyncresp)

    #SPC2C2 SYNA REQ
    chartc2c2syncreq = workbook.add_chart({'type': 'line'})
    chartc2c2syncreq.height=350
    chartc2c2syncreq.width=500
    chartc2c2syncreq.add_series({
        'categories': '=(PP1A!$C$144:$C$212,PP1A!$C$215:$C$283)',
        'values': '=(PP1A!$L$144:$L$212,PP1A!$L$215:$L$283)',
        'name': '=PP1A!$E$212',
    })
    chartc2c2syncreq.add_series({
        'values': '=(PP1B!$L$144:$L$212,PP1B!$L$215:$L$283)',
        'name': '=PP1B!$E$212',
    })
    chartc2c2syncreq.add_series({
        'values': '=(PP1C!$L$144:$L$212,PP1C!$L$215:$L$283)',
        'name': '=PP1C!$E$212',
    })
    chartc2c2syncreq.add_series({
        'values': '=(PP1D!$L$144:$L$212,PP1D!$L$215:$L$283)',
        'name': '=PP1D!$E$212',
    })
    chartc2c2syncreq.add_series({
        'values': '=(PP1E!$L$144:$L$212,PP1E!$L$215:$L$283)',
        'name': '=PP1E!$E$212',
    })
    chartc2c2syncreq.add_series({
        'values': '=(PP1F!$L$144:$L$212,PP1F!$L$215:$L$283)',
        'name': '=PP1F!$E$212',
    })
    chartc2c2syncreq.add_series({
        'values': '=(PP1G!$L$144:$L$212,PP1G!$L$215:$L$283)',
        'name': '=PP1G!$E$212',
    })
    chartc2c2syncreq.add_series({
        'values': '=(PP1H!$L$144:$L$212,PP1H!$L$215:$L$283)',
        'name': '=PP1H!$E$212',
    })
    chartc2c2syncreq.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartc2c2syncreq.set_legend({'position': 'bottom'})
    chartc2c2syncreq.set_title({'name': 'SPC2C2-SYNC-REQ NUMBER'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B62', chartc2c2syncreq)

    #SPC2C2 SYNC RESP
    chartc2c2syncresp = workbook.add_chart({'type': 'line'})
    chartc2c2syncresp.height=350
    chartc2c2syncresp.width=500
    chartc2c2syncresp.add_series({
        'categories': '=(PP1A!$C$144:$C$212,PP1A!$C$215:$C$283)',
        'values': '=(PP1A!$M$144:$M$212,PP1A!$M$215:$M$283)',
        'name': '=PP1A!$E$212',
    })
    chartc2c2syncresp.add_series({
        'values': '=(PP1B!$M$144:$M$212,PP1B!$M$215:$M$283)',
        'name': '=PP1B!$E$212',
    })
    chartc2c2syncresp.add_series({
        'values': '=(PP1C!$M$144:$M$212,PP1C!$M$215:$M$283)',
        'name': '=PP1C!$E$212',
    })
    chartc2c2syncresp.add_series({
        'values': '=(PP1D!$M$144:$M$212,PP1D!$M$215:$M$283)',
        'name': '=PP1D!$E$212',
    })
    chartc2c2syncresp.add_series({
        'values': '=(PP1E!$M$144:$M$212,PP1E!$M$215:$M$283)',
        'name': '=PP1E!$E$212',
    })
    chartc2c2syncresp.add_series({
        'values': '=(PP1F!$M$144:$M$212,PP1F!$M$215:$M$283)',
        'name': '=PP1F!$E$212',
    })
    chartc2c2syncresp.add_series({
        'values': '=(PP1G!$M$144:$M$212,PP1G!$M$215:$M$283)',
        'name': '=PP1G!$E$212',
    })
    chartc2c2syncresp.add_series({
        'values': '=(PP1H!$M$144:$M$212,PP1H!$M$215:$M$283)',
        'name': '=PP1H!$E$212',
    })
    chartc2c2syncresp.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartc2c2syncresp.set_legend({'position': 'bottom'})
    chartc2c2syncresp.set_title({'name': 'SPC2C2-SYNC RESP TIME'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K62', chartc2c2syncresp)

    #SPC2C2 ASYNA REQ
    chartc2c2asyncreq = workbook.add_chart({'type': 'line'})
    chartc2c2asyncreq.height=350
    chartc2c2asyncreq.width=500
    chartc2c2asyncreq.add_series({
        'categories': '=(PP1A!$C$144:$C$212,PP1A!$C$215:$C$283)',
        'values': '=(PP1A!$N$144:$N$212,PP1A!$N$215:$N$283)',
        'name': '=PP1A!$E$212',
    })
    chartc2c2asyncreq.add_series({
        'values': '=(PP1B!$N$144:$N$212,PP1B!$N$215:$N$283)',
        'name': '=PP1B!$E$212',
    })
    chartc2c2asyncreq.add_series({
        'values': '=(PP1C!$N$144:$N$212,PP1C!$N$215:$N$283)',
        'name': '=PP1C!$E$212',
    })
    chartc2c2asyncreq.add_series({
        'values': '=(PP1D!$N$144:$N$212,PP1D!$N$215:$N$283)',
        'name': '=PP1D!$E$212',
    })
    chartc2c2asyncreq.add_series({
        'values': '=(PP1E!$N$144:$N$212,PP1E!$N$215:$N$283)',
        'name': '=PP1E!$E$212',
    })
    chartc2c2asyncreq.add_series({
        'values': '=(PP1F!$N$144:$N$212,PP1F!$N$215:$N$283)',
        'name': '=PP1F!$E$212',
    })
    chartc2c2asyncreq.add_series({
        'values': '=(PP1G!$N$144:$N$212,PP1G!$N$215:$N$283)',
        'name': '=PP1G!$E$212',
    })
    chartc2c2asyncreq.add_series({
        'values': '=(PP1H!$N$144:$N$212,PP1H!$N$215:$N$283)',
        'name': '=PP1H!$E$212',
    })
    chartc2c2asyncreq.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartc2c2asyncreq.set_legend({'position': 'bottom'})
    chartc2c2asyncreq.set_title({'name': 'SPC2C2-ASYNC-REQ NUMBER'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B82', chartc2c2asyncreq)

    #SPC2C2 ASYNC RESP
    chartc2c2asyncresp = workbook.add_chart({'type': 'line'})
    chartc2c2asyncresp.height=350
    chartc2c2asyncresp.width=500
    chartc2c2asyncresp.add_series({
        'categories': '=(PP1A!$C$144:$C$212,PP1A!$C$215:$C$283)',
        'values': '=(PP1A!$O$144:$O$212,PP1A!$O$215:$O$283)',
        'name': '=PP1A!$E$212',
    })
    chartc2c2asyncresp.add_series({
        'values': '=(PP1B!$O$144:$O$212,PP1B!$O$215:$O$283)',
        'name': '=PP1B!$E$212',
    })
    chartc2c2asyncresp.add_series({
        'values': '=(PP1C!$O$144:$O$212,PP1C!$O$215:$O$283)',
        'name': '=PP1C!$E$212',
    })
    chartc2c2asyncresp.add_series({
        'values': '=(PP1D!$O$144:$O$212,PP1D!$O$215:$O$283)',
        'name': '=PP1D!$E$3',
    })
    chartc2c2asyncresp.add_series({
        'values': '=(PP1E!$O$144:$O$212,PP1E!$O$215:$O$283)',
        'name': '=PP1E!$E$212',
    })
    chartc2c2asyncresp.add_series({
        'values': '=(PP1F!$O$144:$O$212,PP1F!$O$215:$O$283)',
        'name': '=PP1F!$E$212',
    })
    chartc2c2asyncresp.add_series({
        'values': '=(PP1G!$O$144:$O$212,PP1G!$O$215:$O$283)',
        'name': '=PP1G!$E$212',
    })
    chartc2c2asyncresp.add_series({
        'values': '=(PP1H!$O$144:$O$212,PP1H!$O$215:$O$283)',
        'name': '=PP1H!$E$212',
    })
    chartc2c2asyncresp.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartc2c2asyncresp.set_legend({'position': 'bottom'})
    chartc2c2asyncresp.set_title({'name': 'SPC2C2-ASYNC RESP TIME'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K82', chartc2c2asyncresp)
    writer.save()




