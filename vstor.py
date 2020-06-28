"""Performance comparison with python, 1.0
Yangmeichen 2020/6/22"""
import pandas as pd
from pandas import Series, DataFrame
from pandas import to_datetime
from datetime import datetime
import xlsxwriter

def vstor():
    filenamebefore = input('Enter vstorbefore: ')
    filenameafter = input('Enter vstorafter: ')
    # read data
    df1= pd.read_excel(filenamebefore,encoding='utf8')
    df2 = pd.read_excel(filenameafter, encoding='utf8')
    df1.reset_index()
    df2.reset_index()
        # 设置TIME为索引
    df1 = df1.set_index('TIME')
    df2 = df2.set_index('TIME')
        # 提取联机数据
    df1 = pd.DataFrame(df1['00.00.00': '12.00.00'])
    df2 = pd.DataFrame(df2['00.00.00': '12.00.00'])
    dfpp1a1=df1[(df1.SYSNAME == 'PP1A')]
    dfpp1b1=df1[(df1.SYSNAME == 'PP1B')]
    dfpp1c1=df1[(df1.SYSNAME == 'PP1C')]
    dfpp1d1=df1[(df1.SYSNAME == 'PP1D')]
    dfpp1e1=df1[(df1.SYSNAME == 'PP1E')]
    dfpp1f1=df1[(df1.SYSNAME == 'PP1F')]
    dfpp1g1=df1[(df1.SYSNAME == 'PP1G')]
    dfpp1h1=df1[(df1.SYSNAME == 'PP1H')]

    dfpp1a2=df2[(df2.SYSNAME == 'PP1A')]
    dfpp1b2=df2[(df2.SYSNAME == 'PP1B')]
    dfpp1c2=df2[(df2.SYSNAME == 'PP1C')]
    dfpp1d2=df2[(df2.SYSNAME == 'PP1D')]
    dfpp1e2=df2[(df2.SYSNAME == 'PP1E')]
    dfpp1f2=df2[(df2.SYSNAME == 'PP1F')]
    dfpp1g2=df2[(df2.SYSNAME == 'PP1G')]
    dfpp1h2=df2[(df2.SYSNAME == 'PP1H')]



    writer = pd.ExcelWriter('VSTOR.xlsx', engine='xlsxwriter', datetime_format='YYYY/MM/DD')
    dfpp1a1.to_excel(writer, sheet_name='PP1A')
    dfpp1b1.to_excel(writer, sheet_name='PP1B')
    dfpp1c1.to_excel(writer, sheet_name='PP1C')
    dfpp1d1.to_excel(writer, sheet_name='PP1D')
    dfpp1e1.to_excel(writer, sheet_name='PP1E')
    dfpp1f1.to_excel(writer, sheet_name='PP1F')
    dfpp1g1.to_excel(writer, sheet_name='PP1G')
    dfpp1h1.to_excel(writer, sheet_name='PP1H')

    dfpp1a2.to_excel(writer, sheet_name='PP1A',startrow = 52)
    dfpp1b2.to_excel(writer, sheet_name='PP1B',startrow = 52)
    dfpp1c2.to_excel(writer, sheet_name='PP1C',startrow = 52)
    dfpp1d2.to_excel(writer, sheet_name='PP1D',startrow = 52)
    dfpp1e2.to_excel(writer, sheet_name='PP1E',startrow = 52)
    dfpp1f2.to_excel(writer, sheet_name='PP1F',startrow = 52)
    dfpp1g2.to_excel(writer, sheet_name='PP1G',startrow = 52)
    dfpp1h2.to_excel(writer, sheet_name='PP1H',startrow = 52)

    workbook  = writer.book
    worksheet = writer.sheets['PP1A']
    #CSA
    chartcsa = workbook.add_chart({'type': 'line'})
    chartcsa.height=350
    chartcsa.width=500
    chartcsa.add_series({
        'categories': '=(PP1A!$C$2:$C$51,PP1A!$C$54:$C$103)',
        'values': '=(PP1A!$S$2:$S$51,PP1A!$S$54:$S$103)',
        'name': '=PP1A!$D$3',
    })
    chartcsa.add_series({
        'values': '=(PP1B!$S$2:$S$51,PP1B!$S$54:$S$103)',
        'name': '=PP1B!$D$3',
    })
    chartcsa.add_series({
        'values': '=(PP1C!$S$2:$S$51,PP1C!$S$54:$S$103)',
        'name': '=PP1C!$D$3',
    })
    chartcsa.add_series({
        'values': '=(PP1D!$S$2:$S$51,PP1D!$S$54:$S$103)',
        'name': '=PP1D!$D$3',
    })
    chartcsa.add_series({
        'values': '=(PP1E!$S$2:$S$51,PP1E!$S$54:$S$103)',
        'name': '=PP1E!$D$3',
    })
    chartcsa.add_series({
        'values': '=(PP1F!$S$2:$S$51,PP1F!$S$54:$S$103)',
        'name': '=PP1F!$D$3',
    })
    chartcsa.add_series({
        'values': '=(PP1G!$S$2:$S$51,PP1G!$S$54:$S$103)',
        'name': '=PP1G!$D$3',
    })
    chartcsa.add_series({
        'values': '=(PP1H!$S$2:$S$51,PP1H!$S$54:$S$103)',
        'name': '=PP1H!$D$3',
    })

    chartcsa.set_y_axis({
        'num_format': '0.00%',
        'minor_unit': 0.01, 'major_unit': 0.05,
        'min': 0, 'max': 0.3
    })
    chartcsa.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartcsa.set_legend({'position': 'bottom'})
    chartcsa.set_title({'name': 'VSTOR CSA%'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B2', chartcsa)

    #ECSA
    chartecsa = workbook.add_chart({'type': 'line'})
    chartecsa.height=350
    chartecsa.width=500
    chartecsa.add_series({
        'categories': '=(PP1A!$C$2:$C$51,PP1A!$C$54:$C$103)',
        'values': '=(PP1A!$T$2:$T$51,PP1A!$T$54:$T$103)',
        'name': '=PP1A!$D$3',
    })
    chartecsa.add_series({
        'values': '=(PP1B!$T$2:$T$51,PP1B!$T$54:$T$103)',
        'name': '=PP1B!$D$3',
    })
    chartecsa.add_series({
        'values': '=(PP1C!$T$2:$T$51,PP1C!$T$54:$T$103)',
        'name': '=PP1C!$D$3',
    })
    chartecsa.add_series({
        'values': '=(PP1D!$T$2:$T$51,PP1D!$T$54:$T$103)',
        'name': '=PP1D!$D$3',
    })
    chartecsa.add_series({
        'values': '=(PP1E!$T$2:$T$51,PP1E!$T$54:$T$103)',
        'name': '=PP1E!$D$3',
    })
    chartecsa.add_series({
        'values': '=(PP1F!$T$2:$T$51,PP1F!$T$54:$T$103)',
        'name': '=PP1F!$D$3',
    })
    chartecsa.add_series({
        'values': '=(PP1G!$T$2:$T$51,PP1G!$T$54:$T$103)',
        'name': '=PP1G!$D$3',
    })
    chartecsa.add_series({
        'values': '=(PP1H!$T$2:$T$51,PP1H!$T$54:$T$103)',
        'name': '=PP1H!$D$3',
    })

    chartecsa.set_y_axis({
        'num_format': '0.00%',
        'minor_unit': 0.02, 'major_unit': 0.1,
        'min': 0, 'max': 0.8
    })
    chartecsa.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartecsa.set_legend({'position': 'bottom'})
    chartecsa.set_title({'name': 'VSTOR ECSA%'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K2', chartecsa)

    #SQA
    chartsqa = workbook.add_chart({'type': 'line'})
    chartsqa.height=350
    chartsqa.width=500
    chartsqa.add_series({
        'categories': '=(PP1A!$C$2:$C$51,PP1A!$C$54:$C$103)',
        'values': '=(PP1A!$U$2:$U$51,PP1A!$U$54:$U$103)',
        'name': '=PP1A!$D$3',
    })
    chartsqa.add_series({
        'values': '=(PP1B!$U$2:$U$51,PP1B!$U$54:$U$103)',
        'name': '=PP1B!$D$3',
    })
    chartsqa.add_series({
        'values': '=(PP1C!$U$2:$U$51,PP1C!$U$54:$U$103)',
        'name': '=PP1C!$D$3',
    })
    chartsqa.add_series({
        'values': '=(PP1D!$U$2:$U$51,PP1D!$U$54:$U$103)',
        'name': '=PP1D!$D$3',
    })
    chartsqa.add_series({
        'values': '=(PP1E!$U$2:$U$51,PP1E!$U$54:$U$103)',
        'name': '=PP1E!$D$3',
    })
    chartsqa.add_series({
        'values': '=(PP1F!$U$2:$U$51,PP1F!$U$54:$U$103)',
        'name': '=PP1F!$D$3',
    })
    chartsqa.add_series({
        'values': '=(PP1G!$U$2:$U$51,PP1G!$U$54:$U$103)',
        'name': '=PP1G!$D$3',
    })
    chartsqa.add_series({
        'values': '=(PP1H!$U$2:$U$51,PP1H!$U$54:$U$103)',
        'name': '=PP1H!$D$3',
    })

    chartsqa.set_y_axis({
        'num_format': '0.00%',
        'minor_unit': 0.02, 'major_unit': 0.1,
        'min': 0, 'max': 0.8
    })
    chartsqa.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartsqa.set_legend({'position': 'bottom'})
    chartsqa.set_title({'name': 'VSTOR SQA%'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B22', chartsqa)

    #ESQA
    chartesqa = workbook.add_chart({'type': 'line'})
    chartesqa.height=350
    chartesqa.width=500
    chartesqa.add_series({
        'categories': '=(PP1A!$C$2:$C$51,PP1A!$C$54:$C$103)',
        'values': '=(PP1A!$V$2:$V$51,PP1A!$V$54:$V$103)',
        'name': '=PP1A!$D$3',
    })
    chartesqa.add_series({
        'values': '=(PP1B!$V$2:$V$51,PP1B!$V$54:$V$103)',
        'name': '=PP1B!$D$3',
    })
    chartesqa.add_series({
        'values': '=(PP1C!$V$2:$V$51,PP1C!$V$54:$V$103)',
        'name': '=PP1C!$D$3',
    })
    chartesqa.add_series({
        'values': '=(PP1D!$V$2:$V$51,PP1D!$V$54:$V$103)',
        'name': '=PP1D!$D$3',
    })
    chartesqa.add_series({
        'values': '=(PP1E!$V$2:$V$51,PP1E!$V$54:$V$103)',
        'name': '=PP1E!$D$3',
    })
    chartesqa.add_series({
        'values': '=(PP1F!$V$2:$V$51,PP1F!$V$54:$V$103)',
        'name': '=PP1F!$D$3',
    })
    chartesqa.add_series({
        'values': '=(PP1G!$V$2:$V$51,PP1G!$V$54:$V$103)',
        'name': '=PP1G!$D$3',
    })
    chartesqa.add_series({
        'values': '=(PP1H!$V$2:$V$51,PP1H!$V$54:$V$103)',
        'name': '=PP1H!$D$3',
    })

    chartesqa.set_y_axis({
        'num_format': '0.00%',
        'minor_unit': 0.02, 'major_unit': 0.1,
        'min': 0, 'max': 1
    })
    chartesqa.set_plotarea({
        'layout': {
            'x': 0.15,
            'y': 0.13,
            'width':  0.75,
            'height': 0.35,
        }
    })
    chartesqa.set_legend({'position': 'bottom'})
    chartesqa.set_title({'name': 'VSTOR ESQA%'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K22', chartesqa)
    writer.save()







