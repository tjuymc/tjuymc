"""Performance comparison with python, 1.0
   Yangmeichen 2020/6/22"""
import pandas as pd
from pandas import Series, DataFrame
from pandas import to_datetime
from datetime import datetime
import xlsxwriter

def csusage():
    filenamebefore = input('Enter pagingbefore: ')
    filenameafter = input('Enter pagingafter: ')


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



    writer = pd.ExcelWriter('CSUSAGE.xlsx', engine='xlsxwriter', datetime_format='YYYY/MM/DD')
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
    #PP1A
    chartpp1a = workbook.add_chart({'type': 'line'})
    chartpp1a.height=350
    chartpp1a.width=500
    chartpp1a.add_series({
        'categories': '=PP1A!$A$2:$A$51',
        'values': '=PP1A!$Q$2:$Q$51',
        'name': '=PP1A!$B$3',
    })
    chartpp1a.add_series({
        'values': '=PP1A!$Q$54:$Q$103',
        'name': '=PP1A!$B$54'
    })
    chartpp1a.set_y_axis({'num_format': '0.00%'})
    chartpp1a.set_legend({'position': 'bottom'})
    chartpp1a.set_title({'name': 'CS USAGE%-PP1A'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B2', chartpp1a)

    #PP1B
    chartpp1b = workbook.add_chart({'type': 'line'})
    chartpp1b.height=350
    chartpp1b.width=500
    chartpp1b.add_series({
        'categories': '=PP1B!$A$2:$A$51',
        'values': '=PP1B!$Q$2:$Q$51',
        'name': '=PP1B!$B$3',
    })
    chartpp1b.add_series({
        'values': '=PP1B!$Q$54:$Q$103',
         'name': '=PP1B!$B$54'
    })
    chartpp1b.set_y_axis({'num_format': '0.00%'})
    chartpp1b.set_legend({'position': 'bottom'})
    chartpp1b.set_title({'name': 'CS USAGE%-PP1B'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K2', chartpp1b)

    #PP1C
    chartpp1c = workbook.add_chart({'type': 'line'})
    chartpp1c.height=350
    chartpp1c.width=500
    chartpp1c.add_series({
        'categories': '=PP1C!$A$2:$A$51',
        'values': '=PP1C!$Q$2:$Q$51',
        'name': '=PP1C!$B$3',
    })
    chartpp1c.add_series({
        'values': '=PP1C!$Q$54:$Q$103',
        'name': '=PP1C!$B$54'
    })
    chartpp1c.set_y_axis({'num_format': '0.00%'})
    chartpp1c.set_legend({'position': 'bottom'})
    chartpp1c.set_title({'name': 'CS USAGE%-PP1C'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B22', chartpp1c)

    #PP1D
    chartpp1d = workbook.add_chart({'type': 'line'})
    chartpp1d.height=350
    chartpp1d.width=500
    chartpp1d.add_series({
        'categories': '=PP1D!$A$2:$A$51',
        'values': '=PP1D!$Q$2:$Q$51',
        'name': '=PP1D!$B$3',
    })
    chartpp1d.add_series({
        'values': '=PP1D!$Q$54:$Q$103',
        'name': '=PP1D!$B$54',
    })
    chartpp1d.set_y_axis({'num_format': '0.00%'})
    chartpp1d.set_legend({'position': 'bottom'})
    chartpp1d.set_title({'name': 'CS USAGE%-PP1D'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K22', chartpp1d)

    #PP1E
    chartpp1e = workbook.add_chart({'type': 'line'})
    chartpp1e.height=350
    chartpp1e.width=500
    chartpp1e.add_series({
        'categories': '=PP1E!$A$2:$A$51',
        'values': '=PP1E!$Q$2:$Q$51',
        'name': '=PP1E!$B$3',
    })
    chartpp1e.add_series({
        'values': '=PP1E!$Q$54:$Q$103',
        'name': '=PP1E!$B$54',
    })
    chartpp1e.set_y_axis({'num_format': '0.00%'})
    chartpp1e.set_legend({'position': 'bottom'})
    chartpp1e.set_title({'name': 'CS USAGE%-PP1E'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B42', chartpp1e)

    #PP1F
    chartpp1f = workbook.add_chart({'type': 'line'})
    chartpp1f.height=350
    chartpp1f.width=500
    chartpp1f.add_series({
        'categories': '=PP1F!$A$2:$A$51',
        'values': '=PP1F!$Q$2:$Q$51',
        'name': '=PP1F!$B$3',
    })
    chartpp1f.add_series({
        'values': '=PP1F!$Q$54:$Q$103',
        'name': '=PP1F!$B$54',
    })
    chartpp1f.set_y_axis({'num_format': '0.00%'})
    chartpp1f.set_legend({'position': 'bottom'})
    chartpp1f.set_title({'name': 'CS USAGE%-PP1F'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K42', chartpp1f)

    #PP1G
    chartpp1g = workbook.add_chart({'type': 'line'})
    chartpp1g.height=350
    chartpp1g.width=500
    chartpp1g.add_series({
        'categories': '=PP1G!$A$2:$A$51',
        'values': '=PP1G!$Q$2:$Q$51',
        'name': '=PP1G!$B$3',
    })
    chartpp1g.add_series({
        'values': '=PP1G!$Q$54:$Q$103',
        'name': '=PP1G!$B$54',
    })
    chartpp1g.set_y_axis({'num_format': '0.00%'})
    chartpp1g.set_legend({'position': 'bottom'})
    chartpp1g.set_title({'name': 'CS USAGE%-PP1G'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('B62', chartpp1g)

    #PP1H
    chartpp1h = workbook.add_chart({'type': 'line'})
    chartpp1h.height=350
    chartpp1h.width=500
    chartpp1h.add_series({
        'categories': '=PP1H!$A$2:$A$51',
        'values': '=PP1H!$Q$2:$Q$51',
        'name': '=PP1H!$B$3',
    })
    chartpp1h.add_series({
        'values': '=PP1H!$Q$54:$Q$103',
        'name': '=PP1H!$B$54',
    })
    chartpp1h.set_y_axis({'num_format': '0.00%'})
    chartpp1h.set_legend({'position': 'bottom'})
    chartpp1h.set_title({'name': 'CS USAGE%-PP1H'})
    # Insert the chart into the worksheet.
    worksheet.insert_chart('K62', chartpp1h)
    writer.save()







