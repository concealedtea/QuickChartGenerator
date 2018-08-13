import pyodbc
import sys, csv, os, json, math
import pandas as pd
from io import StringIO
from pandas import DataFrame
from pandas.tools.plotting import table
import itertools as IT
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
from scipy.stats import linregress
from collections import defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import smtplib
import email

def exec_query(query_text,region,platform,driver= "{ODBC Driver 13 for SQL Server};",server= "REMOVED",
               database= "REMOVED",UID= "REMOVED",PWD= "REMOVED",return_value= True):
    conn= pyodbc.connect(
        r'DRIVER='+driver+
        r'SERVER='+server+
        r'DATABASE='+database+
        r'UID='+UID+
        r'PWD='+PWD
        )
    return conn

def emailer():
    fromaddr = "REMOVED"
    toaddr = "REMOVED"
    regions = ['US','GB','CA','IN']
    platforms = ['Desktop', 'Mobile']
    types = ['RPM', 'Receives']
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = 'Daily RPM/Revenue Charts'
    msgRoot['From'] = fromaddr
    msgRoot['To'] = toaddr
    msgRoot.preamble = 'This is a multipart message in MIME format.'
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)
    msgText = MIMEText('This is the alternative plain text message')
    msgAlternative.attach(msgText)
    count = 0
    for region in regions:
        for platform in platforms:
            for type in types:
                msgText = MIMEText('<b>Daily Charts</b><br><img src="cid:image0">'
                                   '<img src="cid:image1"><img src="cid:image2">'
                                   '<img src="cid:image3"><img src="cid:image4">'
                                   '<img src="cid:image5"><img src="cid:image6">'
                                   '<img src="cid:image7"><img src="cid:image8">'
                                   '<img src="cid:image9"><img src="cid:image10">'
                                   '<img src="cid:image11"><img src="cid:image12">'
                                   '<img src="cid:image13"><img src="cid:image14">'
                                   '<img src="cid:image15">'
                                   '<br>', 'html')
                msgAlternative.attach(msgText)
                filename = region+platform+type+'.png'
                fp = open(filename, 'rb')
                msgImage = MIMEImage(fp.read())
                fp.close()
                msgImage.add_header('Content-ID', '<image'+str(count)+'>')
                msgRoot.attach(msgImage)
                count += 1
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("Charts.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="Charts.xlsx"')
    msgRoot.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, 'REMOVED')
    server.sendmail(fromaddr, toaddr, msgRoot.as_string())
    server.quit()


def main():
    regions = ['US','GB','CA','IN']
    platforms = ['Desktop', 'Mobile']
    writer = pd.ExcelWriter('Charts.xlsx')
    for region in regions:
        for platform in platforms:
            query = ("""SELECT * FROM
                                (SELECT
                                        [Date] as date,
                                		'Blended' as advertiser,
                                        SUM([Receives]) as receives,
                                        ROUND(ISNULL(1000 * SUM(revenue) / NULLIF(CAST(SUM(receives) AS FLOAT),0),0),2) AS RPM
                                FROM [Reports].[dbo].[DailyRPC]
                                where 1=1
                                and date BETWEEN CAST(DATEADD(DAY,-9,GETDATE()) AS DATE) AND CAST(DATEADD(DAY,-1,GETDATE()) AS DATE)
                                and country = '"""+region+"""'
                                and platform = '"""+platform+"""'
                                group by
                                    date
                                HAVING
                                    ISNULL(1000 * SUM(revenue) / NULLIF(CAST(SUM(receives) AS FLOAT),0),0) > 0) as A
                                UNION ALL (SELECT
                                        [Date] as date,
                                        [Advertiser] as advertiser,
                                        SUM([Receives]) as receives,
                                        ROUND(ISNULL(1000 * SUM(revenue) / NULLIF(CAST(SUM(receives) AS FLOAT),0),0),2) AS RPM
                                FROM [Reports].[dbo].[DailyRPC]
                                where 1=1
                                and date BETWEEN CAST(DATEADD(DAY,-9,GETDATE()) AS DATE) AND CAST(DATEADD(DAY,-1,GETDATE()) AS DATE)
                                and country = '"""+region+"""'
                                and platform = '"""+platform+"""'
                                group by
                                    date, advertiser
                                HAVING
                                    ISNULL(1000 * SUM(revenue) / NULLIF(CAST(SUM(receives) AS FLOAT),0),0) > 0)""")
            query2 =  ("""SELECT
                                            usermetrics.[Date],
											usermetrics.platform,
											case when usermetrics.country = 'us' then 'US'
													when usermetrics.country = 'gb' then 'GB'
													when usermetrics.country = 'ca' then 'CA'
													when usermetrics.country = 'in' then 'IN'
													else 'Other'
													end,
                                            usermetrics.[Advertiser],
											SUM(usermetrics.clicks * DailyRPC.RPC) as revenue,
                                            SUM(usermetrics.[Receives]) as receives,
                                            ROUND(ISNULL(1000.0 * SUM(usermetrics.clicks * DailyRPC.RPC) / NULLIF(CAST(SUM(usermetrics.receives) AS FLOAT),0),0),2) AS RPM,
                                            CASE
                                                WHEN age < 8 THEN '0-7'
                                                WHEN age > 7 AND age < 22 THEN  '8-21'
                                                WHEN age > 22 THEN '22+'
                                            ELSE '22+'
                                            END as age_bucket
                                    FROM [Reports].[dbo].UserMetrics AS usermetrics
                                    LEFT OUTER JOIN
                                        Reports.dbo.DailyRPC AS DailyRPC ON DailyRPC.[date] = usermetrics.[date]
                                            AND DailyRPC.Country = usermetrics.country
                                            AND DailyRPC.advertiser = usermetrics.advertiser
                                            AND DailyRPC.[platform] = usermetrics.[platform]
                                    where 1=1
                                    and usermetrics.date BETWEEN CAST(DATEADD(DAY,-7,GETDATE()) AS DATE) AND CAST(DATEADD(DAY,-1,GETDATE()) AS DATE)
                                    group by
                                        usermetrics.date,usermetrics.platform, case when usermetrics.country = 'us' then 'US'
													when usermetrics.country = 'gb' then 'GB'
													when usermetrics.country = 'ca' then 'CA'
													when usermetrics.country = 'in' then 'IN'
													else 'Other'
													end, usermetrics.advertiser, CASE
                                                WHEN age < 8 THEN '0-7'
                                                WHEN age > 7 AND age < 22 THEN  '8-21'
                                                WHEN age > 22 THEN '22+'
                                            ELSE '22+'
                                            END
                                    HAVING
                                        ROUND(ISNULL(1000.0 * SUM(usermetrics.clicks * DailyRPC.RPC) / NULLIF(CAST(SUM(usermetrics.receives) AS FLOAT),0),0),2) > 0""")
            msql_db = exec_query(query, region, platform)
            df = pd.read_sql(query, msql_db)
            # Bar Graphs for Advertisers
            dfBars = df.drop('RPM', 1)
            dfBarsWBlend = dfBars.replace('Blended', 'Total')
            dfBars = dfBars[df.advertiser != 'Blended']
            dfBars_pivot = dfBars.pivot(index = 'date', columns = 'advertiser', values = 'receives')
            fig, ax= plt.subplots()
            dfBarsWBlend.loc[dfBarsWBlend['advertiser'] == 'Total'].plot.line(ax = ax)
            dfBars_pivot.plot.bar(fontsize = 7, stacked=True, title = region + ' ' + platform + ' Receives', rot = 0, ax = ax)
            plt.subplots_adjust(bottom=0.25, top = 0.95)
            fmt = '{x:,.00f}'
            tick = mtick.StrMethodFormatter(fmt)
            ax.yaxis.set_major_formatter(tick)
            ax.grid(linestyle = 'dotted')
            ax.set_axisbelow(True)
            ax.xaxis.set_major_formatter(plt.NullFormatter())
            ax.xaxis.label.set_visible(False)
            plt.legend(loc='upper center', frameon=False, ncol = len(dfBars['advertiser']), prop={'size': 6})
            # Bar Table
            temp = []
            for item in dfBarsWBlend.receives:
                temp.append('{:,}'.format(item))
            dfBarsWBlend.receives = temp
            tableBars = dfBarsWBlend.pivot(index = 'advertiser', columns = 'date', values = 'receives')
            dfTemp = tableBars.loc[tableBars.index == 'Total']
            tableBars = tableBars.loc[tableBars.index != 'Total']
            tableNew = pd.concat([tableBars,dfTemp], axis = 0)
            tableNew = tableNew.fillna('-')
            plt.table(cellText = tableNew.values, rowLabels = tableNew.index, colLabels = tableNew.columns, loc = 'bottom')
            plt.tight_layout()
            fig = plt.gcf()
            fig.savefig(region+platform+'Receives.png', bbox_inches='tight')

            # Line Graphs for Advertisers
            dfLines = df.drop('receives', 1)
            dfLines_pivot = dfLines.pivot(index = 'date', columns = 'advertiser', values = 'RPM')
            fig, ax = plt.subplots()
            dfLines_pivot.fillna('-')
            pos = 0
            for i in range(0,len(dfLines_pivot.columns)):
                if (dfLines_pivot.columns[i] == 'Blended'):
                    pos = i
            styles = ['-'] * len(dfLines_pivot.columns)
            styles[pos] = ':'
            dfLines_pivot.plot(fontsize = 8, title = region + ' ' + platform + ' RPM', rot = 0, ax = ax, linewidth = 2, style = styles)
            plt.subplots_adjust(bottom=0.25, top = 0.95)
            fmt = '${x:}'
            tick = mtick.StrMethodFormatter(fmt)
            ax.yaxis.set_major_formatter(tick)
            fmt = '${x:,.00f}'
            ax.grid(linestyle = 'dotted')
            ax.set_axisbelow(True)
            ax.xaxis.set_major_formatter(plt.NullFormatter())
            ax.xaxis.label.set_visible(False)
            plt.yticks(np.arange((math.floor(min(dfLines['RPM']))), (math.ceil(max(dfLines['RPM'])) * 1.1), 0.5))
            plt.legend(loc='upper center', frameon=False, ncol = len(dfBars['advertiser']), prop={'size': 6})
            # Line Table
            temp = []
            for item in dfLines.RPM:
                temp.append('${:.2f}'.format(item))
            dfLines.RPM = temp
            tableLines = dfLines.pivot(index = 'advertiser', columns = 'date', values = 'RPM')
            dfTemp = tableLines.loc[tableLines.index == 'Blended']
            tableLines = tableLines.loc[tableLines.index != 'Blended']
            tableNew = pd.concat([tableLines,dfTemp], axis = 0)
            tableNew = tableNew.fillna('-')
            plt.table(cellText = tableNew.values, rowLabels = tableNew.index, colLabels = tableNew.columns, loc = 'bottom')
            plt.tight_layout()
            fig = plt.gcf()
            fig.savefig(region+platform+'RPM.png', bbox_inches='tight')
    df2 = pd.read_sql(query2, msql_db)
    df2.to_excel(writer, region+platform)
    writer.save()
    emailer()

if __name__ == '__main__':
    main()
    print('Finished, program will now exit')
    os._exit(1)
