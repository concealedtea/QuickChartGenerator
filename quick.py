import pyodbc
import sys, csv, os, json, math
import pandas as pd
from pandas import DataFrame
from pandas.tools.plotting import table
import itertools as IT
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
from scipy.stats import linregress
from collections import defaultdict

def exec_query(query_text,region,platform,driver= "{ODBC Driver 13 for SQL Server};",server= "REMOVED;",
               database= "REMOVED;",UID= "REMOVED;",PWD= "REMOVED;",return_value= True):
    conn= pyodbc.connect(
        r'DRIVER='+driver+
        r'SERVER='+server+
        r'DATABASE='+database+
        r'UID='+UID+
        r'PWD='+PWD
        )
    return conn

def main():
    regions = ['US','GB','CA','IN','Other']
    platforms = ['Desktop', 'Mobile']
    for region in regions:
        for platform in platforms:
            query = ("""SELECT
                              [Date] as date,
                              [Advertiser] as advertiser,
                              SUM([Receives]) as receives,
                              ISNULL(1000 * SUM(revenue) / NULLIF(CAST(SUM(receives) AS FLOAT),0),0) AS RPM
                        FROM [Reports].[dbo].[DailyRPC]
                        where 1=1
                        and date BETWEEN CAST(DATEADD(DAY,-7,GETDATE()) AS DATE) AND CAST(DATEADD(DAY,-1,GETDATE()) AS DATE)
                        and country = '""" + region + """'
                        and platform = '""" + platform + """'
                        group by
                        	date, Advertiser
                        HAVING
                        	sum([Receives]) > 500
                        	AND	ISNULL(1000 * SUM(revenue) / NULLIF(CAST(SUM(receives) AS FLOAT),0),0) > 0
                        order by date desc""")
            msql_db = exec_query(query, region, platform)
            df = pd.read_sql(query, msql_db)
            # Bar Graphs for Advertisers
            dfBars = df.drop('RPM', 1)
            dfBars_pivot = dfBars.pivot(index = 'date', columns = 'advertiser', values = 'receives')
            fig, ax = plt.subplots()
            dfBars_pivot.plot.bar(fontsize = 7, stacked=True, title = region + ' ' + platform + ' Receives', rot = 0, ax = ax, table = True)
            ax.grid(linestyle = 'dotted')
            ax.set_axisbelow(True)
            plt.legend(loc='upper center', frameon=False, ncol = len(dfBars['advertiser']), prop={'size': 6})
            plt.tick_params(axis='x',          # changes apply to the x-axis
                            which='both',      # both major and minor ticks are affected
                            bottom=False,      # ticks along the bottom edge are off
                            top=False,         # ticks along the top edge are off
                            labelbottom=False)
            ax.xaxis.label.set_visible(False)
            plt.tight_layout()
            fig = plt.gcf()
            fig.savefig(region+platform+'Receives.png', bbox_inches='tight')

            # Line Graphs for Advertisers
            dfLines = df.drop('receives', 1)
            dfLines_pivot = dfLines.pivot(index = 'date', columns = 'advertiser', values = 'RPM')
            fig, ax = plt.subplots()
            dfLines_pivot.plot(fontsize = 8, title = region + ' ' + platform + ' RPM', rot = 0, ax = ax, figsize=(20,5), table = True)
            fmt = '${x:,.00f}'
            tick = mtick.StrMethodFormatter(fmt)
            ax.grid(linestyle = 'dotted')
            ax.set_axisbelow(True)
            plt.yticks(np.arange(0, math.ceil(max(dfLines['RPM']) + 1), 1.0))
            plt.legend(loc='upper center', frameon=False, ncol = len(dfBars['advertiser']), prop={'size': 8})
            plt.tick_params(axis='x',          # changes apply to the x-axis
                            which='both',      # both major and minor ticks are affected
                            bottom=False,      # ticks along the bottom edge are off
                            top=False,         # ticks along the top edge are off
                            labelbottom=False)
            ax.xaxis.label.set_visible(False)
            ax.tick_params(labelbottom='off')
            plt.tight_layout()
            fig = plt.gcf()
            fig.savefig(region+platform+'RPM.png', bbox_inches='tight')

if __name__ == '__main__':
    main()
    input("Process Finished, please close the window")
    os._exit(1)
