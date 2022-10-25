global System, DataTable, AMO, ADOMD

import pandas as pd
import ssas_api as ssas
import seaborn as sns
import sys
import os
sys.path.insert(0,os.path.abspath(' ./' ))
import clr
clr.AddReference('System.Data' )
#r = clr.AddReference(r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.AnalysisServices.Tabular\v4.0_15.0.0.0__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.dll")
#r = clr.AddReference(r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.AnalysisServices.AdomdClient\v4.0_15.0.0.0__89845dcd8080cc91\Microsoft.AnalysisServices.AdomdClient.dll")
from System import Data
from System.Data import DataTable
import Microsoft.AnalysisServices.Tabular as TOM
import Microsoft.AnalysisServices.AdomdClient as ADOMD



server = 'powerbi://api.powerbi.com/v1.0/myorg/ModelosMayores1GB'

username = 'j.guzman@bdigitalsolutions.com'

password = 'Bogota2022'

print('ok1\n')
conn1 = ssas.set_conn_string(
    server=server,
    db_name='',
    username=username,
    password=password
    )

try:
    print('ok2\n')
    TOMServer = TOM.Server()
    print('o3k\n')
    TOMServer.Connect(conn1)
    print("Connection to Workspace Successful !")

except:
    print("Connection to Workspace Failed")

datasets = pd.DataFrame(columns=['Dataset_Name', 'Compatibility', 'ID', 'Size_MB','Created_Date','Last_Update' ])

for item in TOMServer.Databases:


    datasets = datasets.append({'Dataset_Name' :item.Name,
                     'Compatibility':item.CompatibilityLevel,
                     'Created_Date' :item.CreatedTimestamp,
                     'ID'           :item.ID,
                     'Last_Update'  :item.LastUpdate,
                     'Size_MB'      :(item.EstimatedSize*1e-06)    },
                     ignore_index=True)

for i in datasets:
    print(datasets[i])

ds = TOMServer.Databases['bc8741ce-3063-4647-9b1e-34c5a855fc8e']

for table in ds.Model.Tables:
    print(table.Name)

conn2 = (ssas.set_conn_string(
    server=server,
    db_name='SatisfactionReport',
    username = username,
    password = password
    #Catalog=**YourModelName**
 ))

dax_string = '''
    EVALUATE
    data_dictionary
    '''

df_dataset2 = (ssas.get_DAX(
                   connection_string=conn2,
                   dax_string=dax_string)

              )

print(df_dataset2.head())
