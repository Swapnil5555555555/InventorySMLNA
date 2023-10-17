#Import Modules

import oracledb
import csv
from azure.storage.blob import *
import dotenv
from datetime import datetime
import win32com.client
import time

#Initializing variables and clients
oracle_client_connection = oracledb.init_oracle_client()
storage_account_key="aco259GEouCjzFWyQBqvRn6J3syKa9dAgMsopOXLyh+ygyf0WtRyGCi+3H6WnYucTgzcDuqo55N6uliv4ebMLw=="
storage_account_name="smlnorthamericaanalytics"
container_name="smlnainventorygoldendataset"
connection_string="DefaultEndpointsProtocol=https;AccountName=smlnorthamericaanalytics;AccountKey=aco259GEouCjzFWyQBqvRn6J3syKa9dAgMsopOXLyh+ygyf0WtRyGCi+3H6WnYucTgzcDuqo55N6uliv4ebMLw==;EndpointSuffix=core.windows.net"

#Setting up Connection
oracle_connection = oracledb.connect(user='BYHREPO', password='byhalia123', dsn='GODS2P01')
cursor = oracle_connection.cursor()
#Oracle Query
invgdquery="""SELECT DISTINCT TRUNC(SYSDATE-0.19) AS REPORT_DATE, INV.SITE_ID, INV.TAG_ID, INV.SKU_ID, INV.DESCRIPTION, 
INV.CONFIG_ID, INV.LOCATION_ID, INV.QTY_ON_HAND, INV.ZONE_1 AS ZONE, INV.QTY_ALLOCATED,
 TRUNC(MOVE_DSTAMP) AS LAST_MOVE_DATE,INV.PALLET_ID, EACH_HEIGHT, EACH_WIDTH, EACH_DEPTH, 
 HAZMAT, SK.USER_DEF_TYPE_3 AS OPC, STANDARD_COST, SALES_MULTIPLE,  ABC_COUNT,
FIRST_TRY_LOC, V_PUTAWAY_GROUP AS PUTAWAY_GROUP, LOC.LOC_TYPE, INV.V_RIP AS RIP, TRUNC(RECEIPT_DSTAMP) AS RECEIPT_DATE, 
EACH_WEIGHT
FROM DCSDBA.INVENTORY INV
LEFT JOIN DCSDBA.SKU SK ON INV.SKU_ID=SK.SKU_ID AND INV.CLIENT_ID=SK.CLIENT_ID
LEFT JOIN DCSDBA.V_SKU_PROPERTIES VSK ON INV.SKU_ID=VSK.SKU_ID AND INV.SITE_ID=VSK.SITE_ID
LEFT JOIN DCSDBA.PICK_FACE PF ON INV.LOCATION_ID=PF.LOCATION_ID AND PF.SITE_ID=INV.SITE_ID
LEFT JOIN DCSDBA.LOCATION LOC ON INV.LOCATION_ID=LOC.LOCATION_ID AND INV.SITE_ID=LOC.SITE_ID
LEFT JOIN (SELECT SITE_ID, SKU_ID, SUM(ABC_COUNT) AS ABC_COUNT
            FROM DCSDBA.SKU_RANKING 
            WHERE CLIENT_ID='VPNA'
            GROUP BY SITE_ID, SKU_ID) SR ON INV.SKU_ID=SR.SKU_ID AND INV.SITE_ID=SR.SITE_ID
LEFT JOIN DCSDBA.PUTAWAY_LOCATION PL ON INV.SKU_ID=PL.SKU_ID AND INV.SITE_ID=PL.SITE_iD
WHERE INV.CLIENT_ID='VPNA' """
t2=datetime.now().strftime("%H:%M:%S")
print(t2)
cursor.execute(invgdquery)
rows = cursor.fetchall()
print(datetime.now().strftime("%H:%M:%S"))

header=["REPORT_DATE","SITE_ID","TAG_ID","SKU_ID","DESCRIPTION",
        "CONFIG_ID","LOCATION_ID","QTY_ON_HAND","ZONE","QTY_ALLOCATED",
        "LAST_MOVE_DATE","PALLET_ID","EACH_HEIGHT","EACH_WIDTH","EACH_DEPTH",
        "HAZMAT","OPC","STANDARD_COST","SALES_MULTIPLE","ABC_COUNT",
        "FIRST_TRY_LOC","PUTAWAY_GROUP", "LOC_TYPE","RIP"
        "RECEIPT_DATE","EACH_WEIGHT"]
curr_date = datetime.now().strftime('%Y-%m-%d')
x="invsmlnagd" + curr_date +".csv"
dotenv.load_dotenv()

try:
    with open(x,'w',newline='') as f:
        writer=csv.writer(f)
        writer.writerow(header)
        for i in rows:
            writer.writerow(i)



    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    blob_client = blob_service_client.get_blob_client(container_name, x)
    blob_client_2= blob_service_client.get_blob_client(container_name,"todayinvsmlnagd.csv")
    with open(x, 'rb') as data:
        blob_client.upload_blob(data,overwrite=True)

    with open(x, 'rb') as data:
        blob_client_2.upload_blob(data,overwrite=True)

    t3 = datetime.now().strftime("%H:%M:%S")
    print(t3)
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0
    nm = ol.CreateItem(olmailitem)
    nm.To = "swapnil.pednekar@volvo.com"
    nm.Subject = 'Inventory Golden Dataset Uploads Done'
    nm.Body = f'Hello, this is an automated email for Inventory Golden Dataset Dumps that ran from {t2} to {t3}'
    nm.Send()
except:
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0
    nm = ol.CreateItem(olmailitem)
    nm.To = "swapnil.pednekar@volvo.com"
    nm.Subject = 'Inventory Golden Dataset Uploads Failed'
    nm.Body = f'Hello, this is an automated email for Inventory Golden Dataset Dumps that ran from {t2} to {t3}'
    nm.Send()
