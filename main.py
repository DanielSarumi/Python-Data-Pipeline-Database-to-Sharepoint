import pandas as pd
from passmanager import start_db_connection, sql_query, sharepoint_cred, delete_list_from_sharepoint, map_dataframe_to_sharepoint, delete_large_list_powershell, test_powershell_connection
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(messages)s')
logger = logging.getLogger(__name__)

conn = start_db_connection()
query = """SELECT [Bar Code], 'GROCERIES' AS Department, Description FROM activesku """
result = sql_query(connection=conn, query=query)
df = pd.read_sql_query(query, conn)

print(df.head())

target_list = "Groceries List"

site = sharepoint_cred()
sp_list = site.List(f"{target_list}")


try:
    logger.info("Checking list size...")
    sp_data = sp_list.GetListItems(fields=["ID"])
    item_count = len(sp_data)
    logger.info(f"Found {item_count} items in list")
    
    if item_count >= 5000:
        logger.info("List has more than 5000 items, using PowerShell deletion...")
        success = delete_large_list_powershell(f"{target_list}")
    else:
        logger.info("List has less than 5000 items, using Python deletion...")
        delete_list_from_sharepoint(sp_list=sp_list, list=sp_list, fields=["ID"])
        success = True
        
except Exception as e:
    logger.info(f"Could not get list count: {e}")
    logger.info("Defaulting to PowerShell deletion...")
    
    # Test PowerShell connection first
    if test_powershell_connection(f"{target_list}"):
        logger.info("PowerShell connection successful, proceeding with deletion...")
        success = delete_large_list_powershell(f"{target_list}")
    else:
        logger.info("PowerShell connection failed too!")
        success = False

if success:
    logger.info("Deletion completed, uploading new data...")
    map_dataframe_to_sharepoint(df=df, sp_list=sp_list)
else:
    logger.info("Deletion failed, skipping upload")