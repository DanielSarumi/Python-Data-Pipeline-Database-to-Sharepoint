import os
from dotenv import load_dotenv
import urllib
from sqlalchemy import create_engine
from shareplum import Site, Office365
from shareplum.site import Version
import logging
import pandas as pd
import subprocess
import tempfile

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

load_dotenv('C:/Users/Business Analyst/OneDrive - Justrite Superstore/Desktop/Dynamics Python File/expiry/.env')



def _database_credential():
    return {
        'server': os.getenv('server'),
        'database': os.getenv('database'),
        'user': os.getenv('user'),
        'pass': os.getenv('pass'),
        'email':os.getenv('email'),
        'emailpass':os.getenv('emailpass'),
        'site':os.getenv('site'),
        'sharepointlink':os.getenv('sharepointlink')
    }


def start_db_connection():
    creds = _database_credential()
    params = urllib.parse.quote_plus("DRIVER={SQL Server Native Client 11.0};"
                                    f"SERVER={creds['server']};"
                                    f"DATABASE={creds['database']};"
                                    f"UID={creds['user']};"
                                    f"PWD={creds['pass']}")

    connection = create_engine("mssql+pyodbc:///?odbc_connect={}".format(params))
    # connection_string = (f"mssql+pyodbc://{creds['user']}:{creds['pass']}@{creds['server']}/{creds['database']}?driver=SQL+Server")
    # connection = create_engine(connection_string)
    return connection

def sql_query(connection, query):
    conn = connection
    
def sharepoint_cred():
    creds = _database_credential()
    authcookie = Office365(f"{creds['sharepointlink']}",
                           username=f"{creds['email']}",
                           password=f"{creds['emailpass']}",
                           ).GetCookies()
    site = Site(
        f"{creds['site']}",
        version=Version.v365,
        authcookie=authcookie
    )
    return site

def delete_list_from_sharepoint(sp_list, list: str, fields: list):
    sp_data = list.GetListItems(fields=fields)
    if len(sp_data) == 0:
        try:
            logger.info(f"No item retrieved or deleted")  
        except Exception as e:
            logger.info(f"No item retrieved: {e}")
    else:
        logger.info(f"retrieved {len(sp_data)} rows")
        del_df = pd.DataFrame.from_dict(sp_data)
        list_data = del_df["ID"].tolist()
        sp_list.UpdateListItems(data=list_data, kind="Delete")
        logger.info("All items deleted successfully")

def map_dataframe_to_sharepoint(df, sp_list):
    batch = []
    for index, row in df.iterrows():
        item = {
            'Description': row['Description'],
            'Department': row['Department'],
            'Bar Code': row['Bar Code']
        }
        batch.append(item)
        # Send in batches of 1000 items to avoid SharePoint limitations
    if batch:
        try:
            logger.info(f'Uploading batch of {len(batch)} items')
            sp_list.UpdateListItems(data=batch, kind='New')
            logger.info(f'{len(batch)} rows Data uploaded successfully to sharepoint list')
        except Exception as e:
            logger.error(f'Error uploading batch: {e}')


# NEW FUNCTION - Test PowerShell connection first
def test_powershell_connection(list_name):
    """Test if PowerShell can connect to SharePoint"""
    try:
        creds = _database_credential()
        
        # Simple connection test script
        test_script = f'''
Add-Type -Path "C:\\Program Files\\Common Files\\Microsoft Shared\\Web Server Extensions\\16\\ISAPI\\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\\Program Files\\Common Files\\Microsoft Shared\\Web Server Extensions\\16\\ISAPI\\Microsoft.SharePoint.Client.Runtime.dll"

$SiteURL = "{creds['site']}"
$ListName = "{list_name}"
$Username = "{creds['email']}"
$Password = ConvertTo-SecureString "{creds['emailpass']}" -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential($Username, $Password)

Try {{
    Write-Host "Testing connection to SharePoint..."
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName, $Credentials.Password)

    $Web = $Context.Web
    $TargetList = $Web.Lists.GetByTitle($ListName)
    $Context.Load($TargetList)
    $Context.ExecuteQuery()

    Write-Host "SUCCESS: Connected to list '$ListName'"
    Write-Host "Current item count: $($TargetList.ItemCount)"
}}
Catch {{
    Write-Host "ERROR: Connection failed - $($_.Exception.Message)"
    exit 1
}}
'''
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ps1', delete=False, encoding='utf-8') as temp_file:
            temp_file.write(test_script)
            temp_script_path = temp_file.name

        print("Testing PowerShell connection...")
        result = subprocess.run([
            'powershell.exe', 
            '-ExecutionPolicy', 'Bypass',
            '-File', temp_script_path
        ], capture_output=True, text=True, timeout=60, encoding='utf-8', errors='replace')
        
        # Cleanup
        try:
            os.unlink(temp_script_path)
        except:
            pass
        
        print(f"Test result: {result.stdout}")
        if result.stderr:
            print(f"Test errors: {result.stderr}")
            
        return result.returncode == 0
        
    except Exception as e:
        print(f"PowerShell test error: {e}")
        return False


# NEW FUNCTION - PowerShell deletion for large lists
def delete_large_list_powershell(list_name):
    """Use PowerShell to delete large SharePoint lists (5000+ items)"""
    try:
        creds = _database_credential()
        
        # Create PowerShell script
        ps_script = f'''
Add-Type -Path "C:\\Program Files\\Common Files\\Microsoft Shared\\Web Server Extensions\\16\\ISAPI\\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\\Program Files\\Common Files\\Microsoft Shared\\Web Server Extensions\\16\\ISAPI\\Microsoft.SharePoint.Client.Runtime.dll"

$SiteURL = "{creds['site']}"
$ListName = "{list_name}"
$BatchSize = 500
$Username = "{creds['email']}"
$Password = ConvertTo-SecureString "{creds['emailpass']}" -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential($Username, $Password)

Try {{
    Write-Host "Connecting to SharePoint..."
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName, $Credentials.Password)

    $Web = $Context.Web
    $TargetList = $Web.Lists.GetByTitle($ListName)
    $Context.Load($TargetList)
    $Context.ExecuteQuery()

    Write-Host "Connected to list: $ListName"

    $BatchQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $BatchQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit Paged='TRUE'>$BatchSize</RowLimit></View>"

    $TotalDeleted = 0
    $BatchNumber = 1

    Do {{
        $Context.Load($TargetList)
        $Context.ExecuteQuery()
        
        $CurrentBatch = $TargetList.GetItems($BatchQuery)
        $Context.Load($CurrentBatch)
        $Context.ExecuteQuery()

        If ($CurrentBatch.Count -eq 0) {{ Break }}

        Write-Host "Batch #$BatchNumber - Deleting $($CurrentBatch.Count) items (Remaining: $($TargetList.ItemCount))"

        ForEach ($ListItem in $CurrentBatch) {{
            $TargetList.GetItemById($ListItem.Id).DeleteObject()
        }}
        
        $Context.ExecuteQuery()
        $TotalDeleted += $CurrentBatch.Count
        $BatchNumber++
        
        Start-Sleep -Milliseconds 100

    }} While ($true)

    Write-Host "SUCCESS: Deleted $TotalDeleted items total" -ForegroundColor Green
}}
Catch {{
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}}
'''
        
        # Write to temp file and execute
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ps1', delete=False, encoding='utf-8') as temp_file:
            temp_file.write(ps_script)
            temp_script_path = temp_file.name

        logger.info("Executing PowerShell deletion...")
        
        # Execute with real-time output
        process = subprocess.Popen([
            'powershell.exe', 
            '-ExecutionPolicy', 'Bypass',
            '-File', temp_script_path
        ], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', errors='replace')
        
        # Print output in real-time
        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                print(f"PowerShell: {output.strip()}")
        
        # Wait for completion
        stdout, stderr = process.communicate()
        result_code = process.returncode
        
        # Cleanup
        try:
            os.unlink(temp_script_path)
        except:
            pass
        
        # Check results
        if result_code == 0:
            logger.info("PowerShell deletion completed successfully")
            if stdout:
                print(f"Final output: {stdout}")
            return True
        else:
            logger.error("PowerShell deletion failed")
            if stderr:
                print(f"PowerShell errors: {stderr}")
            if stdout:
                print(f"PowerShell output: {stdout}")
            return False
            
    except Exception as e:
        logger.error(f"Error with PowerShell deletion: {e}")
        return False