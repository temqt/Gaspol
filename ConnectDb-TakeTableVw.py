import pyodbc
import pandas as pd

# Database connection details for DEV, UAT, PROD (SSO authentication)
environments = {
    'DEV': {
        'server': 'sql01-pl-gp-d.database.windows.net',
        'database': 'DB01-PL-GP-D',
        'driver': '{ODBC Driver 17 for SQL Server}'
    },
    'UAT': {
        'server': 'sql01-pl-gp-a.database.windows.net',
        'database': 'DB01-PL-GP-A',
        'driver': '{ODBC Driver 17 for SQL Server}'
    },
    'PROD': {
        'server': 'sql01-pl-gp-p.database.windows.net',
        'database': 'DB01-PL-GP-P',
        'driver': '{ODBC Driver 17 for SQL Server}'
    }
}

# SQL query to get all tables and views
query = """
SELECT 
    TABLE_TYPE AS Object_Type,
    TABLE_NAME AS Object_Name
FROM 
    INFORMATION_SCHEMA.TABLES
WHERE 
    TABLE_TYPE IN ('BASE TABLE', 'VIEW')
"""

# Function to establish connection and fetch data from a specific environment using SSO (Windows Authentication)
def fetch_data(env_name, env_details):
    conn_str = f"DRIVER={env_details['driver']};SERVER={env_details['server']};DATABASE={env_details['database']};Trusted_Connection=yes"
    conn = pyodbc.connect(conn_str)
    df = pd.read_sql(query, conn)
    conn.close()
    return df

# Create a dictionary to store DataFrames for each environment
dfs = {}

# Fetch data for each environment and store it in the dictionary
for env_name, env_details in environments.items():
    dfs[env_name] = fetch_data(env_name, env_details)

# Export all data to a single Excel file with separate sheets for DEV, UAT, and PROD
output_file = 'database_objects_dev_uat_prod.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for env_name, df in dfs.items():
        df.to_excel(writer, index=False, sheet_name=f'{env_name}_Tables_Views')

print(f'Exported tables and views from DEV, UAT, and PROD to {output_file}')
