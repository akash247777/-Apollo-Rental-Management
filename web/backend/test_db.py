import pyodbc

# Database connection parameters
DB_DRIVER = "{SQL Server}"  # Use the ODBC driver name
DB_SERVER = "APHSC0095-PC"
DB_DATABASE = "RENT"
DB_USER = "apposcr"
DB_PASSWORD = "2#06A9a"

def test_connection():
    try:
        # Create the connection string
        conn_str = f'DRIVER={DB_DRIVER};SERVER={DB_SERVER};DATABASE={DB_DATABASE};UID={DB_USER};PWD={DB_PASSWORD}'
        
        # Connect to the database
        print(f"Attempting to connect to {DB_SERVER}/{DB_DATABASE} as {DB_USER}...")
        conn = pyodbc.connect(conn_str)
        
        # Check if connection is successful
        cursor = conn.cursor()
        cursor.execute("SELECT @@VERSION")
        row = cursor.fetchone()
        
        print(f"Connection successful!")
        print(f"Server version: {row[0]}")
        
        # List all tables in the database
        print("\nTables in the database:")
        cursor.execute("""
            SELECT TABLE_NAME 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_TYPE = 'BASE TABLE'
            ORDER BY TABLE_NAME
        """)
        tables = cursor.fetchall()
        for table in tables:
            print(f"  - {table[0]}")
        
        # Try to find tables related to sites or rental data
        print("\nAttempting to examine potential site tables...")
        for table_name in [t[0] for t in tables]:
            try:
                cursor.execute(f"SELECT TOP 1 * FROM [{table_name}]")
                columns = [column[0] for column in cursor.description]
                print(f"\nTable: {table_name}")
                print(f"Columns: {', '.join(columns)}")
                
                # If we find site_id or similar columns, show some sample data
                potential_id_columns = [col for col in columns if 'id' in col.lower() or 'site' in col.lower() or 'store' in col.lower()]
                if potential_id_columns:
                    id_col = potential_id_columns[0]
                    cursor.execute(f"SELECT TOP 5 {id_col} FROM [{table_name}]")
                    sample_ids = cursor.fetchall()
                    if sample_ids:
                        print(f"Sample {id_col} values: {', '.join([str(row[0]) for row in sample_ids])}")
            except Exception as e:
                print(f"Error examining table {table_name}: {str(e)}")
        
        # Close the connection
        cursor.close()
        conn.close()
        
    except Exception as e:
        print(f"Connection failed: {str(e)}")

if __name__ == "__main__":
    test_connection() 