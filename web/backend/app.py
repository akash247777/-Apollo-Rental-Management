from flask import Flask, request, jsonify, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from flask_jwt_extended import JWTManager, create_access_token, jwt_required, get_jwt_identity
from flask_cors import CORS
from datetime import datetime, timedelta, date
import bcrypt
import os
import pyodbc
from dotenv import load_dotenv
import pandas as pd
import urllib.parse
from dateutil.relativedelta import relativedelta

# Load environment variables
load_dotenv()

app = Flask(__name__, static_folder='../', static_url_path='')
CORS(app)

# Get database connection details from environment variables
DB_DRIVER = os.getenv('DB_DRIVER', '{SQL Server}')
DB_SERVER = os.getenv('DB_SERVER', 'APHSC0095-PC')
DB_DATABASE = os.getenv('DB_DATABASE', 'RENT')
DB_USER = os.getenv('DB_USER', 'apposcr')
DB_PASSWORD = os.getenv('DB_PASSWORD', '2#06A9a')

# Create the connection string
password = urllib.parse.quote_plus(DB_PASSWORD)
connection_string = f"mssql+pyodbc://{DB_USER}:{password}@{DB_SERVER}/{DB_DATABASE}?driver={DB_DRIVER}"

# Configuration
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key')
app.config['SQLALCHEMY_DATABASE_URI'] = connection_string
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['JWT_SECRET_KEY'] = os.getenv('JWT_SECRET_KEY', 'jwt-secret-key')
app.config['JWT_ACCESS_TOKEN_EXPIRES'] = timedelta(hours=1)

# Initialize extensions
db = SQLAlchemy(app)
jwt = JWTManager(app)

# Models
class User(db.Model):
    __tablename__ = 'USERS'  # If your SQL Server table has a different name
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='user')

# Column mapping for database to frontend field conversion
column_mapping = {
    'SITE': 'site',
    'STORE NAME': 'store_name',
    'REGION': 'region',
    'DIV': 'div',
    'MANAGER': 'manager',
    'ASST.MANAGER': 'asst_manager',
    'EXECUTIVE': 'executive',
    'D.O.O': 'doo',
    'SQ.FT': 'sqft',
    'AGREEMENT DATE': 'agreement_date',
    'RENT POSITION DATE': 'rent_position_date',
    'RENT EFFECTIVE DATE': 'rent_effective_date',
    'AGREEMENT VALID UPTO': 'agreement_valid_upto',
    'CURRENT DATE': 'current_date',
    'LEASE PERIOD': 'lease_period',
    'RENT_FREE_PERIOD_DAYS': 'rent_free_period_days',
    'RENT EFFECTIVE AMOUNT': 'rent_effective_amount',
    'PRESENT RENT': 'present_rent',
    'HIKE %': 'hike_percentage',
    'HIKE YEAR': 'hike_year',
    'RENT DEPOSIT': 'rent_deposit',
    'OWNER NAME-1': 'owner_name1',
    'OWNER NAME-2': 'owner_name2',
    'OWNER NAME-3': 'owner_name3',
    'OWNER NAME-4': 'owner_name4',
    'OWNER NAME-5': 'owner_name5',
    'OWNER NAME-6': 'owner_name6',
    'OWNER MOBILE NUMBER': 'owner_mobile',
    'CURRENT DATE 1': 'current_date1',
    'VALIDITY DATE': 'validity_date',
    'GST_NUMBER': 'gst_number',
    'PAN_NUMBER': 'pan_number',
    'TDS_PERCENTAGE': 'tds_percentage',
    'MATURE': 'mature',
    'STATUS': 'status',
    'REMARKS': 'remarks'
}

# Reverse mapping for frontend to database field conversion
field_mapping = {
    'site_id': 'SITE',
    'store_name': '[STORE NAME]',
    'region': 'REGION',
    'div': 'DIV',
    'manager': 'MANAGER',
    'asst_manager': '[ASST.MANAGER]',
    'executive': 'EXECUTIVE',
    'doo': '[D.O.O]',
    'sqft': '[SQ.FT]',
    'agreement_date': '[AGREEMENT DATE]',
    'rent_position_date': '[RENT POSITION DATE]',
    'rent_effective_date': '[RENT EFFECTIVE DATE]',
    'agreement_valid_upto': '[AGREEMENT VALID UPTO]',
    'current_date': '[CURRENT DATE]',
    'lease_period': '[LEASE PERIOD]',
    'rent_free_period_days': 'RENT_FREE_PERIOD_DAYS',
    'rent_effective_amount': '[RENT EFFECTIVE AMOUNT]',
    'present_rent': '[PRESENT RENT]',
    'hike_percentage': '[HIKE %]',
    'hike_year': '[HIKE YEAR]',
    'rent_deposit': '[RENT DEPOSIT]',
    'owner_name1': '[OWNER NAME-1]',
    'owner_name2': '[OWNER NAME-2]',
    'owner_name3': '[OWNER NAME-3]',
    'owner_name4': '[OWNER NAME-4]',
    'owner_name5': '[OWNER NAME-5]',
    'owner_name6': '[OWNER NAME-6]',
    'owner_mobile': '[OWNER MOBILE NUMBER]',
    'gst_number': 'GST_NUMBER',
    'pan_number': 'PAN_NUMBER',
    'tds_percentage': 'TDS_PERCENTAGE',
    'mature': 'MATURE',
    'status': 'STATUS',
    'remarks': 'REMARKS'
}

class Site(db.Model):
    __tablename__ = 'RENTDETAILS'  # Actual table name from the database
    ENTRY_NO = db.Column(db.Integer, primary_key=True)
    SITE = db.Column('SITE', db.String(10), unique=True, nullable=False)
    STORE_NAME = db.Column('STORE NAME', db.String(100), nullable=False)
    REGION = db.Column(db.String(50), nullable=False)
    DIV = db.Column(db.String(10), nullable=False)
    MANAGER = db.Column(db.String(100), nullable=False)
    ASST_MANAGER = db.Column('ASST.MANAGER', db.String(100), nullable=False)
    EXECUTIVE = db.Column(db.String(100), nullable=False)
    DOO = db.Column('D.O.O', db.Date, nullable=False)
    SQFT = db.Column('SQ.FT', db.Integer, nullable=False)
    AGREEMENT_DATE = db.Column('AGREEMENT DATE', db.Date, nullable=False)
    RENT_POSITION_DATE = db.Column('RENT POSITION DATE', db.Date, nullable=False)
    RENT_EFFECTIVE_DATE = db.Column('RENT EFFECTIVE DATE', db.Date, nullable=False)
    AGREEMENT_VALID_UPTO = db.Column('AGREEMENT VALID UPTO', db.Date)
    CURRENT_DATE = db.Column('CURRENT DATE', db.Date)
    LEASE_PERIOD = db.Column('LEASE PERIOD', db.Integer, nullable=False)
    RENT_FREE_PERIOD_DAYS = db.Column('RENT FREE PERIOD_DAYS', db.Integer, nullable=False)
    RENT_EFFECTIVE_AMOUNT = db.Column('RENT EFFECTIVE AMOUNT', db.Float, nullable=False)
    PRESENT_RENT = db.Column('PRESENT RENT', db.Float, nullable=False)
    HIKE_PERCENTAGE = db.Column('HIKE %', db.Float, nullable=False)
    HIKE_YEAR = db.Column('HIKE YEAR', db.Integer, nullable=False)
    RENT_DEPOSIT = db.Column('RENT DEPOSIT', db.Float, nullable=False)
    OWNER_NAME1 = db.Column('OWNER NAME-1', db.String(100), nullable=False)
    OWNER_NAME2 = db.Column('OWNER NAME-2', db.String(100))
    OWNER_NAME3 = db.Column('OWNER NAME-3', db.String(100))
    OWNER_NAME4 = db.Column('OWNER NAME-4', db.String(100))
    OWNER_NAME5 = db.Column('OWNER NAME-5', db.String(100))
    OWNER_NAME6 = db.Column('OWNER NAME-6', db.String(100))
    OWNER_MOBILE = db.Column('OWNER MOBILE NUMBER', db.String(20))
    CURRENT_DATE1 = db.Column('CURRENT DATE 1', db.String(50))
    VALIDITY_DATE = db.Column('VALIDITY DATE', db.String(50))
    GST_NUMBER = db.Column('GST_NUMBER', db.String(20), nullable=False)
    PAN_NUMBER = db.Column('PAN_NUMBER', db.String(20), nullable=False)
    TDS_PERCENTAGE = db.Column('TDS_PERCENTAGE', db.Float, nullable=False)
    MATURE = db.Column(db.String(3), nullable=False)
    STATUS = db.Column(db.String(10), nullable=False)
    REMARKS = db.Column(db.Text)

# Direct connection for custom queries
def get_db_connection():
    conn = pyodbc.connect(f'DRIVER={{{DB_DRIVER}}};SERVER={DB_SERVER};DATABASE={DB_DATABASE};UID={DB_USER};PWD={DB_PASSWORD}')
    return conn

# Routes
@app.route('/')
def serve_index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/<path:path>')
def serve_static(path):
    return send_from_directory(app.static_folder, path)

@app.route('/api/auth/login', methods=['POST'])
def login():
    try:
        data = request.get_json()
        
        # First check fallback credentials
        if data['username'] in ['krishna', 'kuber'] and data['password'] in ['krishna@123', 'kuber@123']:
            access_token = create_access_token(identity=data['username'])
            return jsonify({
                'access_token': access_token,
                'user': {
                    'username': data['username'],
                    'role': 'admin'
                }
            }), 200
        
        # Then try database authentication
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            
            # Query for user
            cursor.execute("SELECT password, role FROM USERS WHERE username = ?", (data['username'],))
            user_data = cursor.fetchone()
            
            if user_data:
                stored_hash = user_data[0].encode('utf-8')
                input_password = data['password'].encode('utf-8')
                
                if bcrypt.checkpw(input_password, stored_hash):
                    access_token = create_access_token(identity=data['username'])
                    return jsonify({
                        'access_token': access_token,
                        'user': {
                            'username': data['username'],
                            'role': user_data[1]
                        }
                    }), 200
            
            return jsonify({'message': 'Invalid credentials'}), 401
            
        except Exception as e:
            print(f"Database error during login: {str(e)}")
            # If database authentication fails, still return invalid credentials
            return jsonify({'message': 'Invalid credentials'}), 401
        finally:
            if 'cursor' in locals():
                cursor.close()
            if 'conn' in locals():
                conn.close()
                
    except Exception as e:
        print(f"Login error: {str(e)}")
        return jsonify({'message': 'An error occurred during login'}), 500

# Modified site search to use direct SQL connection if ORM fails
@app.route('/api/sites', methods=['GET'])
@jwt_required()
def get_sites():
    site_id = request.args.get('site_id')
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        if site_id:
            # Query a specific site
            cursor.execute("SELECT * FROM RENTDETAILS WHERE SITE = ?", (site_id,))
            row = cursor.fetchone()
            
            if row:
                # Get column names
                columns = [column[0] for column in cursor.description]
                # Create a dictionary from the row
                site_data = {}
                
                # Get current date for calculations
                today_date = datetime.today().date()
                current_date = today_date.strftime("%d-%m-%Y")  # Format current date as dd-mm-yyyy
                
                # First pass: Process all columns and store date values
                rent_position_date = None
                agreement_valid_upto = None
                
                # Debug the raw database values first
                print("\nDEBUG RAW DATA:")
                for i, col in enumerate(columns):
                    value = row[i]
                    if col in ['RENT POSITION DATE', 'AGREEMENT VALID UPTO']:
                        print(f"Raw database {col}: {value}, Type: {type(value).__name__}")
                
                for i, col in enumerate(columns):
                    value = row[i]
                    
                    # Handle different types of date values
                    if col in ['RENT POSITION DATE', 'AGREEMENT VALID UPTO'] and value:
                        try:
                            # Handle datetime objects
                            if isinstance(value, datetime):
                                date_str = value.strftime("%d-%m-%Y")
                                print(f"Converted datetime {col}: {date_str}")
                                if col == 'RENT POSITION DATE':
                                    rent_position_date = value.date()
                                elif col == 'AGREEMENT VALID UPTO':
                                    agreement_valid_upto = value.date()
                            # Handle date objects
                            elif isinstance(value, date):
                                date_str = value.strftime("%d-%m-%Y")
                                print(f"Converted date {col}: {date_str}")
                                if col == 'RENT POSITION DATE':
                                    rent_position_date = value
                                elif col == 'AGREEMENT VALID UPTO':
                                    agreement_valid_upto = value
                            # Handle string values
                            else:
                                date_str = str(value)
                                print(f"String value for {col}: {date_str}")
                                # Try to parse date string (handle multiple formats)
                                try:
                                    formats_to_try = ['%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y']
                                    parsed_date = None
                                    
                                    for fmt in formats_to_try:
                                        try:
                                            parsed_date = datetime.strptime(date_str[:10], fmt).date()
                                            date_str = parsed_date.strftime("%d-%m-%Y")
                                            print(f"Successfully parsed {col} using format {fmt}: {date_str}")
                                            break
                                        except ValueError:
                                            continue
                                    
                                    if parsed_date:
                                        if col == 'RENT POSITION DATE':
                                            rent_position_date = parsed_date
                                        elif col == 'AGREEMENT VALID UPTO':
                                            agreement_valid_upto = parsed_date
                                        else:
                                            print(f"Failed to parse {col} with any known format: {date_str}")
                                    else:
                                        print(f"Failed to parse {col} with any known format: {date_str}")
                                except Exception as e:
                                    print(f"Error parsing {col}: {str(e)}")
                        except Exception as e:
                            print(f"Error processing {col}: {str(e)}")
                    
                    # Map column names to frontend keys
                    key = column_mapping.get(col, col.lower())
                    
                    # Special handling for the SITE column
                    if col == 'SITE':
                        site_data['site'] = value  # Standard mapping
                        site_data['site_id'] = value  # Additional mapping for consistency
                        print(f"Set site ID value: {value}")
                    # Convert any datetime values to strings in dd-mm-yyyy format
                    elif isinstance(value, datetime) or isinstance(value, date):
                        value = value.strftime("%d-%m-%Y")
                    
                    site_data[key] = value
                
                # Debug the dates we extracted
                print(f"\nProcessed dates:")
                print(f"Current date: {today_date} ({type(today_date)})")
                print(f"RENT POSITION DATE: {rent_position_date} ({type(rent_position_date) if rent_position_date else 'None'})")
                print(f"AGREEMENT VALID UPTO: {agreement_valid_upto} ({type(agreement_valid_upto) if agreement_valid_upto else 'None'})")
                
                # Calculate CURRENT DATE 1 (difference between current date and rent position date)
                if rent_position_date:
                    try:
                        diff = relativedelta(today_date, rent_position_date)
                        years = abs(diff.years)
                        months = abs(diff.months)
                        days = abs(diff.days)
                        current_date1 = f"{years} Years, {months} Months, {days} Days"
                        print(f"CURRENT DATE 1 calculation: {today_date} - {rent_position_date} = {current_date1}")
                        site_data['current_date1'] = current_date1
                    except Exception as e:
                        print(f"Error calculating CURRENT DATE 1: {str(e)}")
                        site_data['current_date1'] = ""
                else:
                    print("No valid RENT POSITION DATE found for CURRENT DATE 1 calculation")
                    site_data['current_date1'] = ""
                
                # Calculate VALIDITY DATE (difference between agreement valid upto and current date)
                if agreement_valid_upto:
                    try:
                        diff = relativedelta(agreement_valid_upto, today_date)
                        years = abs(diff.years)
                        months = abs(diff.months)
                        days = abs(diff.days)
                        validity_date = f"{years} Years, {months} Months, {days} Days"
                        print(f"VALIDITY DATE calculation: {agreement_valid_upto} - {today_date} = {validity_date}")
                        site_data['validity_date'] = validity_date
                    except Exception as e:
                        print(f"Error calculating VALIDITY DATE: {str(e)}")
                        site_data['validity_date'] = ""
                else:
                    print("No valid AGREEMENT VALID UPTO found for VALIDITY DATE calculation")
                    site_data['validity_date'] = ""
                
                # Special handling for HIKE %
                if 'hike_percentage' in site_data and site_data['hike_percentage'] is not None:
                    try:
                        hike_val = float(site_data['hike_percentage'])
                        if hike_val < 1:  # If value is in decimal form (e.g., 0.15)
                            hike_val *= 100
                        site_data['hike_percentage'] = hike_val
                    except (ValueError, TypeError):
                        site_data['hike_percentage'] = 0
                
                # Set current date
                site_data['current_date'] = current_date
                
                # Debug: Print the final data being sent
                print("Sending site data:", site_data)
                print("Special fields check:")
                print("site:", site_data.get('site'))
                print("site_id:", site_data.get('site_id'))
                print("current_date1:", site_data.get('current_date1'))
                print("validity_date:", site_data.get('validity_date'))
                print("hike_percentage:", site_data.get('hike_percentage'))
                
                return jsonify({'site': site_data}), 200
            else:
                return jsonify({'message': 'Site not found'}), 404
        else:
            # Query all sites (limit to 100)
            cursor.execute("SELECT TOP 100 SITE, [STORE NAME], REGION, DIV, [PRESENT RENT], [LEASE PERIOD], [HIKE %], STATUS FROM RENTDETAILS")
            rows = cursor.fetchall()
            
            sites_data = []
            for row in rows:
                sites_data.append({
                    'site_id': row[0],
                    'store_name': row[1],
                    'region': row[2],
                    'div': row[3],
                    'present_rent': row[4],
                    'lease_period': row[5],
                    'hike_percentage': row[6],
                    'status': row[7]
                })
            
            return jsonify({'sites': sites_data}), 200
    
    except Exception as e:
        print(f"SQL error: {str(e)}")
        return jsonify({'message': f'Database error: {str(e)}'}), 500
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'conn' in locals():
            conn.close()

@app.route('/api/sites', methods=['POST'])
@jwt_required()
def create_site():
    data = request.get_json()
    
    # Validate required fields
    required_fields = ['site_id', 'store_name', 'region', 'div', 'manager', 'asst_manager', 
                      'executive', 'doo', 'sqft', 'agreement_date', 'rent_position_date',
                      'rent_effective_date', 'lease_period', 'rent_free_period_days',
                      'rent_effective_amount', 'present_rent', 'hike_percentage', 'hike_year',
                      'rent_deposit', 'owner_name1', 'gst_number', 'pan_number',
                      'tds_percentage', 'mature', 'status']
    
    for field in required_fields:
        if field not in data:
            return jsonify({'message': f'Missing required field: {field}'}), 400
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Check if site ID already exists
        cursor.execute("SELECT COUNT(*) FROM RENTDETAILS WHERE SITE = ?", (data['site_id'],))
        count = cursor.fetchone()[0]
        if count > 0:
            return jsonify({'message': f"Site ID {data['site_id']} already exists"}), 400
        
        # Prepare the SQL query
        columns = [
            "SITE", "[STORE NAME]", "REGION", "DIV", "MANAGER", "[ASST.MANAGER]", "EXECUTIVE", 
            "[D.O.O]", "[SQ.FT]", "[AGREEMENT DATE]", "[RENT POSITION DATE]", "[RENT EFFECTIVE DATE]", 
            "[LEASE PERIOD]", "RENT_FREE_PERIOD_DAYS", "[RENT EFFECTIVE AMOUNT]", "[PRESENT RENT]", 
            "[HIKE %]", "[HIKE YEAR]", "[RENT DEPOSIT]", "[OWNER NAME-1]", "[OWNER NAME-2]", 
            "[OWNER NAME-3]", "[OWNER NAME-4]", "[OWNER NAME-5]", "[OWNER NAME-6]", 
            "[OWNER MOBILE NUMBER]", "GST_NUMBER", "PAN_NUMBER", "TDS_PERCENTAGE", "MATURE", 
            "STATUS", "REMARKS"
        ]
        
        # Optional fields
        if 'agreement_valid_upto' in data and data['agreement_valid_upto']:
            columns.append("[AGREEMENT VALID UPTO]")
        
        if 'current_date' in data and data['current_date']:
            columns.append("[CURRENT DATE]")
            
        if 'current_date1' in data and data['current_date1']:
            columns.append("[CURRENT DATE 1]")
            
        if 'validity_date' in data and data['validity_date']:
            columns.append("[VALIDITY DATE]")
        
        # Create placeholders for values
        placeholders = ['?'] * len(columns)
        
        # Prepare values
        values = [
            data['site_id'], data['store_name'], data['region'], data['div'], 
            data['manager'], data['asst_manager'], data['executive'], 
            data['doo'], data['sqft'], data['agreement_date'], data['rent_position_date'], 
            data['rent_effective_date'], data['lease_period'], data['rent_free_period_days'], 
            data['rent_effective_amount'], data['present_rent'], data['hike_percentage'], 
            data['hike_year'], data['rent_deposit'], data['owner_name1'],
            data.get('owner_name2'), data.get('owner_name3'), data.get('owner_name4'),
            data.get('owner_name5'), data.get('owner_name6'), data.get('owner_mobile'),
            data['gst_number'], data['pan_number'], data['tds_percentage'], 
            data['mature'], data['status'], data.get('remarks')
        ]
        
        # Add optional values
        if 'agreement_valid_upto' in data and data['agreement_valid_upto']:
            values.append(data['agreement_valid_upto'])
        
        if 'current_date' in data and data['current_date']:
            values.append(data['current_date'])
            
        if 'current_date1' in data and data['current_date1']:
            values.append(data['current_date1'])
            
        if 'validity_date' in data and data['validity_date']:
            values.append(data['validity_date'])
        
        # Execute the SQL query
        query = f"INSERT INTO RENTDETAILS ({', '.join(columns)}) VALUES ({', '.join(placeholders)})"
        cursor.execute(query, values)
        conn.commit()
        
        return jsonify({'message': 'Site created successfully'}), 201
    except Exception as e:
        print(f"Error creating site: {str(e)}")
        return jsonify({'message': f"Error creating site: {str(e)}"}), 400
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'conn' in locals():
            conn.close()

@app.route('/api/sites/<site_id>', methods=['PUT'])
@jwt_required()
def update_site(site_id):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        data = request.get_json()
        print("Received update data:", data)  # Debug print

        # Define numeric fields that need cleaning
        numeric_fields = [
            'sqft', 'lease_period', 'rent_free_period_days',
            'rent_effective_amount', 'present_rent',
            'hike_percentage', 'hike_year',
            'rent_deposit', 'tds_percentage'
        ]

        # Clean numeric fields
        for field in numeric_fields:
            if field in data and data[field]:
                try:
                    cleaned_value = str(data[field]).replace('₹', '').replace(',', '').replace('%', '').strip()
                    if cleaned_value:
                        if field in ['sqft', 'lease_period', 'rent_free_period_days', 'hike_year']:
                            data[field] = int(float(cleaned_value))
                        else:
                            data[field] = float(cleaned_value)
                except Exception as e:
                    print(f"Error cleaning numeric field {field}: {str(e)}")
                    continue

        # Handle date fields with proper conversion
        date_fields = [
            'agreement_date',
            'rent_position_date', 
            'rent_effective_date',
            'agreement_valid_upto',
            'current_date',
            'doo'
        ]
        
        # Convert date strings to SQL Server format (YYYY-MM-DD)
        for field in date_fields:
            if field in data and data[field]:
                try:
                    # First convert to string if not already
                    date_str = str(data[field]).strip()
                    print(f"Original {field} value: '{date_str}'")
                    
                    # Skip empty strings
                    if not date_str:
                        data.pop(field, None)
                        continue
                        
                    formatted_date = None
                    
                    # Try different date formats
                    if '-' in date_str:
                        date_parts = date_str.split('-')
                        if len(date_parts) == 3:
                            # Handle YYYY-MM-DD format
                            if len(date_parts[0]) == 4 and date_parts[0].isdigit():
                                try:
                                    # Already in correct format, validate it
                                    temp_date = datetime.strptime(date_str, '%Y-%m-%d')
                                    formatted_date = temp_date.strftime('%Y-%m-%d')
                                    print(f"Validated YYYY-MM-DD format: {formatted_date}")
                                except ValueError as e:
                                    print(f"Failed to validate YYYY-MM-DD: {e}")
                                    pass
                                
                            # Handle DD-MM-YYYY format
                            elif len(date_parts[2]) == 4 and date_parts[2].isdigit():
                                try:
                                    day, month, year = date_parts
                                    # Create a valid date string in SQL format
                                    formatted_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                                    # Validate it can be parsed
                                    datetime.strptime(formatted_date, '%Y-%m-%d')
                                    print(f"Converted DD-MM-YYYY to YYYY-MM-DD: {formatted_date}")
                                except ValueError as e:
                                    print(f"Failed to convert DD-MM-YYYY: {e}")
                                    formatted_date = None
            
                    # Try parsing with multiple formats if not yet parsed
                    if not formatted_date:
                        formats_to_try = ['%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d']
                        for fmt in formats_to_try:
                            try:
                                temp_date = datetime.strptime(date_str, fmt)
                                formatted_date = temp_date.strftime('%Y-%m-%d')
                                print(f"Parsed with format {fmt}: {formatted_date}")
                                break
                            except ValueError:
                                continue
            
                    if formatted_date:
                        print(f"Final {field} value: '{formatted_date}'")
                        data[field] = formatted_date
                    else:
                        print(f"Could not parse date for {field}: '{date_str}', removing from update")
                        data.pop(field, None)
                        
                except Exception as e:
                    print(f"Date conversion error for {field}: {str(e)}, value: '{data.get(field)}'")
                    # Remove invalid date field to prevent SQL error
                    data.pop(field, None)
                    continue

        # Build the SET clause and values for the update query
        set_clauses = []
        values = []
        
        for key, value in data.items():
            db_field = field_mapping.get(key)
            if db_field and value is not None:
                set_clauses.append(f"{db_field} = ?")
                values.append(value)
        
        if not set_clauses:
            return jsonify({'message': 'No fields to update'}), 400
        
        values.append(site_id)
        
        query = f"UPDATE RENTDETAILS SET {', '.join(set_clauses)} WHERE SITE = ?"
        print("Executing query:", query)  # Debug print
        print("With values:", values)  # Debug print
        
        cursor.execute(query, values)
        
        if cursor.rowcount == 0:
            return jsonify({'message': 'No records were updated'}), 404
        
        conn.commit()
        return jsonify({'message': 'Site updated successfully'}), 200
        
    except Exception as e:
        print(f"Error updating site: {str(e)}")
        return jsonify({'message': f"Error updating site: {str(e)}"}), 400
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'conn' in locals():
            conn.close()

@app.route('/api/reports', methods=['GET'])
@jwt_required()
def get_report():
    report_type = request.args.get('type')
    from_date = request.args.get('from_date')
    to_date = request.args.get('to_date')
    lease_period = request.args.get('lease_period')
    
    if not all([report_type, from_date, to_date]):
        return jsonify({'message': 'Missing required parameters'}), 400
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Convert dates to SQL format
        from_date_obj = datetime.strptime(from_date, '%Y-%m-%d').date()
        to_date_obj = datetime.strptime(to_date, '%Y-%m-%d').date()
        
        # Base query
        base_query = """
            SELECT * FROM RENTDETAILS 
            WHERE [AGREEMENT DATE] BETWEEN ? AND ?
        """
        params = [from_date_obj, to_date_obj]
        
        # Add lease period filter if specified
        if lease_period:
            base_query += " AND [LEASE PERIOD] = ?"
            params.append(int(lease_period))
        
        # Execute the query
        cursor.execute(base_query, params)
        rows = cursor.fetchall()
        columns = [column[0] for column in cursor.description]
        
        # Convert rows to dictionaries
        sites = []
        for row in rows:
            site = {}
            for i, col in enumerate(columns):
                value = row[i]
                if isinstance(value, datetime):
                    value = value.strftime('%Y-%m-%d')
                site[col] = value
            sites.append(site)
        
        # Format data based on report type
        data = []
        if report_type == 'Hike Report':
            for site in sites:
                agreement_date = datetime.strptime(site['AGREEMENT DATE'], '%Y-%m-%d')
                hike_year = int(site['HIKE YEAR'])
                last_hike = agreement_date + timedelta(days=365*hike_year)
                next_hike = agreement_date + timedelta(days=365*(hike_year+1))
                
                data.append({
                    'site_id': site['SITE'],
                    'owner_name': site['OWNER NAME-1'],
                    'present_rent': site['PRESENT RENT'],
                    'hike_percentage': site['HIKE %'],
                    'hike_year': site['HIKE YEAR'],
                    'last_hike': last_hike.strftime('%Y-%m-%d'),
                    'next_hike': next_hike.strftime('%Y-%m-%d'),
                    'amount': site['PRESENT RENT'] * (1 + site['HIKE %']/100)
                })
        elif report_type == 'Rent Report':
            for site in sites:
                data.append({
                    'site_id': site['SITE'],
                    'present_rent': site['PRESENT RENT'],
                    'tds_percentage': site['TDS_PERCENTAGE'],
                    'net': site['PRESENT RENT'] * (1 - site['TDS_PERCENTAGE']/100),
                    'jan_2023': site['PRESENT RENT'],
                    'feb_2023': site['PRESENT RENT'],
                    'mar_2023': site['PRESENT RENT'],
                    'apr_2023': site['PRESENT RENT']
                })
        elif report_type == 'Owner Wise Report':
            for site in sites:
                data.append({
                    'site_id': site['SITE'],
                    'owner_name': site['OWNER NAME-1'],
                    'current_date': site['CURRENT DATE'] if 'CURRENT DATE' in site else None,
                    'agreement_date': site['AGREEMENT DATE'],
                    'agreement_valid_upto': site['AGREEMENT VALID UPTO'] if 'AGREEMENT VALID UPTO' in site else None
                })
        elif report_type == 'Negotiation Report':
            for site in sites:
                data.append({
                    'site_id': site['SITE'],
                    'old_hike_percentage': site['HIKE %'],
                    'new_hike_percentage': site['HIKE %'] + 2,  # Example
                    'old_lease_period': site['LEASE PERIOD'],
                    'new_lease_period': site['LEASE PERIOD'] + 1,  # Example
                    'old_present_rent': site['PRESENT RENT'],
                    'new_present_rent': site['PRESENT RENT'] * 1.1  # Example
                })
        elif report_type == 'Lease Period Report':
            for site in sites:
                data.append({
                    'site_id': site['SITE'],
                    'lease_period': site['LEASE PERIOD'],
                    'hike_percentage': site['HIKE %'],
                    'present_rent': site['PRESENT RENT'],
                    'agreement_valid_upto': site['AGREEMENT VALID UPTO'] if 'AGREEMENT VALID UPTO' in site else None
                })
        elif report_type == 'ALL SITES DATA REPORTS':
            for site in sites:
                data.append({
                    'site_id': site['SITE'],
                    'store_name': site['STORE NAME'],
                    'region': site['REGION'],
                    'div': site['DIV'],
                    'status': site['STATUS'],
                    'present_rent': site['PRESENT RENT'],
                    'lease_period': site['LEASE PERIOD'],
                    'hike_percentage': site['HIKE %']
                })
        else:
            return jsonify({'message': 'Invalid report type'}), 400
        
        return jsonify({'data': data}), 200
    except Exception as e:
        print(f"Report generation error: {str(e)}")
        return jsonify({'message': f'Error generating report: {str(e)}'}), 500
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'conn' in locals():
            conn.close()

@app.route('/api/upload', methods=['POST'])
@jwt_required()
def upload_excel():
    if 'file' not in request.files:
        return jsonify({'message': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'message': 'No file selected'}), 400
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'message': 'Invalid file format'}), 400
    
    try:
        # Read the Excel file
        df = pd.read_excel(file)
        
        # Get a database connection
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Define columns that must be present in the Excel file
        required_columns = [
            'SITE', 'STORE NAME', 'REGION', 'DIV', 'MANAGER', 'ASST.MANAGER', 
            'EXECUTIVE', 'D.O.O', 'SQ.FT', 'AGREEMENT DATE', 'RENT POSITION DATE', 
            'RENT EFFECTIVE DATE', 'LEASE PERIOD', 'RENT_FREE_PERIOD_DAYS',
            'RENT EFFECTIVE AMOUNT', 'PRESENT RENT', 'HIKE %', 'HIKE YEAR', 
            'RENT DEPOSIT', 'OWNER NAME-1', 'GST_NUMBER', 'PAN_NUMBER', 
            'TDS_PERCENTAGE', 'MATURE', 'STATUS'
        ]
        
        # Check if required columns are present
        for col in required_columns:
            if col not in df.columns:
                return jsonify({'message': f'Missing required column: {col}'}), 400
        
        # Count of successful inserts
        inserted_count = 0
        
        # Process each row in the Excel file
        for _, row in df.iterrows():
            # Check if site already exists
            site_id = str(row['SITE'])
            cursor.execute("SELECT COUNT(*) FROM RENTDETAILS WHERE SITE = ?", (site_id,))
            count = cursor.fetchone()[0]
            
            # If site exists, skip it
            if count > 0:
                continue
                
            # Prepare columns and values for insertion
            columns = []
            placeholders = []
            values = []
            
            # Process each column in the dataframe
            for col in df.columns:
                if pd.notna(row[col]):  # Only include non-NA values
                    # Format column name for SQL
                    sql_col = col
                    if ' ' in col or '.' in col or '-' in col:
                        sql_col = f"[{col}]"
                    
                    columns.append(sql_col)
                    placeholders.append('?')
                    
                    # Handle date objects
                    value = row[col]
                    if isinstance(value, pd.Timestamp):
                        value = value.to_pydatetime().date()
                    
                    values.append(value)
            
            # Build and execute the INSERT query
            query = f"INSERT INTO RENTDETAILS ({', '.join(columns)}) VALUES ({', '.join(placeholders)})"
            cursor.execute(query, values)
            inserted_count += 1
        
        # Commit the transaction
        conn.commit()
        
        return jsonify({
            'message': f'Data uploaded successfully. {inserted_count} new records inserted.'
        }), 200
    except Exception as e:
        print(f"Upload error: {str(e)}")
        if 'conn' in locals():
            conn.rollback()
        return jsonify({'message': f'Error uploading data: {str(e)}'}), 400
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'conn' in locals():
            conn.close()

if __name__ == '__main__':
    with app.app_context():
        try:
            # Create USERS table if it doesn't exist
            conn = get_db_connection()
            cursor = conn.cursor()
            
            # Check if USERS table exists
            cursor.execute("""
                IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'USERS')
                BEGIN
                    CREATE TABLE USERS (
                        id INT IDENTITY(1,1) PRIMARY KEY,
                        username NVARCHAR(80) UNIQUE NOT NULL,
                        password NVARCHAR(120) NOT NULL,
                        role NVARCHAR(20) NOT NULL DEFAULT 'user'
                    )
                END
            """)
            conn.commit()

            # Create default admin user if not exists
            cursor.execute("SELECT COUNT(*) FROM USERS WHERE username = 'admin'")
            admin_exists = cursor.fetchone()[0]
            
            if not admin_exists:
                hashed = bcrypt.hashpw('admin123'.encode('utf-8'), bcrypt.gensalt())
                cursor.execute(
                    "INSERT INTO USERS (username, password, role) VALUES (?, ?, ?)",
                    ('admin', hashed.decode('utf-8'), 'admin')
                )
                conn.commit()
            
            cursor.close()
            conn.close()
            
        except Exception as e:
            print(f"Database initialization error: {str(e)}")
            
    app.run(debug=True)