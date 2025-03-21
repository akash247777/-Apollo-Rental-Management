import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import datetime as dt
import math
import pyodbc
from dateutil.relativedelta import relativedelta
from tkcalendar import DateEntry  # pip install tkcalendar
import os
import calendar  # For monthrange
from PIL import Image, ImageTk

def login():
    login_win = tk.Tk()
    login_win.title("Login")
    login_win.state('zoomed')  # Keep it fullscreen if you like

    # 1) Load and place the background image
    try:
        bg_path = r"C:\Rent\Apollo.png"  # Use a raw string or escape backslashes
        bg_image = Image.open(bg_path)
        # Resize to match the window (this can be tweaked as needed)
        bg_image = bg_image.resize((1400, 980), Image.Resampling.LANCZOS)
        bg_photo = ImageTk.PhotoImage(bg_image)
        
        bg_label = tk.Label(login_win, image=bg_photo)
        bg_label.image = bg_photo  # Keep a reference
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
    except Exception as e:
        messagebox.showerror("Image Error", f"Could not load background image: {e}")
    
    # 2) Create a Frame to hold the login widgets
    #    Specify relative width/height so it's bigger:
    form_frame = tk.Frame(login_win, bg="#FFB400")
    # For example, make it 30% of the window's width and 40% of its height,
    # placing it at relx=0.8 (80% from the left) and vertically centered (0.5).
    form_frame.place(relx=0.8, rely=0.5, anchor="center", relwidth=0.3, relheight=0.6)

    # 3) Add a title label
    tk.Label(
        form_frame, 
        text="RENTAL DATA MANAGEMENT", 
        font=("Times New Roman", 18, "bold"), 
        bg="white"
    ).pack(pady=20)  # Increased padding for more space

    # 4) Username and Password fields with bold labels
    tk.Label(form_frame, text="Username", font=("Times New Roman", 15), bg="white").pack(pady=30)  # Increased padding for more space
    username_entry = tk.Entry(form_frame, font=("Times New Roman", 15), width=30)
    username_entry.pack(pady=5)
    tk.Label(form_frame, text="Password", font=("Times New Roman", 15), bg="white").pack(pady=30)  # Increased padding for more space
    password_entry = tk.Entry(form_frame, font=("Times New Roman", 15), show="*", width=30)
    password_entry.pack(pady=5)
    
    def check_login():
        username = username_entry.get()
        password = password_entry.get()
        # Dummy log_event function
        def log_event(msg):
            print(msg)
        if (username == "krishna" and password == "krishna@123") or (username == "kuber" and password == "kuber@123"):
            log_event(f"User '{username}' logged in successfully")
            login_win.destroy()
            global root  # Add this line
            root = main_window()  # Store the return value
            root.protocol("WM_DELETE_WINDOW")  # Add this line
            root.mainloop()  # Start the main event loop
        else:
            log_event(f"Failed login attempt with username '{username}'")
            messagebox.showerror("Login Failed", "Invalid username or password")
    
    # 5) Login button
    tk.Button(form_frame, text="Login", font=("Times New Roman", 15), command=check_login).pack(pady=30)  # Increased padding for more space
    
    login_win.mainloop()

data_labels = {}
upd_win = None  # Add this line
LOG_FILE = os.path.join(os.path.dirname(__file__), "edit_log.txt")

def log_event(event):
    try:
        log_dir = os.path.dirname(LOG_FILE)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        with open(LOG_FILE, "a", encoding='utf-8') as log_file:
            timestamp = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_file.write(f"{timestamp} - {event}\n")
    except Exception as e:
        messagebox.showerror("Logging Error", f"Failed to write to log file: {str(e)}")

def log_login(username):
    log_event(f"User '{username}' logged in")

def log_logout(username):
    log_event(f"User '{username}' logged out")

def log_modification(username, site, changes):
    change_details = ", ".join([f"{col}: '{old}' -> '{new}'" for col, old, new in changes])
    log_event(f"User '{username}' modified site '{site}': {change_details}")

def clear_display():
    for lbl in data_labels.values():
        lbl.config(text="")

# ---------------- Database Connection ----------------
DB_DRIVER = "{SQL Server}"
DB_SERVER = "APHSC0095-PC"     # Your server name/instance
DB_DATABASE = "RENT"           # Your database name
DB_USER = "apposcr"            # Your DB user
DB_PASSWORD = "2#06A9a"         # Your DB password

CONNECTION_STRING = f"DRIVER={DB_DRIVER};SERVER={DB_SERVER};DATABASE={DB_DATABASE};UID={DB_USER};PWD={DB_PASSWORD}"
TABLE_NAME = "RENTDETAILS"

# Global variable for currently displayed record (for editing)
last_record = None

def get_db_connection():
    try:
        return pyodbc.connect(CONNECTION_STRING)
    except Exception as e:
        messagebox.showerror("Database Error", f"Error connecting to database: {str(e)}")
        return None

def load_filtered_data():
    conn = get_db_connection()
    if not conn:
        return pd.DataFrame()
    try:
        today_str = dt.datetime.today().strftime("%Y-%m-%d")
        query = f"SELECT * FROM {TABLE_NAME} WHERE CONVERT(date, [CURRENT DATE]) = ?"
        df = pd.read_sql(query, conn, params=[today_str])
        return df
    except Exception as e:
        messagebox.showerror("Database Error", f"Error fetching data: {str(e)}")
        return pd.DataFrame()
    finally:
        conn.close()

# ---------------- Column Definitions ----------------
columns = [
    "SITE", "STORE NAME", "REGION", "DIV", "MANAGER", "ASST.MANAGER", "EXECUTIVE", "D.O.O",
    "SQ.FT", "AGREEMENT DATE", "RENT POSITION DATE", "RENT EFFECTIVE DATE", "AGREEMENT VALID UPTO",
    "CURRENT DATE", "LEASE PERIOD", "RENT_FREE_PERIOD_DAYS", "RENT EFFECTIVE AMOUNT", "PRESENT RENT",
    "HIKE %", "HIKE YEAR", "RENT DEPOSIT", "OWNER NAME-1", "OWNER NAME-2", "OWNER NAME-3",
    "OWNER NAME-4", "OWNER NAME-5", "OWNER NAME-6", "OWNER MOBILE NUMBER", "CURRENT DATE 1",
    "VALIDITY DATE", "GST_NUMBER", "PAN_NUMBER", "TDS_PERCENTAGE", "MATURE","STATUS","REMARKS"
]

df = load_filtered_data()

# ---------------- Helper Functions ----------------
def convert_to_iso(date_str):
    """Convert a dd-mm-yyyy string to yyyy-mm-dd (ISO format)"""
    if not date_str or date_str.lower() == "dd-mm-yyyy":
        raise ValueError("Invalid date input. Please enter a valid date in dd-mm-yyyy format.")
    try:
        d = dt.datetime.strptime(date_str, "%d-%m-%Y")
        return d.strftime("%Y-%m-%d")
    except Exception as e:
        raise ValueError(f"Error converting date '{date_str}': {str(e)}")

def iso_to_ddmm(iso_str):
    """Convert a 'yyyy-mm-dd' string to 'dd-mm-yyyy'. If invalid, return empty."""
    if not iso_str or not isinstance(iso_str, str):
        return ""
    try:
        return dt.datetime.strptime(iso_str, "%Y-%m-%d").strftime("%d-%m-%Y")
    except:
        return ""

def recalc_current_date1(rp_ddmm, current_ddmm):
    """
    Calculate CURRENT DATE 1 as the difference between CURRENT DATE and RENT POSITION DATE.
    For example, if CURRENT DATE = 01-03-2025 and RENT POSITION DATE = 01-06-2024,
    the result is "0 Years, 9 Months, 0 Days".
    """
    try:
        rp = dt.datetime.strptime(rp_ddmm, "%d-%m-%Y").date()
        cd = dt.datetime.strptime(current_ddmm, "%d-%m-%Y").date()
        diff = relativedelta(cd, rp)
        return f"{diff.years} Years, {diff.months} Months, {diff.days} Days"
    except:
        return ""

def recalc_validity_date(avu_ddmm, current_ddmm):
    """
    Calculate VALIDITY DATE as the difference between AGREEMENT VALID UPTO and CURRENT DATE.
    For example, if CURRENT DATE = 01-03-2025 and AGREEMENT VALID UPTO = 31-05-2033,
    the result is "8 Years, 2 Months, 30 Days".
    """
    try:
        avu = dt.datetime.strptime(avu_ddmm, "%d-%m-%Y").date()
        cd = dt.datetime.strptime(current_ddmm, "%d-%m-%Y").date()
        diff = relativedelta(avu, cd)
        return f"{diff.years} Years, {diff.months} Months, {diff.days} Days"
    except:
        return ""

def update_display(record):
    for col, lbl in data_labels.items():
        val = record.get(col, "")
        if col == "HIKE %":
            try:
                num = float(str(val).replace("%", "").strip())
                if num < 1:
                    num *= 100
                val = str(int(round(num)))
            except:
                pass
        lbl.config(text=val)

# ---------------- Security Check for Edit ----------------
def security_check():
    global root
    sec_win = tk.Toplevel(root)
    sec_win.title("Security Access")
    sec_win.geometry("300x300")
    tk.Label(sec_win, text="Enter credentials", font=("Times New Roman", 12, "bold")).pack(pady=10)
    tk.Label(sec_win, text="Username:", font=("Times New Roman", 12)).pack()
    username_entry = tk.Entry(sec_win, font=("Times New Roman", 12))
    username_entry.pack(pady=5)
    tk.Label(sec_win, text="Password:", font=("Times New Roman", 12)).pack()
    password_entry = tk.Entry(sec_win, font=("Times New Roman", 12), show="*")
    password_entry.pack(pady=5)
    result = {"authenticated": False}
    def check_credentials():
        if username_entry.get() == "admin" and password_entry.get() == "admin":
            result["authenticated"] = True
            sec_win.destroy()
        else:
            messagebox.showerror("Access Denied", "Invalid credentials.")
    tk.Button(sec_win, text="Login", font=("Times New Roman", 12), command=check_credentials).pack(pady=10)
    sec_win.grab_set()
    sec_win.wait_window()
    return result["authenticated"]

# ---------------- Search Function ----------------
def search_site():
    global last_record
    site = search_entry.get().strip()
    if not site:
        messagebox.showwarning("Input Error", "Please enter a site name.")
        return
    conn = get_db_connection()
    if not conn:
        return
    try:
        query = f"SELECT * FROM {TABLE_NAME} WHERE [SITE] LIKE ?"
        df_search = pd.read_sql(query, conn, params=[f"%{site}%"])
        if df_search.empty:
            messagebox.showinfo("No Results", "No matching records found.")
            return
        record = df_search.iloc[0].to_dict()
        
        # Update CURRENT DATE with the current system date
        record["CURRENT DATE"] = dt.datetime.today().strftime("%d-%m-%Y")
        for field in ["RENT POSITION DATE", "AGREEMENT VALID UPTO", "RENT EFFECTIVE DATE", "D.O.O"]:
            record[field] = iso_to_ddmm(record.get(field, ""))
        rp_ddmm = record.get("RENT POSITION DATE", "")
        if rp_ddmm:
            record["CURRENT DATE 1"] = recalc_current_date1(rp_ddmm, record["CURRENT DATE"])
        else:
            record["CURRENT DATE 1"] = ""
        avu_ddmm = record.get("AGREEMENT VALID UPTO", "")
        if avu_ddmm:
            record["VALIDITY DATE"] = recalc_validity_date(avu_ddmm, record["CURRENT DATE"])
        else:
            record["VALIDITY DATE"] = ""
        # Convert AGREEMENT DATE to dd-mm-yyyy format
        if record.get("AGREEMENT DATE"):
            try:
                agreement_date = dt.datetime.strptime(record["AGREEMENT DATE"], "%Y-%m-%d")
                record["AGREEMENT DATE"] = agreement_date.strftime("%d-%m-%Y")
            except:
                record["AGREEMENT DATE"] = ""
        last_record = record
        update_display(record)
    except Exception as e:
        messagebox.showerror("Database Error", f"Error during search: {str(e)}")
    finally:
        conn.close()

# ---------------- Update Window ----------------
def open_update_window():
    if not security_check():
        return
    global last_record
    if not last_record:
        messagebox.showinfo("No Record", "No record selected. Please search for a site first.")
        return
    upd_win = tk.Toplevel(root)
    upd_win.title("Edit Record")
    upd_win.state('zoomed')
    tk.Label(upd_win, text="Edit Record", font=("Times New Roman", 14, "bold")).pack(pady=10)
    upd_fields = {}
    upd_frame = tk.Frame(upd_win)
    upd_frame.pack(pady=10, fill="x")
    optional_fields = {"OWNER NAME-2", "OWNER NAME-3", "OWNER NAME-4",
                       "OWNER NAME-5", "OWNER NAME-6", "OWNER MOBILE NUMBER",
                       "REMARKS", "CURRENT DATE 1", "VALIDITY DATE"}
    auto_calc_fields = {"CURRENT DATE 1", "VALIDITY DATE"}
    date_fields = ["D.O.O", "AGREEMENT DATE", "RENT POSITION DATE",
                   "RENT EFFECTIVE DATE", "AGREEMENT VALID UPTO", "CURRENT DATE"]
    for i, col in enumerate(columns):
        lbl_text = col if col in optional_fields else f"{col} *"
        tk.Label(upd_frame, text=lbl_text, font=("Times New Roman", 12)).grid(
            row=i//2, column=(i%2)*2, padx=5, pady=5, sticky="w")
        entry = tk.Entry(upd_frame, font=("Times New Roman", 12), width=30)
        entry_val = last_record.get(col, "")
        entry.insert(0, str(entry_val))
        if col == "SITE" or col in auto_calc_fields:
            entry.config(state="readonly")
        entry.grid(row=i//2, column=(i%2)*2+1, padx=5, pady=5)
        upd_fields[col] = entry

    def recalc_agree_valid_upto(*args):
        try:
            ad_str = upd_fields["AGREEMENT DATE"].get().strip()
            if ad_str.lower() == "dd-mm-yyyy":
                return
            agree_date = dt.datetime.strptime(ad_str, "%d-%m-%Y").date()
            lease_period = int(upd_fields["LEASE PERIOD"].get())
            valid_date = agree_date + relativedelta(years=+lease_period)
            widget = upd_fields["AGREEMENT VALID UPTO"]
            widget.config(state="normal")
            widget.delete(0, tk.END)
            widget.insert(0, valid_date.strftime("%d-%m-%Y"))
            widget.config(state="readonly")
            recalc_validity_date()
        except Exception:
            pass

    def recalc_validity_date():
        try:
            avu_str = upd_fields["AGREEMENT VALID UPTO"].get().strip()
            cd_str = upd_fields["CURRENT DATE"].get().strip()
            avu = dt.datetime.strptime(avu_str, "%d-%m-%Y").date()
            cd = dt.datetime.strptime(cd_str, "%d-%m-%Y").date()
            diff = relativedelta(avu, cd)
            result = f"{diff.years} Years, {diff.months} Months, {diff.days} Days"
            widget = upd_fields["VALIDITY DATE"]
            widget.config(state="normal")
            widget.delete(0, tk.END)
            widget.insert(0, result)
            widget.config(state="readonly")
        except Exception:
            pass

    def recalc_current_date1_event(event):
        try:
            rp_str = upd_fields["RENT POSITION DATE"].get().strip()
            cd_str = upd_fields["CURRENT DATE"].get().strip()
            rp = dt.datetime.strptime(rp_str, "%d-%m-%Y").date()
            cd = dt.datetime.strptime(cd_str, "%d-%m-%Y").date()
            diff = relativedelta(cd, rp)
            result = f"{diff.years} Years, {diff.months}, {diff.days} Days"
            widget = upd_fields["CURRENT DATE 1"]
            widget.config(state="normal")
            widget.delete(0, tk.END)
            widget.insert(0, result)
            widget.config(state="readonly")
        except Exception:
            pass

    

    upd_fields["AGREEMENT DATE"].bind("<FocusOut>", recalc_agree_valid_upto)
    upd_fields["LEASE PERIOD"].bind("<FocusOut>", recalc_agree_valid_upto)
    upd_fields["RENT POSITION DATE"].bind("<FocusOut>", recalc_current_date1_event)
    

    def save_update():
        try:
            global last_record, data_labels, upd_win
            
            # Get values from update fields
            new_data = {col: upd_fields[col].get().strip() for col in columns}
            site_val = new_data["SITE"]
            
            # Convert date fields to ISO format for SQL Server
            date_fields = [
                "D.O.O", 
                "AGREEMENT DATE", 
                "RENT POSITION DATE",
                "RENT EFFECTIVE DATE", 
                "AGREEMENT VALID UPTO", 
                "CURRENT DATE"
            ]
            
            # Convert dates from dd-mm-yyyy to yyyy-mm-dd
            for field in date_fields:
                try:
                    if new_data[field] and new_data[field].lower() != "dd-mm-yyyy":
                        date_obj = dt.datetime.strptime(new_data[field], "%d-%m-%Y")
                        new_data[field] = date_obj.strftime("%Y-%m-%d")
                except ValueError as e:
                    messagebox.showerror("Date Error", f"Invalid date format in {field}. Use dd-mm-yyyy")
                    return
            
            # Fields to track for negotiation
            negotiation_fields = [
                "AGREEMENT VALID UPTO",
                "HIKE %",
                "LEASE PERIOD",
                "PRESENT RENT"
            ]
            
            # Validate mandatory fields
            mandatory_fields = [col for col in columns if col not in optional_fields]
            for field in mandatory_fields:
                if not new_data.get(field):
                    messagebox.showwarning("Input Error", f"'{field}' is mandatory!")
                    if field in upd_fields:
                        upd_fields[field].focus_set()
                    return False  # Return without closing window
                    
                    
            
            conn = get_db_connection()
            if not conn:
                return
                
            try:
                cursor = conn.cursor()
                
                # Record changes in RENT_CHANGES table
                for field in negotiation_fields:
                    old_val = str(last_record.get(field, ''))
                    new_val = str(new_data.get(field, ''))
                    if old_val != new_val:
                        insert_query = """
                            INSERT INTO RENT_CHANGES (Site, ChangeDate, Field, OldValue, NewValue)
                            VALUES (?, GETDATE(), ?, ?, ?)
                        """
                        cursor.execute(insert_query, (site_val, field, old_val, new_val))
                
                # Update the main record
                update_cols = [col for col in columns if col != "SITE"]
                set_clause = ", ".join(f"[{col}] = ?" for col in update_cols)
                update_query = f"UPDATE {TABLE_NAME} SET {set_clause} WHERE [SITE] = ?"
                values = [new_data[col] for col in update_cols] + [site_val]
                
                cursor.execute(update_query, values)
                conn.commit()
                
                log_event(f"Updated record for site: {site_val}")
                messagebox.showinfo("Success", "Record updated successfully!")
                last_record = new_data
                update_display(new_data)
                
            except Exception as e:
                messagebox.showerror("Database Error", f"Error updating data: {str(e)}")
                conn.rollback()
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")

    tk.Button(upd_win, text="Save", font=("Times New Roman", 12, "bold"),
              command=save_update, bg="lightgreen").pack(pady=10)
    tk.Button(upd_win, text="Cancel", font=("Times New Roman", 12),
              command=upd_win.destroy, bg="lightgray").pack(pady=5)
    

# ---------------- New Data Entry Window ----------------
def open_entry_window():
    global last_record, data_labels, entry_fields
    
    # Move the global declarations to the top of the function
    entry_win = tk.Toplevel(root)
    entry_win.title("New Data Entry")
    entry_win.state('zoomed')
    
    # Create new frame for the entry window
    entry_frame = tk.Frame(entry_win)
    entry_frame.pack(fill="both", expand=True)
    
    # Then clear the last record and display
    last_record = None
    clear_display()
    
    tk.Label(entry_win, text="Enter New Record", font=("Times New Roman", 14, "bold")).pack(pady=10)
    entry_fields = {}
    frm = tk.Frame(entry_win)
    frm.pack(pady=10, fill="x")
    optional_fields = {"OWNER NAME-2", "OWNER NAME-3", "OWNER NAME-4",
                       "OWNER NAME-5", "OWNER NAME-6", "OWNER MOBILE NUMBER",
                       "REMARKS", "CURRENT DATE 1", "VALIDITY DATE"}
    date_fields = ["D.O.O", "AGREEMENT DATE", "RENT POSITION DATE",
                   "RENT EFFECTIVE DATE", "AGREEMENT VALID UPTO", "CURRENT DATE"]
    
    def validate_site(P):
        if len(P) > 5:
            return False
        if not P.isdigit():
            return False
        return True

    vcmd = (entry_win.register(validate_site), '%P')
    
    for col in columns:
        lbl_text = col if col in optional_fields else f"{col} *"
        tk.Label(frm, text=lbl_text, font=("Times New Roman", 12)).grid(
            row=columns.index(col)//2, column=(columns.index(col)%2)*2,
            padx=5, pady=5, sticky="w")
        if col == "SITE":
            widget = tk.Entry(frm, font=("Times New Roman", 12), width=30, validate="key", validatecommand=vcmd)
        elif col == "AGREEMENT DATE":
            widget = tk.Entry(frm, font=("Times New Roman", 12), width=30)
            widget.insert(0, "dd-mm-yyyy")
            widget.config(foreground="gray")
            
            def on_focus_in(event):
                if widget.get() == "dd-mm-yyyy":
                    widget.delete(0, tk.END)
                    widget.config(foreground="black")
                    
            def on_focus_out(event):
                if not widget.get():
                    widget.insert(0, "dd-mm-yyyy")
                    widget.config(foreground="gray")
            
            widget.bind("<FocusIn>", on_focus_in)
            widget.bind("<FocusOut>", on_focus_out)
        elif col == "CURRENT DATE":
            widget = tk.Entry(frm, font=("Times New Roman", 12), width=30)
            widget.insert(0, dt.datetime.today().strftime("%d-%m-%Y"))
            widget.config(state="readonly")
        elif col == "LEASE PERIOD":
            widget = ttk.Combobox(frm, values=[str(i) for i in range(0,21)],
                                  font=("Times New Roman", 12), state="readonly", width=28)
            widget.current(0)
        elif col == "AGREEMENT VALID UPTO":
            widget = tk.Entry(frm, font=("Times New Roman", 12), width=30)
            widget.config(state="readonly")
        else:
            widget = tk.Entry(frm, font=("Times New Roman", 12), width=30)
        widget.grid(row=columns.index(col)//2, column=(columns.index(col)%2)*2+1,
                    padx=5, pady=5)
        entry_fields[col] = widget

    def recalc_agree_valid_upto(*args):
        try:
            ad_str = entry_fields["AGREEMENT DATE"].get().strip()
            if ad_str.lower() == "dd-mm-yyyy":
                return
            agree_date = dt.datetime.strptime(ad_str, "%d-%m-%Y").date()
            lease_period = int(entry_fields["LEASE PERIOD"].get())
            valid_date = agree_date + relativedelta(years=+lease_period)
            widget = entry_fields["AGREEMENT VALID UPTO"]
            widget.config(state="normal")
            widget.delete(0, tk.END)
            widget.insert(0, valid_date.strftime("%d-%m-%Y"))
            widget.config(state="readonly")
            recalc_validity_date()
        except Exception:
            pass

    def recalc_validity_date():
        try:
            avu_str = entry_fields["AGREEMENT VALID UPTO"].get().strip()
            cd_str = entry_fields["CURRENT DATE"].get().strip()
            avu = dt.datetime.strptime(avu_str, "%d-%m-%Y").date()
            cd = dt.datetime.strptime(cd_str, "%d-%m-%Y").date()
            diff = relativedelta(avu, cd)
            result = f"{diff.years} Years, {diff.months} Months, {diff.days} Days"
            widget = entry_fields["VALIDITY DATE"]
            widget.config(state="normal")
            widget.delete(0, tk.END)
            widget.insert(0, result)
            widget.config(state="readonly")
        except Exception:
            pass

    def recalc_current_date1_event(event):
        try:
            rp_str = entry_fields["RENT POSITION DATE"].get().strip()
            cd_str = entry_fields["CURRENT DATE"].get().strip()
            rp = dt.datetime.strptime(rp_str, "%d-%m-%Y").date()
            cd = dt.datetime.strptime(cd_str, "%d-%m-%Y").date()
            diff = relativedelta(cd, rp)
            result = f"{diff.years} Years, {diff.months} Months, {diff.days} Days"
            widget = entry_fields["CURRENT DATE 1"]
            widget.config(state="normal")
            widget.delete(0, tk.END)
            widget.insert(0, result)
            widget.config(state="readonly")
        except Exception:
            pass

    entry_fields["AGREEMENT DATE"].bind("<FocusOut>", recalc_agree_valid_upto)
    entry_fields["LEASE PERIOD"].bind("<<ComboboxSelected>>", recalc_agree_valid_upto)
    entry_fields["RENT POSITION DATE"].bind("<FocusOut>", recalc_current_date1_event)

    def save_entry():
        try:
            # First get the data from entry fields
            new_data = {col: entry_fields[col].get().strip() for col in columns}
            
            # Validate mandatory fields
            mandatory_fields = [col for col in columns if col not in optional_fields]
            for field in mandatory_fields:
                if not new_data.get(field):
                    messagebox.showwarning("Input Error", f"'{field}' is mandatory!")
                    # Focus on the field that needs attention
                    if field in entry_fields:
                        entry_fields[field].focus_set()
                    return False  # Return False to prevent window from closing
                    
                     # Simply return without trying to lift any window

            # Now we can log the event because we have new_data
            log_event(f"New entry created:\n"
                     f"Site: {new_data['SITE']}\n"
                     f"Store: {new_data['STORE NAME']}\n"
                     f"Owner: {new_data['OWNER NAME-1']}\n"
                     f"Region: {new_data['REGION']}\n"
                     f"Present Rent: {new_data['PRESENT RENT']}\n"
                     f"Agreement Date: {new_data['AGREEMENT DATE']}\n"
                     f"Lease Period: {new_data['LEASE PERIOD']}")

            # Rest of the function remains the same...
            recalc_agree_valid_upto()
            for field in date_fields:
                try:
                    new_data[field] = convert_to_iso(new_data[field])
                except Exception as e:
                    messagebox.showerror("Date Conversion Error", f"Error in field '{field}': {str(e)}")
                    return
            try:
                agreement_date = entry_fields["AGREEMENT DATE"].get().strip()
                if agreement_date and agreement_date != "dd-mm-yyyy":
                    try:
                        agreement_date = dt.datetime.strptime(agreement_date, "%d-%m-%Y")
                        new_data["AGREEMENT DATE"] = agreement_date.strftime("%Y-%m-%d")
                    except ValueError:
                        messagebox.showerror("Date Error", "Invalid AGREEMENT DATE format. Please use dd-mm-yyyy")
                        return
            except Exception as e:
                messagebox.showerror("Error", f"Unexpected error: {str(e)}")
                return
            conn = get_db_connection()
            if not conn:
                return
            try:
                cursor = conn.cursor()
                check_query = f"SELECT COUNT(*) FROM {TABLE_NAME} WHERE [SITE] = ?"
                cursor.execute(check_query, new_data["SITE"])
                if cursor.fetchone()[0] > 0:
                    messagebox.showwarning("Duplicate Entry", "A record for this site already exists.")
                    conn.close()
                    return
                cursor.fast_executemany = True
                cols = list(new_data.keys())
                placeholders = ", ".join(["?"] * len(cols))
                col_names = ", ".join(f"[{col}]" for col in cols)
                insert_query = f"INSERT INTO {TABLE_NAME} ({col_names}) VALUES ({placeholders})"
                cursor.execute(insert_query, list(new_data.values()))
                conn.commit()
                
                # Add this logging code here
                log_event(f"New entry created - Site: {new_data['SITE']}, "
                         f"Store: {new_data['STORE NAME']}, "
                         f"Owner: {new_data['OWNER NAME-1']}")
                         
                messagebox.showinfo("Success", "New entry added successfully!")
                entry_win.destroy()
            except Exception as e:
                messagebox.showerror("Database Error", f"Error inserting data: {str(e)}")
            finally:
                if conn:
                    conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")

    def upload_excel():
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_path:
            return
        try:
            excel_data = pd.read_excel(file_path, engine="openpyxl")
            excel_data.fillna("", inplace=True)
            date_cols_for_upload = [
                "D.O.O", "AGREEMENT DATE", "RENT POSITION DATE",
                "RENT EFFECTIVE DATE", "AGREEMENT VALID UPTO", "CURRENT DATE"
            ]
            for col in date_cols_for_upload:
                if col in excel_data.columns:
                    excel_data[col] = pd.to_datetime(excel_data[col], format="%d-%m-%Y", errors='coerce')
            today = dt.datetime.today().date()
            if "CURRENT DATE" in excel_data.columns:
                excel_data["CURRENT DATE"] = excel_data["CURRENT DATE"].apply(lambda x: x if pd.notna(x) else today)
            else:
                excel_data["CURRENT DATE"] = today
            conn = get_db_connection()
            if not conn:
                return
            cursor = conn.cursor()
            cursor.fast_executemany = True
            cursor.execute(f"SELECT [SITE] FROM {TABLE_NAME}")
            existing_sites = set(row[0] for row in cursor.fetchall())
            rows_to_insert = []
            for i, row in excel_data.iterrows():
                new_data = row.to_dict()
                for date_col in date_cols_for_upload:
                    if date_col in new_data and pd.notna(new_data[date_col]):
                        new_data[date_col] = new_data[date_col].strftime("%Y-%m-%d")
                    else:
                        if date_col == "CURRENT DATE":
                            new_data[date_col] = today.strftime("%Y-%m-%d")
                        else:
                            new_data[date_col] = ""
                site_val = str(new_data.get("SITE", "")).strip()
                if site_val in existing_sites:
                    continue
                new_data["CURRENT DATE 1"] = ""
                row_tuple = tuple(new_data.get(col, "") for col in columns)
                rows_to_insert.append(row_tuple)
            cols = list(columns)
            placeholders = ", ".join(["?"] * len(cols))
            col_names = ", ".join(f"[{col}]" for col in cols)
            insert_query = f"INSERT INTO {TABLE_NAME} ({col_names}) VALUES ({placeholders})"
            cursor.executemany(insert_query, rows_to_insert)
            conn.commit()
            messagebox.showinfo("Success", "Excel data uploaded successfully!")
        except Exception as e:
            messagebox.showerror("File Error", f"Error uploading data: {str(e)}")
        finally:
            if conn:
                conn.close()

    def clear_entries():
        for col, widget in entry_fields.items():
            if col == "SITE":
                widget.config(validate="none")  # Temporarily disable validation
                widget.delete(0, tk.END)
                widget.config(validate="key")  # Re-enable validation
            else:
                widget.delete(0, tk.END)
                
    def close_window():
        entry_win.destroy()

    btn_frame = tk.Frame(entry_win)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="Upload Excel", font=("Times New Roman", 12),
              command=upload_excel, bg="lightgreen").pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="Save Record", font=("Times New Roman", 12, "bold"),
              command=save_entry, bg="lightblue").pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="Clear Data", font=("Times New Roman", 12),
              command=clear_entries, bg="lightyellow").pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="Previous", font=("Times New Roman", 12),
              command=close_window, bg="lightgray").pack(side=tk.LEFT, padx=10)

# ---------------- Report Window ----------------
def open_report_window():
    global div_combobox, status_combobox, upd_win
    report_win = tk.Toplevel(root)
    report_win.title("Generate Report")
    report_win.state('zoomed')

    # Top frame for report type selection
    top_frame = tk.Frame(report_win)
    top_frame.pack(pady=10, fill="x")
    
    report_type = tk.StringVar(value="Hike Report")
    options = ["Hike Report", "Rent Report", "Owner Wise Report","Negotiation Report","Lease Period Report", "ALL SITES DATA REPORTS"]

    tk.Label(top_frame, text="Report Type:", font=("Times New Roman", 12)).pack(side=tk.LEFT, padx=5)
    
    # Frame for DIV and Status filters (initially empty)
    filter_frame = tk.Frame(report_win)
    filter_frame.pack(pady=10, fill="x")

    # Create the DIV filter widgets (but do not pack them immediately)
    div_label = tk.Label(filter_frame, text="DIV:", font=("Times New Roman", 12))
    div_combobox = ttk.Combobox(filter_frame, values=["SAP", "BOT", "ALL"], font=("Times New Roman", 12), state="readonly")
    div_combobox.current(2)  # Default to "ALL"

    # Create the Status filter widgets (but do not pack them immediately)
    status_label = tk.Label(filter_frame, text="Status:", font=("Times New Roman", 12))
    status_combobox = ttk.Combobox(filter_frame, values=["ONLINE", "OFFLINE", "ALL"], font=("Times New Roman", 12), state="readonly")
    status_combobox.current(2)  # Default to "ALL"

    # Callback function to update filter widgets based on report type selection
    def update_filters(*args):
        if report_type.get() == "ALL SITES DATA REPORTS":
            # Pack DIV and Status filters if the selected report type is "ALL SITES DATA REPORTS"
            div_label.pack(side=tk.LEFT, padx=5)
            div_combobox.pack(side=tk.LEFT, padx=5)
            status_label.pack(side=tk.LEFT, padx=5)
            status_combobox.pack(side=tk.LEFT, padx=5)
        else:
            # Remove the DIV and Status filters if another report type is selected
            div_label.pack_forget()
            div_combobox.pack_forget()
            status_label.pack_forget()
            status_combobox.pack_forget()

    # Attach a trace to update when the report type changes
    report_type.trace("w", update_filters)
    
    # Initialize filter visibility based on the default report type
    update_filters()
    
    # ... rest of your code to generate and display the report ...


    def toggle_lease_frame():
        if report_type.get() == "Lease Period Report":
            lease_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5)
        else:
            lease_frame.grid_forget()
    for opt in options:
        tk.Radiobutton(top_frame, text=opt.upper(), variable=report_type, value=opt,
                       font=("Times New Roman", 12), command=toggle_lease_frame).pack(side=tk.LEFT, padx=3)
    date_frame = tk.Frame(report_win)
    date_frame.pack(pady=10)
    tk.Label(date_frame, text="From Date:", font=("Times New Roman", 12)).grid(row=0, column=0, padx=5, pady=5, sticky="e")
    from_date_entry = DateEntry(date_frame, date_pattern='dd-mm-yyyy', font=("Times New Roman", 12), width=28)
    from_date_entry.grid(row=0, column=1, padx=5, pady=5)
    tk.Label(date_frame, text="To Date:", font=("Times New Roman", 12)).grid(row=1, column=0, padx=5, pady=5, sticky="e")
    to_date_entry = DateEntry(date_frame, date_pattern='dd-mm-yyyy', font=("Times New Roman", 12), width=28)
    to_date_entry.grid(row=1, column=1, padx=5, pady=5)
    lease_frame = tk.Frame(date_frame)
    tk.Label(lease_frame, text="Lease Period:", font=("Times New Roman", 12)).grid(row=0, column=0, padx=5, pady=5, sticky="e")
    lease_combo = ttk.Combobox(lease_frame, values=[str(i) for i in range(0,21)],
                                font=("Times New Roman", 12), state="readonly", width=28)
    lease_combo.current(0)
    lease_combo.grid(row=0, column=1, padx=5, pady=5)
    if report_type.get() == "Lease Period Report":
        lease_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5)
    else:
        lease_frame.grid_forget()
    btn_frame = tk.Frame(report_win)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="Generate Report", font=("Times New Roman", 12, "bold"),
              command=lambda: generate_report(report_type.get(), from_date_entry.get(), to_date_entry.get(),
                                                lease_combo.get() if report_type.get() == "Lease Period Report" else None),
              bg="lightblue").pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="Back", font=("Times New Roman", 12),
              command=report_win.destroy, bg="lightgray").pack(side=tk.LEFT, padx=10)

def generate_report(report_type, from_date, to_date, lease_period=None):
    conn = get_db_connection()
    if not conn:
        return

    # Retrieve selected values from the dropdowns
    selected_div = div_combobox.get()       # Expected values: "SAP", "BOT", "ALL"
    selected_status = status_combobox.get()   # Expected values: "ONLINE", "OFFLINE", "ALL"

    try:
        if report_type == "Owner Wise Report":
            query = f"""
                SELECT SITE, [OWNER NAME-1] as OWNERNAME,
                       [CURRENT DATE], [AGREEMENT DATE], [AGREEMENT VALID UPTO]
                FROM {TABLE_NAME}
            """
            df = pd.read_sql(query, conn)
            report_df = df.copy()
            columns_report = ["SITE", "OWNERNAME", "CURRENT DATE", "AGREEMENT DATE", "AGREEMENT VALID UPTO"]

        elif report_type == "Lease Period Report":
            try:
                lease_val = int(lease_period)  # Convert lease period to integer
            except ValueError:
                messagebox.showerror("Input Error", "Please enter a valid lease period.")
                return

            today = dt.datetime.today()  # System date

            if lease_val != 0:
                # Auto-calculate date range based on system date for report columns
                from_date = today.strftime("%d-%m-%Y")
                try:
                    to_date = today.replace(year=today.year + lease_val).strftime("%d-%m-%Y")
                except ValueError:
                    to_date = (today + relativedelta(years=lease_val)).strftime("%d-%m-%Y")
            else:
                # If lease period is 0, expect manual date input and use those dates for filtering
                try:
                    dt.datetime.strptime(from_date, "%d-%m-%Y")
                    dt.datetime.strptime(to_date, "%d-%m-%Y")
                except ValueError:
                    messagebox.showerror("Date Error", "Please enter valid dates (dd-mm-yyyy).")
                    return

            # Convert to ISO format for date conversion (used only for report columns in nonzero lease period case)
            try:
                from_iso = dt.datetime.strptime(from_date, "%d-%m-%Y").strftime("%Y-%m-%d")
                to_iso = dt.datetime.strptime(to_date, "%d-%m-%Y").strftime("%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Date Error", "Error in converting dates.")
                return

            # Debugging prints
            print("Lease Period:", lease_val)
            print("From Date (ISO):", from_iso)
            print("To Date (ISO):", to_iso)

            # Build SQL Query
            if lease_val == 0:
                # For manual date input, filter by CURRENT DATE range
                query = f"""
                    SELECT SITE, [LEASE PERIOD], [HIKE %], [PRESENT RENT], [RENT DEPOSIT],
                        [RENT POSITION DATE], [CURRENT DATE], [AGREEMENT DATE], [AGREEMENT VALID UPTO]
                    FROM {TABLE_NAME}
                    WHERE CONVERT(date, [CURRENT DATE]) BETWEEN ? AND ?
                """
                params = [from_iso, to_iso]
            else:
                # For nonzero lease period, filter solely on lease period
                query = f"""
                    SELECT SITE, [LEASE PERIOD], [HIKE %], [PRESENT RENT], [RENT DEPOSIT],
                        [RENT POSITION DATE], [CURRENT DATE], [AGREEMENT DATE], [AGREEMENT VALID UPTO]
                    FROM {TABLE_NAME}
                    WHERE [LEASE PERIOD] = ?
                """
                params = [lease_val]

            # Execute SQL Query
            df = pd.read_sql(query, conn, params=params)

            # Debugging - Check if data is retrieved
            print("Data Retrieved:")
            print(df.head())

            if df.empty:
                messagebox.showinfo("No Data", "No records found for the given filters.")
                return

            # Convert key date columns
            df["CURRENT DATE"] = pd.to_datetime(df["CURRENT DATE"], format="%Y-%m-%d", errors="coerce")
            df["Month"] = df["CURRENT DATE"].dt.strftime("%b %Y")

            # Pivot Table: Use the month labels from the CURRENT DATE (if any)
            pivot_df = df.pivot_table(
                index=["SITE", "LEASE PERIOD"],
                columns="Month",
                values="PRESENT RENT",
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            # Aggregate Data
            agg_df = df.groupby([
                "SITE", "LEASE PERIOD", "HIKE %", "AGREEMENT DATE", "RENT POSITION DATE",
                "AGREEMENT VALID UPTO", "RENT DEPOSIT"
            ]).agg({"PRESENT RENT": "sum"}).reset_index()

            # Merge Pivot and Aggregated Data
            report_df = pd.merge(agg_df, pivot_df, on=["SITE", "LEASE PERIOD"], how="left")

            # Build Date Range for Report Columns using auto-calculated from_date and to_date
            from_dt = dt.datetime.strptime(from_date, "%d-%m-%Y").replace(day=1)
            to_dt = dt.datetime.strptime(to_date, "%d-%m-%Y").replace(day=1)
            date_range = pd.date_range(from_dt, to_dt, freq='MS').strftime("%b %Y").tolist()

            # Function to Compute Monthly Rent
            def compute_month_rent(row, month_label):
                try:
                    base_rent = float(row["PRESENT RENT"])
                except:
                    base_rent = 0.0
                try:
                    hike_pct = float(row["HIKE %"])
                except:
                    hike_pct = 0.0
                try:
                    valid_date = dt.datetime.strptime(str(row["AGREEMENT VALID UPTO"]), "%Y-%m-%d")
                except:
                    valid_date = None

                new_rent = base_rent + (base_rent * hike_pct / 100)
                try:
                    month_dt = dt.datetime.strptime(month_label, "%b %Y")
                except:
                    return round(base_rent, 2)

                if not valid_date or (month_dt.year < valid_date.year) or (month_dt.year == valid_date.year and month_dt.month < valid_date.month):
                    return round(base_rent, 2)
                if (month_dt.year > valid_date.year) or (month_dt.year == valid_date.year and month_dt.month > valid_date.month):
                    return round(new_rent, 2)

                days_in_month = calendar.monthrange(month_dt.year, month_dt.month)[1]
                old_daily = base_rent / days_in_month
                new_daily = new_rent / days_in_month
                old_days = valid_date.day
                new_days = days_in_month - old_days
                total_for_month = old_days * old_daily + new_days * new_daily
                return round(total_for_month, 2)

            # Apply Rent Calculation for each month in date_range
            for m in date_range:
                report_df[m] = report_df.apply(lambda row: compute_month_rent(row, m), axis=1)

            # Final Column Order for the report
            columns_report = [
                "SITE", "LEASE PERIOD", "HIKE %", "AGREEMENT DATE",
                "RENT POSITION DATE", "AGREEMENT VALID UPTO", "RENT DEPOSIT",
                "PRESENT RENT"
            ] + date_range

        
        elif report_type == "Hike Report":
            try:
                from_iso = dt.datetime.strptime(from_date, "%d-%m-%Y").strftime("%Y-%m-%d")
                to_iso = dt.datetime.strptime(to_date, "%d-%m-%Y").strftime("%Y-%m-%d")
            except Exception:
                messagebox.showerror("Date Error", "Please enter valid dates (dd-mm-yyyy).")
                return
            query = f"""
                SELECT SITE, [OWNER NAME-1] as OWNERNAME, [PRESENT RENT], [HIKE %], [HIKE YEAR],
                       [CURRENT DATE], [RENT POSITION DATE]
                FROM {TABLE_NAME}
                WHERE CONVERT(date, [CURRENT DATE]) BETWEEN ? AND ?
            """
            df = pd.read_sql(query, conn, params=[from_iso, to_iso])
            report_data = []
            today_date = dt.datetime.today().date()
            for _, row in df.iterrows():
                site = row["SITE"]
                owner = row["OWNERNAME"]
                try:
                    present_rent = float(row["PRESENT RENT"])
                except:
                    present_rent = 0
                try:
                    hike_val = float(row["HIKE %"])
                    if hike_val < 1:
                        hike_val *= 100
                except:
                    hike_val = 0
                try:
                    period_val = int(row["HIKE YEAR"])
                except:
                    period_val = 0
                if period_val <= 0:
                    last_hike = "Not Hiked"
                    next_hike = "Not Hiked"
                else:
                    join_date_str = row["RENT POSITION DATE"]
                    try:
                        join_date = dt.datetime.strptime(join_date_str, "%Y-%m-%d").date()
                    except:
                        join_date = None
                    if join_date is None:
                        last_hike = "Not Hiked"
                        next_hike = "Not Hiked"
                    else:
                        diff = relativedelta(today_date, join_date)
                        total_years = diff.years + diff.months/12 + diff.days/365.25
                        n = int(total_years // period_val)
                        if n < 1:
                            last_hike = "Not Hiked"
                            next_hike_date = join_date + relativedelta(years=+period_val)
                            next_hike = next_hike_date.strftime("%d-%m-%Y")
                        else:
                            last_hike_date = join_date + relativedelta(years=+n * period_val)
                            last_hike = last_hike_date.strftime("%d-%m-%Y")
                            next_hike_date = last_hike_date + relativedelta(years=+period_val)
                            next_hike = next_hike_date.strftime("%d-%m-%Y")
                if period_val <= 0:
                    amount = "N/A"
                else:
                    amount = present_rent + (present_rent * (hike_val/100) * period_val)
                current_date = row["CURRENT DATE"]
                report_data.append([site, owner, present_rent, hike_val, period_val,
                                    current_date, row["RENT POSITION DATE"], last_hike, next_hike, amount])
            report_df = pd.DataFrame(report_data, columns=[
                "SITE", "OWNER NAME", "PRESENT RENT", "HIKE %", "HIKE YEAR",
                "CURRENT DATE", "RENT POSITION DATE", "LAST HIKE", "NEXT HIKE", "AMOUNT"
            ])
            columns_report = report_df.columns.tolist()

        elif report_type == "Rent Report":
            try:
                # Parse the from and to dates
                from_dt = dt.datetime.strptime(from_date, "%d-%m-%Y")
                to_dt = dt.datetime.strptime(to_date, "%d-%m-%Y")
            except Exception:
                messagebox.showerror("Date Error", "Please enter valid dates (dd-mm-yyyy).")
                return

            # Query the database and include [CURRENT DATE] for building the Month column.
            query = f"""
                SELECT SITE, [CURRENT DATE], [HIKE %], [LEASE PERIOD], [AGREEMENT VALID UPTO], [PRESENT RENT], [TDS_PERCENTAGE]
                FROM {TABLE_NAME}
                WHERE CONVERT(date, [CURRENT DATE]) BETWEEN ? AND ?
            """
            df = pd.read_sql(query, conn, params=[from_dt.strftime("%Y-%m-%d"), to_dt.strftime("%Y-%m-%d")])
            
            # Convert numeric columns and compute Net = PRESENT RENT - (PRESENT RENT * TDS_PERCENTAGE/100)
            df["PRESENT RENT"] = pd.to_numeric(df["PRESENT RENT"], errors="coerce").fillna(0)
            df["TDS_PERCENTAGE"] = pd.to_numeric(df["TDS_PERCENTAGE"], errors="coerce").fillna(0)
            df["Net"] = df["PRESENT RENT"] - (df["PRESENT RENT"] * (df["TDS_PERCENTAGE"] / 100))
            
            # Convert CURRENT DATE to datetime and create a Month column (e.g. "Apr 2025")
            df["CURRENT DATE"] = pd.to_datetime(df["CURRENT DATE"], errors="coerce")
            df["Month"] = df["CURRENT DATE"].dt.strftime("%b %Y")
            
            # Build a pivot table: index = SITE, columns = Month, values = Net (summing if multiple rows exist)
            pivot_df = df.pivot_table(
                index="SITE",
                columns="Month",
                values="Net",
                aggfunc="sum",
                fill_value=0
            ).reset_index()
            
            # Aggregate other fields per SITE (use first for HIKE %, LEASE PERIOD, AGREEMENT VALID UPTO,
            # sum for PRESENT RENT and Net, and mean for TDS_PERCENTAGE)
            agg_df = df.groupby("SITE").agg({
                "HIKE %": "first",
                "LEASE PERIOD": "first",
                "AGREEMENT VALID UPTO": "first",
                "PRESENT RENT": "sum",
                "TDS_PERCENTAGE": "mean",
                "Net": "sum"
            }).reset_index()
            
            # Merge the aggregated data with the pivot table.
            report_df = pd.merge(agg_df, pivot_df, on="SITE", how="left")
            
            # Generate a full list of month labels for the selected date range
            date_range = pd.date_range(from_dt, to_dt, freq='MS').strftime("%b %Y").tolist()
            
            # Force creation of a column for each month in date_range, if not present
            for month_label in date_range:
                if month_label not in report_df.columns:
                    report_df[month_label] = 0
            
            # Sort the month columns chronologically
            def parse_month(m_str):
                return dt.datetime.strptime(m_str, "%b %Y")
            # Use the list from date_range
            month_cols = [m for m in date_range if m in report_df.columns]
            month_cols_sorted = sorted(month_cols, key=lambda m: parse_month(m))
            
            # Set the final column order: static columns then month columns
            columns_report = ["SITE", "HIKE %", "LEASE PERIOD", "AGREEMENT VALID UPTO", "PRESENT RENT", "TDS_PERCENTAGE", "NET"] + month_cols_sorted
            
            # Calculate rent for each month considering agreement validity and hike
            for index, row in report_df.iterrows():
                agreement_valid_upto = row["AGREEMENT VALID UPTO"]
                present_rent = row["PRESENT RENT"]
                hike_percentage = row["HIKE %"]
                tds_percentage = row["TDS_PERCENTAGE"]
                agreement_valid_upto_date = dt.datetime.strptime(agreement_valid_upto, "%Y-%m-%d")
                
                for month_label in month_cols_sorted:
                    month_date = dt.datetime.strptime(month_label, "%b %Y")
                    if month_date <= agreement_valid_upto_date:
                        net_rent = present_rent - (present_rent * (tds_percentage / 100))
                    else:
                        hiked_rent = present_rent + (present_rent * (hike_percentage / 100))
                        net_rent = hiked_rent - (hiked_rent * (tds_percentage / 100))
                    report_df.at[index, month_label] = net_rent
                    
                    
        elif report_type == "Negotiation Report":
            try:
                from_iso = dt.datetime.strptime(from_date, "%d-%m-%Y").strftime("%Y-%m-%d")
                to_iso = dt.datetime.strptime(to_date, "%d-%m-%Y").strftime("%Y-%m-%d")
                
                # First, let's check if there are any records in the date range
                check_query = """
                    SELECT COUNT(*) FROM RENT_CHANGES 
                    WHERE CONVERT(date, ChangeDate) BETWEEN ? AND ?
                """
                cursor = conn.cursor()
                cursor.execute(check_query, [from_iso, to_iso])
                count = cursor.fetchone()[0]
                
                if count == 0:
                    messagebox.showinfo("No Data", f"No changes found between {from_date} and {to_date}")
                    return
                    
                # If records exist, proceed with the main query
                query = """
                    SELECT 
                        Site,
                        MAX(CASE WHEN Field = 'AGREEMENT VALID UPTO' THEN OldValue END) as 'Old Agreement Valid Upto',
                        MAX(CASE WHEN Field = 'AGREEMENT VALID UPTO' THEN NewValue END) as 'New Agreement Valid Upto',
                        MAX(CASE WHEN Field = 'HIKE %' THEN OldValue END) as 'Old HIKE %',
                        MAX(CASE WHEN Field = 'HIKE %' THEN NewValue END) as 'New HIKE %',
                        MAX(CASE WHEN Field = 'LEASE PERIOD' THEN OldValue END) as 'Old Lease Period',
                        MAX(CASE WHEN Field = 'LEASE PERIOD' THEN NewValue END) as 'New Lease Period',
                        MAX(CASE WHEN Field = 'PRESENT RENT' THEN OldValue END) as 'Old Present Rent',
                        MAX(CASE WHEN Field = 'PRESENT RENT' THEN NewValue END) as 'New Present Rent'
                    FROM RENT_CHANGES
                    WHERE CONVERT(date, ChangeDate) BETWEEN ? AND ?
                    GROUP BY Site
                """
                
                report_df = pd.read_sql(query, conn, params=[from_iso, to_iso])
                
                # Debug print
                print(f"Records found: {len(report_df)}")
                print(f"Date range: {from_iso} to {to_iso}")
                
                if report_df.empty:
                    messagebox.showinfo("No Data", "No negotiation changes found for the selected period.")
                    return
                    
                columns_report = [
                    "Site",
                    "Old Agreement Valid Upto", "New Agreement Valid Upto",
                    "Old HIKE %", "New HIKE %",
                    "Old Lease Period", "New Lease Period",
                    "Old Present Rent", "New Present Rent"
                ]
            except Exception as e:
                messagebox.showerror("Report Error", f"Error generating negotiation report: {str(e)}")
                print(f"Error details: {str(e)}")  # Debug print
                return
            
        elif report_type == "ALL SITES DATA REPORTS":
            try:
                from_iso = dt.datetime.strptime(from_date, "%d-%m-%Y").strftime("%Y-%m-%d")
                to_iso = dt.datetime.strptime(to_date, "%d-%m-%Y").strftime("%Y-%m-%d")
            except Exception:
                messagebox.showerror("Date Error", "Please enter valid dates (dd-mm-yyyy).")
                return

            # Retrieve selected values from the dropdowns
            selected_div = div_combobox.get()       # Expected values: "SAP", "BOT", "ALL"
            selected_status = status_combobox.get()   # Expected values: "ONLINE", "OFFLINE", "ALL"

            query = f"""
                SELECT *
                FROM {TABLE_NAME}
                WHERE CONVERT(date, [CURRENT DATE]) BETWEEN ? AND ?
            """
            params = [from_iso, to_iso]

            # Add DIV filter if a specific value (other than ALL) is selected
            if selected_div.upper() != "ALL":
                query += " AND [DIV] = ?"
                params.append(selected_div)

            # Add Status filter if a specific value (other than ALL) is selected
            if selected_status.upper() != "ALL":
                query += " AND [status] = ?"
                params.append(selected_status)

            report_df = pd.read_sql(query, conn, params=params)
            columns_report = columns  # Use the existing columns list defined at the top of the script

# Disambiguate owner name if needed
        if "OWNER NAME" in report_df.columns:
            owner_col = "OWNER NAME"
        elif "OWNERNAME" in report_df.columns:
            owner_col = "OWNERNAME"
        else:
            owner_col = None
        if owner_col:
            counts = {}
            new_vals = []
            for name in report_df[owner_col]:
                if name in counts:
                    counts[name] += 1
                    new_vals.append(f"{name} - {counts[name]}")
                else:
                    counts[name] = 1
                    new_vals.append(name)
            report_df[owner_col] = new_vals
        
        report_display = tk.Toplevel(root)
        report_display.title(report_type)
        report_display.state('zoomed')
        back_btn = tk.Button(report_display, text="Back", font=("Times New Roman", 12),
                             command=report_display.destroy, bg="lightgray")
        back_btn.pack(pady=5)
        top_frame = tk.Frame(report_display)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        search_label = "Search (Owner/Site):" if report_type == "Owner Wise Report" else "Search Site:"
        tk.Label(top_frame, text=search_label, font=("Times New Roman", 12)).pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry_report = tk.Entry(top_frame, textvariable=search_var, font=("Times New Roman", 12))
        search_entry_report.pack(side=tk.LEFT, padx=5)
        def search_in_report():
            qs = search_var.get().strip().lower()
            if qs == "":
                filtered_df = report_df
            else:
                if report_type == "Owner Wise Report":
                    filtered_df = report_df[
                        (report_df["SITE"].astype(str).str.lower().str.contains(qs)) |
                        (report_df["OWNERNAME"].astype(str).str.lower().str.contains(qs))
                    ]
                else:
                    filtered_df = report_df[report_df["SITE"].astype(str).str.lower().str.contains(qs)]
            update_tree(filtered_df)
        tk.Button(top_frame, text="Search", font=("Times New Roman", 12), command=search_in_report).pack(side=tk.LEFT, padx=5)
        def export_to_excel():
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel Files", "*.xlsx;*.xls")])
            if file_path:
                try:
                    report_df.to_excel(file_path, index=False)
                    messagebox.showinfo("Export", "Report exported successfully!")
                except Exception as e:
                    messagebox.showerror("Export Error", f"Error exporting report: {str(e)}")
        tk.Button(top_frame, text="Export to Excel", font=("Times New Roman", 12),
                  command=export_to_excel, bg="lightgreen").pack(side=tk.RIGHT, padx=5)
        tree_frame = tk.Frame(report_display)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree = ttk.Treeview(tree_frame, columns=columns_report, show="headings",
                            yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill=tk.BOTH, expand=True)
        for col in columns_report:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor="w")
        def update_tree(df_to_display):
            for item in tree.get_children():
                tree.delete(item)
            for _, row in df_to_display.iterrows():
                tree.insert("", "end", values=list(row))
        update_tree(report_df)
    except Exception as e:
        messagebox.showerror("Report Error", f"Error generating report: {str(e)}")
    finally:
        conn.close()

# ---------------- Main Window ----------------
def main_window():
    global root, search_entry, data_labels
    root = tk.Tk()
    root.title("Excel Site Search & Data Entry")
    root.state('zoomed')
    search_frame = tk.Frame(root)
    search_frame.pack(pady=10)
    tk.Label(search_frame, text="Enter Site ID :", font=("Times New Roman", 12)).pack(side=tk.LEFT, padx=(0,5))
    search_entry = tk.Entry(search_frame, width=30, font=("Times New Roman", 12))
    search_entry.pack(side=tk.LEFT, padx=(0,5))
    tk.Button(search_frame, text="Search", font=("Times New Roman", 12), command=search_site).pack(side=tk.LEFT)
    tk.Button(search_frame, text="New Entry", font=("Times New Roman", 12), command=open_entry_window, bg="lightblue").pack(side=tk.LEFT, padx=10)
    tk.Button(search_frame, text="Edit", font=("Times New Roman", 12, "bold"), command=open_update_window, bg="yellow").pack(side=tk.LEFT, padx=10)
    tk.Button(search_frame, text="REPORT", font=("Times New Roman", 12, "bold"), command=open_report_window, bg="orange").pack(side=tk.LEFT, padx=10)
    canvas = tk.Canvas(root)
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0,0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    data_labels = {}
    non_remarks_cols = [col for col in columns if col != "REMARKS"]
    
    # Define the columns that need special colors
    special_color_columns = {
        "SITE" : "Lightblue",
        "STORE NAME" : "Lightblue",
        "PRESENT RENT": "Lightblue",
        "RENT DEPOSIT": "Lightblue",
        "LEASE PERIOD": "Lightblue",
        "HIKE %": "Lightblue",
        "HIKE YEAR": "Lightblue",
        "RENT POSITION DATE": "Lightblue",
        "AGREEMENT VALID UPTO": "Lightblue",
        "OWNER NAME-1": "Lightblue",
        "AGREEMENT DATE": "Lightblue",
        "RENT_FREE_PERIOD_DAYS": "Lightblue",
        "RENT EFFECTIVE DATE": "Lightblue",
        "SQ.FT": "Lightblue",
        "PAN_NUMBER": "Lightblue"
    }
    
    for i, col in enumerate(non_remarks_cols):
        box = tk.Frame(scrollable_frame, bd=1, relief="solid", padx=10, pady=5)
        row = i // 5
        col_pos = i % 5
        box.grid(row=row, column=col_pos, padx=10, pady=10, sticky="nsew")
        
        # Set the background color based on the column
        bg_color = special_color_columns.get(col, "lightgray")
        
        lbl_title = tk.Label(box, text=col, font=("Times New Roman", 12, "bold"), bg=bg_color, width=25)
        lbl_title.pack()
        lbl_value = tk.Label(box, text="", font=("Times New Roman", 12), bg="white", width=25)
        lbl_value.pack()
        data_labels[col] = lbl_value
    
    remarks_box = tk.Frame(scrollable_frame, bd=2, relief="solid", padx=10, pady=5)
    num_rows = math.ceil(len(non_remarks_cols)/5)
    remarks_box.grid(row=num_rows, column=0, columnspan=5, padx=10, pady=10, sticky="nsew")
    remarks_width = 100
    lbl_title = tk.Label(remarks_box, text="REMARKS", font=("Times New Roman", 12, "bold"), bg="lightblue", width=remarks_width)
    lbl_title.pack()
    lbl_value = tk.Label(remarks_box, text="", font=("Times New Roman", 12), bg="white", width=remarks_width, wraplength=600)
    lbl_value.pack()
    data_labels["REMARKS"] = lbl_value
    for i in range(5):
        scrollable_frame.grid_columnconfigure(i, weight=1, uniform="group1")
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    return root  # Add this line to return the root window

login()

def show_login_page():
    login_win = tk.Toplevel(root)
    login_win.title("Login")
    login_win.state('zoomed')
    tk.Label(login_win, text="Username:", font=("Times New Roman", 12)).pack(pady=5)
    username_entry = tk.Entry(login_win, font=("Times New Roman", 12))
    username_entry.pack(pady=5)
    tk.Label(login_win, text="Password:", font=("Times New Roman", 12)).pack(pady=5)
    password_entry = tk.Entry(login_win, font=("Times New Roman", 12), show="*")
    password_entry.pack(pady=5)
    def check_login():
        if username_entry.get() == "admin" and password_entry.get() == "admin":
            login_win.destroy()
        else:
            messagebox.showerror("Login Failed", "Invalid username or password.")
    tk.Button(login_win, text="Login", font=("Times New Roman", 12), command=check_login).pack(pady=10)
    login_win.grab_set()
    login_win.wait_window()
