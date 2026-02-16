import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import gspread
from datetime import datetime
import threading
import requests
import string
import random
import firebase_admin
from firebase_admin import credentials, auth
import os
import re
import socket
import sys
import subprocess
import tempfile
import csv
import json
import hashlib
import time

try:
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    GOOGLE_LOGIN_AVAILABLE = True
except ImportError:
    GOOGLE_LOGIN_AVAILABLE = False

SHEET_NAME = "Logbook_db"
GITHUB_OWNER = "Syano18" # REPLACE WITH YOUR GITHUB USERNAME
GITHUB_REPO = "Kalinga-OpsHUB"  # REPLACE WITH YOUR REPO NAME

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

GOOGLE_CREDENTIALS_FILE = resource_path("credentials.json")
FIREBASE_ADMIN_KEY = resource_path("firebase_admin_key.json")
FIREBASE_WEB_API_KEY = "AIzaSyAVTSOd5gN4QhFWjXBiQvpTIXLMzT_HSMU" 
CLIENT_SECRETS_FILE = resource_path("client_secret.json")
SCOPES = ['https://www.googleapis.com/auth/userinfo.email', 'openid']
CURRENT_VERSION = "1.7"

# Save session and cache in a user-writable directory
try:
    # Typically C:\Users\USER\AppData\Local\KalingaOpsHub
    base_dir = os.getenv("LOCALAPPDATA", os.path.expanduser("~"))
    session_dir = os.path.join(base_dir, "KalingaOpsHub")
    if not os.path.exists(session_dir):
        os.makedirs(session_dir)
    SESSION_FILE = os.path.join(session_dir, "session.txt")
    CACHE_FILE = os.path.join(session_dir, "user_cache.json")
    GOOGLE_TOKEN_FILE = os.path.join(session_dir, "token.json")
    LOGS_CACHE_FILE = os.path.join(session_dir, "logs_cache.json")
except Exception:
    SESSION_FILE = "user_session.txt"
    CACHE_FILE = "user_cache.json"
    GOOGLE_TOKEN_FILE = "token.json"
    LOGS_CACHE_FILE = "logs_cache.json"

def get_cache_data(email):
    """Retrieves cached user data if valid (less than 10 days old)."""
    try:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, "r") as f:
                data = json.load(f)
                if data.get("email") == email:
                    # Check 10 days expiration (864000 seconds)
                    if time.time() - data.get("timestamp", 0) < 864000:
                        return data
    except: pass
    return None

def save_cache_data(email, password, user_data):
    """Saves user data and password hash to local cache."""
    try:
        data = {
            "email": email,
            "password_hash": hashlib.sha256(password.encode()).hexdigest(),
            "user_data": user_data,
            "timestamp": time.time()
        }
        with open(CACHE_FILE, "w") as f:
            json.dump(data, f)
    except: pass

LOGO_PATH = resource_path("assets/Logo.png")
Agency_Logo = resource_path("assets/PSA.png")

DEFAULT_EMAIL = ""
DEFAULT_PASSWORD = ""

CLR_PRIMARY = "#0066cc"      
CLR_PRIMARY_HOVER = "#0052a3"
CLR_BG = "#f8f9fa"           
CLR_SIDEBAR = "#202124"      
CLR_CARD = "#ffffff"         
CLR_TEXT = "#3c4043"         
CLR_SUCCESS = "#28a745"

if not firebase_admin._apps:
    try:
        cred = credentials.Certificate(FIREBASE_ADMIN_KEY)
        firebase_admin.initialize_app(cred)
    except Exception as e:
        print(f"Firebase Admin Error: {e}")

def load_logo(path, size=(100, 100)):
    """Safely loads and resizes an image."""
    try:
        if os.path.exists(path):
            img = Image.open(path)
            img = img.resize(size, Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
    except Exception as e:
        print(f"Logo Load Error: {e}")
    return None

def create_modern_button(parent, text, command, bg=CLR_PRIMARY, fg="white", width=None):
    btn = tk.Button(parent, text=text, command=command, bg=bg, fg=fg, 
                    font=("Segoe UI", 10, "bold"), relief=tk.FLAT, 
                    cursor="hand2", pady=8, padx=15, activebackground=CLR_PRIMARY_HOVER)
    btn.bind("<Enter>", lambda e: btn.config(bg=CLR_PRIMARY_HOVER))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg))
    if width: btn.config(width=width)
    return btn

def create_labeled_entry(parent, label_text, var=None, show=None):
    container = tk.Frame(parent, bg=CLR_CARD)
    container.pack(fill="x", pady=8)
    tk.Label(container, text=label_text, font=("Segoe UI", 9, "bold"), 
             bg=CLR_CARD, fg=CLR_TEXT).pack(anchor="w")
    border = tk.Frame(container, bg="#dadce0", padx=1, pady=1)
    border.pack(fill="x", pady=(4, 0))
    entry = tk.Entry(border, textvariable=var, show=show, font=("Segoe UI", 11), 
                     bg="white", relief=tk.FLAT, bd=0)
    entry.pack(fill="x", ipady=8, padx=5)
    return entry

def create_multi_select_dropdown(parent, options, result_var):
    container = tk.Frame(parent, bg="#dadce0", padx=1, pady=1)
    inner = tk.Frame(container, bg="white")
    inner.pack(fill="x")
    
    display_var = tk.StringVar()
    def update_display(*args):
        val = result_var.get()
        display_var.set(val if val else "Select Mode(s)")
    result_var.trace("w", update_display)
    update_display()
    
    lbl = tk.Label(inner, textvariable=display_var, bg="white", anchor="w", 
                   font=("Segoe UI", 11), cursor="hand2", relief="flat")
    lbl.pack(fill="x", ipady=5, padx=5)
    tk.Label(inner, text="▼", bg="white", fg="#5f6368", font=("Segoe UI", 8)).place(relx=1.0, rely=0.5, anchor="e", x=-5)
    
    def open_popup(event=None):
        top = tk.Toplevel(parent)
        top.title("Select Mode(s)")
        top.geometry("250x250")
        top.configure(bg="white")
        top.transient(parent)
        top.grab_set()
        try: top.geometry(f"+{inner.winfo_rootx()}+{inner.winfo_rooty() + inner.winfo_height()}")
        except: pass
        
        current_val = result_var.get()
        current_selection = [x.strip() for x in current_val.split(',')] if current_val else []
        vars_dict = {}
        
        def on_change():
            selected = [opt for opt in options if vars_dict[opt].get()]
            result_var.set(", ".join(selected))

        for opt in options:
            var = tk.BooleanVar(value=(opt in current_selection))
            vars_dict[opt] = var
            tk.Checkbutton(top, text=opt, variable=var, command=on_change, bg="white", anchor="w", font=("Segoe UI", 10)).pack(fill="x", padx=10, pady=2)
            
        tk.Button(top, text="Done", command=top.destroy, bg=CLR_PRIMARY, fg="white", relief="flat").pack(fill="x", padx=10, pady=10, side="bottom")

    lbl.bind("<Button-1>", open_popup)
    return container

def check_internet():
    """Checks for active internet connection."""
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except OSError:
        return False

class ManageUsersView:
    def __init__(self, parent_frame, google_sheet_connection, web_api_key):
        self.frame = parent_frame
        self.worksheet = google_sheet_connection
        self.api_key = web_api_key
        self.search_var = tk.StringVar()
        self.create_widgets()
        self.search_var.trace("w", self.refresh_table)
        self.refresh_table()

    def create_widgets(self):
        header = tk.Frame(self.frame, bg=CLR_CARD, pady=20, padx=30)
        header.pack(fill="x", side="top")
        tk.Label(header, text="User Access Control", font=("Segoe UI", 18, "bold"), 
                 bg=CLR_CARD, fg=CLR_SIDEBAR).pack(side="left")

        toolbar = tk.Frame(self.frame, bg=CLR_CARD, padx=20, pady=15)
        toolbar.pack(fill="x", side="top", padx=30, pady=(0, 20))
        create_modern_button(toolbar, "+ New Staff", self.open_invite_window, bg=CLR_SUCCESS).pack(side="left")
        
        tk.Entry(toolbar, textvariable=self.search_var, font=("Segoe UI", 10), 
                 bg="#f1f3f4", relief="flat", width=30).pack(side="right", ipady=7, padx=10)
        tk.Label(toolbar, text="Search User:", bg=CLR_CARD).pack(side="right")

        table_card = tk.Frame(self.frame, bg=CLR_CARD, padx=2, pady=2)
        table_card.pack(fill="both", expand=True, padx=30, pady=(0, 30), side="top")

        cols = ("Email", "Name", "Role", "Position", "Status")
        self.tree = ttk.Treeview(table_card, columns=cols, show="headings")
        
        self.tree.heading("Email", text="EMAIL"); self.tree.column("Email", width=220)
        self.tree.heading("Name", text="NAME"); self.tree.column("Name", width=200)
        self.tree.heading("Role", text="ROLE"); self.tree.column("Role", width=100)
        self.tree.heading("Position", text="POSITION"); self.tree.column("Position", width=150)
        self.tree.heading("Status", text="STATUS"); self.tree.column("Status", width=100)
        
        vsb = ttk.Scrollbar(table_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        
        self.menu = tk.Menu(self.frame, tearoff=0)
        self.menu.add_command(label="Edit Info", command=self.edit_info)
        self.menu.add_command(label="Modify Role", command=self.edit_role)
        self.menu.add_command(label="Reset Password", command=self.reset_password)
        self.menu.add_command(label="Toggle Status", command=self.toggle_status)
        self.menu.add_command(label="Remove User", command=self.delete_user, foreground="red")
        self.tree.bind("<Button-3>", lambda e: self.menu.post(e.x_root, e.y_root) if self.tree.identify_row(e.y) else None)

    def refresh_table(self, *args):
        for item in self.tree.get_children(): self.tree.delete(item)
        self.tree.insert("", "end", values=("Loading...", "", ""))
        threading.Thread(target=self._load_rows_thread, daemon=True).start()

    def _load_rows_thread(self):
        error_message = None
        try:
            if not self.worksheet:
                error_message = "Tab 'User_Permissions' not found."
                rows = []
            else:
                rows = self.worksheet.get_all_values()
        except Exception as e:
            error_message = f"Error: {e}"
            rows = []
        
        def update_ui():
            for item in self.tree.get_children(): 
                self.tree.delete(item)
            
            if error_message:
                self.tree.insert("", "end", values=(error_message, "", "", "", ""))
                return

            q = self.search_var.get().lower()
            for row in rows[1:]:
                if not row: 
                    continue
                email = row[0] if len(row) > 0 else ""
                role = row[1] if len(row) > 1 else ""
                
                fname = row[2] if len(row) > 2 else ""
                mi = row[3] if len(row) > 3 else ""
                lname = row[4] if len(row) > 4 else ""
                name = f"{fname} {mi} {lname}".strip()
                pos = row[5] if len(row) > 5 else ""
                status = row[8] if len(row) > 8 else "Active"
                
                if q and q not in email.lower() and q not in name.lower():
                    continue
                self.tree.insert("", "end", values=(email, name, role, pos, status))
        
        try:
            self.frame.after(0, update_ui)
        except Exception:
            pass

    def open_invite_window(self):
        win = tk.Toplevel(self.frame.winfo_toplevel())
        win.title("Enrol New Staff")
        win.geometry("450x500")
        win.configure(bg=CLR_CARD, padx=30, pady=20)

        # center on parent
        win.update_idletasks()
        root = self.frame.winfo_toplevel()
        rx = root.winfo_x(); ry = root.winfo_y()
        rw = root.winfo_width(); rh = root.winfo_height()
        ww = 450; wh = 500
        win.geometry(f"{ww}x{wh}+{rx + (rw-ww)//2}+{ry + (rh-wh)//2}")

        # Make modal
        win.transient(root)
        win.grab_set()

        # --- Field Variables ---
        email_var = tk.StringVar()
        fname_var = tk.StringVar()
        mi_var = tk.StringVar()
        lname_var = tk.StringVar()
        pos_var = tk.StringVar()
        sal_var = tk.StringVar()
        sg_var = tk.StringVar()
        role_var = tk.StringVar(value="Staff")
        ht_role_var = tk.StringVar(value="User")

        # --- Layout ---

        # Email (Full width)
        create_labeled_entry(win, "Email Address", email_var)

        # Name fields (First, Middle, Last)
        name_frame = tk.Frame(win, bg=CLR_CARD)
        name_frame.pack(fill="x", pady=8)

        fname_container = tk.Frame(name_frame, bg=CLR_CARD)
        fname_container.pack(side="left", expand=True, fill="x", padx=(0, 4))
        tk.Label(fname_container, text="First Name", font=("Segoe UI", 9, "bold"), bg=CLR_CARD, fg=CLR_TEXT).pack(anchor="w")
        fname_border = tk.Frame(fname_container, bg="#dadce0", padx=1, pady=1)
        fname_border.pack(fill="x", pady=(4, 0))
        tk.Entry(fname_border, textvariable=fname_var, font=("Segoe UI", 11), bg="white", relief=tk.FLAT, bd=0).pack(fill="x", ipady=8, padx=5)
        
        mi_container = tk.Frame(name_frame, bg=CLR_CARD)
        mi_container.pack(side="left", fill="x", padx=4)
        tk.Label(mi_container, text="M.I.", font=("Segoe UI", 9, "bold"), bg=CLR_CARD, fg=CLR_TEXT).pack(anchor="w")
        mi_border = tk.Frame(mi_container, bg="#dadce0", padx=1, pady=1)
        mi_border.pack(fill="x", pady=(4, 0))
        tk.Entry(mi_border, textvariable=mi_var, font=("Segoe UI", 11), width=5, bg="white", relief=tk.FLAT, bd=0).pack(fill="x", ipady=8, padx=5)

        lname_container = tk.Frame(name_frame, bg=CLR_CARD)
        lname_container.pack(side="left", expand=True, fill="x", padx=(4, 0))
        tk.Label(lname_container, text="Last Name", font=("Segoe UI", 9, "bold"), bg=CLR_CARD, fg=CLR_TEXT).pack(anchor="w")
        lname_border = tk.Frame(lname_container, bg="#dadce0", padx=1, pady=1)
        lname_border.pack(fill="x", pady=(4, 0))
        tk.Entry(lname_border, textvariable=lname_var, font=("Segoe UI", 11), bg="white", relief=tk.FLAT, bd=0).pack(fill="x", ipady=8, padx=5)

        # Position (Full width)
        create_labeled_entry(win, "Position", pos_var)

        # Salary and SG
        salary_frame = tk.Frame(win, bg=CLR_CARD)
        salary_frame.pack(fill="x", pady=8)

        sal_container = tk.Frame(salary_frame, bg=CLR_CARD)
        sal_container.pack(side="left", expand=True, fill="x", padx=(0, 8))
        tk.Label(sal_container, text="Salary", font=("Segoe UI", 9, "bold"), bg=CLR_CARD, fg=CLR_TEXT).pack(anchor="w")
        sal_border = tk.Frame(sal_container, bg="#dadce0", padx=1, pady=1)
        sal_border.pack(fill="x", pady=(4, 0))
        tk.Entry(sal_border, textvariable=sal_var, font=("Segoe UI", 11), bg="white", relief=tk.FLAT, bd=0).pack(fill="x", ipady=8, padx=5)

        sg_container = tk.Frame(salary_frame, bg=CLR_CARD)
        sg_container.pack(side="left", expand=True, fill="x", padx=(8, 0))
        tk.Label(sg_container, text="Salary Grade", font=("Segoe UI", 9, "bold"), bg=CLR_CARD, fg=CLR_TEXT).pack(anchor="w")
        sg_border = tk.Frame(sg_container, bg="#dadce0", padx=1, pady=1)
        sg_border.pack(fill="x", pady=(4, 0))
        tk.Entry(sg_border, textvariable=sg_var, font=("Segoe UI", 11), bg="white", relief=tk.FLAT, bd=0).pack(fill="x", ipady=8, padx=5)

        # Roles (OpsHub Role, HireTrack Role)
        roles_frame = tk.Frame(win, bg=CLR_CARD)
        roles_frame.pack(fill="x", pady=8)

        role_container = tk.Frame(roles_frame, bg=CLR_CARD)
        role_container.pack(side="left", expand=True, fill="x", padx=(0, 8))
        tk.Label(role_container, text="Role", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w")
        ttk.Combobox(role_container, textvariable=role_var, values=["Staff", "Admin"], state="readonly", font=("Segoe UI", 11)).pack(fill="x", pady=5)

        ht_role_container = tk.Frame(roles_frame, bg=CLR_CARD)
        ht_role_container.pack(side="left", expand=True, fill="x", padx=(8, 0))
        tk.Label(ht_role_container, text="HireTrack Role", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w")
        ttk.Combobox(ht_role_container, textvariable=ht_role_var, values=["Admin", "PACD", "Focal Person", "User"], state="readonly", font=("Segoe UI", 11)).pack(fill="x", pady=5)

        btn = create_modern_button(win, "Send Invitation", None)
        btn.pack(fill="x", pady=20)

        def submit():
            # Collect values
            email = email_var.get().strip()
            role = role_var.get()
            ht_role = ht_role_var.get()
            fname = fname_var.get().strip()
            mi = mi_var.get().strip()
            lname = lname_var.get().strip()
            pos = pos_var.get().strip()
            sal = sal_var.get().strip()
            sg = sg_var.get().strip()

            if not all([email, role, fname, mi, lname]):
                return messagebox.showwarning("Error", "Email, First Name, Middle Initial, Last Name, and Role are required")
            
            if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
                return messagebox.showwarning("Error", "Invalid Email Address format")
            
            if not check_internet():
                return messagebox.showerror("Network Error", "No internet connection detected.")
            
            win.protocol("WM_DELETE_WINDOW", lambda: None)
            
            user_data = [email, role, fname, mi, lname, pos, sal, sg, "Active", ht_role]

            btn.config(text="Sending...", state="disabled")
            threading.Thread(target=self._invite_thread, args=(user_data, win, btn)).start()

        btn.config(command=submit)

    def _invite_thread(self, user_data, win, btn):
        email = user_data[0]
        try:
            # Create user (Firebase)
            temp_pass = ''.join(random.choices(string.ascii_letters + string.digits, k=12))
            auth.create_user(email=email, password=temp_pass)
            
            # Add to sheet
            self.worksheet.append_row(user_data)
            
            # Send reset email
            url = f"https://identitytoolkit.googleapis.com/v1/accounts:sendOobCode?key={self.api_key}"
            requests.post(url, json={"requestType": "PASSWORD_RESET", "email": email})
            
            self.frame.after(0, lambda: self._invite_success(win, email))
        except Exception as e:
            self.frame.after(0, lambda: self._invite_error(e, btn, win))

    def _invite_success(self, win, email):
        win.grab_release()
        messagebox.showinfo("Success", f"User {email} invited. Please inform them to check their email for the password setup link.")
        win.destroy()
        self.refresh_table()

    def _invite_error(self, e, btn, win):
        win.grab_release()
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        messagebox.showerror("Error", f"Failed to invite: {e}")
        btn.config(text="Send Invitation", state="normal")

    def edit_info(self):
        sel = self.tree.selection()
        if not sel: return
        val = self.tree.item(sel[0])['values']
        email = val[0]

        try:
            cell = self.worksheet.find(email)
            row_vals = self.worksheet.row_values(cell.row)
            if len(row_vals) < 8: row_vals += [""] * (8 - len(row_vals))
        except Exception as e:
            messagebox.showerror("Error", f"Could not load user details: {e}")
            return

        win = tk.Toplevel(self.frame.winfo_toplevel())
        win.title(f"Edit Info: {email}")
        win.geometry("450x650")
        win.configure(bg=CLR_CARD, padx=30, pady=20)
        win.transient(self.frame.winfo_toplevel()); win.grab_set()
        
        root = self.frame.winfo_toplevel()
        win.geometry(f"+{root.winfo_x() + (root.winfo_width()-450)//2}+{root.winfo_y() + (root.winfo_height()-650)//2}")

        vars_map = {
            "First Name": tk.StringVar(value=row_vals[2]),
            "Middle Initial": tk.StringVar(value=row_vals[3]),
            "Last Name": tk.StringVar(value=row_vals[4]),
            "Position": tk.StringVar(value=row_vals[5]),
            "Salary": tk.StringVar(value=row_vals[6]),
            "Salary Grade": tk.StringVar(value=row_vals[7])
        }
        
        for lbl, var in vars_map.items():
            create_labeled_entry(win, lbl, var)

        def save():
            if not check_internet(): return messagebox.showerror("Error", "No internet.")
            btn.config(text="Saving...", state="disabled")
            
            def _t():
                try:
                    self.worksheet.update_cell(cell.row, 3, vars_map["First Name"].get().strip())
                    self.worksheet.update_cell(cell.row, 4, vars_map["Middle Initial"].get().strip())
                    self.worksheet.update_cell(cell.row, 5, vars_map["Last Name"].get().strip())
                    self.worksheet.update_cell(cell.row, 6, vars_map["Position"].get().strip())
                    self.worksheet.update_cell(cell.row, 7, vars_map["Salary"].get().strip())
                    self.worksheet.update_cell(cell.row, 8, vars_map["Salary Grade"].get().strip())
                    
                    self.frame.after(0, lambda: (messagebox.showinfo("Success", "User info updated."), win.destroy(), self.refresh_table()))
                except Exception as e:
                    self.frame.after(0, lambda: (messagebox.showerror("Error", str(e)), btn.config(text="Save Changes", state="normal")))
            
            threading.Thread(target=_t).start()

        btn = create_modern_button(win, "Save Changes", save)
        btn.pack(fill="x", pady=20)

    def reset_password(self):
        sel = self.tree.selection()
        if not sel: return
        email = self.tree.item(sel[0])['values'][0]
        
        if messagebox.askyesno("Confirm Reset", f"Send password reset email to {email}?"):
            threading.Thread(target=self._send_reset, args=(email,)).start()

    def _send_reset(self, email):
        try:
            url = f"https://identitytoolkit.googleapis.com/v1/accounts:sendOobCode?key={self.api_key}"
            requests.post(url, json={"requestType": "PASSWORD_RESET", "email": email})
            self.frame.after(0, lambda: messagebox.showinfo("Success", f"Reset link sent to {email}"))
        except Exception as e:
            self.frame.after(0, lambda: messagebox.showerror("Error", f"Failed to send reset link: {e}"))

    def edit_role(self):
        sel = self.tree.selection()
        if not sel: return
        val = self.tree.item(sel[0])['values']
        email = val[0]

        try:
            cell = self.worksheet.find(email)
            row_vals = self.worksheet.row_values(cell.row)
            if len(row_vals) < 10: row_vals += [""] * (10 - len(row_vals))
            current_role = row_vals[1]
            current_ht_role = row_vals[9]
        except Exception as e:
            messagebox.showerror("Error", f"Could not load user details: {e}")
            return

        win = tk.Toplevel(self.frame.winfo_toplevel())
        win.title("Edit Roles")
        win.geometry("350x320")
        win.configure(bg=CLR_CARD, padx=20, pady=20)
        
        win.update_idletasks()
        root = self.frame.winfo_toplevel()
        x = root.winfo_x() + (root.winfo_width() - 350) // 2
        y = root.winfo_y() + (root.winfo_height() - 320) // 2
        win.geometry(f"+{x}+{y}")
        win.transient(root); win.grab_set()

        tk.Label(win, text=f"Update Roles for {email}", font=("Segoe UI", 10), bg=CLR_CARD, wraplength=300).pack(pady=(0, 15))
        
        tk.Label(win, text="OpsHub Role", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w")
        role_var = tk.StringVar(value=current_role if current_role in ["Staff", "Admin"] else "Staff")
        ttk.Combobox(win, textvariable=role_var, values=["Staff", "Admin"], state="readonly", font=("Segoe UI", 11)).pack(fill="x", pady=(5, 15))

        tk.Label(win, text="HireTrack Role", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w")
        ht_role_var = tk.StringVar(value=current_ht_role if current_ht_role in ["Admin", "PACD", "Focal Person", "User"] else "User")
        ttk.Combobox(win, textvariable=ht_role_var, values=["Admin", "PACD", "Focal Person", "User"], state="readonly", font=("Segoe UI", 11)).pack(fill="x", pady=(5, 15))

        def save():
            try:
                self.worksheet.update_cell(cell.row, 2, role_var.get())
                self.worksheet.update_cell(cell.row, 10, ht_role_var.get())
                messagebox.showinfo("Success", "Roles updated.")
                win.destroy()
                self.refresh_table()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update roles: {e}")

        create_modern_button(win, "Save Changes", save).pack(fill="x", pady=10)

    def toggle_status(self):
        sel = self.tree.selection()
        if not sel: return
        val = self.tree.item(sel[0])['values']
        email = val[0]
        current_status = val[4]
        
        new_status = "Inactive" if current_status == "Active" else "Active"
        
        if messagebox.askyesno("Confirm", f"Set user {email} to {new_status}?"):
            try:
                try:
                    user = auth.get_user_by_email(email)
                    auth.update_user(user.uid, disabled=(new_status == "Inactive"))
                except: pass

                cell = self.worksheet.find(email)
                self.worksheet.update_cell(cell.row, 9, new_status)
                messagebox.showinfo("Success", f"User updated to {new_status}.")
                self.refresh_table()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update status: {e}")

    def delete_user(self):
        sel = self.tree.selection()
        if not sel: return
        val = self.tree.item(sel[0])['values']
        email = val[0]
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to permanently delete user {email}?"):
            try:
                try:
                    user = auth.get_user_by_email(email)
                    auth.delete_user(user.uid)
                except: pass
                
                cell = self.worksheet.find(email)
                self.worksheet.delete_rows(cell.row)
                messagebox.showinfo("Success", "User deleted.")
                self.refresh_table()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete user: {e}")

class LoginWindow:
    def __init__(self, root, on_success):
        self.root = root
        self.on_success = on_success
        self.root.title("Kalinga OpsHUB - Login")
        
        # Ensure window is restored to normal state (not maximized)
        self.root.state('normal')
        
        width, height = 400, 600
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.resizable(False, False)
        self.root.configure(bg=CLR_CARD)
        
        container = tk.Frame(self.root, bg=CLR_CARD, padx=40)
        container.pack(expand=True, fill="both")
        
        self.info_btn = tk.Button(self.root, text="ℹ", font=("Segoe UI", 16), bg=CLR_CARD, fg=CLR_PRIMARY,
                  relief=tk.FLAT, bd=0, cursor="hand2")
        self.info_btn.place(x=350, y=10)
        self.info_btn.bind("<Enter>", self.show_about_tooltip)
        self.info_btn.bind("<Leave>", self.hide_about_tooltip)
        
        self.logo_img = load_logo(LOGO_PATH, size=(120, 120))
        if self.logo_img:
            tk.Label(container, image=self.logo_img, bg=CLR_CARD).pack(pady=(40, 5))
        else:
            tk.Label(container, text="Kalinga OpsHUB", font=("Segoe UI", 28, "bold"), 
                     bg=CLR_CARD, fg=CLR_PRIMARY).pack(pady=(50, 5))

        tk.Label(container, text="Kalinga Operations HUB", font=("Segoe UI", 10), 
                 bg=CLR_CARD, fg="#5f6368").pack(pady=(0, 40))
        
        self.email_var = tk.StringVar(value=DEFAULT_EMAIL)
        self.e_mail = create_labeled_entry(container, "Email", self.email_var)
        self.e_mail.bind('<Return>', lambda e: self.attempt_login())
        self.pass_var = tk.StringVar(value=DEFAULT_PASSWORD)
        pass_container = tk.Frame(container, bg=CLR_CARD)
        pass_container.pack(fill="x", pady=8)
        tk.Label(pass_container, text="Password", font=("Segoe UI", 9, "bold"), 
             bg=CLR_CARD, fg=CLR_TEXT).pack(anchor="w")
        
        pass_frame = tk.Frame(pass_container, bg="#dadce0", padx=1, pady=1)
        pass_frame.pack(fill="x", pady=(4, 0))
        
        inner_frame = tk.Frame(pass_frame, bg="white")
        inner_frame.pack(fill="x")
        
        self.e_pass = tk.Entry(inner_frame, textvariable=self.pass_var, font=("Segoe UI", 11), 
                       bg="white", relief=tk.FLAT, bd=0, show="●")
        self.e_pass.pack(side="left", fill="x", expand=True, ipady=8, padx=5)
        self.e_pass.bind('<Return>', lambda e: self.attempt_login())
        
        self.show_pass_btn = tk.Button(inner_frame, text="👁", bg="white", relief=tk.FLAT, 
                        bd=0, cursor="hand2", command=self.toggle_password)
        self.show_pass_btn.pack(side="right", padx=5)
        
        self.btn = create_modern_button(container, "Login", self.attempt_login)
        self.btn.pack(fill="x", pady=20)
        
        if GOOGLE_LOGIN_AVAILABLE:
            self.google_btn = tk.Button(container, text="Sign in with Google", command=self.login_with_google,
                                   bg="#db4437", fg="white", font=("Segoe UI", 10, "bold"), 
                                   relief=tk.FLAT, cursor="hand2", pady=8)
            self.google_btn.pack(fill="x", pady=(0, 20))
        
        forgot_frame = tk.Frame(container, bg=CLR_CARD)
        forgot_frame.pack(fill="x", pady=(0, 10))
        forgot_btn = tk.Button(forgot_frame, text="Forgot Password?", bg=CLR_CARD, fg=CLR_PRIMARY, 
                               font=("Segoe UI", 9, "underline"), relief=tk.FLAT, bd=0, 
                               cursor="hand2", command=self.open_forgot_password)
        forgot_btn.pack(anchor="e")
        
        tk.Label(self.root, text=f"v{CURRENT_VERSION}", font=("Segoe UI", 8), bg=CLR_CARD, fg="#9aa0a6").place(relx=0.98, rely=0.99, anchor="se")

        self.update_btn = tk.Button(self.root, text="Check for Updates", font=("Segoe UI", 8, "underline"), bg=CLR_CARD, fg=CLR_PRIMARY, 
                                    relief=tk.FLAT, bd=0, cursor="hand2", command=self.check_updates)
        self.update_btn.place(relx=0.02, rely=0.99, anchor="sw")

    def check_updates(self):
        self.update_btn.config(text="Checking...", state="disabled")
        threading.Thread(target=self._update_check_thread, daemon=True).start()

    def _update_check_thread(self):
        if not check_internet():
            self.root.after(0, lambda: (messagebox.showerror("Error", "No internet connection."), self.update_btn.config(text="Check for Updates", state="normal")))
            return
        
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                latest_ver = data.get("tag_name", "").lstrip("v").strip() # Remove 'v' prefix if present
                assets = data.get("assets", [])
                dl_link = next((a.get("browser_download_url") for a in assets if a.get("name", "").endswith(".exe")), None)
                
                if not dl_link:
                    raise Exception("No .exe file found in the latest release assets.")

                if latest_ver and latest_ver != CURRENT_VERSION:
                    self.root.after(0, lambda: (self._prompt_update(latest_ver, dl_link), self.update_btn.config(text="Check for Updates", state="normal")))
                else:
                    self.root.after(0, lambda: (messagebox.showinfo("Up to Date", "You are using the latest version."), self.update_btn.config(text="Check for Updates", state="normal")))
            else:
                raise Exception(f"GitHub API returned {response.status_code}")
        except Exception as e:
            self.root.after(0, lambda: (messagebox.showerror("Error", f"Update check failed: {e}"), self.update_btn.config(text="Check for Updates", state="normal")))

    def _prompt_update(self, ver, url):
        if messagebox.askyesno("Update Available", f"A new version ({ver}) is available.\nWould you like to download and install it now?"):
            self.download_and_install(url)

    def download_and_install(self, url):
        win = tk.Toplevel(self.root)
        win.title("Updating")
        win.geometry("350x150")
        win.configure(bg=CLR_CARD)
        win.transient(self.root); win.grab_set()
        
        win.update_idletasks()
        win.geometry(f"+{self.root.winfo_x() + (self.root.winfo_width()-350)//2}+{self.root.winfo_y() + (self.root.winfo_height()-150)//2}")
        
        tk.Label(win, text="Downloading Update...", font=("Segoe UI", 10, "bold"), bg=CLR_CARD).pack(pady=(20, 10))
        pb = ttk.Progressbar(win, orient="horizontal", length=280, mode="determinate")
        pb.pack(pady=5)
        lbl = tk.Label(win, text="Initializing...", font=("Segoe UI", 8), bg=CLR_CARD, fg="#5f6368")
        lbl.pack()

        def _d():
            try:
                dest = os.path.join(tempfile.gettempdir(), "KalingaOpsHUB_Update.exe")
                r = requests.get(url, stream=True)

                total = int(r.headers.get('content-length', 0))
                
                if total == 0:
                    self.root.after(0, lambda: (pb.configure(mode="indeterminate"), pb.start(10)))
                
                with open(dest, 'wb') as f:
                    dl = 0
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk); dl += len(chunk)
                            if total > 0:
                                perc = int(dl * 100 / total)
                                self.root.after(0, lambda p=perc: (pb.configure(value=p), lbl.configure(text=f"{p}% Complete")))
                
                if os.path.getsize(dest) < 1024:
                    raise Exception("Download failed (file too small).")

                self.root.after(0, lambda: (lbl.configure(text="Running Installer..."), subprocess.Popen([dest]), self.root.destroy(), sys.exit()))
            except Exception as e:
                self.root.after(0, lambda: (win.destroy(), messagebox.showerror("Update Error", f"Failed: {e}")))
        
        threading.Thread(target=_d).start()
    
    def toggle_password(self):
        """Toggle password visibility."""
        if self.e_pass.cget('show') == '●':
            self.e_pass.config(show='')
            self.show_pass_btn.config(text='🙈')
        else:
            self.e_pass.config(show='●')
            self.show_pass_btn.config(text='👁')
    
    def open_forgot_password(self):
        """Open forgot password dialog."""
        win = tk.Toplevel(self.root)
        win.title("Reset Password")
        win.geometry("350x350")
        win.configure(bg=CLR_CARD)
        win.resizable(False, False)
        
        self.root.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        win_width = 350
        win_height = 250
        x = root_x + (root_width - win_width) // 2
        y = root_y + (root_height - win_height) // 2
        win.geometry(f"350x350+{x}+{y}")
        
        container = tk.Frame(win, bg=CLR_CARD, padx=30, pady=30)
        container.pack(fill="both", expand=True)
        
        tk.Label(container, text="Reset Password", font=("Segoe UI", 14, "bold"), 
                 bg=CLR_CARD, fg=CLR_SIDEBAR).pack(anchor="w", pady=(0, 20))
        
        tk.Label(container, text="Enter your email address and we'll send you a password reset link.", 
                 font=("Segoe UI", 9), bg=CLR_CARD, fg=CLR_TEXT, wraplength=280, justify="left").pack(anchor="w", pady=(0, 15))
        
        email_var = tk.StringVar()
        email_entry = create_labeled_entry(container, "Email Address", email_var)
        
        def send_reset():
            email = email_var.get()
            if not email:
                messagebox.showwarning("Error", "Please enter your email address")
                return
            
            if not check_internet():
                messagebox.showerror("Network Error", "No internet connection detected.")
                return

            reset_btn.config(text="Sending...", state="disabled")
            threading.Thread(target=self._send_password_reset, args=(email, reset_btn, win)).start()
        
        reset_btn = create_modern_button(container, "Send Reset Link", send_reset, bg=CLR_SUCCESS)
        reset_btn.pack(fill="x", pady=20)
    
    def _send_password_reset(self, email, btn, win):
        """Send password reset email."""
        try:
            requests.post(f"https://identitytoolkit.googleapis.com/v1/accounts:sendOobCode?key={FIREBASE_WEB_API_KEY}", 
                          json={"requestType": "PASSWORD_RESET", "email": email})
            self.root.after(0, lambda: (messagebox.showinfo("Success", "Password reset link sent to your email"), 
                                       win.destroy()))
        except Exception as e:
            self.root.after(0, lambda: (messagebox.showerror("Error", "Failed to send reset link"), 
                                       btn.config(text="Send Reset Link", state="normal")))

    def show_about_tooltip(self, event):
        if hasattr(self, 'tw') and self.tw: return
        x = self.info_btn.winfo_rootx() - 260
        y = self.info_btn.winfo_rooty() + 30
        self.tw = tk.Toplevel(self.root)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x}+{y}")
        self.tw.configure(bg="#333333")

        msg = ("The Kalinga OpHub is a digital innovation tool designed to simplify office operations. "
               "It combines two essential functions into one platform: an Outgoing Document Logbook for tracking "
               "official correspondence with unique reference numbers, and a Leave Monitoring System for the "
               "accurate management of personnel leave credits.")
        tk.Label(self.tw, text=msg, justify='left', background="#333333", fg="white", 
                 relief='flat', borderwidth=0, font=("Segoe UI", 9), wraplength=280, padx=10).pack(anchor="w", pady=(10, 5))
        
        tk.Label(self.tw, text="Developer: Chano\nEmail: officialchano18@gmail.com", justify='left', background="#333333", fg="white", 
                 relief='flat', borderwidth=0, font=("Segoe UI", 9, "bold"), wraplength=280, padx=10).pack(anchor="w", pady=(0, 10))

    def hide_about_tooltip(self, event):
        if hasattr(self, 'tw') and self.tw: self.tw.destroy(); self.tw = None

    def login_with_google(self):
        if not check_internet():
            messagebox.showerror("Network Error", "No internet connection detected.")
            return
        
        if not os.path.exists(CLIENT_SECRETS_FILE):
            messagebox.showerror("Configuration Error", "client_secret.json not found.\nPlease add the OAuth2 client secret file.")
            return

        self.btn.config(state="disabled")
        if hasattr(self, 'google_btn'): self.google_btn.config(state="disabled")
        threading.Thread(target=self._google_auth_flow, daemon=True).start()

    def _google_auth_flow(self):
        creds = None
        if os.path.exists(GOOGLE_TOKEN_FILE):
            try: creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_FILE, SCOPES)
            except: creds = None
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try: creds.refresh(Request())
                except: creds = None
            
            if not creds:
                try:
                    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
                    creds = flow.run_local_server(port=0, access_type='offline', prompt='consent')
                except Exception as e:
                    self.root.after(0, lambda: self._reset_login_ui(f"Google Sign-In Error: {e}"))
                    return
            
            try:
                with open(GOOGLE_TOKEN_FILE, 'w') as token: token.write(creds.to_json())
            except: pass

        try:
            session = requests.Session()
            session.headers.update({'Authorization': f'Bearer {creds.token}'})
            user_info = session.get('https://www.googleapis.com/oauth2/v1/userinfo').json()
            email = user_info.get('email')
            
            if email:
                self.root.after(0, lambda: self._process_google_login(email))
            else:
                raise Exception("Email not found in Google account info.")
        except Exception as e:
            self.root.after(0, lambda: self._reset_login_ui(f"Google Sign-In Error: {e}"))

    def _process_google_login(self, email):
        def check_gs():
            try:
                gc = gspread.service_account(filename=GOOGLE_CREDENTIALS_FILE)
                sh = gc.open(SHEET_NAME)
                ws = sh.worksheet("User_Permissions")
                all_users = ws.get_all_records()
                user_data = next((user for user in all_users if user.get('Email') and user.get('Email').lower() == email.lower()), None)
                self.root.after(0, lambda: self._finish_google_login(email, user_data))
            except Exception as e:
                self.root.after(0, lambda: self._reset_login_ui(f"Database Error: {e}"))
        threading.Thread(target=check_gs, daemon=True).start()

    def _finish_google_login(self, email, user_data):
        if not user_data:
            self._reset_login_ui("This Google account is not registered in the system.")
            return
            
        if user_data.get('Status') == "Inactive":
            self._reset_login_ui("Account is Inactive.\nPlease contact your administrator.")
            return

        save_cache_data(email, "GOOGLE_AUTH_TOKEN", user_data)
        try:
            with open(SESSION_FILE, "w") as f: f.write(email)
        except: pass
        
        self.on_success(email, user_data)

    def _reset_login_ui(self, error_msg=None):
        if error_msg: messagebox.showerror("Login Error", error_msg)
        self.btn.config(state="normal")
        if hasattr(self, 'google_btn'): self.google_btn.config(state="normal")

    def attempt_login(self):
        if not check_internet():
            messagebox.showerror("Network Error", "No internet connection detected.\nPlease check your connection.")
            return

        self.btn.config(text="Authenticating...", state="disabled")
        em = self.email_var.get()
        pw = self.pass_var.get()
        threading.Thread(target=self._auth, args=(em, pw)).start()

    def _auth(self, em, pw):
        # Check local cache first to speed up login
        cached = get_cache_data(em)
        if cached:
            ph = hashlib.sha256(pw.encode()).hexdigest()
            if cached.get("password_hash") == ph:
                try:
                    with open(SESSION_FILE, "w") as f: f.write(em)
                except: pass
                self.root.after(0, lambda: self.on_success(em, cached["user_data"]))
                return

        # Start Google Sheets check in parallel to speed up login
        gs_result = {"user_data": None}
        def check_gs_status():
            try:
                gc = gspread.service_account(filename=GOOGLE_CREDENTIALS_FILE)
                sh = gc.open(SHEET_NAME)
                ws = sh.worksheet("User_Permissions")
                all_users = ws.get_all_records()
                gs_result["user_data"] = next((user for user in all_users if user.get('Email') and user.get('Email').lower() == em.lower()), None)
            except Exception as e:
                print(f"Status check warning: {e}")

        gs_thread = threading.Thread(target=check_gs_status)
        gs_thread.start()

        url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={FIREBASE_WEB_API_KEY}"
        try:
            r = requests.post(url, json={"email": em, "password": pw, "returnSecureToken": True})
            
            # Wait for GS thread to finish
            gs_thread.join()

            if r.status_code == 200:
                user_data = gs_result["user_data"]
                
                save_cache_data(em, pw, user_data)

                if user_data and user_data.get('Status') == "Inactive":
                    self.root.after(0, lambda: self._fail("Account is Inactive.\nPlease contact your administrator."))
                    return

                try:
                    with open(SESSION_FILE, "w") as f: f.write(em)
                except: pass
                
                self.root.after(0, lambda: self.on_success(em, user_data))
            else:
                msg = "Invalid login details"
                try:
                    err = r.json().get('error', {}).get('message', '')
                    if err == "USER_DISABLED": msg = "Account is disabled."
                except: pass
                self.root.after(0, lambda: self._fail(msg))
        except Exception as e: self.root.after(0, lambda: self._fail(f"Connection Error: {e}"))

    def _fail(self, msg="Invalid login details"):
        messagebox.showerror("Error", msg)
        self.btn.config(text="Login", state="normal")

class LogbookApp:
    def __init__(self, root, email, user_data=None, on_success_callback=None):
        self.root = root
        self.user_data = user_data
        self.email = email
        self.role = "Staff"
        
        if self.user_data:
            role_val = self.user_data.get('Role')
            self.role = str(role_val).strip() if role_val else "Staff"
            
        self.master_logs = []
        self.headers = []
        self.fetching_logs = False
        self.data_loaded_from_net = False
        self.connecting = True
        self.user_names_cache = []
        self.current_user_name = ""
        self.on_success_callback = on_success_callback
        self.root.title("Kalinga OpsHUB")

        self.ws_users = None
        self.ws_logs = None
        self.gc = None
        self.sh = None

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=35, font=("Segoe UI", 10), borderwidth=0)
        style.configure("Treeview.Heading", background="#f1f3f4", font=("Segoe UI", 10, "bold"), relief="flat")
        # Configure larger font for Combobox dropdown
        style.configure("TCombobox", fieldbackground="white", background="white", font=("Segoe UI", 13))
        style.map("TCombobox", fieldbackground=[("readonly", "white")], background=[("readonly", "white")])
        
        width, height = 1200, 800
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.resizable(True, True)
        
        self.root.configure(bg=CLR_BG)
        self.setup_layout()
        self.show_page("logs")
        self.monitor_connection()
        self.load_local_logs_cache()
        
        # Start connection in background to prevent startup freeze
        threading.Thread(target=self.connect_background, daemon=True).start()

    def connect_background(self):
        if not check_internet():
            self.connecting = False
            self.root.after(0, self.draw_logs)
            return

        try:
            self.gc = gspread.service_account(filename=GOOGLE_CREDENTIALS_FILE)
            self.sh = self.gc.open(SHEET_NAME)
            self.ws_logs = self.sh.sheet1
            
            role_updated = False
            try:
                self.ws_users = self.sh.worksheet("User_Permissions")
                
                if not self.user_data:
                    all_records = self.ws_users.get_all_records()
                    u = next((x for x in all_records if x.get('Email') and x.get('Email').lower() == self.email.lower()), None)
                    if u:
                        self.user_data = u
                        r = str(u.get('Role')).strip()
                        if r != self.role:
                            self.role = r
                            role_updated = True
            except: pass
            
            self.root.after(0, lambda: self.on_connected(role_updated))
        except Exception as e:
            print(f"Connection Error: {e}")
            self.connecting = False
            self.root.after(0, self.draw_logs)

    def on_connected(self, role_updated=False):
        self.connecting = False
        if role_updated:
            self.build_sidebar()
        
        if not self.fetching_logs and not self.data_loaded_from_net:
            self.fetching_logs = True
            threading.Thread(target=self._fetch_logs_thread, daemon=True).start()
        
        threading.Thread(target=self.fetch_user_names, daemon=True).start()
    
    def load_local_logs_cache(self):
        try:
            if os.path.exists(LOGS_CACHE_FILE):
                with open(LOGS_CACHE_FILE, 'r') as f:
                    data = json.load(f)
                    self.headers = data.get("headers", [])
                    self.master_logs = data.get("rows", [])
        except: pass

    def fetch_user_names(self):
        """Fetches user names for dropdowns in background."""
        try:
            if not self.ws_users: return
            rows = self.ws_users.get_all_values()
            names = []
            for r in rows[1:]:
                if len(r) > 4:
                    fname = r[2].strip()
                    mi = r[3].strip()
                    if mi: mi = f"{mi[0]}."
                    lname = r[4].strip()
                    full = f"{fname} {mi} {lname}".strip()
                    full = " ".join(full.split())
                    if full: names.append(full)
                    
                    if r[0].strip().lower() == self.email.lower():
                        self.current_user_name = full
            self.user_names_cache = sorted(names)
        except Exception as e:
            print(f"User fetch error: {e}")

    def logout(self):
        """Logout and return to login screen."""
        if os.path.exists(SESSION_FILE):
            try: os.remove(SESSION_FILE)
            except: pass
        for w in self.root.winfo_children():
            w.destroy()
        LoginWindow(self.root, self.on_success_callback)

    def monitor_connection(self):
        """Continuously checks internet connection and updates sidebar status."""
        def check():
            is_online = check_internet()
            self.root.after(0, lambda: self._update_status(is_online))
        threading.Thread(target=check, daemon=True).start()
        self.root.after(5000, self.monitor_connection)

    def _update_status(self, is_online):
        try:
            if not self.status_dot.winfo_exists(): return
            if is_online:
                self.status_dot.config(fg=CLR_SUCCESS)
                self.status_text.config(text="SYSTEM ONLINE", fg="#9aa0a6")
            else:
                self.status_dot.config(fg="#d32f2f") # Red color
                self.status_text.config(text="SYSTEM OFFLINE", fg="#d32f2f")
        except: pass

    def setup_layout(self):
        self.sidebar = tk.Frame(self.root, bg=CLR_SIDEBAR, width=220)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        
        self.build_sidebar()

        self.content = tk.Frame(self.root, bg=CLR_BG)
        self.content.pack(side="right", fill="both", expand=True)
        self.p_logs = tk.Frame(self.content, bg=CLR_BG)
        self.p_users = tk.Frame(self.content, bg=CLR_BG)
        self.users_view = None

    def build_sidebar(self):
        for w in self.sidebar.winfo_children(): w.destroy()

        self.side_logo = load_logo(LOGO_PATH, size=(100, 100))
        if self.side_logo:
            tk.Label(self.sidebar, image=self.side_logo, bg=CLR_SIDEBAR).pack(pady=(20, 30))
        
        self.nav_btns = {}
        def add_nav(id, txt, cmd):
            b = tk.Button(self.sidebar, text=f"  {txt}", font=("Segoe UI", 10), anchor="w",
                          bg=CLR_SIDEBAR, fg="#9aa0a6", relief="flat", bd=0, pady=12, command=cmd)
            b.pack(fill="x", padx=10); self.nav_btns[id] = b

        add_nav("logs", "Digital Logbook", lambda: self.show_page("logs"))
        if self.role == "Admin":
            add_nav("users", "User Management", lambda: self.show_page("users"))

        bottom_frame = tk.Frame(self.sidebar, bg=CLR_SIDEBAR)
        bottom_frame.pack(side="bottom", fill="x")
        
        tk.Label(bottom_frame, text="Logged in as:", font=("Segoe UI", 8), bg=CLR_SIDEBAR, fg="#9aa0a6").pack(anchor="w", padx=20, pady=(10, 0))
        tk.Label(bottom_frame, text=self.email, font=("Segoe UI", 9, "bold"), bg=CLR_SIDEBAR, fg="white", wraplength=180, justify="left").pack(anchor="w", padx=20, pady=(0, 10))
        
        logout_btn = create_modern_button(bottom_frame, "🚪 Logout", self.logout, bg="#d32f2f")
        logout_btn.pack(fill="x", padx=10, pady=(10, 15))
        
        status_frame = tk.Frame(bottom_frame, bg=CLR_SIDEBAR)
        status_frame.pack(fill="x", pady=(0, 15), padx=20)
        self.status_dot = tk.Label(status_frame, text="●", fg=CLR_SUCCESS, bg=CLR_SIDEBAR)
        self.status_dot.pack(side="left")
        self.status_text = tk.Label(status_frame, text="SYSTEM ONLINE", font=("Segoe UI", 8), bg=CLR_SIDEBAR, fg="#9aa0a6")
        self.status_text.pack(side="left", padx=5)

    def show_page(self, pid):
        if pid == "users" and self.role != "Admin":
            messagebox.showerror("Access Denied", "Restricted to Administrators only.")
            return

        for p in [self.p_logs, self.p_users]: p.pack_forget()
        for b in self.nav_btns.values(): b.config(bg=CLR_SIDEBAR, fg="#9aa0a6")
        if pid in self.nav_btns: self.nav_btns[pid].config(bg=CLR_PRIMARY, fg="white")
        
        if pid == "logs": self.p_logs.pack(fill="both", expand=True); self.draw_logs()
        elif pid == "users":
            self.p_users.pack(fill="both", expand=True)
            if not self.users_view:
                self.users_view = ManageUsersView(self.p_users, self.ws_users, FIREBASE_WEB_API_KEY)

    def open_profile_settings(self):
        if not self.ws_users:
            return messagebox.showerror("Error", "User database not connected.")
        
        try:
            cell = self.ws_users.find(self.email)
            row_vals = self.ws_users.row_values(cell.row)
            if len(row_vals) < 8: row_vals += [""] * (8 - len(row_vals))
        except Exception as e:
            return messagebox.showerror("Error", f"Could not load profile: {e}")

        win = tk.Toplevel(self.root)
        win.title("Edit Profile")
        win.geometry("400x750")
        win.configure(bg=CLR_CARD, padx=30, pady=30)
        win.transient(self.root); win.grab_set()
        self.root.update_idletasks()
        win.geometry(f"+{self.root.winfo_x() + (self.root.winfo_width()-400)//2}+{self.root.winfo_y() + (self.root.winfo_height()-750)//2}")

        tk.Label(win, text="My Profile", font=("Segoe UI", 16, "bold"), bg=CLR_CARD, fg=CLR_SIDEBAR).pack(anchor="w", pady=(0,5))
        tk.Label(win, text=self.email, font=("Segoe UI", 10), bg=CLR_CARD, fg="#5f6368").pack(anchor="w", pady=(0,20))

        vars_map = {
            "First Name": tk.StringVar(value=row_vals[2]),
            "Middle Initial": tk.StringVar(value=row_vals[3]),
            "Last Name": tk.StringVar(value=row_vals[4]),
            "Position": tk.StringVar(value=row_vals[5]),
            "Salary": tk.StringVar(value=row_vals[6]),
            "Salary Grade": tk.StringVar(value=row_vals[7])
        }
        for lbl, var in vars_map.items(): create_labeled_entry(win, lbl, var)

        def save():
            if not check_internet(): return messagebox.showerror("Error", "No internet.")
            btn.config(text="Saving...", state="disabled")
            def _t():
                try:
                    self.ws_users.update_cell(cell.row, 3, vars_map["First Name"].get().strip())
                    self.ws_users.update_cell(cell.row, 4, vars_map["Middle Initial"].get().strip())
                    self.ws_users.update_cell(cell.row, 5, vars_map["Last Name"].get().strip())
                    self.ws_users.update_cell(cell.row, 6, vars_map["Position"].get().strip())
                    self.ws_users.update_cell(cell.row, 7, vars_map["Salary"].get().strip())
                    self.ws_users.update_cell(cell.row, 8, vars_map["Salary Grade"].get().strip())
                    self.root.after(0, lambda: (messagebox.showinfo("Success", "Profile Updated"), win.destroy()))
                except Exception as e:
                    self.root.after(0, lambda: (messagebox.showerror("Error", str(e)), btn.config(text="Save Changes", state="normal")))
            threading.Thread(target=_t).start()

        btn = create_modern_button(win, "Save Changes", save)
        btn.pack(fill="x", pady=20)

    def export_logs(self):
        if not self.master_logs:
            return messagebox.showwarning("Export", "No data available to export.")

        # Get available years
        years = set()
        for r in self.master_logs:
            if r and len(r) > 0:
                try: years.add(str(datetime.strptime(r[0], "%Y-%m-%d %H:%M").year))
                except: pass
        sorted_years = sorted(list(years), reverse=True)
        if not sorted_years: sorted_years = [datetime.now().strftime("%Y")]

        win = tk.Toplevel(self.root)
        win.title("Export Options")
        win.geometry("350x350")
        win.configure(bg=CLR_CARD, padx=20, pady=20)
        win.transient(self.root); win.grab_set()
        
        self.root.update_idletasks()
        win.geometry(f"+{self.root.winfo_x() + (self.root.winfo_width()-350)//2}+{self.root.winfo_y() + (self.root.winfo_height()-350)//2}")

        tk.Label(win, text="Export Data", font=("Segoe UI", 12, "bold"), bg=CLR_CARD).pack(anchor="w", pady=(0, 15))

        mode_var = tk.StringVar(value="Current View")
        year_var = tk.StringVar(value=sorted_years[0])
        month_var = tk.StringVar(value=datetime.now().strftime("%B"))
        
        opts_frame = tk.Frame(win, bg=CLR_CARD)
        opts_frame.pack(fill="x", pady=5)
        
        year_frame = tk.Frame(win, bg=CLR_CARD)
        month_frame = tk.Frame(win, bg=CLR_CARD)

        def update_ui():
            year_frame.pack_forget(); month_frame.pack_forget()
            if mode_var.get() == "By Year": year_frame.pack(fill="x", pady=10)
            elif mode_var.get() == "By Month": month_frame.pack(fill="x", pady=10)

        for txt, val in [("Current View (Filtered)", "Current View"), ("By Year", "By Year"), ("By Month", "By Month")]:
            tk.Radiobutton(opts_frame, text=txt, variable=mode_var, value=val, command=update_ui, bg=CLR_CARD, font=("Segoe UI", 10)).pack(anchor="w", pady=2)

        tk.Label(year_frame, text="Select Year:", bg=CLR_CARD, font=("Segoe UI", 9)).pack(anchor="w")
        ttk.Combobox(year_frame, textvariable=year_var, values=sorted_years, state="readonly").pack(fill="x", pady=5)

        tk.Label(month_frame, text="Select Month & Year:", bg=CLR_CARD, font=("Segoe UI", 9)).pack(anchor="w")
        m_row = tk.Frame(month_frame, bg=CLR_CARD)
        m_row.pack(fill="x", pady=5)
        months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        ttk.Combobox(m_row, textvariable=month_var, values=months, state="readonly", width=15).pack(side="left", padx=(0,5))
        ttk.Combobox(m_row, textvariable=year_var, values=sorted_years, state="readonly", width=10).pack(side="left")

        def proceed():
            mode = mode_var.get()
            data = []
            suffix = ""
            
            if mode == "Current View":
                data = [self.tree.item(i)['values'] for i in self.tree.get_children()]
                suffix = "CurrentView"
            elif mode == "By Year":
                y = year_var.get()
                data = [r for r in self.master_logs if len(r) > 0 and str(r[0]).startswith(y)]
                suffix = f"Year_{y}"
            elif mode == "By Month":
                y = year_var.get()
                m = month_var.get()
                try:
                    m_num = datetime.strptime(m, "%B").month
                    prefix = f"{y}-{m_num:02d}"
                    data = [r for r in self.master_logs if len(r) > 0 and str(r[0]).startswith(prefix)]
                    suffix = f"{m}_{y}"
                except: pass

            if not data: return messagebox.showwarning("Export", "No records found.")
            
            win.destroy()
            self._perform_export(data, suffix)

        create_modern_button(win, "Export", proceed).pack(fill="x", side="bottom")

    def _perform_export(self, data, suffix):
        try:
            filename = f"Logbook_{suffix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            file_path = filedialog.asksaveasfilename(initialfile=filename, defaultextension=".csv", filetypes=[("CSV files", "*.csv")], title="Export Data")
            if file_path:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(self.headers)
                    writer.writerows(data)
                if messagebox.askyesno("Success", f"Exported to {file_path}\nOpen now?"):
                    try: os.startfile(file_path)
                    except: pass
        except Exception as e: messagebox.showerror("Error", f"Export failed: {e}")

    def draw_logs(self):
        for w in self.p_logs.winfo_children(): w.destroy()
        h = tk.Frame(self.p_logs, bg=CLR_CARD, pady=20, padx=30); h.pack(fill="x")
        tk.Label(h, text="Digital Logbook", font=("Segoe UI", 18, "bold"), bg=CLR_CARD).pack(side="left")
        
        tk.Button(h, text="👤", font=("Segoe UI", 16), bg=CLR_CARD, fg=CLR_PRIMARY, 
                  bd=0, relief="flat", cursor="hand2", command=self.open_profile_settings).pack(side="right", padx=(10, 0))

        btn_frame = tk.Frame(h, bg=CLR_CARD); btn_frame.pack(side="right", padx=5)
        create_modern_button(btn_frame, "+ Add New Record", self.open_new_record_dialog).pack(side="left", padx=5)
        create_modern_button(btn_frame, "Export to Excel", self.export_logs, bg=CLR_SUCCESS).pack(side="left", padx=5)
        
        s_frame = tk.Frame(h, bg="#f1f3f4", padx=10); s_frame.pack(side="right", padx=10)
        self.s_var = tk.StringVar(); self.s_var.trace("w", self.filter_data)
        tk.Label(s_frame, text="🔍", bg="#f1f3f4").pack(side="left")
        tk.Entry(s_frame, textvariable=self.s_var, font=("Segoe UI", 11), bg="#f1f3f4", relief="flat", width=30, bd=0).pack(side="left", ipady=8)

        t_card = tk.Frame(self.p_logs, bg=CLR_CARD, padx=1, pady=1); t_card.pack(fill="both", expand=True, padx=30, pady=30)
        
        t_card.grid_rowconfigure(0, weight=1)
        t_card.grid_columnconfigure(0, weight=1)

        # Trigger background fetch if needed
        if self.ws_logs and not self.fetching_logs and not self.data_loaded_from_net:
            self.fetching_logs = True
            threading.Thread(target=self._fetch_logs_thread, daemon=True).start()

        if not self.master_logs:
            # Show loading state
            self.tree = ttk.Treeview(t_card, columns=("Status",), show="headings")
            self.tree.heading("Status", text="Status")
            self.tree.column("Status", width=200, anchor="center")
            self.tree.grid(row=0, column=0, sticky="nsew")
            
            msg = "Loading data..." if (self.fetching_logs or self.connecting) else "No Data Available"
            self.tree.insert("", "end", values=(msg,))
            return

        self.tree = ttk.Treeview(t_card, columns=self.headers, show="headings")
        
        vsb = ttk.Scrollbar(t_card, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(t_card, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        for c in self.headers: self.tree.heading(c, text=c.upper()); self.tree.column(c, width=150, minwidth=100)
        
        current_year = datetime.now().strftime("%Y")
        for r in sorted(self.master_logs, key=lambda x: x[1] if len(x) > 1 else "", reverse=True):
            if len(r) > 0 and str(r[0]).startswith(current_year):
                self.tree.insert("", "end", values=r)
        
        self.log_menu = tk.Menu(self.root, tearoff=0)
        self.log_menu.add_command(label="Edit Record", command=self.edit_log_entry)
        self.tree.bind("<Button-3>", self.show_log_menu)

    def _fetch_logs_thread(self):
        try:
            raw = self.ws_logs.get_all_values()
            if raw:
                self.headers = raw[0]
                self.master_logs = raw[1:]
                try:
                    with open(LOGS_CACHE_FILE, 'w') as f:
                        json.dump({"headers": self.headers, "rows": self.master_logs}, f)
                except: pass
            else:
                self.headers = ["No Data"]
                self.master_logs = []
        except Exception as e:
            print(f"Log fetch error: {e}")
            if not self.master_logs: self.headers = ["Error"]
        
        self.fetching_logs = False
        self.data_loaded_from_net = True
        self.root.after(0, self.draw_logs)

    def generate_next_ref_number(self):
        """Generates the next reference number: YYCAR32-XXX"""
        try:
            yr = datetime.now().strftime("%y")
            prefix = f"{yr}CAR32"
            
            # Fetch existing references (Column 2)
            col_values = self.ws_logs.col_values(2)
            
            # Find max sequence for current year
            max_seq = 0
            for ref in col_values:
                if str(ref).startswith(prefix):
                    try:
                        seq = int(str(ref).split('-')[-1])
                        if seq > max_seq: max_seq = seq
                    except: pass
            
            return f"{prefix}-{max_seq + 1:03d}"
        except Exception as e: raise Exception(f"Ref Gen Error: {e}")

    def open_new_record_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("New Record")
        ww, wh = 540, 620
        win.geometry(f"{ww}x{wh}")
        win.configure(bg=CLR_CARD, padx=30, pady=30)

        win.transient(self.root)
        win.grab_set()

        self.root.update_idletasks()
        rx = self.root.winfo_x(); ry = self.root.winfo_y()
        rw = self.root.winfo_width(); rh = self.root.winfo_height()
        x = rx + max(0, (rw - ww) // 2)
        y = ry + max(0, (rh - wh) // 2)
        win.geometry(f"{ww}x{wh}+{x}+{y}")

        vars_local = {k: tk.StringVar() for k in ["part", "addr", "trans", "sec", "mode", "rem"]}
        if self.current_user_name: vars_local["trans"].set(self.current_user_name)
        create_labeled_entry(win, "Particulars", vars_local["part"])
        
        create_labeled_entry(win, "Addressee", vars_local["addr"])
        
        tk.Label(win, text="Transmitter", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w", pady=(8, 0))
        cb_trans = ttk.Combobox(win, textvariable=vars_local["trans"], font=("Segoe UI", 11), values=self.user_names_cache, state="readonly")
        cb_trans.pack(fill="x", pady=6, ipady=4)

        valid_sections = ["PSO", "Admin", "CRS", "PhilSys", "Statistical"]
        valid_modes = ["Email", "Walk-in", "Hand Carry", "JRS", "Routing", "Google Link"]

        tk.Label(win, text="Section", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w", pady=(12, 0))
        cb_sec = ttk.Combobox(win, textvariable=vars_local["sec"], font=("Segoe UI", 11), height=10, values=valid_sections, state="readonly")
        cb_sec.pack(fill="x", pady=8, ipady=4)

        tk.Label(win, text="Mode of Transmittal", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w", pady=(8, 0))
        create_multi_select_dropdown(win, valid_modes, vars_local["mode"]).pack(fill="x", pady=6)

        create_labeled_entry(win, "Remarks (Optional)", vars_local["rem"])

        btn = create_modern_button(win, "Save Record", None, width=24)
        btn.pack(fill="x", pady=20)

        def submit():
            if not all(vars_local[k].get().strip() for k in ["part", "addr", "trans", "sec", "mode"]):
                return messagebox.showwarning("Error", "All fields are required")
            
            addr_clean = vars_local["addr"].get().replace(" ", "").lower()
            if addr_clean in ["co", "ro", "centraloffice", "regionaloffice"]:
                return messagebox.showwarning("Invalid Input", "Generic addressees like 'CO', 'RO', 'Central Office', or 'Regional Office' are not allowed.")
            
            if vars_local["sec"].get() not in valid_sections:
                return messagebox.showwarning("Invalid Input", "Please select a valid Section from the list.")
            
            if not vars_local["mode"].get():
                return messagebox.showwarning("Invalid Input", "Please select at least one Mode.")

            if self.user_names_cache and vars_local["trans"].get() not in self.user_names_cache:
                return messagebox.showwarning("Invalid Input", "Transmitter must be a registered user.")
            
            if not check_internet():
                return messagebox.showerror("Network Error", "No internet connection detected.")

            win.protocol("WM_DELETE_WINDOW", lambda: None)

            d = [datetime.now().strftime("%Y-%m-%d %H:%M"), "REF", vars_local["part"].get(), vars_local["addr"].get(), vars_local["trans"].get(), vars_local["sec"].get(), vars_local["mode"].get(), vars_local["rem"].get(), self.email]
            btn.config(text="Saving...", state="disabled")
            threading.Thread(target=self._save_and_close, args=(d, win, btn)).start()
        
        btn.config(command=submit)

    def _save_and_close(self, d, win, btn):
        try:
            # Generate Auto-Reference Number
            d[1] = self.generate_next_ref_number()
            
            self.ws_logs.append_row(d)
            self.master_logs = [] # Force refresh from sheet on next draw
            self.data_loaded_from_net = False
            self.root.after(0, lambda: self._save_success(win, d[1]))
        except Exception as e:
            self.root.after(0, lambda: self._save_error(e, btn, win))

    def _save_success(self, win, ref_num=None):
        win.grab_release()
        msg = "Saved successfully"
        if ref_num:
            msg += f"\n\nReference Number: {ref_num}"
        messagebox.showinfo("Success", msg)
        win.destroy()
        self.draw_logs()

    def _save_error(self, e, btn, win):
        win.grab_release()
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.focus_force()

        messagebox.showerror("Error", f"Failed to save record:\n{e}")

        btn.config(text="Save Record", state="normal")

    def show_log_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.log_menu.post(event.x_root, event.y_root)

    def edit_log_entry(self):
        sel = self.tree.selection()
        if not sel: return
        val = self.tree.item(sel[0])['values']
        if len(val) > 8 and str(val[8]) != self.email:
            return messagebox.showerror("Access Denied", "You can only edit your own records.")
        self.open_edit_dialog(val)

    def open_edit_dialog(self, val):
        win = tk.Toplevel(self.root)
        win.title(f"Edit Record {val[1]}")
        win.geometry("540x620")
        win.configure(bg=CLR_CARD, padx=30, pady=30)
        win.transient(self.root); win.grab_set()
        
        self.root.update_idletasks()
        rx = self.root.winfo_x(); ry = self.root.winfo_y()
        win.geometry(f"+{rx + (self.root.winfo_width()-540)//2}+{ry + (self.root.winfo_height()-540)//2}")

        rem_val = val[7] if len(val) > 7 else ""
        v = {k: tk.StringVar(value=val[i]) for k, i in zip(["part", "addr", "trans", "sec", "mode"], [2, 3, 4, 5, 6])}
        v["rem"] = tk.StringVar(value=rem_val)
        create_labeled_entry(win, "Particulars", v["part"])
        
        create_labeled_entry(win, "Addressee", v["addr"])
        
        tk.Label(win, text="Transmitter", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w", pady=(8, 0))
        cb_trans = ttk.Combobox(win, textvariable=v["trans"], font=("Segoe UI", 11), values=self.user_names_cache, state="readonly")
        cb_trans.pack(fill="x", pady=6, ipady=4)

        valid_sections = ["PSO", "Admin", "CRS", "PhilSys", "Statistical"]
        valid_modes = ["Email", "Walk-in", "Hand Carry", "JRS", "Routing", "Google Link"]

        tk.Label(win, text="Section", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w", pady=(12, 0))
        cb_sec = ttk.Combobox(win, textvariable=v["sec"], font=("Segoe UI", 11), values=valid_sections, state="readonly")
        cb_sec.pack(fill="x", pady=8, ipady=4)

        tk.Label(win, text="Mode of Transmittal", font=("Segoe UI", 9, "bold"), bg=CLR_CARD).pack(anchor="w", pady=(8, 0))
        create_multi_select_dropdown(win, valid_modes, v["mode"]).pack(fill="x", pady=6)

        create_labeled_entry(win, "Remarks (Optional)", v["rem"])

        btn = create_modern_button(win, "Update Record", None)
        btn.pack(fill="x", pady=20)
        
        def submit():
            if v["sec"].get() not in valid_sections:
                return messagebox.showwarning("Invalid Input", "Please select a valid Section.")
            
            if not v["mode"].get():
                return messagebox.showwarning("Invalid Input", "Please select at least one Mode.")

            if self.user_names_cache and v["trans"].get() not in self.user_names_cache:
                return messagebox.showwarning("Invalid Input", "Transmitter must be a registered user.")

            win.protocol("WM_DELETE_WINDOW", lambda: None)
            btn.config(text="Updating...", state="disabled")
            threading.Thread(target=self._update_thread, args=(val[1], [v[k].get() for k in ["part", "addr", "trans", "sec", "mode", "rem"]], win, btn)).start()
        btn.config(command=submit)

    def _update_thread(self, ref, data, win, btn):
        try:
            r = self.ws_logs.find(ref).row
            # Update columns C through H (Particulars to Remarks) in a single call
            cells_to_update = self.ws_logs.range(f"C{r}:H{r}")
            for i, cell in enumerate(cells_to_update):
                cell.value = data[i]
            self.ws_logs.update_cells(cells_to_update)

            self.master_logs = [] # Force refresh from sheet on next draw
            self.data_loaded_from_net = False
            self.root.after(0, lambda: self._save_success(win, ref))
        except Exception as e:
            self.root.after(0, lambda: self._save_error(e, btn, win))

    def filter_data(self, *args):
        q = self.s_var.get().lower()
        current_year = datetime.now().strftime("%Y")
        for i in self.tree.get_children(): self.tree.delete(i)
        for r in sorted(self.master_logs, key=lambda x: x[1] if len(x) > 1 else "", reverse=True):
            if len(r) > 0 and str(r[0]).startswith(current_year):
                if any(q in str(cell).lower() for cell in r): self.tree.insert("", "end", values=r)

if __name__ == "__main__":
    root = tk.Tk()
    
    try:
        if os.path.exists(Agency_Logo):
            root.iconphoto(True, ImageTk.PhotoImage(Image.open(Agency_Logo)))
    except Exception: pass

    def start(email, user_data=None):
        for w in root.winfo_children(): w.destroy()
        LogbookApp(root, email, user_data=user_data, on_success_callback=start)
    
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE, "r") as f:
                saved_email = f.read().strip()
            if saved_email: 
                cached = get_cache_data(saved_email)
                u_data = cached["user_data"] if cached else None
                start(saved_email, user_data=u_data)
            else: 
                LoginWindow(root, start)
        except: 
            LoginWindow(root, start)
    else:
        LoginWindow(root, start)
        
    root.mainloop()