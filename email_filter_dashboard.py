#!/usr/bin/env python3
"""
Email Filter and Dashboard System
A Tkinter-based desktop application for filtering and displaying emails.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime, parseaddr
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import calendar as py_calendar
from demo_data import load_demo_emails_between
import csv
import json
import threading
import time


class EmailFilterApp:
    """Main application class for Email Filter and Dashboard system."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Email Filter & Dashboard")
        self.root.geometry("1000x720")
        self.root.minsize(800, 600)
        
        # Theme palette and styles
        self.theme_var = tk.StringVar(value='Light')
        self.setup_theme_palette(self.theme_var.get())
        self.root.configure(bg=self.colors['app_bg'])
        
        # Email data storage
        self.emails_data = []
        self.dashboard_window = None
        self.filtered_emails = []
        
        # UI state variables
        self.animation_running = False
        self.notification_queue = []
        
        # Configure style
        self.setup_styles()
        
        # Set up keyboard shortcuts
        self.setup_keyboard_shortcuts()
        
        # Create main interface
        self.create_email_filter_screen()
        
        # Bind Enter key to fetch emails at application level
        self.root.bind('<Return>', lambda e: self.fetch_emails())
        
        # Start notification system
        self.process_notifications()
        
    def setup_theme_palette(self, theme_name: str):
        """Define professional, accessible color palettes and set active colors."""
        # Persist themes on the instance so the selector only shows Light and Dark
        themes = {
            'Light': {
                'app_bg': '#f8fafc', 'surface': '#ffffff', 'surface_alt': '#f1f5f9',
                'text': '#0f172a', 'text_muted': '#334155',
                'primary': '#2563eb', 'primary_hover': '#1d4ed8', 'accent': '#0891b2',
                'border': '#e2e8f0', 'input_bg': '#ffffff', 'input_fg': '#0f172a',
                'placeholder': '#64748b', 'success': '#059669', 'warning': '#d97706', 'danger': '#dc2626'
            },
            'Dark': {
                'app_bg': '#000000', 'surface': '#0a0a0a', 'surface_alt': '#0f0f0f',
                'text': '#e5e7eb', 'text_muted': '#cbd5e1',
                'primary': '#3b82f6', 'primary_hover': '#2563eb', 'accent': '#22d3ee',
                'border': '#1a1a1a', 'input_bg': '#0a0a0a', 'input_fg': '#e5e7eb',
                'placeholder': '#94a3b8', 'success': '#10b981', 'warning': '#f59e0b', 'danger': '#ef4444'
            }
        }
        self.themes = themes
        self.colors = themes.get(theme_name, themes['Dark'])

    def setup_styles(self):
        """Configure modern styling for the application."""
        style = ttk.Style()
        style.theme_use('clam')

        c = self.colors

        # Base containers
        style.configure('TFrame', background=c['app_bg'])
        style.configure('TLabelframe', background=c['surface'], bordercolor=c['border'], relief='solid')
        style.configure('TLabelframe.Label', background=c['surface'], foreground=c['text'])

        # Labels
        style.configure('Title.TLabel', font=('Segoe UI', 20, 'bold'), background=c['app_bg'], foreground=c['text'])
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'), background=c['surface'], foreground=c['text_muted'])
        style.configure('TLabel', background=c['app_bg'], foreground=c['text'])

        # Buttons
        style.configure('Custom.TButton', font=('Segoe UI', 11, 'bold'), padding=8, background=c['primary'], foreground='#ffffff', bordercolor=c['primary'])
        style.map('Custom.TButton',
                  background=[('active', c['primary_hover']), ('pressed', c['primary_hover'])],
                  foreground=[('disabled', '#7c8596')])
        
        # Success button style
        style.configure('Success.TButton', font=('Segoe UI', 11, 'bold'), padding=8, background=c['success'], foreground='#ffffff', bordercolor=c['success'])
        style.map('Success.TButton',
                  background=[('active', '#047857'), ('pressed', '#047857')])
        
        # Warning button style
        style.configure('Warning.TButton', font=('Segoe UI', 11, 'bold'), padding=8, background=c['warning'], foreground='#ffffff', bordercolor=c['warning'])
        style.map('Warning.TButton',
                  background=[('active', '#b45309'), ('pressed', '#b45309')])
        
        # Danger button style
        style.configure('Danger.TButton', font=('Segoe UI', 11, 'bold'), padding=8, background=c['danger'], foreground='#ffffff', bordercolor=c['danger'])
        style.map('Danger.TButton',
                  background=[('active', '#b91c1c'), ('pressed', '#b91c1c')])

        # Quick select button styles
        style.configure('Today.TButton', font=('Segoe UI', 11, 'bold'), padding=8, background='#3b82f6', foreground='#ffffff', bordercolor='#3b82f6')
        style.map('Today.TButton', background=[('active', '#2563eb'), ('pressed', '#2563eb')])
        
        style.configure('7Days.TButton', font=('Segoe UI', 11, 'bold'), padding=8, background='#10b981', foreground='#ffffff', bordercolor='#10b981')
        style.map('7Days.TButton', background=[('active', '#059669'), ('pressed', '#059669')])

        style.configure('30Days.TButton', font=('Segoe UI', 11, 'bold'), padding=8, background='#f97316', foreground='#ffffff', bordercolor='#f97316')
        style.map('30Days.TButton', background=[('active', '#ea580c'), ('pressed', '#ea580c')])

        style.configure('90Days.TButton', font=('Segoe UI', 11, 'bold'), padding=8, background='#8b5cf6', foreground='#ffffff', bordercolor='#8b5cf6')
        style.map('90Days.TButton', background=[('active', '#7c3aed'), ('pressed', '#7c3aed')])

        # Entries
        style.configure('TEntry', fieldbackground=c['input_bg'], foreground=c['input_fg'], insertcolor=c['input_fg'], bordercolor=c['border'])

        # Combobox
        style.configure('TCombobox', fieldbackground=c['input_bg'], background=c['input_bg'], foreground=c['input_fg'], arrowcolor=c['text'])

        # Treeview
        style.configure('Custom.Treeview', background=c['surface'], fieldbackground=c['surface'], foreground=c['text'], bordercolor=c['border'], rowheight=40)
        style.configure('Treeview.Heading', background=c['surface_alt'], foreground=c['text_muted'], relief='flat')
        style.map('Treeview.Heading', background=[('active', c['surface_alt'])])

        # Scrollbar
        style.configure('Vertical.TScrollbar', background=c['surface_alt'], troughcolor=c['surface'], bordercolor=c['border'])
        style.configure('Horizontal.TScrollbar', background=c['surface_alt'], troughcolor=c['surface'], bordercolor=c['border'])
        
        # Root background can change by theme
        try:
            self.root.configure(bg=c['app_bg'])
        except Exception:
            pass

        # Separators
        style.configure('TSeparator', background=c['border'])

        # Footer/status label
        style.configure('Status.TLabel', background=c['surface'], foreground=c['text_muted'], font=('Segoe UI', 9))
        
        # Notification styles
        style.configure('Success.TLabel', background=c['success'], foreground='#ffffff', font=('Segoe UI', 10, 'bold'), padding=8)
        style.configure('Warning.TLabel', background=c['warning'], foreground='#ffffff', font=('Segoe UI', 10, 'bold'), padding=8)
        style.configure('Error.TLabel', background=c['danger'], foreground='#ffffff', font=('Segoe UI', 10, 'bold'), padding=8)

    def apply_theme(self, theme_name: str):
        """Apply a theme by name across the app without changing functionality."""
        self.setup_theme_palette(theme_name)
        self.setup_styles()
        # Update existing toplevels if present
        try:
            self.root.configure(bg=self.colors['app_bg'])
        except Exception:
            pass
        try:
            if self.dashboard_window and self.dashboard_window.winfo_exists():
                self.dashboard_window.configure(bg=self.colors['app_bg'])
        except Exception:
            pass
    
    def toggle_theme(self):
        """Toggle between light and dark themes."""
        current_theme = self.theme_var.get()
        if current_theme == 'Light':
            new_theme = 'Dark'
            self.theme_button.config(text="☀️")
        else:
            new_theme = 'Light'
            self.theme_button.config(text="🌙")
        
        self.theme_var.set(new_theme)
        self.apply_theme(new_theme)
        self.show_notification(f"Theme changed to {new_theme}", "success", 2000)
    
    def setup_keyboard_shortcuts(self):
        """Set up keyboard shortcuts for common actions."""
        self.root.bind('<Control-f>', lambda e: self.focus_search() if hasattr(self, 'search_var') else None)
        self.root.bind('<Control-r>', lambda e: self.fetch_emails())
        self.root.bind('<Control-d>', lambda e: self.open_dashboard())
        self.root.bind('<Control-t>', lambda e: self.test_imap_connection())
        self.root.bind('<F5>', lambda e: self.refresh_dashboard() if hasattr(self, 'refresh_dashboard') else None)
        self.root.bind('<Escape>', lambda e: self.clear_search() if hasattr(self, 'search_var') else None)
    
    def focus_search(self):
        """Focus on search entry if dashboard is open."""
        if hasattr(self, 'search_entry') and self.search_entry.winfo_exists():
            self.search_entry.focus_set()
    
    def clear_search(self):
        """Clear search query."""
        if hasattr(self, 'search_var'):
            self.set_search_query("")
    
    def show_notification(self, message: str, notification_type: str = "info", duration: int = 3000):
        """Show a temporary notification at the top of the window."""
        self.notification_queue.append({
            'message': message,
            'type': notification_type,
            'duration': duration,
            'timestamp': time.time()
        })
    
    def process_notifications(self):
        """Process notification queue and display notifications."""
        if self.notification_queue and not self.animation_running:
            notification = self.notification_queue.pop(0)
            self.display_notification(notification)
        
        # Schedule next check
        self.root.after(100, self.process_notifications)
    
    def display_notification(self, notification):
        """Display a notification with fade-in/fade-out animation."""
        if self.animation_running:
            return
        
        self.animation_running = True
        
        # Create notification frame
        notification_frame = ttk.Frame(self.root)
        notification_frame.place(relx=0.5, y=10, anchor='n')
        
        # Style based on type
        style_map = {
            'success': 'Success.TLabel',
            'warning': 'Warning.TLabel',
            'error': 'Error.TLabel',
            'info': 'Status.TLabel'
        }
        
        style = style_map.get(notification['type'], 'Status.TLabel')
        
        # Create label
        label = ttk.Label(notification_frame, text=notification['message'], style=style)
        label.pack(padx=20, pady=10)
        
        # Animate in
        self.animate_notification_in(notification_frame, notification['duration'])
    
    def animate_notification_in(self, frame, duration):
        """Animate notification sliding in from top."""
        def slide_in(step=0):
            if step < 10:
                y = 10 - (step * 2)  # Slide down
                frame.place(relx=0.5, y=y, anchor='n')
                self.root.after(20, lambda: slide_in(step + 1))
            else:
                # Keep visible for duration, then slide out
                self.root.after(duration, lambda: self.animate_notification_out(frame))
        
        slide_in()
    
    def animate_notification_out(self, frame):
        """Animate notification sliding out to top."""
        def slide_out(step=0):
            if step < 10:
                y = 10 - (step * 2)  # Slide up
                frame.place(relx=0.5, y=y, anchor='n')
                self.root.after(20, lambda: slide_out(step + 1))
            else:
                frame.destroy()
                self.animation_running = False
        
        slide_out()
        
    def create_email_filter_screen(self):
        """Create the main email filter interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header bar with enhanced layout
        header_bar = ttk.Frame(main_frame)
        header_bar.pack(fill=tk.X)
        
        # Title with icon-like symbol
        title_frame = ttk.Frame(header_bar)
        title_frame.pack(side=tk.LEFT)
        ttk.Label(title_frame, text="📧", font=('Segoe UI', 20)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(title_frame, text="Email Filter & Dashboard", style='Title.TLabel').pack(side=tk.LEFT)

        # Control panel (right side)
        control_frame = ttk.Frame(header_bar)
        control_frame.pack(side=tk.RIGHT)
        
        # Help button
        help_btn = ttk.Button(control_frame, text="❓", width=3, 
                             command=self.show_help_dialog)
        help_btn.pack(side=tk.RIGHT, padx=(5, 0))
        self._add_tooltip(help_btn, "Show keyboard shortcuts and help")
        
        # Theme toggle button
        self.theme_button = ttk.Button(control_frame, text="🌙", width=3, 
                                     command=self.toggle_theme)
        self.theme_button.pack(side=tk.RIGHT, padx=(0, 5))
        self._add_tooltip(self.theme_button, "Toggle light/dark theme")

        # Structural separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=(12, 16))

        # Date selection frame
        date_frame = ttk.LabelFrame(main_frame, text="Select Date Range", padding="15")
        date_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Quick date selection buttons
        quick_date_frame = ttk.Frame(date_frame)
        quick_date_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(quick_date_frame, text="Quick Select:", style='Header.TLabel').pack(side=tk.LEFT)
        
        quick_buttons = [
            ("Today", lambda: self.set_quick_date_range(0), "Today.TButton"),
            ("Last 7 Days", lambda: self.set_quick_date_range(7), "7Days.TButton"),
            ("Last 30 Days", lambda: self.set_quick_date_range(30), "30Days.TButton"),
            ("Last 90 Days", lambda: self.set_quick_date_range(90), "90Days.TButton")
        ]
        
        for text, command, style in quick_buttons:
            btn = ttk.Button(quick_date_frame, text=text, command=command, style=style)
            btn.pack(side=tk.LEFT, padx=(10, 0))
            self._add_tooltip(btn, f"Set date range to {text.lower()}")
        
        # Start date
        start_date_frame = ttk.Frame(date_frame)
        start_date_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(start_date_frame, text="Start Date:", style='Header.TLabel').pack(side=tk.LEFT)
        today_str = datetime.today().strftime('%Y-%m-%d')
        self.start_date_var = tk.StringVar(value=today_str)
        self.start_date_entry = ttk.Entry(start_date_frame, textvariable=self.start_date_var, width=15)
        self.start_date_entry.pack(side=tk.LEFT, padx=(10, 0))
        self.start_date_entry.bind('<Return>', lambda e: self.fetch_emails())
        
        start_cal_btn = ttk.Button(start_date_frame, text="📅", width=3,
                                  command=lambda: self.select_date('start'))
        start_cal_btn.pack(side=tk.LEFT, padx=(5, 0))
        self._add_tooltip(start_cal_btn, "Open calendar picker")
        
        # End date
        end_date_frame = ttk.Frame(date_frame)
        end_date_frame.pack(fill=tk.X)
        
        ttk.Label(end_date_frame, text="End Date:", style='Header.TLabel').pack(side=tk.LEFT)
        self.end_date_var = tk.StringVar(value=today_str)
        self.end_date_entry = ttk.Entry(end_date_frame, textvariable=self.end_date_var, width=15)
        self.end_date_entry.pack(side=tk.LEFT, padx=(10, 0))
        self.end_date_entry.bind('<Return>', lambda e: self.fetch_emails())
        
        end_cal_btn = ttk.Button(end_date_frame, text="📅", width=3,
                                command=lambda: self.select_date('end'))
        end_cal_btn.pack(side=tk.LEFT, padx=(5, 0))
        self._add_tooltip(end_cal_btn, "Open calendar picker")
        
        # Date validation indicators
        self.start_date_indicator = ttk.Label(start_date_frame, text="✓", foreground=self.colors['success'])
        self.start_date_indicator.pack(side=tk.LEFT, padx=(5, 0))
        self.end_date_indicator = ttk.Label(end_date_frame, text="✓", foreground=self.colors['success'])
        self.end_date_indicator.pack(side=tk.LEFT, padx=(5, 0))
        
        # Bind validation
        self.start_date_var.trace('w', self.validate_dates)
        self.end_date_var.trace('w', self.validate_dates)
        
        # Email configuration frame
        config_frame = ttk.LabelFrame(main_frame, text="Email Configuration", padding="15")
        config_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Server selection
        server_frame = ttk.Frame(config_frame)
        server_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(server_frame, text="Email Server:", style='Header.TLabel').pack(side=tk.LEFT)
        self.server_var = tk.StringVar(value="Gmail")
        server_combo = ttk.Combobox(server_frame, textvariable=self.server_var, 
                                   values=["Gmail", "Outlook", "Custom"], width=15)
        server_combo.pack(side=tk.LEFT, padx=(10, 0))
        
        # Credentials
        cred_frame = ttk.Frame(config_frame)
        cred_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(cred_frame, text="Email:", style='Header.TLabel').pack(side=tk.LEFT)
        self.email_var = tk.StringVar()
        ttk.Entry(cred_frame, textvariable=self.email_var, width=25).pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Label(cred_frame, text="Password:", style='Header.TLabel').pack(side=tk.LEFT, padx=(20, 0))
        self.password_var = tk.StringVar()
        ttk.Entry(cred_frame, textvariable=self.password_var, show="*", width=20).pack(side=tk.LEFT, padx=(10, 0))
        
        # Demo mode checkbox (default unchecked)
        self.demo_mode_var = tk.BooleanVar(value=False)
        demo_check = ttk.Checkbutton(config_frame, text="Use Demo Data (for testing)", 
                                   variable=self.demo_mode_var)
        demo_check.pack(anchor=tk.W, pady=(10, 0))
        
        # Action buttons with enhanced layout
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Primary actions (left side)
        primary_frame = ttk.Frame(button_frame)
        primary_frame.pack(side=tk.LEFT)
        
        self.fetch_button = ttk.Button(primary_frame, text="🔄 Fetch Emails", 
                                      command=self.fetch_emails, style='Custom.TButton',
                                      default='active')  # Makes this the default button for Enter key
        self.fetch_button.pack(side=tk.LEFT, padx=(0, 10))
        self._add_tooltip(self.fetch_button, "Fetch emails for selected date range (Ctrl+R or Enter)")
        
        self.test_button = ttk.Button(primary_frame, text="🔗 Test Connection", 
                                     command=self.test_imap_connection, style='Custom.TButton')
        self.test_button.pack(side=tk.LEFT, padx=(0, 10))
        self._add_tooltip(self.test_button, "Test IMAP connection (Ctrl+T)")

        self.dashboard_button = ttk.Button(primary_frame, text="📊 Dashboard", 
                                          command=self.open_dashboard, style='Success.TButton')
        self.dashboard_button.pack(side=tk.LEFT, padx=(0, 10))
        self._add_tooltip(self.dashboard_button, "Open email dashboard (Ctrl+D)")
        
        # Secondary actions (right side)
        secondary_frame = ttk.Frame(button_frame)
        secondary_frame.pack(side=tk.RIGHT)
        
        export_button = ttk.Button(secondary_frame, text="📤 Export", 
                                  command=self.show_export_options)
        export_button.pack(side=tk.LEFT, padx=(10, 0))
        self._add_tooltip(export_button, "Export email data to file")
        
        settings_button = ttk.Button(secondary_frame, text="⚙️ Settings", 
                                    command=self.show_settings_dialog)
        settings_button.pack(side=tk.LEFT, padx=(10, 0))
        self._add_tooltip(settings_button, "Application settings")
        
        # Status label
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=(16, 10))
        self.status_label = ttk.Label(main_frame, text="Ready to fetch emails", style='Status.TLabel')
        self.status_label.pack(fill=tk.X)
        # Loading progress bar (hidden by default)
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=(6, 0))
        self.progress_bar.pack_forget()

        # Simple tooltips for key controls
        try:
            self._install_tooltips({
                'fetch': self.fetch_button,
                'test': button_frame.winfo_children()[1],
                'dash': button_frame.winfo_children()[2],
            })
        except Exception:
            pass
        
    def select_date(self, date_type):
        """Open a calendar-style date picker with month grid and navigation."""
        def parse_default_date(value: str) -> datetime:
            try:
                return datetime.strptime(value, '%Y-%m-%d')
            except Exception:
                return datetime.today()

        default_dt = parse_default_date(self.start_date_var.get() if date_type == 'start' else self.end_date_var.get())

        date_window = tk.Toplevel(self.root)
        date_window.title(f"Select {date_type.title()} Date")
        date_window.geometry("340x360")
        date_window.transient(self.root)
        date_window.grab_set()

        # Apply top-level background
        try:
            date_window.configure(bg=self.colors['app_bg'])
        except Exception:
            pass

        container = ttk.Frame(date_window, padding="12")
        container.pack(fill=tk.BOTH, expand=True)

        # Header with navigation and direct month/year selection
        header = ttk.Frame(container)
        header.pack(fill=tk.X)

        view_year = tk.IntVar(value=default_dt.year)
        view_month = tk.IntVar(value=default_dt.month)

        def month_name(y: int, m: int) -> str:
            return f"{py_calendar.month_name[m]} {y}"

        title_var = tk.StringVar(value=month_name(view_year.get(), view_month.get()))

        # Month and Year Comboboxes for direct selection
        months_list = [py_calendar.month_name[m] for m in range(1, 13)]
        month_cb_var = tk.StringVar(value=py_calendar.month_name[view_month.get()])
        year_cb_var = tk.StringVar(value=str(view_year.get()))
        current_year = datetime.today().year
        years_list = [str(y) for y in range(current_year - 50, current_year + 51)]

        def go_prev():
            y, m = view_year.get(), view_month.get()
            if m == 1:
                y -= 1
                m = 12
            else:
                m -= 1
            view_year.set(y)
            view_month.set(m)
            title_var.set(month_name(y, m))
            render_days()

        def go_next():
            y, m = view_year.get(), view_month.get()
            if m == 12:
                y += 1
                m = 1
            else:
                m += 1
            view_year.set(y)
            view_month.set(m)
            title_var.set(month_name(y, m))
            render_days()

        left_group = ttk.Frame(header)
        left_group.pack(side=tk.LEFT)
        ttk.Button(left_group, text="<", width=3, command=go_prev).pack(side=tk.LEFT)

        mid_group = ttk.Frame(header)
        mid_group.pack(side=tk.LEFT, expand=True)
        ttk.Label(mid_group, text="Month:").pack(side=tk.LEFT, padx=(8, 4))
        month_cb = ttk.Combobox(mid_group, state='readonly', values=months_list, textvariable=month_cb_var, width=12)
        month_cb.pack(side=tk.LEFT)
        ttk.Label(mid_group, text="Year:").pack(side=tk.LEFT, padx=(12, 4))
        year_cb = ttk.Combobox(mid_group, state='readonly', values=years_list, textvariable=year_cb_var, width=8)
        year_cb.pack(side=tk.LEFT)

        right_group = ttk.Frame(header)
        right_group.pack(side=tk.RIGHT)
        ttk.Button(right_group, text=">", width=3, command=go_next).pack(side=tk.RIGHT)

        def on_month_change(_=None):
            try:
                m_name = month_cb_var.get()
                m = months_list.index(m_name) + 1
                view_month.set(m)
                title_var.set(month_name(view_year.get(), view_month.get()))
                render_days()
            except Exception:
                pass

        def on_year_change(_=None):
            try:
                y = int(year_cb_var.get())
                view_year.set(y)
                title_var.set(month_name(view_year.get(), view_month.get()))
                render_days()
            except Exception:
                pass

        month_cb.bind('<<ComboboxSelected>>', on_month_change)
        year_cb.bind('<<ComboboxSelected>>', on_year_change)

        # Weekday headers
        # Title (optional, subtle under comboboxes)
        title_row = ttk.Frame(container)
        title_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(title_row, textvariable=title_var, style='Header.TLabel').pack(side=tk.LEFT)

        days_header = ttk.Frame(container)
        days_header.pack(fill=tk.X, pady=(10, 0))
        # Monday first (1..7) to match ISO
        weekday_names = [py_calendar.day_abbr[i] for i in range(7)]  # Mon..Sun depending on locale (Mon=0)
        for name in weekday_names:
            lbl = ttk.Label(days_header, text=name, width=4, anchor='center')
            lbl.pack(side=tk.LEFT, expand=True)

        # Days grid
        grid_frame = ttk.Frame(container)
        grid_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 6))

        # Action row
        action_row = ttk.Frame(container)
        action_row.pack(fill=tk.X)

        def on_today():
            today = datetime.today()
            view_year.set(today.year)
            view_month.set(today.month)
            title_var.set(month_name(today.year, today.month))
            render_days()

        def on_select_date(y: int, m: int, d: int):
            try:
                picked = datetime(year=y, month=m, day=d)
                value = picked.strftime('%Y-%m-%d')
                if date_type == 'start':
                    self.start_date_var.set(value)
                else:
                    self.end_date_var.set(value)
                date_window.destroy()
            except Exception as exc:
                messagebox.showerror("Invalid date", f"Please select a valid date.\n{exc}")

        def clear_grid():
            for child in grid_frame.winfo_children():
                child.destroy()

        def render_days():
            clear_grid()
            y, m = view_year.get(), view_month.get()
            cal = py_calendar.Calendar(firstweekday=0)  # 0=Monday
            today = datetime.today()
            for week in cal.monthdayscalendar(y, m):
                row = ttk.Frame(grid_frame)
                row.pack(fill=tk.X)
                for day in week:
                    if day == 0:
                        # Empty cell
                        btn = ttk.Label(row, text=' ', width=4, anchor='center')
                        btn.pack(side=tk.LEFT, expand=True, padx=1, pady=1)
                    else:
                        is_today = (y == today.year and m == today.month and day == today.day)
                        txt = str(day)
                        btn = ttk.Button(row, text=txt, width=4,
                                         command=lambda dd=day: on_select_date(y, m, dd))
                        if is_today:
                            # Visual hint for today
                            btn.configure(style='Today.TButton')
                        btn.pack(side=tk.LEFT, expand=True, padx=1, pady=1)

        # Optional style for today's button highlight
        try:
            style = ttk.Style()
            style.configure('Today.TButton', foreground=self.colors['accent'])
        except Exception:
            pass

        # Buttons
        ttk.Button(action_row, text="Today", command=on_today).pack(side=tk.LEFT)
        ttk.Button(action_row, text="Cancel", command=date_window.destroy).pack(side=tk.RIGHT)

        # Initial render
        render_days()
    
    def set_quick_date_range(self, days_back: int):
        """Set date range based on days back from today."""
        end_date = datetime.today()
        start_date = end_date - timedelta(days=days_back)
        
        self.start_date_var.set(start_date.strftime('%Y-%m-%d'))
        self.end_date_var.set(end_date.strftime('%Y-%m-%d'))
        
        # Show notification
        if days_back == 0:
            self.show_notification("Date range set to today", "success", 2000)
        else:
            self.show_notification(f"Date range set to last {days_back} days", "success", 2000)
    
    def validate_dates(self, *args):
        """Validate date entries and show visual feedback."""
        try:
            start_valid = True
            end_valid = True
            
            # Validate start date
            try:
                start_date = datetime.strptime(self.start_date_var.get(), '%Y-%m-%d')
            except:
                start_valid = False
            
            # Validate end date
            try:
                end_date = datetime.strptime(self.end_date_var.get(), '%Y-%m-%d')
            except:
                end_valid = False
            
            # Update indicators
            if hasattr(self, 'start_date_indicator'):
                if start_valid:
                    self.start_date_indicator.config(text="✓", foreground=self.colors['success'])
                else:
                    self.start_date_indicator.config(text="✗", foreground=self.colors['danger'])
            
            if hasattr(self, 'end_date_indicator'):
                if end_valid:
                    self.end_date_indicator.config(text="✓", foreground=self.colors['success'])
                else:
                    self.end_date_indicator.config(text="✗", foreground=self.colors['danger'])
            
            # Check date range validity
            if start_valid and end_valid:
                if start_date > end_date:
                    self.start_date_indicator.config(text="⚠", foreground=self.colors['warning'])
                    self.end_date_indicator.config(text="⚠", foreground=self.colors['warning'])
        except Exception:
            pass
    
    def show_help_dialog(self):
        """Show help dialog with keyboard shortcuts."""
        help_window = tk.Toplevel(self.root)
        help_window.title("Help & Keyboard Shortcuts")
        help_window.geometry("500x400")
        help_window.transient(self.root)
        help_window.grab_set()
        
        try:
            help_window.configure(bg=self.colors['app_bg'])
        except Exception:
            pass
        
        frame = ttk.Frame(help_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Help & Keyboard Shortcuts", style='Title.TLabel').pack(pady=(0, 20))
        
        help_text = """
Keyboard Shortcuts:
• Ctrl+R - Fetch emails
• Ctrl+T - Test connection
• Ctrl+D - Open dashboard
• Ctrl+F - Focus search (in dashboard)
• F5 - Refresh dashboard
• Escape - Clear search

Quick Tips:
• Use Demo Data checkbox for testing without email credentials
• Click calendar icons for easy date selection
• Right-click on email rows for context menu
• Double-click email rows to view details
• Use quick date buttons for common ranges

Troubleshooting:
• For Gmail: Enable IMAP and use App Password
• For Outlook: Basic Auth may be disabled
• Green checkmarks indicate valid dates
• Red X marks indicate invalid dates
• Yellow warnings show date range issues
        """
        
        text_widget = tk.Text(frame, wrap=tk.WORD, height=15, width=60,
                             bg=self.colors['surface'], fg=self.colors['text'],
                             insertbackground=self.colors['text'])
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)
        
        ttk.Button(frame, text="Close", command=help_window.destroy).pack(pady=(10, 0))
    
    def show_export_options(self):
        """Show export options dialog."""
        if not self.emails_data:
            self.show_notification("No email data to export. Fetch emails first.", "warning")
            return
        
        export_window = tk.Toplevel(self.root)
        export_window.title("Export Options")
        export_window.geometry("400x300")
        export_window.transient(self.root)
        export_window.grab_set()
        
        try:
            export_window.configure(bg=self.colors['app_bg'])
        except Exception:
            pass
        
        frame = ttk.Frame(export_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Export Email Data", style='Title.TLabel').pack(pady=(0, 20))
        
        # Export format selection
        format_frame = ttk.LabelFrame(frame, text="Export Format", padding="10")
        format_frame.pack(fill=tk.X, pady=(0, 20))
        
        format_var = tk.StringVar(value="csv")
        ttk.Radiobutton(format_frame, text="CSV (Comma Separated Values)", 
                       variable=format_var, value="csv").pack(anchor=tk.W)
        ttk.Radiobutton(format_frame, text="JSON (JavaScript Object Notation)", 
                       variable=format_var, value="json").pack(anchor=tk.W)
        
        # Export options
        options_frame = ttk.LabelFrame(frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 20))
        
        include_body_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Include email body content", 
                       variable=include_body_var).pack(anchor=tk.W)
        
        filtered_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Export filtered results only", 
                       variable=filtered_only_var).pack(anchor=tk.W)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        def do_export():
            try:
                data_to_export = self.filtered_emails if filtered_only_var.get() and hasattr(self, 'filtered_emails') else self.emails_data
                
                if format_var.get() == "csv":
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".csv",
                        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
                    )
                    if filename:
                        self.export_to_csv(data_to_export, filename, include_body_var.get())
                else:
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".json",
                        filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                    )
                    if filename:
                        self.export_to_json(data_to_export, filename, include_body_var.get())
                
                export_window.destroy()
            except Exception as e:
                self.show_notification(f"Export failed: {str(e)}", "error")
        
        ttk.Button(button_frame, text="Export", command=do_export, 
                  style='Success.TButton').pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Cancel", 
                  command=export_window.destroy).pack(side=tk.RIGHT)
    
    def show_settings_dialog(self):
        """Show application settings dialog."""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("450x350")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        try:
            settings_window.configure(bg=self.colors['app_bg'])
        except Exception:
            pass
        
        frame = ttk.Frame(settings_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Application Settings", style='Title.TLabel').pack(pady=(0, 20))
        
        # UI Settings
        ui_frame = ttk.LabelFrame(frame, text="User Interface", padding="10")
        ui_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Auto-refresh option
        auto_refresh_var = tk.BooleanVar(value=getattr(self, 'auto_refresh', False))
        ttk.Checkbutton(ui_frame, text="Auto-refresh dashboard every 30 seconds", 
                       variable=auto_refresh_var).pack(anchor=tk.W)
        
        # Show notifications option
        show_notifications_var = tk.BooleanVar(value=getattr(self, 'show_notifications', True))
        ttk.Checkbutton(ui_frame, text="Show notification messages", 
                       variable=show_notifications_var).pack(anchor=tk.W)
        
        # Email Settings
        email_frame = ttk.LabelFrame(frame, text="Email Processing", padding="10")
        email_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Max emails to fetch
        max_emails_frame = ttk.Frame(email_frame)
        max_emails_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(max_emails_frame, text="Max emails to fetch:").pack(side=tk.LEFT)
        max_emails_var = tk.StringVar(value=str(getattr(self, 'max_emails', 1000)))
        ttk.Entry(max_emails_frame, textvariable=max_emails_var, width=10).pack(side=tk.RIGHT)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        def save_settings():
            try:
                self.auto_refresh = auto_refresh_var.get()
                self.show_notifications = show_notifications_var.get()
                self.max_emails = int(max_emails_var.get())
                self.show_notification("Settings saved successfully", "success")
                settings_window.destroy()
            except Exception as e:
                self.show_notification(f"Failed to save settings: {str(e)}", "error")
        
        ttk.Button(button_frame, text="Save", command=save_settings, 
                  style='Success.TButton').pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Cancel", 
                  command=settings_window.destroy).pack(side=tk.RIGHT)
        
    def fetch_emails(self):
        """Fetch emails based on selected criteria (non-blocking with loading state)."""
        # Validate date inputs
        if not self.start_date_var.get() or not self.end_date_var.get():
            messagebox.showerror("Error", "Please select both start and end dates.")
            return
        try:
            start_date = datetime.strptime(self.start_date_var.get(), '%Y-%m-%d')
            end_date = datetime.strptime(self.end_date_var.get(), '%Y-%m-%d')
        except Exception:
            messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD.")
            return
        if start_date > end_date:
            messagebox.showerror("Error", "Start date cannot be after end date.")
            return

        self._show_loading("Fetching emails...")
        try:
            self.fetch_button.state(['disabled'])
        except Exception:
            pass

        use_demo = self.demo_mode_var.get()

        import threading
        def worker():
            try:
                if use_demo:
                    data = load_demo_emails_between(start_date, end_date)
                    status = f"Demo: Found {len(data)} emails"
                else:
                    data = self.fetch_real_emails(start_date, end_date)
                    status = f"Found {len(data)} emails"
                self.root.after(0, lambda: self._on_fetch_done(data, None, status))
            except Exception as exc:
                self.root.after(0, lambda: self._on_fetch_done([], exc, "Error occurred"))
        threading.Thread(target=worker, daemon=True).start()

    def _on_fetch_done(self, data, error, status_text: str):
        self._hide_loading()
        try:
            self.fetch_button.state(['!disabled'])
        except Exception:
            pass
        if error:
            err = str(error)
            if 'AUTHENTICATIONFAILED' in err.upper():
                provider = self.server_var.get()
                hint = (
                    "For Gmail, enable IMAP and use an App Password.\n"
                    "For Outlook/Office365, Basic Auth may be disabled; try an App Password or use Demo Data."
                )
                self.show_notification("Authentication failed - check credentials", "error", 5000)
                messagebox.showerror("Authentication Failed", f"Login rejected by server.\n\nProvider: {provider}\n\n{err}\n\n{hint}")
            else:
                self.show_notification(f"Fetch failed: {str(error)[:50]}...", "error", 5000)
                messagebox.showerror("Error", f"Failed to fetch emails: {err}")
            self.status_label.config(text="Error occurred")
            return
        self.emails_data = data
        self._build_search_cache()
        self.filtered_emails = list(self.emails_data)
        self.status_label.config(text=status_text)
        if not self.emails_data:
            self.show_notification("No emails found for selected date range", "warning", 4000)
            messagebox.showinfo("No Emails", "No emails found for the selected date range.")
        else:
            self.show_notification(f"Successfully fetched {len(self.emails_data)} emails", "success", 3000)
            messagebox.showinfo("Success", f"Successfully fetched {len(self.emails_data)} emails.")

    def _show_loading(self, text: str):
        try:
            self.status_label.config(text=text)
            self.progress_bar.pack(fill=tk.X, pady=(6, 0))
            self.progress_bar.start(12)
        except Exception:
            pass

    def _hide_loading(self):
        try:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
        except Exception:
            pass
    
    def export_to_csv(self, data: List[Dict], filename: str, include_body: bool = True):
        """Export email data to CSV file."""
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['name', 'email', 'subject', 'date']
                if include_body:
                    fieldnames.append('body')
                
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                for email_data in data:
                    row = {
                        'name': email_data.get('name', ''),
                        'email': email_data.get('email', ''),
                        'subject': email_data.get('subject', ''),
                        'date': email_data.get('date', '').strftime('%Y-%m-%d %H:%M:%S') if isinstance(email_data.get('date'), datetime) else str(email_data.get('date', ''))
                    }
                    if include_body:
                        row['body'] = email_data.get('body', '')
                    writer.writerow(row)
            
            self.show_notification(f"Successfully exported {len(data)} emails to CSV", "success", 3000)
        except Exception as e:
            raise Exception(f"CSV export failed: {str(e)}")
    
    def export_to_json(self, data: List[Dict], filename: str, include_body: bool = True):
        """Export email data to JSON file."""
        try:
            export_data = []
            for email_data in data:
                row = {
                    'name': email_data.get('name', ''),
                    'email': email_data.get('email', ''),
                    'subject': email_data.get('subject', ''),
                    'date': email_data.get('date', '').strftime('%Y-%m-%d %H:%M:%S') if isinstance(email_data.get('date'), datetime) else str(email_data.get('date', ''))
                }
                if include_body:
                    row['body'] = email_data.get('body', '')
                export_data.append(row)
            
            with open(filename, 'w', encoding='utf-8') as jsonfile:
                json.dump(export_data, jsonfile, indent=2, ensure_ascii=False)
            
            self.show_notification(f"Successfully exported {len(data)} emails to JSON", "success", 3000)
        except Exception as e:
            raise Exception(f"JSON export failed: {str(e)}")

    def test_imap_connection(self):
        """Test IMAP login and INBOX access without fetching emails."""
        if self.demo_mode_var.get():
            messagebox.showinfo("Demo Mode Enabled", "Demo mode is enabled. Disable it to test a real connection.")
            return
        try:
            server_config = self.get_server_config()
            email_user = self.email_var.get().strip()
            email_pass = self.password_var.get().strip()
            if not email_user or not email_pass:
                messagebox.showwarning("Missing Credentials", "Enter your email and password (or app password) to test.")
                return
            mail = None
            try:
                mail = imaplib.IMAP4_SSL(server_config['host'], server_config['port'])
                mail.login(email_user, email_pass)
                status, _ = mail.select('INBOX', readonly=True)
            finally:
                try:
                    if mail is not None:
                        try:
                            mail.close()
                        except Exception:
                            pass
                        try:
                            mail.logout()
                        except Exception:
                            pass
                except Exception:
                    pass
            if status == 'OK':
                messagebox.showinfo("Connection Successful", f"Connected to {server_config['host']} and accessed INBOX successfully.")
            else:
                messagebox.showwarning("Partial Success", "Logged in, but could not open INBOX. Check mailbox permissions.")
        except imaplib.IMAP4.error as e:
            messagebox.showerror("Authentication Failed", f"Server rejected login: {e}\n\nIf using Gmail, use an App Password. If using Outlook 365, Basic Auth may be disabled.")
        except Exception as e:
            messagebox.showerror("Connection Error", f"Unable to connect: {e}")
    
    # Removed unused local demo generator; external demo loader is used instead.
    
    def fetch_real_emails(self, start_date: datetime, end_date: datetime) -> List[Dict]:
        """Fetch real emails using IMAP connection with SINCE/BEFORE filtering and robust parsing."""
        # Validate credentials
        email_user = self.email_var.get().strip()
        email_pass = self.password_var.get().strip()
        if not email_user or not email_pass:
            raise Exception("Email and password are required when demo mode is off.")

        mail = None
        try:
            server_config = self.get_server_config()
            mail = imaplib.IMAP4_SSL(server_config['host'], server_config['port'])
            mail.login(email_user, email_pass)
            # Select INBOX read-only to avoid flag changes
            status, _ = mail.select('INBOX', readonly=True)
            if status != 'OK':
                raise Exception('Unable to select INBOX')

            # IMAP date range: SINCE is inclusive, BEFORE is exclusive. Add 1 day to end_date
            since_str = start_date.strftime('%d-%b-%Y')
            before_date = end_date + timedelta(days=1)
            before_str = before_date.strftime('%d-%b-%Y')
            search_criteria = f'(SINCE "{since_str}" BEFORE "{before_str}")'

            status, messages = mail.search(None, search_criteria)
            if status != 'OK':
                raise Exception('IMAP search failed')
            id_bytes = messages[0]
            if not id_bytes:
                return []
            email_ids = id_bytes.split()

            emails_data: List[Dict] = []
            # Fetch messages; avoid downloading entire large mailboxes at once
            for email_id in email_ids:
                status, msg_data = mail.fetch(email_id, '(RFC822)')
                if status != 'OK' or not msg_data or not msg_data[0]:
                    continue
                try:
                    email_message = email.message_from_bytes(msg_data[0][1])
                except Exception:
                    continue

                # Parse and filter by Date header to be safe
                msg_date_raw = email_message.get('Date')
                msg_dt: Optional[datetime] = None
                try:
                    if msg_date_raw:
                        msg_dt = parsedate_to_datetime(msg_date_raw)
                        if msg_dt.tzinfo is not None:
                            msg_dt = msg_dt.astimezone(tz=None).replace(tzinfo=None)
                except Exception:
                    msg_dt = None

                # Skip messages with unparsable date to avoid incorrect range
                if not msg_dt or msg_dt < start_date or msg_dt > end_date:
                    # We rely on server-side search, but double-check here
                    pass

                subject = self.decode_header(email_message.get('Subject'))
                from_header = self.decode_header(email_message.get('From'))
                body = self.extract_email_body(email_message)

                name, email_addr = self.parse_sender(from_header)

                emails_data.append({
                    'name': name,
                    'email': email_addr,
                    'subject': subject,
                    'body': body[:200] + '...' if len(body) > 200 else body,
                    'date': msg_dt or start_date
                })

            # Sort emails by date in descending order (latest first)
            emails_data.sort(key=lambda x: x['date'], reverse=True)
            return emails_data

        except imaplib.IMAP4.error as e:
            raise Exception(f"IMAP authentication/search error: {e}")
        except Exception as e:
            raise Exception(f"IMAP connection failed: {e}")
        finally:
            # Ensure cleanup even when exceptions occur
            try:
                if mail is not None:
                    try:
                        mail.close()
                    except Exception:
                        pass
                    try:
                        mail.logout()
                    except Exception:
                        pass
            except Exception:
                pass
    
    def get_server_config(self) -> Dict:
        """Get IMAP server configuration based on selection."""
        configs = {
            'Gmail': {'host': 'imap.gmail.com', 'port': 993},
            'Outlook': {'host': 'outlook.office365.com', 'port': 993},
            'Custom': {'host': 'imap.example.com', 'port': 993}
        }
        return configs.get(self.server_var.get(), configs['Gmail'])
    
    def decode_header(self, header_value):
        """Decode email header values."""
        if header_value is None:
            return ""
        
        decoded_parts = decode_header(header_value)
        decoded_string = ""
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                if encoding:
                    decoded_string += part.decode(encoding)
                else:
                    decoded_string += part.decode('utf-8', errors='ignore')
            else:
                decoded_string += part
        return decoded_string
    
    def extract_email_body(self, email_message):
        """Extract text body from email message."""
        # Streamlined extraction with short-circuit after reasonable length
        MAX_LEN = 2000
        body_chunks = []
        remaining = MAX_LEN
        try:
            if email_message.is_multipart():
                for part in email_message.walk():
                    if part.get_content_type() == "text/plain":
                        payload = part.get_payload(decode=True)
                        if not payload:
                            continue
                        text = payload.decode('utf-8', errors='ignore')
                        if not text:
                            continue
                        if len(text) >= remaining:
                            body_chunks.append(text[:remaining])
                            remaining = 0
                        else:
                            body_chunks.append(text)
                            remaining -= len(text)
                        if remaining <= 0:
                            break
            else:
                payload = email_message.get_payload(decode=True)
                if payload:
                    text = payload.decode('utf-8', errors='ignore')
                    body_chunks.append(text[:MAX_LEN])
        except Exception:
            pass
        return ''.join(body_chunks)
    
    def parse_sender(self, from_header):
        """Parse sender name and email from header using email.utils.parseaddr."""
        try:
            name, addr = parseaddr(from_header or '')
            name = name.strip().strip('"') if name else addr
            addr = addr or ''
            return name, addr
        except Exception:
            return from_header or '', from_header or ''
    
    def open_dashboard(self):
        """Open the dashboard window."""
        if not self.emails_data:
            messagebox.showwarning("No Data", "Please fetch emails first.")
            return
        
        if self.dashboard_window and self.dashboard_window.winfo_exists():
            self.dashboard_window.lift()
            return
        
        self.dashboard_window = tk.Toplevel(self.root)
        self.dashboard_window.title("Email Dashboard")
        self.dashboard_window.geometry("1200x700")
        self.dashboard_window.configure(bg=self.colors['app_bg'])
        
        self.create_dashboard_screen()
    
    def create_dashboard_screen(self):
        """Create the dashboard interface with email table."""
        # Main frame
        main_frame = ttk.Frame(self.dashboard_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title and controls
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X)
        ttk.Label(header_frame, text="Email Dashboard", style='Title.TLabel').pack(side=tk.LEFT)
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=(12, 12))
        
        # Enhanced search and controls
        controls_frame = ttk.Frame(header_frame)
        controls_frame.pack(side=tk.RIGHT)
        
        # Filter controls
        filter_frame = ttk.Frame(controls_frame)
        filter_frame.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Label(filter_frame, text="Filter by:").pack(side=tk.LEFT, padx=(0, 5))
        self.filter_type_var = tk.StringVar(value="All")
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_type_var, 
                                   values=["All", "Sender", "Subject", "Body"], 
                                   state="readonly", width=10)
        filter_combo.pack(side=tk.LEFT)
        filter_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_search_filter(self.search_var.get()))
        
        # Search box
        search_frame = ttk.Frame(controls_frame)
        search_frame.pack(side=tk.LEFT)
        
        ttk.Label(search_frame, text="🔍").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=25)
        self.search_entry.pack(side=tk.LEFT)
        self.search_entry.bind('<Return>', lambda e: self.apply_search_filter(self.search_var.get()))
        
        # Search controls
        search_controls = ttk.Frame(search_frame)
        search_controls.pack(side=tk.LEFT, padx=(5, 0))
        
        clear_btn = ttk.Button(search_controls, text="✖", width=3, 
                              command=lambda: self.set_search_query(""))
        clear_btn.pack(side=tk.LEFT, padx=(0, 5))
        self._add_tooltip(clear_btn, "Clear search (Escape)")
        
        refresh_btn = ttk.Button(search_controls, text="🔄", width=3,
                               command=self.refresh_dashboard)
        refresh_btn.pack(side=tk.LEFT)
        self._add_tooltip(refresh_btn, "Refresh dashboard (F5)")

        # Live search on typing
        # Debounced live search on typing
        self._search_after_id = None
        def on_search_key(_e=None):
            try:
                if self._search_after_id:
                    self.root.after_cancel(self._search_after_id)
            except Exception:
                pass
            def do_filter():
                self.apply_search_filter(self.search_var.get())
            self._search_after_id = self.root.after(200, do_filter)
        self.search_entry.bind('<KeyRelease>', on_search_key)
        
        # Enhanced statistics frame
        stats_frame = ttk.LabelFrame(main_frame, text="📊 Statistics & Quick Actions", padding="15")
        stats_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Statistics row
        stats_row = ttk.Frame(stats_frame)
        stats_row.pack(fill=tk.X, pady=(0, 10))
        
        self.total_count_var = tk.StringVar()
        self.unique_senders_var = tk.StringVar()
        self.filtered_count_var = tk.StringVar()
        self.range_var = tk.StringVar(value=f"📅 {self.start_date_var.get()} to {self.end_date_var.get()}")
        
        # Statistics labels with icons
        ttk.Label(stats_row, textvariable=self.total_count_var, font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(stats_row, textvariable=self.unique_senders_var, font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(stats_row, textvariable=self.filtered_count_var, font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(stats_row, text="Date Range:", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(stats_row, textvariable=self.range_var).pack(side=tk.LEFT)
        
        # Quick action buttons
        actions_row = ttk.Frame(stats_frame)
        actions_row.pack(fill=tk.X)
        
        ttk.Label(actions_row, text="Quick Actions:", font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 10))
        
        export_selected_btn = ttk.Button(actions_row, text="📤 Export Visible", 
                                        command=lambda: self.show_export_options())
        export_selected_btn.pack(side=tk.LEFT, padx=(0, 5))
        self._add_tooltip(export_selected_btn, "Export currently visible emails")
        
        # Email table
        table_frame = ttk.LabelFrame(main_frame, text="Email List", padding="10")
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create Treeview
        columns = ('Name', 'Email Address', 'Start Date', 'End Date', 'Subject', 'Body')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', style='Custom.Treeview')
        
        # Configure columns
        column_widths = [120, 180, 100, 100, 200, 250]
        for i, (col, width) in enumerate(zip(columns, column_widths)):
            # Make headings clickable for sorting
            self.tree.heading(col, text=col, command=lambda c=col: self._sort_tree_by_column(c, False))
            self.tree.column(col, width=width, minwidth=80)
        
        # Add vertical scrollbar only
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        # Pack widgets
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate table
        # initialize filtered list and populate
        self.filtered_emails = list(self.emails_data)
        self.populate_table()
        
        # Enhanced event bindings
        self.tree.bind('<Double-1>', self.on_email_double_click)
        self.tree.bind('<Button-1>', self.on_tree_click)  # Single click for selection feedback
        self.tree.bind('<Motion>', self.on_tree_motion)   # Mouse hover effects
        
        # Keyboard navigation
        self.tree.bind('<Return>', lambda e: self.on_email_double_click(e))
        self.tree.bind('<Delete>', self.on_delete_key)
        self.tree.bind('<Control-c>', self.on_copy_selection)
        self.tree.bind('<Control-a>', lambda e: self.select_all_emails())
        
        # Context menu on right-click
        self._attach_tree_context_menu()
        
        # Drag and drop setup
        self.setup_drag_drop()
        
        # Tooltip for search entry
        try:
            self._add_tooltip(self.search_entry, "Type to filter by name, email, subject, or body (Ctrl+F)")
        except Exception:
            pass

    def set_search_query(self, query: str):
        """Set the search box text and apply filter."""
        if hasattr(self, 'search_var'):
            self.search_var.set(query)
        # trigger debounced filter rather than immediate heavy filter
        try:
            if hasattr(self, '_search_after_id') and self._search_after_id:
                self.root.after_cancel(self._search_after_id)
        except Exception:
            pass
        self._search_after_id = self.root.after(10, lambda: self.apply_search_filter(query))
    
    def populate_table(self):
        """Populate the table with email data."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add email data
        active_list = getattr(self, 'filtered_emails', self.emails_data)
        # Sort by date descending to show latest first
        active_list = sorted(active_list, key=lambda x: x['date'], reverse=True)
        start_date = self.start_date_var.get()
        end_date = self.end_date_var.get()
        # Configure zebra tags once
        try:
            c = self.colors
            self.tree.tag_configure('oddrow', background=c['surface'])
            self.tree.tag_configure('evenrow', background=c['surface_alt'])
        except Exception:
            pass
        for index, email_data in enumerate(active_list):
            
            row_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert('', 'end', values=(
                email_data['name'],
                email_data['email'],
                start_date,
                end_date,
                email_data['subject'],
                email_data['body'][:100] + '...' if len(email_data['body']) > 100 else email_data['body']
            ), tags=(row_tag,))

        # Update enhanced stats
        total_emails = len(active_list)
        unique_senders = len(set(item['email'] for item in active_list))
        total_in_dataset = len(self.emails_data)
        
        if hasattr(self, 'total_count_var'):
            self.total_count_var.set(f"📧 Email Count : {total_in_dataset}")
        if hasattr(self, 'unique_senders_var'):
            self.unique_senders_var.set(f"👥 UniqueSenders: {unique_senders}")
        if hasattr(self, 'filtered_count_var'):
            if total_emails != total_in_dataset:
                self.filtered_count_var.set(f"🔍 Filtered: {total_emails}")
            else:
                self.filtered_count_var.set("")

    def _sort_tree_by_column(self, col: str, reverse: bool):
        # Obtain items and their column text
        try:
            items = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
            # Attempt numeric/date sort for date columns, fallback to string
            def try_cast(value: str):
                try:
                    return datetime.strptime(value, '%Y-%m-%d')
                except Exception:
                    return value.lower() if isinstance(value, str) else value
            items.sort(key=lambda t: try_cast(t[0]), reverse=reverse)
            for index, (_, k) in enumerate(items):
                self.tree.move(k, '', index)
            # Toggle sort on next click
            self.tree.heading(col, command=lambda: self._sort_tree_by_column(col, not reverse))
        except Exception:
            pass
    
    def refresh_dashboard(self):
        """Refresh the dashboard with current data."""
        if self.dashboard_window and self.dashboard_window.winfo_exists():
            # Reset filter based on current query
            query = self.search_var.get() if hasattr(self, 'search_var') else ''
            self.apply_search_filter(query)
            messagebox.showinfo("Refresh", "Dashboard refreshed successfully.")
    
    def on_email_double_click(self, event):
        """Handle double-click on email row with enhanced email details."""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            values = item['values']
            
            # Find the full email data by matching sender and subject
            full_email_data = None
            active_list = getattr(self, 'filtered_emails', self.emails_data)
            
            for email_data in active_list:
                if (email_data.get('name', '') == values[0] and 
                    email_data.get('email', '') == values[1] and 
                    email_data.get('subject', '') == values[4]):
                    full_email_data = email_data
                    break
            
            # Create enhanced email detail window
            detail_window = tk.Toplevel(self.dashboard_window)
            detail_window.title("📧 Email Details")
            detail_window.geometry("800x600")
            detail_window.transient(self.dashboard_window)
            detail_window.grab_set()
            
            try:
                detail_window.configure(bg=self.colors['app_bg'])
            except Exception:
                pass
            
            # Main container
            main_container = ttk.Frame(detail_window, padding="20")
            main_container.pack(fill=tk.BOTH, expand=True)
            
            # Header with title and close button
            header_frame = ttk.Frame(main_container)
            header_frame.pack(fill=tk.X, pady=(0, 15))
            
            ttk.Label(header_frame, text="📧 Email Details", style='Title.TLabel').pack(side=tk.LEFT)
            
            close_btn = ttk.Button(header_frame, text="✖", width=3, 
                                  command=detail_window.destroy)
            close_btn.pack(side=tk.RIGHT)
            self._add_tooltip(close_btn, "Close window (Escape)")
            
            # Email metadata frame
            meta_frame = ttk.LabelFrame(main_container, text="📋 Email Information", padding="15")
            meta_frame.pack(fill=tk.X, pady=(0, 15))
            
            # Create metadata grid
            meta_grid = ttk.Frame(meta_frame)
            meta_grid.pack(fill=tk.X)
            
            # From field
            from_frame = ttk.Frame(meta_grid)
            from_frame.pack(fill=tk.X, pady=(0, 8))
            ttk.Label(from_frame, text="From:", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT)
            from_text = f"{values[0]} <{values[1]}>"
            from_label = ttk.Label(from_frame, text=from_text)
            from_label.pack(side=tk.LEFT, padx=(10, 0))
            
            # Subject field
            subject_frame = ttk.Frame(meta_grid)
            subject_frame.pack(fill=tk.X, pady=(0, 8))
            ttk.Label(subject_frame, text="Subject:", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT)
            subject_label = ttk.Label(subject_frame, text=values[4])
            subject_label.pack(side=tk.LEFT, padx=(10, 0))
            
            # Date field
            date_frame = ttk.Frame(meta_grid)
            date_frame.pack(fill=tk.X, pady=(0, 8))
            ttk.Label(date_frame, text="Date:", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT)
            
            if full_email_data and 'date' in full_email_data:
                email_date = full_email_data['date']
                if isinstance(email_date, datetime):
                    date_text = email_date.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    date_text = str(email_date)
            else:
                date_text = f"{values[2]} to {values[3]}"
            
            date_label = ttk.Label(date_frame, text=date_text)
            date_label.pack(side=tk.LEFT, padx=(10, 0))
            
            # Body frame with scrollbar
            body_frame = ttk.LabelFrame(main_container, text="📄 Email Body", padding="15")
            body_frame.pack(fill=tk.BOTH, expand=True)
            
            # Create text widget with scrollbar
            text_container = ttk.Frame(body_frame)
            text_container.pack(fill=tk.BOTH, expand=True)
            
            # Get full body text
            if full_email_data:
                full_body = full_email_data.get('body', 'No body content available.')
            else:
                full_body = values[5] if len(values) > 5 else 'No body content available.'
            
            # Themed text widget with scrollbar
            c = self.colors
            text_widget = tk.Text(text_container, wrap=tk.WORD, 
                                 bg=c['surface'], fg=c['text'], 
                                 insertbackground=c['text'],
                                 highlightbackground=c['border'], 
                                 highlightcolor=c['border'],
                                 font=('Segoe UI', 10),
                                 padx=10, pady=10)
            
            # Scrollbar for text widget
            scrollbar = ttk.Scrollbar(text_container, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            # Pack text widget and scrollbar
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Insert full body text
            text_widget.insert(tk.END, full_body)
            text_widget.config(state=tk.DISABLED)
            
            # Action buttons frame
            action_frame = ttk.Frame(main_container)
            action_frame.pack(fill=tk.X, pady=(15, 0))
            
            # Copy body button
            copy_btn = ttk.Button(action_frame, text="📋 Copy Body", 
                                 command=lambda: self.copy_email_body(full_body))
            copy_btn.pack(side=tk.LEFT)
            self._add_tooltip(copy_btn, "Copy email body to clipboard")
            
            # Word count info
            word_count = len(full_body.split()) if full_body else 0
            char_count = len(full_body) if full_body else 0
            
            info_label = ttk.Label(action_frame, 
                                  text=f"📊 {word_count} words, {char_count} characters",
                                  font=('Segoe UI', 9))
            info_label.pack(side=tk.RIGHT)
            
            # Keyboard shortcuts
            detail_window.bind('<Escape>', lambda e: detail_window.destroy())
            detail_window.bind('<Control-c>', lambda e: self.copy_email_body(full_body))
            
            # Focus on the window
            detail_window.focus_set()
    
    def copy_email_body(self, body_text: str):
        """Copy email body text to clipboard."""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(body_text)
            self.show_notification("Email body copied to clipboard", "success", 2000)
        except Exception as e:
            self.show_notification(f"Copy failed: {str(e)}", "error")

    def _attach_tree_context_menu(self):
        menu = tk.Menu(self.tree, tearoff=0)
        def view_details():
            # Reuse double-click behavior
            try:
                self.on_email_double_click(None)
            except Exception:
                pass
        def copy_cell():
            try:
                sel = self.tree.selection()
                if not sel:
                    return
                item_id = sel[0]
                # Copy the subject cell by default
                values = self.tree.item(item_id, 'values')
                cell_text = values[4] if len(values) > 4 else ''
                self.root.clipboard_clear()
                self.root.clipboard_append(cell_text)
            except Exception:
                pass
        def copy_row():
            try:
                sel = self.tree.selection()
                if not sel:
                    return
                values = self.tree.item(sel[0], 'values')
                line = '\t'.join(str(v) for v in values)
                self.root.clipboard_clear()
                self.root.clipboard_append(line)
            except Exception:
                pass
        menu.add_command(label='View Details', command=view_details)
        menu.add_separator()
        menu.add_command(label='Copy Subject', command=copy_cell)
        menu.add_command(label='Copy Row', command=copy_row)

        def show_menu(event):
            try:
                row_id = self.tree.identify_row(event.y)
                if row_id:
                    self.tree.selection_set(row_id)
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                try:
                    menu.grab_release()
                except Exception:
                    pass
        self.tree.bind('<Button-3>', show_menu)

    # Lightweight tooltip implementation
    def _add_tooltip(self, widget, text: str):
        tooltip = tk.Toplevel(widget)
        tooltip.withdraw()
        tooltip.overrideredirect(True)
        try:
            tooltip.configure(bg=self.colors['app_bg'])
        except Exception:
            pass
        label = ttk.Label(tooltip, text=text, padding=6)
        label.pack()
        def enter(_e):
            try:
                x = widget.winfo_rootx() + 10
                y = widget.winfo_rooty() + widget.winfo_height() + 6
                tooltip.geometry(f"+{x}+{y}")
                tooltip.deiconify()
            except Exception:
                pass
        def leave(_e):
            tooltip.withdraw()
        widget.bind('<Enter>', enter)
        widget.bind('<Leave>', leave)

    def _install_tooltips(self, refs: Dict[str, object]):
        if 'fetch' in refs:
            self._add_tooltip(refs['fetch'], 'Fetch emails for the selected date range')
        if 'test' in refs:
            self._add_tooltip(refs['test'], 'Test IMAP connection with provided credentials')
        if 'dash' in refs:
            self._add_tooltip(refs['dash'], 'Open the dashboard to explore results')

    def select_all_emails(self):
        """Select all visible emails in the tree."""
        try:
            # Select all children
            children = self.tree.get_children()
            self.tree.selection_set(children)
            self.show_notification(f"Selected {len(children)} emails", "info", 2000)
        except Exception as e:
            self.show_notification(f"Selection failed: {str(e)}", "error")
    
    def apply_search_filter(self, query: str):
        """Enhanced filter with type-specific searching."""
        q = (query or '').strip().lower()
        filter_type = getattr(self, 'filter_type_var', tk.StringVar(value="All")).get()
        
        # Avoid redundant work if query and filter type haven't changed
        current_filter = f"{q}|{filter_type}"
        if getattr(self, '_last_filter', None) == current_filter:
            return
        self._last_filter = current_filter
        
        if not q:
            self.filtered_emails = list(self.emails_data)
            self.populate_table()
            return
        
        def matches(item: Dict) -> bool:
            if filter_type == "Sender":
                return q in item.get('name', '').lower() or q in item.get('email', '').lower()
            elif filter_type == "Subject":
                return q in item.get('subject', '').lower()
            elif filter_type == "Body":
                return q in item.get('body', '').lower()
            else:  # All
                blob = item.get('_search_blob')
                if blob is None:
                    name = item.get('name', '')
                    email_addr = item.get('email', '')
                    subject = item.get('subject', '')
                    body = item.get('body', '')
                    blob = f"{name}\n{email_addr}\n{subject}\n{body}".lower()

                return q in blob
        
        self.filtered_emails = [it for it in self.emails_data if matches(it)]
        self.populate_table()
        
        # Show search results notification
        if self.filtered_emails:
            self.show_notification(f"Found {len(self.filtered_emails)} matching emails", "success", 2000)
        else:
            self.show_notification("No emails match your search", "warning", 3000)

    def on_tree_click(self, event):
        """Handle single click on tree for visual feedback."""
        try:
            item = self.tree.selection()[0] if self.tree.selection() else None
            if item:
                # Visual feedback for selection
                self.tree.focus(item)
        except Exception:
            pass
    
    def on_tree_motion(self, event):
        """Handle mouse motion for hover effects."""
        try:
            item = self.tree.identify_row(event.y)
            if item and item != getattr(self, '_last_hover_item', None):
                # Could add hover effects here if needed
                self._last_hover_item = item
        except Exception:
            pass
    
    def on_delete_key(self, event):
        """Handle delete key press."""
        try:
            selection = self.tree.selection()
            if selection:
                self.show_notification(f"Selected {len(selection)} emails (Delete key pressed)", "info", 2000)
        except Exception:
            pass
    
    def on_copy_selection(self, event):
        """Handle Ctrl+C to copy selected email info."""
        try:
            selection = self.tree.selection()
            if not selection:
                return
            
            # Copy first selected item's subject
            item_id = selection[0]
            values = self.tree.item(item_id, 'values')
            if values and len(values) > 4:
                subject = values[4]
                self.root.clipboard_clear()
                self.root.clipboard_append(subject)
                self.show_notification("Email subject copied to clipboard", "success", 2000)
        except Exception as e:
            self.show_notification(f"Copy failed: {str(e)}", "error")
    
    def setup_drag_drop(self):
        """Set up drag and drop functionality for the tree."""
        try:
            # Enable drag detection
            self.tree.bind('<ButtonPress-1>', self.on_drag_start)
            self.tree.bind('<B1-Motion>', self.on_drag_motion)
            self.tree.bind('<ButtonRelease-1>', self.on_drag_end)
            
            # Drag state
            self._drag_data = {'item': None, 'start_y': 0}
        except Exception:
            pass
    
    def on_drag_start(self, event):
        """Start drag operation."""
        try:
            item = self.tree.identify_row(event.y)
            if item:
                self._drag_data['item'] = item
                self._drag_data['start_y'] = event.y
        except Exception:
            pass
    
    def on_drag_motion(self, event):
        """Handle drag motion."""
        try:
            if self._drag_data['item'] and abs(event.y - self._drag_data['start_y']) > 10:
                # Visual feedback for dragging
                self.tree.configure(cursor='hand2')
        except Exception:
            pass
    
    def on_drag_end(self, event):
        """End drag operation."""
        try:
            self.tree.configure(cursor='')
            if self._drag_data['item']:
                # Could implement reordering or other drag operations here
                self.show_notification("Drag operation completed", "info", 1500)
            self._drag_data = {'item': None, 'start_y': 0}
        except Exception:
            pass

    def _build_search_cache(self):
        """Precompute lowercase search blobs for faster incremental filtering."""
        for item in self.emails_data:
            try:
                name = item.get('name', '')
                email_addr = item.get('email', '')
                subject = item.get('subject', '')
                body = item.get('body', '')
                item['_search_blob'] = f"{name}\n{email_addr}\n{subject}\n{body}".lower()
            except Exception:
                item['_search_blob'] = ''
    
    def run(self):
        """Start the application."""
        self.root.mainloop()


def main():
    """Main function to run the application."""
    app = EmailFilterApp()
    app.run()


if __name__ == "__main__":
    main()
