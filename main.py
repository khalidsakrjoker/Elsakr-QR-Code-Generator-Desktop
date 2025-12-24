"""
Elsakr QR Code Generator - Desktop Version
Generate QR codes with custom colors, logos, and batch processing.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import qrcode
from PIL import Image, ImageTk, ImageDraw
import os
import pyperclip
from io import BytesIO

class ElsakrQRGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Elsakr QR Code Generator")
        self.root.geometry("1300x1200")
        self.root.minsize(800, 600)
        
        # Dark theme colors
        self.colors = {
            'bg_primary': '#0a0a0f',
            'bg_secondary': '#12121a',
            'bg_tertiary': '#1a1a25',
            'accent': '#8b5cf6',
            'accent_hover': '#7c3aed',
            'text_primary': '#ffffff',
            'text_secondary': '#a0a0b0',
            'border': '#2a2a3a'
        }
        
        # Configure root window
        self.root.configure(bg=self.colors['bg_primary'])
        
        # Set window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'fav.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass
        
        # Variables
        self.qr_type = tk.StringVar(value='url')
        self.fg_color = '#000000'
        self.bg_color = '#FFFFFF'
        self.logo_image = None
        self.current_qr_image = None
        
        # Configure styles
        self.configure_styles()
        
        # Build UI
        self.create_widgets()
        
        # Generate initial QR
        self.generate_qr()
    
    def configure_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure common styles
        style.configure('TFrame', background=self.colors['bg_primary'])
        style.configure('Card.TFrame', background=self.colors['bg_secondary'])
        style.configure('TLabel', 
                       background=self.colors['bg_primary'], 
                       foreground=self.colors['text_primary'],
                       font=('Segoe UI', 10))
        style.configure('Header.TLabel',
                       font=('Segoe UI', 24, 'bold'),
                       foreground=self.colors['text_primary'])
        style.configure('Subheader.TLabel',
                       font=('Segoe UI', 11),
                       foreground=self.colors['text_secondary'])
        
        # Button styles
        style.configure('Accent.TButton',
                       background=self.colors['accent'],
                       foreground='white',
                       font=('Segoe UI', 11, 'bold'),
                       padding=(20, 12))
        style.map('Accent.TButton',
                 background=[('active', self.colors['accent_hover'])])
        
        style.configure('Secondary.TButton',
                       background=self.colors['bg_tertiary'],
                       foreground=self.colors['text_primary'],
                       font=('Segoe UI', 10),
                       padding=(15, 10))
        
        # Entry style
        style.configure('TEntry',
                       fieldbackground=self.colors['bg_tertiary'],
                       foreground=self.colors['text_primary'],
                       insertcolor=self.colors['text_primary'])
        
        # Radiobutton style
        style.configure('Tab.TRadiobutton',
                       background=self.colors['bg_tertiary'],
                       foreground=self.colors['text_primary'],
                       font=('Segoe UI', 10),
                       padding=(12, 8))
        style.map('Tab.TRadiobutton',
                 background=[('selected', self.colors['accent'])])
    
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        header_frame = ttk.Frame(main_frame, style='TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Load and display logo
        try:
            logo_path = os.path.join(os.path.dirname(__file__), 'assets', 'Sakr-logo.png')
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((50, 50), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                logo_label = ttk.Label(header_frame, image=self.logo_photo, background=self.colors['bg_primary'])
                logo_label.pack(side=tk.LEFT, padx=(0, 15))
        except:
            pass
        
        title_frame = ttk.Frame(header_frame, style='TFrame')
        title_frame.pack(side=tk.LEFT)
        
        ttk.Label(title_frame, text="QR Code Generator", style='Header.TLabel').pack(anchor='w')
        ttk.Label(title_frame, text="Create beautiful QR codes with custom colors and logos", 
                 style='Subheader.TLabel').pack(anchor='w')
        
        # Content area (two columns)
        content_frame = ttk.Frame(main_frame, style='TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left panel - Controls
        left_panel = tk.Frame(content_frame, bg=self.colors['bg_secondary'], 
                             highlightthickness=1, highlightbackground=self.colors['border'])
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        left_inner = ttk.Frame(left_panel, style='Card.TFrame')
        left_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # QR Type Selection
        ttk.Label(left_inner, text="QR Code Type", style='TLabel',
                 background=self.colors['bg_secondary']).pack(anchor='w', pady=(0, 10))
        
        type_frame = ttk.Frame(left_inner, style='Card.TFrame')
        type_frame.pack(fill=tk.X, pady=(0, 20))
        
        qr_types = [
            ('🔗 URL', 'url'),
            ('📝 Text', 'text'),
            ('📶 WiFi', 'wifi'),
            (' Email', 'email'),
            (' SMS', 'sms')
        ]
        
        for i, (label, value) in enumerate(qr_types):
            rb = tk.Radiobutton(type_frame, text=label, variable=self.qr_type, value=value,
                               bg=self.colors['bg_tertiary'], fg=self.colors['text_primary'],
                               selectcolor=self.colors['accent'], activebackground=self.colors['bg_tertiary'],
                               activeforeground=self.colors['text_primary'],
                               font=('Segoe UI', 10), padx=10, pady=5,
                               command=self.on_type_change)
            rb.pack(side=tk.LEFT, padx=2)
        
        # Input fields container
        self.input_container = ttk.Frame(left_inner, style='Card.TFrame')
        self.input_container.pack(fill=tk.X, pady=(0, 20))
        
        self.create_input_fields()
        
        # Color selection
        color_frame = ttk.Frame(left_inner, style='Card.TFrame')
        color_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Foreground color
        fg_frame = ttk.Frame(color_frame, style='Card.TFrame')
        fg_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))
        
        ttk.Label(fg_frame, text="QR Color", style='TLabel',
                 background=self.colors['bg_secondary']).pack(anchor='w')
        
        self.fg_btn = tk.Button(fg_frame, bg=self.fg_color, width=6, height=2,
                               command=lambda: self.choose_color('fg'),
                               relief=tk.FLAT, cursor='hand2')
        self.fg_btn.pack(anchor='w', pady=5)
        
        # Background color
        bg_frame = ttk.Frame(color_frame, style='Card.TFrame')
        bg_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        ttk.Label(bg_frame, text="Background", style='TLabel',
                 background=self.colors['bg_secondary']).pack(anchor='w')
        
        self.bg_btn = tk.Button(bg_frame, bg=self.bg_color, width=6, height=2,
                               command=lambda: self.choose_color('bg'),
                               relief=tk.FLAT, cursor='hand2')
        self.bg_btn.pack(anchor='w', pady=5)

        # Reset Colors Button
        reset_btn = tk.Button(color_frame, text="⏪ Reset Colors", bg=self.colors['bg_tertiary'],
                              fg=self.colors['text_primary'], font=('Segoe UI', 9),
                              command=self.reset_colors, relief=tk.FLAT, padx=10, pady=5,
                              cursor='hand2')
        reset_btn.pack(fill=tk.X, pady=(10, 5))

        # Contrast Warning
        ttk.Label(color_frame, text="⚠️ Ensure high contrast (Dark on Light) for best results!", 
                 style='Subheader.TLabel', font=('Segoe UI', 9), foreground='#ef4444').pack(fill=tk.X, pady=(5, 0))
        
        # Logo upload
        logo_frame = ttk.Frame(left_inner, style='Card.TFrame')
        logo_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(logo_frame, text="Logo (optional)", style='TLabel',
                 background=self.colors['bg_secondary']).pack(anchor='w')
        
        logo_btn_frame = ttk.Frame(logo_frame, style='Card.TFrame')
        logo_btn_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(logo_btn_frame, text="📷 Upload Logo", bg=self.colors['bg_tertiary'],
                 fg=self.colors['text_primary'], font=('Segoe UI', 10),
                 command=self.upload_logo, relief=tk.FLAT, padx=15, pady=8,
                 cursor='hand2').pack(side=tk.LEFT)
        
        self.logo_label = ttk.Label(logo_btn_frame, text="No logo selected", 
                                   style='Subheader.TLabel', background=self.colors['bg_secondary'])
        self.logo_label.pack(side=tk.LEFT, padx=10)
        
        tk.Button(logo_btn_frame, text="✕", bg=self.colors['bg_tertiary'],
                 fg='#ef4444', font=('Segoe UI', 10),
                 command=self.remove_logo, relief=tk.FLAT, padx=10, pady=8,
                 cursor='hand2').pack(side=tk.LEFT)
        
        # Batch import
        batch_frame = ttk.Frame(left_inner, style='Card.TFrame')
        batch_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Button(batch_frame, text="📄 Batch Import (TXT)", bg=self.colors['bg_tertiary'],
                 fg=self.colors['text_primary'], font=('Segoe UI', 10),
                 command=self.batch_import, relief=tk.FLAT, padx=15, pady=8,
                 cursor='hand2').pack(side=tk.LEFT)
        
        # Generate button
        generate_btn = tk.Button(left_inner, text="⚡ Generate QR Code",
                                bg=self.colors['accent'], fg='white',
                                font=('Segoe UI', 12, 'bold'),
                                command=self.generate_qr, relief=tk.FLAT,
                                padx=20, pady=12, cursor='hand2')
        generate_btn.pack(fill=tk.X, pady=(10, 0))
        
        # Right panel - Preview
        right_panel = tk.Frame(content_frame, bg=self.colors['bg_secondary'],
                              highlightthickness=1, highlightbackground=self.colors['border'])
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        right_inner = ttk.Frame(right_panel, style='Card.TFrame')
        right_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        ttk.Label(right_inner, text="Preview", style='TLabel',
                 background=self.colors['bg_secondary']).pack(anchor='w', pady=(0, 10))
        
        # QR Preview canvas
        preview_frame = tk.Frame(right_inner, bg='white', width=280, height=280)
        preview_frame.pack(pady=10)
        preview_frame.pack_propagate(False)
        
        self.qr_label = ttk.Label(preview_frame, background='white')
        self.qr_label.pack(expand=True)
        
        # Download buttons
        btn_frame = ttk.Frame(right_inner, style='Card.TFrame')
        btn_frame.pack(fill=tk.X, pady=(20, 0))
        
        tk.Button(btn_frame, text="📥 Save PNG", bg=self.colors['bg_tertiary'],
                 fg=self.colors['text_primary'], font=('Segoe UI', 10),
                 command=lambda: self.save_qr('png'), relief=tk.FLAT,
                 padx=15, pady=10, cursor='hand2').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        
        tk.Button(btn_frame, text="📐 Save SVG", bg=self.colors['bg_tertiary'],
                 fg=self.colors['text_primary'], font=('Segoe UI', 10),
                 command=lambda: self.save_qr('svg'), relief=tk.FLAT,
                 padx=15, pady=10, cursor='hand2').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))
        
        # Copy button
        tk.Button(right_inner, text="📋 Copy to Clipboard", bg=self.colors['bg_tertiary'],
                 fg=self.colors['text_primary'], font=('Segoe UI', 10),
                 command=self.copy_to_clipboard, relief=tk.FLAT,
                 padx=15, pady=10, cursor='hand2').pack(fill=tk.X, pady=(10, 0))
    
    def create_input_fields(self):
        # Clear existing fields
        for widget in self.input_container.winfo_children():
            widget.destroy()
        
        qr_type = self.qr_type.get()
        
        if qr_type == 'url':
            self.create_entry("Website URL", "url_entry", "https://elsakr.company")
        
        elif qr_type == 'text':
            self.create_text_area("Text Content", "text_area", "Hello from Elsakr! Test QR Code.")
        
        elif qr_type == 'wifi':
            self.create_entry("Network Name (SSID)", "wifi_ssid", "ElsakrWiFi")
            self.create_entry("Password", "wifi_password", "Elsakr2024")
            self.create_dropdown("Encryption", "wifi_encryption", ["WPA/WPA2", "WEP", "None"])
        
        elif qr_type == 'vcard':
            self.create_entry("First Name", "vcard_firstname", "Khalid")
            self.create_entry("Last Name", "vcard_lastname", "Sakr")
            self.create_entry("Phone", "vcard_phone", "+201016495229")
            self.create_entry("Email", "vcard_email", "hello@elsakr.company")
            self.create_entry("Company", "vcard_company", "Elsakr Software House")
        
        elif qr_type == 'email':
            self.create_entry("Email Address", "email_address", "hello@elsakr.company")
            self.create_entry("Subject (optional)", "email_subject", "Hello from QR Code")
            self.create_entry("Body (optional)", "email_body", "This email was generated by Elsakr QR!")
        
        elif qr_type == 'phone':
            self.create_entry("Phone Number", "phone_number", "+201016495229")
        
        elif qr_type == 'sms':
            self.create_entry("Phone Number", "sms_phone", "+201016495229")
            self.create_entry("Message (optional)", "sms_message", "Hello from Elsakr QR!")
    
    def create_entry(self, label, name, default=""):
        frame = ttk.Frame(self.input_container, style='Card.TFrame')
        frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(frame, text=label, style='TLabel',
                 background=self.colors['bg_secondary']).pack(anchor='w')
        
        entry = tk.Entry(frame, bg=self.colors['bg_tertiary'], fg=self.colors['text_primary'],
                        font=('Segoe UI', 11), insertbackground=self.colors['text_primary'],
                        relief=tk.FLAT, bd=0)
        entry.insert(0, default)
        entry.pack(fill=tk.X, pady=5, ipady=8, padx=2)
        
        setattr(self, name, entry)
    
    def create_text_area(self, label, name, default=""):
        frame = ttk.Frame(self.input_container, style='Card.TFrame')
        frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(frame, text=label, style='TLabel',
                 background=self.colors['bg_secondary']).pack(anchor='w')
        
        text = tk.Text(frame, bg=self.colors['bg_tertiary'], fg=self.colors['text_primary'],
                      font=('Segoe UI', 11), insertbackground=self.colors['text_primary'],
                      relief=tk.FLAT, height=4, wrap=tk.WORD)
        text.pack(fill=tk.X, pady=5)
        if default:
            text.insert("1.0", default)
        
        setattr(self, name, text)
    
    def create_dropdown(self, label, name, options):
        frame = ttk.Frame(self.input_container, style='Card.TFrame')
        frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(frame, text=label, style='TLabel',
                 background=self.colors['bg_secondary']).pack(anchor='w')
        
        var = tk.StringVar(value=options[0])
        dropdown = ttk.Combobox(frame, textvariable=var, values=options, state='readonly')
        dropdown.pack(fill=tk.X, pady=5)
        
        setattr(self, name, var)
    
    def on_type_change(self):
        self.create_input_fields()
    
    def choose_color(self, color_type):
        color = colorchooser.askcolor(title=f"Choose {'QR' if color_type == 'fg' else 'Background'} Color")
        if color[1]:
            if color_type == 'fg':
                self.fg_color = color[1]
                self.fg_btn.configure(bg=self.fg_color)
            else:
                self.bg_color = color[1]
                self.bg_btn.configure(bg=self.bg_color)
    
    def reset_colors(self):
        self.fg_color = '#000000'
        self.bg_color = '#FFFFFF'
        self.fg_btn.configure(bg=self.fg_color)
        self.bg_btn.configure(bg=self.bg_color)
        messagebox.showinfo("Colors Reset", "Colors have been reset to default Black & White.")

    def create_input_fields(self):
        # Clear existing fields
        for widget in self.input_container.winfo_children():
            widget.destroy()
        
        qr_type = self.qr_type.get()
        
        if qr_type == 'url':
            self.create_entry("Website URL", "url_entry", "https://elsakr.company")
        
        elif qr_type == 'text':
            self.create_text_area("Text Content", "text_area", "Hello from Elsakr! Test QR Code.")
        
        elif qr_type == 'wifi':
            self.create_entry("Network Name (SSID)", "wifi_ssid", "ElsakrWiFi")
            self.create_entry("Password", "wifi_password", "Elsakr2024")
            self.create_dropdown("Encryption", "wifi_encryption", ["WPA/WPA2", "WEP", "None"])
        
        elif qr_type == 'email':
            self.create_entry("Email Address", "email_address", "hello@elsakr.company")
            self.create_entry("Subject (optional)", "email_subject", "Hello from QR Code")
            self.create_entry("Body (optional)", "email_body", "This email was generated by Elsakr QR!")
        
        elif qr_type == 'sms':
            self.create_entry("Phone Number", "sms_phone", "+201016495229")
            self.create_entry("Message (optional)", "sms_message", "Hello from Elsakr QR!")
    
    def remove_logo(self):
        self.logo_image = None
        self.logo_label.configure(text="No logo selected")
    
    def upload_logo(self):
        file_path = filedialog.askopenfilename(
            title="Select Logo",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp")]
        )
        if file_path:
            try:
                self.logo_image = Image.open(file_path)
                self.logo_label.configure(text=os.path.basename(file_path))
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load logo: {e}")

    def safe_get(self, attr_name, is_text=False):
        """Safely get value from a widget attribute, returning empty string if widget is destroyed"""
        try:
            widget = getattr(self, attr_name, None)
            if widget is None:
                return ''
            if is_text:
                val = widget.get("1.0", tk.END).strip()
            else:
                val = widget.get()
            return val
        except Exception:
            return ''
    
    def get_qr_data(self):
        qr_type = self.qr_type.get()
        
        if qr_type == 'url':
            return self.safe_get('url_entry') or 'https://elsakr.company'
        
        elif qr_type == 'text':
            return self.safe_get('text_area', is_text=True) or 'Hello from Elsakr!'
        
        elif qr_type == 'wifi':
            ssid = self.safe_get('wifi_ssid')
            password = self.safe_get('wifi_password')
            enc_val = self.safe_get('wifi_encryption') or 'WPA/WPA2'
            enc_map = {'WPA/WPA2': 'WPA', 'WEP': 'WEP', 'None': 'nopass'}
            return f"WIFI:T:{enc_map.get(enc_val, 'WPA')};S:{ssid};P:{password};;"
        
        elif qr_type == 'email':
            email_addr = self.safe_get('email_address')
            subject = self.safe_get('email_subject')
            body = self.safe_get('email_body')
            url = f"mailto:{email_addr}"
            params = []
            if subject:
                params.append(f"subject={subject}")
            if body:
                params.append(f"body={body}")
            if params:
                url += "?" + "&".join(params)
            return url
        
        elif qr_type == 'sms':
            sms_phone = self.safe_get('sms_phone')
            sms_msg = self.safe_get('sms_message')
            if sms_msg:
                return f"sms:{sms_phone}?body={sms_msg}"
            return f"sms:{sms_phone}"
        
        return 'https://elsakr.company'
    
    def generate_qr(self, data=None):
        if data is None:
            data = self.get_qr_data()
        
        # Create QR code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=2
        )
        qr.add_data(data)
        qr.make(fit=True)
        
        # Create black and white QR image first
        qr_image = qr.make_image(fill_color='black', back_color='white')
        
        # Convert to PIL Image
        if hasattr(qr_image, 'get_image'):
            qr_image = qr_image.get_image()
        
        # Convert to RGB
        qr_image = qr_image.convert('RGB')
        
        # Manually replace colors using PIL
        # Get pixel data
        pixels = qr_image.load()
        width, height = qr_image.size
        
        # Parse hex colors to RGB tuples
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        fg_rgb = hex_to_rgb(self.fg_color)
        bg_rgb = hex_to_rgb(self.bg_color)
        
        # Replace black with foreground color, white with background color
        for y in range(height):
            for x in range(width):
                r, g, b = pixels[x, y]
                # If pixel is dark (black), use foreground color
                if r < 128 and g < 128 and b < 128:
                    pixels[x, y] = fg_rgb
                else:
                    # If pixel is light (white), use background color
                    pixels[x, y] = bg_rgb
        
        # Add logo if present
        if self.logo_image:
            logo = self.logo_image.copy()
            logo_size = int(qr_image.size[0] * 0.25)
            logo = logo.resize((logo_size, logo_size), Image.Resampling.LANCZOS)
            
            # Create background for logo matching QR bg color
            bg = Image.new('RGB', (logo_size + 10, logo_size + 10), bg_rgb)
            pos = ((qr_image.size[0] - logo_size - 10) // 2, (qr_image.size[1] - logo_size - 10) // 2)
            qr_image.paste(bg, pos)
            
            # Paste logo
            logo_pos = ((qr_image.size[0] - logo_size) // 2, (qr_image.size[1] - logo_size) // 2)
            if logo.mode == 'RGBA':
                qr_image.paste(logo.convert('RGB'), logo_pos)
            else:
                qr_image.paste(logo, logo_pos)
        
        self.current_qr_image = qr_image
        
        # Display in preview
        display_size = 250
        display_img = qr_image.resize((display_size, display_size), Image.Resampling.LANCZOS)
        self.qr_photo = ImageTk.PhotoImage(display_img)
        self.qr_label.configure(image=self.qr_photo)
    
    def save_qr(self, format_type):
        if self.current_qr_image is None:
            messagebox.showwarning("Warning", "Generate a QR code first!")
            return
        
        if format_type == 'png':
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png")],
                initialname="elsakr-qrcode.png"
            )
            if file_path:
                self.current_qr_image.save(file_path, 'PNG')
                messagebox.showinfo("Success", f"QR code saved to:\n{file_path}")
        
        elif format_type == 'svg':
            # For SVG, regenerate using qrcode library
            data = self.get_qr_data()
            file_path = filedialog.asksaveasfilename(
                defaultextension=".svg",
                filetypes=[("SVG files", "*.svg")],
                initialname="elsakr-qrcode.svg"
            )
            if file_path:
                import qrcode.image.svg
                factory = qrcode.image.svg.SvgPathImage
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_H,
                    box_size=10,
                    border=2
                )
                qr.add_data(data)
                qr.make(fit=True)
                img = qr.make_image(image_factory=factory)
                img.save(file_path)
                messagebox.showinfo("Success", f"QR code saved to:\n{file_path}")
    
    def copy_to_clipboard(self):
        if self.current_qr_image is None:
            messagebox.showwarning("Warning", "Generate a QR code first!")
            return
        
        # Save to bytes
        output = BytesIO()
        self.current_qr_image.save(output, format='PNG')
        
        # For Windows, use win32clipboard
        try:
            import win32clipboard
            from io import BytesIO
            
            output = BytesIO()
            self.current_qr_image.convert('RGB').save(output, 'BMP')
            data = output.getvalue()[14:]  # Remove BMP header
            output.close()
            
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            
            messagebox.showinfo("Success", "QR code copied to clipboard!")
        except ImportError:
            messagebox.showinfo("Info", "Install pywin32 for clipboard support.\nUse 'Save PNG' instead.")
    
    def batch_import(self):
        file_path = filedialog.askopenfilename(
            title="Select Text File",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv")]
        )
        if not file_path:
            return
        
        # Ask for output folder
        output_folder = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder:
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip()]
            
            count = 0
            for i, line in enumerate(lines):
                self.generate_qr(data=line)
                if self.current_qr_image:
                    output_path = os.path.join(output_folder, f"qr_{i+1:04d}.png")
                    self.current_qr_image.save(output_path, 'PNG')
                    count += 1
            
            messagebox.showinfo("Batch Complete", f"Generated {count} QR codes in:\n{output_folder}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Batch processing failed:\n{e}")


def main():
    root = tk.Tk()
    app = ElsakrQRGenerator(root)
    root.mainloop()


if __name__ == "__main__":
    main()
