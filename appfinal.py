import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
import os

class CertificateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate Generator - Click to Place Fields")
        self.root.geometry("1400x900") # Increased window size
        
        self.template_path = None
        self.excel_path = None
        self.template_image = None
        self.display_image = None
        self.fields = {}
        self.excel_data = None
        self.current_column = None
        self.scale_factor = 1.0
        
        self.setup_ui()

        # Attempt to load and display the last used template if available (e.g., from a config file)
        # For now, we'll keep it simple and just show a blank canvas until a template is loaded.
        # In a real application, you'd save/load self.template_path.
        self.display_template()
    
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title = ttk.Label(main_frame, text="Certificate Generator", 
                         font=('Arial', 20, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Left side - Controls
        left_frame = ttk.Frame(main_frame, padding="10")
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        # Template Upload Section
        ttk.Label(left_frame, text="Step 1: Upload Template", 
                 font=('Arial', 12, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        ttk.Button(left_frame, text="Browse Template Image", 
                  command=self.load_template).grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.template_label = ttk.Label(left_frame, text="No template selected", 
                                        wraplength=250)
        self.template_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        # Excel Upload Section
        ttk.Separator(left_frame, orient='horizontal').grid(row=3, column=0, 
                                                            sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(left_frame, text="Step 2: Upload Excel Data", 
                 font=('Arial', 12, 'bold')).grid(row=4, column=0, sticky=tk.W, pady=5)
        
        ttk.Button(left_frame, text="Browse Excel File", 
                  command=self.load_excel).grid(row=5, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.excel_label = ttk.Label(left_frame, text="No Excel file selected", 
                                     wraplength=250)
        self.excel_label.grid(row=6, column=0, sticky=tk.W, pady=5)
        
        # Fields Selection Section
        ttk.Separator(left_frame, orient='horizontal').grid(row=7, column=0, 
                                                            sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(left_frame, text="Step 3: Select Field Positions", 
                 font=('Arial', 12, 'bold')).grid(row=8, column=0, sticky=tk.W, pady=5)
        
        instruction_text = "Click on the template to place each field:"
        ttk.Label(left_frame, text=instruction_text, 
                 wraplength=250).grid(row=9, column=0, sticky=tk.W, pady=5)
        
        # Scrollable fields list
        fields_canvas = tk.Canvas(left_frame, height=200)
        fields_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", 
                                        command=fields_canvas.yview)
        self.fields_list_frame = ttk.Frame(fields_canvas)
        
        self.fields_list_frame.bind(
            "<Configure>",
            lambda e: fields_canvas.configure(scrollregion=fields_canvas.bbox("all"))
        )
        
        fields_canvas.create_window((0, 0), window=self.fields_list_frame, anchor="nw")
        fields_canvas.configure(yscrollcommand=fields_scrollbar.set)
        
        fields_canvas.grid(row=10, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        fields_scrollbar.grid(row=10, column=1, sticky=(tk.N, tk.S))
        
        # Status label
        self.status_label = ttk.Label(left_frame, text="", foreground="blue", 
                                     wraplength=250)
        self.status_label.grid(row=11, column=0, sticky=tk.W, pady=5)
        
        # Generate buttons
        ttk.Separator(left_frame, orient='horizontal').grid(row=12, column=0, 
                                                            sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(left_frame, text="Step 4: Generate", 
                 font=('Arial', 12, 'bold')).grid(row=13, column=0, sticky=tk.W, pady=5)
        
        ttk.Button(left_frame, text="Preview First Certificate", 
                  command=self.preview_certificate).grid(row=14, column=0, 
                                                         sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(left_frame, text="Generate All Certificates", 
                  command=self.generate_certificates, 
                  style='Accent.TButton').grid(row=15, column=0, 
                                              sticky=(tk.W, tk.E), pady=5)
        
        # Right side - Template Canvas
        right_frame = ttk.Frame(main_frame, padding="10")
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        ttk.Label(right_frame, text="Certificate Template (Click to place fields)", 
                 font=('Arial', 11)).grid(row=0, column=0, pady=5)
        
        # Canvas for template
        self.canvas = tk.Canvas(right_frame, width=800, height=600, 
                               bg='lightgray', cursor='cross', scrollregion=(0,0,800,600))
        self.canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.h_scrollbar = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.v_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set, yscrollcommand=self.v_scrollbar.set)
        
        self.h_scrollbar.grid(row=2, column=0, sticky=(tk.W, tk.E))
        self.v_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))

        self.canvas.bind('<Button-1>', self.on_canvas_click)
        self.canvas.bind('<Configure>', self.on_canvas_resize)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        left_frame.rowconfigure(10, weight=1)
        right_frame.rowconfigure(1, weight=1)
        right_frame.columnconfigure(0, weight=1)
    
    def on_canvas_resize(self, event):
        self.display_template()
    
    def load_template(self):
        filepath = filedialog.askopenfilename(
            title="Select Certificate Template",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp")]
        )
        if filepath:
            self.template_path = filepath
            self.template_label.config(text=f"✓ {os.path.basename(filepath)}")
            self.template_image = Image.open(filepath)
            self.display_template()
            messagebox.showinfo("Success", "Template loaded successfully!")
    
    def display_template(self):
        self.canvas.delete('all') # Clear existing image and markers

        if not self.template_image:
            # Display a placeholder if no template is loaded
            self.canvas.create_text(self.canvas.winfo_width() / 2, self.canvas.winfo_height() / 2,
                                   text="Upload a certificate template image to begin",
                                   font=('Arial', 16), fill='gray')
            return
        
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        if canvas_width == 0 or canvas_height == 0:
            # Canvas not yet rendered, defer display
            self.root.after(100, self.display_template)
            return
        
        img_width, img_height = self.template_image.size
        
        scale_w = canvas_width / img_width
        scale_h = canvas_height / img_height
        
        # Use a maximum scale of 1.0 to prevent pixelation if the image is smaller than the canvas
        self.scale_factor = min(scale_w, scale_h, 1.0) 
        
        new_width = int(img_width * self.scale_factor)
        new_height = int(img_height * self.scale_factor)
        
        if new_width == 0 or new_height == 0:
            # Avoid issues with very small or zero-sized images after scaling
            return

        self.display_image = self.template_image.resize((new_width, new_height), 
                                                        Image.Resampling.LANCZOS)
        self.photo_image = ImageTk.PhotoImage(self.display_image)
        
        # Calculate position to center the image on the canvas
        x_center = (canvas_width - new_width) // 2
        y_center = (canvas_height - new_height) // 2
        
        self.canvas.create_image(x_center, y_center, anchor=tk.NW, image=self.photo_image)
        
        # Update scroll region to match the image size (if image is larger than canvas)
        self.canvas.config(scrollregion=(0, 0, new_width, new_height))

        # Redraw existing markers
        self.redraw_markers()
    
    def load_excel(self):
        filepath = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filepath:
            self.excel_path = filepath
            try:
                self.excel_data = pd.read_excel(filepath)
                self.excel_label.config(text=f"✓ {os.path.basename(filepath)}")
                
                # Clear previous fields
                for widget in self.fields_list_frame.winfo_children():
                    widget.destroy()
                self.fields = {}
                
                # Create button for each column
                for idx, column in enumerate(self.excel_data.columns):
                    self.create_field_button(column, idx)
                
                messagebox.showinfo("Success", 
                    f"Excel loaded!\n{len(self.excel_data)} rows found\n"
                    f"{len(self.excel_data.columns)} columns: {', '.join(self.excel_data.columns[:3])}...")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel: {str(e)}")
    
    def create_field_button(self, column, idx):
        frame = ttk.Frame(self.fields_list_frame)
        frame.grid(row=idx, column=0, sticky=(tk.W, tk.E), pady=3, padx=5)
        
        btn = ttk.Button(frame, text=f"📍 Click to place: {column}", 
                        command=lambda c=column: self.select_column(c))
        btn.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)
        
        # Font size
        ttk.Label(frame, text="Size:").grid(row=0, column=1, padx=5)
        font_size = ttk.Entry(frame, width=5)
        font_size.insert(0, "40")
        font_size.grid(row=0, column=2, padx=5)
        
        # Status indicator
        status = ttk.Label(frame, text="⚪ Not placed", foreground="gray")
        status.grid(row=0, column=3, padx=5)
        
        # Clear button
        clear_btn = ttk.Button(frame, text="Clear", width=6,
                              command=lambda c=column: self.clear_field(c))
        clear_btn.grid(row=0, column=4, padx=5)
        
        self.fields[column] = {
            'button': btn,
            'status': status,
            'font_size': font_size,
            'x': None,
            'y': None,
            'marker': None
        }
        
        frame.columnconfigure(0, weight=1)
    
    def select_column(self, column):
        self.current_column = column
        self.status_label.config(text=f"Click on template to place: {column}")
        # Highlight selected button
        for col, field in self.fields.items():
            if col == column:
                field['button'].state(['pressed'])
            else:
                field['button'].state(['!pressed'])
    
    def on_canvas_click(self, event):
        if not self.current_column or not self.template_image:
            messagebox.showwarning("Warning", 
                "Please select a field to place first!")
            return
        
        # Convert canvas coordinates to image coordinates
        canvas_width = 800
        canvas_height = 600
        
        img_display_width = int(self.template_image.width * self.scale_factor)
        img_display_height = int(self.template_image.height * self.scale_factor)
        
        offset_x = (canvas_width - img_display_width) // 2
        offset_y = (canvas_height - img_display_height) // 2
        
        click_x = event.x - offset_x
        click_y = event.y - offset_y
        
        # Check if click is within image bounds
        if 0 <= click_x <= img_display_width and 0 <= click_y <= img_display_height:
            # Convert to original image coordinates
            orig_x = int(click_x / self.scale_factor)
            orig_y = int(click_y / self.scale_factor)
            
            # Store coordinates
            field = self.fields[self.current_column]
            field['x'] = orig_x
            field['y'] = orig_y
            field['status'].config(text=f"✓ ({orig_x}, {orig_y})", foreground="green")
            
            # Draw marker
            self.draw_marker(event.x, event.y, self.current_column)
            
            self.status_label.config(text=f"✓ Placed {self.current_column} at ({orig_x}, {orig_y})")
            
            # Reset button state
            field['button'].state(['!pressed'])
            self.current_column = None
    
    def draw_marker(self, x, y, label):
        # Remove old marker if exists
        if self.fields[label]['marker']:
            self.canvas.delete(self.fields[label]['marker'])
        
        # Draw new marker
        marker = self.canvas.create_oval(x-5, y-5, x+5, y+5, 
                                        fill='red', outline='white', width=2)
        text_id = self.canvas.create_text(x, y-15, text=label, 
                                         fill='red', font=('Arial', 10, 'bold'))
        
        self.fields[label]['marker'] = [marker, text_id]
    
    def redraw_markers(self):
        canvas_width = 800
        canvas_height = 600
        
        img_display_width = int(self.template_image.width * self.scale_factor)
        img_display_height = int(self.template_image.height * self.scale_factor)
        
        offset_x = (canvas_width - img_display_width) // 2
        offset_y = (canvas_height - img_display_height) // 2
        
        for label, field in self.fields.items():
            if field['x'] is not None and field['y'] is not None:
                display_x = int(field['x'] * self.scale_factor) + offset_x
                display_y = int(field['y'] * self.scale_factor) + offset_y
                self.draw_marker(display_x, display_y, label)
    
    def clear_field(self, column):
        field = self.fields[column]
        field['x'] = None
        field['y'] = None
        field['status'].config(text="⚪ Not placed", foreground="gray")
        
        if field['marker']:
            for item in field['marker']:
                self.canvas.delete(item)
            field['marker'] = None
    
    def create_certificate(self, row_data):
        cert = self.template_image.copy()
        draw = ImageDraw.Draw(cert)
        
        for column, field in self.fields.items():
            if field['x'] is None or field['y'] is None:
                continue
            
            try:
                x = field['x']
                y = field['y']
                font_size = int(field['font_size'].get())
                text = str(row_data[column])
                
                try:
                    font = ImageFont.truetype("arial.ttf", font_size)
                except:
                    try:
                        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size)
                    except:
                        font = ImageFont.load_default()
                
                draw.text((x, y), text, fill='black', font=font)
            except Exception as e:
                print(f"Error drawing field {column}: {e}")
        
        return cert
    
    def preview_certificate(self):
        if not self.validate_inputs():
            return
        
        if len(self.excel_data) == 0:
            messagebox.showwarning("Warning", "Excel file is empty!")
            return
        
        # Create preview window
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Certificate Preview")
        preview_window.geometry("900x700")
        
        # Use first row for preview
        preview_data = self.excel_data.iloc[0]
        cert_image = self.create_certificate(preview_data)
        
        # Resize for preview
        preview_width = 850
        aspect_ratio = cert_image.height / cert_image.width
        preview_height = int(preview_width * aspect_ratio)
        
        cert_image = cert_image.resize((preview_width, preview_height), 
                                      Image.Resampling.LANCZOS)
        
        canvas = tk.Canvas(preview_window, width=preview_width, height=preview_height)
        canvas.pack(padx=20, pady=20)
        
        photo = ImageTk.PhotoImage(cert_image)
        canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        canvas.image = photo
        
        ttk.Label(preview_window, text="Preview of first certificate", 
                 font=('Arial', 12)).pack(pady=10)
    
    def generate_certificates(self):
        if not self.validate_inputs():
            return
        
        output_dir = filedialog.askdirectory(title="Select Output Folder for Certificates")
        if not output_dir:
            return
        
        try:
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Generating Certificates")
            progress_window.geometry("400x150")
            
            ttk.Label(progress_window, text="Generating certificates...", 
                     font=('Arial', 12)).pack(pady=20)
            
            progress = ttk.Progressbar(progress_window, length=300, mode='determinate')
            progress.pack(pady=10)
            
            status_label = ttk.Label(progress_window, text="")
            status_label.pack(pady=10)
            
            total = len(self.excel_data)
            progress['maximum'] = total
            
            for idx, row in self.excel_data.iterrows():
                cert = self.create_certificate(row)
                
                # Use first column value for filename
                filename = f"certificate_{idx+1}_{str(row.iloc[0]).replace(' ', '_')[:30]}.png"
                filepath = os.path.join(output_dir, filename)
                cert.save(filepath)
                
                progress['value'] = idx + 1
                status_label.config(text=f"Generated {idx+1} of {total}")
                progress_window.update()
            
            progress_window.destroy()
            messagebox.showinfo("Success!", 
                f"✓ Successfully generated {total} certificates!\n\n"
                f"Saved to: {output_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate certificates:\n{str(e)}")
    
    def validate_inputs(self):
        if not self.template_path:
            messagebox.showwarning("Warning", "Please load a template first!")
            return False
        
        if self.excel_data is None:
            messagebox.showwarning("Warning", "Please load Excel file first!")
            return False
        
        placed_fields = sum(1 for f in self.fields.values() if f['x'] is not None)
        if placed_fields == 0:
            messagebox.showwarning("Warning", 
                "Please place at least one field on the template!")
            return False
        
        return True

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateGenerator(root)
    root.mainloop()