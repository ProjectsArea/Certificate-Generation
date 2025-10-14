import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
import openpyxl
import os
import sys
from datetime import datetime

# ✅ Helper: handle relative paths for fonts (PyInstaller friendly)
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ✅ Load font safely
def load_font(font_file, size):
    paths_to_try = [
        resource_path(font_file),
        os.path.join("C:\\Windows\\Fonts", font_file),
        os.path.join("C:\\Windows\\Fonts", font_file.replace('.ttf', '.TTF'))
    ]
    
    for path in paths_to_try:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    
    print(f"Warning: Font file '{font_file}' not found. Using default font.")
    return ImageFont.load_default()

# ✅ Smart text wrapping function
def wrap_text(text, font, max_width, draw):
    words = text.split()
    lines = []
    current_line = []
    
    for word in words:
        test_line = ' '.join(current_line + [word])
        bbox = draw.textbbox((0, 0), test_line, font=font)
        width = bbox[2] - bbox[0]
        
        if width <= max_width:
            current_line.append(word)
        else:
            if current_line:
                lines.append(' '.join(current_line))
                current_line = [word]
            else:
                lines.append(word)
    
    if current_line:
        lines.append(' '.join(current_line))
    
    return lines

# ✅ Draw multi-line text
def draw_multiline_text(draw, lines, start_x, start_y, font, fill, line_spacing=5):
    y = start_y
    for line in lines:
        draw.text((start_x, y), line, font=font, fill=fill)
        bbox = draw.textbbox((0, 0), line, font=font)
        y += bbox[3] - bbox[1] + line_spacing

# ✅ Format date helper
def format_date(date_value):
    if date_value is None:
        return ""
    
    if isinstance(date_value, str):
        try:
            dt = datetime.strptime(date_value, "%d-%m-%Y")
            return dt.strftime("%d - %m - %Y")
        except:
            try:
                dt = datetime.strptime(date_value, "%Y-%m-%d")
                return dt.strftime("%d - %m - %Y")
            except:
                return date_value
    else:
        return date_value.strftime("%d - %m - %Y")

# ✅ Core certificate generation function
def generate_certificates(template_path, excel_path, output_dir, save_format):
    try:
        # Load template
        template = Image.open(template_path).convert("RGB")

        # Load fonts
        font_name = load_font("times.ttf", 28)
        font_course = load_font("times.ttf", 28)
        font_grade = load_font("times.ttf", 28)
        font_location = load_font("times.ttf", 28)
        font_center = load_font("times.ttf", 28)
        font_dates = load_font("times.ttf", 28)

        # Load Excel
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook.active

        os.makedirs(output_dir, exist_ok=True)

        # Excel columns mapping (Certificate_No REMOVED)
        # Column 0: Name
        # Column 1: Course
        # Column 2: Grade
        # Column 3: Centre_Location
        # Column 4: Centre_Code
        # Column 5: From_Date
        # Column 6: To_Date
        # Column 7: Date_of_Issue

        certificates_generated = 0

        # Iterate through each row (skip header)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 8:
                continue
                
            student_name = row[0]
            course_name = row[1]
            grade = row[2]
            centre_location = row[3]
            centre_code = row[4]
            from_date = row[5]
            to_date = row[6]
            date_of_issue = row[7]

            if not (student_name and course_name):
                continue

            cert_image = template.copy()
            draw = ImageDraw.Draw(cert_image)

            # 📌 COORDINATES (Certificate No REMOVED)
            name_position = (400, 275)
            course_position = (565, 345)
            grade_position = (880, 400)
            location_position = (175, 470)
            center_code_position = (960, 470)
            from_date_position = (312, 540)
            to_date_position = (820, 540)
            issue_date_position = (240, 755)

            text_color = (60, 40, 20)

            # Draw Student Name
            draw.text(name_position, str(student_name).upper(), font=font_name, fill=text_color)

            # Draw Course Name (with wrapping)
            course_lines = wrap_text(str(course_name), font_course, 580, draw)
            draw_multiline_text(draw, course_lines, course_position[0], course_position[1], font_course, fill=text_color, line_spacing=3)

            # Draw Grade
            if grade:
                draw.text(grade_position, str(grade).upper(), font=font_grade, fill=text_color)

            # Draw Centre Location
            if centre_location:
                location_lines = wrap_text(str(centre_location).upper(), font_location, 680, draw)
                draw_multiline_text(draw, location_lines, location_position[0], location_position[1], font_location, fill=text_color, line_spacing=3)

            # Draw Centre Code
            if centre_code:
                draw.text(center_code_position, str(centre_code), font=font_center, fill=text_color)

            # Draw Dates
            if from_date:
                draw.text(from_date_position, format_date(from_date), font=font_dates, fill=text_color)
            if to_date:
                draw.text(to_date_position, format_date(to_date), font=font_dates, fill=text_color)
            if date_of_issue:
                draw.text(issue_date_position, format_date(date_of_issue), font=font_dates, fill=text_color)

            # Save file
            safe_name = str(student_name).replace(" ", "_").replace("/", "-").replace(":", "-").replace("\\", "-")
            filename = f"{safe_name}"
            ext = "pdf" if save_format.upper() == "PDF" else "png"
            out_path = os.path.join(output_dir, f"{filename}.{ext}")

            counter = 1
            while os.path.exists(out_path):
                out_path = os.path.join(output_dir, f"{safe_name}_{counter}.{ext}")
                counter += 1

            if save_format.upper() == "PDF":
                cert_image.save(out_path, "PDF", resolution=100.0)
            else:
                cert_image.save(out_path, "PNG")
            
            certificates_generated += 1
            print(f"Generated: {filename}")

        messagebox.showinfo("Success", f"✅ {certificates_generated} Certificate(s) generated successfully!")
        
    except Exception as e:
        messagebox.showerror("Error", f"❌ An error occurred:\n{str(e)}")
        print(f"Error details: {str(e)}")

# ✅ File selectors
def select_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Image/Excel Files", "*.png *.jpg *.jpeg *.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def select_directory(entry):
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

# ✅ GUI
app = tk.Tk()
app.title("Datapro Certificate Generator")
app.geometry("650x350")
app.resizable(False, False)

frame = tk.Frame(app)
frame.pack(pady=20)

# Template
tk.Label(frame, text="Certificate Template:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
template_entry = tk.Entry(frame, width=45)
template_entry.grid(row=0, column=1, padx=5)
tk.Button(frame, text="Browse", command=lambda: select_file(template_entry)).grid(row=0, column=2, padx=5)

# Excel
tk.Label(frame, text="Excel File:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
excel_entry = tk.Entry(frame, width=45)
excel_entry.grid(row=1, column=1, padx=5)
tk.Button(frame, text="Browse", command=lambda: select_file(excel_entry)).grid(row=1, column=2, padx=5)

# Output directory
tk.Label(frame, text="Output Folder:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
output_dir_entry = tk.Entry(frame, width=45)
output_dir_entry.grid(row=2, column=1, padx=5)
tk.Button(frame, text="Browse", command=lambda: select_directory(output_dir_entry)).grid(row=2, column=2, padx=5)

# Format
tk.Label(frame, text="Save Format:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
format_options = ["PDF", "Image"]
format_var = tk.StringVar(value=format_options[0])
format_menu = tk.OptionMenu(frame, format_var, *format_options)
format_menu.grid(row=3, column=1, sticky="w")

# Generate Button
tk.Button(app, text="Generate Certificates", bg="#007bff", fg="white", font=("Arial", 12, "bold"), 
          command=lambda: generate_certificates(
              template_entry.get(),
              excel_entry.get(),
              output_dir_entry.get(),
              format_var.get()
          )).pack(pady=20)

app.mainloop()