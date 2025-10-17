import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
import openpyxl
import os
import sys

# ✅ Helper: handle relative paths for fonts (PyInstaller friendly)
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ✅ Load font safely
def load_font(font_file, size):
    # Try multiple paths: local folder, Windows fonts, system _MEIPASS
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
    """Wrap text to fit within max_width, returning list of lines"""
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
                # Single word is too long, add it anyway
                lines.append(word)
    
    if current_line:
        lines.append(' '.join(current_line))
    
    return lines

# ✅ Draw multi-line text centered
def draw_multiline_text(draw, lines, start_x, start_y, font, fill, line_spacing=5):
    """Draw multiple lines of text, each line centered at start_x"""
    y = start_y
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        line_width = bbox[2] - bbox[0]
        x = start_x - (line_width / 2)  # Center each line
        draw.text((x, y), line, font=font, fill=fill)
        y += bbox[3] - bbox[1] + line_spacing

# ✅ Core certificate generation function
def generate_certificates(template_path, excel_path, output_dir, save_format):
    # Load template
    template = Image.open(template_path).convert("RGB")
    width, height = template.size

    # Load fonts - matching certificate style
    font_number = load_font("arial.ttf", 36)   # Certificate Number
    font_name = load_font("arialbd.ttf", 55)  # Student Name - fixed size
    font_course = load_font("arial.ttf", 45)  # Course Name - optimized size

    # Load Excel
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active

    os.makedirs(output_dir, exist_ok=True)

    # Iterate through each row (skip header)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        cert_no = row[1]        # CERTIFICATE NO.
        full_name = row[2]      # FullName
        course_name = row[3]    # COURSE
        college_name = row[4]   # NAME OF THE COLLEGE
        start_date = row[5]     # START DATE
        end_date = row[6]       # END DATE

        if not (cert_no and full_name and course_name):
            continue

        cert_image = template.copy()
        draw = ImageDraw.Draw(cert_image)

        # 📌 Coordinate positions
        cert_no_position = (105, 424)         # Certificate number position
        name_position = (400, 350)            # Student name position
        course_center_x = 640                 # Center X for course name (centered on page)
        course_start_y = 515                  # Starting Y for course name (below the main text)
        max_course_width = 800                # Maximum width for course text

        # Draw certificate number and name (single line)
        draw.text(cert_no_position, str(cert_no), font=font_number, fill=(0, 0, 0))
        draw.text(name_position, str(full_name), font=font_name, fill=(0, 0, 0))

        # Draw course name with wrapping if needed
        course_lines = wrap_text(str(course_name), font_course, max_course_width, draw)
        draw_multiline_text(draw, course_lines, course_center_x, course_start_y, 
                          font_course, fill=(0, 0, 0), line_spacing=3)

        # File name
        filename = f"{full_name}_{cert_no}".replace(" ", "_")
        out_path = os.path.join(output_dir, f"{filename}.{'pdf' if save_format.upper() == 'PDF' else 'png'}")

        if save_format.upper() == "PDF":
            cert_image.save(out_path, "PDF", resolution=100.0)
        else:
            cert_image.save(out_path, "PNG")

    messagebox.showinfo("Success", "✅ Certificates generated successfully!")

# ✅ File selector helpers
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
app.title("Internship Certificate Generator")
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