import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
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

# ✅ Auto-scale text to fit within max width
def get_fitted_font(text, font_file, initial_size, max_width, draw):
    font = load_font(font_file, initial_size)
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    
    if text_width <= max_width:
        return font
    
    size = initial_size
    while size > 20:
        size -= 2
        font = load_font(font_file, size)
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        if text_width <= max_width:
            return font
    
    return font

# ✅ Generate single test certificate
def generate_test_certificate(template_path, cert_no, full_name, course_name, output_dir, save_format):
    if not template_path or not cert_no or not full_name or not course_name or not output_dir:
        messagebox.showerror("Error", "Please fill all fields!")
        return
    
    # Load template
    template = Image.open(template_path).convert("RGB")
    
    # Load fonts
    font_number = load_font("arial.ttf", 36)
    font_name = load_font("arialbd.ttf", 55)
    font_course = load_font("arial.ttf", 40)
    
    cert_image = template.copy()
    draw = ImageDraw.Draw(cert_image)
    
    # Positions (matching your original code)
    cert_no_position = (105, 424)
    name_position = (500, 350)
    course_position = (994, 470)
    
    # Maximum widths
    max_name_width = 430
    max_course_width = 290
    
    # Get fitted fonts
    fitted_name_font = get_fitted_font(str(full_name), "arialbd.ttf", 55, max_name_width, draw)
    fitted_course_font = get_fitted_font(str(course_name), "arial.ttf", 40, max_course_width, draw)
    
    # Draw text
    draw.text(cert_no_position, str(cert_no), font=font_number, fill=(0, 0, 0))
    draw.text(name_position, str(full_name), font=fitted_name_font, fill=(0, 0, 0))
    draw.text(course_position, str(course_name), font=fitted_course_font, fill=(0, 0, 0))
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Save file
    filename = f"{full_name}_{cert_no}".replace(" ", "_")
    ext = "pdf" if save_format.upper() == "PDF" else "png"
    out_path = os.path.join(output_dir, f"{filename}.{ext}")
    
    if save_format.upper() == "PDF":
        cert_image.save(out_path, "PDF", resolution=100.0)
    else:
        cert_image.save(out_path, "PNG")
    
    messagebox.showinfo("Success", f"✅ Certificate generated!\n\nSaved to:\n{out_path}")

# ✅ File selectors
def select_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png *.jpg *.jpeg")])
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
app.title("Single Certificate Test Generator")
app.geometry("650x450")
app.resizable(False, False)

frame = tk.Frame(app)
frame.pack(pady=20)

# Template
tk.Label(frame, text="Certificate Template:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
template_entry = tk.Entry(frame, width=45)
template_entry.grid(row=0, column=1, padx=5)
tk.Button(frame, text="Browse", command=lambda: select_file(template_entry)).grid(row=0, column=2, padx=5)

# Certificate Number
tk.Label(frame, text="Certificate No.:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
cert_no_entry = tk.Entry(frame, width=45)
cert_no_entry.grid(row=1, column=1, padx=5)
cert_no_entry.insert(0, "TEST001")  # Default value

# Full Name
tk.Label(frame, text="Full Name:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
name_entry = tk.Entry(frame, width=45)
name_entry.grid(row=2, column=1, padx=5)
name_entry.insert(0, "John Doe")  # Default value

# Course Name
tk.Label(frame, text="Course Name:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
course_entry = tk.Entry(frame, width=45)
course_entry.grid(row=3, column=1, padx=5)
course_entry.insert(0, "Web Development")  # Default value

# Output directory
tk.Label(frame, text="Output Folder:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
output_dir_entry = tk.Entry(frame, width=45)
output_dir_entry.grid(row=4, column=1, padx=5)
tk.Button(frame, text="Browse", command=lambda: select_directory(output_dir_entry)).grid(row=4, column=2, padx=5)

# Format
tk.Label(frame, text="Save Format:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
format_options = ["PDF", "Image"]
format_var = tk.StringVar(value=format_options[1])  # Default to Image for testing
format_menu = tk.OptionMenu(frame, format_var, *format_options)
format_menu.grid(row=5, column=1, sticky="w")

# Generate Button
tk.Button(app, text="Generate Test Certificate", bg="#28a745", fg="white", font=("Arial", 12, "bold"), 
          command=lambda: generate_test_certificate(
              template_entry.get(),
              cert_no_entry.get(),
              name_entry.get(),
              course_entry.get(),
              output_dir_entry.get(),
              format_var.get()
          )).pack(pady=20)

# Info Label
info_label = tk.Label(app, text="Test with a single certificate before batch processing", 
                      font=("Arial", 9), fg="gray")
info_label.pack()

app.mainloop()