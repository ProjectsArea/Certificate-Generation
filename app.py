from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
import uuid
from io import BytesIO
import base64
import zipfile

app = Flask(__name__)
app.secret_key = 'a_very_long_and_random_secret_key_that_you_should_change'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['FONT_FOLDER'] = 'static/fonts'

# Create necessary folders
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['FONT_FOLDER'], exist_ok=True)

ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'bmp'}
ALLOWED_EXCEL_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename, extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in extensions

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_template', methods=['POST'])
def upload_template():
    if 'template' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['template']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename, ALLOWED_IMAGE_EXTENSIONS):
        return jsonify({'error': 'Invalid file type'}), 400
    
    # Generate unique filename
    filename = secure_filename(f"{uuid.uuid4()}_{file.filename}")
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    # Get image dimensions
    img = Image.open(filepath)
    width, height = img.size
    
    # Store in session
    session['template_path'] = filepath
    session['template_dimensions'] = {'width': width, 'height': height}
    
    # Convert image to base64 for preview
    buffered = BytesIO()
    img.save(buffered, format=img.format or 'PNG')
    img_str = base64.b64encode(buffered.getvalue()).decode()
    
    return jsonify({
        'success': True,
        'filename': file.filename,
        'dimensions': {'width': width, 'height': height},
        'preview': f"data:image/png;base64,{img_str}"
    })

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'excel' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['excel']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename, ALLOWED_EXCEL_EXTENSIONS):
        return jsonify({'error': 'Invalid file type'}), 400
    
    filename = secure_filename(f"{uuid.uuid4()}_{file.filename}")
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    try:
        df = pd.read_excel(filepath)
        columns = df.columns.tolist()
        row_count = len(df)
        
        # Store in session
        session['excel_path'] = filepath
        session['excel_columns'] = columns
        
        # Get first row as sample data
        sample_data = df.iloc[0].to_dict() if row_count > 0 else {}
        
        return jsonify({
            'success': True,
            'filename': file.filename,
            'columns': columns,
            'row_count': row_count,
            'sample_data': {k: str(v) for k, v in sample_data.items()}
        })
    except Exception as e:
        return jsonify({'error': f'Failed to read Excel: {str(e)}'}), 400

@app.route('/save_fields', methods=['POST'])
def save_fields():
    data = request.json
    session['fields'] = data.get('fields', {})
    return jsonify({'success': True})

@app.route('/preview_certificate', methods=['POST'])
def preview_certificate():
    try:
        template_path = session.get('template_path')
        excel_path = session.get('excel_path')
        fields = session.get('fields', {})
        
        if not template_path or not excel_path:
            return jsonify({'error': 'Missing template or data'}), 400
        
        # Load data
        df = pd.read_excel(excel_path)
        if len(df) == 0:
            return jsonify({'error': 'Excel file is empty'}), 400
        
        # Create certificate
        cert = create_certificate(template_path, df.iloc[0], fields)
        
        # Convert to base64
        buffered = BytesIO()
        cert.save(buffered, format='PNG')
        img_str = base64.b64encode(buffered.getvalue()).decode()
        
        return jsonify({
            'success': True,
            'preview': f"data:image/png;base64,{img_str}"
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    try:
        template_path = session.get('template_path')
        excel_path = session.get('excel_path')
        fields = session.get('fields', {})
        
        if not template_path or not excel_path:
            return jsonify({'error': 'Missing template or data'}), 400
        
        df = pd.read_excel(excel_path)
        
        # Create unique output folder
        output_id = str(uuid.uuid4())
        output_dir = os.path.join(app.config['OUTPUT_FOLDER'], output_id)
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate certificates
        generated_files = []
        for idx, row in df.iterrows():
            cert = create_certificate(template_path, row, fields)
            filename = f"certificate_{idx+1}_{str(row.iloc[0]).replace(' ', '_')[:30]}.png"
            filepath = os.path.join(output_dir, filename)
            cert.save(filepath)
            generated_files.append(filename)
        
        session['last_output_id'] = output_id # Store output_id in session
        session['last_generated_files'] = generated_files # Store generated filenames

        # Create zip file
        zip_filename = f"certificates_{output_id}.zip"
        zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for filename in generated_files:
                file_path = os.path.join(output_dir, filename)
                zipf.write(file_path, filename)
        
        session['last_zip'] = zip_filename
        session['last_output_dir'] = output_dir
        
        return jsonify({
            'success': True,
            'count': len(generated_files),
            'download_url': f'/download/{zip_filename}'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    return jsonify({'error': 'File not found'}), 404

@app.route('/print_certificates')
def print_certificates():
    """Route to display all generated certificates for printing"""
    try:
        output_id = session.get('last_output_id')
        generated_files = session.get('last_generated_files')

        if not output_id or not generated_files:
            return "No certificates generated yet. Please generate certificates first.", 400

        certificate_urls = [
            f'/display_certificate/{output_id}/{filename}'
            for filename in generated_files
        ]
        
        return render_template('print_view.html', certificate_urls=certificate_urls)
    except Exception as e:
        return f"Error generating print view: {str(e)}", 500

@app.route('/display_certificate/<output_id>/<filename>')
def display_certificate(output_id, filename):
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], output_id, filename)
    if os.path.exists(filepath):
        return send_file(filepath, mimetype='image/png')
    return jsonify({'error': 'Certificate not found'}), 404

def create_certificate(template_path, row_data, fields):
    """Create a certificate image with data filled in"""
    img = Image.open(template_path)
    draw = ImageDraw.Draw(img)
    
    for column, field_data in fields.items():
        x = field_data.get('x')
        y = field_data.get('y')-100
        font_size = field_data.get('fontSize', 40)
        
        if x is None or y is None:
            continue
        
        text = str(row_data.get(column, ''))
        
        font = ImageFont.load_default() # Fallback to default

        try:
            # Try to load Arial.ttf first
            arial_font_path = os.path.join(os.environ.get('WINDIR', 'C:/Windows'), 'Fonts', 'arial.ttf')
            if os.path.exists(arial_font_path):
                font = ImageFont.truetype(arial_font_path, font_size)
            else:
                # Fallback to the provided custom fonts if Arial is not found
                custom_font_path1 = os.path.join(app.config['FONT_FOLDER'], "Emotional Rescue Personal Use.ttf")
                if os.path.exists(custom_font_path1):
                    font = ImageFont.truetype(custom_font_path1, font_size)
                else:
                    custom_font_path2 = os.path.join(app.config['FONT_FOLDER'], "FontsFree-Net-GOTHICB0.ttf")
                    if os.path.exists(custom_font_path2):
                        font = ImageFont.truetype(custom_font_path2, font_size)

        except Exception as e:
            print(f"Error loading font: {e}. Using default font.")

        draw.text((x, y), text, fill='black', font=font)
    
    return img

if __name__ == '__main__':
    app.run(debug=True, port=5000)