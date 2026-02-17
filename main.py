import os
import cv2
import numpy as np
import fitz  # PyMuPDF
import subprocess
import uuid
import json
import requests
import tempfile
import shutil
import platform
from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename
from flask import Flask
from flask_cors import CORS

app = Flask(__name__)
CORS(
    app,
    origins=["http://localhost:3939"],
    supports_credentials=True
)

# --- Configuration ---
UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS_DOC = {'pdf', 'docx', 'doc'}
ALLOWED_EXTENSIONS_IMG = {'jpg', 'jpeg', 'png'}
PRICING_API_URL = "http://localhost:8000/api/config/pricing"

# Saturation Threshold: 0-255. 
# Pixels with saturation < 20 are considered gray/black/white.
SATURATION_THRESHOLD = 20 

def allowed_file(filename):
    return '.' in filename and \
           (filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_DOC or \
            filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_IMG)

def get_file_type(filename):
    ext = filename.rsplit('.', 1)[1].lower()
    if ext in ALLOWED_EXTENSIONS_IMG:
        return 'image'
    return 'document'

def fetch_pricing_config():
    """
    Fetches dynamic pricing from the external API.
    Defaults to standard pricing if API is unreachable.
    Full Color is hardcoded to 1500 as requested.
    """
    prices = {
        "bnw": 500,     
        "color": 1000,  
        "full_color": 1500 
    }
    
    try:
        response = requests.get(PRICING_API_URL, timeout=5)
        if response.status_code == 200:
            data = response.json()
            if data.get('success') and 'data' in data and 'prices' in data['data']:
                api_prices = data['data']['prices']
                prices["bnw"] = api_prices.get('bnw', prices["bnw"])
                prices["color"] = api_prices.get('color', prices["color"])
    except Exception:
        # Silently fail back to defaults if API is down
        pass
        
    return prices

def calculate_price_for_percentage(percentage, price_config):
    if percentage <= 0:
        return "black_and_white", price_config["bnw"]
    elif percentage <= 50:
        return "color", price_config["color"]
    else:
        return "full_color", price_config["full_color"]

def analyze_image_array(img_array):
    """
    Takes an OpenCV image array (BGR), converts to HSV, 
    calculates color percentage based on Saturation.
    """
    try:
        # Convert BGR to HSV
        hsv_img = cv2.cvtColor(img_array, cv2.COLOR_BGR2HSV)
        
        # Extract Saturation channel
        saturation_channel = hsv_img[:, :, 1]
        
        # Count pixels where saturation > threshold
        colored_pixels = np.count_nonzero(saturation_channel > SATURATION_THRESHOLD)
        total_pixels = img_array.shape[0] * img_array.shape[1]
        
        if total_pixels == 0:
            return 0.0
            
        percentage = (colored_pixels / total_pixels) * 100
        return round(percentage, 2)
    except Exception as e:
        print(f"Error analyzing image: {e}")
        return 0.0

def convert_doc_to_pdf(input_path, output_dir):
    """
    Uses LibreOffice to convert DOC/DOCX to PDF.
    Uses a temporary UserInstallation to prevent permission/lock errors.
    """
    possible_cmds = ['soffice', 'lowriter', 'libreoffice']
    if platform.system() == 'Windows':
        possible_cmds.append(r"C:\Program Files\LibreOffice\program\soffice.exe")
        pass

    libreoffice_cmd = None
    for cmd in possible_cmds:
        if shutil.which(cmd):
            libreoffice_cmd = cmd
            break
            
    if libreoffice_cmd is None:
        raise EnvironmentError("LibreOffice not found in PATH.")

    # Create a specific temp directory for LibreOffice user profile
    lo_profile_dir = os.path.join(output_dir, 'lo_profile')
    os.makedirs(lo_profile_dir, exist_ok=True)
    
    user_install_arg = f"-env:UserInstallation=file:///{lo_profile_dir.replace(os.sep, '/')}"

    cmd = [
        libreoffice_cmd,
        user_install_arg,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', output_dir,
        input_path
    ]
    
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        pdf_path = os.path.join(output_dir, base_name + '.pdf')
        
        if os.path.exists(pdf_path):
            return pdf_path
        else:
            raise FileNotFoundError(f"PDF output not found at {pdf_path}")
            
    except subprocess.CalledProcessError as e:
        error_msg = e.stderr.decode() if e.stderr else "Unknown Error"
        raise RuntimeError(f"LibreOffice conversion failed: {error_msg}")
    finally:
        # Cleanup profile
        if os.path.exists(lo_profile_dir):
            shutil.rmtree(lo_profile_dir, ignore_errors=True)

def process_file_logic(filepath, filename, price_config):
    file_type = get_file_type(filename)
    result_entry = {
        "filename": filename,
        "type": file_type,
        "total_pages": 0,  # Initialize
        "total_price": 0,
        "total_price_bnw": 0,
        "colors": []
    }
    
    ext = filename.rsplit('.', 1)[1].lower()
    processing_path = filepath
    temp_pdf_created = False
    
    try:
        # 1. Handle Document Conversion (DOC/DOCX -> PDF)
        if ext in ['doc', 'docx']:
            try:
                processing_path = convert_doc_to_pdf(filepath, os.path.dirname(filepath))
                temp_pdf_created = True
            except Exception as e:
                print(f"DOCX Conversion failed for {filename}: {e}")
                return result_entry

        # 2. Process Image
        if file_type == 'image':
            img = cv2.imread(filepath)
            if img is None:
                raise ValueError("Could not read image file.")
            
            # Images are treated as 1 page
            result_entry["total_pages"] = 1
            
            percentage = analyze_image_array(img)
            category, price = calculate_price_for_percentage(percentage, price_config)
            
            result_entry["total_price"] = price
            result_entry["total_price_bnw"] = price_config["bnw"]
            result_entry["colors"].append({
                "page": 1,
                "color": category,
                "price": price,
                "price_bnw": price_config["bnw"],
                "percentage": percentage
            })
            
        # 3. Process PDF
        else: 
            doc = fitz.open(processing_path)
            
            # Set Total Pages
            result_entry["total_pages"] = doc.page_count
            
            total_doc_price = 0
            total_doc_price_bnw = 0
            
            for page_num, page in enumerate(doc):
                # Render page to pixmap
                # Standard matrix ensures good resolution for analysis (simulating A4 density)
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5)) 
                
                if pix.n < 3:
                    img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                    img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2BGR)
                else:
                    img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                    img_array = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)

                percentage = analyze_image_array(img_array)
                category, price = calculate_price_for_percentage(percentage, price_config)
                
                total_doc_price += price
                total_doc_price_bnw += price_config["bnw"]
                
                result_entry["colors"].append({
                    "page": page_num + 1,
                    "color": category,
                    "price": price,
                    "price_bnw": price_config["bnw"],
                    "percentage": percentage
                })
            
            result_entry["total_price"] = total_doc_price
            result_entry["total_price_bnw"] = total_doc_price_bnw
            doc.close()

    except Exception as e:
        print(f"Error processing file logic {filename}: {str(e)}")
    finally:
        # Clean up generated PDF
        if temp_pdf_created and processing_path and os.path.exists(processing_path):
            try:
                os.remove(processing_path)
            except:
                pass

    return result_entry

@app.route('/detect', methods=['POST'])
def detect_colors():
    if 'files' not in request.files:
        return jsonify({"success": "false", "error": "No files part in the request"}), 400
    
    uploaded_files = request.files.getlist('files')
    if not uploaded_files:
        return jsonify({"success": "false", "error": "No selected file"}), 400

    current_prices = fetch_pricing_config()
    response_data = []
    
    # Unique temp dir for this request
    request_temp_dir = os.path.join(UPLOAD_FOLDER, str(uuid.uuid4()))
    os.makedirs(request_temp_dir, exist_ok=True)

    try:
        for file in uploaded_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                save_path = os.path.join(request_temp_dir, filename)
                file.save(save_path)
                
                file_result = process_file_logic(save_path, filename, current_prices)
                response_data.append(file_result)
                
                if os.path.exists(save_path):
                    try:
                        os.remove(save_path)
                    except:
                        pass
        
        return jsonify({
            "success": "true",
            "data": response_data
        })

    except Exception as e:
        return jsonify({"success": "false", "error": str(e)}), 500
        
    finally:
        shutil.rmtree(request_temp_dir, ignore_errors=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)