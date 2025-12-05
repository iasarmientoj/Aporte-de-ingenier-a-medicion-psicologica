import os
import cv2
from pyzbar.pyzbar import decode, ZBarSymbol
import numpy as np

# Define the folder path containing the scanned images
folder_path = "resultados-hojas-escaneadas-png"

# Ensure the folder exists
if not os.path.exists(folder_path):
    print(f"Error: The folder '{folder_path}' does not exist.")
    exit()

# Get a list of all image files in the folder (supporting common image extensions)
image_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.tiff')
image_files = [f for f in os.listdir(folder_path) if f.lower().endswith(image_extensions)]

# Check if there are any images in the folder
if not image_files:
    print(f"No images found in the folder '{folder_path}'.")
    exit()

def preprocess_image(image):
    # Convert to grayscale
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    return gray

# Process each image
for image_file in image_files:
    print(image_file)
    # Construct full path to the image
    image_path = os.path.join(folder_path, image_file)
    
    # Read the image using OpenCV
    image = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    if image is None:
        print(f"Error: Could not read image '{image_file}'.")
        continue
    
    # Preprocess the image
    processed_image = preprocess_image(image)
    
    # Try multiple scales to improve detection
    scales = [1.0, 0.75, 1.25]  # Adjusted for high-resolution images
    barcodes = []
    
    for scale in scales:
        # Resize processed image
        scaled_image = cv2.resize(processed_image, None, fx=scale, fy=scale, interpolation=cv2.INTER_LINEAR)
        
        # Focus on Code 39 barcodes only
        barcode_types = [ZBarSymbol.CODE39]
        
        # Detect barcodes in the scaled image
        detected_barcodes = decode(scaled_image, symbols=barcode_types)
        
        if detected_barcodes:
            barcodes.extend(detected_barcodes)
            break  # Stop if barcodes are found at this scale
    
    if barcodes:
        # Get the first barcode found
        barcode_data = barcodes[0].data.decode('utf-8')
        print(f"Detected barcode in '{image_file}': {barcode_data}")
        
        # Create new filename based on barcode (ensure valid filename)
        new_filename = barcode_data.replace('/', '_').replace('\\', '_') + os.path.splitext(image_file)[1]
        new_filepath = os.path.join(folder_path, new_filename)
        
        # Check if a file with the new name already exists
        if os.path.exists(new_filepath):
            print(f"Warning: File '{new_filename}' already exists. Skipping rename for '{image_file}'.")
            continue
        
        # Rename the image file
        try:
            os.rename(image_path, new_filepath)
            print(f"Renamed '{image_file}' to '{new_filename}'.")
        except Exception as e:
            print(f"Error renaming '{image_file}' to '{new_filename}': {e}")
    else:
        print(f"No barcode found in image '{image_file}'.")

print("Processing complete.")