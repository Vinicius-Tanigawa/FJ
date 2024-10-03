import pandas as pd
import qrcode
import os
import re

def sanitize_filename(name):
    """Remove or replace invalid characters in filenames."""
    return re.sub(r'[\\/*?:"<>|\n]', "_", name)

def load_data(file_path):
    """Load the Excel file and return the data as a DataFrame."""
    return pd.read_excel(file_path)

def create_qr_code(content, output_path):
    """Generate a QR code for the given content and save it to the specified path."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(content)
    qr.make(fit=True)
    
    img = qr.make_image(fill='black', back_color='white')
    img.save(output_path)

def generate_qr_codes(data, output_dir, starting_id):
    """Generate QR codes for each entry in the data and save them in the output directory."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    ids_names_paths = []

    for index, row in data.iterrows():
        name = row['Etiqueta']
        sanitized_name = sanitize_filename(name)
        
        # Generate ID number
        id_number = starting_id + index
        
        # Generate QR code based on ID number
        img_path = os.path.join(output_dir, f'{sanitized_name}_{id_number}.png')
        create_qr_code(id_number, img_path)
        
        # Add ID, name, and path to the list
        ids_names_paths.append({'ID': id_number, 'Name': name, 'QRCodePath': img_path})
    
    return pd.DataFrame(ids_names_paths)

def save_to_excel(data, output_file_path):
    """Save the DataFrame to an Excel file."""
    data.to_excel(output_file_path, index=False)

def main(input_file_path, output_dir, output_file_path, starting_id=100):
    data = load_data(input_file_path)
    output_data = generate_qr_codes(data, output_dir, starting_id)
    save_to_excel(output_data, output_file_path)
    print(f"QR codes generated and saved in {output_dir}")
    print(f"New Excel file with QR code paths and IDs saved as {output_file_path}")

if __name__ == "__main__":
    input_file_path = './FJ2024/adicional.xlsx'
    output_dir = './FJ2024/qr_codes_adicional'
    output_file_path = './FJ2024/output_adicional.xlsx'
    starting_id = 1061

    main(input_file_path, output_dir, output_file_path, starting_id)
