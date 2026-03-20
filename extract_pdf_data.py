import os
import glob
import re
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

def extract_data_from_image(image):
    # Perform OCR on the image
    text = pytesseract.image_to_string(image, lang='hin+eng')

    data = {}

    # Extract basic info
    year_match = re.search(r'निर्वाचक नामावली (\d{4})', text)
    if year_match:
        data['निर्वाचक नामावली वर्ष'] = year_match.group(1)

    assembly_match = re.search(r'विधान सभा क्षेत्र की संख्या , नाम व आरक्षण स्थिति : (.*?)\n', text)
    if assembly_match:
        data['विधान सभा क्षेत्र'] = assembly_match.group(1).strip()

    part_match = re.search(r'भाग संख्या :[ \:]+(\d+)', text)
    if part_match:
        data['भाग संख्या'] = part_match.group(1).strip()

    parliament_match = re.search(r'संसदीय निर्वाचन क्षेत्र की संख्या, नाम व आरक्षण स्थिति.*? : (.*?)\n', text)
    if parliament_match:
        data['संसदीय निर्वाचन क्षेत्र'] = parliament_match.group(1).strip()

    # Extract part 1 info
    rev_year = re.search(r'पुनरीक्षण का वर्ष : (\d{4})', text)
    if rev_year:
        data['पुनरीक्षण का वर्ष'] = rev_year.group(1)

    qual_date = re.search(r'अ[हर्ह]ता तिथि : ([\d-]+)', text)
    if qual_date:
        data['अर्हता तिथि'] = qual_date.group(1)

    # Extract part 2 info
    main_town = re.search(r'मुख्य शहर / मुख्य ग्राम.*?[:४]\s*(.*?)\n', text)
    if main_town:
        data['मुख्य शहर / ग्राम'] = main_town.group(1).strip()

    ward = re.search(r'वार्ड.*?[:४]\s*(.*?)\n', text)
    if ward:
        data['वार्ड'] = ward.group(1).strip()

    post_office = re.search(r'पोस्ट ऑफिस.*?[:४]\s*(.*?)\n', text)
    if post_office:
        data['पोस्ट ऑफिस'] = post_office.group(1).strip()

    police_station = re.search(r'पुलिस थाना.*?[:४]\s*(.*?)\n', text)
    if police_station:
        data['पुलिस थाना'] = police_station.group(1).strip()

    tehsil = re.search(r'तहसील.*?[:४]\s*(.*?)\n', text)
    if tehsil:
        data['तहसील'] = tehsil.group(1).strip()

    district = re.search(r'जिला.*?[:४]\s*(.*?)\n', text)
    if district:
        data['जिला'] = district.group(1).strip()

    pincode = re.search(r'पिन कोड.*?[:४]?\s*(\d{6})', text)
    if pincode:
        data['पिन कोड'] = pincode.group(1)

    # Extract part 3 info
    polling_station_no_name = re.search(r'मतदान केंद्र की संख्या व नाम :\s*(.*?)\n', text)
    if polling_station_no_name:
        data['मतदान केंद्र की संख्या व नाम'] = polling_station_no_name.group(1).strip()

    polling_station_address = re.search(r'मतदान केंद्र का भवन व पता :\s*(.*?)\n\n', text)
    if polling_station_address:
        data['मतदान केंद्र का भवन व पता'] = polling_station_address.group(1).strip()

    # Extract part 4 info
    voters_match = re.search(r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d*)\s+(\d+)', text.split('मतदाताओं की संख्या')[-1])
    if voters_match:
        data['आरंभिक क्रम संख्या'] = voters_match.group(1)
        data['अंतिम क्रम संख्या'] = voters_match.group(2)
        data['पुरुष मतदाता'] = voters_match.group(3)
        data['महिला मतदाता'] = voters_match.group(4)
        data['तृतीय लिंग मतदाता'] = voters_match.group(5) if voters_match.group(5) else "0"
        data['कुल मतदाता'] = voters_match.group(6)

    return data

def process_pdfs_in_folder(folder_path, output_excel='extracted_voter_data.xlsx'):
    # Supported extensions
    pdf_files = glob.glob(os.path.join(folder_path, '*.pdf'))
    # Also support processing image directly if needed for testing
    img_files = glob.glob(os.path.join(folder_path, '*.png')) + glob.glob(os.path.join(folder_path, '*.jpg'))

    all_files = pdf_files + img_files

    if not all_files:
        print(f"No PDF or Image files found in {folder_path}")
        return

    print(f"Found {len(all_files)} files. Processing...")
    all_data = []

    for file_path in all_files:
        print(f"Processing: {os.path.basename(file_path)}")
        try:
            if file_path.lower().endswith('.pdf'):
                # Convert the first page of the PDF to an image
                # poppler-utils must be installed on the system
                images = convert_from_path(file_path, first_page=1, last_page=1)
                if images:
                    image = images[0]
                else:
                    print(f"Could not extract image from {file_path}")
                    continue
            else:
                # Direct image file
                image = Image.open(file_path)

            data = extract_data_from_image(image)
            data['File Name'] = os.path.basename(file_path)
            all_data.append(data)

        except Exception as e:
            print(f"Error processing {file_path}: {e}")

    if all_data:
        df = pd.DataFrame(all_data)

        # Reorder columns to have File Name first
        cols = ['File Name'] + [col for col in df.columns if col != 'File Name']
        df = df[cols]

        df.to_excel(output_excel, index=False)
        print(f"Extraction complete! Data saved to {output_excel}")
    else:
        print("No data was extracted.")

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Extract data from first pages of Electoral Roll PDFs into Excel.')
    parser.add_argument('folder', nargs='?', default='.', help='Folder containing the PDFs (default: current directory)')
    parser.add_argument('--output', default='extracted_voter_data.xlsx', help='Output Excel file name')

    args = parser.parse_args()
    process_pdfs_in_folder(args.folder, args.output)
