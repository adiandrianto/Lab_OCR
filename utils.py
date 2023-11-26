from pdf2image import convert_from_path
from PIL import Image
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)
import pytesseract
import xlsxwriter
import pandas as pd
import tempfile
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\Tesseract.exe'

parameter = ['','HEMATOLOGY', 'Routine Hematology', 'Hemoglobin', 'Hematocrit', 'Leukocyte',
    'Erythrocyte', 'Thrombocyte', 'ESR', 'Erythrocyte Index', 'MCV', 'MCH', 'MCHC',
    'Differential Count', 'Basophil', 'Eosinophil', 'Band Neutrophil', 'Neutrophil',
    'Lymphocyte', 'Monocyte', '', 'BLOOD CHEMICAL', 'Liver Function', 'Total Bilirubbin',
    'Direct Bilirubin', 'Alkaline Phosphatase', 'SGOT', 'SGPT', 'Diabetes',
    'Random Blood Glucose', 'Kidney Function', 'Ureum', 'Creatinine', 'eGFR/Creatinine', '',
    'URINALYSIS', 'Macroscopic', 'Color', 'Appearance', 'Odor', 'Urine Chemical',
    'pH', 'Spec. Gravity', 'Albumin', 'Glucose', 'Keton', 'Urobilinogen',
    'Occult Blood', 'Leukocyte Esterase', 'Nitrit', 'Microscopic', 'Erythrocyte',
    'Leucocyte', 'Epithel', 'Crystal', 'Bacteria', 'Others']

parameter_df = pd.DataFrame(parameter)

def calculate_sum(a):
    total = []
    for item in a:
        total.append(item)
    return sum(total)

def pdf_to_image(pdf_path):
    imgs = convert_from_path(pdf_path, grayscale=True, dpi=300, use_pdftocairo=True, size=(1754, 2480), poppler_path='poppler-23.08.0/Library/bin')
    return imgs

def combine_images(images):
    total_width = images[0].width * 2  # Assuming 2 pages per image
    max_height = max(img.height for img in images)
    # Create a new image with white background
    combined_image = Image.new('L', (total_width, max_height))
    # Paste the images side by side
    x_offset = 0
    for img in images:
        combined_image.paste(img, (x_offset, 0))
        x_offset += img.width

    return combined_image.point(lambda x: 0 if x < 128 else 255, '1')

def crop_img(image):
    blood_1_crop = image.crop((500, 850, 800, 1140))
    blood_2_crop = image.crop((500, 1140, 800, 1290))
    blood_3_crop = image.crop((500, 1290, 800, 1600))
    blood_4_crop = image.crop((500, 1600, 800, 1880))
    blood_5_crop = image.crop((500, 1880, 800, 1950))
    blood_6_crop = image.crop((500, 1950, 800, 2100))
    urine_1_crop = image.crop((2250, 780, 2600, 1050)) 
    urine_2_crop = image.crop((2250, 1050, 2500, 1440))
    urine_3_crop = image.crop((2250, 1480, 2500, 1800))
    initial_crop = image.crop((320, 465, 480, 510))
    
    return blood_1_crop, blood_2_crop, blood_3_crop, blood_4_crop, blood_5_crop, blood_6_crop, urine_1_crop, urine_2_crop, urine_3_crop, initial_crop

def convert_img_to_df(*args):
    df_list = []
    for param in args :
        text = pytesseract.image_to_string(param,config = '--oem 1 --psm 6' )
        
        # Split the string into individual words
        param_words = text.split('\n\n')
        param_words = text.split('\n')

        # Create empty dict
        param_data = []
        for word in param_words:
            param_data.append({str(param) : word})

        # Create a pandas dataframe from the list of dict
        param_df = pd.DataFrame(param_data)
        param_df.iloc[-1] = param_df.iloc[-1].str.replace('\n', '') # delete '\n' on the last row of each dataframe
        param_df[str(param)] = param_df[str(param)].str.replace('- ', '') # remove '- ' from lab parameter
        param_df[str(param)] = param_df[str(param)].str.replace('L (?=\d)', '*', regex=True)
        param_df[str(param)] = param_df[str(param)].str.replace('H (?=\d)', '*', regex=True)
        param_df.replace({'()': 'Negative'}, inplace=True) # give (-) on urine result, otherwise it gives '()'
        param_df.replace({')': 'Negative'}, inplace=True)
        param_df.replace({'(¢)': 'Negative'}, inplace=True)
        param_df.replace({'©)': 'Negative'}, inplace=True)
        param_df.replace({'(-)': 'Negative'}, inplace=True)
        param_df.replace({'(+)': 'Positive'}, inplace=True)
        empty_rows = param_df.apply(lambda row: row.str.strip().eq(''), axis=1).all(axis=1)

        # Remove empty rows
        df_filtered = param_df[~empty_rows]
        df_filtered.reset_index(drop=True, inplace=True)
        df_list.append(df_filtered)

    return df_list

def df_to_excel(df_list,col):
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
        output = temp_file.name
        #output = "hasil.xlsx"
        writer = pd.ExcelWriter(output, engine="xlsxwriter")
    
        parameter_df.to_excel(writer, sheet_name="Sheet1", startrow=0, startcol=0, header=False, index=False) 

        for df in df_list:    
            df[9].to_excel(writer, sheet_name="Sheet1", startrow=0, startcol=col, header=False, index=False)
            df[0].to_excel(writer, sheet_name="Sheet1", startrow=3, startcol=col, header=False, index=False)
            df[1].to_excel(writer, sheet_name="Sheet1", startrow=10, startcol=col, header=False, index=False)
            df[2].to_excel(writer, sheet_name="Sheet1", startrow=14, startcol=col, header=False, index=False)
            df[3].to_excel(writer, sheet_name="Sheet1", startrow=23, startcol=col, header=False, index=False)
            df[4].to_excel(writer, sheet_name="Sheet1", startrow=29, startcol=col, header=False, index=False)
            df[5].to_excel(writer, sheet_name="Sheet1", startrow=31, startcol=col, header=False, index=False)
            df[6].to_excel(writer, sheet_name="Sheet1", startrow=37, startcol=col, header=False, index=False)
            df[7].to_excel(writer, sheet_name="Sheet1", startrow=41, startcol=col, header=False, index=False)
            df[8].to_excel(writer, sheet_name="Sheet1", startrow=51, startcol=col, header=False, index=False)
            col += 1
        writer.close()
    return output