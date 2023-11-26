import streamlit as st
from utils import pdf_to_image, combine_images, crop_img, convert_img_to_df, df_to_excel, calculate_sum
import glob
import tempfile
import os
from pathlib import Path


def uploaded_path(uploaded_files):
    file_paths = []
    temp_dir = tempfile.mkdtemp()
    
    for uploaded_file in uploaded_files:
        file_path = os.path.join(temp_dir, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.read())
        file_paths.append(file_path)

    if len(file_paths)==1:
        return file_paths[0]
    else:
        common_path = Path(os.path.commonpath(file_paths))
        relative_path = common_path.joinpath("*.pdf")
        abs_path = Path(relative_path).resolve()
        raw_abs_path = r"{}".format(abs_path)

    return raw_abs_path

def process_pdf_files(path):
    pdf_file_paths = glob.glob(path)
    combined_images_list = []
    combined_df_list = []
    
    for file in pdf_file_paths:
        images = pdf_to_image(file)
        combined_image = combine_images(images)
        combined_images_list.append(combined_image)

    for image in combined_images_list:
        a,b,c,d,e,f,g,h,i,j = crop_img(image)
        dataframe = convert_img_to_df(a,b,c,d,e,f,g,h,i,j)
        combined_df_list.append(dataframe)
    
    return df_to_excel(combined_df_list,1)

uploaded_files = st.file_uploader("Silahkan upload hasil lab format pdf", type="pdf",accept_multiple_files=True)
    
if st.button("Convert to Excel"):
    if uploaded_files:
        path = uploaded_path(uploaded_files)
        st.write("Memproses, tunggu sebentar")
        output = process_pdf_files(path)
        st.write("Proses selesai, silahkan download")
        with open(output, "rb") as file:
            file_bytes = file.read()
        st.download_button(label="Download Excel File", data=file_bytes, file_name='hasil.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        st.warning("Please upload PDF files.")