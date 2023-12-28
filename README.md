## OCR using pytesseract

Convert lab result in pdf to excel

This is a problem I face in my workplace when mycoworker is required to input lab result to excel manually. And this is my solution for that problem. By using OCR (in this case pytesseract) we can extract required result and automatically put it in excel format.

Dataset : On dataset folder I provide 5 lab result that I used in this repo.

Steps :

1. Using pdf2images to convert .pdf to image format because pytesseract work on images and not on pdf doc.
2. Preprcess converted images, in this case I convert image to grayscale and using threshold technique to make text clearer and erase unnecessary lines.
3. Make a crops of lab result (I crop lab results only because I dont need normal value range (nilai rujukan), unit (satuan), and method(metode).
4. Using pytesseract (or other OCR) extract each cropped image and make a dataframes from each one of it.
5. Using xlsxwriter, place each dataframe to excel columns and rows accordingly.

## Streamlit Deployment

to see the deployed streamlit, click links :
https://labocr-pml.streamlit.app/

you can download lab result example in dataset folder.
