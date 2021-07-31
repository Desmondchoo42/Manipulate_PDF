import streamlit as st
from pdfrw import PdfReader, PdfWriter, PageMerge, IndirectPdfDict
import pathlib
import os
from os import path
from glob import glob
from PIL import Image
import numpy as np
#import comtypes.client
from pathlib import Path
from zipfile import ZipFile
import base64
import tempfile

import io
from io import BytesIO
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage

###################################Creating all the functions##################################

def load_image(img):
    im = Image.open(img)
    image = np.array(im)
    return image

def extract_nullpage_from_pdf(pdf_path,remove_word):

    to_del = []
    try:
        for i, page in enumerate(PDFPage.get_pages(pdf_path, caching=True, check_extractable=True)):

            resource_manager = PDFResourceManager()
            fake_file_handle = io.StringIO()
            converter = TextConverter(resource_manager, fake_file_handle)
            page_interpreter = PDFPageInterpreter(resource_manager, converter)
            page_interpreter.process_page(page)
            text = fake_file_handle.getvalue()
            # close open handles
            converter.close()
            fake_file_handle.close()

            if  remove_word in text:
                to_del.append(i)

    except AttributeError:
        fp = open(pdf_path, 'rb')
        for i, page in enumerate(PDFPage.get_pages(fp, caching=True, check_extractable=True)):

            resource_manager = PDFResourceManager()
            fake_file_handle = io.StringIO()
            converter = TextConverter(resource_manager, fake_file_handle)
            page_interpreter = PDFPageInterpreter(resource_manager, converter)
            page_interpreter.process_page(page)
            text = fake_file_handle.getvalue()
            # close open handles
            converter.close()
            fake_file_handle.close()

            if  remove_word in text:
                to_del.append(i)

    return to_del

def extract_null_from_pdf(pdf_path):

    blank_pg = []

    try:
        for i, page in enumerate(PDFPage.get_pages(pdf_path, caching=True, check_extractable=True)):

            resource_manager = PDFResourceManager()
            fake_file_handle = io.StringIO()
            converter = TextConverter(resource_manager, fake_file_handle)
            page_interpreter = PDFPageInterpreter(resource_manager, converter)
            page_interpreter.process_page(page)
            text = fake_file_handle.getvalue()
            # close open handles
            converter.close()
            fake_file_handle.close()

            if text == "  \f":
                blank_pg.append(i)

    except AttributeError:
        fp = open(pdf_path, 'rb')
        for i, page in enumerate(PDFPage.get_pages(fp, caching=True, check_extractable=True)):

            resource_manager = PDFResourceManager()
            fake_file_handle = io.StringIO()
            converter = TextConverter(resource_manager, fake_file_handle)
            page_interpreter = PDFPageInterpreter(resource_manager, converter)
            page_interpreter.process_page(page)
            text = fake_file_handle.getvalue()
            # close open handles
            converter.close()
            fake_file_handle.close()

            if text == "  \f":
                blank_pg.append(i)

    return blank_pg

def find_ext(dr, ext):
    return glob(path.join(dr,"*.{}".format(ext)))

####################    Streamlit  ####################


st.image(load_image(os.getcwd()+"/Title.png"))
st.write("##")

st.subheader("Choose Options")

uploaded_file_pdf = st.file_uploader("Upload PDF Files",type=['pdf'], accept_multiple_files=True)

st.write("-----")
st.subheader("Choose type of manipulation to PDF")

config_select_manipulation = st.multiselect("Select one or more options:", ["Add Watermark", "Remove Metadata", "Concatenate PDFs","Remove blank pages", "Remove pages that contain words/phrases"], ["Add Watermark", "Remove Metadata", "Concatenate PDFs"])
if "Add Watermark" in config_select_manipulation:
    uploaded_file_wmp = st.file_uploader("Upload watermark PDF for portrait",type=['pdf'])
    uploaded_file_wml = st.file_uploader("Upload watermark PDF for landscape",type=['pdf'])
remove_word = ""
if "Remove pages that contain words/phrases" in config_select_manipulation:
    remove_word = st.text_input("Input words / phrases contain in the page so that the page will be removed:","To remove")

####################    Actual Code  ####################

def main():
    ## Set up progress bar
    st.write("-----")
    st.subheader("Status")

    progress_bar = st.progress(0)
    status_text = st.empty()

    status_text.text("In progress... Please Wait.")

    ## Making sure that all residue annex files were deleted in the inetrmediate folder
    filepath_todel = find_ext(os.getcwd()+"/Intermediate_Data","pdf")
    for file in filepath_todel:
        os.remove(file)    
    status_text.text("Files in Intermediate folder deleted. Proceeding to next step...")

    ## Define the reader and writer objects

    writer = PdfWriter()
    if "Add Watermark" in config_select_manipulation:
        watermark_input_P = PdfReader(uploaded_file_wmp)
        watermark_input_LS = PdfReader(uploaded_file_wml)
        watermark_P = watermark_input_P.pages[0]
        watermark_LS = watermark_input_LS.pages[0]
        status_text.text("Loaded Watermark PDF. Progressing to the next step...")
        progress_bar.progress(0.25)

    filepath = uploaded_file_pdf

    ## Create a loop for all the paths
    for i in range(len(filepath)):
        file = filepath[i]
        reader_input = PdfReader(file)
        status_text.text(f"Processing {file} now...")

        if "Remove pages that contain words/phrases" in config_select_manipulation:
            to_del = (extract_nullpage_from_pdf(file,remove_word))
        else:
            to_del = []

        if "Remove blank pages" in config_select_manipulation:
            blank_pg = (extract_null_from_pdf(file))
        else:
            blank_pg = []

        if "Add Watermark" in config_select_manipulation:
            ## go through the pages one after the next
            for current_page in range(len(reader_input.pages)):
                if current_page in to_del:
                    pass
                elif current_page in blank_pg:
                    pass
                else:
                    #if reader_input.pages[current_page].contents is not None:
                    merger = PageMerge(reader_input.pages[current_page])

                    try:
                        mediabox = reader_input.pages[current_page].values()[1]['/Kids'][0]['/MediaBox']

                    except TypeError:
                        mediabox = reader_input.pages[0].values()[1]


                    if mediabox[2] < mediabox[3]:
                        merger.add(watermark_P).render()
                    else:
                        merger.add(watermark_LS).render()
                    writer.addpage(reader_input.pages[current_page])
            status_text.text(f"Watermark done for {file}...")
        else:
            writer.addpages(reader_input.pages)

        if "Remove Metadata" in config_select_manipulation:
            # Remove metadata
            writer.trailer.Info = IndirectPdfDict(
                Title='',
                Author='',
                Subject='',
                Creator='',
            )

        if "Concatenate PDFs" not in config_select_manipulation:
            writer.write(os.getcwd()+"/Intermediate_Data/"+"Annex "+str(i+1)+".pdf")
            writer = PdfWriter()


        status_text.text(f"{file} completed...")
        progress_bar.progress(0.25+(0.75/len(filepath))*(i+1))

    if "Concatenate PDFs" not in config_select_manipulation:
        status_text.text(f"All done!!!")
        st.balloons()
    else:
        # write the modified content to disk
        writer.write(os.getcwd()+"/Intermediate_Data/Annex.pdf")

    st.balloons()

    ## Creating a zipfile for download

    zipObj = ZipFile("download.zip", "w")
    # Add multiple files to the zip
    filepath = find_ext(os.getcwd()+"/Intermediate_Data","pdf")
    for file in filepath:
        zipObj.write(file)
    # close the Zip File
    zipObj.close()

    ZipfileDotZip = "download.zip"

    with open(ZipfileDotZip, "rb") as f:
        bytes = f.read()
        b64 = base64.b64encode(bytes).decode()
        href = f"<a href=\"data:file/zip;base64,{b64}\" download='{ZipfileDotZip}.zip'>\
            Click here to download amended PDFs\
        </a>"
    st.markdown(href, unsafe_allow_html=True)

st.write("-----")
st.write("Once you have selected the required options above, you can click on the button below to start processing. ")

if st.button("Click here to start!"):
    main()
