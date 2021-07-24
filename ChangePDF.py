import streamlit as st
from pdfrw import PdfReader, PdfWriter, PageMerge, IndirectPdfDict
import pathlib
import os
from os import path
from glob import glob
from PIL import Image
import numpy as np
import comtypes.client
from pathlib import Path

import io
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage

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

####################    Streamlit  ####################
def load_image(img):
    im = Image.open(img)
    image = np.array(im)
    return image

st.image(load_image(os.getcwd()+"\Title.png"))
st.write("##")

st.subheader("Choose Options")

config_select_options = st.selectbox("Select option:", ["Input files manually", "Input path"], 0)

if config_select_options == "Input files manually":
    uploaded_file_pdf = st.file_uploader("Upload PDF Files",type=['pdf'], accept_multiple_files=True)
    # uploaded_file_doc = st.file_uploader("Upload doc/docx Files",type=['docx','doc'], accept_multiple_files=True)
else:
    input_path = st.text_input("Please input the path of your folder")
    uploaded_file = []
    if len(input_path) > 0:
        st.write(f"PDFs in this path {input_path} will be uploaded")

output_path = st.text_input("Please input the output path to house your PDFs")
if len(output_path) > 0:
    st.write(f"Amended PDFs will be housed in this path {output_path}")

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

    ## Checking the output directory
    if not os.path.exists('output_path'):
        os.makedirs('output_path')
    status_text.text("Output path checked okay. Proceeding to next step...")

    ## Define the reader and writer objects

    writer = PdfWriter()
    if "Add Watermark" in config_select_manipulation:
        watermark_input_P = PdfReader(uploaded_file_wmp)
        watermark_input_LS = PdfReader(uploaded_file_wml)
        watermark_P = watermark_input_P.pages[0]
        watermark_LS = watermark_input_LS.pages[0]
        status_text.text("Loaded Watermark PDF. Progressing to the next step...")
        progress_bar.progress(0.25)

    def find_ext(dr, ext):
        return glob(path.join(dr,"*.{}".format(ext)))

    wdFormatPDF = 17

    if config_select_options == "Input path":

        filepath_doc = find_ext(input_path,"doc")
        filepath_docx = find_ext(input_path,"docx")
        filepath_all_doc = filepath_doc + filepath_docx

        for file in filepath_all_doc:
            name = Path(file).name.split(".")[0]
            word = comtypes.client.CreateObject('Word.Application', dynamic = True)
            word.Visible = True
            doc = word.Documents.Open(file)
            doc.SaveAs(input_path+"\\"+ name +".pdf", wdFormatPDF)
            doc.Close()
            word.Quit()
        filepath = find_ext(input_path,"pdf")
    else:
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
            writer.write(output_path+"\Annex "+str(i+1)+".pdf")
            writer = PdfWriter()

        status_text.text(f"{file} completed...")
        progress_bar.progress(0.25+(0.75/len(filepath))*(i+1))

    if "Concatenate PDFs" not in config_select_manipulation:
        status_text.text(f"All done!!!")
        st.balloons()
    else:
        # write the modified content to disk
        writer.write(output_path+"\Annex.pdf")

    if config_select_options != "Input files manually":
        for file in filepath_pdf:
            if os.path.exists(file):
                os.remove(file)

    st.balloons()


st.write("-----")
st.write("Once you have selected the required options above, you can click on the button below to start processing. ")

if st.button("Click here to start!"):
    main()
