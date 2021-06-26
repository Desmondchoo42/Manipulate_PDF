# Manipulating PDF

## Introduction

This script utilizes the package Streamlit and PDFRW to add watermark, remove metadata or concatenate PDF or Document (Doc/Docx) and outputting them either into individual PDF or one single PDF (if concatenation option is chosen).  

## How to use
* Main Package
  * Streamlit
  * comtypes (to convert doc/docx to PDF)
  * PDFRW
* How to use
  * Direct to the working directory and run the syntax *streamlit run ChangePDF.py*   
* Script
  * Watermark document
    * Currently the watermark features is done by overlaying another PDF with the watermark on the actual PDF. It has two version, landscape and portrait. User must drag and drop    these files if watermark feature is selected. Script will auto detect the orientation of the PDF and incoporate the right watermark PDF onto the PDF  
  * Input path (doc/docx can only be ran on this option)
    * User has to input the path of the folder and the script will take into account both PDF and doc/docx files within that folder 
  * Input files manually
    * User just drag and drop the required files into the drop down boxes. Option only support PDF files but can handle multiple files at one go
  * Concatenation vs non-concatenation
    * If concatenation option is selected, all PDFs will be save into one PDF call "Annex.pdf" in the prescribed output path
    * If not, individual PDF will be outputted as "Annex 1.pdf, Annex 2.pdf, ..., Annex X.pdf"    

## Possible improvement in the future
* File types
  * Incorporate doc/docx file types into the "Input files manually" option
* Ability to identify size of document (i.e. A4, B5, letter and etc) so that more dynamic watermark can be coded
* Watermark features:
  * Ability for user to choose exactly the location where they want to put the watermark features on the PDF
  * What is the word, font / colour and orientation for the watermark

## Example of the script in Streamlit interface
![alt text](https://github.com/Desmondchoo42/Manipulate_PDF/blob/main/Preview.png?raw=true)
