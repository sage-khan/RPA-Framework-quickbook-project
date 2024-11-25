*** Settings ***
Library           RPA.PDF
Library           OperatingSystem
Library           Process

*** Variables ***
${PDF_FILE}       D:\\VBPO\\Auromation\\RPA_PDF_Images\\a2.pdf
${OUTPUT_FOLDER}  D:\\VBPO\\Auromation\\RPA_PDF_Images\\images

*** Keywords ***
*** Keyword ***
Figures to Images
    ${image_filenames} =    Save figures as images
    ...             source_path=${PDF_FILE}
    ...             images_folder=${OUTPUT_FOLDER}
    ...             pages=${1}
    ...             file_prefix=output
*** Tasks ***
Extract Images
    Figures to Images
