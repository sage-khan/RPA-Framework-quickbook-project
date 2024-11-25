*** Settings ***
Library         RPA.PDF
Library         OperatingSystem

*** Variables ***
${PDF_FILE}     D:\\VBPO\\Auromation\\RPA_PDFtoTXT\\files\\invoice_1.pdf
${TEXT_FILE}    ./textfile.txt

*** Variables ***
${PDF_FILE}     path/to/your/document.pdf
${TEXT_FILE}    path/to/output/textfile.txt

*** Tasks ***
Extract Text From PDF
    ${text}=    Read Text From PDF    ${PDF_FILE}
    Write Text To File    ${TEXT_FILE}    ${text}
    Log    Text extracted and saved to ${TEXT_FILE}

*** Keywords ***
Read Text From PDF
    [Arguments]    ${file_path}
    ${pages}=      Get Text From PDF    ${file_path}
    ${text}=       Set Variable         ${EMPTY}
    FOR    ${page}    ${text_content}    IN    &{pages}
        ${text}=    Catenate    ${text}    ${text_content}
    END
    [Return]    ${text}

Write Text To File
    [Arguments]    ${file_path}    ${text}
    Create File    ${file_path}
    Append To File    ${file_path}    ${text}