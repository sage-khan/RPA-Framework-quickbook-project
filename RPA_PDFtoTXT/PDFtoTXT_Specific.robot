*** Settings ***
Library         RPA.PDF
Library         OperatingSystem
Library         String

*** Variables ***
${PDF_FILE}        D:\\VBPO\\Auromation\\RPA_PDFtoTXT\\files\\invoice_1.pdf
${TEXT_FILE}       ./invoice_numbers.txt
${INVOICE_REGEX}   (?i)Invoice no[:.\s]*\s*(\d+)

*** Tasks ***
Extract Invoice Numbers From PDF
    ${text}=    Read Text From PDF    ${PDF_FILE}
    ${invoice_numbers}=    Extract Invoice Numbers    ${text}
    Write Invoice Numbers To File    ${TEXT_FILE}    ${invoice_numbers}
    Log    Invoice numbers extracted and saved to ${TEXT_FILE}

*** Keywords ***
Read Text From PDF
    [Arguments]    ${file_path}
    ${text}=       Get Text From PDF    ${file_path}
    ${ocr_text}=   Get Text From Pdf    ${file_path}
    ${combined_text}=    Set Variable    ${text}\n${ocr_text}
    [Return]       ${combined_text}

Extract Invoice Numbers
    [Arguments]    ${text}
    ${matches}=    Get Regexp Matches    ${text}    ${INVOICE_REGEX}
    ${invoice_numbers}=    Create List
    FOR    ${match}    IN    @{matches}
        ${invoice_numbers}=    Create List    ${invoice_numbers}    ${match[1]}
    END
    IF    '${invoice_numbers}' == '[]'
        Log    No invoice numbers found.    WARN
    END
    [Return]    ${invoice_numbers}

Write Invoice Numbers To File
    [Arguments]    ${file_path}    ${invoice_numbers}
    FOR    ${number}    IN    @{invoice_numbers}
        Append To File    ${file_path}    ${number}\n
    END
