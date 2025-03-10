*** Settings ***
Library    RPA.Desktop
Library    RPA.recognition
Library    RPA.Excel.Files
Library    RPA.Tables

*** Variables ***
${OUTLOOK_PATH}    C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE
${EXCEL_FILE}      C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\emails.csv
${ATTACHED_FILES_PATH}    C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\final_report

*** Keywords ***
Open Outlook
    RPA.Desktop.Open Application    ${OUTLOOK_PATH}
    Sleep    5

Read Excel Data
    Open Workbook    ${EXCEL_FILE}
    ${data}=    Read Worksheet As Table    header=True
    Sleep    5
    Close Workbook
    [Return]    ${data}

Create And Send Email
    [Arguments]    ${Email}    ${cc}    ${Subject}    ${Message}
    RPA.Desktop.Open Application    ${OUTLOOK_PATH}
    Click    (image:D:\\VBPO\\Auromation\\RPA_Outlook_Automation\\images\\NewEmail.png or image:D:\\VBPO\\Auromation\\RPA_Outlook_Automation\\images\\NewEmail1.png)
    Sleep    5
    Type Text    ${Email}
    Sleep    3
    Press Keys    tab
    Press Keys    tab
    Type Text    ${cc}
    Sleep    3
    Press Keys    tab
    Type Text    ${Subject}
    Press Keys    tab
    Type Text    ${Message}
    Sleep    5
    Click    (image:D:\\VBPO\\Auromation\\RPA_Outlook_Automation\\images\\SendEmail.png or image:D:\\VBPO\Auromation\\RPA_Outlook_Automation\\images\\SendEmail1.png)
    Sleep    5
*** Tasks ***
Open Outlook And Click New Email
    Open Outlook

Send Emails From Excel
    ${data}=    Read Excel Data
    FOR    ${row}    IN    @{data}
        ${email}=    Set Variable    ${row}[Email]
        ${cc}=    Set Variable    ${row}[cc]
        ${subject}=    Set Variable    ${row}[Subject]
        ${message}=    Set Variable    ${row}[Message]
        Create And Send Email    ${email}    ${cc}    ${subject}    ${message}
        Sleep    5
    END