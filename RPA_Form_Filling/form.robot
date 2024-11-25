*** Settings ***
Library           RPA.Desktop
Library           RPA.recognition

*** Variables ***
${URL}            https://forms.gle/GyFwrpvVawpxRcTVA
${NAME}           John_Doe
${EMAIL}          john.doe@example.com
${ADDRESS}        1234 Elm Street
${PHONE}          123-456-7890
${COMMENTS}       This_is_a_test_comment.
${BROWSER_PATH}    C:\\Users\\mjawa\\AppData\\Local\\BraveSoftware\\Brave-Browser\\Application\\brave.exe

*** Keywords ***
Browser
    RPA.Desktop.Open Application    ${BROWSER_PATH}
    Sleep    3
    Type Text    ${URL}
    Press Keys    ENTER
    Sleep    5
Fill Form
    Press Keys    tab    tab    tab    tab
    Type Text    ${NAME}
    Press Keys    tab
    Type Text    ${EMAIL}
    Press Keys    tab
    Type Text    ${ADDRESS}
    Press Keys    tab
    Type Text    ${PHONE}
    Press Keys    tab
    Type Text    ${COMMENTS}
    Press Keys    tab
    Press Keys    ENTER
*** Tasks ***
Fill
    Browser
    Fill Form