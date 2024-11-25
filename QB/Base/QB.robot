*** Settings ***
Library    RPA.Desktop
Library    RPA.recognition
Library    Process


*** Variables ***
${QB_PATH}    C:\\Program Files\\Intuit\\QuickBooks 2024\\QBWPro.exe
${QBW_PROCESS}    QBW

*** Keywords ***
Open QB
    RPA.Desktop.Open Application    ${QB_PATH}

Close QB
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Close.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Close.png)

QB
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\QB.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\QB2.png)
    Sleep    2

Click Open
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\Open.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\Open2.png)

Enter Pswd
    Type Text    text=OCTATE2019

Click Warning OK
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\W_OK2.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\W_OK2.png)

Click OK
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\OK.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\OK2.png)

Click Report
    Sleep    2
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\Report.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\Report2.png) 
    Sleep    1

Click Payrolls
    Sleep    2
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Payrolls.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Payrolls2.png) 
    Sleep    1

Click Summary
    Sleep    2
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Summary.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Summary2.png) 
    Sleep    1

Click Customize
    Sleep    2
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Cusomize.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Cusomize2.png)
    Sleep    1

Click To
    Sleep     2
    Click With Offset    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\To.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\To2.png)    x=110
    Press Keys    backspace    backspace    backspace    backspace   
    Sleep    2
    Press Keys    backspace    backspace    backspace    backspace
    Sleep    2
    Press Keys    backspace    backspace    backspace    backspace
    Sleep    2
    Press Keys    backspace    backspace    backspace    backspace
    Sleep    2
    Type Text    text=20240630

Click From
    Sleep     2
    Click With Offset   (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\From.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\From2.png)    x=110
    Press Keys    backspace    backspace    backspace    backspace   
    Sleep    2
    Press Keys    backspace    backspace    backspace    backspace
    Sleep    2
    Press Keys    backspace    backspace    backspace    backspace
    Sleep    2
    Press Keys    backspace    backspace    backspace    backspace
    Sleep    2
    Type Text    text=20240616

Export Excel
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Excel.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Excel2.png)
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Worksheet.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Worksheet2.png)
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Export.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Export2.png)

Open Excle
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\ExcelLogo2.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\ExcelLogo.png)
    Press Keys    ctrl    s
    Sleep    2

File Name
    Click With Offset    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=B2

Select Path
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Path.png)
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Desktop.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Desktop2.png)

Save Report
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Save.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Save2.png)
    Click    (image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Cross.png or image:C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\images\\Cross2.png)

*** Tasks ***
Open QuickBook
    Open QB
    Sleep    11
    QB
    Click Open
    Sleep    6

Enter Password
    Enter Pswd
    Click OK
    Sleep    11
    Click Warning OK
    Sleep    6
    QB
    Sleep    9

Generate Payrolls
    QB
    Click Report
    Click Payrolls
    Click Summary

Customize Report
    Click From
    Click To
    Press Keys    Enter
    Sleep    6

Export
    Export Excel
    Sleep    7

Save Excel
    Open Excle
    Select Path
    Save Report

Close QuickBooks
    Sleep    5s
    Run Process    powershell    Stop-Process -Name ${QBW_PROCESS} -Force
    Sleep    5s