*** Settings ***
Library    RPA.Desktop
Library    RPA.recognition
Library    Process

*** Variables ***
${QB_PATH}    C:\\Program Files\\Intuit\\QuickBooks 2024\\QBWPro.exe
${TO_DATE}	20240330
${FROM_DATE}	20240316
${QBW_PROCESS}    QBW
${EXCEL_PROCESS}    EXCEL

*** Keywords ***
Open QB
    RPA.Desktop.Open Application    ${QB_PATH}

Close QB
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Close.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Close.png)

QB
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\QB.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\QB2.png)
    Sleep    2

Click Open
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\Open.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\Open2.png)

Enter Pswd
    Type Text    text=OCTATE2019

Click Warning OK
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\W_OK2.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\W_OK2.png)
    Sleep    1
Click OK
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\OK.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\OK2.png)
    Sleep    1

Click Report
    Sleep    2
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\Report.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\Images\\Report2.png) 
    Sleep    1

Click Payrolls
    Sleep    2
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Payrolls.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Payrolls2.png) 
    Sleep    1

Click Summary
    Sleep    1
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Summary.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Summary2.png) 
    Sleep    1

Click Customize
    Sleep    2
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Cusomize.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Cusomize2.png)
    Sleep    1

Click Filters
    Sleep    1
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filters.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filters2.png)
    Sleep    1

Click Class
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Class.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Class2.png)
    Sleep    1

Class Select Drop
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Class_Select.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Class_Select2.png)
    Sleep    1

Click All Classes
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\All_Classes.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\All_Classes2.png)
    Sleep    1

Click Customize_OK
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\OK.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\OK2.png)

Click Show_Column
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Show_Columns.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Show_Columns2.png)    x=90
    Sleep    1

Click Scroll
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Scroll.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Scroll2.png)    action=double_click

Total Only
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Total_Only.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Total_Only2.png)
    Sleep    1

Click Class_Column
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Class_Column.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Class_Column2.png)
    Sleep    1
    
#   Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\.png)

Click To
    Sleep     2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\To.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\To2.png)    x=110
    Press Keys    backspace    backspace    backspace    backspace   
    Press Keys    backspace    backspace    backspace    backspace
    Press Keys    backspace    backspace    backspace    backspace
    Press Keys    backspace    backspace    backspace    backspace
    Press Keys    backspace    backspace    backspace    backspace
    Type Text    text=${TO_DATE}

Click From
    Sleep     2
    Click With Offset   (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\From.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\From2.png)    x=110
    Press Keys    backspace    backspace    backspace    backspace   
    Press Keys    backspace    backspace    backspace    backspace
    Press Keys    backspace    backspace    backspace    backspace
    Press Keys    backspace    backspace    backspace    backspace
    Press Keys    backspace    backspace    backspace    backspace
    Type Text    text=${FROM_DATE}

Export Excel
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Excel.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Excel2.png)
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Worksheet.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Worksheet2.png)
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Export.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Export2.png)

Open Excel
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\ExcelLogo3.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\ExcelLogo4.png)
    Sleep    8
    Press Keys    ctrl    s

Select Path
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Path.png)
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Desktop.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Desktop2.png)

Save Report
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Save.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Save2.png)
    Sleep    2
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Cross.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Cross2.png)

Export
    Export Excel
    Sleep    7

Show Columns
    Click Show_Column
    Sleep    1

Customize Report
    Click Customize
    Sleep    1
    Click Filters
    Sleep    1
    Click Class
    Sleep    1

Click Glendale
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Glendale.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Glendale2.png)

Click Thorold
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Thorold.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Thorold2.png)

Click Welland
   Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Welland.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Welland2.png)

Click Welland Ave
   Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Welland_Ave.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Welland_Ave2.png)

Click Mcleod Rd
   Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Mcleod_Rd.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Mcleod_Rd2.png)

Click Multiple Location#1
   Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Multiple_Location#1.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Multiple_Location#12.png)

Click Multiple Location#2
   Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Multiple_Location#2.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Multiple_Location#22.png)

Set Date
    Click From
    Click To
    Press Keys    Enter
    Sleep    2

Click Memorized Report
    Sleep    1
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Memorized_Report.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Memorized_Report2.png)
    Sleep    1

Click Accountant
    Sleep    1
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Accountant.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Accountant2.png)
    Sleep    1

Click Pay Cheque Details
    Sleep    1
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Pay_Cheque_Details.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Pay_Cheque_Details2.png)
    Sleep    1

Click Payroll Deduction
    Sleep    1
    Click    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Payroll_Deduction.png or image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Payroll_Deduction2.png)
    Sleep    1

*** Tasks ***
# Open QuickBook
#     # Open QB
#     # Sleep    11
#     QB
#     # Click Open
#     Sleep    6

# Enter Password
#     Enter Pswd
#     Click OK
#     Sleep    12
#     Click Warning OK
#     Sleep    9
#     QB
#     Sleep    9

Generate Payrolls
    QB
    Click Report
    Click Payrolls
    Click Summary
    Set Date
    
Select All Class
    Customize Report
    Sleep    1
    Class Select Drop
    Sleep    1
    # Click All Classes
    # Sleep    1
    Click Customize_OK
    Sleep    1

Total Only Report
    Show Columns
    Sleep    1
    Click Scroll
    Sleep    1
    Total Only
    Sleep    1

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Total Only Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Class Report
    Show Columns
    Sleep    1
    Click Class_Column
    Sleep    1

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Class Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Glendale Report
    Customize Report
    Sleep    1
    Class Select Drop
    Sleep    1
    Click Glendale
    Sleep    1
    Click Customize_OK
    Sleep    1

    Export
    Sleep    2
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Glendale Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Thorold Report
    Customize Report
    Sleep    1
    Class Select Drop
    Sleep    1
    Click Thorold
    Sleep    1
    Click Customize_OK
    Sleep    1

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Thorold Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Welland Report
    Customize Report
    Sleep    1
    Class Select Drop
    Sleep    1
    Click Welland
    Sleep    1
    Click Customize_OK
    Sleep    1

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Welland Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Welland Ave Report
    Customize Report
    Sleep    1
    Class Select Drop
    Sleep    1
    Click Welland Ave
    Sleep    1
    Click Customize_OK
    Sleep    1

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Welland Ave Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Mcleod Rd Report
    Customize Report
    Sleep    1
    Class Select Drop
    Sleep    1
    Click Mcleod Rd
    Sleep    1
    Click Customize_OK
    Sleep    1
    
    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Mcleod Rd Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Multiple Location#1 Report
    Customize Report
    Sleep    1
    Class Select Drop
    Sleep    1
    Click Multiple Location#1
    Sleep    1
    Click Customize_OK
    Sleep    1

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Multiple Location#1 Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Multiple Location#2 Report
    Customize Report
    Sleep    1
    Class Select Drop
    Sleep    1
    Click Multiple Location#2
    Sleep    1
    Click Customize_OK
    Sleep    1

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Multiple Location#2 Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Payroll Deduction
    Click Report
    Click Memorized Report
    Click Accountant
    Click Payroll Deduction
    Set Date

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Payroll Deduction Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

Pay Cheque Details
    Click Report
    Click Memorized Report
    Click Accountant
    Click Pay Cheque Details
    Set Date

    Export
    Sleep    2
    Open Excel
    Sleep    2
    Click With Offset    (image:C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\images\\Filename.png)    x=10    y=-10
    Type Text    text=Pay Cheque Details Report ${FROM_DATE} - ${TO_DATE}
    Sleep    2
    Select Path
    Save Report

# Close QuickBooks
#     Sleep    5s
#     Run Process    powershell    Stop-Process -Name ${QBW_PROCESS} -Force
#     Sleep    5s

# Close Excel
#     Sleep    5s
#     Run Process    powershell    Stop-Process -Name ${EXCEL_PROCESS} -Force
#     Sleep    5s
#     QB
