*** Settings ***
Library    RPA.Desktop
Library    RPA.recognition

*** Variables ***
${PPT_PATH}    C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.exe

*** Keywords ***
Open PPT
    RPA.Desktop.Open Application    ${PPT_PATH}
    Sleep    2

add new slide
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\NewSlide.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\NewSlide1.png)
    Sleep    1

Save PPT
    Sleep    1
    Click       (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\File.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\File1.png)
    Sleep    1
    Click       (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Save.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Save1.png)
    Sleep    1
    Click       (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Browse.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Browse1.png)
    Sleep    1
    Click       (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\SaveA.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\SaveA1.png)

Centure Allign
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Centure.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Centure1.png)

Slide Title
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\TitleN.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\TitleN1.png)

Slide Body
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Body.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Body1.png)

Animation
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Animation.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Animation1.png)

Fade Animation
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Fade.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Fade1.png)

Transition
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Transition.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Transition1.png)

Push Transition
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Push.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Push1.png)

Home
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Home.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Home1.png)

Create PPT
    # Create PPT
    Sleep    1
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\BlankPPT.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\BlankPPT1.png)

 Make PPT
    # Title Page
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Title.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\Title1.png)
    Type Text    Creating PPT with RPA Framework
    Sleep    2
    Click    (image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\SubTitle.png or image:D:\\VBPO\\Auromation\\RPA_PPT\\images\\SubTitle1.png)
    Type Text    By Muhammad Jawad Khan

    # Slide 1
    add new slide
    Sleep    1
    Slide Title
    Centure Allign
    Type Text    Introduction to RPA Framework for Automation
    Slide Body
    Type Text    A robust tool for creating RPA solutions
    Press Keys    Enter
    Type Text    Supports automating processes across various applications.
    Press Keys    Enter
    Type Text    Provides an easy-to-use interface for designing automation workflows.
    Press Keys    Enter
    Type Text    Uses OCR and image processing technology to screen reading.
    Sleep    1
    Animation
    Fade Animation
    Home
    Sleep    1
    Transition
    Push Transition
    Home
    
    # Slide 2
    add new slide
    Sleep    1
    Slide Title
    Centure Allign
    Type Text    Benefits of Automation with RPA
    Slide Body
    Type Text    Reduces manual effort and human errors.
    Press Keys    Enter
    Type Text    Increases productivity and operational efficiency.
    Press Keys    Enter
    Type Text    Ensures timely and consistent delivery.
    Sleep    1
    Animation
    Fade Animation
    Home
    Sleep    1
    Transition
    Push Transition
    Home
    
    # Slide 3
    add new slide
    Sleep    1
    Slide Title
    Centure Allign
    Type Text    Steps to Automate Presentations Using RPA Framework
    Slide Body
    Type Text    Take images of the buttons you need to cick and save it in a commin directory
    Press Keys    Enter
    Type Text    Write the .robot script
    Press Keys    Enter
    Type Text    Execute in CLI
    Press Keys    Enter
    Type Text    Considering the currenet script it will:
    Press Keys    Enter
    Type Text    Open new PowerPoint Presentation
    Press Keys    Enter
    Type Text    Open a blank PPT slide
    Press Keys    Enter
    Type Text    Add Title and author name
    Press Keys    Enter
    Type Text    Click on new slide
    Press Keys    Enter
    Type Text    Write relavent pre-coded information in the Title and the body and allign the title to centure and repeat
    Press Keys    Enter
    Type Text    After that it will save the Presentation
    Sleep    1
    Animation
    Fade Animation
    Home
    Sleep    1
    Transition
    Push Transition
    Home
    
    # Slide 4
    add new slide
    Sleep    1
    Slide Title
    Centure Allign
    Type Text    Key Words to Use
    Sleep    1
    Slide Body
    Type Text    Use "Click" keyword to click the button by comparing the input on screen to the image provided
    Press Keys    Enter
    Type Text    Use "Type Text" keyword to enter or write text
    Press Keys    Enter
    Type Text    Use "Press Keys" keyword to use direct keyboard keys
    Press Keys    Enter
    Type Text    Use "Sleep" keyword to add delay to actions to counter system delays
    Sleep    1
    Animation
    Fade Animation
    Home
    Sleep    1
    Transition
    Push Transition
    Home

    # Slide 5
    add new slide
    Sleep    1
    Slide Title
    Centure Allign
    Type Text    Thank You
    Sleep    1
    Transition
    Push Transition
    Home

*** Tasks ***
Open PowerPoint
  Open PPT

Create and Make PPT
    Create PPT
    Make PPT

# Save PPT
#     Save PPT