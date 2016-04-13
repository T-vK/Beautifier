#SingleInstance Force

;Default gui element sizes
elementWidth := 800
codeBoxHeight := 850
buttonHeight := 30

;Create hidden internet explorer
wb := ComObjCreate("InternetExplorer.Application")
OnExit("CleanUp") ;kill the internet explorer again on exit
wb.Visible := False ;make sure it's hidden
wb.Navigate("http://jsbeautifier.org/?without-codemirror") ;navigate to that url
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ;wait until the site is ready
    Sleep, 10

;Create GUI
Gui, Font, s11 ;font size 11
Gui, +Resize ;allow resize

buttonProps :=  "v" "BeautifyButton" ;button id
buttonProps .= " g" "BeautifyIt" ;label to execute on click
buttonProps .= " x" 0 ;set position/size:
buttonProps .= " y" codeBoxHeight-buttonHeight
buttonProps .= " w" elementWidth
buttonProps .= " h" buttonHeight
Gui, Add, Button, %buttonProps%, Beautify it!

Gui, Font,, Consolas ;use Consolas font for the edit box
codeBoxProps :=  "v" "CodeBox" ;code box id
codeBoxProps .= " x" 0 ;set position/size:
codeBoxProps .= " y" 0
codeBoxProps .= " w" elementWidth
codeBoxProps .= " h" codeBoxHeight-buttonHeight
codeBoxProps .= " -Wrap +HScroll"
Gui, Add, Edit, %codeBoxProps%

Gui, Show ;Display the GUI
Return

GuiSize(guiHwnd, eventInfo, newWidth, newHeight) {
    Global buttonHeight ;get global variable
    
    newCodeBoxSize :=  "w" newWidth
    newCodeBoxSize .= " h" newHeight-buttonHeight
    GuiControl, Move, CodeBox, %newCodeBoxSize%
    
    newButtonSize :=  "y" newHeight-buttonHeight
    newButtonSize .= " w" newWidth
    newButtonSize .= " h" buttonHeight
    GuiControl, Move, BeautifyButton, %newButtonSize%
}

BeautifyIt(a:="",b:="",c:="",d:="") {
    Global wb ;get global variable
    GuiControlGet, CodeBox ;get content of the code box
    ;Correct CR/LF escapes:
    code := StrReplace(CodeBox, "\r\n", "`n")
    code := StrReplace(code, "\r", "`n")
    code := StrReplace(code, "\n", "`n")
    code := StrReplace(code, "\t", "`t")
    GuiControl,, CodeBox, %code% ;display changes
    
    wb.document.getElementById("source").click()
    wb.document.getElementById("source").value := code " "
    wb.document.getElementsByClassName("submit")[1].click()
    ;wait until code has been beautified
    While (wb.document.getElementById("source").value = code)
        Sleep, 100
    GuiControl,, CodeBox, % wb.document.getElementById("source").value
}

GuiClose() {
    CleanUp()
}

CleanUp() {
    Global wb ;get global variable
    wb.Quit
    ExitApp
}