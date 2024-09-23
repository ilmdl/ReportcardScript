#Requires AutoHotkey v2.0

check := 1
len := 1000
; clipboard := ""

; f12::reload
!e::reload
#^z::{ ; Windows ctrl z
    loop 31 {
        ;Get name
        WinActivate "Book1 - Excel"
        ; if (check = 0)
        ;     break
        Sleep len
        Send '{f2}' ;copy
        Sleep len
        Send '+{Home}'
        Sleep len
        Send '^c' ;copy

        Sleep len
        Send '{Enter}'
        Sleep len
        ; MsgBox A_Clipboard
        
        ; Send "{Alt down}{Tab}{Alt up}" ;Switch app
        WinActivate "REPORT CARD TERM 3 2024 - Google Sheets - Google Chrome"
        Sleep len
        
        ;Create student file
        Send "{Up}"         ;Return to selection box
        Sleep len
        Send A_Clipboard   ;Type student name
        Sleep len
        Send '{Enter}'     ;Select Name
        Sleep len
        Send '{Ctrl down}{p}{Ctrl up}'         ;Open print settings
        Sleep len
        Send '{Tab}'        ;Select Next Button  
        Send '{Enter}'
        Sleep len * 5
        Send '{Enter}'
        Sleep len * 6
        Send A_Clipboard
        Sleep len
        Send ' REPORT CARD END TERM 3 2024'
        Sleep len
        Send '{Enter}'
        Sleep len

        ; Send "{Alt down}{Tab}{Alt up}" ;Switch app
    }
}
