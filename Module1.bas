Attribute VB_Name = "Module1"
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_MOUSEWHEEL = &H20A ' window message for mouse wheel
Private MouseWheelUp As Boolean     ' true mouse wheel up false if down
Private I As Long                   ' used to hold our counter value
Public Const GWL_WNDPROC = (-4)
Public OldProc As Long              ' Holds the old TWndProc

Public Function TWndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If wMsg = WM_MOUSEWHEEL Then ' Have we got a mouse wheel message
    
    If wParam > 0 Then MouseWheelUp = True Else MouseWheelUp = False
    
    Select Case MouseWheelUp
        Case True ' mouse up value is found
            Form1.lblmouseval.Caption = "The value of the mouse wheel is set to Up" ' update the label caption
            I = I + 1 ' update our counter
        Case False ' mouse value down is found
            Form1.lblmouseval.Caption = "The value of the mouse wheel is set to Down" ' update the label caption
            I = I - 1 ' update our counter
            If I <= 0 Then I = 0 'reset our counter if below zero
    End Select
        Form1.txt.Text = "The value of the mouse wheel is " & I ' update the text in the textbox
    End If
    
    TWndProc = CallWindowProc(OldProc, hWnd, wMsg, wParam, lParam)

End Function
