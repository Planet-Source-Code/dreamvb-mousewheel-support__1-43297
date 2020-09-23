VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Useing the mouse wheel in vb"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   195
      TabIndex        =   1
      Text            =   "Place your mouse here and use the mouse wheel see the value change"
      Top             =   540
      Width           =   6405
   End
   Begin VB.Label lblmouseval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mouse wheel by Ben Jones
' Email dreamvb@yahoo.com
' I hope you find this code helpfull

Private Sub Command1_Click()
    Unload Form1    ' unload the form
End Sub

Private Sub Form_Load()
    OldProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
    SetWindowLong txt.hWnd, GWL_WNDPROC, AddressOf TWndProc ' Subclass the textbox
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong txt.hWnd, GWL_WNDPROC, OldProc
    Set Form1 = Nothing ' Release the form from memory
End Sub


