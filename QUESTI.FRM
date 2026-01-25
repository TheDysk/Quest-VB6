VERSION 5.00
Begin VB.Form nquesti 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Command"
   ClientHeight    =   1095
   ClientLeft      =   1200
   ClientTop       =   6885
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1095
   ScaleWidth      =   6180
   Begin VB.CheckBox ShowDebug 
      Caption         =   "Show Debug Data"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox InputLine 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "nquesti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const MF_BYCOMMAND = &H0&
Private Const SC_CLOSE = &HF060&

' QUESTI.FRM
Option Explicit


Private Sub InputLine_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   gv0090 = InputLine.Text
   InputLine.Text = ""
   Unload Me
End If

End Sub

Private Sub Form_Load()
    

    Dim hMenu As Long
    hMenu = GetSystemMenu(Me.hwnd, False)
    Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    
    ShowDebug.Value = IIf(DebugData.Visible, 1, 0)
    
End Sub



Private Sub ShowDebug_Click()
    ' Check if the checkbox is checked (value of 1)
    If ShowDebug.Value = vbChecked Then
        DebugData.Visible = True
        ' Bring it to the front if it was hidden behind other windows
        DebugData.ZOrder 0
    Else
        DebugData.Visible = False
    End If

End Sub

