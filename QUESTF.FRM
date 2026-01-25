VERSION 5.00
Begin VB.Form nquestf 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Select Saved Game"
   ClientHeight    =   1485
   ClientLeft      =   3135
   ClientTop       =   5355
   ClientWidth     =   3450
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1485
   ScaleWidth      =   3450
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   120
      Pattern         =   "*.sav"
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "nquestf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' QUESTF.FRM
Option Explicit

Private Sub File1_Click()
gv0094 = File1.FileName
End Sub

Private Sub File1_DblClick()
Unload Me
End Sub
