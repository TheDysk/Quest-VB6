VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DebugData 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug Data"
   ClientHeight    =   4905
   ClientLeft      =   10545
   ClientTop       =   1575
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   12375
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Save Debug Data to txt file"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear All Debug Data"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Selected to Clipboard"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "DebugData"
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


Public Sub DebugAdd(recNum As String, recType As String, recDesc As String, recData As String)

    Dim itm As ListItem

    Set itm = ListView1.ListItems.Add(, , recNum)
    itm.SubItems(1) = recType
    itm.SubItems(2) = recDesc
    itm.SubItems(3) = recData

    ' Auto-scroll to newest entry
    ListView1.ListItems(itm.Index).EnsureVisible

End Sub


Private Sub Command1_Click()

    Dim itm As ListItem
    Dim s As String
    Dim i As Integer, j As Integer

    If ListView1.ListItems.Count = 0 Then Exit Sub

    For i = 1 To ListView1.ListItems.Count
        Set itm = ListView1.ListItems(i)

        If itm.Selected Then
            Dim line As String
            line = itm.Text

            For j = 1 To ListView1.ColumnHeaders.Count - 1
                line = line & vbTab & itm.SubItems(j)
            Next j

            s = s & line & vbCrLf
        End If
    Next i

    If Len(s) > 0 Then
        Clipboard.Clear
        Clipboard.SetText s
    End If

End Sub



Private Sub Command2_Click()

    ListView1.ListItems.Clear
    
End Sub

Private Sub Command3_Click()

    Dim i As Integer, j As Integer
    Dim itm As ListItem
    Dim s As String
    Dim f As Integer

    f = FreeFile
    Open App.Path & "\debug_output.txt" For Output As #f

    For i = 1 To ListView1.ListItems.Count
        Set itm = ListView1.ListItems(i)

        s = itm.Text
        For j = 1 To ListView1.ColumnHeaders.Count - 1
            s = s & vbTab & itm.SubItems(j)
        Next j

        Print #f, s
    Next i

    Close #f

End Sub

Private Sub Form_Load()

    Dim hMenu As Long
    hMenu = GetSystemMenu(Me.hwnd, False)
    Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    
    With ListView1.ColumnHeaders
        .Add , , "Rec#", 1000 ' 4-digit record number
        .Add , , "Type", 800 ' 1-digit record type
        .Add , , "Description", 1500 ' Up to 8 chars
        .Add , , "Data", 8300 ' Up to 100 chars
    End With
    
End Sub


