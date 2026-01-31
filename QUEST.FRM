VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Quest 
   Appearance      =   0  'Flat
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest"
   ClientHeight    =   6750
   ClientLeft      =   555
   ClientTop       =   1575
   ClientWidth     =   9795
   FillColor       =   &H00008000&
   BeginProperty Font 
      Name            =   "@Malgun Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "QUEST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6750
   ScaleWidth      =   9795
   Begin VB.TextBox FocusSink 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   6360
      Width           =   375
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
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   5760
      Width           =   6495
   End
   Begin VB.CheckBox ShowDebug 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      Caption         =   "Show Debug Data"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   6360
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4575
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8070
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"QUEST.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox scorebox 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "  Quest   Click to run"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      TabIndex        =   0
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Command"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   3120
      Y1              =   2040
      Y2              =   4200
   End
End
Attribute VB_Name = "Quest"
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
Private IsLoading As Boolean

' QUEST.FRM
Option Explicit
Dim AL(60) As Integer
Dim AR(60) As Integer
Dim GL(60) As Integer
Dim GR(60) As Integer
Dim GF(60) As Integer
Dim RP(5) As Integer
Dim m00B6 As T0109
Dim m00D0 As T013D
Dim m0136 As T013D
Dim m019C As T014E


Private Sub Command1_Click()

Dim RecTypeTxt$
Dim recType$
Dim DirPath As String
Dim RN As Integer
Dim CC As Integer
Dim A2 As String
Dim R As Integer
Dim RC As String
Dim PN As Integer
Dim SI As Integer
Dim HO As Integer
Dim WR As Integer
Dim DK As Integer
Dim FG As Integer
Dim NA As Integer
Dim NG As Integer
Dim M As Integer
Dim NewGame As String
Dim X As Integer
Dim ValidFile As String
Dim ARR As String ' AR$ (string variable) changed to ARR$ as there is also a AR% (integer variables)
Dim S2 As String
Dim S1 As String
Dim G As Integer
Dim ST As Integer
Dim TX As String
Dim NP As Integer
Dim PL As String
Dim PS As Integer
Dim OM As Integer
Dim TG As Integer
Dim Y As Integer
Dim CO As String
Dim LW As String
Dim L1 As String
Dim W1 As String
Dim L2 As String
Dim W2 As String
Dim AC As String
Dim DC As String
Dim KG As String
Dim s As String
Dim MsgText As String
Dim Score As Variant
Dim MM As String ' M$ (string variable) changed to MM$ as there is also a M% (integer variables)
Dim LV As Integer
Dim N As String
Dim NN As Integer ' N% (integer variable) changed to NN% as there is also a N$ (string variables)
Dim CR As String
Dim STT As String ' ST$ (string variable) changed to STT$ as there is also a ST% (integer variables)
Dim l026C As Integer
Dim TY As String
Dim FL As Integer
Dim CA As String
Dim NM As Integer
Dim RR As String
Dim OB As String
Dim SV As Integer
Dim l028E As Variant
Dim l0292 As String
Dim l0294 As String
Dim l0296 As String
Dim l0298 As String
Dim l029A As Variant
Dim l029E As Variant
Dim l02A2 As Variant
Dim xxxx As Variant
Dim l02AA As Variant
Dim l02AE As Variant
            
      Load DebugData ' Form is now in memory but invisible
      DebugData.Visible = False ' Make sure its visible property is set false
      
      Command1.Visible = False
      GoSub 12000:
      Randomize (Val(Right$(Time$, 2)))
      ' DirPath$ = "C:\Users\TheDysk\Desktop\temp\"
      
      'Get the path where the exe has run from. The datafiles should be in the same folder
      If Right$(App.Path, 1) = "\" Then
        DirPath$ = App.Path
      Else
        DirPath$ = App.Path & "\"
      End If
      
      Open DirPath$ + "ADATA1" For Random As #1 Len = 100
      Open DirPath$ + "ADATA2" For Random As #2 Len = 100
      RN% = 1: CC% = 0: A2$ = "0"

100
101   R% = 0: GoSub 9010
102   PN% = -Val(Mid$(RC$, 82, 4))
105   SI% = Val(Mid$(RC$, 78, 4)): HO% = Val(Mid$(RC$, 25, 4)): WR% = Val(Mid$(RC$, 29, 4))
110   DK% = Val(Mid$(RC$, 33, 4)): R% = Val(Mid$(RC$, 11, 4)): FG% = Val(Mid$(RC$, 17, 4))
115   NA% = Val(Mid$(RC$, 9, 2)): NG% = Val(Mid$(RC$, 15, 2)): M% = Val(Mid$(RC$, 74, 4))
116   If NewGame <> "Y" Then GoTo 185

117   SI% = Val(Mid$(RC$, 21, 4)): PN% = 0: M% = 100
150   GoSub 9010: X% = Val(Mid$(RC$, 91, 2))
155   AL%(X%) = Val(Mid$(RC$, 93, 4)): AR%(X%) = R%
160   R% = Val(Left$(RC$, 4)): If R% > 0 Then GoTo 150 Else R% = FG%
170   GoSub 9010: X% = Val(Mid$(RC$, 91, 2))
175   GL%(X%) = Val(Mid$(RC$, 93, 4)): GR%(X%) = R%: GF%(X%) = Val(Mid$(RC$, 81, 1))
180   R% = Val(Left$(RC$, 4)): If R% > 0 Then GoTo 170
181   GoTo 190

185   GoSub 16000: If ValidFile$ = "NOWAY" Then ValidFile$ = "": GoTo 117

190

191   ARR$ = "N":
192   GoTo 600

500   ARR$ = "Y": S2$ = "The": S1$ = "X":
505   For G% = 1 To NG%
510   If GL%(G%) <> SI% Then GoTo 550
515   R% = GR%(G%): GoSub 9010: ST% = Val(Mid$(RC$, 9, 1))
520   If Val(Mid$(RC$, 80 + 2 * ST%, 1)) <= Rnd(RN%) * 9 Then GoTo 550
525   TX$ = S2$: GoSub 9230
530   If S2$ = "The" Then S1$ = "%4has" Else S1$ = "%4have"
535   S2$ = "%4and": GoSub 8010: GL%(G%) = -NP%
550   Next G%
560   SI% = NP%: If S1$ = "X" Then GoTo 600
565   TX$ = S1$ & " followed you. ": GoSub 9230: ARR$ = "N"

600   R% = SI%: GoSub 9010: PL$ = RC$: R% = Val(Left$(PL$, 4)):
605   For X% = 1 To 5
610   If SI% = RP%(X%) Then R% = Val(Mid$(PL$, 5, 4))
615   RP%(X% - 1) = RP%(X%)
620   Next X%
625   RP%(5) = SI%: GoSub 9120
630   PS% = Val(Mid$(PL$, 9, 1)): R% = Val(Mid$(PL$, 6 + 4 * PS%, 4)): GoSub 9120
635   OM% = 9999: For G% = 1 To NG%
645   If GL%(G%) = SI% Then R% = GR%(G%): GoSub 8100
647   If GL%(G%) = -SI% Then GL%(G%) = SI%
650   Next G%
660   For G% = 1 To NA%
665   If AL%(G%) = SI% Then R% = AR%(G%): GoSub 8100
670   Next G%
700   If ARR$ = "Y" Then ARR$ = "N": GoTo 900 Else TG% = 0
705   For X% = 1 To NG%
710   If GL%(X%) <> SI% Then GoTo 720
715   If Rnd(RN%) * 9 < GF%(X%) Then TG% = X%: GF%(X%) = 9
717   GF%(X%) = GF%(X%) + 1
720   Next X%
730   If TG% = 0 Then GoTo 900 Else R% = GR%(TG%): GoSub 9010
735   GF%(TG%) = Val(Mid$(RC$, 86, 1))
740   ST% = Val(Mid$(RC$, 9, 1))
745   R% = Val(Mid$(RC$, 87, 4)): If R% = 0 Then GoTo 900
750   Y% = Val(Mid$(RC$, 81 + 2 * ST%, 1)): GoSub 9010: ARR$ = "Y"
755   If Rnd(RN%) * 9 < Y% Then GoTo 3500 Else GoTo 3600
900   GoSub 9850

910
      GoSub 7000: RN% = Len(CO$): M% = M% + 1
920   GoSub 8200: L1$ = LW$
930   If Len(L1$) = 0 Then W1$ = "*   " Else W1$ = Left$(L1$ & "   ", 4)
940   GoSub 8200
950   L2$ = LW$: GoSub 8200: If Len(LW$) > 0 Then GoTo 950
960   If Len(L2$) = 0 Then W2$ = "*   " Else W2$ = Left$(L2$ & "   ", 4)
970   For X% = 1 To 4
980   If Mid$(W1$, X%, 1) > "@" Then Mid$(W1$, X%, 1) = Chr$(Asc(Mid$(W1$, X%, 1)) And &H5F)
990   If Mid$(W2$, X%, 1) > "@" Then Mid$(W2$, X%, 1) = Chr$(Asc(Mid$(W2$, X%, 1)) And &H5F)
995   Next X%
1000  CO$ = "?": AC$ = "?"
1030  DC$ = "NORTSOUTEASTWESTUP  DOWN" & Mid$(PL$, 26, 8) & "ENTEEXIT"
1033  For X% = 1 To 37 Step 4
1035  If W1$ = Mid$(DC$, X%, 4) Then CO$ = "C" & Mid$("01020304050607080910", (X% + 1) / 2, 2)
1040  Next X%: If CO$ <> "?" Then GoTo 2000
1042  If Mid$(PL$, 26, 4) <> "* KR" Then GoTo 1050
1044  R% = Val(Mid$(PL$, 30, 4)): GoSub 8300: If AC$ <> "?" Then GoTo 3000
1050  G% = NG%
1060  If GL%(G%) <> SI% Then GoTo 1100
1070  R% = GR%(G%): GoSub 9010: If Mid$(RC$, 26, 4) <> "* KR" Then GoTo 1100
1080  R% = Val(Mid$(RC$, 30, 4)): GoSub 8300
1090  If AC$ <> "?" Then GoTo 3000
1100  G% = G% - 1: If G% > 0 Then GoTo 1060
1110  Y% = 0: For X% = 1 To 25 Step 4
1120  If Mid$("LOOKINVEFEEDSCOREND ATTAKILL", X%, 4) = W1$ Then Y% = (X% + 3) / 4
1130  Next X%: G% = 1
1140  On Y% GoTo 1560, 1500, 1410, 1570, 1600, 1400, 1410
1150  If AL%(G%) <> SI% And AL%(G%) <> 0 Then GoTo 1190 Else R% = AR%(G%): GoSub 9010
1160  If Mid$(RC$, 26, 4) <> "* KR" Then GoTo 1180
1170  R% = Val(Mid$(RC$, 30, 4)): GoSub 8300: If AC$ = "?" Then GoTo 1190 Else GoTo 3000
1180  If Mid$(RC$, 26, 4) = W2$ Or Mid$(RC$, 30, 4) = W2$ Then GoTo 1210
1190  G% = G% + 1: If G% <= NA% Then GoTo 1150
1200  R% = DK%: GoSub 8300: GoTo 3000:
1210  If W1$ = "CARR" Then GoTo 1330
1211  If W1$ = "GET " Then GoTo 1330
1212  If W1$ = "TAKE" Then GoTo 1330
1220  If W1$ = "DROP" Then GoTo 1300
1230  If W1$ = "THRO" Then GoTo 1300
1250  R% = AR%(G%): GoSub 9010: CO$ = "X": For X% = 67 To 75 Step 4
1260  If W1$ = Mid$(RC$, X%, 4) Then CO$ = Right$(Str$((X% - 63) / 4), 1)
1270  Next X%: If CO$ = "X" Then R% = 338: GoSub 9120: GoTo 700
1280  CO$ = CO$ & Mid$(Str$(G% / 100) & "0", 3, 2): GoTo 2000
1300  If AL%(G%) <> 0 Then R% = 339: GoSub 9120: GoTo 700
1310  CC% = CC% - 1: GoSub 8400: If W1$ = "DROP" Then GoTo 700 Else CO$ = "4": GoTo 1280
1330  If AL%(G%) = 0 Then R% = 349: GoSub 9120: GoTo 700
1335  R% = AR%(G%): GoSub 9010: ST% = Val(Mid$(RC$, 9, 1))
1340  NP% = Val(Mid$(RC$, 86, 4)): If NP% <> 0 Then GoTo 3205
1350  If CC% > 6 Then R% = 340: GoSub 9120: GoTo 700
1360  CC% = CC% + 1: AL%(G%) = 0: GoTo 3100
1400  W1$ = "K":
1410  TG% = 1
1420  If GL%(TG%) <> SI% Then GoTo 1450
1430  R% = GR%(TG%): GoSub 9010: If W2$ = Mid$(RC$, 26, 4) Then GoTo 1460
1440  If W2$ = Mid$(RC$, 30, 4) Then GoTo 1460
1450  TG% = TG% + 1: If TG% > NG% Then R% = 339: GoSub 9120: GoTo 700 Else GoTo 1420
1460  CO$ = Left$(W1$, 1) & Mid$(Str$(TG% * 4 / 100) & "0", 3, 2): KG$ = RC$: GoTo 2000
1500  R% = 342: GoSub 9120:
1510  s$ = "%1NOTHING%1": For G% = 1 To NA%
1520  If AL%(G%) <> 0 Then GoTo 1540
1530  TX$ = "%1A ": GoSub 9200: R% = AR%(G%): GoSub 9010: GoSub 8000: s$ = " %1"
1540  Next G%
1550  TX$ = s$: GoSub 9230: GoTo 700
1560  R% = Val(Left$(PL$, 4)): GoSub 9120: GoTo 630
1570  GoSub 9500:
      MsgText$ = MsgText$ & "SCORE:" & Str$(Score):
      GoTo 910
1600  GoSub 9500:
      MsgText$ = "Your Score was " & Str$(Score):
      GoSub 11000:
      MsgBox (MsgText$)
      Unload Me
      End

2000  AC$ = "?": RC$ = PL$:
2010  GoSub 8505: If AC$ <> "?" Then GoTo 3000
2015  If Left$(CO$, 1) <> "C" Then GoTo 2030
2020  If Right$(CO$, 2) < "09" Then R% = 343: GoSub 9120: GoTo 700
2025  R% = 358: GoSub 9120: GoTo 700
2030  G% = 1:
2040  If GL%(G%) <> SI% Then GoTo 2060
2050  R% = GR%(G%): GoSub 8500: If AC$ <> "?" Then GoTo 3000
2060  G% = G% + 1: If G% <= NG% Then GoTo 2040
2070  G% = 1:
2080  If AL%(G%) <> SI% And AL%(G%) <> 0 Then GoTo 2100
2090  R% = AR%(G%): GoSub 8500: If AC$ <> "?" Then GoTo 3000
2100  G% = G% + 1: If G% <= NA% Then GoTo 2080
2120  If Left$(CO$, 1) = "F" Or Left$(CO$, 1) = "K" Then GoTo 2200
2130  R% = 344: GoSub 9120: GoTo 700
2150
2200  If Left$(W1$, 1) = "F" Then s$ = Mid$(KG$, 67, 8): M% = 354: GoTo 2220
2210  s$ = "00" & Mid$(KG$, 75, 6): M% = 351
2220  X% = 1:
2230  G% = Val(Mid$(s$, X%, 2)): If G% = 0 Then GoTo 2250
2240  If AL%(G%) = 0 Then GoTo 2300
2250  If X% < 6 Then X% = X% + 2: GoTo 2230
2260  GoSub 9110: TX$ = L2$: GoSub 9200
2270  If Left$(W1$, 1) = "F" Then R% = 341: GoSub 9120
2280  GoTo 700
2300  TX$ = "The " & L2$: GoSub 9200: R% = M% + 1: GoSub 9120
2310  R% = AR%(G%): GoSub 9010: GoSub 8000: If Left$(W1$, 1) = "F" Then GoTo 2400
2320  GL%(TG%) = 9999: GoTo 700
2400  AL%(G%) = 9999: CC% = CC% - 1: GF%(TG%) = Val(Mid$(KG$, 86, 1)): GoTo 700
3000  X% = Val(Left$(AC$, 1)) + 1:
3010  On X% GoTo 700, 3050, 3200, 5000, 3850, 3870, 4000, 4100, 4200, 5000
3020  GoTo 9999
3050  MM$ = Right$(AC$, 4): GoSub 9100:
3100  AC$ = A2$: A2$ = "0": GoTo 3000:
3200  NP% = Val(Right$(AC$, 4)):
3205  R% = NP%: GoSub 9010
3210  If Right$(RC$, 1) < "6" Then GoTo 500
3220  If Right$(RC$, 1) > "6" Then GoTo 3800
3250  X% = Val(Mid$(RC$, 19, 1)): On X% GoTo 3270, 3270, 3280, 3600, 3700
3270  If X% <> PS% Then GoTo 3600
3280  LV% = 0:
3290  N$ = Mid$(RC$, LV% + 20, 2): GoSub 8600: If NN% < 0 Then GoTo 3320
3300  If NN% = 0 Then GoTo 3330
3310  If AL%(NN%) <> 0 Then GoTo 3600 Else GoTo 3330
3320  If GL%(-NN%) <> SI% Then GoTo 3600
3330  N$ = Mid$(RC$, LV% + 32, 2): GoSub 8600: If NN% < 0 Then GoTo 3360
3340  If NN% = 0 Then GoTo 3370
3350  If AL%(NN%) = 0 Then GoTo 3600 Else GoTo 3370
3360  If GL%(NN%) = SI% Then GoTo 3600
3370  If LV% < 10 Then LV% = LV% + 2: GoTo 3290
3380  CR$ = RC$: R% = Val(Mid$(RC$, 44, 4)): If R% = 0 Then GoTo 3410
3390  GoSub 9010: STT$ = Mid$(RC$, 9, 1): RC$ = CR$
3400  If Mid$(RC$, 48, 1) <> STT$ Then GoTo 3600
3410  If Val(Mid$(RC$, 49, 1)) > Rnd(RN%) * 10 Then GoTo 3600
3500  R% = Val(Mid$(RC$, 11, 4)):
3510  AC$ = Left$(RC$, 5): GoSub 9120: GoTo 3000
3600  R% = Val(Mid$(RC$, 15, 4)):
3610  AC$ = Mid$(RC$, 6, 5): GoSub 9120: GoTo 3000
3700  A2$ = Mid$(RC$, 6, 5): GoTo 3500:
3800  CR$ = RC$: For l026C% = 1 To 86 Step 5
3810  R% = Val(Mid$(CR$, l026C%, 4)): If R% = 0 Then GoTo 3830
3820  GoSub 9010: Mid$(RC$, 9, 1) = Mid$(CR$, l026C% + 4, 1): GoSub 9600
3830  Next l026C%: AC$ = Mid$(CR$, 91, 5): GoTo 3000
3850  PS% = 1: Mid$(PL$, 9, 1) = "1": RC$ = PL$: R% = SI%
3860  GoSub 9600: GoTo 3200
3870  N$ = Right$(AC$, 4): GoSub 8600: PN% = PN% + NN%: GoTo 3100
4000  MM$ = Right$(AC$, 4): GoSub 9100:
4010  R% = 345: GoSub 9120: GoSub 9850
4015  MsgText$ = MsgText$ & ":"
      GoSub 7000: s$ = CO$
4020  s$ = Chr$(Asc(Left$(s$, 1)) And &H5F): If s$ = "N" Then GoTo 1600
4030  M% = M% + 10: For G% = 1 To NA%
4040  If AL%(G%) = 0 Then GoSub 8400
4050  Next G%: CC% = 0: SI% = HO%: ARR$ = "Y": GoTo 600
4100  CO$ = Mid$(AC$, 2, 3): GoTo 2000
4200  R% = Val(Right$(AC$, 4)): GoSub 9010: X% = Val(Mid$(RC$, 9, 1))
4220  Mid$(RC$, 9, 1) = Mid$("21", X%, 1): GoSub 9600: GoTo 3100
5000  TY$ = Mid$(AC$, 3, 1): FL% = Val(Mid$(AC$, 2, 1)):
5010  If Right$(AC$, 2) = "??" Then Mid$(AC$, 4, 2) = CA$
5020  If TY$ = "0" Then R% = SI%: GoTo 6000
5030  G% = Val(Right$(AC$, 2)): If G% > 0 Then GoTo 5160
5040  If FL% = 3 Then GoTo 5200
5050  If TY$ <> "1" Then GoTo 5120
5060  For X% = 1 To NA%
5070  If AL%(X%) = 0 Or AL%(X%) = SI% Then GoSub 8700
5080  Next X%
5090  If G% > 0 Then R% = AR%(G%): GoTo 6000
5100  R% = 344: GoSub 9120: GoTo 3100
5120  For X% = 1 To NG%
5130  If GL%(X%) = SI% Then GoSub 8700
5140  Next X%
5150  If G% > 0 Then R% = GR%(G%): GoTo 6000 Else GoTo 5100
5160  If TY$ = "1" Then R% = AR%(G%) Else R% = GR%(G%)
5170  GoTo 6000
5200  If TY$ <> "1" Then GoTo 5250
5210  G% = Int(Rnd(RN%) * NA%) + 1
5220  If AL%(G%) = 0 Or AL%(G%) = SI% Then GoTo 5100
5230  R% = AR%(G%)
5240  GoSub 9010: If Mid$(RC$, 99, 1) = "1" Then GoTo 5100 Else GoTo 6100
5250  G% = Int(Rnd(RN%) * NG%) + 1
5260  If GL%(G%) = SI% Then GoTo 5100 Else R% = GR%(G%): GoTo 5240
6000  GoSub 9010:
6100
6110  On FL% + 1 GoTo 6200, 6300, 6370, 6400, 1360, 1350
6200  X% = Val(Mid$(RC$, 9, 1)): Mid$(RC$, 9, 1) = Mid$("21", X%, 1)
6230  GoSub 9600: If TY$ = "0" Then PS% = Val(Mid$(RC$, 9, 1)): PL$ = RC$
6240  R% = Val(Mid$(RC$, 14 + 4 * Val(Mid$(RC$, 9, 1)), 4)): GoSub 9120
6250  GoTo 3100
6300  NP% = Val(Mid$(RC$, 93, 4)): If NP% <> SI% Then GoTo 6330
6310  NP% = WR%: If NP% = SI% Then NP% = 9999
6330  If TY$ <> "1" Then GoTo 6380
6340  If AL%(G%) = 0 Then CC% = CC% - 1
6350  AL%(G%) = NP%
6360  R% = 346: GoSub 9900: GoTo 3100
6370  NP% = 9999: GoTo 6330
6380  GL%(G%) = NP%: If FL% = 1 Then GoTo 6360
6390  R% = 347: GoSub 9900: GoTo 3100
6400  If Left$(AC$, 1) = "9" Then GoTo 6430
6410  TX$ = "A ": GoSub 9220: GoSub 9750: GoSub 9110
6420  R% = 348: GoSub 9120
6430  If TY$ <> "1" Then GoTo 6460
6440  If AL%(G%) = 0 Then CC% = CC% - 1
6450  GoSub 8400: GoTo 3100
6460  GL%(G%) = SI%: GoTo 3100
7000

7020  If MsgText$ <> "" Then Text1.Text = MsgText$
      GoSub 9500
      scorebox.Text = "Score: " & Score

      ' Loop waiting for command input
7021  InputLine.Visible = True ' nquesti.Show
      DoEvents
      
      On Error Resume Next
      InputLine.SetFocus ' Set the inputline box as the focus so the cursor moves here for input
      On Error GoTo 0
      
      ' Do While InputLine.Visible = True ' nquesti.Visible
      Do While gv0090 = ""
        DoEvents
      Loop

      CO$ = gv0090
      gv0090 = ""
      MsgText$ = ""
7025  If CO$ = "" Then
      GoTo 7021
      End If
7040  Return
8000  ST% = Val(Mid$(RC$, 9, 1))
8010  R% = Val(Mid$(RC$, 6 + 4 * ST%, 4)): GoTo 9120
8100  GoSub 9010: NM% = Val(Mid$(RC$, 6 + 4 * Val(Mid$(RC$, 9, 1)), 4))
8110  R% = Val(Mid$(RC$, 5, 4)): If R% = 0 Then GoTo 8130
8120  If R% = OM% Then TX$ = "%4and a": GoSub 9230 Else GoSub 9120
8130  OM% = R%: R% = NM%: GoTo 9120
8200  If Left$(CO$, 1) = " " Then CO$ = Right$(CO$, Len(CO$) - 1): GoTo 8200
8210  X% = InStr(CO$, " "): If X% = 0 Then X% = Len(CO$) + 1
8220  LW$ = Left$(CO$, X% - 1): CO$ = Right$(CO$, Len(CO$) - X% + 1)
8230  Return
8300  GoSub 9010: For X% = 1 To 91 Step 13
8305  If Mid$(RC$, X%, 4) <> "?   " And Mid$(RC$, X%, 4) <> W1$ Then GoTo 8330
8310  If Mid$(RC$, X% + 4, 4) <> "?   " And Mid$(RC$, X% + 4, 4) <> W2$ Then GoTo 8330
8320  AC$ = Mid$(RC$, X% + 8, 5): X% = 91
8330  Next X%
8340  R% = Val(Mid$(RC$, 92, 4)): If R% = 0 Or AC$ <> "?" Then Return Else GoTo 8300
8400  If Mid$(PL$, 90, 3) = "***" Then AL%(G%) = Val(Mid$(PL$, 94, 4)) Else AL%(G%) = SI%
8405 GoSub 12400:
8410  Return
8500  GoSub 9010:
8505  X% = 34:
8510  If Mid$(RC$, X%, 1) = "E" Then Return
8520  If Mid$(RC$, X%, 3) = CO$ Then GoTo 8590
8530  If Mid$(RC$, X% + 1, 2) = "??" And Mid$(RC$, X%, 1) = Left$(CO$, 1) Then GoTo 8580
8540  If X% < 90 Then X% = X% + 8: GoTo 8510 Else Return
8580  CA$ = Right$(CO$, 2)
8590  AC$ = Mid$(RC$, X% + 3, 5): Return
8600  NN% = Val(N$): If Right$(N$, 1) < "P" Then Return:
8610  NN% = 0 - 10 * NN% - (Asc(Right$(N$, 1)) And &HF): Return
8700  If G% = 0 Then GoTo 8720
8710  If Rnd(RN%) * 10 > 2 Then Return
8720  G% = X%: Return
9000  Let R% = Val(RR$):
9005  Stop
9010  GoSub 9050: RC$ = m00D0.M0149:
9020  Return

9050  m00D0.M0149 = Space$(100)
      If R% > 2499 Then
         Get #2, R% - 2499, m00D0
         Get #2, R% - 2499, m019C
      Else
         Get #1, R% + 1, m00D0
         Get #1, R% + 1, m019C
      End If
      
      recType$ = Right$(m00D0.M0149, 1)
      
      Select Case recType$
            Case "0"
                RecTypeTxt$ = "Unused"
            Case "1"
                RecTypeTxt$ = "Control"
            Case "2"
                RecTypeTxt$ = "Text"
            Case "3"
                RecTypeTxt$ = "Place"
            Case "4"
                RecTypeTxt$ = "Artefact"
            Case "5"
                RecTypeTxt$ = "Creature"
            Case "6"
                RecTypeTxt$ = "Actions"
            Case "7"
                RecTypeTxt$ = "Keys"
            Case "8"
                RecTypeTxt$ = "State"
            Case Else
                RecTypeTxt$ = "Invalid Rec"
      End Select
      
      DebugData.DebugAdd (R% + 1), recType$, RecTypeTxt$, m00D0.M0149

        
9055  Return

9100  If Left$(MM$, 2) = "??" Then GoSub 9700 Else M% = Val(MM$)
9110  R% = M%:
9120  If R% = 0 Then GoTo 9190 Else GoSub 9050: TX$ = m019C.M0161: GoSub 9200
9130  If m019C.M015B = "1" Then GoTo 9190 Else TX$ = m019C.M0167: GoSub 9200
9140  If m019C.M015B = "3" Then R% = R% + 1: GoTo 9120
9190  Return

9200  If Right$(TX$, 1) <> " " Then GoTo 9225
9210  If Right$(TX$, 2) = ". " Then GoTo 9230 Else TX$ = Left$(TX$, Len(TX$) - 1)
9220  If Len(TX$) = 0 Then GoTo 9190 Else GoTo 9200
9225  If Right$(TX$, 1) = "." Then TX$ = TX$ & " "
9230  Let X% = InStr(TX$, "%")
9240  If X% = 0 Then X% = Len(TX$) + 1

9250  OB$ = OB$ & Left$(TX$, X% - 1) & " "
9260  If Len(OB$) > 70 Then GoSub 9300
9270  If X% > Len(TX$) Then GoTo 9190
9280  On Val(Mid$(TX$, X% + 1, 1)) GoSub 9400, 9410, 9420, 9430
9290  TX$ = Right$(TX$, Len(TX$) - X% - 1): If Len(TX$) = 0 Then GoTo 9190 Else GoTo 9230
9300  Y% = InStr(60, OB$, " "): If Y% = 0 Then Y% = Len(OB$)
9310  MsgText$ = MsgText$ & Left$(OB$, Y%)
9320  OB$ = Right$(OB$, Len(OB$) - Y%)
9325  If Left$(OB$, 1) = " " Then OB$ = Right$(OB$, Len(OB$) - 1): GoTo 9325
9330  Return

9400  MsgText$ = MsgText$ & OB$ & Chr(13) & Chr(10): OB$ = "": Return
9410  OB$ = OB$ & "%": Return
9420  OB$ = Left$(OB$, Len(OB$) - 1): Return

9430  If Len(OB$) < 4 Then Return Else OB$ = Left$(OB$, Len(OB$) - 4) & " ": Return

9500
9505 M% = M% - 1
9540 Score = -PN%
     Return

9600 m0136.M0149 = RC$
     If R% < 2500 Then
        Put #1, R% + 1, m0136
     Else
        Put #2, R% - 2499, m0136
     End If
     Return

9700  If Mid$(MM$, 4, 1) = "0" Then GoTo 9720
9710  R% = GR%(Val(CA$)): GoTo 9730
9720  R% = AR%(Val(CA$))
9730  GoSub 9010: GoSub 9750: TX$ = "The ": GoSub 9200: Return
9750  M% = Val(Mid$(RC$, 6 + 4 * Val(Mid$(RC$, 9, 1)), 4)): Return
9850  If Len(OB$) > 0 Then GoSub 9300 Else Return
9860  If Len(OB$) > 0 Then
      MsgText$ = MsgText$ & OB$:
      OB$ = ""
      End If
9870  Return
9900  If Left$(AC$, 1) = "9" Then Return Else SV% = R%
9905  TX$ = "The ": GoSub 9200: GoSub 9750: GoSub 9110
9910  R% = SV%: GoTo 9120
9999  MsgBox ("NOT SUPPORTED"): Exit Sub

10000
10005
10010 R% = 0
10020 GoSub 9050
10030 l028E = Val(Mid$(m00D0.M0149, 70, 4))
10040 l0292$ = Mid$(m00D0.M0149, 40, 30)
      l0294$ = ""
10045 If l0296$ = "Y" Then l0294$ = "NOW "

10050 l0294$ = l0294$ & "HIGH SCORE IS " & l028E & Chr(10)
10056 If l0296$ = "Y" Then l0294$ = l0294$ & "RECENTLY "
10060 l0294$ = l0294$ & "SET BY " & l0292$
      MsgBox (l0294$)
10070 Return

11000
11005 l0296$ = "N"
11010 R% = 0
11020 GoSub 9050
11030 l028E = Val(Mid$(m00D0.M0149, 70, 4))
11040 If l028E > Score Then GoTo 11095
11045 l0296$ = "Y"
      l0298$ = "QUEST"
11060 l0298$ = InputBox$("WHAT IS YOUR NAME", "CONGRATULATIONS! NEW HIGH SCORE")
11070 Mid$(m00D0.M0149, 40, 30) = l0298$ & Space$(30 - Len(l0298$))
11080 Mid$(m00D0.M0149, 70, 4) = Right$(Str$(10000 + Int(Score)), 4)
11090 GoSub 11500:
11095 GoSub 12100:
11100 GoSub 10000
11110 Close 1: Close 2
11120 Return

11500
11510 If R% > 2499 Then Put 2, R% - 2499, m00D0.M0149 Else Put 1, R% + 1, m00D0.M0149
11520 Return

12000
      l029A = 3 + 32:
      l029E = MsgBox("Is this a new game", l029A, "Welcome to Quest")
      If l029E = 6 Then NewGame = "Y": Return
      If l029E = 7 Then NewGame = "N": Return
      Unload Me
      End

12100
12105 GoSub 12300
12106 If Left$(s$, 1) <> "Y" And Left$(s$, 1) <> "y" Then Return
12110 LSet m00D0.M0149 = Space$(100): l02A2 = 1
12120 For xxxx = 1 To NA%
12130 If AL%(xxxx) = 0 Then Mid$(m00D0.M0149, l02A2, 4) = Right$(Str$(10000 + xxxx), 4): l02A2 = l02A2 + 4
12140 Next xxxx
12150 R% = 2229
12160 GoSub 11500
12165 GoSub 15000:
12170 Return
12200
12210 R% = 2229: GoSub 9050
12220 l02A2 = 1
12230 If Mid$(m00D0.M0149, l02A2, 4) = "    " Then Return
12240 AL%(Val(Mid$(m00D0.M0149, l02A2, 4))) = 0
12250 l02A2 = l02A2 + 4
12260 GoTo 12230

12300
      l029A = 4 + 32:
      l02AA = MsgBox("Save Game?", l029A, "Thank you for playing Quest")
      If l02AA <> 6 Then Return

12340 Mid$(m00D0.M0149, 74, 4) = Right$(Str$(10000 + M%), 4)
12350 Mid$(m00D0.M0149, 78, 4) = Right$(Str$(10000 + SI%), 4)
12355 Mid$(m00D0.M0149, 82, 4) = Right$(Str$(10000 + Int(Score)), 4)

12360 R% = 0: GoSub 11500:
      GoSub 15000
      Close #3
      Close #1
      Close #2
      Unload Me
      End

12400
12410 If AL%(G%) <> HO% Then Return:
12420 R% = AR%(G%): GoSub 9010
12430 ST% = Val(Mid$(RC$, 9, 1))
12440 If Mid$(RC$, 80 + ST% + ST%, 1) <> "2" Then Return
12441 l02AE = Val(Mid$(RC$, 79, 3)): If l02AE = 0 Then Return
12443 PN% = PN% - l02AE
12445 AL%(G%) = WR%:
12450 GoSub 12500:
12460 Return
12500
12540 Return

15000
    N$ = "QUEST"
    N$ = InputBox("Name for Saved Game", "Quest")
    If N$ = "" Then Return
    Open DirPath$ & N$ & ".SAV" For Random As #3
    For xxxx = 1 To 60
        LSet m00B6.M0114 = Right$(Str$(10000 + AL%(xxxx)), 4)
        LSet m00B6.M011B = Right$(Str$(10000 + AR%(xxxx)), 4)
        LSet m00B6.M0122 = Right$(Str$(10000 + GL%(xxxx)), 4)
        LSet m00B6.M0129 = Right$(Str$(10000 + GR%(xxxx)), 4)
        LSet m00B6.M0130 = Right$(Str$(10000 + GF%(xxxx)), 4)
        LSet m00B6.M0137 = Left$(N$ & "    ", 4)
        Put #3, xxxx, m00B6
    Next xxxx
    Close #3
    Return

16000
      ' nquestf.Show 1
      nquestf.Show
      N$ = gv0094
      If N$ = "" Then GoTo 16150

16010 Open DirPath$ & N$ For Random As #3
16020
16025 Get #3, 1, m00B6
16026 If Left$(N$ & "    ", 4) <> m00B6.M0137 Then GoTo 16130
16030 For xxxx = 1 To 60
16035 Get #3, xxxx, m00B6
16040 AL%(xxxx) = Val(m00B6.M0114)
16050 AR%(xxxx) = Val(m00B6.M011B)
16060 GL%(xxxx) = Val(m00B6.M0122)
16070 GR%(xxxx) = Val(m00B6.M0129)
16080 GF%(xxxx) = Val(m00B6.M0130)
16100 Next xxxx
16110 Close #3
16120 Return
16130 Close #3
16140 Kill DirPath$ & N$
16150 ValidFile$ = "NOWAY"
16160 Return
End Sub

Private Sub Command2_Click()

    About.Show
    
End Sub

Private Sub Form_Load()

    Dim hMenu As Long
    hMenu = GetSystemMenu(Me.hwnd, False)
    Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)

    InputLine.Visible = False ' make the input line invisible
    
    ShowDebug.Value = IIf(DebugData.Visible, 1, 0)

    IsLoading = True ' used to avoid the combo_click code from trying to change the text box alignment before the form has completed loading

    ' Set the default text alignment options in the combo box at form load time.
    Combo1.AddItem "Left"
    Combo1.AddItem "Center"
    Combo1.AddItem "Right"
    Combo1.ListIndex = 1
    
    IsLoading = False
    
    ' set the default text alignment in the text box to centre
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SelAlignment = rtfCenter

End Sub


Private Sub Combo1_Click()

    ' if the form is still loading then exit the sub
    If IsLoading Then Exit Sub
    
    ' set the main text box alignment depending on the option selected in the combo box
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    
    Select Case Combo1.Text
        Case "Left"
            Text1.SelAlignment = rtfLeft
        Case "Center"
            Text1.SelAlignment = rtfCenter
        Case "Right"
            Text1.SelAlignment = rtfRight
    End Select

End Sub


Private Sub InputLine_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
        
        If Trim$(InputLine.Text) = "" Then
            ' Blank input — ignore it and keep the textbox visible
            Exit Sub
        End If
        
        
        gv0090 = InputLine.Text
        
        InputLine.Text = ""
        
        ' Unload Me
        FocusSink.SetFocus
        
        InputLine.Visible = False
    End If

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


