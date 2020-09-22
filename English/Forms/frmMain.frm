VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "DATE-CALCULATOR - (C)2003 by CodeXP"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdAbout 
      Caption         =   "AB&OUT"
      Height          =   285
      Left            =   2280
      TabIndex        =   24
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit &Program"
      Height          =   285
      Left            =   3600
      TabIndex        =   25
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Frame fraRechopt 
      Caption         =   " Operations: "
      Height          =   3015
      Left            =   2280
      TabIndex        =   13
      Top             =   0
      Width           =   3135
      Begin VB.CheckBox chkSetDate 
         Caption         =   "Use Result as Date"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   1815
      End
      Begin VB.ComboBox lstInt 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optRO 
         Caption         =   "Diff"
         Height          =   285
         Index           =   2
         Left            =   1080
         Style           =   1  'Grafisch
         TabIndex        =   21
         ToolTipText     =   " Difference (Date - Date = Interval) "
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optRO 
         Caption         =   "-"
         Height          =   285
         Index           =   1
         Left            =   600
         Style           =   1  'Grafisch
         TabIndex        =   20
         ToolTipText     =   " Subtraction (Date - Interval = Date) "
         Top             =   1680
         Width           =   375
      End
      Begin VB.OptionButton optRO 
         Caption         =   "+"
         Height          =   285
         Index           =   0
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   19
         ToolTipText     =   " Addition (Date + Interval = Date) "
         Top             =   1680
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   15
         Text            =   "0"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operation:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame fraErgebnis 
      Caption         =   " Result: "
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
      Begin VB.CommandButton cmdSetDatum 
         Caption         =   "&Use as Date"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdCopyErg 
         Caption         =   "&Copy"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   " Copy to Clipboard "
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtErgWTN 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtErg 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weekday:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame fraDatum 
      Caption         =   " Date: "
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdSetNow 
         Caption         =   "Current &Date"
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtTDJ 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MousePointer    =   1  'Pfeil
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtMonat 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Pfeil
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtWTN 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Pfeil
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day of Year:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   510
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weekday:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
  Dim sTxt As String
  sTxt = "DATE CALCULATOR (C)2003 by CodeXP" & vbCrLf & _
         "Coder: CodeXP" & vbCrLf & _
         "E-Mail: CodeXP@Lycos.de" & vbCrLf & _
         "Language: English" & vbCrLf & _
         "Coded in: VB6" & vbCrLf & _
         "Date: 28.09.2003" & vbCrLf & vbCrLf & _
         "Thanx to: VB-Runtime" & vbCrLf & _
         "Greetz to:" & vbCrLf & _
         "  ooOimaxOoo [VB-Coder]," & vbCrLf & _
         "  G-Man.Net [damn C++/.Net-Coder] ^^," & vbCrLf & _
         "  Unknown [great VB-Coder]"
  MsgBox sTxt, vbInformation
End Sub

Private Sub cmdCalc_Click()
  Dim sInt As String
  Dim sErg As String
  Dim iOpt As Integer
  
  Select Case lstInt.ListIndex
  Case 0 ' yyyy - Jahr           '
    sInt = "yyyy"
  Case 1 ' q    - Quartal        '
    sInt = "q"
  Case 2 ' m    - Monat          '
    sInt = "m"
  Case 3 ' y    - Tag des Jahres '
    sInt = "y"
  Case 4 ' d    - Tag            '
    sInt = "d"
  Case 5 ' w    - Wochentag      '
    sInt = "w"
  Case 6 ' ww   - Woche          '
    sInt = "ww"
  Case 7 ' h    - Stunde         '
    sInt = "h"
  Case 8 ' n    - Minute         '
    sInt = "n"
  Case 9 ' s    - Sekunde        '
    sInt = "s"
  Case Else
    MsgBox "Please Choose the Interval first!", vbExclamation
    Exit Sub
  End Select
  
  iOpt = IIf(optRO(0).Value, 1, _
         IIf(optRO(1).Value, 2, _
         IIf(optRO(2).Value, 3, 0)))
  
  On Error GoTo Calc_Error
  Select Case iOpt
  Case 1  ' Add '
    txtValue = Replace(Val(Trim(txtValue)), ",", ".")
    sErg = DateAdd(sInt, Val(txtValue), CDate(Trim(txtDate)))
  Case 2  ' Sub '
    txtValue = Replace(Val(Trim(txtValue)), ",", ".")
    sErg = DateAdd(sInt, -Val(txtValue), CDate(Trim(txtDate)))
  Case 3  ' Dif '
    sErg = DateDiff(sInt, CDate(Trim(txtValue)), CDate(Trim(txtDate)))
  Case Else
    MsgBox "Wählen Sie zuerst eine Rechenoperation aus!", vbExclamation
    Exit Sub
  End Select
  
  txtErg = sErg
  If chkSetDate.Value = vbChecked Then txtDate = sErg
  
  Exit Sub
Calc_Error:
  txtErg = ""
  MsgBox "Error! Check your Inputs!", vbCritical
End Sub

Private Sub cmdCopyErg_Click()
  Clipboard.Clear
  Clipboard.SetText txtErg
End Sub

Private Sub cmdQuit_Click()
  Unload Me
End Sub

Private Sub cmdSetDatum_Click()
  If Len(txtErg) > 0 Then txtDate = txtErg
End Sub

Private Sub cmdSetNow_Click()
  txtDate = Now
End Sub

Private Sub Form_Load()
  txtDate = Now

  ' Intervale füllen              '
  lstInt.AddItem "Year"
  lstInt.AddItem "Quarter"
  lstInt.AddItem "Month"
  lstInt.AddItem "Day of Year"
  lstInt.AddItem "Day"
  lstInt.AddItem "Weekday"
  lstInt.AddItem "Week"
  lstInt.AddItem "Hour"
  lstInt.AddItem "Minute"
  lstInt.AddItem "Second"
  lstInt.ListIndex = 0
End Sub

Private Sub txtDate_ValidationCheck()
  Dim oTxt As String
  Dim nTxt As String
  Dim nDat As Date
  nTxt = Trim(txtDate)
  nTxt = Replace(nTxt, "/", ".")
  nTxt = Replace(nTxt, "\", ".")
  nTxt = Replace(nTxt, "-", ".")
  Do
    oTxt = nTxt
    nTxt = Replace(Trim(oTxt), "  ", " ")
  Loop While nTxt <> oTxt
  On Error Resume Next
  nDat = CDate(nTxt)
  If Err Then
    txtDate = nTxt
    MsgBox "Date you have entered is invalid!", vbCritical
  Else
    txtDate = nDat
  End If
End Sub

Private Sub optRO_Click(Index As Integer)
  On Error Resume Next
  If Index = 2 Then
    lblCap(2) = "Date:"
  Else
    lblCap(2) = "Value:"
  End If
  cmdCalc.SetFocus
End Sub

Private Sub txtDate_Change()
  FillInfo
End Sub

Private Sub txtDate_GotFocus()
  txtDate.SelStart = 0
  txtDate.SelLength = Len(txtDate)
End Sub

Private Sub txtDate_LostFocus()
  txtDate_ValidationCheck
End Sub

Private Sub txtErg_Change()
  On Error Resume Next
  txtErgWTN = WeekdayName(Weekday(CDate(txtErg), vbMonday))
  If Err Then txtErgWTN = ""
End Sub

Private Sub txtMonat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  If txtMonat.SelLength = 0 Then
    txtDate.SetFocus
  End If
End Sub

Private Sub txtTDJ_Change()
  On Error Resume Next
  If txtTDJ.SelLength = 0 Then
    txtDate.SetFocus
  End If
End Sub

Private Sub txtWTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  If txtWTN.SelLength = 0 Then
    txtDate.SetFocus
  End If
End Sub

Private Sub FillInfo()
  Dim nDat As Date
  On Error GoTo FillInfo_Error
  nDat = CDate(Trim(txtDate))
  txtWTN = WeekdayName(Weekday(nDat, vbMonday))
  txtMonat = MonthName(Month(nDat))
  txtTDJ = DatePart("y", nDat, vbMonday)
  Exit Sub
FillInfo_Error:
  txtTDJ = ""
  txtWTN = ""
  txtMonat = ""
End Sub

