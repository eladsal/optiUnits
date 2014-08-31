VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "optiUnits by Elad Salomons"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   5400
      TabIndex        =   14
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox cmbOut 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3840
      Width           =   4095
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdInpFile 
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3960
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "EPANet unit convertor tool"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   555
      Left            =   2040
      TabIndex        =   16
      Top             =   960
      Width           =   6075
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Elad Salomons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7080
      MouseIcon       =   "frmMain.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4560
      Width           =   2550
   End
   Begin VB.Label Label5 
      Caption         =   "Convert to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label lblInpUnits 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblInpUnits"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "INP units:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OptiWater"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   555
      Left            =   5040
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   0
      Width           =   2325
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "More products at: www.optiwater.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4560
      MouseIcon       =   "frmMain.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   600
      Width           =   3225
   End
   Begin VB.Label Label3 
      Caption         =   "INP file name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "optiUnits 1.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   3150
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   285
   End
   Begin VB.Label Label11 
      Caption         =   "Output file name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblInpFileName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblInpFileName"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   6255
   End
   Begin VB.Label lblOutputFileName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblOutputFileName"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   2400
      Width           =   6255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InpFileName As String
Dim OutputFileName As String
Dim UnitsNames() As String

Dim W(50) As String
Dim N As Long
Dim InpUnits As String
Dim InpHeadloss As String

Const EN_CFS = 0           ' Flow units types
Const EN_GPM = 1
Const EN_MGD = 2
Const EN_IMGD = 3
Const EN_AFD = 4
Const EN_LPS = 5
Const EN_LPM = 6
Const EN_MLD = 7
Const EN_CMH = 8
Const EN_CMD = 9

Dim multiDistance As Single
Dim multiDiameter As Single
Dim multiFlow As Single
Dim multiVolume As Single
Dim multiHeadloss As Single
Dim multiPower As Single

Sub DoConversion()
'---------------------------------------------------------------------------
Dim Fin As Long, Fout As Long
Dim Section As String
Dim L As String
Dim NewSection As Boolean
Dim CurveType As String
Dim colLinks As New Collection
Dim EpanetComment As String

Fin = FreeFile
Open InpFileName For Input As #Fin
Fout = FreeFile
Open OutputFileName For Output As #Fout
Do
    EpanetComment = ""
    Line Input #Fin, L
    StringToWords L, W(), N, EpanetComment
    If N > 0 Then
        NewSection = False
        If CStr(Left(W(1), 1)) = "[" Then
            Section = UCase(W(1)): NewSection = True
        End If
        
        If Not NewSection Then
            Select Case Section
                Case "[TITLE]"
                    Print #Fout, L
                Case "[JUNCTIONS]"
                    Print #Fout, W(1), CSng(W(2)) * multiDistance;
                    If N > 2 Then
                        Print #Fout, CSng(W(3)) * multiFlow;
                        If N > 3 Then
                            Print #Fout, W(4), EpanetComment
                        Else
                            Print #Fout, EpanetComment
                        End If
                    Else
                        Print #Fout, EpanetComment
                    End If
                Case "[RESERVOIRS]"
                    Print #Fout, W(1), CSng(W(2)) * multiDistance;
                    If N > 2 Then
                        'changed 29-06-2012
                        'Print #Fout, CSng(W(3)), EpanetComment
                        Print #Fout, CStr(W(3)), EpanetComment
                    Else
                        Print #Fout, EpanetComment
                    End If
                Case "[TANKS]"
                    Print #Fout, W(1), CSng(W(2)) * multiDistance, , CSng(W(3)) * multiDistance, CSng(W(4)) * multiDistance, CSng(W(5)) * multiDistance, CSng(W(6)) * multiDistance, CSng(W(7)) * multiVolume;
                    If N > 7 Then
                        Print #Fout, W(8), EpanetComment
                    Else
                        Print #Fout, EpanetComment
                    End If
                Case "[PIPES]"
                    colLinks.Add "PIPE", W(1)
                    Print #Fout, W(1), W(2), W(3), CSng(W(4)) * multiDistance, CSng(W(5)) * multiDiameter, CSng(W(6)) * multiHeadloss;
                    If N > 6 Then
                        Print #Fout, CSng(W(7));
                        If N > 7 Then
                            Print #Fout, W(8), EpanetComment
                        Else
                            Print #Fout, EpanetComment
                        End If
                    Else
                        Print #Fout, EpanetComment
                    End If
                Case "[PUMPS]"
                    colLinks.Add "PUMP", W(1)
                    Print #Fout, W(1), W(2), W(3);
                    k = 4
                    For i = k To N Step 2
                        Select Case UCase(W(k))
                            Case "HEAD", "SPEED"
                                Print #Fout, " " & W(k), W(k + 1);
                            Case "POWER"
                                Print #Fout, " " & W(k), W(k + 1) * multiPower;
                        End Select
                    Next i
                    Print #Fout, EpanetComment
                Case "[VALVES]"
                    colLinks.Add W(5), W(1)
                    Print #Fout, W(1), W(2), W(3), CSng(W(4)) * multiDiameter, W(5);
                    Select Case UCase(W(5))
                        Case "PRV", "PSV", "PBV"
                            Print #Fout, CSng(W(6)) * multiDistance, W(7);
                        Case "FCV"
                            Print #Fout, CSng(W(6)) * multiFlow, W(7);
                        Case "TCV", "GPV"
                            Print #Fout, " " & W(6), W(7);
                    End Select
                    Print #Fout, EpanetComment
                Case "[EMITTERS]"
                    Print #Fout, W(1), CSng(W(2)) * multiFlow, EpanetComment
                Case "[CURVES]"
                    Print #Fout, W(1);
                    Select Case CurveType
                        Case "PUMP"
                            Print #Fout, CSng(W(2)) * multiFlow, CSng(W(3)) * multiDistance;
                        Case "EFFICIENCY"
                            Print #Fout, CSng(W(2)) * multiFlow, W(3);
                        Case "VOLUME"
                            Print #Fout, CSng(W(2)) * multiVolume, CSng(W(3)) * multiDistance;
                        Case "HEADLOSS"
                            Print #Fout, CSng(W(2)) * multiDistance, CSng(W(3)) * multiFlow;
                    End Select
                    Print #Fout, EpanetComment
                Case "[PATTERNS]"
                    Print #Fout, L
                Case "[ENERGY]"
                    Print #Fout, L
                Case "[STATUS]"
                    If UCase(W(2)) = "OPEN" Or UCase(W(2)) = "CLOSED" Or UCase(W(2)) = "ACTIVE" Then
                        Print #Fout, L
                    Else
                        Select Case colLinks(W(1))
                            Case "PIPE", "PUMP"
                                Print #Fout, L
                            Case "PRV", "PSV", "PBV"
                                Print #Fout, W(1), CSng(W(2)) * multiDistance, EpanetComment
                            Case "FCV"
                                Print Fout, W(1), CSng(W(2)) * multiFlow, EpanetComment
                            Case "TCV", "GPV"
                                Print #Fout, L
                        End Select
                    End If
                Case "[CONTROLS]"
                    Print #Fout, W(1), W(2);
                    If UCase(W(3)) = "OPEN" Or UCase(W(3)) = "CLOSED" Or UCase(W(3)) = "ACTIVE" Then
                        Print #Fout, " " & W(3);
                    Else
                        Select Case colLinks(W(2))
                            Case "PIPE", "PUMP"
                                Print #Fout, " " & W(3);
                            Case "PRV", "PSV", "PBV"
                                Print #Fout, " " & CSng(W(3)) * multiDistance;
                            Case "FCV"
                                Print Fout, " " & CSng(W(3)) * multiFlow;
                            Case "TCV", "GPV"
                                Print #Fout, " " & W(3);
                        End Select
                    End If
                    If N = 6 Then
                        Print #Fout, " " & W(4), W(5), W(6);
                    ElseIf N = 7 Then
                        Print #Fout, " " & W(4), W(5), W(6), W(7);
                    ElseIf N = 8 Then
                        Print #Fout, " " & W(4), W(5), W(6), W(7), CSng(W(8)) * multiDistance;
                    End If
                    Print #Fout, EpanetComment
                Case "[RULES]"
                    Print #Fout, L
                Case "[DEMANDS]"
                    Print #Fout, W(1), CSng(W(2)) * multiFlow;
                    If N > 2 Then
                        Print #Fout, W(3), EpanetComment
                    Else
                        Print #Fout, EpanetComment
                    End If
                Case "[QUALITY]"
                    Print #Fout, L
                Case "[REACTIONS]"
                    Print #Fout, L
                Case "[SOURCES]"
                    Print #Fout, L
                Case "[MIXING]"
                    Print #Fout, L
                Case "[OPTIONS]"
                    If UCase(W(1)) = "UNITS" Then
                        Print #Fout, W(1), UnitsNames(cmbOut.ListIndex), EpanetComment
                    Else
                        Print #Fout, L
                    End If
                Case "[TIMES]"
                    Print #Fout, L
                Case "[REPORT]"
                    Print #Fout, L
                Case "[COORDINATES]"
                    Print #Fout, L
                Case "[LABELS]"
                    Print #Fout, L
                Case "[BACKDROP]"
                    Print #Fout, L
                Case "[VERTICES]"
                    Print #Fout, L
                Case "[TAGS]" 'added 3-aug-2007
                    Print #Fout, L
                Case Else
                    Stop
            End Select
        Else
            Print #Fout, L
        End If
    Else
        Print #Fout, L
        If Section = "[CURVES]" Then
            Select Case UCase(Mid(L, 2, 4))
                Case "PUMP"
                    CurveType = "PUMP"
                Case "EFFI"
                    CurveType = "EFFICIENCY"
                Case "VOLU"
                    CurveType = "VOLUME"
                Case "HEAD"
                    CurveType = "HEADLOSS"
            End Select
        End If
    End If
Loop Until EOF(Fin)

Close #Fin, #Fout

End Sub

Sub GetInpUnits()
'---------------------------------------------------------------------------
Dim F As Integer
Dim Section As String
Dim L As String
Dim NewSection As Boolean

F = FreeFile
Open InpFileName For Input As #F
Do
    Line Input #F, L
    StringToWords L, W(), N
    If N > 0 Then
        NewSection = False
        If CStr(Left(W(1), 1)) = "[" Then
            If W(1) = "[OPTIONS]" Then
                Section = W(1): NewSection = True
            Else
                Section = ""
            End If
        End If
        
        If Not NewSection Then
            If Section = "[OPTIONS]" Then
                If UCase(W(1)) = "UNITS" Then InpUnits = UCase(W(2))
                If UCase(W(1)) = "HEADLOSS" Then InpHeadloss = UCase(W(2))
            End If
        End If
    End If
Loop Until EOF(F)

Close #F

End Sub

Sub SetConversionValues()
'----------------------------------------------------
Dim ConversionString As String
Dim ToUnits As String

ToUnits = UnitsNames(cmbOut.ListIndex)
ConversionString = InpUnits & "2" & ToUnits

multiDistance = 1
multiDiameter = 1
multiFlow = 1
multiVolume = 1
multiHeadloss = 1
multiPower = 1

If (InpUnits = "CFS" Or InpUnits = "GPM" Or InpUnits = "MGD" Or InpUnits = "IMGD" Or InpUnits = "AFD") And _
    (ToUnits = "LPS" Or ToUnits = "LPM" Or ToUnits = "MLD" Or ToUnits = "CMH" Or ToUnits = "CMD") Then
        multiDistance = 0.3048
        multiDiameter = 25.4
        multiVolume = 0.02831685
        If InpHeadloss = "D-W" Then
            multiHeadloss = 0.3048
        Else
            multiHeadloss = 1
        End If
        multiPower = 0.7456999
End If
If (ToUnits = "CFS" Or ToUnits = "GPM" Or ToUnits = "MGD" Or ToUnits = "IMGD" Or ToUnits = "AFD") And _
    (InpUnits = "LPS" Or InpUnits = "LPM" Or InpUnits = "MLD" Or InpUnits = "CMH" Or InpUnits = "CMD") Then
        multiDistance = 3.28083
        multiDiameter = 0.03936996
        multiVolume = 35.31467
        If InpHeadloss = "D-W" Then
            multiHeadloss = 3.28083
        Else
            multiHeadloss = 1
        End If
        multiPower = 1.341022
End If

Select Case ConversionString
'----------CFS-------------------------------------
    Case "CFS2GPM"
        multiFlow = 448.8312
    Case "CFS2MGD"
        multiFlow = 646316.9 / 1000000#
    Case "CFS2IMGD"
        multiFlow = 538171.1 / 1000000#
    Case "CFS2AFD"
        multiFlow = 1.98347107
    Case "CFS2LPS"
        multiFlow = 28.31685
    Case "CFS2LPM"
        multiFlow = 1699.011
    Case "CFS2MLD"
        multiFlow = 2446576 / 1000000#
    Case "CFS2CMH"
        multiFlow = 101.9406
    Case "CFS2CMD"
        multiFlow = 2446.576
    
    Case "GPM2CFS"
        multiFlow = 1 / 448.8312
    Case "MGD2CFS"
        multiFlow = 1 / (646316.9 / 1000000#)
    Case "IMGD2CFS"
        multiFlow = 1 / (538171.1 / 1000000#)
    Case "AFD2CFS"
        multiFlow = 1 / 1.98347107
    Case "LPS2CFS"
        multiFlow = 1 / 28.31685
    Case "LPM2CFS"
        multiFlow = 1 / 1699.011
    Case "MLD2CFS"
        multiFlow = 1 / (2446576 / 1000000#)
    Case "CMH2CFS"
        multiFlow = 1 / 101.9406
    Case "CMD2CFS"
        multiFlow = 1 / 2446.576
    
'----------GPM-------------------------------------
    Case "GPM2MGD"
        multiFlow = 1440 / 1000000#
    Case "GPM2IMGD"
        multiFlow = 1199.05 / 1000000#
    Case "GPM2AFD"
        multiFlow = 0.00441919194
    Case "GPM2LPS"
        multiFlow = 0.0630902
    Case "GPM2LPM"
        multiFlow = 3.785412
    Case "GPM2MLD"
        multiFlow = 5450.993 / 1000000#
    Case "GPM2CMH"
        multiFlow = 0.2271247
    Case "GPM2CMD"
        multiFlow = 5.450993

    Case "MGD2GPM"
        multiFlow = 1 / (1440 / 1000000#)
    Case "IMGD2GPM"
        multiFlow = 1 / (1199.05 / 1000000#)
    Case "AFD2GPM"
        multiFlow = 1 / 0.00441919194
    Case "LPS2GPM"
        multiFlow = 1 / 0.0630902
    Case "LPM2GPM"
        multiFlow = 1 / 3.785412
    Case "MLD2GPM"
        multiFlow = 1 / (5450.993 / 1000000#)
    Case "CMH2GPM"
        multiFlow = 1 / 0.2271247
    Case "CMD2GPM"
        multiFlow = 1 / 5.450993

'----------MGD-------------------------------------
    Case "MGD2IMGD"
        multiFlow = 0.8326738
    Case "MGD2AFD"
        multiFlow = 3.06888329
    Case "MGD2LPS"
        multiFlow = 43.81264
    Case "MGD2LPM"
        multiFlow = 2628.758
    Case "MGD2MLD"
        multiFlow = 3.785412
    Case "MGD2CMH"
        multiFlow = 157.7255
    Case "MGD2CMD"
        multiFlow = 3785.412
        
    Case "IMGD2MGD"
        multiFlow = 1 / 0.8326738
    Case "AFD2MGD"
        multiFlow = 1 / 3.06888329
    Case "LPS2MGD"
        multiFlow = 1 / 43.81264
    Case "LPM2MGD"
        multiFlow = 1 / 2628.758
    Case "MLD2MGD"
        multiFlow = 1 / 3.785412
    Case "CMH2MGD"
        multiFlow = 1 / 157.7255
    Case "CMD2MGD"
        multiFlow = 1 / 3785.412
        
'----------IMGD-------------------------------------
    Case "IMGD2AFD"
        multiFlow = 3.68557667
    Case "IMGD2LPS"
        multiFlow = 52.61681
    Case "IMGD2LPM"
        multiFlow = 3157.008
    Case "IMGD2MLD"
        multiFlow = 4.546092
    Case "IMGD2CMH"
        multiFlow = 189.4205
    Case "IMGD2CMD"
        multiFlow = 4546.092
        
    Case "AFD2IMGD"
        multiFlow = 1 / 3.68557667
    Case "LPS2IMGD"
        multiFlow = 1 / 52.61681
    Case "LPM2IMGD"
        multiFlow = 1 / 3157.008
    Case "MLD2IMGD"
        multiFlow = 1 / 4.546092
    Case "CMH2IMGD"
        multiFlow = 1 / 189.4205
    Case "CMD2IMGD"
        multiFlow = 1 / 4546.092
        
'----------AFD-------------------------------------
    Case "AFD2LPS"
        multiFlow = 14.2764102
    Case "AFD2LPM"
        multiFlow = 856.584609
    Case "AFD2MLD"
        multiFlow = 1233481.84 / 1000000#
    Case "AFD2CMH"
        multiFlow = 51.3950766
    Case "AFD2CMD"
        multiFlow = 1233.48184
        
    Case "LPS2AFD"
        multiFlow = 1 / 14.2764102
    Case "LPM2AFD"
        multiFlow = 1 / 856.584609
    Case "MLD2AFD"
        multiFlow = 1 / (1233481.84 / 1000000#)
    Case "CMH2AFD"
        multiFlow = 1 / 51.3950766
    Case "CMD2AFD"
        multiFlow = 1 / 1233.48184
        
'----------LPS-------------------------------------
    Case "LPS2LPM"
        multiFlow = 60
    Case "LPS2MLD"
        multiFlow = 86400 / 1000000#
    Case "LPS2CMH"
        multiFlow = 3.6
    Case "LPS2CMD"
        multiFlow = 86.4
        
    Case "LPM2LPS"
        multiFlow = 1 / 60
    Case "MLD2LPS"
        multiFlow = 1 / (86400 / 1000000#)
    Case "CMH2LPS"
        multiFlow = 1 / 3.6
    Case "CMD2LPS"
        multiFlow = 1 / 86.4
        
'----------LPM-------------------------------------
    Case "LPM2MLD"
        multiFlow = 1440 / 1000000#
    Case "LPM2CMH"
        multiFlow = 0.06
    Case "LPM2CMD"
        multiFlow = 1.44
        
    Case "MLD2LPM"
        multiFlow = 1 / (1440 / 1000000#)
    Case "CMH2LPM"
        multiFlow = 1 / 0.06
    Case "CMD2LPM"
        multiFlow = 1 / 1.44
        
'----------MLD-------------------------------------
    Case "MLD2CMH"
        multiFlow = 41.66667
    Case "MLD2CMD"
        multiFlow = 1000
        
    Case "CMH2MLD"
        multiFlow = 1 / 41.66667
    Case "CMD2MLD"
        multiFlow = 1 / 1000
        
'----------CMH-------------------------------------
    Case "CMH2CMD"
        multiFlow = 24
        
    Case "CMD2CMH"
        multiFlow = 1 / 24
        
End Select

End Sub

Sub StringToWords(ByVal T$, W$(), N, Optional EpanetComment As String = "")
'===========================================================================
' Extracts N% words (delimited by a space, comma or tab)
' from the string T$ and places them in the array W$.
' (Text between double quotes is treated as one word.)
'=======================================================
Dim Inword As Integer
Dim i As Long, Nchar As Long
Dim First As Long, Last As Long
Dim SepStr As String * 3
Dim c As String * 1

  SepStr = " " + "," + Chr$(9)
  T$ = RTrim$(LTrim$(T$))   'Trim both ends of T$
  i = InStr(T$, ";")        'Trim comment
  If i > 0 Then
    EpanetComment = Mid$(T$, i, Len(T$))
    T$ = Mid$(T$, 1, i - 1)
  End If
    
  N = 0                    'Initialize number of words
  Nchar = Len(T$)           'Make sure somethings left
  If Nchar = 0 Then Exit Sub
  Inword = False

  For i = 1 To Nchar        'Scan characters of T$
    'Is character I a space, comma, or tab character?
    c = Mid$(T$, i, 1)
    If InStr(SepStr, c) > 0 Then
      If Inword Then        'Found end of current word
        N = N + 1         'Increment word counter
        If N > UBound(W$) Then Exit Sub
        'Extract the last word
        W$(N) = Mid$(T$, First, i - First)
        Inword = False      'Set in-word flag to false
      End If
    'Does double quote sub-string begin at I?
    ElseIf InStr(Chr$(34), c) > 0 Then
      N = N + 1
      If N > UBound(W$) Then Exit Sub
      'Find positions of first & last double quote
      First = i + 1
      Last = InStr(First, T$, Chr$(34))
      If Last = 0 Then Last = Nchar
      W$(N) = Mid$(T$, First, Last - First)
      Inword = False
      i = Last + 1
    Else
      If Not Inword Then     'Was scanning spaces, so
        Inword = True        'Set in-word flag to true
        First = i            'Mark beginning of word
      End If
    End If
  Next i

  If Inword Then             'Extract last pending word
    N = N + 1
    If N > UBound(W$) Then Exit Sub
    W(N) = Mid$(T$, First, Len(T$) + 1 - First)
  End If

End Sub



Private Sub cmbOut_Click()

If InpFileName <> "" Then
    OutputFileName = InpFileName & "." & UnitsNames(cmbOut.ListIndex) & ".inp"
    lblOutputFileName = OutputFileName
End If

End Sub

Private Sub cmdConvert_Click()

Call SetConversionValues

Call DoConversion

MsgBox "Conversion Done !"

End Sub

Private Sub cmdEnd_Click()

End

End Sub


Private Sub cmdInpFile_Click()
'---------------------------------------------------------------

If CD.InitDir = "" Then CD.InitDir = App.Path

CD.CancelError = False
CD.DialogTitle = "Select INP file..."
CD.Filter = "INP Files (*.inp)|*.inp"
CD.ShowOpen

If CD.FileName <> "" Then
    lblInpFileName = CD.FileName
    InpFileName = CD.FileName
    
    Call GetInpUnits
    
    OutputFileName = InpFileName & "." & UnitsNames(cmbOut.ListIndex) & ".inp"
    lblOutputFileName = OutputFileName
    
    lblInpUnits = InpUnits
    
    cmdConvert.Enabled = True
End If

End Sub


Private Sub Form_Load()

lblInpFileName = ""
lblOutputFileName = ""
lblInpUnits = ""
cmdConvert.Enabled = False

cmbOut.AddItem "CFS (cubic feet / sec)"
cmbOut.AddItem "GPM (gallons / min)"
cmbOut.AddItem "MGD (million gal / day)"
cmbOut.AddItem "IMGD (Imperial MGD)"
cmbOut.AddItem "AFD (acre-feet / day)"
cmbOut.AddItem "LPS (liters / sec)"
cmbOut.AddItem "LPM (liters / min)"
cmbOut.AddItem "MLD (megaliters / day)"
cmbOut.AddItem "CMH (cubic meters / hr)"
cmbOut.AddItem "CMD (cubic meters / day)"

ReDim UnitsNames(0 To 9)
UnitsNames(0) = "CFS"
UnitsNames(1) = "GPM"
UnitsNames(2) = "MGD"
UnitsNames(3) = "IMGD"
UnitsNames(4) = "AFD"
UnitsNames(5) = "LPS"
UnitsNames(6) = "LPM"
UnitsNames(7) = "MLD"
UnitsNames(8) = "CMH"
UnitsNames(9) = "CMD"

cmbOut.Font.Size = 12
cmbOut.ListIndex = 0
cmdConvert.Font.Size = 14
cmdEnd.Font.Size = 14

End Sub


Private Sub Label2_Click()

Dim x
x = ShellExecute(hwnd, "Open", "http://www.optiwater.com", &O0, &O0, SW_NORMAL)

End Sub


Private Sub Label6_Click()

Dim x
x = ShellExecute(hwnd, "Open", "mailto: selad@optiwater.com", &O0, &O0, SW_NORMAL)

End Sub


