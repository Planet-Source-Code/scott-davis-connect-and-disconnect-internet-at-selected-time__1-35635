VERSION 5.00
Begin VB.Form frmNetDialer 
   Caption         =   "NetDial"
   ClientHeight    =   2055
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "News Gothic MT"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDis 
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkCon 
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   480
      Width           =   255
   End
   Begin VB.ComboBox cmbDisAMPM 
      Enabled         =   0   'False
      Height          =   360
      Left            =   4440
      TabIndex        =   18
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox cmbConAMPM 
      Enabled         =   0   'False
      Height          =   360
      Left            =   4440
      TabIndex        =   17
      Top             =   360
      Width           =   615
   End
   Begin VB.Timer TMR1 
      Interval        =   50
      Left            =   1920
      Top             =   1320
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Disconnect at:"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Connect at:"
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cmbDis 
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   3600
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox cmbDis 
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox cmbDis 
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox cmbCon 
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox cmbCon 
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox cmbCon 
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCurrent 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblDis 
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblCon 
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "AM / PM"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Seconds"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Minutes"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Hours"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmNetDialer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '""""""""""""""""""""""""""""""""""""""""""
    '""""""""""Code submitted by:""""""""""""""
    '"""""""""""Scott Davis""""""""""""""""""""
    
    'I have a funky way of doing some things,
    'so this may be a little confusing.
    'I tend to make more variables than needed,
    'but it works.  All I ask, if you use any of this
    'code in your own program, give me credit for my
    'work.  I did this one day, because I was trying
    'to sleep, and phone solicitors wouldnt leave me alone.
    'So I made this, so it would get on when my parents left
    'for work, and then got off when I really needed to get up,
    'and wouldnt tie up the line anymore.  I am hoping to
    'get a new version out soon.  Email me with questions or
    'problems
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Dim i As Integer, s As String, s2 As String, b As Boolean
Dim CTime As String, STime() As String
Dim TTimeC As String, TTimeD As String, TTimeI As String
Dim UTimeCon(0 To 2) As String, UTimeDis(0 To 2) As String
Dim ConAMPM As String, DisAMPM As String
Dim ConChk As Boolean, DisChk As Boolean
Const z As String = ":"

Private Sub chkCon_Click()
    If chkCon.Value = Checked Then
        ConChk = True
        For i = 0 To 2
            cmbCon(i).Enabled = True
        Next
        cmbConAMPM.Enabled = True
        lblCon.Enabled = True
    ElseIf chkCon.Value = Unchecked Then
        ConChk = False
        For i = 0 To 2
            cmbCon(i).Enabled = False
        Next
        cmbConAMPM.Enabled = False
        lblCon.Caption = ""
        lblCon.Enabled = False
    End If
End Sub

Private Sub chkDis_Click()
    If chkDis.Value = Checked Then
        DisChk = True
        For i = 0 To 2
            cmbDis(i).Enabled = True
        Next
        cmbDisAMPM.Enabled = True
        lblDis.Enabled = True
    ElseIf chkDis.Value = Unchecked Then
        DisChk = False
        For i = 0 To 2
            cmbDis(i).Enabled = False
        Next
        cmbDisAMPM.Enabled = False
        lblDis.Caption = ""
        lblDis.Enabled = False
    End If
End Sub

Private Sub cmbCon_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        s2 = CInt(cmbCon(Index).Text)
        If s2 > cmbCon(Index).ListCount Then
            MsgBox "Number is too large, please enter a lower number!", vbInformation, "ERROR"
            If cmbCon(0) = cmbCon(Index) Then
                cmbCon(0).Text = "01"
            ElseIf cmbCon(1) = cmbCon(Index) Or cmbCon(2) = cmbCon(Index) Then
                cmbCon(Index).Text = "00"
            End If
            Exit Sub
        End If
        If s2 < 10 Then
            s = 0 & s2
            cmbCon(Index).Text = s
        Else
            s = s2
            cmbCon(Index).Text = s
        End If
    End If
End Sub

Private Sub cmbCon_LostFocus(Index As Integer)
    s2 = CInt(cmbCon(Index).Text)
    If s2 > cmbCon(Index).ListCount Then
        MsgBox "Number is too large, please enter a lower number!", vbInformation, "ERROR"
        If cmbCon(0) = cmbCon(Index) Then
            cmbCon(0).Text = "01"
        ElseIf cmbCon(1) = cmbCon(Index) Or cmbCon(2) = cmbCon(Index) Then
            cmbCon(Index).Text = "00"
        End If
        Exit Sub
    End If
    If s2 < 10 Then
        s = 0 & s2
        cmbCon(Index).Text = s
    Else
        s = s2
        cmbCon(Index).Text = s
    End If
End Sub

Private Sub cmbDis_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        s2 = CInt(cmbDis(Index).Text)
        If s2 > cmbDis(Index).ListCount Then
            MsgBox "Number is too large, please enter a lower number!", vbInformation, "ERROR"
            If cmbDis(0) = cmbDis(Index) Then
                cmbDis(0).Text = "01"
            ElseIf cmbDis(1) = cmbDis(Index) Or cmbDis(2) = cmbDis(Index) Then
                cmbDis(Index).Text = "00"
            End If
            Exit Sub
        End If
        If s2 < 10 Then
            s = 0 & s2
            cmbDis(Index).Text = s
        Else
            s = s2
            cmbDis(Index).Text = s
        End If
    End If
End Sub

Private Sub cmbDis_LostFocus(Index As Integer)
    s2 = CInt(cmbDis(Index).Text)
    If s2 > cmbDis(Index).ListCount Then
        MsgBox "Number is too large, please enter a lower number!", vbInformation, "ERROR"
        If cmbDis(0) = cmbDis(Index) Then
            cmbDis(0).Text = "01"
        ElseIf cmbDis(1) = cmbDis(Index) Or cmbDis(2) = cmbDis(Index) Then
            cmbDis(Index).Text = "00"
        End If
        Exit Sub
    End If
    If s2 < 10 Then
        s = 0 & s2
        cmbDis(Index).Text = s
    Else
        s = s2
        cmbDis(Index).Text = s
    End If
End Sub

Private Sub cmdSet_Click()
    For i = 0 To 2
        UTimeCon(i) = cmbCon(i).Text
        UTimeDis(i) = cmbDis(i).Text
    Next
    TTimeC = UTimeCon(0) & z & UTimeCon(1) & z & UTimeCon(2)
    TTimeD = UTimeDis(0) & z & UTimeDis(1) & z & UTimeDis(2)
    If CInt(UTimeCon(0)) < 10 Then
        UTimeCon(0) = CStr(Mid(UTimeCon(0), 2))
    End If
    If CInt(UTimeCon(1)) < 10 And CInt(UTimeCon(1)) <> 0 Then
        UTimeCon(1) = CStr(Mid(UTimeCon(1), 2))
    End If
    If CInt(UTimeDis(0)) < 10 Then
        UTimeDis(0) = CStr(Mid(UTimeDis(0), 2))
    End If
    If CInt(UTimeDis(1)) < 10 And CInt(UTimeDis(1)) <> 0 Then
        UTimeDis(1) = CStr(Mid(UTimeDis(1), 2))
    End If
    ConAMPM = cmbConAMPM.Text
    DisAMPM = cmbDisAMPM.Text
    lblCon.Caption = UTimeCon(0) & z & UTimeCon(1) & z & UTimeCon(2) & z & ConAMPM
    lblDis.Caption = UTimeDis(0) & z & UTimeDis(1) & z & UTimeDis(2) & z & DisAMPM
    b = True
End Sub

Private Sub cmdStop_Click()
    b = False
    lblCon.Caption = ""
    lblDis.Caption = ""
End Sub

Private Sub Form_Load()
    For i = 0 To 11 Step 0
        i = i + 1
        If i < 10 Then
            s = "0" & CStr(i)
        Else
            s = CStr(i)
        End If
        cmbCon(0).AddItem s
        cmbDis(0).AddItem s
    Next
    For i = 0 To 59
        If i < 10 Then
            s = "0" & CStr(i)
        Else
            s = CStr(i)
        End If
        cmbCon(1).AddItem s
        cmbCon(2).AddItem s
        cmbDis(1).AddItem s
        cmbDis(2).AddItem s
    Next
    cmbConAMPM.AddItem "AM"
    cmbConAMPM.AddItem "PM"
    cmbDisAMPM.AddItem "AM"
    cmbDisAMPM.AddItem "PM"
    cmbCon(0).Text = "01"
    cmbCon(1).Text = "00"
    cmbCon(2).Text = "00"
    cmbConAMPM.Text = "AM"
    cmbDis(0).Text = "01"
    cmbDis(1).Text = "00"
    cmbDis(2).Text = "00"
    cmbDisAMPM.Text = "AM"
    chkCon.Value = Checked
    chkDis.Value = Checked
    ConChk = True
    DisChk = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
    End If
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub TMR1_Timer()
Dim iTime As Integer
    CTime = Format(Time, "HH:MM:SS")
    STime() = Split(CTime, z)
    If STime(0) > 12 Then
        iTime = CInt(STime(0))
        iTime = iTime - 12
        lblCurrent.Caption = CStr(iTime) & z & STime(1) & z & STime(2)
    Else
        lblCurrent.Caption = CTime
    End If
    If b = True Then
        If ConChk = True Then
            If ConAMPM = "AM" Then
                If CTime = TTimeC Then
                    If (InternetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0)) Then
                        lblCon.Caption = ""
                    End If
                End If
            ElseIf ConAMPM = "PM" Then
                TTimeI = Val(cmbCon(0).Text) + 12
                If CTime = CStr(TTimeI) & Mid(TTimeC, 3) Then
                    If (InternetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0)) Then
                        lblCon.Caption = ""
                    End If
                End If
            End If
        End If
        If DisChk = True Then
            If DisAMPM = "AM" Then
                If CTime = TTimeD Then
                    If (InternetAutodialHangup(0)) Then
                        lblDis.Caption = ""
                    End If
                End If
            ElseIf DisAMPM = "PM" Then
                TTimeI = Val(cmbDis(0).Text) + 12
                If CTime = CStr(TTimeI) & Mid(TTimeD, 3) Then
                    If (InternetAutodialHangup(0)) Then
                        lblDis.Caption = ""
                    End If
                End If
            End If
        End If
    End If
    TMR1.Interval = 50
End Sub
