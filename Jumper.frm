VERSION 5.00
Begin VB.Form frmJumper 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jumper"
   ClientHeight    =   2190
   ClientLeft      =   945
   ClientTop       =   2760
   ClientWidth     =   9255
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
   Icon            =   "Jumper.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2190
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5175
      TabIndex        =   1
      Top             =   1560
      Width           =   1080
   End
   Begin VB.CommandButton cmdAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4050
      TabIndex        =   2
      Top             =   1560
      Width           =   1080
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2925
      TabIndex        =   3
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   90
      Left            =   1050
      Top             =   1575
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.PictureBox picWinbox 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         Picture         =   "Jumper.frx":030A
         ScaleHeight     =   495
         ScaleWidth      =   1815
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image imgBox 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   4
         Left            =   3120
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgBox 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   3
         Left            =   2160
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgBox 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   2
         Left            =   1200
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgBox 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   9
         Left            =   7920
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgBox 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   8
         Left            =   6960
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgBox 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   7
         Left            =   6000
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgBox 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   6
         Left            =   5040
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgBox 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   1
         Left            =   240
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Image imgBox_left 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   7050
      Picture         =   "Jumper.frx":08B6
      Top             =   2325
      Width           =   855
   End
   Begin VB.Image imgBox_right 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   8040
      Picture         =   "Jumper.frx":1058
      Top             =   2340
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Jumper - Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmJumper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TT(1 To 9) As Integer
Dim SS(1 To 9) As Integer
Dim RR(1 To 9) As Integer
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim N As Integer
Dim Ap As String
Dim Win As Boolean

Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Sub imgBox_Click(Index As Integer)

Dim E As Integer

For N = 1 To 9
    If RR(N) = Index Then E = N
Next
If TT(E) = -1 Then
    If E - 1 > 0 Then
        If TT(E - 1) = 0 Then
            imgBox(Index).Left = SS(E - 1)
            TT(E) = 0
            TT(E - 1) = -1
            RR(E) = 0
            RR(E - 1) = Index
       Else
            If E - 2 > 0 Then
                If TT(E - 2) = 0 Then
                    imgBox(Index).Left = SS(E - 2)
                    TT(E) = 0
                    TT(E - 2) = -1
                    RR(E) = 0
                    RR(E - 2) = Index
                Else
                    Beep
                End If
            End If
       End If
    End If
End If
If TT(E) = 1 Then
    If E + 1 < 10 Then
        If TT(E + 1) = 0 Then
            imgBox(Index).Left = SS(E + 1)
            TT(E) = 0
            TT(E + 1) = 1
            RR(E) = 0
            RR(E + 1) = Index
       Else
            If E + 2 < 10 Then
                If TT(E + 2) = 0 Then
                    imgBox(Index).Left = SS(E + 2)
                    TT(E) = 0
                    TT(E + 2) = 1
                    RR(E) = 0
                    RR(E + 2) = Index
                Else
                    Beep
                End If
            End If
       End If
    End If
End If
CheckPos

End Sub

Sub CheckPos()

Dim E As Integer

For N = 1 To 4
   If TT(N) = -1 Then E = E + 1
Next
For N = 6 To 9
   If TT(N) = 1 Then E = E + 1
Next
If E = 8 Then
    For N = 1 To 4
        Beep
    Next
    Win = True
    picMain.SetFocus
Else
    E = 0
End If

End Sub

Sub cmdAbout_Click()

frmAbout.Show 1

End Sub

Sub cmdExit_Click()

End

End Sub

Sub cmdNew_Click()

picMain.SetFocus
Init

End Sub

Sub Form_Load()

If App.PrevInstance = True Then
    MsgBox "This Application is already running!", 48, "Applcation Running"
    End
End If
If Right(App.Path, 1) = "\" Then
    Ap = App.Path
Else
    Ap = App.Path & "\"
End If

CentreMe Me
Me.Show
App.HelpFile = Ap & "Jumper.hlp"
picMain.BackColor = &HC0C0C0
For N = 1 To 4
    SS(N) = imgBox(N).Left
Next
For N = 6 To 9
    SS(N) = imgBox(N).Left
Next
SS(5) = imgBox(4).Left + imgBox(2).Left - imgBox(1).Left
Init

End Sub

Sub Form_Paint()

Dim E As Integer

E = 0
For N = 1 To 2
    frmJumper.Line (E, E)-(frmJumper.ScaleWidth - E - 15, E), RGB(255, 255, 255)
    frmJumper.Line (E, E)-(E, frmJumper.ScaleHeight - E - 15), RGB(255, 255, 255)
    frmJumper.Line (E, frmJumper.ScaleHeight - E - 15)-(frmJumper.ScaleWidth, frmJumper.ScaleHeight - E - 15)
    frmJumper.Line (frmJumper.ScaleWidth - E - 15, E)-(frmJumper.ScaleWidth - E - 15, frmJumper.ScaleHeight)
    E = E + 15
Next
E = 0

End Sub

Sub mnuAbout_Click()

cmdAbout_Click

End Sub

Private Sub mnuContents_Click()

Dim RetVal As Long

RetVal = WinHelp(Me.hwnd, App.HelpFile, 1, CLng(10))

End Sub

Sub mnuNew_Click()

cmdNew_Click

End Sub

Sub mnuExit_Click()

End

End Sub

Sub Init()

Win = False
picWinbox.Visible = False
For N = 1 To 4
    imgBox(N).Visible = False
    TT(N) = 1
Next
For N = 6 To 9
    imgBox(N).Visible = False
    TT(N) = -1
Next
TT(5) = 0
For N = 1 To 4
    RR(N) = N
Next
For N = 6 To 9
    RR(N) = N
Next
RR(5) = 0
For N = 1 To 4
    imgBox(N).Left = SS(N)
    imgBox(N).Picture = imgBox_left.Picture
    imgBox(N).Visible = True
Next
For N = 6 To 9
    imgBox(N).Left = SS(N)
    imgBox(N).Picture = imgBox_right.Picture
    imgBox(N).Visible = True
Next

End Sub

Sub picMain_Paint()

Dim P As Integer

For N = 1 To 2
    picMain.Line (P, P)-(picMain.Width - P, P)
    picMain.Line (P, P)-(P, picMain.Height - P)
    picMain.Line (P, picMain.Height - P)-(picMain.Width, picMain.Height - P), RGB(255, 255, 255)
    picMain.Line (picMain.Width - P, P)-(picMain.Width - P, picMain.Height), RGB(255, 255, 255)
    P = P + 15
Next
P = 0

End Sub

Sub Timer1_Timer()

If Win = True Then
    For N = 1 To 4
        imgBox(N).Visible = False
    Next
    For N = 6 To 9
        imgBox(N).Visible = False
    Next
    picWinbox.Visible = True
    A = Abs(Sin(B)) * 300
    If C = 0 Then
        B = B + 75
    Else
        B = B - 75
    End If
    picWinbox.Move B + 45, A + 75
    If B > 9015 - 1815 - 90 Then C = -1
    If B < 75 Then C = 0
End If

End Sub

Private Sub CentreMe(F As Form)

F.Left = (Screen.Width - F.Width) / 2
F.Top = (Screen.Height - F.Height) / 2

End Sub

