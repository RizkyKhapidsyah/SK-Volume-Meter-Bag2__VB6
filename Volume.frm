VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Volume"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Chiudi"
      Default         =   -1  'True
      Height          =   360
      Left            =   270
      TabIndex        =   0
      Top             =   3810
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   2805
      Left            =   195
      TabIndex        =   3
      Top             =   180
      Width           =   1305
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1305
         Left            =   390
         ScaleHeight     =   1305
         ScaleWidth      =   540
         TabIndex        =   4
         Top             =   765
         Width           =   540
         Begin VB.Image Image1 
            Height          =   165
            Left            =   120
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Massimo"
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   8
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimo"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   7
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   270
         Left            =   150
         TabIndex        =   6
         Top             =   2175
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   210
         Left            =   45
         TabIndex        =   5
         Top             =   465
         Width           =   1230
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   285
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Volume
Dim vol As Long
Dim hmixer As Long          ' mixer handle
Dim volCtrl As MIXERCONTROL ' waveout volume control
Dim rc As Long              ' return code
Dim ok As Boolean           ' boolean return code
Dim volMin As Long
Dim volMax As Long
'Cursore
Dim Mouse_Button As Integer
Dim Mouse_Y As Single
Dim PicFrame As Integer
Private Sub Cursore()
   Timer1.Enabled = False
   Timer1.Interval = 10
   PicFrame = 2 * Screen.TwipsPerPixelY
   Image1.Top = (Picture4.Height - Image1.Height) / 4
   Image1.Left = Picture4.Width / 2 - Image1.Width / 2
   Text1.Text = Fix(Abs(CDbl(volMax * (Image1.Top) / (Picture4.Height - Image1.Height)) - volMax))
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    InitVolume
    Cursore
End Sub
Private Sub Timer1_Timer()
   Dim TempTop As Integer
   If Mouse_Button = 1 Then
      TempTop = Image1.Top
      TempTop = Mouse_Y - TempHeight / 2
      If TempTop + Image1.Height > Picture4.Height Then
         TempTop = Picture4.Height - Image1.Height '- PicFrame
      End If
      If TempTop < 0 Then TempTop = 0
   End If
   If Image1.Top <> TempTop Then Image1.Top = TempTop
   Text1.Text = Fix(Abs(CDbl(volMax * (Image1.Top) / (Picture4.Height - Image1.Height)) - volMax))
End Sub
Private Sub InitVolume()
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Non posso aprire il mixer."
        Exit Sub
    End If
    ok = GetVolumeControl(hmixer, _
                         MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                         MIXERCONTROL_CONTROLTYPE_VOLUME, _
                         volCtrl)
    If (ok = True) Then
        volMin = volCtrl.lMinimum
        volMax = volCtrl.lMaximum
        Label2.Caption = volMin
        Label3.Caption = volMax
    End If
End Sub
Private Sub SettaVolume()
    vol = CLng(Text1.Text)
    SetVolumeControl hmixer, volCtrl, vol
End Sub
Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer1.Enabled = True
  Mouse_Button = Button
End Sub
Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Mouse_Button = Button
   Mouse_Y = Image1.Top + Y
   SettaVolume
End Sub
Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Timer1.Enabled = False
   Mouse_Button = Button
    SettaVolume
End Sub
