VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmdesktop 
   BorderStyle     =   0  'None
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   5385
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   5385
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   9499
      ButtonWidth     =   873
      ButtonHeight    =   1244
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Default         =   -1  'True
         Height          =   735
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   4440
         Width           =   735
         Begin VB.Label shotnum 
            Caption         =   "12"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   615
         End
      End
      Begin MSComctlLib.ProgressBar wait 
         Height          =   8175
         Left            =   480
         TabIndex        =   3
         Top             =   4800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   14420
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar Clip 
         Height          =   8175
         Left            =   0
         TabIndex        =   2
         Top             =   4800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   14420
         _Version        =   393216
         Appearance      =   1
         Max             =   12
         Orientation     =   1
      End
   End
   Begin VB.Timer tmrClip 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   0
   End
   Begin VB.Timer check 
      Interval        =   1
      Left            =   840
      Top             =   0
   End
   Begin VB.Image bullet 
      Height          =   240
      Index           =   0
      Left            =   2280
      Picture         =   "frmdesktop.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin Project1.MikesGun MikesGun 
      Height          =   2895
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
   End
End
Attribute VB_Name = "frmdesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim a As Long
Dim s As Long
Dim b As Integer
Dim z As Integer

Dim xx As Integer
Dim yy As Integer

Dim random As Integer
Private Sub cmdExit_Click()
    End
End Sub
'[FORM LOAD]
Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    MikesGun.Top = frmdesktop.Height - MikesGun.Height
    
    z = 60
    Clip.Value = 12
'[GET SCREEN SHOT / SAVE SCREEN SHOT / LOAD SCREEN SHOT]
    frmdesktop.AutoRedraw = True
    frmdesktop.ScaleMode = 1
    a = GetDesktopWindow()
    s = GetDC(a)
    BitBlt frmdesktop.hDC, 0, 0, Screen.Width, Screen.Height, s, 0, 0, vbSrcCopy
    SavePicture Image, ("Desktop.bmp")
    DoEvents
    frmdesktop.Picture = LoadPicture("Desktop.bmp")
    frmdesktop.AutoRedraw = False
End Sub
'[FORM MOUSE MOVE]
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'[X,Y]
    xx = X
    yy = Y
'[PISTOL]
    MikesGun.Move X
End Sub
Private Sub bullet_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Randomize
'[PISTOL]
    If Clip.Value > 0 Then
        Play_Sound (App.Path & "\GunFire.wav"), &H0
        b = b + 1
        Load bullet(b)
        random = Int(Rnd * 2)
        Select Case random
            Case 0
                bullet(b).Left = X + Rnd * 200
                bullet(b).Top = Y + Rnd * 200
            Case 1
                bullet(b).Left = X - Rnd * 200
                bullet(b).Top = Y - Rnd * 200
        End Select
        bullet(b).Visible = True
        Clip.Value = Clip.Value - 1
    End If
'[CLIP]
    If Clip.Value = 0 Then
        tmrWait.Enabled = True
    End If
End Sub
'[MOUSE DOWN / UP]
'[FORM MOUSE DOWN]
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Randomize
    '[PISTOL]
    If Clip.Value > 0 Then
        Play_Sound (App.Path & "\GunFire.wav"), &H0
        b = b + 1
        Load bullet(b)
        random = Int(Rnd * 2)
        Select Case random
            Case 0
                bullet(b).Left = X + Rnd * 200
                bullet(b).Top = Y + Rnd * 200
            Case 1
                bullet(b).Left = X - Rnd * 200
                bullet(b).Top = Y - Rnd * 200
        End Select
        bullet(b).Visible = True
        Clip.Value = Clip.Value - 1
    End If
    '[CLIP]
    If Clip.Value = 0 Then
    tmrWait.Enabled = True
    End If
End Sub
'[FORM MOUSE UP]
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrClip.Enabled = False
End Sub
Private Sub check_Timer()
    shotnum.Caption = Clip.Value
End Sub
'[CLIP]
Private Sub tmrClip_Timer()
'[PISTOL]
If Clip.Value > 0 Then
Clip.Value = Clip.Value - 1
End If
'[WAIT]
If Clip.Value = 0 Then
tmrWait.Enabled = True
End If
End Sub
'[RELOAD]
Private Sub tmrWait_Timer()
    '[WAIT]
    If wait.Value < 100 Then
        wait.Value = wait.Value + 1
    End If
    '[PISTOL RELOAD]
    If wait.Value = 100 Then
        Play_Sound (App.Path & "\SHOTGUNRELOAD.wav"), &H0
        tmrWait.Enabled = False
        wait.Value = 0
        Clip.Value = 12
    End If
End Sub

