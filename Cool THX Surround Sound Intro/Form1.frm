VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cool THX Surround Sound Intro"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   2378
      Picture         =   "Form1.frx":0D9C
      ScaleHeight     =   2835
      ScaleWidth      =   7245
      TabIndex        =   0
      Top             =   3083
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF5000&
      BorderWidth     =   2
      Height          =   4890
      Left            =   960
      Top             =   2048
      Visible         =   0   'False
      Width           =   10065
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF4000&
      BorderWidth     =   4
      Height          =   4815
      Left            =   1006
      Top             =   2086
      Visible         =   0   'False
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created By: Derek Skeba - Everything but the resolution changer
'Email me at frillyozz@comcast.net
'Thanks for downloading this source code

Option Explicit
Const WM_DISPLAYCHANGE = &H7E
Const HWND_BROADCAST = &HFFFF&
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H4
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const BITSPIXEL = 12
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Dim OldX As Long, OldY As Long, nDC As Long
Dim X As Double
Private Sub Form_KeyPress(KeyAscii As Integer)
sndPlaySound App.Path + "", SND_NODEFAULT + SND_ASYNC
ChangeRes OldX, OldY, GetDeviceCaps(nDC, BITSPIXEL)
DeleteDC nDC
sndPlaySound App.Path + "", SND_NODEFAULT + SND_ASYNC
End
End Sub
Private Sub Form_Load()
Dim nDC As Long
OldX = Screen.Width / Screen.TwipsPerPixelX
OldY = Screen.Height / Screen.TwipsPerPixelY
nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
ChangeRes 800, 600, GetDeviceCaps(nDC, BITSPIXEL)
Me.CurrentX = 10
Me.CurrentY = 10
Me.Print "(Press Any Key To Exit)"
End Sub
Private Sub Timer1_Timer()
X = X + 0.25
If X = 1 Then Shape1.Visible = True Else
If X = 1.25 Then Shape2.Visible = True Else
If X = 2 Then Picture1.Visible = True Else
If X = 2 Then sndPlaySound App.Path + "\thx.WAV", SND_NODEFAULT + SND_ASYNC
If X = 7 Then Shape2.Visible = False
If X = 7.25 Then Shape1.Visible = False
If X = 8.25 Then Picture1.Visible = False
If X = 8.5 Then
Me.CurrentX = 200
Me.CurrentY = 250
Me.FontSize = 25
Me.Print "Created by: Derek Skeba"
End If
If X = 9 Then
sndPlaySound App.Path + "", SND_NODEFAULT + SND_ASYNC
ChangeRes OldX, OldY, GetDeviceCaps(nDC, BITSPIXEL)
DeleteDC nDC
sndPlaySound App.Path + "", SND_NODEFAULT + SND_ASYNC
End
End If
End Sub
Sub ChangeRes(X As Long, Y As Long, Bits As Long)
Dim DevM As DEVMODE, ScInfo As Long, erg As Long, an As VbMsgBoxResult
erg = EnumDisplaySettings(0&, 0&, DevM)
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
DevM.dmPelsWidth = X
DevM.dmPelsHeight = Y
DevM.dmBitsPerPel = Bits
erg = ChangeDisplaySettings(DevM, CDS_TEST)
End Sub
