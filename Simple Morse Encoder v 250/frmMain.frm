VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Simple Morse Encoder v 2.5"
   ClientHeight    =   2235
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Speed / Frequency:"
      Height          =   975
      Left            =   5400
      TabIndex        =   10
      Top             =   1140
      Width           =   2475
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
         _Version        =   393216
         Max             =   200
         SelStart        =   146
         TickStyle       =   1
         TickFrequency   =   10
         Value           =   146
      End
      Begin MSComctlLib.Slider sldFreq 
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   600
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
         _Version        =   393216
         Max             =   2000
         SelStart        =   520
         TickStyle       =   1
         TickFrequency   =   100
         Value           =   520
      End
      Begin VB.Label Label2 
         Caption         =   "Hz"
         Height          =   255
         Left            =   1740
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "mSec"
         Height          =   240
         Left            =   1740
         TabIndex        =   11
         Top             =   300
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdTransmit 
      Caption         =   "&Transmit"
      Default         =   -1  'True
      Height          =   435
      Left            =   6600
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdPause 
      Cancel          =   -1  'True
      Caption         =   "&Pause"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   180
      Width           =   1215
   End
   Begin VB.Frame fraTranslate 
      Caption         =   "Translate:"
      Height          =   975
      Left            =   2820
      TabIndex        =   6
      Top             =   1140
      Width           =   2475
      Begin VB.CommandButton cmdTranslateE2M 
         Caption         =   "English > Morse"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   600
         Width           =   2115
      End
      Begin VB.CommandButton cmdTranslateM2E 
         Caption         =   "Morse > English"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   2115
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output Through:"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1140
      Width           =   2475
      Begin VB.OptionButton optSoundcard 
         Caption         =   "Soundcard (DirectX7)"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optSpeaker 
         Caption         =   "Internal Speaker"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1755
      End
   End
   Begin VB.TextBox txtEnglish 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Text            =   "HELLO"
      Top             =   660
      Width           =   6195
   End
   Begin VB.TextBox txtMorse 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   6195
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------------------
'   ALWAYS REMEMBER:
'                 "When everything else fails, read the instructions"
'--------------------------------------------------------------------
'       Please refer to the README module for more information
'--------------------------------------------------------------------
Public aPause As Boolean

Private Sub cmdPause_Click()
    If aPause = False Then
        aPause = True
        cmdPause.Caption = "&Resume"
    Else
        aPause = False
        cmdPause.Caption = "&Pause"
    End If
End Sub

Private Sub cmdTranslateE2M_Click()
    txtMorse.Text = modMorseCode.EncodeToMorse(Trim(txtEnglish.Text))
End Sub

Private Sub cmdTranslateM2E_Click()
    txtEnglish.Text = modMorseCode.DecodeToEnglish(Trim(txtMorse.Text))
End Sub

Private Sub cmdTransmit_Click()
    ' if the txtMorse textbox is empty, encode the english text into morse
    ' and then use it... If it is not empty, then probably the user want to type
    ' in morse him/her self. Allow it.
    If Trim(txtMorse.Text) = "" Then
        txtMorse.Text = modMorseCode.EncodeToMorse(Trim(txtEnglish.Text))
    End If
    
    ' We will now choose our function call... depending on which output
    ' device is chosen.
    ' NOTE another thing, the speed is subtracted from the max.
    ' Well, what the user sets through sldSpeed is the Delay, and we sent the
    ' speed from here.
    If optSpeaker.Value = True Then
        modMorseCode.PlayMorse txtMorse.Text, (sldSpeed.Max - sldSpeed.Value), sldFreq.Value, PCSpeaker
    Else
        modMorseCode.PlayMorse txtMorse.Text, (sldSpeed.Max - sldSpeed.Value), sldFreq.Value, Soundcard
    End If
    
    ' Now clear the txtMorse textbox
    txtMorse.Text = ""
    
    ' Select the text in txtenglish textbox, so that if the user wishes,
    ' s/he can type on and the selected text will be replaced OR the user
    ' may want to retransmit the message.
    txtEnglish.SetFocus
    ' send the HOME keystroke
    SendKeys "{HOME}"
    ' allow OS to complete the task...
    DoEvents
    ' send the SHIFT-END keystroke (+ Shift, ^ Ctrl, % Alt)
    SendKeys "+{END}"
End Sub

Private Sub Form_Load()
    'confirm that our form is being shown NOW
    Me.Show
    ' let OS do it...
    DoEvents
    ' we have to first iniatialize the SimplestOscillator because is uses
    ' DirectX for outputting the sound.
    modOscillator.InitializeBeep Me.hWnd
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Simple Morse Encoder v 2.5" & vbCrLf & "By: Harshad Sharma (harshad.sharma@bigfoot.com)", vbInformation, "Simple Morse Encoder 2.5"
End Sub

Private Sub mnuHelpContents_Click()
    Dim Message As String
    
End Sub

