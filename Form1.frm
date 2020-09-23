VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3435
   ClientLeft      =   2760
   ClientTop       =   1665
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   Begin VB.TextBox R02Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   3435
      TabIndex        =   29
      Top             =   3060
      Width           =   495
   End
   Begin VB.TextBox BYTESsentText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   4410
      TabIndex        =   27
      Top             =   3060
      Width           =   540
   End
   Begin VB.TextBox SPENDINGLIMITtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox RATINGtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox GUIDEtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TIMEZONEtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox IRDText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox USWtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox FUSEtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox CardIDtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox COMMlist 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0002
      TabIndex        =   8
      Text            =   "Select COM Port"
      Top             =   1080
      Width           =   1905
   End
   Begin VB.TextBox BuffCntText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   2895
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3060
      Width           =   180
   End
   Begin VB.TextBox TextInReadBuffer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1410
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   6180
      Width           =   855
   End
   Begin VB.CommandButton CARDinfoBtn 
      Appearance      =   0  'Flat
      Caption         =   "Get Card Info"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1380
   End
   Begin VB.TextBox COMtext 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5370
      Width           =   615
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2445
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   6180
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "The basics of reading DSS digital TV smartcard  H / P2 series "
      Height          =   405
      Left            =   1695
      TabIndex        =   35
      Top             =   600
      Width           =   3075
   End
   Begin VB.Label Label9 
      Caption         =   "          Visual Basic Smartcard Interface      Compliments of The 2001 The Hickware CREW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   135
      TabIndex        =   31
      Top             =   1725
      Width           =   4440
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PORT"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   2445
      TabIndex        =   34
      Top             =   1155
      Width           =   405
   End
   Begin VB.Image PORTLITE 
      Appearance      =   0  'Flat
      Height          =   210
      Left            =   2205
      Picture         =   "Form1.frx":0004
      Top             =   1140
      Width           =   210
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RX"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   33
      Top             =   3150
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TX"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   32
      Top             =   3150
      Width           =   255
   End
   Begin VB.Image RXLITE 
      Appearance      =   0  'Flat
      Height          =   210
      Left            =   6600
      Picture         =   "Form1.frx":02AE
      Top             =   3120
      Width           =   210
   End
   Begin VB.Image TXLITE 
      Appearance      =   0  'Flat
      Height          =   210
      Left            =   6120
      Picture         =   "Form1.frx":0558
      Top             =   3120
      Width           =   210
   End
   Begin VB.Image RXOFF 
      Height          =   210
      Left            =   2310
      Picture         =   "Form1.frx":0802
      Top             =   3915
      Width           =   210
   End
   Begin VB.Image TXOFF 
      Height          =   210
      Left            =   2100
      Picture         =   "Form1.frx":0AAC
      Top             =   3915
      Width           =   210
   End
   Begin VB.Image PortOFF 
      Height          =   210
      Left            =   1920
      Picture         =   "Form1.frx":0D56
      Top             =   3930
      Width           =   210
   End
   Begin VB.Image RXON 
      Height          =   210
      Left            =   1725
      Picture         =   "Form1.frx":1000
      Top             =   3930
      Width           =   210
   End
   Begin VB.Image TXON 
      Height          =   210
      Left            =   1530
      Picture         =   "Form1.frx":12AA
      Top             =   3930
      Width           =   210
   End
   Begin VB.Image PortON 
      Height          =   210
      Left            =   1335
      Picture         =   "Form1.frx":1554
      Top             =   3930
      Width           =   210
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Realtime Card Responses   -    "
      Height          =   195
      Left            =   195
      TabIndex        =   30
      Top             =   3060
      Width           =   2190
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Sent:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3960
      TabIndex        =   28
      Top             =   3120
      Width           =   450
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   180
      TabIndex        =   26
      Top             =   2685
      Width           =   4770
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "LIMIT:"
      Height          =   255
      Left            =   5115
      TabIndex        =   25
      Top             =   2715
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "RATING:"
      Height          =   255
      Left            =   4905
      TabIndex        =   23
      Top             =   2355
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "TIMEZONE:"
      Height          =   195
      Left            =   4710
      TabIndex        =   21
      Top             =   2025
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "GUIDE:"
      Height          =   255
      Left            =   5025
      TabIndex        =   20
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IRD:"
      Height          =   255
      Left            =   5235
      TabIndex        =   19
      Top             =   540
      Width           =   375
   End
   Begin VB.Label USWlabel 
      AutoSize        =   -1  'True
      Caption         =   "USW:"
      Height          =   255
      Left            =   5130
      TabIndex        =   15
      Top             =   915
      Width           =   495
   End
   Begin VB.Label FuseLabel 
      AutoSize        =   -1  'True
      Caption         =   "FUSE:"
      Height          =   255
      Left            =   5115
      TabIndex        =   14
      Top             =   1275
      Width           =   495
   End
   Begin VB.Label CardIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "CARDID:"
      Height          =   195
      Left            =   4935
      TabIndex        =   11
      Top             =   195
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "INS:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3090
      TabIndex        =   9
      Top             =   3135
      Width           =   360
   End
   Begin VB.Label BuufCntLabel 
      AutoSize        =   -1  'True
      Caption         =   "BytesIn:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   2160
      TabIndex        =   7
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   " Text in COMM buffer for DEBUG"
      Height          =   585
      Left            =   90
      TabIndex        =   5
      Top             =   6210
      Width           =   1320
   End
   Begin VB.Label ATRlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   225
      TabIndex        =   1
      Top             =   120
      Width           =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'This was done by us as a "fun project" and is by no means the
'clean crisp code as we did in C/C++. This is being released to
'the public domain as open source and is meant for your
'educational purposes ONLY. Seems NO ONE on the web knows how to
'interface with a smartcard, thus this donation. If you get
'in trouble using this source code, you only have yourself
'to blame. We take NO responsibility for what you do with this
'code. This code can easily be modified to be used in ANY type
'setting where smartcard standards are strictly ISO7816 compliant.
'Of course you need a ISO7816 smartcard read/write device to
'use this program. If you dont know what that is then please
'delete this code ASAP ;-)
'
'YOU USE THIS CODE AT YOUR OWN RISK!
'
'Compliments of The 2001/2002 Hickware CREW
'

Private Sub CardIDtext_GotFocus()

On Error Resume Next
 CARDinfoBtn.SetFocus
Exit Sub


End Sub

Private Sub Cardinfobtn_Click()

CardInserted = True

ShowStatus "Clearing Fields"

Call ToggleButtons

Call ClearVariables

If port = "" Then
   MsgBox "No COM port selected!"
   Call CloseCOMM
   Exit Sub
End If

ShowStatus "Resetting ATR"
Call ResetForWrite

DelaySecs 0.25

If AtrLen < 38 Then Call CloseCOMM: Call ToggleButtons: Exit Sub

ShowStatus "Sending Data"
Call SendData(CardInfoStr)

ShowStatus "Reading Data"
Call ReadDATA

ShowStatus "Parsing Data": DoEvents
Call ShowDATA: Call CardInfo2A(CardInfoBuffer)
'--------------------
Call ClearVariables
'--------------------
ShowStatus "Sending Data"
Call SendData(IRDinfoStr)
ShowStatus "Reading Data"
Call ReadDATA
ShowStatus "Displaying Data"
Call ShowDATA: Call CardInfo58(CardInfoBuffer)
ShowStatus "Sending Data"
Call SendData(PPVinfoStr)
ShowStatus "Reading Data"
Call ReadDATA
ShowStatus "Displaying PPV Data":
Call ShowDATA: Call CardInfoPPV(CardInfoBuffer)
ShowStatus "Closing Comport"
Call CloseCOMM

Call ToggleButtons

ShowStatus "Done"

CardInserted = False

End Sub

Private Sub COMMlist_Click()

If port > "" Then Call CloseCOMM

port = Form1.COMMlist.Text

COMtext.Text = port

Call CheckCOM(port)
 DelaySecs 0.25

End Sub

Private Sub Form_Load()

TFile$ = App.Path & "\" & App.EXEName & ".exe"
'Use ascii to make it hard to hex edit our text`s
'=====================================================================================
'H<space>Card
titleA$ = Chr$(72) + Chr$(32) + _
          Chr$(67) + Chr$(97) + Chr$(114) + Chr$(100)

'<space>
titleB$ = Chr$(32)
'Utility
titleC$ = Chr$(85) + Chr$(116) + Chr$(105) + Chr$(108) + _
          Chr$(105) + Chr$(116) + Chr$(121)
          
'<72 spaces> v2.0
titleD$ = Space$(72) + "1.0b"

titleE$ = titleA$ + titleB$ + titleC$ + titleD$
Form1.Caption = titleE$

COMMlist.AddItem "COM1"
COMMlist.AddItem "COM2"
COMMlist.AddItem "COM3"
COMMlist.AddItem "COM4"

Me.Show

Call GetState

If port > "" Then
   Select Case port
    Case Is = "COM1"
     COMMlist.Text = COMMlist.List(0)
     Call ToggleButtons:
    Case Is = "COM2"
     COMMlist.Text = COMMlist.List(1)
     Call ToggleButtons:
    Case Is = "COM3"
     COMMlist.Text = COMMlist.List(2)
     Call ToggleButtons:
    Case Is = "COM4"
     COMMlist.Text = COMMlist.List(3)
     Call ToggleButtons:
    Case Else
    
  End Select
 
 End If
  
 Call CheckCOM(port)
 

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call SaveState

Call CloseCOMM
DelaySecs 0.0025

Unload Me
End

End Sub

Private Sub R02Label_Change()

If Len(R02Label.Text) > 2 Then
   Label5.Caption = "ACK:"
 Else
  Label5.Caption = "INS:"
End If


End Sub

