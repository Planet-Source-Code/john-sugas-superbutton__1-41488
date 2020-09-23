VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Testing SuperButton"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0E42
   ScaleHeight     =   7215
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin Project1.SuperButton SuperButton11 
      Height          =   1215
      Left            =   5820
      TabIndex        =   21
      ToolTipText     =   "Brushes, Pens and ROP's...."
      Top             =   5460
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   2143
      FontShadowLeft  =   14737632
      FontShadowRight =   49344
      BackColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   3
      Object.ToolTipText     =   "Brushes, Pens and ROP's...."
      CaptionLeft     =   -3
      CaptionTop      =   8
      ToolTipShowDelay=   100
      ToolTipTimeShown=   1000
      CaptionCenterY  =   0   'False
      ToolTipBackColor=   8438015
      Caption         =   "SuperButton!"
      FontUsePen      =   -1  'True
      FontColorBrush  =   8438015
      FontUseBrush    =   -1  'True
      FontEscapement  =   -10
      FontCharSpacing =   20
      FontROP         =   2
      BackStyle       =   1
   End
   Begin Project1.SuperButton SuperButton10 
      Height          =   975
      Left            =   7800
      TabIndex        =   20
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      CaptionLeft     =   36
      CaptionTop      =   32
      Picture         =   "Form1.frx":BD14
   End
   Begin Project1.SuperButton SuperButton9 
      Height          =   435
      Left            =   2340
      TabIndex        =   19
      Top             =   6060
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      MaskColor       =   12632256
      PictureDown     =   "Form1.frx":CB66
      CaptionLeft     =   16
      CaptionTop      =   14
      PicDownMove     =   0   'False
      MaskColorDownPic=   12632256
      Picture         =   "Form1.frx":CCDD
      EdgeType        =   0
   End
   Begin Project1.SuperButton SuperButton8 
      Height          =   795
      Left            =   9660
      TabIndex        =   18
      ToolTipText     =   "Exit Demo"
      Top             =   6300
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1402
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      Object.ToolTipText     =   "Exit Demo"
      PictureUp       =   "Form1.frx":CE73
      CaptionLeft     =   30
      CaptionTop      =   26
      ToolTipShowDelay=   100
      PicDownMove     =   0   'False
      MaskColorDownPic=   16777215
      Picture         =   "Form1.frx":D18D
   End
   Begin Project1.SuperButton SuperButton7 
      Height          =   1335
      Left            =   780
      TabIndex        =   17
      Top             =   3060
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      FontColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   12632256
      PictureUp       =   "Form1.frx":D4A7
      PictureDown     =   "Form1.frx":DB21
      CaptionLeft     =   16
      CaptionTop      =   60
      ToolTipShowDelay=   100
      MaskColorUpPic  =   12632256
      MaskColorDownPic=   12632256
      CaptionCenterY  =   0   'False
      Picture         =   "Form1.frx":E19B
      Caption         =   "Bitmaps"
   End
   Begin Project1.SuperButton SuperButton6 
      Height          =   1335
      Left            =   780
      TabIndex        =   16
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      FontColor       =   16777215
      FontShadowLeft  =   12632256
      FontShadowRight =   16576
      BackColor       =   4210752
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   48
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   3
      CaptionLeft     =   18
      CaptionTop      =   10
      ToolTipShowDelay=   100
      CaptionCenterX  =   0   'False
      CaptionCenterY  =   0   'False
      Caption         =   "T"
      FontUsePen      =   -1  'True
      FontUseBrush    =   -1  'True
      FontShadowOffsetY=   -3
      FontShadow      =   1
      BackStyle       =   1
      EdgeType        =   4
   End
   Begin Project1.SuperButton SuperButton5 
      Height          =   465
      Index           =   0
      Left            =   5820
      TabIndex        =   12
      ToolTipText     =   "BMP Pic's"
      Top             =   0
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   3
      AutoSize        =   -1  'True
      MaskColor       =   12632256
      Object.ToolTipText     =   "BMP Pic's"
      CaptionLeft     =   13
      CaptionTop      =   13
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":E815
      EdgeType        =   3
   End
   Begin Project1.SuperButton SuperButton4 
      Height          =   555
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   "Icon Pictures"
      Top             =   0
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Icon Pictures"
      CaptionLeft     =   18
      CaptionTop      =   18
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":E927
   End
   Begin Project1.SuperButton SuperButton3 
      Height          =   2175
      Left            =   6900
      TabIndex        =   6
      Top             =   1140
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3836
      FontColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   5
      CaptionLeft     =   100
      CaptionTop      =   72
      PictureOffsetY  =   -40
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":EC41
      EdgeType        =   2
   End
   Begin Project1.SuperButton SuperButton2 
      Height          =   1335
      Left            =   780
      TabIndex        =   5
      Top             =   1140
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      FontColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   3
      PictureUp       =   "Form1.frx":FA93
      PictureDown     =   "Form1.frx":108E5
      CaptionLeft     =   21
      CaptionTop      =   60
      ToolTipShowDelay=   100
      CaptionCenterY  =   0   'False
      Picture         =   "Form1.frx":11737
      Caption         =   "Icons"
      EdgeType        =   3
   End
   Begin Project1.SuperButton SuperButton1 
      Height          =   915
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "VB SuperButton"
      Top             =   3840
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      FontShadowLeft  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   3
      Object.ToolTipText     =   "VB SuperButton"
      CaptionLeft     =   10
      CaptionTop      =   12
      ToolTipShowDelay=   100
      ToolTipFontColor=   16711680
      ToolTipBackColor=   12632256
      Caption         =   "VB"
      FontShadowOffsetX=   2
      FontShadowOffsetY=   2
      FontShadow      =   2
   End
   Begin Project1.SuperButton SuperButton1 
      Height          =   915
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Char Pic's"
      Top             =   2940
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      FontShadowLeft  =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   36
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      EdgeWidth       =   3
      Object.ToolTipText     =   "Char Pic's"
      CaptionLeft     =   8
      CaptionTop      =   4
      ToolTipShowDelay=   100
      Caption         =   "Ù"
      FontShadowOffsetX=   2
      FontShadowOffsetY=   2
      FontShadow      =   2
      EdgeType        =   2
   End
   Begin Project1.SuperButton SuperButton1 
      Height          =   915
      Index           =   2
      Left            =   4980
      TabIndex        =   2
      ToolTipText     =   "Char Pic's"
      Top             =   3840
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      FontShadowLeft  =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   36
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      EdgeWidth       =   3
      Object.ToolTipText     =   "Char Pic's"
      CaptionLeft     =   8
      CaptionTop      =   6
      ToolTipShowDelay=   100
      CaptionCenterX  =   0   'False
      CaptionCenterY  =   0   'False
      Caption         =   "Ø"
      FontShadowOffsetX=   2
      FontShadowOffsetY=   2
      FontShadow      =   2
      EdgeType        =   2
   End
   Begin Project1.SuperButton SuperButton1 
      Height          =   915
      Index           =   3
      Left            =   3180
      TabIndex        =   3
      ToolTipText     =   "Char Pic's"
      Top             =   3840
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      FontShadowRight =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   36
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      EdgeWidth       =   3
      Object.ToolTipText     =   "Char Pic's"
      CaptionLeft     =   5
      CaptionTop      =   6
      ToolTipShowDelay=   100
      CaptionCenterX  =   0   'False
      CaptionCenterY  =   0   'False
      Caption         =   "×"
      FontShadowOffsetX=   2
      FontShadowOffsetY=   2
      FontShadow      =   1
      EdgeType        =   2
   End
   Begin Project1.SuperButton SuperButton1 
      Height          =   915
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Char Pic's"
      Top             =   4740
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      FontShadowRight =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   36
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      EdgeWidth       =   3
      Object.ToolTipText     =   "Char Pic's"
      CaptionLeft     =   0
      CaptionTop      =   4
      ToolTipShowDelay=   100
      CaptionCenterX  =   0   'False
      Caption         =   "Ú"
      FontShadowOffsetX=   2
      FontShadowOffsetY=   2
      FontShadow      =   1
      EdgeType        =   2
   End
   Begin Project1.SuperButton SuperButton4 
      Height          =   555
      Index           =   1
      Left            =   3300
      TabIndex        =   8
      ToolTipText     =   "Icon Pictures"
      Top             =   0
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Icon Pictures"
      PictureDown     =   "Form1.frx":12589
      CaptionLeft     =   18
      CaptionTop      =   18
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":129DB
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   9
      Top             =   0
      Width           =   0
   End
   Begin Project1.SuperButton SuperButton4 
      Height          =   555
      Index           =   2
      Left            =   3960
      TabIndex        =   10
      ToolTipText     =   "Icon Pictures"
      Top             =   0
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Icon Pictures"
      CaptionLeft     =   18
      CaptionTop      =   18
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":12E2D
   End
   Begin Project1.SuperButton SuperButton4 
      Height          =   555
      Index           =   3
      Left            =   4620
      TabIndex        =   11
      ToolTipText     =   "Icon Pictures"
      Top             =   0
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Icon Pictures"
      CaptionLeft     =   18
      CaptionTop      =   18
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":1327F
   End
   Begin Project1.SuperButton SuperButton5 
      Height          =   465
      Index           =   1
      Left            =   6265
      TabIndex        =   13
      ToolTipText     =   "BMP Pic's"
      Top             =   0
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   3
      AutoSize        =   -1  'True
      MaskColor       =   12632256
      Object.ToolTipText     =   "BMP Pic's"
      CaptionLeft     =   13
      CaptionTop      =   13
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":136D1
      EdgeType        =   3
   End
   Begin Project1.SuperButton SuperButton5 
      Height          =   465
      Index           =   2
      Left            =   6720
      TabIndex        =   14
      ToolTipText     =   "BMP Pic's"
      Top             =   0
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   3
      AutoSize        =   -1  'True
      MaskColor       =   12632256
      Object.ToolTipText     =   "BMP Pic's"
      CaptionLeft     =   13
      CaptionTop      =   13
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":137E3
      EdgeType        =   3
   End
   Begin Project1.SuperButton SuperButton5 
      Height          =   465
      Index           =   3
      Left            =   7165
      TabIndex        =   15
      ToolTipText     =   "BMP Pic's"
      Top             =   0
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EdgeWidth       =   3
      AutoSize        =   -1  'True
      MaskColor       =   12632256
      Object.ToolTipText     =   "BMP Pic's"
      CaptionLeft     =   13
      CaptionTop      =   13
      ToolTipShowDelay=   100
      Picture         =   "Form1.frx":138F5
      EdgeType        =   3
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SuperButton3.Caption = "Super" & vbCrLf & "Button" & vbCrLf & "Testing ToolTips"
    'to set ToolTipText programmically- the extender must be bypassed....
    SuperButton3.object.ToolTipText = "Dissected the tooltip stuff from-> CustomToolTips -" _
                                   & "http://www.Planet-Source-Code.com/vb/scripts/ " _
                                   & "ShowCode.asp?txtCodeId=9185&lngWId=1  by: Ark"
    SuperButton3.ToolTipBackColor = vbYellow
    SuperButton3.ToolTipFontColor = vbBlack
    SuperButton3.ToolTipTimeShown = 5000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub

Private Sub SuperButton1_MouseEnter(Index As Integer)
    SuperButton1(Index).FontColor = vbRed
End Sub

Private Sub SuperButton1_MouseExit(Index As Integer)
    SuperButton1(Index).FontColor = vbBlack
End Sub

Private Sub SuperButton11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static i As Long, iX As Long
    If Abs(x - iX) > 10 Then
        i = i + 1
        If i = 17 Then i = 1
        SuperButton11.FontROP = i
        iX = x
    End If
End Sub

Private Sub SuperButton8_Click()
    Unload Me
End Sub

Private Sub SuperButton9_Click()
    Static fOn As Boolean, pic As StdPicture
    If pic Is Nothing Then Set pic = SuperButton9.Picture
    fOn = Not fOn
    If fOn Then
        Set SuperButton9.Picture = SuperButton9.PictureDown
        SuperButton6.Enabled = True
    Else
        Set SuperButton9.Picture = pic
        SuperButton6.Enabled = False
    End If
End Sub

