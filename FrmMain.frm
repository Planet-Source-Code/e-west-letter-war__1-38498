VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Letter War"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   9195
      TabIndex        =   38
      Top             =   6480
      Width           =   9255
      Begin VB.Frame Frame1 
         Caption         =   "[ Options ]"
         Height          =   975
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Width           =   3375
         Begin VB.OptionButton optDifficulty 
            Caption         =   "Easy"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optDifficulty 
            Caption         =   "Medium"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   44
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDifficulty 
            Caption         =   "Hard"
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdAdvanced 
            Caption         =   "Advanced"
            Height          =   255
            Left            =   1080
            TabIndex        =   42
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   40
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   39
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Score:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6720
         TabIndex        =   46
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Timer timerScroll 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   4560
   End
   Begin VB.Timer TimerRound 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   4560
   End
   Begin VB.Timer TimerScore 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4320
      Top             =   4560
   End
   Begin VB.Timer timerLetters 
      Interval        =   100
      Left            =   4920
      Top             =   4560
   End
   Begin VB.Label lblPress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   37
      Top             =   5760
      Width           =   9135
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   17
      Left            =   8640
      TabIndex        =   36
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   16
      Left            =   8160
      TabIndex        =   35
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   15
      Left            =   7680
      TabIndex        =   34
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   14
      Left            =   7200
      TabIndex        =   33
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   13
      Left            =   6720
      TabIndex        =   32
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   12
      Left            =   6240
      TabIndex        =   31
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   11
      Left            =   5760
      TabIndex        =   30
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   29
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   9
      Left            =   4800
      TabIndex        =   28
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   27
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   26
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   25
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   24
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   23
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   22
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   21
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   20
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   19
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hit Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   8895
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   17
      Left            =   8640
      Picture         =   "FrmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   17
      Left            =   8640
      Picture         =   "FrmMain.frx":030A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   17
      Left            =   8640
      TabIndex        =   17
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   16
      Left            =   8160
      Picture         =   "FrmMain.frx":0614
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   16
      Left            =   8160
      Picture         =   "FrmMain.frx":091E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   16
      Left            =   8160
      TabIndex        =   16
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   15
      Left            =   7680
      Picture         =   "FrmMain.frx":0C28
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   15
      Left            =   7680
      Picture         =   "FrmMain.frx":0F32
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   15
      Left            =   7680
      TabIndex        =   15
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   14
      Left            =   7200
      Picture         =   "FrmMain.frx":123C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   14
      Left            =   7200
      Picture         =   "FrmMain.frx":1546
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   14
      Left            =   7200
      TabIndex        =   14
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   13
      Left            =   6720
      Picture         =   "FrmMain.frx":1850
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   13
      Left            =   6720
      Picture         =   "FrmMain.frx":1B5A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   13
      Left            =   6720
      TabIndex        =   13
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   12
      Left            =   6240
      Picture         =   "FrmMain.frx":1E64
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   12
      Left            =   6240
      Picture         =   "FrmMain.frx":216E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   12
      Left            =   6240
      TabIndex        =   12
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   11
      Left            =   5760
      Picture         =   "FrmMain.frx":2478
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   11
      Left            =   5760
      Picture         =   "FrmMain.frx":2782
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   11
      Left            =   5760
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   10
      Left            =   5280
      Picture         =   "FrmMain.frx":2A8C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   10
      Left            =   5280
      Picture         =   "FrmMain.frx":2D96
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   10
      Left            =   5280
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   9
      Left            =   4800
      Picture         =   "FrmMain.frx":30A0
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   9
      Left            =   4800
      Picture         =   "FrmMain.frx":33AA
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   9
      Left            =   4800
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   8
      Left            =   4320
      Picture         =   "FrmMain.frx":36B4
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   8
      Left            =   4320
      Picture         =   "FrmMain.frx":39BE
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   8
      Left            =   4320
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   7
      Left            =   3840
      Picture         =   "FrmMain.frx":3CC8
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   7
      Left            =   3840
      Picture         =   "FrmMain.frx":3FD2
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   7
      Left            =   3840
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   6
      Left            =   3360
      Picture         =   "FrmMain.frx":42DC
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   6
      Left            =   3360
      Picture         =   "FrmMain.frx":45E6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   6
      Left            =   3360
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   5
      Left            =   2880
      Picture         =   "FrmMain.frx":48F0
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   5
      Left            =   2880
      Picture         =   "FrmMain.frx":4BFA
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   5
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   4
      Left            =   2400
      Picture         =   "FrmMain.frx":4F04
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   4
      Left            =   2400
      Picture         =   "FrmMain.frx":520E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   3
      Left            =   1920
      Picture         =   "FrmMain.frx":5518
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   3
      Left            =   1920
      Picture         =   "FrmMain.frx":5822
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   2
      Left            =   1440
      Picture         =   "FrmMain.frx":5B2C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   2
      Left            =   1440
      Picture         =   "FrmMain.frx":5E36
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   1
      Left            =   960
      Picture         =   "FrmMain.frx":6140
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   1
      Left            =   960
      Picture         =   "FrmMain.frx":644A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   49
      X1              =   0
      X2              =   360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   48
      X1              =   0
      X2              =   360
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   47
      X1              =   0
      X2              =   360
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   46
      X1              =   0
      X2              =   360
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   45
      X1              =   0
      X2              =   360
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   44
      X1              =   0
      X2              =   360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   43
      X1              =   0
      X2              =   360
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   42
      X1              =   0
      X2              =   360
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   41
      X1              =   0
      X2              =   360
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   40
      X1              =   0
      X2              =   360
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   39
      X1              =   0
      X2              =   360
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   38
      X1              =   0
      X2              =   360
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   37
      X1              =   0
      X2              =   360
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   36
      X1              =   0
      X2              =   360
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Image Star 
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "FrmMain.frx":6754
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9240
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Explode 
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "FrmMain.frx":6A5E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   35
      X1              =   0
      X2              =   360
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   34
      X1              =   0
      X2              =   360
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   33
      X1              =   0
      X2              =   360
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   32
      X1              =   0
      X2              =   360
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   31
      X1              =   0
      X2              =   360
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   30
      X1              =   0
      X2              =   360
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   29
      X1              =   0
      X2              =   360
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   28
      X1              =   0
      X2              =   360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   27
      X1              =   0
      X2              =   360
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   26
      X1              =   0
      X2              =   360
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   25
      X1              =   0
      X2              =   360
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   24
      X1              =   0
      X2              =   360
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   23
      X1              =   0
      X2              =   360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   22
      X1              =   0
      X2              =   360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   21
      X1              =   0
      X2              =   360
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   20
      X1              =   0
      X2              =   360
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   19
      X1              =   0
      X2              =   360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   18
      X1              =   0
      X2              =   360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   17
      X1              =   0
      X2              =   360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   16
      X1              =   0
      X2              =   360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   15
      X1              =   0
      X2              =   360
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   14
      X1              =   0
      X2              =   360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   13
      X1              =   0
      X2              =   360
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   12
      X1              =   0
      X2              =   360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   11
      X1              =   0
      X2              =   360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   10
      X1              =   0
      X2              =   360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   9
      X1              =   0
      X2              =   360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   8
      X1              =   0
      X2              =   360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   7
      X1              =   0
      X2              =   360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   6
      X1              =   0
      X2              =   360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   5
      X1              =   0
      X2              =   360
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   4
      X1              =   0
      X2              =   360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   3
      X1              =   0
      X2              =   360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   2
      X1              =   0
      X2              =   360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   1
      X1              =   0
      X2              =   360
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   0
      X2              =   360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   6375
      Left            =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Public lRound As Long
Public RoundMax As Long

Public InGame As Boolean

Public CurrentDifficulty As LetterWarDifficulty
Public FlawlessBonus As Boolean

Private Sub cmdAdvanced_Click()
   frmOptions.Show vbModal
End Sub

Private Sub cmdExit_Click()
   End
End Sub

Private Sub cmdStart_Click()
      Dim Difficulty As LetterWarDifficulty
      
      If InGame Then
         InGame = False
         HideLetters
         lblMessage.Visible = True
         cmdStart.Caption = "Start"
      Else
         GetOptions
         ResetCounters
         InGame = True
         cmdStart.Caption = "Quit"
         cmdStart.Default = False
         If optDifficulty(0).Value = True Then
            Difficulty = LW_EASY
         ElseIf optDifficulty(1).Value = True Then
            Difficulty = LW_MEDIUM
         Else
            Difficulty = LW_HARD
         End If
         CurrentDifficulty = Difficulty
         StartGame Difficulty
      End If
      
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If InGame Then
      If KeyCode <> 16 Then
         lblPress.Caption = Chr(KeyCode)
      End If
   End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   CheckHit Chr(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   lblPress.Caption = ""
End Sub


Private Sub Form_Load()
   GetOptions
   timerLetters.Enabled = True
   HideLetters
   lblMessage.Visible = True
  
End Sub

Private Sub timerLetters_Timer()

   Dim i As Long
   
   For i = 0 To lblLetter.UBound
      If lblLetter(i).Visible = True Then
         lblLetter(i).Top = lblLetter(i).Top + CInt(lblLetter(i).Tag)
      End If
   Next
   
   CheckState
End Sub







Private Sub TimerRound_Timer()
   Dim i As Long
   Static TStep As Integer
   
   TStep = TStep + 1
   
   If TStep = 1 Then
      For i = 0 To Score.UBound
         Score(i).Visible = False
      Next
      If lRound <= ROUND_MAX - 1 Then
         lblMessage.Caption = "Round " & lRound & " complete."
         lblMessage.Visible = True
         lRound = lRound + 1
      Else
         lblMessage.Caption = "Game Over" & vbNewLine & "Score: " & lblTotal.Caption
         lblMessage.Visible = True
         cmdStart.Caption = "Start"
         TimerRound.Enabled = False
         FrmMain.InGame = False
      End If
   ElseIf TStep = 2 Then
      lblMessage.Caption = "Get ready for Round " & lRound
      lblMessage.Visible = True
   Else
      TStep = 0
      lblMessage.Visible = False
      lblMessage.Caption = "Hit Start"
      TimerRound.Enabled = False
      StartGame CurrentDifficulty
   End If
          
      
   
End Sub

Private Sub TimerScore_Timer()
   Dim i As Integer
   Dim lScore As Long
   Dim sScore As String
   
   With Me
      For i = 0 To .Star.UBound
         If .Star(i).Visible = True Then
            'This is a positive score.
            lScore = 5250 - .Star(i).Top
            .Score(i).Top = .Star(i).Top
            .Star(i).Visible = False
            .Score(i).Caption = CStr(lScore)
            .Score(i).Visible = True
            SendSound ScoreCheck
            Exit Sub
         End If
      Next
      
      'Got here, Tally explodes
      For i = 0 To .Explode.UBound
         If .Explode(i).Visible = True Then
            .Score(i).Top = .Explode(i).Top
            .Score(i).Caption = "-100"
            .Explode(i).Visible = False
            .Score(i).Visible = True
            SendSound ScoreCheck
            Exit Sub
         End If
      Next
      
      'Got here tally scores,
      For i = 0 To .Score.UBound
         If .Score(i).Visible = True Then
            If InStr(1, .Score(i).Caption, "-") > 0 Then
               lScore = lScore - 100
            Else
               lScore = lScore + CLng(.Score(i).Caption)
            End If
         End If
      Next
      If FlawlessBonus Then
         lblMessage.Caption = "Flawless Bonus " & vbNewLine & "+10,000 Points!"
         lblMessage.Visible = True
         lScore = lScore + 10000
      Else
         FlawlessBonus = True
      End If
      lblTotal.Caption = CLng(lblTotal.Caption) + lScore
      TimerScore.Enabled = False
      TimerRound.Enabled = True
    
  End With
End Sub


Private Sub timerScroll_Timer()

End Sub


