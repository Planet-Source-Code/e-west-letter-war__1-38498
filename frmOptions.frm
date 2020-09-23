VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   Caption         =   "Options"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   6420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   32
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   31
      Top             =   6480
      Width           =   1455
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4560
      ScaleHeight     =   555
      ScaleWidth      =   855
      TabIndex        =   27
      Top             =   5400
      Width           =   915
      Begin VB.VScrollBar vsHardHigh 
         Height          =   555
         Left            =   600
         Max             =   500
         Min             =   1
         TabIndex        =   28
         Top             =   0
         Value           =   8
         Width           =   255
      End
      Begin VB.Label lblHardHigh 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   855
      TabIndex        =   24
      Top             =   5400
      Width           =   915
      Begin VB.VScrollBar vsHardLow 
         Height          =   555
         Left            =   600
         Max             =   500
         Min             =   1
         TabIndex        =   25
         Top             =   0
         Value           =   8
         Width           =   255
      End
      Begin VB.Label lblHardLow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4560
      ScaleHeight     =   555
      ScaleWidth      =   855
      TabIndex        =   20
      Top             =   4560
      Width           =   915
      Begin VB.VScrollBar vsMedHigh 
         Height          =   555
         Left            =   600
         Max             =   500
         Min             =   1
         TabIndex        =   21
         Top             =   0
         Value           =   8
         Width           =   255
      End
      Begin VB.Label lblMedHigh 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   855
      TabIndex        =   17
      Top             =   4560
      Width           =   915
      Begin VB.VScrollBar vsMedLow 
         Height          =   555
         Left            =   600
         Max             =   500
         Min             =   1
         TabIndex        =   18
         Top             =   0
         Value           =   8
         Width           =   255
      End
      Begin VB.Label lblMedLow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4560
      ScaleHeight     =   555
      ScaleWidth      =   855
      TabIndex        =   10
      Top             =   3720
      Width           =   915
      Begin VB.VScrollBar vsEasyHigh 
         Height          =   555
         Left            =   600
         Max             =   500
         Min             =   1
         TabIndex        =   11
         Top             =   0
         Value           =   8
         Width           =   255
      End
      Begin VB.Label lblEasyHigh 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   855
      TabIndex        =   7
      Top             =   3720
      Width           =   915
      Begin VB.VScrollBar vsEasyLow 
         Height          =   555
         Left            =   600
         Max             =   500
         Min             =   1
         TabIndex        =   8
         Top             =   0
         Value           =   8
         Width           =   255
      End
      Begin VB.Label lblEasyLow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      ScaleHeight     =   675
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   720
      Width           =   1695
      Begin VB.VScrollBar vsRounds 
         Height          =   670
         Left            =   1380
         Max             =   20
         Min             =   1
         TabIndex        =   1
         Top             =   0
         Value           =   8
         Width           =   255
      End
      Begin VB.Label lblRounds 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   30
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3240
      Width           =   1635
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   3240
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Low"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   3240
      Width           =   915
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmOptions.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Falling Letter Rates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Specifies the number of rounds in a single game."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Round Count:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SaveOptions()
   SaveSetting "LetterWar", "FallRates", "EasyLow", FALLING_EASY_LOW
   SaveSetting "LetterWar", "FallRates", "EasyHigh", FALLING_EASY_HIGH
   SaveSetting "LetterWar", "FallRates", "MedLow", FALLING_MED_LOW
   SaveSetting "LetterWar", "FallRates", "MedHigh", FALLING_MED_HIGH
   SaveSetting "LetterWar", "FallRates", "HardLow", FALLING_HARD_LOW
   SaveSetting "LetterWar", "FallRates", "HardHigh", FALLING_HARD_HIGH
   SaveSetting "LetterWar", "Rounds", "Max", ROUND_MAX
   

End Sub


Public Sub SetOptions()
   With Me
      .lblRounds.Caption = CStr(ROUND_MAX)
      .lblEasyHigh.Caption = CStr(FALLING_EASY_HIGH)
      .vsEasyHigh.Value = FALLING_EASY_HIGH
      .lblEasyLow.Caption = CStr(FALLING_EASY_LOW)
      .vsEasyLow.Value = FALLING_EASY_LOW
      .lblHardHigh.Caption = CStr(FALLING_HARD_HIGH)
      .vsHardHigh.Value = FALLING_HARD_HIGH
      .lblHardLow.Caption = CStr(FALLING_HARD_LOW)
      .vsHardLow.Value = FALLING_HARD_LOW
      .lblMedLow.Caption = CStr(FALLING_MED_LOW)
      .vsMedLow.Value = FALLING_MED_LOW
      .lblMedHigh.Caption = CStr(FALLING_MED_HIGH)
      .vsMedHigh.Value = FALLING_MED_HIGH
   End With
End Sub

Private Sub cmdApply_Click()
   FALLING_EASY_LOW = CLng(lblEasyLow.Caption)
   FALLING_EASY_HIGH = CLng(lblEasyHigh.Caption)
   FALLING_MED_LOW = CLng(lblMedLow.Caption)
   FALLING_MED_HIGH = CLng(lblMedHigh.Caption)
   FALLING_HARD_LOW = CLng(lblHardLow.Caption)
   FALLING_HARD_HIGH = CLng(lblHardHigh.Caption)
   ROUND_MAX = CLng(lblRounds.Caption)
   SaveOptions
   cmdApply.Enabled = False
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   SetOptions
   cmdApply.Enabled = False
End Sub

Private Sub VScroll1_Change()
   lblRounds.Caption = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
  
End Sub

Private Sub lblEasyHigh_Change()
   cmdApply.Enabled = True
End Sub

Private Sub lblEasyLow_Change()
   cmdApply.Enabled = True
End Sub

Private Sub lblHardHigh_Change()
   cmdApply.Enabled = True
End Sub

Private Sub lblHardLow_Change()
   cmdApply.Enabled = True
End Sub

Private Sub lblMedHigh_Change()
   cmdApply.Enabled = True
End Sub

Private Sub lblMedLow_Change()
   cmdApply.Enabled = True
End Sub

Private Sub lblRounds_Change()
   cmdApply.Enabled = True
End Sub

Private Sub vsEasyHigh_Change()
   Me.lblEasyHigh.Caption = vsEasyHigh.Value
End Sub

Private Sub vsEasyHigh_Scroll()
   Me.lblEasyHigh.Caption = vsEasyHigh.Value
End Sub


Private Sub vsEasyLow_Change()
   Me.lblEasyLow.Caption = vsEasyLow.Value
End Sub

Private Sub vsEasyLow_Scroll()
   Me.lblEasyLow.Caption = vsEasyLow.Value
End Sub


Private Sub vsHardHigh_Change()
   lblHardHigh.Caption = vsHardHigh.Value
End Sub


Private Sub vsHardHigh_Scroll()
   lblHardHigh.Caption = vsHardHigh.Value
End Sub


Private Sub vsHardLow_Change()
   lblHardLow.Caption = vsHardLow.Value
End Sub


Private Sub vsHardLow_Scroll()
   lblHardLow.Caption = vsHardLow.Value
End Sub


Private Sub vsMedHigh_Change()
   lblMedHigh.Caption = vsMedHigh.Value
End Sub


Private Sub vsMedHigh_Scroll()
   lblMedHigh.Caption = vsMedHigh.Value
End Sub


Private Sub vsMedLow_Change()
   lblMedLow.Caption = vsMedLow.Value
End Sub


Private Sub vsMedLow_Scroll()
    lblMedLow.Caption = vsMedLow.Value
End Sub


Private Sub vsRounds_Change()
   Me.lblRounds.Caption = vsRounds.Value
   
End Sub


Private Sub vsRounds_Scroll()
   lblRounds.Caption = vsRounds.Value
End Sub


