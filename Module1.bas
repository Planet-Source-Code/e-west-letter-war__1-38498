Attribute VB_Name = "Module1"

Option Explicit

Public Enum LetterWarDifficulty
   LW_EASY = 100
   LW_MEDIUM = 50
   LW_HARD = 25
End Enum

Public FALLING_EASY_LOW As Long
Public FALLING_EASY_HIGH As Long
Public FALLING_MED_LOW As Long
Public FALLING_MED_HIGH As Long
Public FALLING_HARD_LOW As Long
Public FALLING_HARD_HIGH As Long
Public ROUND_MAX As Long
Public gHighScores As String

'For sounds:
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Enum WAVEPLAY
   ScoreCheck = 1
   LetterFire = 2
   BottomHit = 3
   LetterMiss = 4
   Cheer = 5
End Enum
Public Sub CheckHit(Letter As String)
   Dim i As Integer
   Dim VisibleCount As Long
   Dim NoHit As Boolean
   
   NoHit = True
   
   With FrmMain
      For i = 0 To .lblLetter.UBound
         If .lblLetter(i).Visible = True Then
            VisibleCount = VisibleCount + 1
            If .lblLetter(i).Caption = Letter Then
               'Success!
               SendSound LetterFire
               NoHit = False
               VisibleCount = VisibleCount - 1
               .Star(i).Top = .lblLetter(i).Top
               .Star(i).Visible = True
               .lblLetter(i).Visible = False
               .lblLetter(i).Top = 0
            End If
         End If
      Next
   End With
   
   If NoHit = True Then
      If Letter <> Chr(34) Then
         FrmMain.FlawlessBonus = False
         SendSound LetterMiss
      End If
   End If
   
   If VisibleCount = 0 Then
      'End of round:
      FrmMain.TimerScore.Enabled = True
   End If
End Sub

Public Sub CheckState()
   Dim i As Long
   
   With FrmMain
      For i = 0 To .lblLetter.UBound
         DoEvents
         If .lblLetter(i).Visible = True Then
            If .lblLetter(i).Top > 5250 Then
               .Explode(i).Top = .lblLetter(i).Top
               .Explode(i).Visible = True
               .lblLetter(i).Visible = False
               FrmMain.FlawlessBonus = False
               SendSound BottomHit
               CheckHit Chr(34) 'send a dummy value in
            End If
         End If
      Next
   End With
End Sub

Public Sub GetOptions()
 
   FALLING_EASY_LOW = CLng(GetSetting("LetterWar", "FallRates", "EasyLow", "5"))
   FALLING_EASY_HIGH = CLng(GetSetting("LetterWar", "FallRates", "EasyHigh", "10"))
   FALLING_MED_LOW = CLng(GetSetting("LetterWar", "FallRates", "MedLow", "15"))
   FALLING_MED_HIGH = CLng(GetSetting("LetterWar", "FallRates", "MedHigh", "30"))
   FALLING_HARD_LOW = CLng(GetSetting("LetterWar", "FallRates", "HardLow", "35"))
   FALLING_HARD_HIGH = CLng(GetSetting("LetterWar", "FallRates", "HardHigh", "50"))
   ROUND_MAX = CLng(GetSetting("LetterWar", "Rounds", "Max", "8"))
   
End Sub

Public Function GetRandomNumber(LowBound As Long, UpperBound As Long) As Long
   Dim lRes As Long
   
   Randomize Time
   
   lRes = Int((UpperBound - LowBound + 1) * Rnd + LowBound)

   GetRandomNumber = lRes
   
End Function


Public Sub HideLetters()
Dim i As Integer
   With FrmMain
      For i = 0 To .lblLetter.UBound
         .lblMessage.Visible = False
         .lblLetter(i).Visible = False
         .lblLetter(i).Alignment = 2
         .lblLetter(i).BackStyle = 1
         .lblLetter(i).BackColor = vbBlack
         .lblLetter(i).BorderStyle = 0
         .lblLetter(i).Top = 0
         .lblLetter(i).Tag = ""
         .Explode(i).Visible = False
         .Star(i).Visible = False
         .Score(i).Visible = False
      Next
   End With
End Sub

Public Sub RandomizeVisibleLetters(Difficulty As LetterWarDifficulty)
   Dim lLetterCount As Long
   Dim bContinue As Boolean
   Dim lRand As Long
   Dim MinLetters As Long
   Dim MaxLetters As Long
   Dim i As Long
   
   If Difficulty = LW_EASY Then
      MinLetters = 1
      MaxLetters = 5
   ElseIf Difficulty = LW_MEDIUM Then
      MinLetters = 5
      MaxLetters = 8
   Else
      MinLetters = 8
      MaxLetters = FrmMain.lblLetter.UBound
   End If
   
   With FrmMain
   
      lLetterCount = GetRandomNumber(MinLetters, MaxLetters)
      
      'Make random labels visible:
      For i = 0 To lLetterCount
         bContinue = False
         Do
            lRand = GetRandomNumber(.lblLetter.LBound, .lblLetter.UBound)
            If .lblLetter(lRand).Visible = False Then
               .lblLetter(lRand).Visible = True
               'Falling rate will be shown on the labels tag:
               'Rate Easy: 5 to 10
               'Rate Medium: 15 to 30
               'Rate Hard: 35 to 55
               If Difficulty = LW_EASY Then
                  .lblLetter(lRand).Tag = GetRandomNumber(FALLING_EASY_LOW, FALLING_EASY_HIGH)
               ElseIf Difficulty = LW_MEDIUM Then
                  .lblLetter(lRand).Tag = GetRandomNumber(FALLING_MED_LOW, FALLING_MED_HIGH)
               Else
                  .lblLetter(lRand).Tag = GetRandomNumber(FALLING_HARD_LOW, FALLING_HARD_HIGH)
               End If
               bContinue = False
            Else
               bContinue = True
            End If
         Loop While bContinue
      Next
      
   End With
   
End Sub

Public Sub ResetCounters()
   FrmMain.lRound = 1
   FrmMain.FlawlessBonus = True
   FrmMain.lblTotal.Caption = "0"
End Sub

Public Sub SendSound(wPlay As WAVEPLAY)
   
   If wPlay = ScoreCheck Then
      PlaySound App.Path & "\blip.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
   ElseIf wPlay = LetterFire Then
      PlaySound App.Path & "\LaserFire.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
   ElseIf wPlay = BottomHit Then
      PlaySound App.Path & "\Explode.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
   ElseIf wPlay = LetterMiss Then
      PlaySound App.Path & "\Miss.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
   End If
End Sub

Public Sub SetLabelChars()
   Dim i As Long
   Dim iRand As Long
   
   With FrmMain
    
      'Characters:
      '(0): Upper Case = 65 to 90
      '(1): Lower Case = 97 to 122
      
      'Numbers:
      '(2): 0 - 9 = 48 to 57
      
      For i = .lblLetter.LBound To .lblLetter.UBound
         If .lblLetter(i).Visible = True Then
            iRand = GetRandomNumber(0, 2)
            Select Case iRand
               Case 0 'Upper case character set
                  .lblLetter(i).Caption = Chr(GetRandomNumber(65, 90))
               Case 1 'Lower case character set
                  .lblLetter(i).Caption = Chr(GetRandomNumber(97, 122))
               Case 2 'Numneric character set
                  .lblLetter(i).Caption = Chr(GetRandomNumber(48, 57))
            End Select
         End If
      Next
      
   End With
   
End Sub

Public Sub SetTimers(LetterRate As Long)
   With FrmMain
      .timerLetters.Interval = LetterRate
      .timerLetters.Enabled = True
   End With
End Sub

Public Sub StartGame(Difficulty As LetterWarDifficulty)
   
   With FrmMain
      .KeyPreview = True
      HideLetters
      RandomizeVisibleLetters Difficulty
      SetLabelChars
      SetTimers Difficulty
   End With
   
   
End Sub
Public Sub StartHighScoresList()
      
   Dim vScores
   Dim sRichText As String
   
   vScores = Split(gHighScores, "|")
   
   With FrmMain.rtfScores
      .Text = "High Scores"
      .
End Sub


