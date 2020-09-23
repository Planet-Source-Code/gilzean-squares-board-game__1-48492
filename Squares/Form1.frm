VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Squares"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List4 
      Height          =   645
      Left            =   2640
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   390
      Left            =   480
      Picture         =   "Form1.frx":01CA
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ListBox List3 
      Height          =   645
      Left            =   3120
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Top Scores "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   5040
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
      Begin VB.Label lblHighScore 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1400
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1400
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1400
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1400
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1400
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton Level 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level 6"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Tag             =   "3"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Level 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level 5"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Tag             =   "8"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Level 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level 4"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Tag             =   "7"
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Level 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level 3"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Tag             =   "6"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Level 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level 2"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Tag             =   "5"
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Level 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level 1"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Tag             =   "4"
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Level 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level 0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "3"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   600
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score = 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Shape Border 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   4800
      Left            =   0
      Top             =   0
      Width           =   4800
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Counter As Integer
Dim Squares_Remaining As Integer
Dim Points As Integer
Dim Game_Level As Integer
Dim Hiscore As Boolean
Dim Prev_User As String
Private Sub Form_Load()
    Randomize
    Create_Grid
    Game_Level = 1
    Prev_User = ""
    Show_Hiscores
    Initialise
End Sub
Private Sub Initialise()
    Randomize_Colours
    Colour_Squares
    Squares_Remaining = 144
    Points = 0
    Label1.Caption = "Points = 0"
    Hiscore = False
    Show_Hiscores
End Sub
Private Sub Create_Grid()
Dim SquareSize As Integer
Dim XPos As Integer
Dim YPos As Integer
Dim BorderWidth As Integer
    SquareSize = 400
    BorderWidth = 30
    XPos = BorderWidth
    YPos = -SquareSize + BorderWidth
    Square(0).Height = SquareSize
    Square(0).Width = SquareSize
    Border.Width = (SquareSize * 12) + BorderWidth
    Border.Height = (SquareSize * 12) + BorderWidth
    For Counter = 1 To 144
        If (Counter Mod 12 = 1) Then
            YPos = YPos + SquareSize
            XPos = BorderWidth
        End If
        Load Square(Counter)
        Square(Counter).Left = XPos
        Square(Counter).Top = YPos
        Square(Counter).Visible = True
        XPos = XPos + SquareSize
    Next
End Sub
Private Sub Randomize_Colours()
Dim x As Integer
    List1.Clear
    For Counter = 1 To 144
        x = (Rnd * 799) + 100
        List1.AddItem x & (Counter Mod (Game_Level + 2))
    Next Counter
End Sub

Private Sub Level_Click(Index As Integer)
    Label2.Caption = "Level " & Index
    Game_Level = Index
    Show_Hiscores
    Initialise
End Sub

Private Sub Square_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Keep_Looking As Boolean
Dim Current As Integer
    If Square(Index).BackColor = &HFFFFFF Then
        Exit Sub
    End If
    List1.Clear
    List1.AddItem Index + 1000
    'up
    Keep_Looking = True
    Current = Index
    Square(Current).BorderStyle = 0
    While ((Current - 12) > 0) And Keep_Looking = True
        DoEvents
        If (Square(Current - 12).BackColor = Square(Index).BackColor) Then
            Square(Current - 12).BorderStyle = 0
            Current = Current - 12
            List1.AddItem Current + 1000
        Else
            Keep_Looking = False
        End If
    Wend
    'down
    Keep_Looking = True
    Current = Index
    While ((Current + 12) < 145) And Keep_Looking = True
        DoEvents
        If (Square(Current + 12).BackColor = Square(Index).BackColor) Then
            Square(Current + 12).BorderStyle = 0
            Current = Current + 12
            List1.AddItem Current + 1000
        Else
            Keep_Looking = False
        End If
    Wend
    'left
    Keep_Looking = True
    Current = Index
    While ((Current - 1) Mod 12) > 0 And Keep_Looking = True
        DoEvents
        If (Square(Current - 1).BackColor = Square(Index).BackColor) Then
            Square(Current - 1).BorderStyle = 0
            Current = Current - 1
            List1.AddItem Current + 1000
        Else
            Keep_Looking = False
        End If
    Wend
    'right
    Keep_Looking = True
    Current = Index
    While ((Current + 1) Mod 12) <> 1 And Keep_Looking = True
        DoEvents
        If (Square(Current + 1).BackColor = Square(Index).BackColor) Then
            Square(Current + 1).BorderStyle = 0
            Current = Current + 1
            List1.AddItem Current + 1000
        Else
            Keep_Looking = False
        End If
    Wend
    If Button = 1 And List1.ListCount > 1 Then
        Remove_Squares
    End If
End Sub

Private Sub Square_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    For Counter = 1 To 144
        Square(Counter).BorderStyle = 1
    Next
End Sub
Private Sub Remove_Squares()
Dim Dead_Square As Integer
Dim Column_Move As Boolean
    Column_Move = False
    List2.Clear
    For Counter = List1.ListCount - 1 To 0 Step -1
        List2.AddItem (List1.List(Counter) - 1000)
    Next Counter
    For Counter = 0 To List2.ListCount - 1
        If List2.List(Counter) > 132 Then
            Column_Move = True
        End If
    Next Counter
    Points = Points + (3 * (List2.ListCount - 1) - 2 + Game_Level)
    On Error Resume Next
    For Counter = (List2.ListCount - 1) To 0 Step -1
        Dead_Square = List2.List(Counter)
        While Dead_Square > 12
            Square(Dead_Square).BackColor = Square(Dead_Square - 12).BackColor
            Dead_Square = Dead_Square - 12
        Wend
        Square(Dead_Square).BackColor = &HFFFFFF
        Squares_Remaining = Squares_Remaining - 1
        Label1.Caption = "Points = " & Points
    Next Counter
    If Squares_Remaining = 0 Then
        Points = Points + 30
        Label1.Caption = "Points = " & Points
        Call Check_Hiscore
    End If
    If Column_Move = True Then
        Move_Columns
    End If
End Sub
Private Sub Move_Columns()
Dim Counter1 As Integer
Dim Counter2 As Integer
    For Counter2 = 1 To 10
        For Counter = 133 To 143
            If Square(Counter).BackColor = &HFFFFFF Then
                For Counter1 = 0 To 11
                    Square(Counter - (Counter1 * 12)).BackColor = Square(Counter - (Counter1 * 12) + 1).BackColor
                    Square(Counter - (Counter1 * 12) + 1).BackColor = &HFFFFFF
                Next Counter1
            End If
        Next Counter
    Next Counter2
End Sub

Private Sub Colour_Squares()
    For Counter = 1 To 144
        Square(Counter).BorderStyle = 1
        Select Case (Right(List1.List(Counter - 1), 1))
            Case 0
                Square(Counter).BackColor = &HFF&
            Case 1
                Square(Counter).BackColor = &H80FF&
            Case 2
                Square(Counter).BackColor = &HFFFF&
            Case 3
                Square(Counter).BackColor = &H80FF80
            Case 4
                Square(Counter).BackColor = &H8000&
            Case 5
                Square(Counter).BackColor = &HFF0000
            Case 6
                Square(Counter).BackColor = &HFFFF80
            Case 7
                Square(Counter).BackColor = &HFF00FF
        End Select
    Next
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblPlay.BackColor = &H8000&
    lblClose.BackColor = &H8000&
    lblQuit.BackColor = &H8000&
End Sub

Private Sub lblQuit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblQuit.BackColor = &H80FF80
End Sub

Private Sub lblQuit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblQuit.BackColor = &HFF00&
End Sub

Private Sub lblQuit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblQuit.BackColor = &H8000&
End Sub
Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblClose.BackColor = &H80FF80
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblClose.BackColor = &HFF00&
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblClose.BackColor = &H8000&
End Sub


Private Sub lblPlay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblPlay.BackColor = &H80FF80
End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblPlay.BackColor = &HFF00&
End Sub

Private Sub lblPlay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblPlay.BackColor = &H8000&
End Sub
Private Sub lblClose_Click()
    For Counter = 1 To 144
        Unload Square(Counter)
    Next
    Unload Me
    End
End Sub
Private Sub lblPlay_Click()
    Call Initialise
End Sub
Private Sub lblQuit_Click()
    Check_Hiscore
End Sub
Private Sub Check_Hiscore()
Dim Temp As String
Dim Message As String
    If Squares_Remaining = 0 Then
        Message = "You Completed Level " & Game_Level & vbCrLf & " With a score of " & Points
    Else
        Message = "You Quit Level " & Game_Level & vbCrLf & " With a score of " & Points
    End If
    If Points > Val(List3.List(4)) Then
        Hiscore = True
        Message = Message & vbCrLf & "Congratulations - a HIGH SCORE"
    End If
    MsgBox Message, vbOKOnly + vbInformation, "Squares"
    If Hiscore = True Then
        Update_Score
    End If
    For Counter = 1 To 144
           Square(Counter).BackColor = &HFFFFFF
    Next Counter
    Call lblPlay_Click
End Sub
Private Sub Show_Hiscores()
Dim Hi_Score As String
Dim File_Name As String
    List3.Clear
    File_Name = App.Path & "\Hiscores" & Game_Level & ".txt"
    Open File_Name For Input As #1
    For Counter = 0 To 4
        Line Input #1, Hi_Score
        List3.AddItem (Val(Left$(Hi_Score, 3)))
        lblHighScore(Counter).Caption = Hi_Score
    Next
    Close #1
End Sub
Private Sub Update_Score()
Dim TempName As String
Dim Temp As String
Dim File_Name As String
Dim Position As Integer
    For Counter = 0 To 4
        If Points > Val(List3.List(Counter)) Then
            Position = Counter
            Counter = 4
        End If
    Next Counter
    File_Name = App.Path & "\Hiscores" & Game_Level & ".txt"
    TempName = InputBox("Enter your Name", "Squares", Prev_User)
    Prev_User = TempName
    Open File_Name For Input As #1
    Open App.Path & "\Hiscorestemp.txt" For Output As #2
    For Counter = 0 To 4
        If Counter = Position Then
            Print #2, Points & "   " & TempName
        Else
            Line Input #1, Temp
            Print #2, Temp
        End If
    Next Counter
    Close #1
    Close #2
    Kill File_Name
    Name App.Path & "\Hiscorestemp.txt" As File_Name
    Show_Hiscores
End Sub
