VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Based on My Child Memory Games ""True Love"""
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer MagicTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   1320
   End
   Begin VB.CommandButton Calculate 
      BackColor       =   &H0000FF00&
      Caption         =   "&CALCULATE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox LoverName 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox YourName 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label MagicWord 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed By : Teddy Siswoyo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT YOUR LOVER NAME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3930
      TabIndex        =   2
      Top             =   2040
      Width           =   2565
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT YOUR NAME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   0
      Picture         =   "Menu.frx":0CCA
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   6615
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed By : Teddy Siswoyo
'Just a fun little game
'This Game is based on my child memory game
'Enjoy It :) you can believe it or not just enjoy

Option Base 1
Dim i As Integer
Dim j As Integer
Dim nStep As Integer
Dim LoveResult As String
Dim ChatWord As String
Dim WordStep As Integer
Dim nSecond As Integer
Private Sub MagicTimer_Timer()
    Select Case nStep
    Case 0: WordStep = 0: nSecond = 0
        ChatWord = "Uhuk Uhukk..Be Patience, Uncle Teddy is calculating your future with he or she": nStep = 1
    Case 1: WordStep = WordStep + 1 'This step is to make the word animation
            If WordStep >= Len(ChatWord) Then nStep = 2: nSecond = 0 'Update the timer and go to the next step
    Case 2: nSecond = nSecond + 1: If nSecond > 2 Then nStep = 3: ChatWord = "1": MagicTimer.Interval = 10: nSecond = 0
    Case 3
         ChatWord = Val(ChatWord) + 1
         If ChatWord >= 100 Then nStep = 4: nSecond = 0: ChatWord = "IT'S DONE !!"
    Case 4
         nSecond = nSecond + 1
         If nSecond > 2 Then nStep = 5: ChatWord = "Your Possible Match with he or she is " & vbCrLf & LoveResult & " %": WordStep = Len(ChatWord): Beep
    Case 5: Calculate.Enabled = True: YourName.Locked = False: LoverName.Locked = False
         nStep = 0: MagicTimer.Interval = 100: MagicTimer.Enabled = False
    End Select
    MagicWord.Caption = Mid(ChatWord, 1, WordStep)
End Sub
Private Sub Calculate_Click() 'This procedure is used to count the possible match
Dim WordType(8) As Integer
Dim CoupleName As String
Dim Amount As Integer
    If Trim(YourName) = "" Or Trim(LoverName) = "" Then
       MsgBox "Hey Come On, Don't You two have a name ? ", vbExclamation, "Hey.."
       Exit Sub
    End If
    Calculate.Enabled = False: YourName.Locked = True
    MagicTimer.Enabled = True: LoverName.Locked = True
    Amount = 8
    CoupleName = YourName.Text + LoverName.Text 'Combine this couple name first
    DelSpace CoupleName 'Delete the empty space
    For j = 1 To UBound(WordType) 'check the couple name if they have a truelove words in it or not
        For i = 1 To Len(CoupleName)
            If Mid(CoupleName, i, 1) = GetTextName(",T,R,U,E,L,O,V,E,", j) Or Mid(CoupleName, i, 1) = GetTextName(",t,r,u,e,l,o,v,e,", j) Then
               WordType(j) = WordType(j) + 1 'Count the amount of T,R,U,E,L,O,V,E word
            End If
        Next i
    Next j
    
    Do While Amount > 2
       For i = 1 To Amount - 1
           WordType(i) = WordType(i) + WordType(i + 1)
           If WordType(i) > 9 Then WordType(i) = WordType(i) - 10
       Next i
       Amount = Amount - 1
    Loop
    
    LoveResult = Str(WordType(1)) + Trim(Str(WordType(2)))
    If Val(LoveResult) < 10 Then LoveResult = Trim(Mid(LoveResult, 2, 1)) 'If loveresult below ten, then take the second digit only
End Sub
Private Sub LoverName_GotFocus() 'Selection
    LoverName.SelStart = 0: LoverName.SelLength = Len(LoverName.Text)
End Sub
Private Sub YourName_GotFocus() 'Selection to the text box
    YourName.SelStart = 0: YourName.SelLength = Len(YourName.Text)
End Sub

Private Sub YourName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       Calculate_Click
    End If
End Sub
Private Sub LoverName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Calculate_Click
End Sub


