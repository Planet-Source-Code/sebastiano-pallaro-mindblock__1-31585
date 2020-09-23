VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MindBlock"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2040
      Top             =   2160
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "Main.frx":0000
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Index           =   6
      Left            =   2880
      Picture         =   "Main.frx":038A
      Top             =   1140
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Index           =   6
      Left            =   2880
      Picture         =   "Main.frx":0714
      Top             =   540
      Width           =   240
   End
   Begin VB.Image imgCode 
      Height          =   240
      Index           =   6
      Left            =   2880
      Picture         =   "Main.frx":0A9E
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Index           =   5
      Left            =   2580
      Picture         =   "Main.frx":0E28
      Top             =   1140
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Index           =   5
      Left            =   2580
      Picture         =   "Main.frx":11B2
      Top             =   540
      Width           =   240
   End
   Begin VB.Image imgCode 
      Height          =   240
      Index           =   5
      Left            =   2580
      Picture         =   "Main.frx":153C
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Index           =   4
      Left            =   2280
      Picture         =   "Main.frx":18C6
      Top             =   1140
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Index           =   4
      Left            =   2280
      Picture         =   "Main.frx":1C50
      Top             =   540
      Width           =   240
   End
   Begin VB.Image imgCode 
      Height          =   240
      Index           =   4
      Left            =   2280
      Picture         =   "Main.frx":1FDA
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Warn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   1440
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   3060
      Width           =   1815
   End
   Begin VB.Image imgCode 
      Height          =   240
      Index           =   0
      Left            =   1080
      Picture         =   "Main.frx":2364
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Index           =   0
      Left            =   1080
      Picture         =   "Main.frx":26EE
      Top             =   540
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Index           =   0
      Left            =   1080
      Picture         =   "Main.frx":2A78
      Top             =   1140
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Index           =   1
      Left            =   1380
      Picture         =   "Main.frx":2E02
      Top             =   540
      Width           =   240
   End
   Begin VB.Image imgCode 
      Height          =   240
      Index           =   1
      Left            =   1380
      Picture         =   "Main.frx":318C
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgCode 
      Height          =   240
      Index           =   2
      Left            =   1680
      Picture         =   "Main.frx":3516
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgCode 
      Height          =   240
      Index           =   3
      Left            =   1980
      Picture         =   "Main.frx":38A0
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Index           =   2
      Left            =   1680
      Picture         =   "Main.frx":3C2A
      Top             =   540
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Index           =   3
      Left            =   1980
      Picture         =   "Main.frx":3FB4
      Top             =   540
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Index           =   1
      Left            =   1380
      Picture         =   "Main.frx":433E
      Top             =   1140
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Index           =   2
      Left            =   1680
      Picture         =   "Main.frx":46C8
      Top             =   1140
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Index           =   3
      Left            =   1980
      Picture         =   "Main.frx":4A52
      Top             =   1140
      Width           =   240
   End
   Begin VB.Image imgGo 
      Height          =   480
      Left            =   540
      Picture         =   "Main.frx":4DDC
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1 - 2 symbols"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   300
      Width           =   1875
   End
   Begin VB.Image imgHave 
      Height          =   240
      Index           =   5
      Left            =   2880
      Picture         =   "Main.frx":5AA6
      Top             =   2700
      Width           =   240
   End
   Begin VB.Image imgHave 
      Height          =   240
      Index           =   4
      Left            =   2580
      Picture         =   "Main.frx":5E30
      Top             =   2700
      Width           =   240
   End
   Begin VB.Image imgHave 
      Height          =   240
      Index           =   3
      Left            =   2280
      Picture         =   "Main.frx":61BA
      Top             =   2700
      Width           =   240
   End
   Begin VB.Image imgHave 
      Height          =   240
      Index           =   2
      Left            =   1980
      Picture         =   "Main.frx":6544
      Top             =   2700
      Width           =   240
   End
   Begin VB.Image imgHave 
      Height          =   240
      Index           =   1
      Left            =   1680
      Picture         =   "Main.frx":68CE
      Top             =   2700
      Width           =   240
   End
   Begin VB.Image imgHave 
      Height          =   240
      Index           =   0
      Left            =   1380
      Picture         =   "Main.frx":6E58
      Top             =   2700
      Width           =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bytLevel As Byte ' Here is the level
Private bytSels(0 To 6) As Byte ' Here is the code entered
Private bytCode(0 To 6) As Byte
Private bytReflex As Byte
Private bytSecondsLeft As Byte '96 ->60

' This sub calculate a new code combination and store it into
' bytCode() vector.
Private Sub CalculateCode()
    Dim i As Byte
    Dim x As Byte
Redo_Calculate_Code:

    ' I don't know why I put this code here, but I thinks that if I do
    ' randomize for a while well got a better randomizing... ^_^
    Randomize
    x = Rnd * 30
    For i = 1 To x
        Randomize
    Next i
    '---------------
    
    ' Now calculate the random numbers of the symbols on the code.
    ' The number of symbols depend to the bytLevel (level number) plus 1.
    ' If level is 5 we must calculate 7 symbols, too. ^_^
    For i = 0 To IIf(bytLevel = 5, bytLevel + 1, bytLevel)
        bytCode(i) = Rnd * bytLevel
    Next i
    
    ' Now write some infos for debugging (label1 is hidden.
    Label1.Caption = bytCode(0) & " " & bytCode(1) & " " & bytCode(2) & " " & bytCode(3) & " " & bytCode(4) & " " & bytCode(5) & " " & bytCode(6)
End Sub

' This sub set all the variables depending the bytLevel is set.
Private Sub SetLevel()
    Dim strLiv As String
    Dim i As Byte
    
    ' Print some infos.
    strLiv = "Level " & bytLevel & " - " & Str(bytLevel + 1) & " symbols"
    lblLevel.Caption = strLiv
    
    ' Show right box for the level.
    For i = 0 To 6
        If i <= bytLevel Then
            imgCode(i).Visible = True
            Set imgCode(i).Picture = imgHave(0).Picture
            imgUp(i).Visible = True
            imgDown(i).Visible = True
            bytSels(i) = 0
        Else
            imgCode(i).Visible = False
            imgUp(i).Visible = False
            imgDown(i).Visible = False
        End If
    Next i
    
    ' If we are on level 5 just show the last box
    If bytLevel = 5 Then
        imgCode(6).Visible = True
        Set imgCode(6).Picture = imgHave(0).Picture
        imgUp(6).Visible = True
        imgDown(6).Visible = True
        bytSels(6) = 0
    End If
    
    ' Set the picture for each box.
    For i = 0 To 41
        Set Image1(i).Picture = imgHave(bytLevel - 1).Picture
    Next i
    lblWarning.Caption = ""
    ' Let's calculate the new code!
    CalculateCode
End Sub

' This sub sets all the vars for startup a new game.
Private Sub SetStart()
    ' We are on level 1.
    bytLevel = 1
    ' Set all the vars for level 1.
    SetLevel
    ' The first box of the square is 0;
    bytReflex = 0
    ' The seconds we left.
    bytSecondsLeft = 252
    ' One second as interval.
    Timer1.Interval = 1000
    ' Set some other vars that are on this event
    Timer1_Timer
End Sub

' The first Sub that run
Private Sub Form_Load()
    Dim i As Byte
    
    ' TIP : this part of the code demonstrate how to load dinamically
    ' some controls (in our case some image box). We only need a control
    ' with index=0.
    For i = 1 To 7
        Load Image1(i)
        Image1(i).Visible = True
        Image1(i).Left = Image1(0).Left
        Image1(i).Top = Image1(i - 1).Top + Image1(i).Height
    Next i
    For i = 8 To 21
        Load Image1(i)
        Image1(i).Visible = True
        Image1(i).Left = Image1(i - 1).Left + Image1(i).Width
        Image1(i).Top = Image1(7).Top
    Next i
    For i = 22 To 28
        Load Image1(i)
        Image1(i).Visible = True
        Image1(i).Left = Image1(21).Left
        Image1(i).Top = Image1(i - 1).Top - Image1(i).Height
    Next i
    For i = 29 To 41
        Load Image1(i)
        Image1(i).Visible = True
        Image1(i).Left = Image1(i - 1).Left - Image1(i).Height
        Image1(i).Top = Image1(28).Top
    Next i
    Me.Width = 3690
    Me.Height = 2295
    ' set the starting vars for the game.
    SetStart
End Sub

' This event is fire by pressing the down arrow on a box.
Private Sub imgDown_Click(Index As Integer)
    If bytSels(Index) = 0 Then
        bytSels(Index) = bytLevel
    Else
        bytSels(Index) = bytSels(Index) - 1
    End If
    
    Set imgCode(Index).Picture = imgHave(bytSels(Index)).Picture
End Sub

Private Sub imgGo_Click()
    Dim i As Byte
    Dim bytOk As Byte
    
    bytOk = 0
    
    For i = 0 To IIf(bytLevel = 5, bytLevel + 1, bytLevel)
        If bytSels(i) = bytCode(i) Then
            bytOk = bytOk + 1
        End If
    Next i
    
    If bytOk = bytLevel + 1 Then
        bytLevel = bytLevel + 1
        If bytLevel > 5 Then
            MsgBox "Victory!", vbInformation
            bytLevel = 1
        Else
            MsgBox "Passing to level " & bytLevel, vbInformation
            Timer1.Interval = Timer1.Interval - 5
        End If
        SetLevel
    Else
        For i = 0 To bytLevel + 1
            Set imgCode(i).Picture = imgHave(0).Picture
            bytSels(i) = 0
        Next i
        lblWarning.Caption = "Wrong : " & bytOk & " symbols ok."
    End If
End Sub

Private Sub imgUp_Click(Index As Integer)
    bytSels(Index) = bytSels(Index) + 1
    If bytSels(Index) > bytLevel Then
        bytSels(Index) = 0
    End If
    Set imgCode(Index).Picture = imgHave(bytSels(Index)).Picture
End Sub

Private Sub Timer1_Timer()
    If bytSecondsLeft = 0 Then
        MsgBox "Time over.", vbExclamation
        SetStart
        Exit Sub
    End If
    
    Set Image1(bytReflex).Picture = imgHave(bytLevel - 1).Picture
    bytReflex = bytReflex + 1
    If bytReflex > 41 Then bytReflex = 0
    Set Image1(bytReflex).Picture = imgHave(bytLevel).Picture
    
    bytSecondsLeft = bytSecondsLeft - 1
    Me.Caption = "MindBlock (" & bytSecondsLeft & " seconds left)"
End Sub
