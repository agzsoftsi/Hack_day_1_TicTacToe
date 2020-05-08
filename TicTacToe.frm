VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Tic Tac Toe - Hack Day"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   7455
      Left            =   0
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tic Tac Toe"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   6420
      Left            =   4800
      Picture         =   "TicTacToe.frx":0000
      Top             =   840
      Width           =   3405
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   8
      Left            =   3360
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   7
      Left            =   1920
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   5
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label T 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Player As Byte

Private Sub Command1_Click()
T(0).BackColor = vbWhite
T(1).BackColor = vbWhite
T(2).BackColor = vbWhite
T(3).BackColor = vbWhite
T(4).BackColor = vbWhite
T(5).BackColor = vbWhite
T(6).BackColor = vbWhite
T(7).BackColor = vbWhite
T(8).BackColor = vbWhite
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Player = 0
End Sub

Private Sub T_Click(Index As Integer)
' Definimos la
If Player = 0 Then
T(Index).BackColor = vbBlack
Player = 1
Else
T(Index).BackColor = vbYellow

Player = 0
End If


If T(0).BackColor = vbBlack And T(1).BackColor = vbBlack And T(2).BackColor = vbBlack Then
MsgBox ("Player Black: WIN")
End If

If T(3).BackColor = vbBlack And T(4).BackColor = vbBlack And T(5).BackColor = vbBlack Then
MsgBox ("Player Black: WIN")
End If

If T(6).BackColor = vbBlack And T(7).BackColor = vbBlack And T(8).BackColor = vbBlack Then
MsgBox ("Player Black: WIN")
End If

If T(0).BackColor = vbBlack And T(3).BackColor = vbBlack And T(6).BackColor = vbBlack Then
MsgBox ("Player Black: WIN")
End If

If T(1).BackColor = vbBlack And T(4).BackColor = vbBlack And T(7).BackColor = vbBlack Then
MsgBox ("Player Black: WIN")
End If

If T(2).BackColor = vbBlack And T(5).BackColor = vbBlack And T(8).BackColor = vbBlack Then
MsgBox ("Player Black: WIN")
End If

If T(0).BackColor = vbBlack And T(4).BackColor = vbBlack And T(8).BackColor = vbBlack Then
MsgBox ("Player Black: WIN")
End If

If T(2).BackColor = vbBlack And T(4).BackColor = vbBlack And T(6).BackColor = vbBlack Then
MsgBox ("Player Black: WIN")
End If




If T(0).BackColor = vbYellow And T(1).BackColor = vbYellow And T(2).BackColor = vbYellow Then
MsgBox ("Player Yellow: WIN")
End If

If T(3).BackColor = vbYellow And T(4).BackColor = vbYellow And T(5).BackColor = vbYellow Then
MsgBox ("Player Yellow: WIN")
End If

If T(6).BackColor = vbYellow And T(7).BackColor = vbYellow And T(8).BackColor = vbYellow Then
MsgBox ("Player Yellow: WIN")
End If

If T(0).BackColor = vbYellow And T(3).BackColor = vbYellow And T(6).BackColor = vbYellow Then
MsgBox ("Player Yellow: WIN")
End If

If T(1).BackColor = vbYellow And T(4).BackColor = vbYellow And T(7).BackColor = vbYellow Then
MsgBox ("Player Yellow: WIN")
End If

If T(2).BackColor = vbYellow And T(5).BackColor = vbYellow And T(8).BackColor = vbYellow Then
MsgBox ("Player Yellow: WIN")
End If

If T(0).BackColor = vbYellow And T(4).BackColor = vbYellow And T(8).BackColor = vbYellow Then
MsgBox ("Player Yellow: WIN")
End If

If T(2).BackColor = vbYellow And T(4).BackColor = vbYellow And T(6).BackColor = vbYellow Then
MsgBox ("Player Yellow: WIN")
End If



End Sub
