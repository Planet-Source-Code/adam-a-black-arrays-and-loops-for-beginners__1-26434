VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2640
      Top             =   1560
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a square"
      Height          =   195
      Left            =   960
      TabIndex        =   24
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   840
      TabIndex        =   23
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Left            =   1320
      TabIndex        =   22
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   195
      Left            =   1800
      TabIndex        =   21
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   195
      Left            =   2280
      TabIndex        =   20
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   1680
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   1320
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   960
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   2160
      TabIndex        =   15
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   1680
      TabIndex        =   14
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   1200
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   720
      TabIndex        =   12
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   1680
      TabIndex        =   10
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   1200
      TabIndex        =   9
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pos(1 To 4, 1 To 4) As Boolean 'whether clicked or not

Private Sub Form_Load()
    Randomize
    Const Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim i As Integer
    For i = 0 To 15
        Label1(i).Caption = Mid(Alpha, Int(Rnd * 26) + 1, 1)
    Next
    
End Sub

Private Sub Label1_Click(Index As Integer)
    Dim x As Byte, y As Byte
    Dim i As Byte
    i = 0
    For x = 1 To 4
        For y = 1 To 4
            If Index = i Then
                If Pos(x, y) = True Then
                    MsgBox "This has been click on already!"
                    Exit Sub
                End If
                MsgBox "You clicked at " & x & ", " & y & ". with a caption value of " & Label1(i).Caption
                Pos(x, y) = True
            End If
            i = i + 1
        Next
    Next
                
End Sub
