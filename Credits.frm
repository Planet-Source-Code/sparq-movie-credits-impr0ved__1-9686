VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   2475
      Width           =   5115
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Credits.frx":0000
         Left            =   600
         List            =   "Credits.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Left            =   3900
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Align:"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   45
      Width           =   5115
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SPARQ CINEMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   4935
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1875
      Left            =   660
      Top             =   600
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   1875
      Left            =   360
      Top             =   600
      Width           =   5115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
       Command1.Visible = False
       Combo1.Visible = False
       Label2.Visible = False
       
       Label1.Alignment = Combo1.ListIndex
       Scroll "CREDITS:", vbRed
       Scroll "The Donkey:  Dom D'amore" & vbCrLf & "Superman:  Joe Dicandia" & vbCrLf & "Worthless Pot Head: SPARQ", vbGreen
       
       Label1.Font.Size = 12
       Scroll "PLEASE VOTE FOR ME ON PSC!", vbMagenta
       Label1.Font.Size = 8
       Label1.FontUnderline = True
       Scroll "THE END", vbWhite
       Label1.FontUnderline = False
       
       Command1.Visible = True
       Combo1.Visible = True
       Label2.Visible = True
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 2
    Frame1.BackColor = &HE0E0E0
    Frame2.BackColor = &HE0E0E0
    Label1.Top = Frame1.Top
End Sub


Private Sub Scroll(Text As String, Color As ColorConstants)
    Dim X
    Label1.Top = Frame1.Top
    Label1.ForeColor = Color
    Label1 = Text
    
    
    Do Until Label1.Top >= 1080
        Label1.Top = Label1.Top + 30
        X = 0
        Do Until X = 1000
            X = X + 1
            DoEvents
        Loop
    Loop
    
    X = 0
    Do Until X = 10000
        DoEvents
        X = X + 1
    Loop
    
    Do Until Label1.Top >= Frame2.Top
        Label1.Top = Label1.Top + 30
        X = 0
        Do Until X = 1000
            X = X + 1
            DoEvents
        Loop
    Loop
End Sub
