VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "settings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3000
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   5655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Settings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "Path Where APPLETVIEWER.EXE File Is Located"
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Enter The Path Where JAVA.EXE File is Located"
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter The Path Where JAVAC.EXE Is Located"
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "EXECUTABLE Files (*.exe)|*.exe"
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName
    
End Sub

Private Sub Command2_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "EXECUTABLE Files (*.exe)|*.exe"
    CommonDialog1.ShowOpen
    Text2.Text = CommonDialog1.FileName

End Sub

Private Sub Command3_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "EXECUTABLE Files (*.exe)|*.exe"
    CommonDialog1.ShowOpen
    Text3.Text = CommonDialog1.FileName

End Sub

Private Sub Command4_Click()
    Open App.Path & "\javac.ji" For Output As #1
  
    Write #1, Text1.Text
    
    Close #1
    Open App.Path & "\java.ji" For Output As #1
  
    Write #1, Text2.Text
    
    Close #1
    
    Open App.Path & "\javaapplet.ji" For Output As #1
  
    Write #1, Text3.Text
    
    Close #1

End Sub


Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
    
    Open App.Path & "\javac.ji" For Input As #1
  
    Input #1, javac
    
    Close #1
    Text1.Text = javac
    Open App.Path & "\java.ji" For Input As #1
  
    Input #1, java
    
    Close #1
    Text2.Text = java
   
    Open App.Path & "\javaapplet.ji" For Input As #1
  
    Input #1, javaa
    
    Close #1
    Text3.Text = javaa

End Sub
