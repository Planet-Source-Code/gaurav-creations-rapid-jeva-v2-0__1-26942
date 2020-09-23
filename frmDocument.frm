VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   3030
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":030A
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_Resize
    Open App.Path & "\javac.ji" For Input As #1
  
    Input #1, javac
    
    Close #1
    Text5.Text = javac
    Open App.Path & "\java.ji" For Input As #1
  
    Input #1, java
    
    Close #1
    Text6.Text = java
    Open App.Path & "\javaapplet.ji" For Input As #1
  
    Input #1, javaa
    
    Close #1
    Text7.Text = javaa
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    Text1.RightMargin = Text1.Width - 400
End Sub

