VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Rapid Jeva v2.0"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Register For Free Updates"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   6135
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   6075
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   -2760
      Width           =   3615
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "Implement    your    imagination     log  on  to   gauravcreations.com ,,,,,,,, , , gauravcreations.com , implementing imagination"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   10215
         Left            =   -120
         ScaleHeight     =   10215
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   6000
         Width           =   3615
         Begin VB.Image Image2 
            Height          =   1020
            Left            =   120
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   3480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "www.gauravcreations.com                                  Sponsor This Software Place Your Ad Here"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   120
            TabIndex        =   3
            Top             =   9120
            Width           =   3375
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            Caption         =   $"frmAbout.frx":030A
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   6015
            Left            =   120
            TabIndex        =   2
            Top             =   2760
            Width           =   3495
         End
         Begin VB.Image Image1 
            Height          =   1140
            Left            =   600
            Top             =   120
            Width           =   2550
         End
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3360
      Width           =   3615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "Explorer http://chelp.bizland.com/updated.html"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\gc.jpg")
Image2.Picture = LoadPicture(App.Path & "\imagination1.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel%)
Unload Me
End Sub

Private Sub Image1_Click()
Shell "Explorer http://www.gauravcreations.com"
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub

Private Sub Label1_Click()
Shell "Explorer http://www.gauravcreations.com"
End Sub
Private Sub Label2_Click()
Shell "Explorer http://www.gauravcreations.com"
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub

Private Sub Picture1_Click()
Shell "Explorer http://www.gauravcreations.com"
End Sub
Private Sub Picture2_Click()
Shell "Explorer http://www.gauravcreations.com"
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Picture2.Top = Picture2.Top - 10
If Picture2.Top = -7000 Then
Picture2.Top = 6120
End If
End Sub
