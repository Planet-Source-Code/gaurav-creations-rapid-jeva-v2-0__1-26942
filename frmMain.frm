VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Rapid Jeva v2.0"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6405
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H8000000B&
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   6120
      Width           =   6405
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000B&
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000B&
         Caption         =   "Check This Box If Your Program Uses Command Line Parameters"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   0
         Width           =   5895
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   6450
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5741
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "7/10/01"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:51 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2520
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0752
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0864
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0976
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A88
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B9A
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CAC
            Key             =   "Drawing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DBE
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ED0
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Macro"
            Object.ToolTipText     =   "Compile"
            ImageKey        =   "Macro"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Run"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Drawing"
            Object.ToolTipText     =   "Run Applet"
            ImageKey        =   "Drawing"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Run Applet In Browser"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu build 
      Caption         =   "&Build"
      Begin VB.Menu compile 
         Caption         =   "Compile"
         Shortcut        =   {F7}
      End
      Begin VB.Menu run 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu rapplet 
         Caption         =   "Run Applet"
      End
      Begin VB.Menu rappibrow 
         Caption         =   "Run Applet In Browser"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.Text1.SelRTF
    ActiveForm.Text1.SelText = vbNullString
End Sub
Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.Text1.SelRTF = Clipboard.GetText
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.Text1.SelRTF
End Sub

Private Sub compile_Click()
vReturnValue = Shell("Command.com /K " & ActiveForm.Text5.Text & " " & ActiveForm.CommonDialog1.FileName, 1)
End Sub

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc
End Sub

Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Document " & lDocumentCount
    frmD.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub rappibrow_Click()
    ActiveForm.CommonDialog1.FileName = ""
    ActiveForm.CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm|All Files (*.*)|*.*"
    ActiveForm.CommonDialog1.ShowOpen
    If ActiveForm.CommonDialog1.FileName = "" Then
    MsgBox ("SELECT A PROGRAM TO RUN"), vbokayonly, "ERROR"
    Exit Sub
    End If
    Shell "Explorer " & ActiveForm.CommonDialog1.FileName
End Sub

Private Sub rapplet_Click()
ActiveForm.CommonDialog1.FileName = ""
ActiveForm.CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm|All Files (*.*)|*.*"
ActiveForm.CommonDialog1.ShowOpen
If ActiveForm.CommonDialog1.FileName = "" Then
MsgBox ("SELECT A FILE TO RUN"), vbokayonly, "ERROR"
Exit Sub
End If
ActiveForm.Caption = "Rapid Jeva 1.0 - Running " & ActiveForm.CommonDialog1.FileTitle
ret3 = Shell("Command.com /K " & ActiveForm.Text7.Text & " " & ActiveForm.CommonDialog1.FileTitle, vbNormalFocus)
End Sub

Private Sub run_Click()
If Check1.Value = 1 Then
Let stext = InputBox(sBox, "Commmand line executer")
ActiveForm.Text2.Text = ActiveForm.CommonDialog1.FileTitle
ActiveForm.Text3.Text = Len(ActiveForm.Text2.Text) - 5
If ActiveForm.Text3.Text < 0 Then
Exit Sub
End If
ActiveForm.Text4.Text = Left(ActiveForm.Text2.Text, ActiveForm.Text3.Text)
ret1 = Shell("Command.com /K " & ActiveForm.Text6.Text & " " & ActiveForm.Text4.Text & " " & stext, vbNormalFocus)
End If

If Check1.Value = 0 Then
ActiveForm.Text2.Text = ActiveForm.CommonDialog1.FileTitle
ActiveForm.Text3.Text = Len(ActiveForm.Text2.Text) - 5
If ActiveForm.Text3.Text < 0 Then
MsgBox ("SELECT A PROGRAM TO RUN"), vbokayonly, "ERROR"
Exit Sub
End If
ActiveForm.Text4.Text = Left(ActiveForm.Text2.Text, ActiveForm.Text3.Text)
ret1 = Shell("Command.com /K " & ActiveForm.Text6.Text & " " & ActiveForm.Text4.Text, vbNormalFocus)
End If
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Macro"
           compile_Click
        Case "Forward"
          run_Click
        Case "Drawing"
           rapplet_Click
        Case "Properties"
            rappibrow_Click
        Case "Help"
         mnuHelpAbout_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click

    End Select
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuViewOptions_Click()
Form2.Show
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    
    With ActiveForm.CommonDialog1
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.Text1.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.Text1.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With ActiveForm.CommonDialog1
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub


Private Sub mnuFileSave_Click()
    ActiveForm.CommonDialog1.FileName = ""
    ActiveForm.CommonDialog1.Filter = "JAVA Source Files (*.java)|*.java|HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm|All Files (*.*)|*.*"
    ActiveForm.CommonDialog1.ShowSave
    If ActiveForm.CommonDialog1.FileName <> "" Then
        Dim iFile As Integer
        iFile = FreeFile
        Open ActiveForm.CommonDialog1.FileName For Output As iFile
        Print #iFile, Text1.Text
        Close iFile
    End If
End Sub

Private Sub mnuFileClose_Click()
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
If ActiveForm Is Nothing Then LoadNewDoc
ActiveForm.CommonDialog1.FileName = ""
    ActiveForm.CommonDialog1.Filter = "JAVA Source Files (*.java)|*.java|HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm|All Files (*.*)|*.*"
    ActiveForm.CommonDialog1.ShowOpen
    If ActiveForm.CommonDialog1.FileName <> "" Then
        Dim iFile As Integer
        iFile = FreeFile
        Open ActiveForm.CommonDialog1.FileName For Input As iFile
        ActiveForm.Text1.Text = Input(LOF(iFile), iFile)
        Close iFile
    End If
ActiveForm.Caption = ActiveForm.CommonDialog1.FileTitle
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

