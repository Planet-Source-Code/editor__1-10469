VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   4395
   ClientLeft      =   2670
   ClientTop       =   3195
   ClientWidth     =   7335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   7335
   Begin RichTextLib.RichTextBox RTFText 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StrikeThrough"
            Object.ToolTipText     =   "Strike Through"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "Font"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FontColor"
            Object.ToolTipText     =   "Font Color"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   15
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   16
            Style           =   2
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   17
            Style           =   2
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HelpMe"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   18
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4125
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7303
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "8/7/00"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "12:11 AM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2880
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":04FC
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":060E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0720
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0832
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0944
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A56
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B68
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C7A
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D8C
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E9E
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FB0
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10C2
            Key             =   "Strike Through"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11D4
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12E6
            Key             =   "FontColor"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":173A
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":184C
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":195E
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A70
            Key             =   "Help"
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
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
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
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Select All"
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
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpMe 
         Caption         =   "&Help Me!"
      End
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
Dim filechanged As Boolean
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub Form_Load()
    sbStatusBar.Panels(1).Text = "Program start: no file is loaded or file has not changed."
    filechanged = False
    Me.Caption = "Editor " & App.Major & "." & App.Minor
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 7000)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    RTFText.Move 0, 400, Me.ScaleWidth, Me.ScaleHeight - 650
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Dim msg, Response   ' Declare variables.
    If filechanged = True Then
    msg = "Your current document has not been saved!"
    msg = msg + vbCrLf + "Save before exiting?"
    Response = MsgBox(msg, vbQuestion + vbYesNo, "Exit Editor")
    Select Case Response
       Case vbYes   ' Don't allow close.
         mnuFileSaveAs_Click
         Unload Me
       Case vbNo
         SaveSetting App.Title, "Settings", "MainLeft", Me.Left
         SaveSetting App.Title, "Settings", "MainTop", Me.Top
         SaveSetting App.Title, "Settings", "MainWidth", Me.Width
         SaveSetting App.Title, "Settings", "MainHeight", Me.Height
         Unload Me
    End Select
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    RTFText.Move 0, 400, Me.ScaleWidth, Me.ScaleHeight - 650
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuEditSelectAll_Click()
      RTFText.SelStart = 0
      RTFText.SelLength = Len(RTFText.Text)
End Sub

Private Sub RTFText_Change()
    filechanged = True
    sbStatusBar.Panels(1).Text = "File has changed - not saved."
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Undo"
            mnuEditUndo_Click
        Case "Bold"
            RTFText.SelBold = Not RTFText.SelBold
            Button.Value = IIf(RTFText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            RTFText.SelItalic = Not RTFText.SelItalic
            Button.Value = IIf(RTFText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            RTFText.SelUnderline = Not RTFText.SelUnderline
            Button.Value = IIf(RTFText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Font"
            Me.dlgCommonDialog.Flags = 1
            Me.dlgCommonDialog.FontSize = RTFText.SelFontSize
            Me.dlgCommonDialog.FontName = RTFText.SelFontName
            Me.dlgCommonDialog.FontBold = RTFText.SelBold
            Me.dlgCommonDialog.FontItalic = RTFText.SelItalic
            Me.dlgCommonDialog.ShowFont
            RTFText.SelFontName = Me.dlgCommonDialog.FontName
            RTFText.SelFontSize = Me.dlgCommonDialog.FontSize
            RTFText.SelItalic = Me.dlgCommonDialog.FontItalic
            RTFText.SelBold = Me.dlgCommonDialog.FontBold
        Case "FontColor"
            dlgCommonDialog.ShowColor
            RTFText.SelColor = Me.dlgCommonDialog.Color
        Case "Align Left"
            RTFText.SelAlignment = rtfLeft
        Case "Center"
            RTFText.SelAlignment = rtfCenter
        Case "Align Right"
            RTFText.SelAlignment = rtfRight
        Case "HelpMe"
            mnuHelpMe_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Editor " & App.Major & "." & App.Minor & " Copyright (c) 2000" + vbCrLf + "by David Bowlin" + vbCrLf + vbCrLf + "Written entirely in Visual Basic 6" + vbCrLf + "Enterprise Edition, Microsoft Corp." + vbCrLf + vbCrLf + "Use at your own risk. Contact/report bugs:" + vbCrLf + "fordpref@home.com", vbOKOnly + vbInformation, "About SimpleWords"
End Sub

Private Sub mnuHelpMe_Click()
    frmHelp.Show vbModal
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    RTFText.SelRTF = Clipboard.GetText
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText RTFText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText RTFText.SelRTF
    RTFText.SelText = vbNullString
End Sub

Private Sub mnuFileExit_Click()
 Dim msg, Response
    If filechanged = True Then
    msg = "Your current document has not been saved!" & vbCrLf & "Are you sure you want to exit?"
    msg = msg + vbCrLf + "(Save your work first.)"
    Response = MsgBox(msg, vbQuestion + vbOKCancel, "Exit Editor")
    Select Case Response
       Case vbCancel
          Cancel = -1
          RTFText.SetFocus
       Case vbOK
         SaveSetting App.Title, "Settings", "MainLeft", Me.Left
         SaveSetting App.Title, "Settings", "MainTop", Me.Top
         SaveSetting App.Title, "Settings", "MainWidth", Me.Width
         SaveSetting App.Title, "Settings", "MainHeight", Me.Height
         Unload Me
    End Select
    Else
        Unload Me
    End If
End Sub

Private Sub mnuEditUndo_Click()
    Dim success&
    success& = SendMessage(RTFText.hwnd, WM_UNDO, 0&, 0&)
End Sub
Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If RTFText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If err <> MSComDlg.cdlCancel Then
            RTFText.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFileSave_Click()
    On Error GoTo err
    Dim sFile As String
        With dlgCommonDialog
            .DialogTitle = "Save RTF File"
            .CancelError = True
            ' Set the options (flags) for the dialog box.
            .Flags = cdlOFNHideReadOnly
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "Rich Text Files (*.rtf)|*.rtf"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        RTFText.SaveFile sFile, rtfRTF
        sbStatusBar.Panels(1).Text = "File saved successfully."
        filechanged = False
        RTFText.SetFocus
    Exit Sub
err:
    Exit Sub
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileSaveAs_Click()
    On Error GoTo err
    Dim sFile As String
        With dlgCommonDialog
            .DialogTitle = "Save As..."
            .CancelError = True
            ' Set the options (flags) for the dialog box.
            .Flags = cdlOFNHideReadOnly
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "Rich Text Files (*.rtf)|*.rtf"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        RTFText.SaveFile sFile, rtfRTF
        sbStatusBar.Panels(1).Text = "File saved successfully."
        filechanged = False
        RTFText.SetFocus
        End With
        Exit Sub
err:
    Exit Sub
End Sub
Private Sub mnuFileOpen_Click()
    Dim sFile As String
    With dlgCommonDialog
        .DialogTitle = "Open RTF File"
        .CancelError = False
        ' Set the options (flags) for the dialog box.
        .Flags = cdlOFNHideReadOnly
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Rich Text Files (*.rtf)|*.rtf"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    RTFText.LoadFile sFile, rtfRTF
End Sub

Private Sub mnuFileNew_Click()
 Dim msg, Response
   If filechanged = True Then
   msg = "Are you sure you want to start a new file?"
   msg = msg + vbCrLf + "(Save your current work first.)"
   Response = MsgBox(msg, vbQuestion + vbOKCancel, "New File?")
   Select Case Response
      Case vbCancel
         Cancel = -1
         RTFText.SetFocus
      Case vbOK
         RTFText.Text = ""
         sbStatusBar.Panels(1).Text = "New file loaded, not saved."
         RTFText.SetFocus
   End Select
   Else
    RTFText.Text = ""
    RTFText.SetFocus
   End If
End Sub

