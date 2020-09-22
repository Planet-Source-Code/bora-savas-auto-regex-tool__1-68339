VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto RegEx Tool"
   ClientHeight    =   8220
   ClientLeft      =   -15
   ClientTop       =   585
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Case Sensitive"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9750
      TabIndex        =   8
      Top             =   7920
      Width           =   1890
   End
   Begin RichTextLib.RichTextBox txtEdit 
      Height          =   7275
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   12832
      _Version        =   393217
      ScrollBars      =   3
      FileName        =   "C:\Documents and Settings\bsavas\My Documents\Projects\RegEx Tool\info.rtf"
      TextRTF         =   $"frmMain.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AutoRegExTool.UserControl_Button cmdApply 
      Height          =   315
      Left            =   9900
      TabIndex        =   0
      Top             =   7440
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      BTYPE           =   4
      TX              =   "&Apply RegEx"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   16576
   End
   Begin AutoRegExTool.UserControl_Button cmdRegEx 
      Height          =   315
      Left            =   8100
      TabIndex        =   3
      Top             =   7440
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      BTYPE           =   4
      TX              =   "&RegEx Options"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin AutoRegExTool.UserControl_Button cmdGrep 
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Top             =   7440
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      BTYPE           =   4
      TX              =   "&Quick RegEx"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin VB.TextBox txtRegEx 
      Height          =   330
      Left            =   150
      TabIndex        =   6
      Text            =   "<RegularExpressions>"
      Top             =   7440
      Visible         =   0   'False
      Width           =   2190
   End
   Begin AutoRegExTool.UserControl_Button cmdEvaluate 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   7440
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      BTYPE           =   4
      TX              =   "&Eval"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin VB.Line Line1 
      X1              =   6225
      X2              =   6225
      Y1              =   7920
      Y2              =   8155
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   6300
      TabIndex        =   9
      Top             =   7920
      Width           =   3330
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   825
      TabIndex        =   2
      Top             =   7920
      Width           =   5415
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   7890
      Width           =   11715
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save as"
         Shortcut        =   ^S
      End
      Begin VB.Menu splitter01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRemoveHighlight 
         Caption         =   "&Remove Highlight"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuFont 
         Caption         =   "&Font Settings"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuRegExOptions 
         Caption         =   "&RegEx Options"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLineCropper 
         Caption         =   "&Line Cropper"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About RegEx"
      End
      Begin VB.Menu showIntro 
         Caption         =   "&Show RegEx Intro"
      End
      Begin VB.Menu mnuShowHelp 
         Caption         =   "About MS Regular Expressions"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------
'
' Copyright Notice :
' All rights reserved by Bora SAVAS 2005, 2007
' Osaka University, Japan
'
' Contact : borasavas@gmail.com
'           http://www.japanalyzer.com
'
'-------------------------------------------------------------------------
'
' Form Name    : AutoRegExTool.frmMain
'
' Description  : The main form of this application
'
'-------------------------------------------------------------------------
'
'   Redistribution and use in source and binary forms, with or without
'   modification, are permitted provided that the following conditions are
'   met:
'
'   -Redistributions of source code must retain the above copyright notice,
'    this list of conditions and the following disclaimer.
'
'   -Redistributions in binary form must reproduce the above copyright
'    notice, this list of conditions and the following disclaimer in the
'    documentation and/or other materials provided with the distribution.
'
'
'   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
'   CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
'   EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
'   PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
'   PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
'   LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
'   NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'   SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'-------------------------------------------------------------------------

Option Explicit


Private Sub Check1_Click()
    isCaseSensitive = Not isCaseSensitive
End Sub

Private Sub cmdApply_Click()
    
    closePressed = False
    
    If totalRegEx < 1 Then
        MsgBox "You should define some 'Regular Expressions' before applying.", vbExclamation, App.Title
        frmRegEx.Show vbModal, Me
        Exit Sub
    End If
    
    'Apply regular expressions
    If Len(txtEdit.Text) > 1 Then
        processTime = GetTickCount
        Call applyRegEx
    Else
        MsgBox "There is no text to apply RegEx.", vbExclamation, App.Title
    End If
    
End Sub

Private Sub cmdEvaluate_Click()
    
    Dim t As Long
    
    If (Len(txtRegEx) = 0) Then Exit Sub
    
    If Not txtRegEx = "<RegularExpressions>" Then
        Call mnuRemoveHighlight_Click
        txtRegEx.Visible = False
        cmdEvaluate.Visible = False
        cmdGrep.Visible = True
        DoEvents
        t = highlightMatches(txtEdit, txtRegEx)
        lblStatus.Caption = "Process done in " & t & " miliseconds"
    Else
        MsgBox "Give me a Regular Expression to evaluate", vbExclamation
        txtRegEx.SetFocus
        txtRegEx.SelStart = 0
        txtRegEx.SelLength = Len(txtRegEx)
    End If
        
End Sub

Private Sub cmdGrep_Click()
    
    cmdGrep.Visible = False
    txtRegEx.Visible = True
    txtRegEx.SetFocus
    txtRegEx.SelStart = 0
    txtRegEx.SelLength = Len(txtRegEx)
    cmdEvaluate.Visible = True
    
End Sub

Private Sub cmdRegEx_Click()
    Call mnuRegExOptions_Click
End Sub

Private Sub Form_Initialize()
    justLoaded = True
End Sub

Private Sub Form_Load()
    
    Dim fName As String
    Dim fSize As Integer
    
    On Error GoTo HANDLER
    
    isCaseSensitive = False
    
    DEFAULT_SETTINGS_FILE = App.Path & "\settings.ini"
    
    'txtEdit.LoadFile "info.rtf"

    'Check the default settings file
    currentRegExFile = ReadINI("RegExSettings", "CurrentRegExFile", DEFAULT_SETTINGS_FILE)
    
    'Get font settings
    fName = ReadINI("RegExSettings", "CurrentFont", DEFAULT_SETTINGS_FILE)
    fSize = ReadINI("RegExSettings", "CurrentFontSize", DEFAULT_SETTINGS_FILE)
    
    changeFontSettings fName, fSize
    
    'Check current regex file
    If Len(currentRegExFile) < 1 Then
        'Set current regex file
        currentRegExFile = App.Path & "\regex.lst"
        WriteINI "RegExSettings", "CurrentRegExFile", currentRegExFile, DEFAULT_SETTINGS_FILE
    End If
    
    Exit Sub
    
HANDLER:
    
    If Err.Number = 13 Then
        MsgBox "Error occured while reading RegEx file." & vbCrLf & _
                "Current RegEx file may be corrupted or misdefined." & vbCrLf & vbCrLf & _
                "Note: 'RegExCount' & 'CurrentFontSize' fields must be an integer.", vbCritical, currentRegExFile
    End If
    
    Load frmRegEx

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Unload frmEntry
    Unload frmRegEx
    Unload frmProgress
    Unload frmAbout
    Unload frmPopup
    
End Sub

Private Sub mnuAbout_Click()

    Load frmAbout
    frmAbout.Show vbModal, Me
    
End Sub

Private Sub mnuCopy_Click()
    If Len(txtEdit.SelText) > 0 Then
        Clipboard.Clear
        Clipboard.SetText (txtEdit.SelText)
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFont_Click()
    cDialog.CancelError = False
    cDialog.ShowFont
    'Change font settings
    If Len(cDialog.FontName) > 1 And cDialog.FontSize > 1 Then
        changeFontSettings cDialog.FontName, cDialog.FontSize
    End If
End Sub

Private Sub mnuLineCropper_Click()
        
    Load frmCropper
    frmCropper.Show vbModal, Me
    
End Sub

Private Sub mnuOpenFile_Click()
    
    If justLoaded = True Then
        justLoaded = False
        txtEdit.Text = vbNullString
    End If
    cDialog.CancelError = False
    cDialog.Filter = "Text files|*.txt"
    cDialog.DialogTitle = "Open document for applying RegEx"
    cDialog.ShowOpen
    currentFile = cDialog.FileTitle
    txtEdit.filename = cDialog.filename
    lblStatus.Caption = Len(txtEdit.Text) & " character(s) "

End Sub

Private Sub mnuRegExOptions_Click()
    Load frmRegEx
    frmRegEx.Show vbModal, Me
End Sub

Private Sub mnuRemoveHighlight_Click()

    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit)
    txtEdit.SelColor = vbBlack

End Sub

Private Sub mnuSave_Click()
    
    cDialog.CancelError = False
    cDialog.Filter = "Text files|*.txt"
    cDialog.DialogTitle = "Save as..."
    cDialog.ShowSave
    
    If Len(cDialog.filename) > 0 Then
        saveFile cDialog.filename & ".txt", txtEdit.Text
        cDialog.filename = vbNullString
    End If
    
End Sub

Private Sub mnuShowHelp_Click()
    ShellExecute Me.hWnd, vbNullString, "http://www.google.com/search?q=Microsoft+VBS+Regular+Expressions&btnG=Search", vbNullString, "C:\", 1
End Sub

Private Sub showIntro_Click()
    On Error GoTo HANDLER
    justLoaded = True
    txtEdit.filename = App.Path & "\info.rtf"
    
    Exit Sub
    
HANDLER:
    If Err.Number = 75 Then
        MsgBox "RegEx info file can not be found.", vbCritical, "I/O Error"
        justLoaded = False
    Else
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Error"
        justLoaded = False
    End If
End Sub

Private Sub txtEdit_Click()
    If justLoaded = True Then
        txtEdit.Text = vbNullString
        justLoaded = False
        MsgBox "For testing, open 'sample_marked.txt' file and click '<Apply RegEx>'." & vbCrLf & _
               "You can also check which regular expressions to be applied and its related actions by clicking '<RegEx Options>'" & _
               vbCrLf & vbCrLf & "I'd appriciate if you give me some comments and some votes :)" & vbCrLf & vbCrLf, vbInformation
        End If
End Sub

Private Sub txtRegEx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdEvaluate_Click
End Sub
