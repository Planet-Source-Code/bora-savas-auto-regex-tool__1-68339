VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegEx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RegEx Options"
   ClientHeight    =   5220
   ClientLeft      =   -765
   ClientTop       =   -75
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDown 
      Height          =   240
      Left            =   7200
      Picture         =   "frmRegEx.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Move current item to down"
      Top             =   4350
      Width           =   240
   End
   Begin VB.CommandButton cmdUp 
      Height          =   240
      Left            =   7425
      Picture         =   "frmRegEx.frx":08E6
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Move current item to up"
      Top             =   4350
      Width           =   240
   End
   Begin MSComctlLib.ListView lstRegEx 
      Height          =   3315
      Left            =   75
      TabIndex        =   7
      Top             =   975
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   5847
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "RegEx"
         Text            =   "RegEx"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Definition"
         Text            =   "Definition"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Action"
         Text            =   "Action"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "ActionParameter"
         Text            =   "Action Parameter"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   6165
   End
   Begin AutoRegExTool.UserControl_Button cmdApply 
      Height          =   300
      Left            =   6375
      TabIndex        =   2
      Top             =   300
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      BTYPE           =   4
      TX              =   "&Browse"
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
      BCOL            =   33023
      FCOL            =   0
   End
   Begin AutoRegExTool.UserControl_Button cmdAddNew 
      Height          =   300
      Left            =   150
      TabIndex        =   4
      Top             =   4800
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   529
      BTYPE           =   4
      TX              =   "&Add New"
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
      BCOL            =   33023
      FCOL            =   0
   End
   Begin AutoRegExTool.UserControl_Button cmdRemove 
      Height          =   300
      Left            =   3300
      TabIndex        =   5
      Top             =   4800
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   529
      BTYPE           =   4
      TX              =   "&Remove Current"
      ENAB            =   0   'False
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
      BCOL            =   33023
      FCOL            =   0
   End
   Begin AutoRegExTool.UserControl_Button cmdCancel 
      Height          =   300
      Left            =   6300
      TabIndex        =   6
      Top             =   4800
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   529
      BTYPE           =   4
      TX              =   "&Close"
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
      BCOL            =   33023
      FCOL            =   0
   End
   Begin AutoRegExTool.UserControl_Button cmdEdit 
      Height          =   300
      Left            =   1575
      TabIndex        =   8
      Top             =   4800
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   529
      BTYPE           =   4
      TX              =   "&Edit Current"
      ENAB            =   0   'False
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
      BCOL            =   33023
      FCOL            =   0
   End
   Begin AutoRegExTool.UserControl_Button UserControl_Button1 
      Height          =   300
      Left            =   6375
      TabIndex        =   10
      Top             =   600
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      BTYPE           =   4
      TX              =   "&New RegEx"
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
      BCOL            =   33023
      FCOL            =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "This list of RegEx will be applied to the text one by one."
      Height          =   225
      Left            =   150
      TabIndex        =   9
      Top             =   4350
      Width           =   6090
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "RegEx List"
      Height          =   225
      Left            =   95
      TabIndex        =   3
      Top             =   750
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RegEx Settings File"
      Height          =   225
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1995
   End
End
Attribute VB_Name = "frmRegEx"
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
' Form Name    : AutoRegExTool.frmRegEx
'
' Description  : RegEx modification functions
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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Item index of lstRegEx
Private currentItemIndex As Long

Dim m_lCurItemIndex As Long

Private Sub cmdAddNew_Click()
    Load frmEntry
    editMode = False
    frmEntry.Show vbModal, Me
End Sub

Private Sub cmdApply_Click()
    cDialog.CancelError = False
    cDialog.Filter = "LST files|*.lst"
    cDialog.DialogTitle = "Open RegEx Settings List file"
    cDialog.ShowOpen
    
    If Len(cDialog.filename) > 1 Then
        txtPath.Text = cDialog.filename
        currentRegExFile = cDialog.filename
        WriteINI "RegExSettings", "CurrentRegExFile", currentRegExFile, DEFAULT_SETTINGS_FILE
        lstRegEx.ListItems.Clear
        loadListView lstRegEx, currentRegExFile
        'Get total count of regex
        totalRegEx = frmRegEx.lstRegEx.ListItems.Count
    End If
End Sub

Private Sub cmdCancel_Click()
    'Save the list
    saveListView frmRegEx.lstRegEx, currentRegExFile
    Me.Hide
End Sub

Private Sub cmdDown_Click()

    If currentItemIndex = lstRegEx.ListItems.Count Or currentItemIndex = 0 Then Exit Sub
    
    Dim liMoving As ListItem
    Dim liNew As ListItem
    Dim newPos As Integer
    
    newPos = currentItemIndex + 2
    lstRegEx.SetFocus

    Set liMoving = lstRegEx.SelectedItem
    
    ' Add the item in its new position
    Set liNew = lstRegEx.ListItems.Add(newPos, , liMoving.Text)
        liNew.SubItems(1) = liMoving.SubItems(1)
        liNew.SubItems(2) = liMoving.SubItems(2)
        liNew.SubItems(3) = liMoving.SubItems(3)
        liNew.Selected = True

    ' Remove the item from its old position
    lstRegEx.ListItems.Remove (currentItemIndex)
    currentItemIndex = newPos - 1
    
    Set liMoving = Nothing
    Set liNew = Nothing

End Sub

Private Sub cmdEdit_Click()
        
    currentSelectedItem = lstRegEx.SelectedItem.Index
        
    Load frmEntry
    With frmEntry
        .txtRegEx = lstRegEx.ListItems(currentSelectedItem).Text
        .txtDefinition = lstRegEx.ListItems(currentSelectedItem).SubItems(1)
        .comboAction.Text = lstRegEx.ListItems(currentSelectedItem).SubItems(2)
        .txtParameter = lstRegEx.ListItems(currentSelectedItem).SubItems(3)
        editMode = True
        .Show vbModal, Me
    End With
    
End Sub

Private Sub cmdRemove_Click()

    currentSelectedItem = lstRegEx.SelectedItem.Index
    
    If MsgBox("Are you sure to remove this item ?", vbYesNo + vbQuestion, lstRegEx.SelectedItem.Text) = vbYes Then
        lstRegEx.ListItems.Remove (currentSelectedItem)
        totalRegEx = totalRegEx - 1
        saveListView lstRegEx, currentRegExFile
    End If
    
End Sub

Private Sub cmdUp_Click()

    If currentItemIndex = 1 Or currentItemIndex = 0 Then Exit Sub
    
    Dim liMoving As ListItem
    Dim liNew As ListItem
    Dim newPos As Integer
    
    newPos = currentItemIndex - 1
    lstRegEx.SetFocus

    Set liMoving = lstRegEx.SelectedItem
    
    ' Add the item in its new position
    Set liNew = lstRegEx.ListItems.Add(newPos, , liMoving.Text)
        liNew.SubItems(1) = liMoving.SubItems(1)
        liNew.SubItems(2) = liMoving.SubItems(2)
        liNew.SubItems(3) = liMoving.SubItems(3)
        liNew.Selected = True

    ' Remove the item from its old position
    lstRegEx.ListItems.Remove (currentItemIndex + 1)
    currentItemIndex = newPos
    
    Set liMoving = Nothing
    Set liNew = Nothing
    
End Sub

Private Sub Form_Load()
    
    txtPath = currentRegExFile
    
    'Load listview data
    If Not Dir(currentRegExFile) = vbNullString And Right$(Dir(currentRegExFile), 4) = ".lst" Then
        loadListView Me.lstRegEx, currentRegExFile
    End If
    
    'Get total count of regex
    totalRegEx = frmRegEx.lstRegEx.ListItems.Count
    
    Me.Caption = "RegEx Options - [" & totalRegEx & "]"
    
    currentSelectedItem = 0
    editMode = False
    
    Set TT = New CTooltip
    TT.Style = TTBalloon
    TT.Icon = TTIconInfo
    TT.DelayTime = 50
    TT.VisibleTime = 20000
    
End Sub

Private Sub lstRegEx_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If Not Item.Text = vbNullString Then
        cmdEdit.Enabled = True
        cmdRemove.Enabled = True
    Else
        cmdEdit.Enabled = False
        cmdRemove.Enabled = False
    End If
    
    currentItemIndex = Item.Index
    
    
End Sub

Private Sub lstRegEx_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


   Dim lvhti As LVHITTESTINFO
   Dim lItemIndex As Long
   
   lvhti.pt.x = x / Screen.TwipsPerPixelX
   lvhti.pt.y = y / Screen.TwipsPerPixelY
   lItemIndex = SendMessage(lstRegEx.hWnd, LVM_HITTEST, 0, lvhti) + 1
   
   If m_lCurItemIndex <> lItemIndex Then
      m_lCurItemIndex = lItemIndex
      If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
         TT.Destroy
      Else
         TT.Title = "RegEx Details"
         TT.TipText = "RegEx: " & lstRegEx.ListItems(m_lCurItemIndex).Text & vbCrLf & _
                        "Definition: " & lstRegEx.ListItems(m_lCurItemIndex).SubItems(1) & vbCrLf & _
                        "Action: " & lstRegEx.ListItems(m_lCurItemIndex).SubItems(2) & vbCrLf & _
                        "Parameter: " & lstRegEx.ListItems(m_lCurItemIndex).SubItems(3)
         TT.Create lstRegEx.hWnd
      End If
   End If
   
End Sub

Private Sub UserControl_Button1_Click()
    
    Dim filename As String
    Dim ret As Boolean
    
    filename = InputBox("Please enter the name of your new RegEx file.", "Creating New RegEx File", "NewRegExFile")
    
    ret = saveFile(filename & ".lst", vbNullString)
    If ret = False Then
        MsgBox "Invalid file name. Your file name should not include such characters like \*/", vbExclamation
        Call UserControl_Button1_Click
    End If
    
    If Len(filename) > 0 Then
        currentRegExFile = App.Path & "\" & filename & ".lst"
        WriteINI "RegExSettings", "CurrentRegExFile", currentRegExFile, DEFAULT_SETTINGS_FILE
        txtPath.Text = currentRegExFile
        lstRegEx.ListItems.Clear
    End If
    
End Sub
