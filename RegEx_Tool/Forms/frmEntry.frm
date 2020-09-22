VERSION 5.00
Begin VB.Form frmEntry 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4665
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AutoRegExTool.UserControl_Button cmdRegExTool 
      Height          =   315
      Left            =   4275
      TabIndex        =   12
      Top             =   525
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BTYPE           =   4
      TX              =   "+"
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
      BCOL            =   14737632
      FCOL            =   0
   End
   Begin VB.TextBox txtParameter 
      Height          =   330
      Left            =   75
      TabIndex        =   10
      Top             =   2550
      Width           =   4515
   End
   Begin VB.ComboBox comboAction 
      Height          =   345
      ItemData        =   "frmEntry.frx":0000
      Left            =   75
      List            =   "frmEntry.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1875
      Width           =   4515
   End
   Begin VB.TextBox txtDefinition 
      Height          =   330
      Left            =   75
      TabIndex        =   7
      Top             =   1200
      Width           =   4515
   End
   Begin VB.TextBox txtRegEx 
      Height          =   330
      Left            =   75
      TabIndex        =   5
      Top             =   525
      Width           =   4140
   End
   Begin AutoRegExTool.UserControl_Button cmdAdd 
      Height          =   330
      Left            =   3300
      TabIndex        =   2
      Top             =   3075
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      BTYPE           =   4
      TX              =   "&Add"
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
   Begin AutoRegExTool.UserControl_Button cmdCancel 
      Height          =   330
      Left            =   1875
      TabIndex        =   3
      Top             =   3075
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      BTYPE           =   4
      TX              =   "&Cancel"
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Action Parameter"
      Height          =   225
      Left            =   75
      TabIndex        =   11
      Top             =   2325
      Width           =   1680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Action"
      Height          =   225
      Left            =   75
      TabIndex        =   8
      Top             =   1650
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Definition (Ex: Remove Spaces)"
      Height          =   225
      Left            =   75
      TabIndex        =   6
      Top             =   975
      Width           =   3150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Regular Expression"
      Height          =   225
      Left            =   75
      TabIndex        =   4
      Top             =   300
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "New RegEx Entry"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4665
   End
End
Attribute VB_Name = "frmEntry"
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
' Form Name    : AutoRegExTool.frmEntry
'
' Description  : RegEx entry related form
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

Dim tip(4) As New CTooltip

Private Sub cmdAdd_Click()
    
    Dim itemCount As Integer
    
    Call replaceQuotations
    
    If Len(txtRegEx) = 0 Then
        MsgBox "Please enter a 'Regular Expression' first.", vbExclamation, App.Title
        txtRegEx.SetFocus
        Exit Sub
    End If
        
    If comboAction.Text = vbNullString Then
        MsgBox "Please choose an action from the list", vbExclamation, App.Title
        comboAction.SetFocus
        Exit Sub
    End If
    
    ' ---- ADD MODE ----
    If editMode = False Then
        With frmRegEx.lstRegEx
            itemCount = .ListItems.Count + 1
            .ListItems.Add itemCount, Text:=txtRegEx
            .ListItems.Item(itemCount).ListSubItems.Add Text:=txtDefinition
            .ListItems.Item(itemCount).ListSubItems.Add Text:=comboAction.Text
            If comboAction.Text = "Highlight" Or comboAction.Text = "Remove" Then
                .ListItems.Item(itemCount).ListSubItems.Add Text:=vbNullString
            Else
                .ListItems.Item(itemCount).ListSubItems.Add Text:=txtParameter
            End If
            'Save the list
            saveListView frmRegEx.lstRegEx, currentRegExFile
            
            'Increment totalRegEx count
            totalRegEx = totalRegEx + 1
            
            editMode = False
            Unload Me
        End With
    ' ---- EDIT MODE ----
    Else
        With frmRegEx.lstRegEx
            .ListItems.Remove (currentSelectedItem)
            .ListItems.Add currentSelectedItem, Text:=txtRegEx
            .ListItems.Item(currentSelectedItem).ListSubItems.Add Text:=txtDefinition
            .ListItems.Item(currentSelectedItem).ListSubItems.Add Text:=comboAction.Text
            If comboAction.Text = "Highlight" Or comboAction.Text = "Remove" Then
                .ListItems.Item(currentSelectedItem).ListSubItems.Add Text:=vbNullString
            Else
                .ListItems.Item(currentSelectedItem).ListSubItems.Add Text:=txtParameter
            End If
            'Save the list
            saveListView frmRegEx.lstRegEx, currentRegExFile
            
            editMode = False
            Unload Me
        End With
    End If
    
End Sub

Private Sub cmdCancel_Click()
    editMode = False
    Unload Me
End Sub

Private Sub cmdRegExTool_Click()
    
    frmPopup.Show vbModal, Me

    'MsgBox "You will be able to add frequently used ready-made Regular Expressions" & vbCrLf & _
    '"But unfortunately, this function is under construction now", vbInformation, "RegEx Library"
    
End Sub

Private Sub comboAction_Click()
    
    On Error Resume Next
    
    If comboAction.Text = "Highlight" Or comboAction.Text = "Remove" Then
        txtParameter.Text = "Not available for this item"
        txtParameter.Locked = True
        txtParameter.Enabled = False
    ElseIf comboAction.Text = "WriteToFile" Then
        txtParameter.Enabled = True
        txtParameter.Locked = False
        txtParameter.Text = "FileName"
        txtParameter.SetFocus
        txtParameter.SelStart = 0
        txtParameter.SelLength = Len(txtParameter)
    Else
        txtParameter.Enabled = True
        txtParameter.Locked = False
        txtParameter.Text = vbNullString
    End If
    
End Sub

Private Sub Form_Activate()
    If editMode Then
        cmdAdd.Caption = "&Edit"
    End If
End Sub

Private Sub Form_Load()
    Load frmPopup
    loadListView frmPopup.lstRegEx, App.Path & "\reglib.rxl"
    frmPopup.Move cmdRegExTool.Left, cmdRegExTool.Top + 30
    
    'RegEx Textbox
    tip(0).DelayTime = 100
    tip(0).Style = TTBalloon
    tip(0).Icon = TTIconInfo
    tip(0).Title = "Regular Expression Field"
    tip(0).TipText = "Please enter your Regular Expression. (Note: MS Regular Expressions Supported)" & vbCrLf & vbCrLf & _
                     "Note that you should use in some cases; '\r\n' (CrLf) instead of using only '\n'"
    tip(0).Create (txtRegEx.hWnd)
    
    'Definition
    tip(1).DelayTime = 100
    tip(1).Style = TTBalloon
    tip(1).Icon = TTIconInfo
    tip(1).Title = "Definition of your Regular Expression."
    tip(1).TipText = "Please describe your Regular Expression"
    tip(1).Create (txtDefinition.hWnd)
    
    'Action
    tip(2).DelayTime = 100
    tip(2).Style = TTBalloon
    tip(2).Icon = TTIconInfo
    tip(2).Title = "Action Type"
    tip(2).TipText = "Please specify the action to do when any matches found"
    tip(2).Create (comboAction.hWnd)
    
    'Parameter
    tip(3).DelayTime = 100
    tip(3).Style = TTBalloon
    tip(3).Icon = TTIconWarning
    tip(3).Title = "Parameter"
    tip(3).TipText = "Please specify the parameter of the action." & vbCrLf & vbCrLf & _
                     "Note that you should use;" & vbCrLf & vbCrLf & " '$1' insted of '\1'" & vbCrLf & _
                     " And in some cases '$r$n' (CrLf) instead of using only '\n'"
    tip(3).Create (txtParameter.hWnd)
    
    'cmdRegExTool
    tip(4).DelayTime = 100
    tip(4).Style = TTBalloon
    tip(4).Icon = TTIconInfo
    tip(4).Title = "RegEx Library"
    tip(4).TipText = "Library of frequently used Regualar Expressions ready to use." & vbCrLf & _
                        "Such as extracting URLs, Emails etc..."
    tip(4).Create (cmdRegExTool.hWnd)
    
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragForm Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragForm Me
End Sub

Private Sub replaceQuotations()
    
    txtRegEx = Replace(txtRegEx, """", "'")
    txtDefinition = Replace(txtDefinition, """", "'")
    txtParameter = Replace(txtParameter, """", "'")

End Sub
