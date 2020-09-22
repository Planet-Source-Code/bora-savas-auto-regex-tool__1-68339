VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPopup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstRegEx 
      Height          =   2865
      Left            =   0
      TabIndex        =   2
      Top             =   225
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "RegEx"
         Text            =   "RegEx"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Definition"
         Text            =   "Definition"
         Object.Width           =   4093
      EndProperty
   End
   Begin AutoRegExTool.UserControl_Button cmdAdd 
      Default         =   -1  'True
      Height          =   315
      Left            =   3600
      TabIndex        =   3
      Top             =   3225
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      BTYPE           =   4
      TX              =   "&Add"
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
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin AutoRegExTool.UserControl_Button cmdCancel 
      Height          =   315
      Left            =   2550
      TabIndex        =   4
      Top             =   3225
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
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
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin AutoRegExTool.UserControl_Button cmdEdit 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   3225
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      BTYPE           =   4
      TX              =   "&Edit Database"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "RegEx Library"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   1365
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
Attribute VB_Name = "frmPopup"
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
' Form Name    : AutoRegExTool.frmPopup
'
' Description  : Popup window <Library>
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

Dim m_lCurItemIndex As Long
   
Private Sub cmdAdd_Click()

    Dim currentSelectedItem As Integer
    
    With frmEntry
        currentSelectedItem = lstRegEx.SelectedItem.Index
        .txtRegEx = lstRegEx.ListItems(currentSelectedItem).Text
        .txtDefinition = lstRegEx.ListItems(currentSelectedItem).SubItems(1)
        .comboAction.Text = "Highlight"
    End With
    
    Me.Hide
    
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdEdit_Click()
    MsgBox "You can edit the RegEx library by modifying the file below;" & vbCrLf & _
            App.Path & "\reglib.rxl", vbInformation
End Sub

Private Sub Form_Activate()
    cmdCancel.SetFocus
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragForm Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragForm Me
End Sub

Private Sub lstRegEx_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Not Item.Text = vbNullString Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
    
End Sub

Private Sub lstRegEx_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   Dim lvhti As LVHITTESTINFO
   Dim lItemIndex As Long
   
   lvhti.pt.x = x / Screen.TwipsPerPixelX
   lvhti.pt.y = y / Screen.TwipsPerPixelY
   lItemIndex = SendMessage(Me.lstRegEx.hWnd, LVM_HITTEST, 0, lvhti) + 1
   
   If m_lCurItemIndex <> lItemIndex Then
      m_lCurItemIndex = lItemIndex
      If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
         TT.Destroy
      Else
         TT.Title = "RegEx Details"
         TT.TipText = "RegEx: " & lstRegEx.ListItems(m_lCurItemIndex).Text & vbCrLf & _
                        "Definition: " & lstRegEx.ListItems(m_lCurItemIndex).SubItems(1)
         TT.Create lstRegEx.hWnd
      End If
   End If

End Sub
