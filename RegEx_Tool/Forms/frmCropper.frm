VERSION 5.00
Begin VB.Form frmCropper 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Line Cropper"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   180
      TabIndex        =   8
      Top             =   1935
      Width           =   2760
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   1485
      Width           =   2760
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   330
      Left            =   2205
      TabIndex        =   6
      Top             =   1035
      Width           =   735
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   1080
      Width           =   1905
   End
   Begin VB.ComboBox combo1 
      Height          =   315
      ItemData        =   "frmCropper.frx":0000
      Left            =   1620
      List            =   "frmCropper.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   315
      Width           =   1365
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   855
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Condition"
      Height          =   195
      Left            =   1620
      TabIndex        =   3
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Threshold Value"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   1140
   End
End
Attribute VB_Name = "frmCropper"
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
' Form Name    : AutoRegExTool.frmCropper
'
' Description  : !!! CURRENTLY NOT USED IN CURRENT PROJECT !!!
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
Dim fName As String


Private Sub cmdApply_Click()
    '= (Equal)
    '!  (Not equal)
    '> (Bigger)
    '< (Smaller)
    
    Dim ifile As Integer
    
    If (Len(Me.txtFilename) = 0) Then Exit Sub
    
    If (Dir$(Me.txtFilename) = cDialog.FileTitle) Then
        
        If Me.combo1.ListIndex = -1 Then
            MsgBox "Please select a condition first", vbExclamation
            Me.combo1.SetFocus
        End If
                        
        Select Case Me.combo1.ListIndex
                
            Case 0
            
            Case 1
            
            Case 2
            
            Case 3
                
        End Select
    Else
        MsgBox "Specified file does not exist.", vbCritical
    End If
    
End Sub

Private Sub cropLine(actionIndex As Integer, threshold As Long)
    
    Dim iFile1 As Integer
    Dim iFile2 As Integer
    Dim pos1, pos2 As Integer
    Dim sLine As String
    Dim value As Long
    
    iFile1 = FreeFile
    iFile2 = FreeFile
    
    Open fName For Input As #iFile1
    Open fName & "_cropped.txt" For Output As #iFile2
    
    While Not EOF(iFile1)
        Line Input #iFile1, sLine
        Debug.Print sLine
    Wend
    
End Sub

Private Sub cmdBrowse_Click()
    cDialog.CancelError = False
    cDialog.Filter = "All Files|*.*"
    cDialog.DialogTitle = "Open document for cropping"
    cDialog.ShowOpen
    fName = cDialog.FileTitle
    Me.txtFilename = cDialog.filename
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
