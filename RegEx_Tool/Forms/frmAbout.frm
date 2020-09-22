VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7380
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
   ScaleHeight     =   4875
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   150
      Top             =   4575
   End
   Begin AutoRegExTool.UserControl_Button cmdClose 
      Height          =   330
      Left            =   5925
      TabIndex        =   2
      Top             =   4425
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      BTYPE           =   4
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   33023
      FCOL            =   0
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warning:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   150
      TabIndex        =   9
      Top             =   3375
      Width           =   840
   End
   Begin VB.Label Label8 
      Caption         =   "http://www.japanalyzer.com"
      Height          =   300
      Left            =   4425
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Visit home page of the author"
      Top             =   1500
      Width           =   2730
   End
   Begin VB.Label Label3 
      Caption         =   "borasavas@gmail.com"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4425
      MouseIcon       =   "frmAbout.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Send email to <Bora Savas> borasavas@gmail.com"
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Label Label7 
      Caption         =   $"frmAbout.frx":02A4
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   150
      TabIndex        =   6
      Top             =   2325
      Width           =   7095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Email    :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3300
      TabIndex        =   5
      Top             =   1200
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Home Page:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3300
      TabIndex        =   4
      Top             =   1500
      Width           =   1050
   End
   Begin VB.Label Label4 
      Caption         =   $"frmAbout.frx":0396
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   150
      TabIndex        =   3
      Top             =   3375
      Width           =   7080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Created by Bora SAVAS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4005
      TabIndex        =   1
      ToolTipText     =   "Bora Savas. Currently @ Osaka University."
      Top             =   585
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Auto RegEx Tool  Ver. 1.0.31"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3300
      TabIndex        =   0
      Top             =   300
      Width           =   2940
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1980
      Left            =   150
      Picture         =   "frmAbout.frx":04B7
      Top             =   150
      Width           =   2985
   End
End
Attribute VB_Name = "frmAbout"
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
' Form Name    : AutoRegExTool.frmAbout
'
' Description  : About the program & Author
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    If splashShowed = False Then
        cmdClose.Visible = False
        Timer1.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    Label1.Caption = "Auto RegEx Tool  Ver. " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label3.ForeColor = vbBlack
    Label8.ForeColor = vbBlack
    Label3.Font.Underline = False
    Label8.Font.Underline = False
End Sub

Private Sub Label10_Click()
    ShellExecute Me.hWnd, vbNullString, "http://bsavas.homelinux.com/", vbNullString, "C:\", 1
End Sub

Private Sub Label3_Click()
    ShellExecute Me.hWnd, vbNullString, "mailto:borasavas@gmail.com", vbNullString, "C:\", 1
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label3.ForeColor = &HFF0000
    Label8.ForeColor = vbBlack
    Label3.Font.Underline = True
    Label8.Font.Underline = False
End Sub

Private Sub Label8_Click()
    ShellExecute Me.hWnd, vbNullString, "http://www.japanalyzer.com/features.asp", vbNullString, "C:\", 1
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label8.ForeColor = &HFF0000
    Label3.ForeColor = vbBlack
    Label8.Font.Underline = True
    Label3.Font.Underline = False
End Sub

Private Sub Timer1_Timer()
    splashShowed = True
    Timer1.Enabled = False
    cmdClose.Visible = True
    Load frmMain
    frmMain.Show
    Unload Me
End Sub
