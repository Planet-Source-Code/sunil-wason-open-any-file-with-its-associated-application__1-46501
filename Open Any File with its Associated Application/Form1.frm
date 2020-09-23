VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Any File"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   3465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   1920
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdOpenAnyFile 
      Caption         =   "Open Any File"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Select File to Open"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub cmdClose_Click()

Unload Me
End

End Sub

Private Sub cmdOpen_Click()

cdlFile.ShowOpen
Text1.Text = cdlFile.FileName

End Sub

Private Sub cmdOpenAnyFile_Click()

If Len(Trim(Text1.Text)) <> 0 Then
ShellExecute Me.hwnd, vbNullString, Text1.Text, vbNullString, "C:\", SW_SHOWNORMAL
End If

End Sub

Private Sub Form_Activate()

Text1.SetFocus

End Sub

