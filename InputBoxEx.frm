VERSION 5.00
Begin VB.Form FormInputBoxEx 
   Caption         =   "“ü—Í‚µ‚Ä‚­‚¾‚³‚¢"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5055
   StartUpPosition =   1  'µ°Å° Ì«°Ñ‚Ì’†‰›
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "·¬¾Ù"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtInput 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label lblMessage 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "FormInputBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m As Integer

Private Sub cmdCancel_Click()
    m = 3
End Sub

Private Sub cmdOK_Click()
    m = 2
End Sub

Private Sub Form_Activate()
    If m = 0 Then
        Unload Me
    Else
        Do
            DoEvents
        Loop While m = 1
        Hide
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        m = 3
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtInput.Move 120, ScaleHeight - txtInput.Height - 120, ScaleWidth - 240
    lblMessage.Move 120, 120, ScaleWidth - 240, txtInput.Top - 240
    cmdOK.Left = ScaleWidth - cmdOK.Width - 120
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 120
End Sub

