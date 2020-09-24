VERSION 5.00
Begin VB.Form Disclaimer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Disclaimer"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Disclaimer.frx":0000
   ScaleHeight     =   4335
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00A40404&
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "Disclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SetTopMostWindow Me.hwnd, True
End Sub
