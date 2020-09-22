VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Short Date Format Handler"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNewDateFormat 
      Height          =   300
      Left            =   2955
      TabIndex        =   3
      Text            =   "yyyy/MM/dd"
      Top             =   870
      Width           =   1410
   End
   Begin VB.CommandButton cmdNewDateFormat 
      Caption         =   "Change Date Format To"
      Height          =   390
      Left            =   315
      TabIndex        =   2
      Top             =   840
      Width           =   2520
   End
   Begin VB.TextBox txtCurrDateFormat 
      Height          =   300
      Left            =   2970
      TabIndex        =   1
      Top             =   360
      Width           =   1410
   End
   Begin VB.CommandButton cmdCurrDateFormat 
      Caption         =   "Read Current Date Format"
      Height          =   390
      Left            =   315
      TabIndex        =   0
      Top             =   315
      Width           =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCurrDateFormat_Click()
txtCurrDateFormat = GetDateFormat
End Sub

Private Sub cmdNewDateFormat_Click()
Call SetDateFormat(txtNewDateFormat)
End Sub
