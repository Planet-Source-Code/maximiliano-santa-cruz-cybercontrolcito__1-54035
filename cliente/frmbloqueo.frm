VERSION 5.00
Begin VB.Form frmbloqueo 
   BorderStyle     =   0  'None
   ClientHeight    =   9435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Caption         =   "ESTACIÓN BLOQUEADA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   11295
   End
End
Attribute VB_Name = "frmbloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Unload Me
End Sub

