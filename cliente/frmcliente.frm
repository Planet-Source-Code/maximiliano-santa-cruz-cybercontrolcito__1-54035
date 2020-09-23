VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmcliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcerrarsesion 
      Caption         =   "Cerrar Sesion"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Monto a pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   4455
      Begin VB.Label lblmonto 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1560
         TabIndex        =   11
         Top             =   480
         Width           =   165
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hora de ingreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4455
      Begin VB.Label lblhoraingreso 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiempo transcurrido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   165
      End
      Begin VB.Label lblhoras 
         AutoSize        =   -1  'True
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblminutos 
         AutoSize        =   -1  'True
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblsegundos 
         AutoSize        =   -1  'True
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   5880
   End
   Begin MSWinsockLib.Winsock wscliente 
      Left            =   4200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdconectar 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label label6 
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmcliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public miid As Long

Private Sub cmdcerrarsesion_Click()
    wscliente.SendData "cerrarsesion+" & miid
End Sub

Private Sub cmdconectar_Click()
    wscliente.Close
    wscliente.Connect "127.0.0.1", 1001
    cmdconectar.Visible = False
End Sub

Private Sub Label2_Click()

End Sub


Private Sub Timer1_Timer()
label6.Caption = "0"

lblsegundos.Caption = Val(lblsegundos.Caption) + 1
If Val(lblsegundos.Caption) < 10 Then
    lblsegundos.Caption = "0" & lblsegundos.Caption
End If


    If Val(lblsegundos.Caption) > 59 Then
        lblsegundos.Caption = "00"
        lblminutos.Caption = Val(lblminutos.Caption) + 1
            If Val(lblminutos.Caption) < 10 Then
                lblminutos.Caption = "0" & lblminutos.Caption
            End If
        If Val(lblminutos.Caption) > 59 Then
            lblminutos.Caption = "00"
            lblhoras.Caption = Val(lblhoras.Caption) + 1
            
            If Val(lblhoras.Caption) < 10 Then
                lblhoras.Caption = "0" & lblhoras.Caption
            End If
        End If
    End If
End Sub

Private Sub wscliente_ConnectionRequest(ByVal requestID As Long)
    wscliente.Accept requestID
End Sub

Private Sub wscliente_DataArrival(ByVal bytesTotal As Long)

Dim strdata As String
Dim accion As String
Dim mensaje As String

wscliente.GetData strdata

accion = Split(strdata, "+")(0)
mensaje = Split(strdata, "+")(1)

If accion = "decimetu" And mensaje = "ip" Then
    wscliente.SendData "miipes+" & wscliente.LocalIP
End If

If accion = "mensaje" Then
    Me.Caption = mensaje
End If

If accion = "puesto" Then
    If mensaje = "bloquear" Then
        wscliente.SendData "monto+" & lblmonto.Caption
        lblhoraingreso.Caption = ""
        lblhoras.Caption = "00"
        lblminutos.Caption = "00"
        lblsegundos.Caption = "00"
        lblmonto.Caption = ""
        frmbloqueo.Show
        Me.Hide
    ElseIf mensaje = "desbloquear" Then
        frmbloqueo.Hide
        Me.Show
        lblhoraingreso.Caption = Format$(Time, "short time")
        lblmonto.Caption = "0.25"
        Timer1.Enabled = True
    End If
End If

If accion = "tuid" Then
    miid = Val(mensaje)
End If

End Sub

