VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmservidor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidor"
   ClientHeight    =   5610
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6840
      TabIndex        =   23
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   23
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   22
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   21
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   20
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   19
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   18
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   17
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   16
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   15
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   14
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   13
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   12
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   11
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   10
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   9
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   8
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   6
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   5
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   4
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   3
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   2
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdpc 
      Height          =   1095
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock wsservidor 
      Index           =   0
      Left            =   7200
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnumenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnubloquearpc 
         Caption         =   "Bloquear PC"
      End
      Begin VB.Menu mnudesbloquearpc 
         Caption         =   "Desbloquear PC"
      End
      Begin VB.Menu mnumensaje 
         Caption         =   "Enviar mensaje"
      End
   End
End
Attribute VB_Name = "frmservidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intmax As Long
Public ipcliente As String
Public indicews As Long
Public indice_boton As Long

Private Sub cmdpc_Click(Index As Integer)
    indice_boton = Index
    ipcliente = Split(cmdpc(Index).Tag, "+")(0)
    indicews = Split(cmdpc(Index).Tag, "+")(1)
    PopupMenu mnumenu, 1
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 1 To cmdpc.Count - 1
        cmdpc(i).Picture = LoadPicture(ruta & "\pc.ico", , vbLPColor)
    Next
    intmax = 0
    wsservidor(0).LocalPort = 1001
    wsservidor(0).Listen
End Sub

Private Sub mnubloquearpc_Click()
    wsservidor(indicews).SendData "puesto+bloquear"
    cmdpc(indice_boton).Picture = LoadPicture(ruta & "\ocultar.ico", , vbLPColor)
End Sub

Private Sub mnudesbloquearpc_Click()
    wsservidor(indicews).SendData "puesto+desbloquear"
    cmdpc(indice_boton).Picture = LoadPicture(ruta & "\pc.ico", , vbLPColor)
End Sub

Private Sub mnumensaje_Click()
Dim mensaje As String
    mensaje = InputBox("Escriba su mensaje", "Mensaje")
    wsservidor(indicews).SendData "mensaje+" & mensaje
End Sub

Private Sub wsservidor_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
    If Index = 0 Then
      intmax = intmax + 1
      Load wsservidor(intmax)
      wsservidor(intmax).LocalPort = 0
      wsservidor(intmax).Accept requestID
      cmdpc(intmax).Visible = True
      wsservidor(intmax).SendData "decimetu+ip"
    End If
    
End Sub

Private Sub wsservidor_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim accion As String
Dim mensaje As String
Dim strdata As String

wsservidor(intmax).GetData strdata

accion = Split(strdata, "+")(0)
mensaje = Split(strdata, "+")(1)

If accion = "miipes" Then
    cmdpc(intmax).Tag = mensaje & "+" & intmax
    cmdpc(intmax).Caption = "Pc_" & intmax
    wsservidor(intmax).SendData "tuid+" & intmax
End If

If accion = "monto" Then
    MsgBox "Este cliente debe pagar: " & mensaje, vbExclamation, "Pc_" & indicews
    de.caja Format(Date, "short date"), Format(Time, "short time"), Val(mensaje), indicews
End If

If accion = "cerrarsesion" Then
    indicews = Val(mensaje)
    Call mnubloquearpc_Click
End If

End Sub
