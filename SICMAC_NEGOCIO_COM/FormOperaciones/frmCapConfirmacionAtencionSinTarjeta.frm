VERSION 5.00
Begin VB.Form frmCapConfirmacionAtencionSinTarjeta 
   Caption         =   "Confirmación de atencion sin tarjeta"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   Icon            =   "frmCapConfirmacionAtencionSinTarjeta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefrescar 
      Caption         =   "Refrescar (Alt + R)"
      Height          =   495
      Left            =   11280
      TabIndex        =   2
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   13680
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdRechazar 
      BackColor       =   &H000000FF&
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H8000000D&
      Caption         =   "Aprobar"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      Begin SICMACT.FlexEdit feLista 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6588
         Cols0           =   11
         HighLight       =   1
         RowSizingMode   =   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-ColumFake-Itm-Fecha-N° Cuenta-Cliente-Usuario-Motivo-Id-Operación-Monto"
         EncabezadosAnchos=   "300-0-400-1700-2000-3000-1000-3000-0-2000-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapConfirmacionAtencionSinTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'***Nombre      : frmCapConfirmacionAtencionSinTarjeta ----SUBIDO DESDE LA 60
'***Descripción : Formulario para aceptar o rechazar la solicitud de atención sin tarjeta
'***Creación    : MARG el 20171201, según TI-ERS 065-2017
'************************************************************************************************
Private Sub cmdAceptar_Click()
    Confimar True
End Sub

Private Sub cmdRechazar_Click()
    Confimar False
End Sub

Private Sub cmdRefrescar_Click()
    CargarDatos
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CargarDatos()
    Dim rsSol As ADODB.Recordset
    Dim row As Long
    Dim oSol As New COMDCaptaGenerales.DCOMCaptaGenerales
    
    Set rsSol = New ADODB.Recordset
    Set rsSol = oSol.ListarCapAutSinTarjetaVisto(gsCodAge)
    
    feLista.Clear
    feLista.Rows = 2
    feLista.FormaCabecera
    Do While Not rsSol.EOF
        Me.feLista.AdicionaFila
        row = feLista.row
        Me.feLista.TextMatrix(row, 1) = "x"
        Me.feLista.TextMatrix(row, 3) = rsSol!dFechaSolicitud
        Me.feLista.TextMatrix(row, 4) = rsSol!cCtaCodCliente
        Me.feLista.TextMatrix(row, 5) = rsSol!Cliente
        Me.feLista.TextMatrix(row, 6) = rsSol!cUserSolicitud
        Me.feLista.TextMatrix(row, 7) = rsSol!cMotivoAutorizacion
        Me.feLista.TextMatrix(row, 8) = rsSol!nIdVisto
        Me.feLista.TextMatrix(row, 9) = rsSol!cOpeDesc
        Me.feLista.TextMatrix(row, 10) = Format(rsSol!nMontoSolicitado, "#,##0.00")
        rsSol.MoveNext
    Loop
End Sub

Private Sub Confimar(ByVal bAceptado As Boolean)
    Dim i As Integer
    Dim nIdVisto As Integer
    Dim nNumUpdate As Integer
    Dim oSol As New COMDCaptaGenerales.DCOMCaptaGenerales
    
   
    If ValidarDatos = False Then Exit Sub
    If MsgBox("Se va a proceder a guardar los datos, desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    nNumUpdate = 0
    For i = 1 To feLista.Rows - 1
        If feLista.TextMatrix(i, 2) = "." Then
             nIdVisto = CInt(Me.feLista.TextMatrix(i, 8))
             oSol.ActualizarCapAutSinTarjetaVisto nIdVisto, gsCodUser, bAceptado
             nNumUpdate = nNumUpdate + 1
        End If
    Next
    If nNumUpdate > 0 Then
        MsgBox "Datos guardados correctamente", vbInformation, "Información"
        CargarDatos
    End If
    Set oSol = Nothing
End Sub

Public Function ValidarDatos() As Boolean
    Dim i As Integer
    Dim C As Integer
    ValidarDatos = True
    C = 0
    For i = 1 To feLista.Rows - 1
        If feLista.TextMatrix(i, 2) = "." Then
            C = C + 1
        End If
    Next
    If feLista.TextMatrix(1, 1) = "" Then
        MsgBox "No Existen Solicitudes para confirmar", vbInformation, "Mensaje"
        ValidarDatos = False
    ElseIf C = 0 Then
        MsgBox "Primero checkea la(s) Solicitud(es)", vbInformation, "Aviso"
        ValidarDatos = False
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyR And Shift = vbAltMask Then
        CargarDatos
    End If
End Sub

'GIPO 20180802
Private Sub Form_Load()
    CargarDatos
End Sub
