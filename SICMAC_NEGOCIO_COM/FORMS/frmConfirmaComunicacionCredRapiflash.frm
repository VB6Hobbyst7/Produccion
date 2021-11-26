VERSION 5.00
Begin VB.Form frmConfirmaComunicacionCredRapiflash 
   Caption         =   "Confirmación de comunicación de Credito Rapiflash"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   Icon            =   "frmConfirmaComunicacionCredRapiflash.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRece 
      Caption         =   "Recepcionado"
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin SICMACT.FlexEdit fePendiente 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5530
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Agencia-Titular-Credito-Moneda-Monto"
      EncabezadosAnchos=   "500-3000-3000-1800-1200-1200"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-R"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmConfirmaComunicacionCredRapiflash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmConfirmaComunicacionCredRapiflash
'** Descripción : Formulario que permitira indicar si se recepciono una comunicacion via email del credito en cuestion (Rapiflash)
'**               creado segun TI-ERS042-2014
'** Creación : FRHU, 20140401 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim mostrarVentana As Boolean

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdRece_Click()
    Dim oCred As New COMDCredito.DCOMCredito
    Dim Credito As String
    Credito = fePendiente.TextMatrix(fePendiente.row, 3)
    If MsgBox("Se marcara el credito: " & Credito & " como Recepcionado" & vbNewLine & _
              "Desea Continuar?", vbYesNo) = vbYes Then
        Call oCred.OpeComunicarRiesgo(Credito, 2, , GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), 1)
        FormateaFlex fePendiente
        Call CargarGrid
    End If
End Sub
Public Sub Inicia()
    Call CargarGrid
    If mostrarVentana Then
        Me.Show 1
    Else
        MsgBox "No hay Datos que Mostrar", vbInformation
    End If
End Sub
Private Sub CargarGrid()
    Dim oCred As New COMDCredito.DCOMCredito
    Dim oRs As ADODB.Recordset
    Dim fila As Integer
    Set oRs = oCred.ObtenerComunicarRiesgoPend()
    fila = 0
    If Not oRs.EOF And Not oRs.BOF Then
        Do While Not oRs.EOF
            fila = fila + 1
            fePendiente.AdicionaFila
            fePendiente.TextMatrix(fila, 1) = oRs!cAgeDescripcion
            fePendiente.TextMatrix(fila, 2) = oRs!cPersNombre
            fePendiente.TextMatrix(fila, 3) = oRs!cCtaCod
            fePendiente.TextMatrix(fila, 4) = oRs!Moneda
            fePendiente.TextMatrix(fila, 5) = Format(oRs!nMonto, "#,###,##0.00")
            oRs.MoveNext
        Loop
        Me.cmdRece.Enabled = True
        mostrarVentana = True
    Else
        mostrarVentana = False
    End If
End Sub

