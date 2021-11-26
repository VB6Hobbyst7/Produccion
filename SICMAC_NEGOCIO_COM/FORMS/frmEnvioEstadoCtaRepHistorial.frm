VERSION 5.00
Begin VB.Form frmEnvioEstadoCtaRepHistorial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de Generación"
   ClientHeight    =   2895
   ClientLeft      =   9555
   ClientTop       =   5490
   ClientWidth     =   3405
   Icon            =   "frmEnvioEstadoCtaRepHistorial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3405
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feGeneracion 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3836
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-Fecha de Generación"
      EncabezadosAnchos=   "400-2560"
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
      ColumnasAEditar =   "X-X"
      ListaControles  =   "0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C"
      FormatosEdit    =   "0-0"
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmEnvioEstadoCtaRepHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmEnvioEstadoCtaRepHistorial
'** Descripción : Formulario para visualizar el historial de generación de Estados de Cuenta TI-ERS036-2017
'** Creación : APRI, 20180520 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim oEnvEstCta As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rs As ADODB.Recordset

Public Function Inicio(ByVal psCtaCod As String)
    If CargarDatos(psCtaCod) Then
        Me.Show 1
    End If
End Function
Private Function CargarDatos(sCtaCod)
    Dim lnFila As Integer
    CargarDatos = True
    
    Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
    
    Set rs = oEnvEstCta.RecuperaHistorialGeneracionEnvioEstadoCta(sCtaCod)
    Set oEnvEstCta = Nothing
        
    Call LimpiaFlex(feGeneracion)
    
    If Not rs.EOF Then
            
        
            Do While Not rs.EOF
                feGeneracion.AdicionaFila
                lnFila = feGeneracion.row
                feGeneracion.TextMatrix(lnFila, 1) = Format(rs!dFechaGen, "dd/MM/yyyy")
                rs.MoveNext
            Loop

        Else
            MsgBox "No existe historial de generación", vbInformation, "Aviso"
             CargarDatos = False
        End If

    
End Function

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
