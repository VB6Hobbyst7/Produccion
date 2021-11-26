VERSION 5.00
Begin VB.Form frmAsigAgeJNTerritorial 
   Caption         =   "Asignacion de Agencia - JEFE DE NEGOCIOS TERRITORIALES"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   Icon            =   "frmAsigAgeJNTerritorial.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   400
      Left            =   6240
      TabIndex        =   3
      Top             =   3960
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   400
      Left            =   7320
      TabIndex        =   2
      Top             =   3960
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Asignaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin SICMACT.FlexEdit feAsignacion 
         Height          =   3255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5741
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencia-Jefe Neg. Territoriales-CodAgencia-CodPersona"
         EncabezadosAnchos=   "400-2000-5000-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X"
         ListaControles  =   "0-0-3-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmAsigAgeJNTerritorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmAsigAgeJNTerritorial
'***     Descripcion:       Permite relacionar una o varias agencias a un jefe de negocios territoriales
'***     Creado por:        FRHU
'***     Fecha-Tiempo:         18/02/2014 01:00:00 PM
'*****************************************************************************************
Option Explicit
Public Sub Inicio()
Call CargarAgencia
Me.Show 1
End Sub
Private Sub CargarAgencia()
Dim oCredito As New COMDCredito.DCOMCreditos
Dim rsAgencia As New ADODB.Recordset
Dim rsPersona As New ADODB.Recordset
Dim cAgeCod As String
Dim cPersCod As String

Dim rs As New ADODB.Recordset
Dim Fila As Integer

Set rsAgencia = oCredito.ObtieneAgenciasActivas
If rsAgencia.BOF Or rsAgencia.EOF Then
    Exit Sub
End If

Do While Not rsAgencia.EOF
    Fila = Fila + 1
    Me.feAsignacion.AdicionaFila
    feAsignacion.TextMatrix(Fila, 1) = rsAgencia!cAgeDescripcion
    feAsignacion.TextMatrix(Fila, 2) = rsAgencia!cPersNombre
    feAsignacion.TextMatrix(Fila, 3) = rsAgencia!cAgeCod
    feAsignacion.TextMatrix(Fila, 4) = rsAgencia!cPersCod
    cAgeCod = rsAgencia!cAgeCod
    cPersCod = rsAgencia!cPersCod
    Set rs = oCredito.ValidarJefeNegocio(cAgeCod, cPersCod)
    feAsignacion.TextMatrix(Fila, 2) = rs!Descripcion
    Set rs = Nothing
    rsAgencia.MoveNext
Loop

Set rsAgencia = Nothing
End Sub
Private Sub cmdGuardar_Click()
    
Dim oCredito As New COMDCredito.DCOMCreditos
Dim rs As New ADODB.Recordset
Dim Fila As Integer
Dim TotalFila As Integer
Dim cAgeCod As String
Dim cPersCod As String
TotalFila = feAsignacion.Rows - 1
Fila = 1
Do While Fila <= TotalFila
   cAgeCod = feAsignacion.TextMatrix(Fila, 3)
   cPersCod = feAsignacion.TextMatrix(Fila, 2)
   If cPersCod = "" Then
    MsgBox "Elegir un Jefe de Negocios Territorial para Cada Agencia", vbInformation, "ADVERTENCIA"
    Exit Sub
   End If
   Fila = Fila + 1
Loop
If MsgBox("Desea Guardar los Datos", vbYesNo, "ADVERTENCIA") = vbNo Then
    Exit Sub
End If
Fila = 1
Do While Fila <= TotalFila
   cAgeCod = feAsignacion.TextMatrix(Fila, 3)
   cPersCod = Right(feAsignacion.TextMatrix(Fila, 2), 13)
   If Not IsNumeric(cPersCod) Then
        cPersCod = feAsignacion.TextMatrix(Fila, 4)
   End If
   Set rs = oCredito.ModificaAgenciaJNTerritorial(cAgeCod, cPersCod)
   Fila = Fila + 1
Loop
Set rs = Nothing
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub feAsignacion_OnCellChange(pnRow As Long, pnCol As Long)
    Dim rs As New ADODB.Recordset
    Dim oCre As New COMDCredito.DCOMCreditos
    
    If Me.feAsignacion.lbEditarFlex Then
        Set rs = oCre.ObtieneJefeTerritorial
        feAsignacion.CargaCombo rs
    End If
    Set rs = Nothing
    Set oCre = Nothing
End Sub

Private Sub feAsignacion_OnRowChange(pnRow As Long, pnCol As Long)
    Dim rs As New ADODB.Recordset
    Dim oCre As New COMDCredito.DCOMCreditos
    
    If Me.feAsignacion.lbEditarFlex Then
        Set rs = oCre.ObtieneJefeTerritorial
        feAsignacion.CargaCombo rs
    End If
    Set rs = Nothing
    Set oCre = Nothing
End Sub
