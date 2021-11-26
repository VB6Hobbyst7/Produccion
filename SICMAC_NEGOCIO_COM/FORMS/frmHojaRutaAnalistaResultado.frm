VERSION 5.00
Begin VB.Form frmHojaRutaAnalistaResultado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hoja de Ruta - Resultado de Visitas de Analistas"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   Icon            =   "frmHojaRutaAnalistaResultado.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   5640
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Programa de Visitas"
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   11415
      Begin SICMACT.FlexEdit grdVisitas 
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7858
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nº Visita-Cod. Cliente-Nombre del Cliente-DOI-Tipo cliente-Resultado-Observaciones-nVisita-Fecha-nHojaRutaCod"
         EncabezadosAnchos=   "700-0-2600-900-1200-2000-3700-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-5-6-X-X-X"
         ListaControles  =   "0-0-0-0-0-3-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-L-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Nº Visita"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   705
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analista"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Label Label1 
         Caption         =   "Analista:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmHojaRutaAnalistaResultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnTpoOpe As Integer

Public Sub Inicio(Optional pnTpoOpe As Integer = 0)
    Dim oDR As New ADODB.Recordset
    Dim oCred As New COMDCredito.DCOMCreditos
    Set oDR = oCred.ObtenerDatosPersonaXUser(gsCodUser)
    lnTpoOpe = pnTpoOpe
    If Not (oDR.EOF And oDR.BOF) Then
        lblUser.Caption = UCase(oDR!cUser)
        lblNombre.Caption = oDR!cPersNombre
        CargarGrillaRutas
    End If
    Me.Show 1
End Sub

Private Sub CargarGrillaRutas()
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim oCredto As New COMDCredito.DCOMCredito
    Dim oDR As New ADODB.Recordset
    Dim i As Integer
    Dim nValor As Integer
    Set oDR = oCred.ObtenerDatosHojaRutaXAnalista(lblUser.Caption, gdFecSis, 1)
    grdVisitas.Clear
    grdVisitas.FormaCabecera
    grdVisitas.Rows = 2
    
    For i = 1 To oDR.RecordCount
        
        grdVisitas.AdicionaFila
        grdVisitas.TextMatrix(i, 1) = oDR!cPersCodCliente
        grdVisitas.TextMatrix(i, 2) = oDR!cPersNombre
        grdVisitas.TextMatrix(i, 3) = oDR!cPersIDnroDNI
        grdVisitas.TextMatrix(i, 7) = oDR!nNumVisita
        grdVisitas.TextMatrix(i, 8) = oDR!dFecha
        grdVisitas.TextMatrix(i, 9) = oDR!nHojaRutaCod
        
        nValor = oCredto.DefineCondicionCredito(oDR!cPersCodCliente, , gdFecSis, False, val(""))
        If nValor = 1 Then
            grdVisitas.TextMatrix(i, 4) = "NUEVO"
        Else
        'ElseIf nValor = 2 Then
            grdVisitas.TextMatrix(i, 4) = "RECURRENTE"
        'ElseIf nValor = 3 Then
        '    grdVisitas.TextMatrix(i, 4) = "PARALELO"
        'ElseIf nValor = 4 Then
        '    grdVisitas.TextMatrix(i, 4) = "REFINANCIADO"
        'ElseIf nValor = 5 Then
        '    grdVisitas.TextMatrix(i, 4) = "AMPLIADO"
        'ElseIf nValor = 6 Then
        '    grdVisitas.TextMatrix(i, 4) = "AUTOMATICO"
        'ElseIf nValor = 7 Then
        '    grdVisitas.TextMatrix(i, 4) = "ADICIONAL"
        End If
        oDR.MoveNext
    Next
    
    Set oDR = Nothing
End Sub



Private Sub cmdGuardar_Click()
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim i As Integer
    Dim nValorResult As String
    If ValidaExisteCeladasVacias = False Then
        For i = 1 To grdVisitas.Rows - 1
            nValorResult = Mid(grdVisitas.TextMatrix(i, 5), Len(grdVisitas.TextMatrix(i, 5)), 1)
            oCred.ActualizarVisitaAnalista grdVisitas.TextMatrix(i, 7), nValorResult, grdVisitas.TextMatrix(i, 6), grdVisitas.TextMatrix(i, 9)
        Next
    
        MsgBox "Los Datos se actualizaron correctamente", vbInformation, "Aviso"
        Unload Me
    Else
        MsgBox "Los Datos no pueden ser vacios", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdsalir_Click()
    If lnTpoOpe = 1 Then
        If MsgBox("Debe registrar los resultados de forma obligatoria,Está seguro que desea salir", vbOKCancel, "Aviso") = vbOk Then
            End
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub grdVisitas_OnCellChange(pnRow As Long, pnCol As Long)
    Dim rs As New ADODB.Recordset
    Dim oConst As New COMDConstantes.DCOMConstantes
    If grdVisitas.lbEditarFlex Then
        Set rs = oConst.RecuperaConstantes(10033)
        grdVisitas.CargaCombo rs
    End If
    Set rs = Nothing
    Set oConst = Nothing
End Sub

Private Sub grdVisitas_OnRowChange(pnRow As Long, pnCol As Long)
    Dim rs As New ADODB.Recordset
    Dim oConst As New COMDConstantes.DCOMConstantes
    If grdVisitas.lbEditarFlex Then
        Set rs = oConst.RecuperaConstantes(10033)
        grdVisitas.CargaCombo rs
    End If
    Set rs = Nothing
    Set oConst = Nothing
End Sub

Public Function ValidaExisteCeladasVacias() As Boolean
    Dim i As Integer
    For i = 1 To grdVisitas.Rows - 1
        If grdVisitas.TextMatrix(i, 4) = "" Or grdVisitas.TextMatrix(i, 5) = "" Then
            ValidaExisteCeladasVacias = True
            Exit Function
        End If
    Next
    ValidaExisteCeladasVacias = False
End Function
