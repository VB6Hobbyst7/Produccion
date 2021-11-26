VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHojaRutaAnalistaConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hoja de Ruta - Consulta de Hoja de Ruta"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   Icon            =   "frmHojaRutaAnalistaConsulta.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   9
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Programa de Visitas"
      Height          =   6135
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   11535
      Begin SICMACT.FlexEdit grdVisitas 
         Height          =   5775
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11295
         _extentx        =   19923
         _extenty        =   10186
         cols0           =   11
         encabezadosnombres=   "Nº-Nombre del Cliente-DOI-Tipo Cliente-Dirección-Actividad-Teléfono-Hora Visita-Resultado-Observaciones-Fecha/Hora Resultado"
         encabezadosanchos=   "400-3400-1000-1200-4000-2000-1200-1200-2000-4500-2000"
         font            =   "frmHojaRutaAnalistaConsulta.frx":030A
         font            =   "frmHojaRutaAnalistaConsulta.frx":0336
         font            =   "frmHojaRutaAnalistaConsulta.frx":0362
         font            =   "frmHojaRutaAnalistaConsulta.frx":038E
         fontfixed       =   "frmHojaRutaAnalistaConsulta.frx":03BA
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "L-L-C-C-L-C-C-C-C-L-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0"
         textarray0      =   "Nº"
         colwidth0       =   405
         rowheight0      =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analista"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton cmdSelec 
         Caption         =   "Seleccionar"
         Height          =   375
         Left            =   10080
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   78118913
         CurrentDate     =   41625
      End
      Begin VB.CommandButton cmdBuscaAnal 
         Caption         =   "..."
         Height          =   280
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   7800
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Analista:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   735
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
   End
End
Attribute VB_Name = "frmHojaRutaAnalistaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscaAnal_Click()
    Dim oDR As New ADODB.Recordset
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim oVentana As New frmListaAnalistas
    oVentana.Show 1
    
    Set oDR = oCred.ObtenerDatosPersonaXUser(oVentana.cUser)
    If Not (oDR.EOF And oDR.BOF) Then
            txtUser.Text = UCase(oDR!cUser)
            txtNombre.Text = oDR!cPersNombre
            txtUser.Enabled = False
            cmdBuscaAnal.Enabled = False
            CmdSelec.Enabled = True
    End If
    'RECO20140621 ERS095-2014**********
    If txtUser.Text = "" Then
        Call cmdCancelar_Click
        CmdSelec.Enabled = False
    End If
    'RECO FIN**************************
End Sub

Private Sub cmdCancelar_Click()
    txtUser.Enabled = True
    cmdBuscaAnal.Enabled = True
    CmdSelec.Enabled = True
    dtpFecha.Enabled = True
    txtNombre.Text = ""
    txtUser.Text = ""
    CmdSelec.Enabled = False 'RECO20140621 ERS095
    dtpFecha.value = gdFecSis 'RECO20140621 ERS095
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Function CargarGrillaRutas() As Boolean
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim oCredto As New COMDCredito.DCOMCredito
    Dim oDR As New ADODB.Recordset
    Dim i As Integer
    Dim nValor As Integer
    Set oDR = oCred.ObtenerDatosHojaRutaXAnalista(txtUser.Text, dtpFecha.value, 2)
    grdVisitas.Clear
    grdVisitas.FormaCabecera
    grdVisitas.Rows = 2
    
    For i = 1 To oDR.RecordCount
        
        grdVisitas.AdicionaFila
        grdVisitas.TextMatrix(i, 1) = oDR!cPersNombre
        grdVisitas.TextMatrix(i, 2) = oDR!cPersIDnroDNI
        grdVisitas.TextMatrix(i, 4) = oDR!cPersDireccDomicilio
        grdVisitas.TextMatrix(i, 5) = oDR!cActiGiro
        grdVisitas.TextMatrix(i, 6) = oDR!cPersTelefono
        grdVisitas.TextMatrix(i, 7) = Format(oDR!dHora, "HH:mm")
        grdVisitas.TextMatrix(i, 8) = oDR!cConsDescripcion
        grdVisitas.TextMatrix(i, 9) = oDR!cObservaciones
        grdVisitas.TextMatrix(i, 10) = oDR!dFecResultado
        nValor = oCredto.DefineCondicionCredito(oDR!cPersCodCliente, , gdFecSis, False, val(""))
        If nValor = 1 Then
            grdVisitas.TextMatrix(i, 3) = "NUEVO"
        Else
            grdVisitas.TextMatrix(i, 3) = "RECURRENTE"
        End If
        oDR.MoveNext
    Next
    
    Set oDR = Nothing
End Function

Private Sub cmdExportar_Click()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim i As Integer: Dim IniTablas As Integer
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    
    lsNomHoja = "Hoja1"
    lsFile = "Reporte_Hoja_Ruta_Analista"
    
    lsArchivo = "\spooler\" & "Reporte_Hoja_Ruta_Analista" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
     
    IniTablas = 9
    
    xlHoja1.Cells(4, 2) = gdFecSis
    xlHoja1.Cells(5, 2) = txtNombre.Text
    xlHoja1.Cells(6, 2) = gsNomAge
    
    For i = 1 To grdVisitas.Rows - 1
        xlHoja1.Cells(IniTablas + i, 1) = grdVisitas.TextMatrix(i, 0)
        xlHoja1.Cells(IniTablas + i, 2) = grdVisitas.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 3) = grdVisitas.TextMatrix(i, 2)
        If grdVisitas.TextMatrix(i, 3) = "RECURRENTE" Then
            xlHoja1.Cells(IniTablas + i, 4) = "X"
        Else
            xlHoja1.Cells(IniTablas + i, 5) = "X"
        End If
        xlHoja1.Cells(IniTablas + i, 6) = grdVisitas.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 7) = grdVisitas.TextMatrix(i, 5)
        xlHoja1.Cells(IniTablas + i, 8) = grdVisitas.TextMatrix(i, 6)
        If grdVisitas.TextMatrix(i, 8) <> "" Then
            If grdVisitas.TextMatrix(i, 8) = "Visitado" Then
                xlHoja1.Cells(IniTablas + i, 9) = "X"
            Else
                xlHoja1.Cells(IniTablas + i, 10) = "X"
            End If
        End If
        xlHoja1.Cells(IniTablas + i, 11) = grdVisitas.TextMatrix(i, 9)
    Next i
    
    xlHoja1.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(i + 8, 11)).Borders.LineStyle = 1
    
    xlHoja1.Cells(IniTablas + i + 2, 1) = "N"
    xlHoja1.Cells(IniTablas + i + 2, 2) = "NUEVO"
    
    xlHoja1.Cells(IniTablas + i + 3, 1) = "R"
    xlHoja1.Cells(IniTablas + i + 3, 2) = "RECURRENTE"
    
    xlHoja1.Cells(IniTablas + i + 4, 1) = "V"
    xlHoja1.Cells(IniTablas + i + 4, 2) = "VISITADO"
    
    xlHoja1.Cells(IniTablas + i + 5, 1) = "NE"
    xlHoja1.Cells(IniTablas + i + 5, 2) = "NO ENCONTRADO"
    
    xlHoja1.Cells(IniTablas + i + 7, 1) = "RESUMEN"
    xlHoja1.Cells(IniTablas + i + 9, 1) = "Numero Visitas al dia"
    xlHoja1.Cells(IniTablas + i + 9, 3) = i - 1
    xlHoja1.Cells(IniTablas + i + 11, 1) = "Numero total de creditos aprobados"
    xlHoja1.Cells(IniTablas + i + 13, 1) = "Numero total de visitas de clientes en mora"
    
    xlHoja1.Cells(IniTablas + i + 13, 7) = "Analista Responsable"
    'xlHoja1.
    xlHoja1.Cells(IniTablas + i + 13, 11) = "Jefe de Agencia/Coordinador"
    
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Private Sub cmdSelec_Click()
    CargarGrillaRutas
    CmdSelec.Enabled = False
End Sub
'RECO20140621 ERS095**************************
Private Sub Form_Load()
    If gsCodCargo = "002026" Or gsCodCargo = "002036" Then
        Me.Caption = "Consulta de Hoja de Ruta - Jefe de Negocio Terrotoriales"
    End If
    CmdSelec.Enabled = False
    dtpFecha.value = gdFecSis
End Sub
'RECO FIN**************************************

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oDR As New ADODB.Recordset
        Dim oCred As New COMDCredito.DCOMCreditos
        Set oDR = oCred.ObtenerDatosPersonaXUser(txtUser.Text)
    
        If Not (oDR.EOF And oDR.BOF) Then
            'lblUser.Caption = UCase(oDR!cUser)
            txtNombre.Text = oDR!cPersNombre
            'Call cmdSelec_Click
            txtUser.Enabled = False
            txtNombre.Enabled = False
            cmdBuscaAnal.Enabled = False
        End If
    End If
End Sub
