VERSION 5.00
Begin VB.Form frmHojaRutaAnalista 
   Caption         =   "Hoja de Ruta"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   ControlBox      =   0   'False
   Icon            =   "frmHojaRutaAnalista.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6090
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   11040
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   12120
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Programa de Visitas"
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   12975
      Begin SICMACT.FlexEdit grdVisitas 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   7858
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nº Visita-Cod. Cliente-Nombre del Cliente-DOI-Dirección-Actividad-Teléfono-Móvil-Hora-Tipo cliente-Estado-cHojaRutaCod"
         EncabezadosAnchos=   "700-1200-2600-900-2600-2500-1200-1200-1200-1200-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-8-X-X-X"
         ListaControles  =   "0-1-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-L-L-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-6-0-0-0"
         TextArray0      =   "Nº Visita"
         SelectionMode   =   1
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
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdBuscaAnal 
         Caption         =   "..."
         Height          =   280
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   1215
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
         TabIndex        =   10
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Analista:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label lblNotificacion 
      Caption         =   "NOTA: El formato a usar para la hora será de 24hs:MM:SS                                 Ej.: 09:30:00 ó 16:45:00"
      Height          =   495
      Left            =   8040
      TabIndex        =   13
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmHojaRutaAnalista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***RECO20131112 ERS154
'***Formulario para registro y mantenimiento de hoja de ruta
'***
'***
Option Explicit

Dim nTipoOpe As Integer
Dim nTipoUser As Integer
'Dim oConect As COMConecta.DCOMConecta

Private Sub cmdAgregar_Click()
    grdVisitas.AdicionaFila
End Sub
Public Sub Inicio(ByVal pnTpoOpe As Integer, ByVal psTitulo As String)
    nTipoOpe = pnTpoOpe
    Me.Caption = Me.Caption & " " & psTitulo
    Dim oDR As New ADODB.Recordset
    Dim oCred As New COMDCredito.DCOMCreditos
    Set oDR = oCred.ObtenerDatosPersonaXUser(gsCodUser)
    If nTipoOpe <> 3 Then
    'If nTipoOpe = 2 Then
        If Not (oDR.EOF And oDR.BOF) Then
            txtUser.Text = UCase(oDR!cUser)
            lblNombre.Caption = oDR!cPersNombre
            If CargarGrillaRutas = True Then
                Me.Show 1
            End If
        End If
    Else
        If nTipoOpe = 3 Then
            cmdBuscaAnal.Visible = True
        End If
        Me.Show 1
    End If
End Sub

Private Sub cmdBuscaAnal_Click()
    Dim oDR As New ADODB.Recordset
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim oVentana As New frmListaAnalistas
    oVentana.Show 1
    
    Set oDR = oCred.ObtenerDatosPersonaXUser(oVentana.cUser)
    If Not (oDR.EOF And oDR.BOF) Then
            txtUser.Text = UCase(oDR!cUser)
            lblNombre.Caption = oDR!cPersNombre
            CargarGrillaRutas
            txtUser.Enabled = False
            cmdBuscaAnal.Visible = False
    End If
End Sub

Private Sub cmdCancelar_Click()
     If nTipoOpe = 3 Then
        cmdBuscaAnal.Visible = True
        txtUser.Text = ""
        txtUser.Enabled = True
        lblNombre.Caption = ""
     End If
     cmdQuitar.Enabled = False
     cmdAgregar.Enabled = False
     cmdCancelar.Enabled = False
     cmdGuardar.Enabled = False
     grdVisitas.Clear
     grdVisitas.FormaCabecera
     grdVisitas.Rows = 2
End Sub

Private Sub cmdGuardar_Click()
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim i As Integer
    Dim a  As String
    If grdVisitas.Rows > 0 Then
        If ValidaFormatoHora = False Then
            MsgBox "Verifique el formato de la hora. Debe coincidir con: 'HH:MM:SS", vbCritical, "Aviso"
            Exit Sub
        End If
        If ValidaDatosVacios = True Then
            For i = 1 To grdVisitas.Rows - 1
                If grdVisitas.TextMatrix(i, 10) = 0 Then
                    'Set oConect = New COMConecta.DCOMConecta
                    'a = Format(gdFecSis & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss")
                    oCred.RegistrarHojaRuta gdFecSis, grdVisitas.TextMatrix(i, 0), grdVisitas.TextMatrix(i, 1), grdVisitas.TextMatrix(i, 8), 1, txtUser.Text, gsCodUser
                End If
            Next
            If MsgBox("¿Desea imprimir la hoja de ruta?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                Call ImprimeHojaRuta
            End If
            Unload Me
        Else
            MsgBox "No se pueden guardar datos en blanco, verifique se información", vbCritical, "Aviso"
        End If
    Else
        MsgBox "Grilla vacia", vbCritical, "Aviso"
    End If
End Sub

Private Sub CmdQuitar_Click()
    If grdVisitas.TextMatrix(grdVisitas.row, 1) = "" Then
        grdVisitas.EliminaFila (grdVisitas.row)
        'MsgBox "Operación incorrecta", vbCritical, "Aviso"
        Exit Sub
    End If
    If nTipoOpe = 1 Then
        If grdVisitas.TextMatrix(grdVisitas.row, 10) = 0 Then
            grdVisitas.EliminaFila (grdVisitas.row)
        Else
            MsgBox "Imposible quitar el ruta, pertenece a un registro previo", vbCritical, "Aviso"
        End If
    Else
        If nTipoOpe = 2 Then
            cmdQuitar.Caption = "Exportar"
            Call ImprimeHojaRuta
        Else
            If grdVisitas.TextMatrix(grdVisitas.row, 10) = 0 Then
                grdVisitas.EliminaFila (grdVisitas.row)
            Else
                MsgBox "Imposible quitar la visita, pertenece a un registro previo", vbCritical, "Aviso"
            End If
        End If
    End If
End Sub

Private Sub cmdsalir_Click()
    If nTipoOpe = 4 Then
        If MsgBox("Este dato es obligatorio", vbOKCancel, "AVISO") = vbOk Then
            End
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
'    Dim oDR As New ADODB.Recordset
'    Dim oCred As New COMDCredito.DCOMCreditos
'    Set oDR = oCred.ObtenerDatosPersonaXUser(gsCodUser)
'    If gsCodCargo = "005002" Or gsCodCargo = "005003" Or gsCodCargo = "005004" Or gsCodCargo = "005005" Then
'        If Not (oDR.EOF And oDR.BOF) Then
'            lblUser.Caption = oDR!cUser
'            lblNombre.Caption = oDR!cPersNombre
'            CargarGrillaRutas
'        End If
'    Else
'        cmdBuscaAnal.Visible = True
'    End If
    If nTipoOpe = 1 Then
        cmdQuitar.Caption = "Quitar"
        txtUser.Enabled = False
    ElseIf nTipoOpe = 2 Then
        cmdQuitar.Caption = "Exportar"
        txtUser.Enabled = False
    ElseIf nTipoOpe = 3 Then
        cmdQuitar.Caption = "Quitar"
        cmdQuitar.Enabled = False
        cmdAgregar.Enabled = False
        cmdCancelar.Enabled = False
        cmdGuardar.Enabled = False
    End If
End Sub

Private Sub grdVisitas_OnCellChange(pnRow As Long, pnCol As Long)
    'If grdVisitas.Col = 1 Then
    '    If ValidaExistePersona = True Then
    '        MsgBox "No se puede registrar 2 veces la misma Persona", vbCritical, "Aviso"
    '        'grdVisitas.EliminaFila (grdVisitas.row)
    '        Exit Sub
    '    End If
    'End If
End Sub

Private Sub grdVisitas_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Dim oCred As COMDCredito.DCOMCredito
    Dim nValor As Integer
    Dim R As New ADODB.Recordset
    
     If ValidaExistePersona = True Then
            MsgBox "No se puede registrar 2 veces la misma Persona", vbCritical, "Aviso"
            grdVisitas.EliminaFila (grdVisitas.row)
            Exit Sub
    End If
        
    If psDataCod <> "" Then
        Set ClsPersona = New COMDPersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(psDataCod, BusquedaCodigo)
        grdVisitas.TextMatrix(pnRow, 3) = R!cPersIDnroDNI
        grdVisitas.TextMatrix(pnRow, 4) = R!cPersDireccDomicilio
        grdVisitas.TextMatrix(pnRow, 5) = R!cActiGiro
        grdVisitas.TextMatrix(pnRow, 6) = R!cPersTelefono
        grdVisitas.TextMatrix(pnRow, 7) = R!cPersCelular
        Set oCred = New COMDCredito.DCOMCredito
        nValor = oCred.DefineCondicionCredito(psDataCod, , gdFecSis, False, val(""))
        If nValor = 1 Then
            grdVisitas.TextMatrix(pnRow, 9) = "NUEVO"
        Else
            grdVisitas.TextMatrix(pnRow, 9) = "RECURRENTE"
        End If
        grdVisitas.TextMatrix(pnRow, 10) = 0
        
        Set oCred = Nothing
        Set R = Nothing
    End If
End Sub

Private Function ValidaFormatoHora() As Boolean
    Dim i As Integer
    Dim lsHora As String
    Dim lsH As String
    Dim lsM As String
    Dim Lss As String
    Dim lsFM As String
    Dim lsSimb1 As String
    Dim lsSimb2 As String
    For i = 1 To grdVisitas.Rows - 1
        lsHora = grdVisitas.TextMatrix(i, 8)
        If Len(lsHora) <> 8 Then
            ValidaFormatoHora = False
            Exit Function
        End If
        
        lsH = Mid(lsHora, 1, 2)
        lsSimb1 = Mid(lsHora, 3, 1)
        lsM = Mid(lsHora, 4, 2)
        lsSimb2 = Mid(lsHora, 6, 1)
        Lss = Mid(lsHora, 7, 2)
        'lsFM = Mid(lsHora, 7, 2)
                        
        If val(lsH) > 24 Then
            ValidaFormatoHora = False
            Exit Function
        End If
        If lsSimb1 <> ":" Then
            ValidaFormatoHora = False
            Exit Function
        End If
        If val(lsM) > 59 Then
            ValidaFormatoHora = False
            Exit Function
        End If
        If lsSimb2 <> ":" Then
            ValidaFormatoHora = False
            Exit Function
        End If
         If val(Lss) > 59 Then
            ValidaFormatoHora = False
            Exit Function
        End If
    Next
    ValidaFormatoHora = True
End Function

Private Function ValidaDatosVacios() As Boolean
    Dim i As Integer
    For i = 1 To grdVisitas.Rows - 1
        If grdVisitas.TextMatrix(i, 1) <> "" And grdVisitas.TextMatrix(i, 8) <> "" Then
            ValidaDatosVacios = True
        Else
            ValidaDatosVacios = False
        End If
    Next
End Function

Private Function CargarGrillaRutas() As Boolean
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim oCredto As New COMDCredito.DCOMCredito
    Dim oDR As New ADODB.Recordset
    Dim i As Integer
    Dim nValor As Integer
    
    If nTipoOpe = 4 Then
        Set oDR = oCred.ObtenerDatosHojaRutaXAnalista(txtUser.Text, gdFecSis, 3) 'Carga Datos desde el MDI
        If oDR.RecordCount > 0 Then
            frmHojaRutaAnalistaResultado.Inicio 1
           
        End If
    Else
        Set oDR = oCred.ObtenerDatosHojaRutaXAnalista(txtUser.Text, gdFecSis, 1)
    End If
    grdVisitas.Clear
    grdVisitas.FormaCabecera
    grdVisitas.Rows = 2
    If nTipoOpe = 4 Then
        If oDR.RecordCount = 0 Then
            MsgBox "Debe Registrar su Hoja de Ruta de forma obligatoria", vbInformation, "AVISO"
            CargarGrillaRutas = True
        Else
            CargarGrillaRutas = False
            Exit Function
        End If
    Else
        If nTipoOpe = 1 And oDR.RecordCount > 0 Then
            MsgBox "Visita ya esta registrada, si desea agregar nuevas visitas acceda a la opción de mantenimiento", vbExclamation, "Aviso"
            CargarGrillaRutas = False
            Exit Function
        End If
        For i = 1 To oDR.RecordCount
            
            grdVisitas.AdicionaFila
            grdVisitas.TextMatrix(i, 1) = oDR!cPersCodCliente
            grdVisitas.TextMatrix(i, 2) = oDR!cPersNombre
            grdVisitas.TextMatrix(i, 3) = oDR!cPersIDnroDNI
            grdVisitas.TextMatrix(i, 4) = oDR!cPersDireccDomicilio
            grdVisitas.TextMatrix(i, 5) = oDR!cActiGiro
            grdVisitas.TextMatrix(i, 6) = oDR!cPersTelefono
            grdVisitas.TextMatrix(i, 7) = oDR!cPersCelular
            grdVisitas.TextMatrix(i, 8) = oDR!dHora
            'grdVisitas.TextMatrix(i, 9) = oDR!cPersCelular
            grdVisitas.TextMatrix(i, 10) = oDR!nEstado
            grdVisitas.TextMatrix(i, 11) = oDR!nHojaRutaCod
            
            nValor = oCredto.DefineCondicionCredito(oDR!cPersCodCliente, , gdFecSis, False, val(""))
            If nValor = 1 Then
                grdVisitas.TextMatrix(i, 9) = "NUEVO"
            Else
            'ElseIf nValor = 2 Then
                grdVisitas.TextMatrix(i, 9) = "RECURRENTE"
            'ElseIf nValor = 3 Then
            '    grdVisitas.TextMatrix(i, 9) = "PARALELO"
            'ElseIf nValor = 4 Then
            '    grdVisitas.TextMatrix(i, 9) = "REFINANCIADO"
            'ElseIf nValor = 5 Then
            '    grdVisitas.TextMatrix(i, 9) = "AMPLIADO"
            'ElseIf nValor = 6 Then
            '    grdVisitas.TextMatrix(i, 9) = "AUTOMATICO"
            'ElseIf nValor = 7 Then
            '    grdVisitas.TextMatrix(i, 9) = "ADICIONAL"
            End If
            
            oDR.MoveNext
        Next
    End If
    cmdQuitar.Enabled = True
    cmdAgregar.Enabled = True
    cmdCancelar.Enabled = True
    cmdGuardar.Enabled = True
    CargarGrillaRutas = True
    Set oDR = Nothing
End Function

Public Function ValidaExistePersona() As Boolean
    Dim i As Integer
    Dim sPers_i As String
    
    sPers_i = grdVisitas.TextMatrix(grdVisitas.row, 1)
    
    For i = 1 To grdVisitas.Rows - 2
        If sPers_i = grdVisitas.TextMatrix(i, 1) Then
            If grdVisitas.TextMatrix(i, 1) <> "" And grdVisitas.TextMatrix(i, 8) = "" Then
                grdVisitas.SetFocus
                grdVisitas.Col = 7
                grdVisitas.row = i
                SendKeys "{Enter}"
            Else
                ValidaExistePersona = True
            Exit Function
            End If
        End If
    Next
    ValidaExistePersona = False
End Function

Public Sub ImprimeHojaRuta()
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
    lsFile = "Hoja_Ruta_Analista"
    
    lsArchivo = "\spooler\" & "Hoja_Ruta_Analista" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
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
    xlHoja1.Cells(5, 2) = lblNombre.Caption
    xlHoja1.Cells(6, 2) = gsNomAge
    
    For i = 1 To grdVisitas.Rows - 1
        xlHoja1.Cells(IniTablas + i, 1) = grdVisitas.TextMatrix(i, 0)
        xlHoja1.Cells(IniTablas + i, 2) = grdVisitas.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 3) = grdVisitas.TextMatrix(i, 3)
        If grdVisitas.TextMatrix(i, 9) = "RECURRENTE" Then
            xlHoja1.Cells(IniTablas + i, 5) = "X"
        Else
            xlHoja1.Cells(IniTablas + i, 4) = "X"
        End If
        xlHoja1.Cells(IniTablas + i, 6) = grdVisitas.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 7) = grdVisitas.TextMatrix(i, 5)
        xlHoja1.Cells(IniTablas + i, 8) = grdVisitas.TextMatrix(i, 6)
        
        xlHoja1.Cells(IniTablas + i, 9) = grdVisitas.TextMatrix(i, 7)
        xlHoja1.Cells(IniTablas + i, 10) = grdVisitas.TextMatrix(i, 8)
    Next i
    
    xlHoja1.Range(xlHoja1.Cells(10, 1), xlHoja1.Cells(i + 8, 10)).Borders.LineStyle = 1
                
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtUser.Text) = 4 Then
                Dim oDR As New ADODB.Recordset
                Dim oCred As New COMDCredito.DCOMCreditos
                
                Set oDR = oCred.ObtenerDatosPersonaXUser(txtUser.Text)
                If Not (oDR.EOF And oDR.BOF) Then
                        txtUser.Text = UCase(oDR!cUser)
                        lblNombre.Caption = oDR!cPersNombre
                        CargarGrillaRutas
                        cmdBuscaAnal.Visible = False
                        txtUser.Enabled = False
                End If
        End If
    End If
End Sub
