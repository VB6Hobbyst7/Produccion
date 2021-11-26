VERSION 5.00
Begin VB.Form frmAdmControlCredDesemb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Créditos para Desembolso"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18735
   Icon            =   "frmAdmControlCredDesemb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   18735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar Excel"
      Height          =   315
      Left            =   9120
      TabIndex        =   7
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   17280
      TabIndex        =   6
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalidaExp 
      Caption         =   "Salida Exp."
      Height          =   315
      Left            =   7320
      TabIndex        =   5
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdRegistroObs 
      Caption         =   "Reingreso Exp."
      Height          =   315
      Left            =   5520
      TabIndex        =   4
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdSalidaObs 
      Caption         =   "Salida por Obs"
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdControlCred 
      Caption         =   "Control de Créditos"
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdIngExp 
      Caption         =   "Ingreso de Exp."
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin SICMACT.FlexEdit feListaCreditos 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   10610
      Cols0           =   14
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Agencia-Crédito-Titular-Producto-Moneda-Monto-Ingreso Exp.-Ult.Salida Obs.-Ult.Ingreso Obs-cUser-IdControl-Revisa-CanObs"
      EncabezadosAnchos=   "300-2400-1800-3500-1800-690-1200-2200-2200-2200-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-L-L-L-R-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmAdmControlCredDesemb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmAdmControlCredDesemb
'** Descripción : Formulario que lista los creditos aprobados que esperan el check de desembolso
'** Creación    : RECO, 20150421 - ERS010-2015
'**********************************************************************************************
Option Explicit
Dim nFila As Integer

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdControlCred_Click()
    Dim nFila As Integer
    Dim sCadSalida As String
    Dim i As Integer
    Screen.MousePointer = 11
    If feListaCreditos.TextMatrix(feListaCreditos.row, 7) <> "" Then
        nFila = feListaCreditos.row
        sCadSalida = frmAdmCredRegControl.Inicio("Pre-Desembolso", feListaCreditos.TextMatrix(feListaCreditos.row, 2), 1)
        Call CargarDatos
        feListaCreditos.TopRow = nFila
        feListaCreditos.row = nFila
        Call feListaCreditos_OnRowChange(feListaCreditos.row, 1)
'        For i = 0 To Len(sCadSalida) - 1
'            If Mid(sCadSalida, i, 1) = "1" Or Mid(sCadSalida, i, 1) = "2" Then
'                If Mid(sCadSalida, i, 1) = "1" Then
'                    cmdSalidaObs.Enabled = True
'                End If
'                cmdSalidaExp.Enabled = False
'            End If
'        Next
    Else
        MsgBox "Debe registrar el ingreso del expediente.", vbInformation, "Alerta"
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdExportar_Click()
    Dim fs As New Scripting.FileSystemObject
    Dim xlsAplicacion As New Excel.Application
    Dim obj As New COMNCredito.NCOMCredito
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim RS As New ADODB.Recordset
    Dim rsObs As New ADODB.Recordset
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo As String, lsFile As String, lsNomHoja As String
    Dim i As Integer, j As Integer, IniTablas As Integer
    Dim lbExisteHoja As Boolean
    Dim nContador As Integer
    
    lsNomHoja = "CREDITOS"
    lsFile = "FormatoControlCredDesemb"
    
    lsArchivo = "\spooler\" & "Creditos_Aprobados" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsFile & ".xls")
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
    
    nContador = 5
    xlHoja1.Cells(1, 3) = gsNomAge
    xlHoja1.Cells(2, 3) = gdFecSis
    
    For i = 1 To feListaCreditos.Rows - 2
        xlHoja1.Cells(nContador, 2) = feListaCreditos.TextMatrix(i, 1)
        xlHoja1.Cells(nContador, 3) = feListaCreditos.TextMatrix(i, 2)
        xlHoja1.Cells(nContador, 4) = feListaCreditos.TextMatrix(i, 3)
        xlHoja1.Cells(nContador, 5) = feListaCreditos.TextMatrix(i, 4)
        xlHoja1.Cells(nContador, 6) = feListaCreditos.TextMatrix(i, 5)
        xlHoja1.Cells(nContador, 7) = feListaCreditos.TextMatrix(i, 6)
        xlHoja1.Cells(nContador, 8) = feListaCreditos.TextMatrix(i, 7)
        xlHoja1.Cells(nContador, 9) = feListaCreditos.TextMatrix(i, 8)
        xlHoja1.Cells(nContador, 10) = feListaCreditos.TextMatrix(i, 9)
        xlHoja1.Range(xlHoja1.Cells(nContador, 2), xlHoja1.Cells(nContador, 10)).Borders.LineStyle = 1
        nContador = nContador + 1
    Next
    
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.Path & lsArchivo
    psArchivoAGrabarC = App.Path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Private Sub cmdIngExp_Click()
    If feListaCreditos.TextMatrix(feListaCreditos.row, 1) <> "" Then
        Dim obj As New COMNCredito.NCOMCredito
        Dim sFecHoraServer As String
        Screen.MousePointer = 11
        sFecHoraServer = Format(GetFechaHoraServer, "yyyy/MM/dd hh:mm:ss")
        Call obj.RegistraControlDesembolso(feListaCreditos.TextMatrix(feListaCreditos.row, 2), sFecHoraServer, gsCodUser)
        nFila = feListaCreditos.row
        Call CargarDatos
        Screen.MousePointer = 0
        feListaCreditos.TopRow = nFila
        feListaCreditos.row = nFila
        cmdIngExp.Enabled = False
        cmdControlCred.Enabled = True
        cmdSalidaObs.Enabled = True
        Call feListaCreditos_OnRowChange(feListaCreditos.row, 1)
        'RECO20161020 ERS060-2016 **********************************************************
        Dim oNCOMColocEval As New NCOMColocEval
        Dim lcMovNro As String
        Dim sCtaCod As String
        
        sCtaCod = feListaCreditos.TextMatrix(feListaCreditos.row, 2)
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Call oNCOMColocEval.insEstadosExpediente(sCtaCod, "Adm. de Creditos", lcMovNro, "", "", "", 1, 2002, gTpoRegCtrlAdmCreditos)
        Set oNCOMColocEval = Nothing
        'RECO FIN **************************************************************************
    Else
        MsgBox "Seleccione un dato valido.", vbInformation, "Alerta"
    End If
End Sub

Private Sub cmdRegistroObs_Click()
    If feListaCreditos.TextMatrix(feListaCreditos.row, 1) <> "" Then
        nFila = feListaCreditos.row
        Call ActualizaDatos(feListaCreditos.TextMatrix(feListaCreditos.row, 11), 3)
        feListaCreditos.TopRow = nFila
        feListaCreditos.row = nFila
        Call feListaCreditos_OnRowChange(feListaCreditos.row, 1)
        
        'RECO20161020 ERS060-2016 *******************************************
        Dim oNCOMColocEval As New NCOMColocEval
        Dim lcMovNro As String
        
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Call oNCOMColocEval.updateEstadoExpediente(feListaCreditos.TextMatrix(feListaCreditos.row, 2), gTpoRegCtrlAdmCreditos) 'BY ARLO 20171027
        Call oNCOMColocEval.insEstadosExpediente(feListaCreditos.TextMatrix(feListaCreditos.row, 2), "Adm. de Creditos", "", "", lcMovNro, "", 1, 2002, gTpoRegCtrlAdmCreditos)
        MsgBox "Re Ingreso de Expediente a Adm. de Creditos", vbInformation, "Aviso"
        Set oNCOMColocEval = Nothing
        'RECO FIN ***********************************************************
    Else
        MsgBox "Seleccione un dato valido.", vbInformation, "Alerta"
    End If
End Sub

Private Sub cmdSalidaExp_Click()
    If feListaCreditos.TextMatrix(feListaCreditos.row, 1) <> "" Then
        If feListaCreditos.TextMatrix(feListaCreditos.row, 8) <> "" Then
            MsgBox "Aún no registro el reingreso por observación..", vbInformation, "Alerta"
            Exit Sub
        End If
        If Trim(feListaCreditos.TextMatrix(feListaCreditos.row, 12)) = "Verdadero" Then
            nFila = feListaCreditos.row
            'Se movio el Codigo ARLO ERS060-2016
            'RECO20161020 ERS060-2016 ******************************************************
            Dim oNCOMColocEval As New NCOMColocEval
            Dim lcMovNro As String
        
            lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            Call oNCOMColocEval.updateEstadoExpediente(feListaCreditos.TextMatrix(feListaCreditos.row, 2), gTpoRegCtrlAdmCreditos) 'BY ARLO 20171027
            Call oNCOMColocEval.insEstadosExpediente(feListaCreditos.TextMatrix(feListaCreditos.row, 2), "Analista de Creditos", "", "", "", lcMovNro, 2, 2002, gTpoRegCtrlAdmCreditos)
            MsgBox "Expediente Salio de Adm. de Creditos", vbInformation, "Aviso"
            Set oNCOMColocEval = Nothing
            'RECO FIN **********************************************************************
            Call ActualizaDatos(feListaCreditos.TextMatrix(feListaCreditos.row, 11), 4)

        Else
            MsgBox "El crédito aún no cuenta con el check de [Tipo de Revisión].", vbInformation, "Alerta"
            Exit Sub
        End If
    Else
        MsgBox "Seleccione un dato valido.", vbInformation, "Alerta"
    End If
End Sub

Private Sub cmdSalidaObs_Click()
    If feListaCreditos.TextMatrix(feListaCreditos.row, 1) <> "" Then
        nFila = feListaCreditos.row
        Call ActualizaDatos(feListaCreditos.TextMatrix(feListaCreditos.row, 11), 2)
        feListaCreditos.TopRow = nFila
        feListaCreditos.row = nFila
        Call feListaCreditos_OnRowChange(feListaCreditos.row, 1)
        
        'RECO20161020 ERS060-2016 *********************************************
        Dim oNCOMColocEval As New NCOMColocEval
        Dim lcMovNro As String
        
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Call oNCOMColocEval.updateEstadoExpediente(feListaCreditos.TextMatrix(feListaCreditos.row, 2))
        Call oNCOMColocEval.insEstadosExpediente(feListaCreditos.TextMatrix(feListaCreditos.row, 2), "Analista de Creditos", "", lcMovNro, "", "", 1, 2002, gTpoRegCtrlAdmCreditos)
        MsgBox "Expediente Salio por Observación de Adm. de Creditos", vbInformation, "Aviso"
        Set oNCOMColocEval = Nothing
        'RECO FIN *************************************************************
    Else
        MsgBox "Seleccione un dato valido.", vbInformation, "Alerta"
    End If
End Sub

Private Sub feListaCreditos_GotFocus()
     If feListaCreditos.TextMatrix(feListaCreditos.row, 7) <> "" Then
        cmdIngExp.Enabled = False
    Else
        cmdIngExp.Enabled = True
    End If
End Sub

Private Sub feListaCreditos_OnRowChange(pnRow As Long, pnCol As Long)
    If feListaCreditos.TextMatrix(feListaCreditos.row, 7) <> "" Then
     
        cmdIngExp.Enabled = False
        
        If feListaCreditos.TextMatrix(feListaCreditos.row, 13) > 0 Then
            cmdSalidaExp.Enabled = False
            cmdSalidaObs.Enabled = True
        Else
            cmdSalidaObs.Enabled = False
            cmdSalidaExp.Enabled = True
        End If
        
        If feListaCreditos.TextMatrix(feListaCreditos.row, 8) = "" Then
            cmdRegistroObs.Enabled = False
            cmdSalidaExp.Enabled = True
            cmdControlCred.Enabled = True
        Else
            cmdRegistroObs.Enabled = True
            cmdSalidaObs.Enabled = False
            cmdSalidaExp.Enabled = False
            cmdControlCred.Enabled = False
        End If
    Else
        cmdIngExp.Enabled = True
        cmdRegistroObs.Enabled = False
        cmdSalidaObs.Enabled = False
        cmdSalidaExp.Enabled = False
    End If
End Sub

Private Sub CargarDatos()
    Dim obj As New COMNCredito.NCOMCredito
    Dim objCred As New COMDCredito.DCOMCreditos
    Dim RS As New ADODB.Recordset
    Dim sCadAge As String
    Dim i As Integer
    sCadAge = ObtieneCadenaAgenciasAutorizadas(gsCodAge)
    feListaCreditos.Clear
    FormateaFlex feListaCreditos
    Set RS = obj.AdmCredListaCredAprobados(sCadAge)
    If Not (RS.EOF And RS.BOF) Then
        For i = 1 To RS.RecordCount
            feListaCreditos.AdicionaFila
            feListaCreditos.TextMatrix(i, 1) = RS!cAgeDescripcion
            feListaCreditos.TextMatrix(i, 2) = RS!cCtaCod
            feListaCreditos.TextMatrix(i, 3) = RS!cPersNombre
            feListaCreditos.TextMatrix(i, 4) = RS!cConsDescripcion
            feListaCreditos.TextMatrix(i, 5) = RS!cMoneda
            feListaCreditos.TextMatrix(i, 6) = Format(RS!nMonto, gsFormatoNumeroView)
            feListaCreditos.TextMatrix(i, 7) = IIf(RS!dIngreso = "01/01/1900", "", RS!dIngreso)
            feListaCreditos.TextMatrix(i, 8) = IIf(RS!dUltSalidaObs = "01/01/1900", "", RS!dUltSalidaObs)
            feListaCreditos.TextMatrix(i, 9) = IIf(RS!dUltIngresoObs = "01/01/1900", "", RS!dUltIngresoObs)
            feListaCreditos.TextMatrix(i, 10) = RS!cUser
            feListaCreditos.TextMatrix(i, 11) = RS!nIdControl
            feListaCreditos.TextMatrix(i, 12) = RS!bRevisaDesemb
            feListaCreditos.TextMatrix(i, 13) = RS!nCanObs
            RS.MoveNext
        Next
    End If
    feListaCreditos.row = 1
    feListaCreditos.TopRow = 1
    Screen.MousePointer = 0
End Sub

Private Function ObtieneCadenaAgenciasAutorizadas(ByVal psCtaCodAge As String) As String
    Dim objCred As New COMDCredito.DCOMCreditos
    Dim RS As New ADODB.Recordset
    Dim lsCadAge As String
    Set RS = objCred.ObtieneCredAdmAgeCodAutirza(gsCodAge)
    If Not (RS.EOF And RS.BOF) Then
        ObtieneCadenaAgenciasAutorizadas = RS!cAgeCadAuto
    End If
End Function

Private Sub Form_Load()
    Call CargarDatos
    Call feListaCreditos_OnRowChange(1, 1)
End Sub

Private Sub ActualizaDatos(ByVal nIdControl As Long, ByVal nTpoOpe As Integer)
    Dim obj As New COMNCredito.NCOMCredito
    Dim sFecHoraServer As String
    Screen.MousePointer = 11
    sFecHoraServer = Format(GetFechaHoraServer, "yyyy/MM/dd hh:mm:ss")
    Call obj.ActualizaControlDesembolso(nIdControl, sFecHoraServer, nTpoOpe)
    Call CargarDatos
End Sub


