VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRegVisitaJefe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Visita Jefe de Agencia"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   Icon            =   "frmCredRegVisitaJefe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraClienteSobreendeudado 
      Caption         =   "Cliente Sobreendeudado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10455
      Begin SICMACT.TxtBuscar txtPersona 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblNombrePersona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   6975
      End
      Begin VB.Label lblNumDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdVerFormato 
      Caption         =   "Ver Formato"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegistrarVisita 
      Caption         =   "Registrar Visita Jefe Ag."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos de las Visitas"
      TabPicture(0)   =   "frmCredRegVisitaJefe.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feVisitas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin SICMACT.FlexEdit feVisitas 
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6800
         Cols0           =   13
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmCredRegVisitaJefe.frx":0326
         EncabezadosAnchos=   "0-1200-1500-1200-2400-0-1200-2400-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-4-X-X-7-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-1-0-1-1-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-L-L-C-L-C-C-C-C-C"
         FormatosEdit    =   "0-0-5-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCredRegVisitaJefe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredRegVisitaJefe
'***     Descripcion:      Registro de Visita Jefe
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     10/09/2013 01:00:00 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Private fsPersCod As String
Private i As Integer

Private Sub cmdCancelar_Click()
LimpiaFlex feVisitas
txtPersona.Text = ""
txtPersona.Enabled = True 'RECO20140712
lblNombrePersona.Caption = ""
lblNumDoc.Caption = ""
cmdRegistrarVisita.Enabled = False
cmdVerFormato.Enabled = False
End Sub

Private Sub cmdRegistrarVisita_Click()
Dim nCodigo As Long
Dim nRealizo As Boolean
nCodigo = feVisitas.TextMatrix(feVisitas.row, 9)

nRealizo = frmCredRegVisitaJefeComentario.Inicio(nCodigo, feVisitas.TextMatrix(feVisitas.row, 2), feVisitas.TextMatrix(feVisitas.row, 1))

If nRealizo Then
    CargaDatos
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub cmdVerFormato_Click()
Dim nCodigo As Long
nCodigo = feVisitas.TextMatrix(feVisitas.row, 9)
Call GenerarExcel(nCodigo)
End Sub

Private Sub FEVisitas_Click()
If Trim(feVisitas.TextMatrix(1, 0)) <> "" Then
    If CLng(feVisitas.TextMatrix(feVisitas.row, 10)) = 0 Then
        cmdRegistrarVisita.Enabled = True
        cmdVerFormato.Enabled = True
    Else
        cmdRegistrarVisita.Enabled = False
        cmdVerFormato.Enabled = True
    End If
End If
End Sub

Private Sub feVisitas_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
Dim psDecripcionCamp As String
Select Case feVisitas.Col
    Case 4:
            psDecripcionCamp = feVisitas.TextMatrix(feVisitas.row, 4)
            psDecripcionCamp = frmCredRegVisitaComentario.Inicio(1, feVisitas.TextMatrix(feVisitas.row, 3), feVisitas.TextMatrix(feVisitas.row, 2), feVisitas.TextMatrix(feVisitas.row, 1), psDecripcionCamp)
    Case 7:
            psDecripcionCamp = feVisitas.TextMatrix(feVisitas.row, 7)
            If CLng(feVisitas.TextMatrix(feVisitas.row, 10)) <> 0 Then
                psDecripcionCamp = frmCredRegVisitaComentario.Inicio(2, feVisitas.TextMatrix(feVisitas.row, 3), CDate(feVisitas.TextMatrix(feVisitas.row, 11)), feVisitas.TextMatrix(feVisitas.row, 12), psDecripcionCamp)
            End If
End Select

psCodigo = psDecripcionCamp
psDescripcion = psCodigo
End Sub


Private Sub feVisitas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Cancel = ValidaFlex(feVisitas, pnCol)
End Sub

Private Sub Form_Load()
cmdRegistrarVisita.Enabled = False
cmdVerFormato.Enabled = False
End Sub

Private Sub txtPersona_EmiteDatos()
On Error GoTo ErrorPersona
    fsPersCod = txtPersona.psCodigoPersona
    
If Trim(fsPersCod) <> "" Then
    lblNombrePersona.Caption = txtPersona.psDescripcion
    lblNumDoc.Caption = txtPersona.sPersNroDoc
    txtPersona.Enabled = False
    CargaDatos
End If

    Exit Sub
ErrorPersona:
    MsgBox err.Description, vbInformation, "Error"
End Sub


Private Sub CargaDatos()
Dim rs As ADODB.Recordset
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito
    
Set rs = oCredito.VisitaAlCliente(fsPersCod)

cmdRegistrarVisita.Enabled = False
cmdVerFormato.Enabled = False
LimpiaFlex feVisitas

If Not (rs.EOF And rs.BOF) Then
    For i = 1 To rs.RecordCount
    feVisitas.AdicionaFila
    feVisitas.TextMatrix(i, 1) = rs!nNumVisita
    feVisitas.TextMatrix(i, 2) = rs!dFecVisita
    feVisitas.TextMatrix(i, 3) = rs!Analista
    feVisitas.TextMatrix(i, 4) = rs!ComAnalista
    feVisitas.TextMatrix(i, 6) = rs!Jefe
    feVisitas.TextMatrix(i, 7) = rs!ComJefe
    feVisitas.TextMatrix(i, 9) = rs!nCodVisita
    feVisitas.TextMatrix(i, 10) = rs!nEstado
    feVisitas.TextMatrix(i, 11) = rs!dFecVisitaJefe
    feVisitas.TextMatrix(i, 12) = rs!nNumVisitaJefe
    rs.MoveNext
    Next i
Else
    MsgBox "No cuenta con visitas registradas", vbInformation, "Aviso"
End If
End Sub


Private Sub GenerarExcel(ByVal pnCodigo As Long)
Dim oCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim rsObs As ADODB.Recordset
Set oCredito = New COMDCredito.DCOMCredito
        
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim ldFechaVisita As Date
    
    Dim lnExcel As Long
    
    On Error GoTo ErrorGeneraExcelFormato
    
    Set rsCredito = oCredito.VisitaAlCliente(, pnCodigo)
    
    If (rsCredito.EOF And rsCredito.BOF) Then
       MsgBox "No hay datos", vbInformation, "Aviso"
       Exit Sub
    End If
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    lsNomHoja = "Visita"
    lsFile = "ClienteSobreendeudado"
    
    
    lsArchivo = "\spooler\" & "ClienteSobreendeudado_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    'Activar Hoja
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
    
    
    xlHoja1.Cells(3, 2) = Trim(UCase(rsCredito!Direccion))
    xlHoja1.Cells(4, 2) = Trim(UCase(rsCredito!Cliente))
    xlHoja1.Cells(5, 2) = Trim(UCase(rsCredito!Entrevistado)) & "(" & Trim(UCase(rsCredito!Relacion)) & ")"
    xlHoja1.Cells(6, 2) = Trim(UCase(rsCredito!GiroNeg))
    xlHoja1.Cells(6, 8) = Trim(UCase(rsCredito!Analista))
    
    xlHoja1.Cells(82, 7) = Trim(rsCredito!nNumVisita)
    xlHoja1.Cells(82, 12) = Trim(rsCredito!dFecVisita)
    xlHoja1.Cells(83, 1) = Trim(UCase(rsCredito!ComAnalista))
    
    
    xlHoja1.Cells(91, 9) = IIf(Trim(rsCredito!nNumVisitaJefe) = "0", "", Trim(rsCredito!nNumVisitaJefe))
    xlHoja1.Cells(91, 12) = Trim(rsCredito!dFecVisitaJefe)
    xlHoja1.Cells(92, 1) = Trim(UCase(rsCredito!ComJefe))
    
    ldFechaVisita = CDate(rsCredito!dFecVisita)

    If Trim(rsCredito!dFecVisitaJefe) = "" Then
        xlHoja1.Cells(91, 7) = ""
    Else
        Dim pnTrimestre As Integer
        pnTrimestre = Month(CDate(rsCredito!dFecVisitaJefe))
        xlHoja1.Cells(91, 7) = IIf(pnTrimestre > 9, "IV", IIf(pnTrimestre > 6, "III", IIf(pnTrimestre > 3, "II", "I"))) & "-" & Year(CDate(rsCredito!dFecVisitaJefe))
    End If

    Set rsCredito = Nothing
    Set rsCredito = oCredito.MostrarDatosCreditosVisita(fsPersCod, ldFechaVisita)
    
    lnExcel = 11
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        For i = 1 To rsCredito.RecordCount
            xlHoja1.Cells(lnExcel, 2).NumberFormat = "@"
            xlHoja1.Cells(lnExcel, 2) = rsCredito!cCtaCod
            xlHoja1.Cells(lnExcel, 3) = rsCredito!Moneda
            xlHoja1.Cells(lnExcel, 4) = rsCredito!FDesem
            xlHoja1.Cells(lnExcel, 5) = rsCredito!MontoDesem
            xlHoja1.Cells(lnExcel, 6) = rsCredito!SalCap
            xlHoja1.Cells(lnExcel, 7) = rsCredito!nCuotas
            xlHoja1.Cells(lnExcel, 8) = rsCredito!CuotasPagadas
            xlHoja1.Cells(lnExcel, 9) = rsCredito!FVecimiento
            lnExcel = lnExcel + 1
            rsCredito.MoveNext
        Next i
    End If
    xlHoja1.Range(xlHoja1.Cells(lnExcel, 1), xlHoja1.Cells(25, 13)).Delete
    lnExcel = lnExcel + 3
    
    rsCredito.MoveFirst
    Dim iObs As Long
    Set rsObs = oCredito.VisitaAlClienteDetalle(pnCodigo)
    
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        For i = 1 To rsCredito.RecordCount
            xlHoja1.Cells(lnExcel, 2).NumberFormat = "@"
            xlHoja1.Cells(lnExcel, 2) = rsCredito!cCtaCod
            xlHoja1.Cells(lnExcel, 3) = rsCredito!Destino
            If Not (rsObs.EOF And rsObs.BOF) Then
                rsObs.MoveFirst
                For iObs = 1 To rsObs.RecordCount
                    If Trim(rsCredito!cCtaCod) = Trim(rsObs!cCtaCod) Then
                        xlHoja1.Cells(lnExcel, 5) = Trim(rsObs!cObservacion)
                    End If
                    rsObs.MoveNext
                Next iObs
            Else
                xlHoja1.Cells(lnExcel, 5) = ""
            End If
            lnExcel = lnExcel + 1
            rsCredito.MoveNext
        Next i
    End If
    xlHoja1.Range(xlHoja1.Cells(lnExcel, 1), xlHoja1.Cells(lnExcel + (15 - i), 13)).Delete
    
    lnExcel = lnExcel + 3
    rsCredito.MoveFirst
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        For i = 1 To rsCredito.RecordCount
            xlHoja1.Cells(lnExcel, 2).NumberFormat = "@"
            xlHoja1.Cells(lnExcel, 2) = rsCredito!cCtaCod
            xlHoja1.Cells(lnExcel, 3) = rsCredito!DiasAtraso6
            xlHoja1.Cells(lnExcel, 4) = rsCredito!DiasAtraso5
            xlHoja1.Cells(lnExcel, 5) = rsCredito!DiasAtraso4
            xlHoja1.Cells(lnExcel, 6) = rsCredito!DiasAtraso3
            xlHoja1.Cells(lnExcel, 7) = rsCredito!DiasAtraso2
            xlHoja1.Cells(lnExcel, 8) = rsCredito!DiasAtraso1
            lnExcel = lnExcel + 1
            rsCredito.MoveNext
        Next i
    End If
    xlHoja1.Range(xlHoja1.Cells(lnExcel, 1), xlHoja1.Cells(lnExcel + (15 - i), 13)).Delete
    
    Set rsCredito = Nothing
    Set rsCredito = oCredito.MostrarDeudaCentralRiesgo(txtPersona.sPersNroDoc, IIf(txtPersona.PersPersoneria = "1", "1", "2"))
    
    lnExcel = lnExcel + 3
    If Not (rsCredito.EOF And rsCredito.BOF) Then
    xlHoja1.Cells(lnExcel - 2, 1) = xlHoja1.Cells(lnExcel - 2, 1) & Trim(rsCredito!Fecha)
        For i = 1 To rsCredito.RecordCount
            xlHoja1.Cells(lnExcel, 2) = rsCredito!Entidad
            xlHoja1.Cells(lnExcel, 3) = rsCredito!Moneda
            xlHoja1.Cells(lnExcel, 4) = rsCredito!Saldo
            xlHoja1.Cells(lnExcel, 5) = rsCredito!Clasificacion
            xlHoja1.Cells(lnExcel, 6) = Trim(rsCredito!Porcentaje) & "%"
            lnExcel = lnExcel + 1
            rsCredito.MoveNext
        Next i
    End If
    xlHoja1.Range(xlHoja1.Cells(lnExcel, 1), xlHoja1.Cells(lnExcel + (15 - i), 13)).Delete
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    Set rsCredito = Nothing
    Set oCredito = Nothing
    
    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
    
    Exit Sub
ErrorGeneraExcelFormato:
    MsgBox err.Description, vbCritical, "Error a Generar El Formato Excel"
End Sub
