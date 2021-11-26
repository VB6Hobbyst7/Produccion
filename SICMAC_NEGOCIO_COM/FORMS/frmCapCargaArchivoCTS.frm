VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCapCargaArchivoCTS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro en Lote de Sueldos de Clientes CTS"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   Icon            =   "frmCapCargaArchivoCTS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Leyenda de Errores"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   6240
      Width           =   7455
      Begin VB.Label Label7 
         Caption         =   "Cuenta no asociada a la Empresa"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Datos incompletos"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   10575
      Begin SICMACT.FlexEdit grdCuentasCTS 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   10305
         _ExtentX        =   18309
         _ExtentY        =   7223
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-DNI-Cuenta-Nombre del Titular-MonedaSueldo-Total_4Sueldos"
         EncabezadosAnchos=   "400-1000-1700-4200-1400-1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-R"
         FormatosEdit    =   "0-0-0-0-1-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   7680
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Abrir archivo"
      Filter          =   "Archivos de Excel (*.xls;*.xlsx)"
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Cargar Archivo Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9120
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdAbrir 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Left            =   9960
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblEmpresaCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   2100
      End
      Begin VB.Label lblEmpresaNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   720
         Width           =   7305
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblPers 
         Caption         =   "Archivo:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   675
      End
      Begin VB.Label txtArchivo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   8985
      End
   End
End
Attribute VB_Name = "frmCapCargaArchivoCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCuentas As New ADODB.Recordset
Dim sNumRUC As String

Private Sub cmdAbrir_Click()
    txtArchivo.Caption = Empty
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Excel (*.xls)|*.xls|Todos los Archivo (*.*)|*.*"
On Error GoTo ErrGraba
    dlgArchivo.ShowOpen
        If dlgArchivo.FileName <> Empty Then
            txtArchivo.Caption = dlgArchivo.FileName
        Else
            txtArchivo.Caption = "NO SE ABRIO NINGUN ARCHIVO"
        End If
        If Me.txtArchivo.Caption <> "" Then
            Me.cmdProcesar.Enabled = True
        Else
            Me.cmdProcesar.Enabled = False
        End If
ErrGraba:
    If dlgArchivo.CancelError <> True Then
        MsgBox Err.Description, vbExclamation, "Error"
    End If
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    Dim i As Integer
    Dim NumFilas As Integer
    txtArchivo.Caption = ""
    Me.lblEmpresaCod.Caption = ""
    Me.lblEmpresaNombre.Caption = ""
    NumFilas = Me.grdCuentasCTS.rows - 1
    
    For i = 1 To NumFilas
        grdCuentasCTS.EliminaFila (1)
    Next
    grdCuentasCTS.BackColorRow (vbWhite)
    cmdGrabar.Enabled = False
    cmdProcesar.Enabled = True
    cmdAbrir.Enabled = True
    Me.cmdAbrir.SetFocus
End Sub
Private Sub ExcelEnd(ByRef xlAplicacion As Excel.Application, ByRef xlLibro As Excel.Workbook, ByRef xlHoja As Excel.Worksheet)
    xlLibro.Close
    Sleep (800)
    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja = Nothing
End Sub

Private Sub cmdProcesar_Click()
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim clsPers As New DCOMPersonas
    Dim rsPers As New ADODB.Recordset
    Dim rsCtasInst As New ADODB.Recordset
    Dim ClsMov As DCOMCaptaMovimiento
    
    Dim lsArchivo As String
    Dim lsNomHoja As String
    Dim CadCuentas As String
    Dim FinRegistros As Boolean
    Dim lbExisteHoja As Boolean
    Dim Valido As Boolean
    Dim i, NumErrores As Integer
    Dim filas As Integer

    Dim resp As VbMsgBoxResult
    
    Set ClsMov = New DCOMCaptaMovimiento
    
'On Error GoTo ErrGraba
    NumErrores = 0
    lsNomHoja = "CuentasCTS"
    If Me.txtArchivo.Caption = "" Then
        MsgBox "No ha seleccionado el archivo a cargar", vbInformation, "Aviso"
        Me.cmdAbrir.SetFocus
        Exit Sub
    End If
    lsArchivo = Me.dlgArchivo.FileName
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(lsArchivo)
    
    For Each xlHoja1 In xlLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
        End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    If Trim(xlHoja1.Range("A1").value) <> "RUC" Or Trim(xlHoja1.Range("A2").value) <> "Nombre" Or Trim(xlHoja1.Range("A3").value) <> "Fecha" Then
        MsgBox "El archivo no tiene el formato correcto", vbInformation, "Aviso"
        Me.cmdAbrir.SetFocus
        Call ExcelEnd(xlAplicacion, xlLibro, xlHoja1)
        Exit Sub
    End If
    If Trim(xlHoja1.Range("A7").value) <> "DNI" Or Trim(xlHoja1.Range("B7").value) <> "Nº CUENTA CTS" Or Trim(xlHoja1.Range("C7").value) <> "APELLIDOS Y NOMBRES" Or Trim(xlHoja1.Range("D7").value) <> "MONEDA DEL SUELDO" Or Trim(xlHoja1.Range("E7").value) <> "Total Sueldo (4 meses)" Then 'APRI20180504 MEJORA CAMBIO Total Sueldo (6 meses) -> Total Sueldo (4 meses)
        MsgBox "El archivo no tiene el formato correcto", vbInformation, "Aviso"
        Me.cmdAbrir.SetFocus
        Call ExcelEnd(xlAplicacion, xlLibro, xlHoja1)
        Exit Sub
    End If
    
    If Len(Trim(xlHoja1.Range("B1").value)) = 11 And IsNumeric(xlHoja1.Range("B1").value) = True Then
        sNumRUC = Trim(xlHoja1.Range("B1").value)
        Set rsPers = clsPers.BuscaCliente(sNumRUC, BusquedaDocumento)
        If Not rsPers.EOF And Not rsPers.BOF Then
            Me.lblEmpresaCod = rsPers!cPersCod
            Me.lblEmpresaNombre = rsPers!cPersNombre
            rsPers.MoveNext
        Else
            MsgBox "El RUC ingresado ha sido registrado", vbInformation, "Error"
            Me.cmdAbrir.SetFocus
            Call ExcelEnd(xlAplicacion, xlLibro, xlHoja1)
            Exit Sub
        End If
    Else
        MsgBox "El RUC de la Empresa no tiene el formato correcto", vbInformation, "Error"
        Me.cmdAbrir.SetFocus
        Exit Sub
    End If
    
    i = 0
    FinRegistros = False
    While FinRegistros = False
       i = i + 1
       grdCuentasCTS.AdicionaFila
       grdCuentasCTS.SetFocus
       grdCuentasCTS.BackColorRow (vbWhite)
       grdCuentasCTS.ForeColorRow (vbBlack)
       grdCuentasCTS.TextMatrix(i, 0) = i
       grdCuentasCTS.TextMatrix(i, 1) = xlHoja1.Range("A" & i + 7).value
       grdCuentasCTS.TextMatrix(i, 2) = xlHoja1.Range("B" & i + 7).value
       grdCuentasCTS.TextMatrix(i, 3) = xlHoja1.Range("C" & i + 7).value
       grdCuentasCTS.TextMatrix(i, 4) = xlHoja1.Range("D" & i + 7).value
       grdCuentasCTS.TextMatrix(i, 5) = xlHoja1.Range("E" & i + 7).value
       If Trim(xlHoja1.Range("A" & i + 7).value) = "" Or Trim(xlHoja1.Range("B" & i + 7).value) = "" Or Trim(xlHoja1.Range("C" & i + 7).value) = "" Or Trim(xlHoja1.Range("D" & i + 7).value) = "" Or Trim(xlHoja1.Range("E" & i + 7).value) = "" Then
           grdCuentasCTS.BackColorRow (&H80FF&)
           If Trim(xlHoja1.Range("A" & i + 7).value) = "" Then
                grdCuentasCTS.TextMatrix(i, 1) = "Error!"
                NumErrores = NumErrores + 1
           End If
           If Trim(xlHoja1.Range("A" & i + 8).value) = "" Then
                grdCuentasCTS.EliminaFila (i)
                NumErrores = NumErrores - 1
                FinRegistros = True
           End If
       End If
    Wend

    Set rsCuentas = Me.grdCuentasCTS.GetRsNew
    Set rsCtasInst = ClsMov.ObtenerCapCuentasCTSInstitucion(sNumRUC)
    Do While Not rsCuentas.EOF
        Valido = False
        Do While Not rsCtasInst.EOF
            If rsCuentas!Cuenta = rsCtasInst!cCtaCod Then
                Valido = True
                Exit Do
            End If
            rsCtasInst.MoveNext
        Loop
        If Valido = False Then
            grdCuentasCTS.BackColorRow (vbRed)
            grdCuentasCTS.ForeColorRow (vbWhite)
            NumErrores = NumErrores + 1
        End If
        rsCtasInst.MoveFirst
        rsCuentas.MoveNext
    Loop
    
    Set rsCtasInst = Nothing
    Set ClsMov = Nothing
    resp = vbYes
    cmdAbrir.Enabled = False
    cmdProcesar.Enabled = False
    If NumErrores > 0 Then
        MsgBox "Se identificaron errores en los datos contenidos en el archivo"
        Me.cmdGrabar.Enabled = False
        Call ExcelEnd(xlAplicacion, xlLibro, xlHoja1)
        Exit Sub
    Else
        Me.cmdGrabar.Enabled = True
    End If
    Call ExcelEnd(xlAplicacion, xlLibro, xlHoja1)
'ErrGraba:
'    MsgBox "El archivo no tiene el formato correcto", vbExclamation, "Error"
'    Exit Sub
End Sub

Private Sub cmdGrabar_Click()
    Dim ClsMov As NCOMCaptaMovimiento
    Dim clsMovN As COMNContabilidad.NCOMContFunciones
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim clsDef As NCOMCaptaDefinicion
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oGen As New COMDConstSistema.DCOMGeneral
    Dim rsCta As ADODB.Recordset
    
    Dim rsCtasInst As New ADODB.Recordset
    Dim sMovNro As String
    Dim nPorcDisp As Double
    Dim nExcedente As Double
    Dim nIntSaldo As Double
    Dim nSaldoRetiro As Double
    Dim dUltMov As Date
    Dim nTasa As Double
    Dim nDiasTranscurridos As Integer
    
    If Me.txtArchivo.Caption = "" Then
        MsgBox "No se seleccionó el archivo Excel a cargar", vbInformation, "Advertencia"
        Me.cmdAbrir.SetFocus
        Exit Sub
    End If

    If MsgBox("Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set ClsMov = New NCOMCaptaMovimiento
        Set clsMovN = New COMNContabilidad.NCOMContFunciones
        Set clsDef = New NCOMCaptaDefinicion
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        
        nPorcDisp = clsDef.GetCapParametro(gPorRetCTS)
        Set clsDef = Nothing
        
        sMovNro = clsMovN.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        rsCuentas.MoveFirst
        Dim nSaldoDisp As Double 'APRI20200330 POR COVID-19
        Dim nDU01 As Double 'APRI20200415 POR COVID-19
        Do While Not rsCuentas.EOF
            Set rsCta = clsMant.GetDatosCuentaCTS(rsCuentas!Cuenta)
            nSaldoRetiro = rsCta("nSaldRetiro")
            nTasa = rsCta("nTasaInteres")
            dUltMov = rsCta("dUltCierre")
            nSaldoDisp = rsCta("nSaldoDisp") * IIf(Mid(rsCuentas!Cuenta, 9, 1) = "1", 1, oGen.GetTipCambio(gdFecSis, TCFijoMes)) 'APRI20200330 POR COVID-19
            nDU01 = rsCta("nDU01") 'APRI20200415 POR COVID-19
            nDiasTranscurridos = DateDiff("d", dUltMov, gdFecSis) - 1
            If nDiasTranscurridos < 0 Then
                nDiasTranscurridos = 0
            End If
            nIntSaldo = ClsMov.GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)

            nExcedente = 0
            Call clsCap.AgregaDatosSueldosClientesCTS(sMovNro, rsCuentas!Cuenta, IIf(Trim(rsCuentas!MonedaSueldo) = "SOLES", 1, 2), rsCuentas!Total_4Sueldos) 'APRI20180504 INC1711070003 CAMBIO Total_6Sueldos -> Total_4Sueldos
            
            Set rsCta = clsCap.ObtenerCapSaldosCuentasCTS(rsCuentas!Cuenta, oGen.GetTipCambio(gdFecSis, TCFijoMes))
            If Not rsCta.EOF Or Not rsCta.BOF Then
                nExcedente = rsCta!TotalSaldos - rsCta!TotalSueldos
                If nExcedente > 0 Then
                    nSaldoRetiro = nExcedente * nPorcDisp / 100
                Else
                    nSaldoRetiro = 0
                End If
                'APRI20200330 CULPA DEL COVID-19
                If gdFecSis <= "2020-04-12" Then
                    nSaldoRetiro = nSaldoRetiro + IIf(nSaldoDisp < 2400, nSaldoDisp, 2400)
                End If
                'END APRI
            End If
            'clsCap.ActualizaSaldoRetiroCTS rsCuentas!Cuenta, nSaldoRetiro, nIntSaldo
            clsCap.ActualizaSaldoRetiroCTS rsCuentas!Cuenta, nSaldoRetiro, nIntSaldo, nDU01 'APRI20200415 POR COVID-19
            rsCuentas.MoveNext
        Loop
        MsgBox "El registro se realizó de forma existosa!", vbInformation, "Aviso"
        Call cmdCancelar_Click
    End If
    Set rsCuentas = Nothing
    Set clsDef = Nothing
    Set ClsMov = Nothing
    Set clsMovN = Nothing
    Set clsCap = Nothing
End Sub


