VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCapCambioTasaCTSLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Tasa CTS"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11535
   Icon            =   "frmCapCambioTasaCTSLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptMoneda 
      Caption         =   "ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   7320
      TabIndex        =   27
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton OptMoneda 
      Caption         =   "MN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6600
      TabIndex        =   26
      Top             =   120
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   15
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCambiaTasa 
      Caption         =   "Cambiar Tasa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame fraTasasProd 
      Caption         =   " Tasa por Producto "
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   8535
      Begin VB.TextBox txtCTSCajaSueldo 
         Alignment       =   1  'Right Justify
         Height          =   325
         Left            =   6840
         TabIndex        =   22
         Top             =   250
         Width           =   855
      End
      Begin VB.TextBox txtCTSNoActivo 
         Alignment       =   1  'Right Justify
         Height          =   325
         Left            =   3960
         TabIndex        =   19
         Top             =   250
         Width           =   855
      End
      Begin VB.TextBox txtCTSActivo 
         Alignment       =   1  'Right Justify
         Height          =   325
         Left            =   1200
         TabIndex        =   16
         Top             =   250
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "%"
         Height          =   255
         Left            =   7800
         TabIndex        =   24
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "CTS Caja Sueldo"
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "%"
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   310
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "CTS No Activo"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   310
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   310
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "CTS Activo"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   310
         Width           =   975
      End
   End
   Begin VB.CheckBox chkCargaManual 
      Caption         =   "Carga Manual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CheckBox chkCargaLote 
      Caption         =   "Carga Lote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame FraCargaManual 
      Enabled         =   0   'False
      Height          =   855
      Left            =   8640
      TabIndex        =   4
      Top             =   3480
      Width           =   2775
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   310
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   310
         Width           =   1095
      End
   End
   Begin VB.Frame FraCargaLote 
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   8415
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   310
         Width           =   1095
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   6600
         TabIndex        =   9
         Top             =   350
         Width           =   375
      End
      Begin VB.CommandButton cmdGeneraFormato 
         Caption         =   "Generar Formato"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "frmCapCambioTasaCTSLote.frx":030A
         TabIndex        =   7
         Top             =   310
         Width           =   1395
      End
      Begin VB.Label txtCarga 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   290
         Left            =   1560
         TabIndex        =   8
         Top             =   350
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cuentas "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   11295
      Begin SICMACT.FlexEdit feCambioTasas 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4683
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nº Cuenta-Nombre del cliente-Fecha Apertura-Saldo Actual-Producto-nTpoPrograma-TEA Act.-Nueva TEA"
         EncabezadosAnchos=   "300-1700-2700-1260-1100-1600-0-1000-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-R-L-L-R-R"
         FormatosEdit    =   "0-0-0-0-2-0-0-2-2"
         CantEntero      =   12
         CantDecimales   =   4
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CheckBox chkCambioCtasVig 
      Caption         =   "Cambio a todas Ctas. Vigentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar pgbExcel 
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "frmCapCambioTasaCTSLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCapCambioTasaCTSLote
'** Descripción : Formulario para cambiar tasas de cuentas CTS en Lote según TI-ERS013-2014
'** Creación : JUEZ, 20140228 06:00:00 PM
'*****************************************************************************************************

Option Explicit

Dim oNCapGen As COMNCaptaGenerales.NCOMCaptaGenerales
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rsDatos As ADODB.Recordset
Dim objPista As COMManejador.Pista
Dim nIndex As Integer

Private Sub chkCambioCtasVig_Click()
    If chkCambioCtasVig.value = 1 Then
        feCambioTasas.Enabled = False
        Call LimpiaFlex(feCambioTasas)
        FraCargaLote.Enabled = False
        fraCargaManual.Enabled = False
        chkCargaManual.value = 0
        chkCargaLote.value = 0
        txtCarga.Caption = ""
    End If
End Sub

Private Sub chkCargaLote_Click()
    If chkCargaLote.value = 1 Then
        feCambioTasas.Enabled = True
        Call LimpiaFlex(feCambioTasas)
        FraCargaLote.Enabled = True
        fraCargaManual.Enabled = False
        chkCargaManual.value = 0
        chkCambioCtasVig.value = 0
    Else
        FraCargaLote.Enabled = False
    End If
End Sub

Private Sub chkCargaManual_Click()
    If chkCargaManual.value = 1 Then
        feCambioTasas.Enabled = True
        Call LimpiaFlex(feCambioTasas)
        FraCargaLote.Enabled = False
        fraCargaManual.Enabled = True
        chkCargaLote.value = 0
        chkCambioCtasVig.value = 0
        txtCarga.Caption = ""
    Else
        fraCargaManual.Enabled = False
    End If
End Sub

Private Sub cmdAgregar_Click()
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
Dim loCuentas As COMDPersona.UCOMProdPersona
Dim nTasaNueva As Double
Dim i As Integer

    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set R = oDCapGen.GetCuentasPersona(oPers.sPersCod, gCapCTS, True, , IIf(optmoneda(1).value, gMonedaNacional, gMonedaExtranjera), , , , True)
        Set oDCapGen = Nothing
        
        If R.RecordCount > 0 Then
            Set loCuentas = New COMDPersona.UCOMProdPersona
            Set loCuentas = frmProdPersona.Inicio(oPers.sPersNombre, R)
            If loCuentas.sCtaCod <> "" Then
                For i = 1 To feCambioTasas.Rows - 1
                    If Mid(loCuentas.sCtaCod, 1, 18) = feCambioTasas.TextMatrix(i, 1) Then
                        MsgBox "La cuenta CTS seleccionada ya se encuentra en lista", vbInformation, "Aviso"
                        Exit Sub
                    End If
                Next i
                Set oNCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
                Set rsDatos = oNCapGen.GetDatosCuenta(Mid(loCuentas.sCtaCod, 1, 18))
                nTasaNueva = 0
                feCambioTasas.AdicionaFila
                feCambioTasas.TextMatrix(feCambioTasas.row, 1) = rsDatos!cCtaCod
                feCambioTasas.TextMatrix(feCambioTasas.row, 2) = oNCapGen.GetPersonaCuenta(rsDatos!cCtaCod, gCapRelPersTitular)!Nombre
                feCambioTasas.TextMatrix(feCambioTasas.row, 3) = Format(rsDatos!dApertura, "dd/MM/yyyy")
                feCambioTasas.TextMatrix(feCambioTasas.row, 4) = Format(rsDatos!nSaldo, "#,##0.00")
                feCambioTasas.TextMatrix(feCambioTasas.row, 5) = rsDatos!cTpoPrograma
                feCambioTasas.TextMatrix(feCambioTasas.row, 6) = rsDatos!nTpoPrograma
                feCambioTasas.TextMatrix(feCambioTasas.row, 7) = Format(rsDatos!nTEA, "#,##0.00")
                If rsDatos!nTpoPrograma = 0 Then
                    If val(Trim(txtCTSCajaSueldo.Text)) Then nTasaNueva = CDbl(Trim(txtCTSCajaSueldo.Text))
                ElseIf rsDatos!nTpoPrograma = 1 Then
                    If val(Trim(txtCTSActivo.Text)) Then nTasaNueva = CDbl(Trim(txtCTSActivo.Text))
                ElseIf rsDatos!nTpoPrograma = 2 Then
                    If val(Trim(txtCTSNoActivo.Text)) Then nTasaNueva = CDbl(Trim(txtCTSNoActivo.Text))
                End If
                feCambioTasas.TextMatrix(feCambioTasas.row, 8) = Format(nTasaNueva, "#,##0.00")
                Set oNCapGen = Nothing
            End If
            Set loCuentas = Nothing
        Else
            MsgBox "El cliente no tiene cuentas CTS activas en la moneda seleccionada", vbInformation, "Aviso"
        End If
    End If
    Set oPers = Nothing
End Sub

Private Sub cmdCambiaTasa_Click()
    If ValidaDatos Then
        If chkCambioCtasVig.value = 0 Then
            If MsgBox("Se grabará la información registrada, desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
        Else
            If MsgBox("Se actualizar las tasas de todas las cuentas vigentes, desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
        End If
        
        Dim oMov As COMDMov.DCOMMov
        Dim VCMovNro As String
        Dim nTNAAnt As Double, nTNA As Double
        Dim i As Integer
        Set oNCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set oMov = New COMDMov.DCOMMov
        Set objPista = New COMManejador.Pista
        
        gsOpeCod = gCTSCambioTasaInteres
        
        VCMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        If chkCambioCtasVig.value = 1 Then
            oNCapGen.CambioTasaCTSCtasVig IIf(optmoneda(1).value, gMonedaNacional, gMonedaExtranjera), VCMovNro, ConvierteTEAaTNA(CDbl(txtCTSCajaSueldo.Text)), ConvierteTEAaTNA(CDbl(txtCTSActivo.Text)), ConvierteTEAaTNA(CDbl(txtCTSNoActivo))
        Else
            For i = 1 To feCambioTasas.Rows - 1
                nTNAAnt = ConvierteTEAaTNA(CDbl(feCambioTasas.TextMatrix(i, 7)))
                nTNA = ConvierteTEAaTNA(CDbl(feCambioTasas.TextMatrix(i, 8)))
                oNCapGen.CambioTasaCTS feCambioTasas.TextMatrix(i, 1), nTNAAnt, nTNA, VCMovNro
            Next i
        End If
        Set oNCapGen = Nothing
        Set oMov = Nothing
        
        objPista.InsertarPista gsOpeCod, VCMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Cambio Tasa"
        MsgBox "Los datos fueron actualizados con éxito", vbInformation, "Aviso"
        If MsgBox("Desea emitir un reporte con los cambios realizados?", vbYesNo, "Aviso") = vbYes Then
            ReporteCambioTasas VCMovNro
        End If
        Limpiar
    End If
End Sub

Private Sub Limpiar()
    Call LimpiaFlex(feCambioTasas)
    chkCambioCtasVig.value = 0
    chkCargaLote.value = 0
    chkCargaManual.value = 0
    txtCTSActivo.Text = ""
    txtCTSNoActivo.Text = ""
    txtCTSCajaSueldo.Text = ""
End Sub

Private Sub cmdCargar_Click()
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim cNombreHoja As String
Dim i As Long, n As Long, j As Integer
Dim pbExisteHoja As Boolean
Dim MatErrores As Variant, nErrores As Integer
Dim lnMoneda As Integer

Call LimpiaFlex(feCambioTasas)
lnMoneda = IIf(optmoneda(1).value, 1, 2)

If Trim(txtCarga.Caption) = "" Then
    MsgBox "Debe indicar la ruta del Archivo Excel", vbInformation, "Mensaje"
    Exit Sub
Else
    Set xlApp = New Excel.Application
    pgbExcel.value = 0
    pgbExcel.Min = 0
    Set xlLibro = xlApp.Workbooks.Open(CommonDialog1.Filename, True, True, , "")
    cNombreHoja = "Formato"
    For Each xlHoja In xlLibro.Worksheets
        If xlHoja.Name = cNombreHoja Then
            pbExisteHoja = True
            Exit For
        End If
    Next
    If pbExisteHoja = False Then
        MsgBox "No existe ninguna hoja con nombre 'Formato'", vbInformation, "Aviso"
        Exit Sub
    End If
    'validar nombre de hoja
    Set xlHoja = xlApp.Worksheets(cNombreHoja)
    varMatriz = xlHoja.Range("A1:A65536").value
    xlLibro.Close SaveChanges:=False
    xlApp.Quit
    Set xlHoja = Nothing
    Set xlLibro = Nothing
    Set xlApp = Nothing
    
    For i = 2 To UBound(varMatriz)
        If Trim(CStr(varMatriz(i, 1))) = "" Then Exit For
        n = n + 1
    Next i
    If n = 0 Then
        MsgBox "No hay datos para la carga", vbInformation, "Aviso"
        Exit Sub
    End If
    
    pgbExcel.Max = n
    
    If varMatriz(1, 1) <> "Nº Cuenta" Then
        MsgBox "Archivo No tiene Estructura Correcta, la cabecera 'Nº Cuenta' debe estar en la fila A:1", vbCritical, "Mensaje"
        Exit Sub
    End If
    
    Set oNCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    nErrores = 0
    ReDim MatErrores(n, 2)
    For i = 1 To n
        If Trim(CStr(varMatriz(i + 1, 1))) = "" Then Exit For
        Set rsDatos = oNCapGen.GetDatosCuenta(CStr(varMatriz(i + 1, 1)))
        If rsDatos.EOF Or rsDatos.BOF Then
            MatErrores(nErrores, 1) = CStr(varMatriz(i + 1, 1))
            MatErrores(nErrores, 2) = "Cuenta no existe"
            nErrores = nErrores + 1
        ElseIf Mid(CStr(varMatriz(i + 1, 1)), 6, 3) <> "234" Then
            MatErrores(nErrores, 1) = CStr(varMatriz(i + 1, 1))
            MatErrores(nErrores, 2) = "Cuenta no es CTS"
            nErrores = nErrores + 1
        ElseIf rsDatos!nPrdEstado = gCapEstAnulada Or rsDatos!nPrdEstado = gCapEstCancelada Then
            MatErrores(nErrores, 1) = CStr(varMatriz(i + 1, 1))
            MatErrores(nErrores, 2) = "Cuenta està anulada o cancelada"
            nErrores = nErrores + 1
        ElseIf Mid(CStr(varMatriz(i + 1, 1)), 9, 1) <> lnMoneda Then
            MatErrores(nErrores, 1) = CStr(varMatriz(i + 1, 1))
            MatErrores(nErrores, 2) = "Cuenta en moneda diferente"
            nErrores = nErrores + 1
        Else
            For j = 1 To n
                If i <> j Then
                    If CStr(varMatriz(i + 1, 1)) = CStr(varMatriz(j + 1, 1)) Then
                        MatErrores(nErrores, 1) = CStr(varMatriz(i + 1, 1))
                        MatErrores(nErrores, 2) = "Cuenta duplicada"
                        nErrores = nErrores + 1
                        Exit For
                    End If
                End If
            Next j
        End If

        pgbExcel.value = pgbExcel.value + 1
    Next i
    
    If nErrores > 0 Then
        If MsgBox("No se puede realizar la carga porque el archivo de carga en lote tiene errores, ¿Desea ver los errores encontrados?", vbYesNo, "Aviso") = vbYes Then
            MuestraErroresExcel MatErrores
        End If
    Else
        Dim nTasaNueva As Double
        For i = 1 To n
            Set rsDatos = oNCapGen.GetDatosCuenta(CStr(varMatriz(i + 1, 1)))
            nTasaNueva = 0
            feCambioTasas.AdicionaFila
            feCambioTasas.TextMatrix(i, 1) = rsDatos!cCtaCod
            feCambioTasas.TextMatrix(i, 2) = oNCapGen.GetPersonaCuenta(rsDatos!cCtaCod, gCapRelPersTitular)!Nombre
            feCambioTasas.TextMatrix(i, 3) = Format(rsDatos!dApertura, "dd/MM/yyyy")
            feCambioTasas.TextMatrix(i, 4) = Format(rsDatos!nSaldo, "#,##0.00")
            feCambioTasas.TextMatrix(i, 5) = rsDatos!cTpoPrograma
            feCambioTasas.TextMatrix(i, 6) = rsDatos!nTpoPrograma
            feCambioTasas.TextMatrix(i, 7) = Format(rsDatos!nTEA, "#,##0.00")
            If rsDatos!nTpoPrograma = 0 Then
                If val(Trim(txtCTSCajaSueldo.Text)) Then nTasaNueva = CDbl(Trim(txtCTSCajaSueldo.Text))
            ElseIf rsDatos!nTpoPrograma = 1 Then
                If val(Trim(txtCTSActivo.Text)) Then nTasaNueva = CDbl(Trim(txtCTSActivo.Text))
            ElseIf rsDatos!nTpoPrograma = 2 Then
                If val(Trim(txtCTSNoActivo.Text)) Then nTasaNueva = CDbl(Trim(txtCTSNoActivo.Text))
            End If
            feCambioTasas.TextMatrix(i, 8) = Format(nTasaNueva, "#,##0.00")
        Next i
        feCambioTasas.row = 1
        txtCarga.Caption = ""
        MsgBox "Carga realizada con éxito", vbInformation, "Carga Finalizada"
    End If

    pgbExcel.value = 0
    pgbExcel.Min = 0
End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGeneraFormato_Click()
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
Dim sColumna As String
    lsArchivo = App.path & "\SPOOLER\FormatoCambioTasasLote_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbLibroOpen Then
        Exit Sub
    End If
    nLin = 1
    lsHoja = "Formato"
    gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
    
    xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 18
    xlHoja1.Range("A1:A1").ColumnWidth = 20
    xlHoja1.Range("B1:B1").ColumnWidth = 40
    xlHoja1.Range("C1:C1").ColumnWidth = 15
    
    xlHoja1.Cells(nLin, 1) = "Nº Cuenta"
    xlHoja1.Cells(nLin, 2) = "Cliente"
    xlHoja1.Cells(nLin, 3) = "DOI"
    
    xlHoja1.Range("A" & nLin & ":C" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":C" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":C" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":C" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":C" & nLin).Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("A" & nLin & ":C" & nLin).Interior.Color = RGB(255, 50, 50)
    xlHoja1.Range("A" & nLin & ":C" & nLin).Font.Color = RGB(255, 255, 255)
    
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
    
    gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    gFunGeneral.CargaArchivo lsArchivo, App.path & "\SPOOLER\"
End Sub

Private Sub cmdLoad_Click()
    CommonDialog1.Filter = "Archivos de Excel (*.xls)|*.xls|Archivos de Excel 2007 (*.xlsx)|*.xlsx|Todos los Archivo (*.*)|*.*"
    CommonDialog1.ShowOpen
    txtCarga.Caption = Replace(CommonDialog1.Filename, " ", "_")
End Sub

Private Sub cmdQuitar_Click()
    If feCambioTasas.TextMatrix(feCambioTasas.row, 0) = "" Then
        MsgBox "Debe seleccionar al menos un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feCambioTasas.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        feCambioTasas.EliminaFila feCambioTasas.row
    End If
End Sub

Private Sub OptMoneda_Click(Index As Integer)
    If nIndex <> Index Then
        If MsgBox("Al cambiar el tipo de moneda se descartarán las cuentas actualmente seleccionadas. Desea continuar?", vbYesNo, "Aviso") = vbYes Then
            Call LimpiaFlex(feCambioTasas)
            txtCTSActivo.Text = ""
            txtCTSCajaSueldo.Text = ""
            txtCTSNoActivo.Text = ""
        Else
            optmoneda(Index).value = False
            nIndex = IIf(Index = 1, 2, 1)
            optmoneda(nIndex).value = True
            nIndex = 0
        End If
    End If
End Sub

Private Sub txtCTSActivo_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtCTSNoActivo.SetFocus
    End If
End Sub

Private Sub txtCTSActivo_LostFocus()
    txtCTSActivo.Text = Format(txtCTSActivo.Text, "#,##0.00")
    If val(Trim(txtCTSActivo.Text)) Then
        Call InsertaTasaNueva(CDbl(IIf(Trim(txtCTSActivo.Text) = "", 0, txtCTSActivo.Text)), 1)
    End If
End Sub

Private Sub txtCTSCajaSueldo_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdCambiaTasa.SetFocus
    End If
End Sub

Private Sub txtCTSCajaSueldo_LostFocus()
    txtCTSCajaSueldo.Text = Format(txtCTSCajaSueldo.Text, "#,##0.00")
    If val(Trim(txtCTSCajaSueldo.Text)) Then
        Call InsertaTasaNueva(CDbl(IIf(Trim(txtCTSCajaSueldo.Text) = "", 0, txtCTSCajaSueldo.Text)), 0)
    End If
End Sub

Private Sub txtCTSNoActivo_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtCTSCajaSueldo.SetFocus
    End If
End Sub

Private Function ValidaDatos() As Boolean
Dim i As Integer

    ValidaDatos = False
    
    If val(Trim(txtCTSActivo.Text)) = 0 Or val(Trim(txtCTSNoActivo.Text)) = 0 Or val(Trim(txtCTSCajaSueldo.Text)) = 0 Then
        MsgBox "Es necesario ingresar las nuevas tasas de los 3 productos CTS", vbInformation, "Aviso"
        txtCTSActivo.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If chkCambioCtasVig.value = 0 Then
        If ValidaFlexVacio Then
            MsgBox "No hay datos en la lista para actualizar", vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If

        For i = 1 To feCambioTasas.Rows - 1
            If CDbl(feCambioTasas.TextMatrix(i, 8)) = 0 Then
                MsgBox "La cuenta " & feCambioTasas.TextMatrix(i, 1) & " no tiene TEA nueva definida", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Next i
    End If
    
    ValidaDatos = True
End Function

Private Sub MuestraErroresExcel(ByVal pMatErrores As Variant)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
Dim sColumna As String
Dim i As Integer
    lsArchivo = App.path & "\SPOOLER\Errores_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbLibroOpen Then
        Exit Sub
    End If
    nLin = 1
    lsHoja = "Hoja1"
    gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
    
    xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 18
    xlHoja1.Range("A1:A1").ColumnWidth = 20
    xlHoja1.Range("B1:B1").ColumnWidth = 40
    
    xlHoja1.Cells(nLin, 1) = "Nº Cuenta"
    xlHoja1.Cells(nLin, 2) = "Error"
    
    xlHoja1.Range("A" & nLin & ":B" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":B" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":B" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":B" & nLin).Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("A" & nLin & ":B" & nLin).Interior.Color = RGB(255, 50, 50)
    xlHoja1.Range("A" & nLin & ":B" & nLin).Font.Color = RGB(255, 255, 255)
    
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
    
    nLin = nLin + 1
    For i = 0 To UBound(pMatErrores)
        xlHoja1.Range("A" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignLeft
        xlHoja1.Cells(nLin, 1) = "'" & pMatErrores(i, 1)
        xlHoja1.Cells(nLin, 2) = pMatErrores(i, 2)
        nLin = nLin + 1
    Next i
    
    gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    gFunGeneral.CargaArchivo lsArchivo, App.path & "\SPOOLER\"
End Sub

Public Function ValidaFlexVacio() As Boolean
    Dim i As Integer
    For i = 1 To feCambioTasas.Rows - 1
        If feCambioTasas.TextMatrix(i, 1) = "" Then
            ValidaFlexVacio = True
            Exit Function
        End If
    Next i
End Function

Private Sub txtCTSNoActivo_LostFocus()
    txtCTSNoActivo.Text = Format(txtCTSNoActivo.Text, "#,##0.00")
    If val(Trim(txtCTSNoActivo.Text)) Then
        Call InsertaTasaNueva(CDbl(IIf(Trim(txtCTSNoActivo.Text) = "", 0, txtCTSNoActivo.Text)), 2)
    End If
End Sub

Private Sub InsertaTasaNueva(ByVal pnTasaNueva As Double, ByVal pnTpoPrograma As Integer)
    If Not ValidaFlexVacio Then
        Dim i As Integer
        For i = 1 To feCambioTasas.Rows - 1
            If CInt(feCambioTasas.TextMatrix(i, 6)) = pnTpoPrograma Then
                feCambioTasas.TextMatrix(i, 8) = Format(pnTasaNueva, "#,##0.00")
            End If
        Next i
    End If
End Sub

Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

Private Sub ReporteCambioTasas(ByVal psMovNroCambio As String)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
Dim sColumna As String
Dim i As Integer
    lsArchivo = App.path & "\SPOOLER\CambioTasasRealizada_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbLibroOpen Then
        Exit Sub
    End If
    nLin = 1
    lsHoja = "Hoja1"
    gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
    
    pgbExcel.value = 0
    pgbExcel.Min = 0
    
    xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 18
    xlHoja1.Range("A1:A1").ColumnWidth = 20
    xlHoja1.Range("B1:B1").ColumnWidth = 50
    xlHoja1.Range("C1:C1").ColumnWidth = 15
    xlHoja1.Range("D1:D1").ColumnWidth = 12
    xlHoja1.Range("E1:E1").ColumnWidth = 20
    xlHoja1.Range("F1:F1").ColumnWidth = 10
    xlHoja1.Range("G1:G1").ColumnWidth = 10
    
    xlHoja1.Cells(nLin, 1) = "Nº Cuenta"
    xlHoja1.Cells(nLin, 2) = "Nombre del Cliente"
    xlHoja1.Cells(nLin, 3) = "Fecha Apertura"
    xlHoja1.Cells(nLin, 4) = "Saldo Actual"
    xlHoja1.Cells(nLin, 5) = "Producto"
    xlHoja1.Cells(nLin, 6) = "TEA Ant"
    xlHoja1.Cells(nLin, 7) = "TEA Nueva"
    
    xlHoja1.Range("A" & nLin & ":G" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":G" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":G" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":G" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":G" & nLin).Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("A" & nLin & ":G" & nLin).Interior.Color = RGB(255, 50, 50)
    xlHoja1.Range("A" & nLin & ":G" & nLin).Font.Color = RGB(255, 255, 255)
    
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
    
    nLin = nLin + 1
        
    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsDatos = oDCapGen.RecuperaDatosCambioTasas(psMovNroCambio)
    Set oDCapGen = Nothing
    
    pgbExcel.Max = rsDatos.RecordCount
    
    Do While Not rsDatos.EOF
        xlHoja1.Range("B" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignLeft
        xlHoja1.Cells(nLin, 1) = "'" & rsDatos!cCtaCod
        xlHoja1.Cells(nLin, 2) = rsDatos!cPersNombre
        xlHoja1.Cells(nLin, 3) = rsDatos!dFechaApert
        xlHoja1.Cells(nLin, 4) = "'" & Format(rsDatos!nSaldo, "#,##0.00")
        xlHoja1.Range("E" & nLin & ":E" & nLin).HorizontalAlignment = xlHAlignLeft
        xlHoja1.Cells(nLin, 5) = rsDatos!cTpoPrograma
        xlHoja1.Cells(nLin, 6) = "'" & Format(rsDatos!nTasaAnt, "#,##0.00") & " % "
        xlHoja1.Cells(nLin, 7) = "'" & Format(rsDatos!nTasaCambio, "#,##0.00") & " % "
        nLin = nLin + 1
        pgbExcel.value = pgbExcel.value + 1
        rsDatos.MoveNext
    Loop
    Set rsDatos = Nothing
    
    gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    gFunGeneral.CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    
    pgbExcel.value = 0
    pgbExcel.Min = 0
End Sub
