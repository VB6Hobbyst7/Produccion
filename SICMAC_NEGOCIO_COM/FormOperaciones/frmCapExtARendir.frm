VERSION 5.00
Begin VB.Form frmCapExtARendir 
   Caption         =   "Extono Retiro de Fondo Fijo"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   Icon            =   "frmCapExtARendir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmMotExtorno 
      Caption         =   "Motivos del Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2700
      Left            =   3720
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   2850
      Begin VB.CommandButton cmdExtContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   860
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCapExtARendir.frx":030A
         Left            =   240
         List            =   "frmCapExtARendir.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdExtonar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtGlosa 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   9975
      Begin SICMACT.FlexEdit FEExtornarARendir 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9735
         _extentx        =   17171
         _extenty        =   4260
         cols0           =   6
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "Nro-Movimiento-Usuario-Colaborador-Moneda-Importe"
         encabezadosanchos=   "500-1200-1200-4000-1200-1200"
         font            =   "frmCapExtARendir.frx":030E
         font            =   "frmCapExtARendir.frx":0332
         font            =   "frmCapExtARendir.frx":0356
         font            =   "frmCapExtARendir.frx":037A
         font            =   "frmCapExtARendir.frx":039E
         fontfixed       =   "frmCapExtARendir.frx":03C2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0"
         encabezadosalineacion=   "C-R-C-L-L-R"
         formatosedit    =   "0-0-0-0-0-0"
         textarray0      =   "Nro"
         colwidth0       =   495
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forma de Busqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtMovimiento 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optBusqueda 
         Caption         =   "&Usuario"
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
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optBusqueda 
         Caption         =   "&Número de Movimiento"
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
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmCapExtARendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'************************************************************************************************
'***Nombre      : frmCapExtARendir
'***Descripción : Formulario para Extornar Desembolsar de las Solicitudes de los Váticos y Otros Gastos A Rendir.
'***Creación    : ELRO el 20120423, según OYP-RFC005-2012 y OYP-RFC016-2012 y OYP-RFC047-2012
'************************************************************************************************

Dim fsCodExtOpc As String

Public Sub iniciarExtornoDesembolso(ByVal pnOpeCodAExt As CaptacOperacion, _
                                    ByVal psTitulo As String, _
                                    ByVal psOpeCod As CaptacOperacion)
                   
If psOpeCod = gOtrOpeExtDesParGas Then
    Me.Caption = psTitulo
    fsCodExtOpc = pnOpeCodAExt
ElseIf psOpeCod = gOtrOpeExtDesParVia Then
    Me.Caption = psTitulo
    fsCodExtOpc = pnOpeCodAExt
ElseIf psOpeCod = gOtrOpeExtDesParCaj Then
    Me.Caption = psTitulo
    fsCodExtOpc = pnOpeCodAExt
End If
Show 1
End Sub

Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    'Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

Private Sub buscarViaticos()
Dim oNCOMCaptaMovimiento As COMNCaptaGenerales.NCOMCaptaMovimiento
Set oNCOMCaptaMovimiento = New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim rsDesembolsoViaticos As ADODB.Recordset
Set rsDesembolsoViaticos = New ADODB.Recordset

Call LimpiaFlex(FEExtornarARendir)
    
Set rsDesembolsoViaticos = oNCOMCaptaMovimiento.devolverDesembolsoAprobacionARendirViaticos(IIf(optBusqueda.item(0).value = False, 0, txtMovimiento), IIf(optBusqueda.item(0).value = False, Trim(UCase(txtMovimiento)), ""), Format(gdFecSis, "yyyyMMdd"), gsCodAge)

If Not rsDesembolsoViaticos.BOF Or Not rsDesembolsoViaticos.EOF Then
    FEExtornarARendir.lbEditarFlex = True
    Do While Not rsDesembolsoViaticos.EOF
        FEExtornarARendir.AdicionaFila
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1) = rsDesembolsoViaticos!nMovNro
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 2) = UCase(rsDesembolsoViaticos!cUser)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 3) = UCase(rsDesembolsoViaticos!cPersNombre)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 4) = UCase(rsDesembolsoViaticos!cmoneda)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 5) = Format(rsDesembolsoViaticos!nMovImporte, "#,##0.00")
        rsDesembolsoViaticos.MoveNext
    Loop
    FEExtornarARendir.lbEditarFlex = False
Else
    MsgBox "No existe movimiento", vbInformation, "Aviso"
End If

Set rsDesembolsoViaticos = Nothing
Set oNCOMCaptaMovimiento = Nothing
End Sub

Private Sub buscarARendirCuentas()
Dim oNCOMCaptaMovimiento As COMNCaptaGenerales.NCOMCaptaMovimiento
Set oNCOMCaptaMovimiento = New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim rsDesembolso As ADODB.Recordset
Set rsDesembolso = New ADODB.Recordset

Call LimpiaFlex(FEExtornarARendir)
    
Set rsDesembolso = oNCOMCaptaMovimiento.devolverDesembolsoAprobacionARendirCuentas(IIf(optBusqueda.item(0).value = False, 0, txtMovimiento), IIf(optBusqueda.item(0).value = False, Trim(UCase(txtMovimiento)), ""), Format(gdFecSis, "yyyyMMdd"), gsCodAge)

If Not rsDesembolso.BOF Or Not rsDesembolso.EOF Then
    FEExtornarARendir.lbEditarFlex = True
    Do While Not rsDesembolso.EOF
        FEExtornarARendir.AdicionaFila
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1) = rsDesembolso!nMovNro
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 2) = UCase(rsDesembolso!cUser)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 3) = UCase(rsDesembolso!cPersNombre)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 4) = UCase(rsDesembolso!cmoneda)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 5) = Format(rsDesembolso!nMovImporte, "#,##0.00")
        rsDesembolso.MoveNext
    Loop
    FEExtornarARendir.lbEditarFlex = False
Else
    MsgBox "No existe movimiento", vbInformation, "Aviso"
End If

Set rsDesembolso = Nothing
Set oNCOMCaptaMovimiento = Nothing
End Sub

Private Sub buscarCajaChica()
Dim oNCOMCaptaMovimiento As COMNCaptaGenerales.NCOMCaptaMovimiento
Set oNCOMCaptaMovimiento = New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim rsDesembolso As ADODB.Recordset
Set rsDesembolso = New ADODB.Recordset

Call LimpiaFlex(FEExtornarARendir)

Set rsDesembolso = oNCOMCaptaMovimiento.devolverDesembolsoAprobacionCH(IIf(optBusqueda.item(0).value = False, 0, txtMovimiento), IIf(optBusqueda.item(0).value = False, Trim(UCase(txtMovimiento)), ""), Format(gdFecSis, "yyyyMMdd"), gsCodAge)

If Not rsDesembolso.BOF Or Not rsDesembolso.EOF Then
    FEExtornarARendir.lbEditarFlex = True
    Do While Not rsDesembolso.EOF
        FEExtornarARendir.AdicionaFila
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1) = rsDesembolso!nMovNro
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 2) = UCase(rsDesembolso!cUser)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 3) = UCase(rsDesembolso!cPersNombre)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 4) = UCase(rsDesembolso!cmoneda)
        FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 5) = Format(rsDesembolso!nMovImporte, "#,##0.00")
        rsDesembolso.MoveNext
    Loop
    FEExtornarARendir.lbEditarFlex = False
Else
    MsgBox "No existe movimiento", vbInformation, "Aviso"
End If

Set rsDesembolso = Nothing
Set oNCOMCaptaMovimiento = Nothing
End Sub

Private Sub cmdBuscar_Click()
Dim lsmensaje As String

lsmensaje = IIf(optBusqueda.item(0).value, "Falta ingresar el Nro Movimiento", "Falta ingresar el Usuario")

If Trim(txtMovimiento) = "" Then
    MsgBox lsmensaje, vbInformation, "Aviso"
    txtMovimiento.SetFocus
    Exit Sub
End If

If optBusqueda.item(0).value = True And Not IsNumeric(txtMovimiento) Then
    MsgBox "Ingrese Nro de Operaciones", vbInformation, "Aviso"
    txtMovimiento.SetFocus
    Exit Sub
End If

If optBusqueda.item(0).value = False And IsNumeric(txtMovimiento) Then
    MsgBox "Ingrese un Usuario", vbInformation, "Aviso"
    txtMovimiento.SetFocus
    Exit Sub
End If

If gsOpeCod = gOtrOpeExtDesParGas Then
    Call buscarARendirCuentas
ElseIf gsOpeCod = gOtrOpeExtDesParVia Then
    Call buscarViaticos
ElseIf gsOpeCod = gOtrOpeExtDesParCaj Then
    Call buscarCajaChica
End If

End Sub
'***CTI3 (ferimoro)    11/10/2018
Sub limpExtorno()
Frame1.Enabled = True
cmdExtonar.Enabled = True
FEExtornarARendir.Enabled = True
frmMotExtorno.Visible = False

End Sub
Private Sub cmdExtContinuar_Click()
Dim oVistoElectronico As frmVistoElectronico
Set oVistoElectronico = New frmVistoElectronico
Dim lbResultadoVisto As Boolean

    If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
        MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'cti3
    Dim DatosExtorna(1) As String
    DatosExtorna(0) = cmbMotivos.Text
    DatosExtorna(1) = txtDetExtorno.Text
    '********************
    
' *** RIRO SEGUN TI-ERS108-2013 ***
    Dim nMovNroOperacion As Long
    nMovNroOperacion = 0
    If FEExtornarARendir.row >= 1 And Len(Trim(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1))) > 0 Then
        nMovNroOperacion = Val(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1))
    End If
' *** FIN RIRO ***

lbResultadoVisto = oVistoElectronico.Inicio(3, fsCodExtOpc, , , nMovNroOperacion) 'RIRO SEGUN TI-ERS108-2013/ Se agrego parametro nMovNroOperacion
If Not lbResultadoVisto Then
    Call limpExtorno  'CTI3
    Exit Sub
End If

If MsgBox("¿Esta seguro que desea eliminar el Nro de Operación: " & FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1) & "?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    Dim oNCOMCaptaMovimiento As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oNCOMCaptaMovimiento = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim lsBoleta As String, lsMovNro As String
    
    oNCOMCaptaMovimiento.IniciaImpresora gImpresora
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If gsOpeCod = gOtrOpeExtDesParGas Then
        Call oNCOMCaptaMovimiento.extornarDesembolsoAprobadoGastos(CLng(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1)), _
                                                                    lsMovNro, _
                                                                    gsOpeCod, _
                                                                    UCase(txtGlosa), _
                                                                    IIf(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 4) = "SOLES", 1, 2), _
                                                                    FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 5), _
                                                                    FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 3), _
                                                                    gsNomAge, _
                                                                    gsCodCMAC, _
                                                                    lsBoleta, _
                                                                    sLpt, _
                                                                    gbImpTMU, DatosExtorna)
    ElseIf gsOpeCod = gOtrOpeExtDesParVia Then
        Call oNCOMCaptaMovimiento.extornarDesembolsoAprobadoViatico(CLng(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1)), _
                                                                    lsMovNro, _
                                                                    gsOpeCod, _
                                                                    UCase(txtGlosa), _
                                                                    IIf(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 4) = "SOLES", 1, 2), _
                                                                    FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 5), _
                                                                    FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 3), _
                                                                    gsNomAge, _
                                                                    gsCodCMAC, _
                                                                    lsBoleta, _
                                                                    sLpt, _
                                                                    gbImpTMU, DatosExtorna)
    ElseIf gsOpeCod = gOtrOpeExtDesParCaj Then
        Call oNCOMCaptaMovimiento.extornarDesembolsoAprobadoCH(CLng(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1)), _
                                                                lsMovNro, _
                                                                gsOpeCod, _
                                                                UCase(txtGlosa), _
                                                                IIf(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 4) = "SOLES", 1, 2), _
                                                                FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 5), _
                                                                FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 3), _
                                                                gsNomAge, _
                                                                gsCodCMAC, _
                                                                lsBoleta, _
                                                                sLpt, _
                                                                gbImpTMU, DatosExtorna)
    
    End If
    
    
    oVistoElectronico.RegistraVistoElectronico (CLng(FEExtornarARendir.TextMatrix(FEExtornarARendir.row, 1)))
     
    If Trim(lsBoleta) <> "" Then ImprimeBoleta (lsBoleta)
    
    FEExtornarARendir.EliminaFila FEExtornarARendir.row
    txtGlosa = ""
    Set oNCOMContFunciones = Nothing
    Set oNCOMCaptaMovimiento = Nothing
    Set oVistoElectronico = Nothing
End If
End Sub
Private Sub cmdExtonar_Click()
'CTI3
If FEExtornarARendir.rows < 2 Then
    MsgBox "No existe movimiento para extornar", vbInformation, "Aviso"
    Exit Sub
End If
'If Len(Trim(txtGlosa)) = 0 Then
'    MsgBox "Por favor ingrese la glosa del extorno", vbInformation, "Aviso"
'
'    txtGlosa.SetFocus
'    Exit Sub
'End If
Frame1.Enabled = False
cmdExtonar.Enabled = False
FEExtornarARendir.Enabled = False
frmMotExtorno.Visible = True
cmbMotivos.SetFocus
End Sub
Private Sub cmdSalir_Click()
fsCodExtOpc = ""
txtMovimiento = ""
Call LimpiaFlex(FEExtornarARendir)
Unload Me
End Sub
'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub

Private Sub Form_Load()
    If optBusqueda.item(0).value = True Then
        txtMovimiento.MaxLength = 9
    Else
        txtMovimiento.MaxLength = 4
    End If
    Call CargaControles 'CTI3
End Sub

Private Sub Form_Unload(Cancel As Integer)
fsCodExtOpc = ""
txtMovimiento = ""
Call LimpiaFlex(FEExtornarARendir)
End Sub

Private Sub optBusqueda_Click(index As Integer)
txtMovimiento = ""
Call LimpiaFlex(FEExtornarARendir)
If optBusqueda.item(0).value = True Then
    txtMovimiento.MaxLength = 9
Else
    txtMovimiento.MaxLength = 4
End If
End Sub

Private Sub txtDetExtorno_Change()

End Sub

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdExtonar.SetFocus
    End If
End Sub

Private Sub txtMovimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub
