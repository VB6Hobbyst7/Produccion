VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmCredFormEvalFormato6_EstFinan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Estados Financieros"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalFormato6_EstFinan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   16815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAudit 
      Caption         =   "Auditado"
      Height          =   195
      Left            =   3840
      TabIndex        =   6
      Top             =   80
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8880
      TabIndex        =   5
      Top             =   8580
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   14208
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Estados Financieros"
      TabPicture(0)   =   "frmCredFormEvalFormato6_EstFinan.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fePasivos"
      Tab(0).Control(1)=   "feActivos"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Estado de Ganancias y Pérdidas"
      TabPicture(1)   =   "frmCredFormEvalFormato6_EstFinan.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
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
         Height          =   7575
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   7815
         Begin SICMACT.FlexEdit feEstaGananPerd 
            Height          =   7095
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   12515
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto-Monto-nConsCod-nConsValor"
            EncabezadosAnchos=   "300-4500-2400-0-0"
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
            ColumnasAEditar =   "X-X-2-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-L-C"
            FormatosEdit    =   "0-0-2-0-0"
            CantEntero      =   12
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
      Begin SICMACT.FlexEdit fePasivos 
         Height          =   7575
         Left            =   -66585
         TabIndex        =   9
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   13361
         Cols0           =   7
         ScrollBars      =   0
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-PASIVO-P.P.-P.E.-Total-nConsCod-nConsValor"
         EncabezadosAnchos=   "300-4000-1300-1300-1300-0-0"
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
         ColumnasAEditar =   "X-X-X-X-4-5-X"
         ListaControles  =   "0-0-0-0-1-1-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-C-C"
         FormatosEdit    =   "0-0-2-0-0-0-0"
         CantEntero      =   12
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit feActivos 
         Height          =   7575
         Left            =   -74880
         TabIndex        =   10
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   13361
         Cols0           =   7
         ScrollBars      =   0
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-ACTIVO-P.P.-P.E.-Total-nConsCod-nConsValor"
         EncabezadosAnchos=   "300-4000-1300-1300-1300-0-0"
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
         ColumnasAEditar =   "X-X-X-X-4-5-X"
         ListaControles  =   "0-0-0-0-1-1-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-C-C"
         FormatosEdit    =   "0-0-2-0-0-0-0"
         CantEntero      =   12
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7680
      TabIndex        =   1
      Top             =   8580
      Width           =   1170
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6480
      TabIndex        =   0
      Top             =   8580
      Width           =   1170
   End
   Begin MSMask.MaskEdBox mskFecRegistro 
      Height          =   300
      Left            =   2280
      TabIndex        =   7
      Top             =   80
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha del Estado EE.FF Al:"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   80
      Width           =   2175
   End
End
Attribute VB_Name = "frmCredFormEvalFormato6_EstFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalFormato6_EstFinan
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 2
'** Referencia  : ERS004-2016
'** Creación    : LUCV, 20160525 09:00:00 AM
'**********************************************************************************************
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Dim fnTipoCliente As Integer
Dim sCtaCod As String
Dim gsOpeCod As String
Dim fnTipoRegMant As Integer
Dim fnTipoPermiso As Integer
Dim fbPermiteGrabar As Boolean
Dim fbBloqueaTodo As Boolean
Dim lnIndMaximaCapPago As Double
Dim lnIndCuotaUNM As Double
Dim lnIndCuotaExcFam As Double
Dim lnCondLocal As Integer
Dim rsCredEval As ADODB.Recordset
Dim rsInd As ADODB.Recordset
Dim fnTotalRef As Currency
'lucv
Dim MatCueva As Variant
Dim lnNumForm As Integer

Dim rsFeGastoNeg As ADODB.Recordset
Dim rsFeDatGastoFam As ADODB.Recordset
Dim rsFeDatOtrosIng As ADODB.Recordset
Dim rsFeDatRef As ADODB.Recordset
Dim rsFeDatActivos As ADODB.Recordset
Dim rsFeDatPasivos As ADODB.Recordset
Dim rsFeDatPasivosNo As ADODB.Recordset
Dim rsFeDatPDT As ADODB.Recordset
Dim rsFeDatPDTDet As ADODB.Recordset
Dim rsFeDatPatrimonio As ADODB.Recordset
Dim rsFeDatPasPat As ADODB.Recordset
Dim rsFeDatRatios As ADODB.Recordset
Dim rsFeDatIngNeg As ADODB.Recordset
Dim rsFeDatActivosForm6 As ADODB.Recordset
Dim rsFeDatPasivosForm6 As ADODB.Recordset
Dim rsFeDatEstadoGanPerdForm6 As ADODB.Recordset

Dim rsFeDatActivosForm6Plantilla As ADODB.Recordset
Dim rsFeDatPasivosForm6Plantilla As ADODB.Recordset
Dim rsFeDatEstadoGanPerdForm6Plantilla As ADODB.Recordset

Dim cuotaifi As Integer ' lucv

Dim rsDatGastoNeg As ADODB.Recordset
Dim rsDatGastoFam As ADODB.Recordset
Dim rsDatOtrosIng As ADODB.Recordset
Dim rsDatRef As ADODB.Recordset

Dim fsCliente As String
Dim fsGiroNego As String 'lucv
Dim fsAnioExp As Integer 'lucv
Dim fsMesExp As Integer 'lucv
Dim fsUserAnalista  As String ' lucv

Dim fnMontoDeudaSbs As Currency 'lucv

Dim rsDatActivos As ADODB.Recordset
Dim rsDatPasivos As ADODB.Recordset
Dim rsDatPasivosNo As ADODB.Recordset
Dim rsDatPDT As ADODB.Recordset
Dim rsDatPDTDet As ADODB.Recordset
Dim rsDatPatrimonio As ADODB.Recordset
Dim rsDatPasPat As ADODB.Recordset

Dim rsDatEstadoGP As ADODB.Recordset
Dim rsDatRatios As ADODB.Recordset
Dim rsDatIngNeg As ADODB.Recordset
Dim nTasaIngNeg As Double
Dim nTasaGastoNeg As Double
Dim nTasaGastoFam As Double
Dim nTasaOtrosIng As Double

Dim fnPasivoPE As Double
Dim fnPasivoPP As Double
Dim fnPasivoTOTAL As Double

Dim fnActivoPE As Double
Dim fnActivoPP As Double
Dim fnActivoTOTAL As Double

Dim cSPrd As String, cPrd As String
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval 'lucv
Dim objPista As COMManejador.Pista
Dim nFormato, nPersoneria As Integer
Dim fnMontoIni As Double
Dim lnMin As Double, lnMax As Double
Dim lnMinDol As Double, lnMaxDol As Double
Dim nTC As Double
Dim i, j As Integer
Dim sMes1 As String, sMes2 As String, sMes3 As String
Dim nMes1 As Integer, nMes2 As Integer, nMes3 As Integer
Dim nAnio1 As Integer, nAnio2 As Integer, nAnio3 As Integer
Dim nMontoPDT, nMontoAct, nMontoPas, nMontoPasN As Double

Dim oFrm6 As frmCredFormEvalDetalleFormato6

'LUCV20170915 *****-> Comentó y agregó, según ERS051-2017
'Dim lvPrincipalActivos() As tForEvalResumenEstFinFormato6 ' matriz para activos
'Dim lvPrincipalPasivos() As tForEvalResumenEstFinFormato6 'matriz para pasivos
'Dim lvPrincipalEstGanPer() As tForEvalResumenEstFinFormato6 'matriz para estado de ganancias y perdidas
Dim lvPrincipalActivos() As tFormEvalPrincipalEstFinFormato6    'Matriz Principal-> Activos
Dim lvPrincipalPasivos() As tFormEvalPrincipalEstFinFormato6    'Matriz Principal-> Pasivos
Dim lvPrincipalEstGanPer() As tFormEvalPrincipalEstFinFormato6  'Matriz Principal-> Ganancias y pérdidas
'<***** Fin LUCV20170915
Dim lcCodifi As String
Dim lcDescripcionDet As String

Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)
    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function


'_____________________________________________________________________________________________________________
'******************************************LUCV20160525: EVENTOS Varios*********************************************
'*************************************************************************************************************
'***** LUCV20160525
'Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    txtGiroNeg2.SetFocus
'    End If
'End Sub

''***** LUCV20160525
'Private Sub feGastosNegocio2_Click()
'    cuotaifi = 9
'        If feGastosNegocio2.Col = 2 Then
'            If CInt(feGastosNegocio2.TextMatrix(feGastosNegocio2.row, 0)) = cuotaifi Then
'                feGastosNegocio2.ListaControles = "0-0-1-0"
'            Else
'                feGastosNegocio2.ListaControles = "0-0-0-0"
'            End If
'        End If
'End Sub
'
''***** LUCV20160525
'Private Sub feGastosNegocio2_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
'    psCodigo = feGastosNegocio2.TextMatrix(feGastosNegocio2.row, 2)
'    psDescripcion = feGastosNegocio2.TextMatrix(feGastosNegocio2.row, 1)
'        frmCredFormEvalIfis.Inicio (CLng(feGastosNegocio2.TextMatrix(feGastosNegocio2.row, 2))), fnTotalRef, MatCueva
'        psCodigo = Format(fnTotalRef, "#,##0.00")
'End Sub
'
''***** LUCV20160525
'Private Sub feGastosNegocio2_EnterCell()
'cuotaifi = 9
'If feGastosNegocio2.Col = 2 Then
'    If CInt(feGastosNegocio2.TextMatrix(feGastosNegocio2.row, 0)) = cuotaifi Then
'    feGastosNegocio2.ListaControles = "0-0-1-0"
'    Else
'    feGastosNegocio2.ListaControles = "0-0-0-0"
'    End If
'End If
'End Sub

'***** LUCV20160525
'Private Sub OptCondLocal2_Click(index As Integer)
'    Select Case index
'    Case 1, 2, 3
'        Me.txtCondLocalOtros2.Visible = False
'        Me.txtCondLocalOtros2.Text = ""
'    Case 4
'        Me.txtCondLocalOtros2.Visible = True
'        Me.txtCondLocalOtros2.Text = ""
'    End Select
'    lnCondLocal = index
'End Sub

'***** LUCV20160528 - FeReferidos2
'Private Sub cmdQuitar2_Click()
'    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        feReferidos2.EliminaFila (feReferidos2.row)
'        'txtTotalIfis.Text = Format(SumarCampo(feReferidos2, 2), "#,##0.00")
'    End If
'End Sub

'***** LUCV20160528
'Private Sub cmdAgregar2_Click()
'    If feReferidos2.Rows - 1 < 25 Then
'        feReferidos2.lbEditarFlex = True
'        feReferidos2.AdicionaFila
'        feReferidos2.SetFocus
'        SendKeys "{Enter}"
'    Else
'    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
'    End If
'
'
'    If ValidaDatosReferencia = False Then
'   ' -'staqb.Tab = 1
'    'validaDatos = False
'      '  Exit Function
'    End If
'End Sub

'***** LUCV20160528
'Private Sub feReferidos2_OnCellChange(pnRow As Long, pnCol As Long)
'    If pnCol = 1 Then
'        feReferidos2.TextMatrix(pnRow, pnCol) = UCase(feReferidos2.TextMatrix(pnRow, pnCol))
'    End If
'End Sub

'________________________________________________________________________________________________________________________
'*************************************************LUCV20160525: METODOS Varios **************************************************
'***** LUCV20160525
Public Sub Inicio(ByVal psCtaCod As String, ByVal pnNumForm As Integer)
Call CargaControlesInicio
    
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredito As ADODB.Recordset
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
'    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    
    Me.mskFecRegistro.Enabled = True
    
 '   nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)
    sCtaCod = psCtaCod
    'fnTipoRegMant = psTipoRegMant
    lnNumForm = pnNumForm
    'ActXCodCta.NroCuenta = sCtaCod
    
    ReDim lvDetalleEstFin(0) 'matriz para activos
    ReDim lvDetalleEstFinPasivos(0) 'matriz para pasivos
    
    Me.SSTab1.Tab = 0
    
    Me.Show 1
End Sub

'***** LUCV20160529 / feReferidos2
Public Function ValidaDatosReferencia() As Boolean
    Dim i As Integer, j As Integer
    ValidaDatosReferencia = False
        
'    If feReferidos2.Rows - 1 < 2 Then
'        MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
'        cmdAgregar2.SetFocus
'        ValidaDatosReferencia = False
'        Exit Function
'    End If
'
'    For i = 1 To feReferidos2.Rows - 1  'Verfica Tipo de Valores del DNI
'        If Trim(feReferidos2.TextMatrix(i, 1)) <> "" Then
'            For j = 1 To Len(Trim(feReferidos2.TextMatrix(i, 2)))
'                If (Mid(feReferidos2.TextMatrix(i, 2), j, 1) < "0" Or Mid(feReferidos2.TextMatrix(i, 2), j, 1) > "9") Then
'                   MsgBox "Uno de los Digitos del primer DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
'                   feReferidos2.SetFocus
'                   ValidaDatosReferencia = False
'                   Exit Function
'                End If
'            Next j
'        End If
'    Next i
'
'    For i = 1 To feReferidos2.Rows - 1  'Verfica Longitud del DNI
'        If Trim(feReferidos2.TextMatrix(i, 1)) <> "" Then
'            If Len(Trim(feReferidos2.TextMatrix(i, 2))) <> gnNroDigitosDNI Then
'                MsgBox "Primer DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
'                feReferidos2.SetFocus
'                ValidaDatosReferencia = False
'                Exit Function
'            End If
'        End If
'    Next i
'
'    For i = 1 To feReferidos2.Rows - 1  'Verfica Tipo de Valores del Telefono
'        If Trim(feReferidos2.TextMatrix(i, 1)) <> "" Then
'            For j = 1 To Len(Trim(feReferidos2.TextMatrix(i, 3)))
'                If (Mid(feReferidos2.TextMatrix(i, 3), j, 1) < "0" Or Mid(feReferidos2.TextMatrix(i, 3), j, 1) > "9") Then
'                   MsgBox "Uno de los Digitos del teléfono de la fila " & i & " no es un Numero", vbInformation, "Aviso"
'                   feReferidos2.SetFocus
'                   ValidaDatosReferencia = False
'                   Exit Function
'                End If
'            Next j
'        End If
'    Next i
'
'    For i = 1 To feReferidos2.Rows - 1 'Verfica Tipo de Valores del DNI 2
'        If Trim(feReferidos2.TextMatrix(i, 1)) <> "" Then
'            For j = 1 To Len(Trim(feReferidos2.TextMatrix(i, 5)))
'                If (Mid(feReferidos2.TextMatrix(i, 5), j, 1) < "0" Or Mid(feReferidos2.TextMatrix(i, 5), j, 1) > "9") Then
'                   MsgBox "Uno de los Digitos del segundo DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
'                   feReferidos2.SetFocus
'                   ValidaDatosReferencia = False
'                   Exit Function
'                End If
'            Next j
'        End If
'    Next i
'
'    For i = 1 To feReferidos2.Rows - 1   'Verfica Longitud del DNI 2
'        If Trim(feReferidos2.TextMatrix(i, 1)) <> "" Then
'            If Len(Trim(feReferidos2.TextMatrix(i, 5))) <> gnNroDigitosDNI Then
'                MsgBox "Segundo DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
'                feReferidos2.SetFocus
'                ValidaDatosReferencia = False
'                Exit Function
'            End If
'        End If
'    Next i
'
'    For i = 1 To feReferidos2.Rows - 1 'Verfica ambos DNI que no sean iguales
'        If Trim(feReferidos2.TextMatrix(i, 1)) <> "" Then
'            If Trim(feReferidos2.TextMatrix(i, 2)) = Trim(feReferidos2.TextMatrix(i, 5)) Then
'                MsgBox "Los DNI de la fila " & feReferidos2.row & " son iguales", vbInformation, "Aviso"
'                feReferidos2.SetFocus
'                ValidaDatosReferencia = False
'                Exit Function
'            End If
'        End If
'    Next i
    ValidaDatosReferencia = True
End Function


Public Function ValidaDatos() As Boolean
'    ValidaDatos = False
'
'If fnTipoPermiso = 3 Then
'            If Not ValidaPorcParticip Then
'                MsgBox "El % de Participacion no debe sumar mas del 100%", vbInformation, "Aviso"
'                validaDatos = False
'                Exit Function
'            End If
'            If Round(ccur(lblMontoMax.Caption), 2) < Round(ccur(txtCalcMonto.Text), 2) Then
'                MsgBox "El Monto Máximo del Credito es menor al ingresado en el calculo", vbInformation, "Aviso"
'                txtCalcMonto.SetFocus
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'            If Round(ccur(lblCuotaEstima.Caption), 2) > Round(ccur(txtCuotaPagar.Text), 2) Then
'                MsgBox "La Couta Estimada a Pagar es mayor a la Probable Cuota por Pagar", vbInformation, "Aviso"
'                txtCuotaPagar.SetFocus
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'            If Trim(txtGiroNeg.Text) = "" Then
'                MsgBox "Falta ingresar el Giro del Negocio", vbInformation, "Aviso"
'                txtGiroNeg.SetFocus
'                validaDatos = False
'                Exit Function
'            End If
'            If Trim(txtGiroNeg.Text) = "" Then
'                MsgBox "Falta ingresar el Giro del Negocio", vbInformation, "Aviso"
'                txtGiroNeg.SetFocus
'                validaDatos = False
'                Exit Function
'            End If
'            If OptCondLocal(1).value = 0 And OptCondLocal(2).value = 0 And OptCondLocal(3).value = 0 And OptCondLocal(4).value = 0 Then
'                MsgBox "Falta elegir la Condicion del local", vbInformation, "Aviso"
'                OptCondLocal(1).SetFocus
'                validaDatos = False
'                Exit Function
'            End If
'            If OptCondLocal(4).value = 1 Then
'                If Trim(txtCondLocalOtros.Text) = "" Then
'                    MsgBox "Falta detallar la Condicion del local", vbInformation, "Aviso"
'                    txtCondLocalOtros.SetFocus
'                    validaDatos = False
'                    Exit Function
'                End If
'            End If
'            If txtCuotaPagar.value = 0 Then
'                MsgBox "Falta ingresar la Probable cuota a pagar", vbInformation, "Aviso"
'                txtCuotaPagar.SetFocus
'                validaDatos = False
'                Exit Function
'            End If
'            If spnCuotas.valor = 0 Then
'                MsgBox "Falta ingresar el nro de cuotas", vbInformation, "Aviso"
'                spnCuotas.SetFocus
'                validaDatos = False
'                Exit Function
'            End If
'            If txtUltEndeuda.value <> 0 Then
'                If Trim(txtFecUltEndeuda.Text) = "__/__/____" Then
'                    MsgBox "Falta ingresar la fecha del ultimo endeudamiento", vbInformation, "Aviso"
'                    txtFecUltEndeuda.SetFocus
'                    validaDatos = False
'                    Exit Function
'                End If
'            End If
'            If cboMontoSol.ListIndex = -1 Then
'                MsgBox "Falta seleccionar la moneda", vbInformation, "Aviso"
'                cboMontoSol.SetFocus
'                validaDatos = False
'                Exit Function
'            End If
'            If txtMontoSol.value = 0 Then
'                MsgBox "Falta ingresar el monto solicitado", vbInformation, "Aviso"
'                txtMontoSol.SetFocus
'                validaDatos = False
'                Exit Function
'            End If
'    If lblVentaProm.Caption = 0 Then
'        MsgBox "Falta ingresar la Venta Promedio en mes", vbInformation, "Aviso"
'        txtVentaProm.SetFocus
'        SSTab2.Tab = 0
'        ValidaDatos = False
'        Exit Function
'    End If
'    If txtCostoVenta.value = 0 Then
'        MsgBox "Falta ingresar el costo de Venta", vbInformation, "Aviso"
'        txtCostoVenta.SetFocus
'        SSTab2.Tab = 0
'        ValidaDatos = False
'        Exit Function
'    End If
'            If Trim(lblUtilNeta.Caption) = "" Then
'                MsgBox "Faltan datos para el calculo de la Utilidad Neta", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'            If Trim(lblExcedenteFam.Caption) = "" Then
'                MsgBox "Faltan datos para el calculo del Excedente Familiar", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'            If Trim(lblMontoMax.Caption) = "" Then
'                MsgBox "Faltan datos para el calculo del Monto maximo del credito", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'            If Trim(lblCuotaEstima.Caption) = "" Then
'                MsgBox "Faltan datos para el calculo de la cuota estimada", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'            If Trim(lblCuotaUNM.Caption) = "" Then
'                MsgBox "Faltan datos para el calculo de la Cuota / Utilidad Neta", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'            If Trim(lblCuotaExcedeFam.Caption) = "" Then
'                MsgBox "Faltan datos para el calculo de la Cuota / Excedente Familiar", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'            If Trim(txtComent.Text) = "" Then
'                MsgBox "Faltan ingresar el comentario", vbInformation, "Aviso"
'                txtComent.SetFocus
'                SSTab2.Tab = 1
'                validaDatos = False
'                Exit Function
'            End If
'
'            Dim i As Integer
'
'            For i = 1 To fgVentasSem.Rows - 1
'                If fgVentasSem.TextMatrix(i, 0) <> "" Then
'                    If Trim(fgVentasSem.TextMatrix(i, 1)) = "" Or Trim(fgVentasSem.TextMatrix(i, 2)) = "" _
'                        Or Trim(fgVentasSem.TextMatrix(i, 3)) = "" Or Trim(fgVentasSem.TextMatrix(i, 4)) = "" _
'                        Or Trim(fgVentasSem.TextMatrix(i, 5)) = "" Or Trim(fgVentasSem.TextMatrix(i, 6)) = "" _
'                        Or Trim(fgVentasSem.TextMatrix(i, 7)) = "" Or Trim(fgVentasSem.TextMatrix(i, 8)) = "" _
'                        Or Trim(fgVentasSem.TextMatrix(i, 9)) = "" Then
'                        MsgBox "Faltan datos en la lista de las Ventas de la Semana", vbInformation, "Aviso"
'                        SSTab2.Tab = 0
'                        validaDatos = False
'                        Exit Function
'                    End If
'                End If
'            Next i
'
'            If fgVentasSem.Rows - 1 < 2 Then
'
'            End If
'
'            If ValidaGrillas(fgGastoNeg) = False Then
'                MsgBox "Faltan datos en la lista de Gastos del Negocio", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'
'            If ValidaGrillas(fgGastoFam) = False Then
'                MsgBox "Faltan datos en la lista de Gastos Familiares", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'
'            If ValidaGrillas(fgOtrosIng) = False Then
'                MsgBox "Faltan datos en la lista de Otros Ingresos", vbInformation, "Aviso"
'                SSTab2.Tab = 0
'                validaDatos = False
'                Exit Function
'            End If
'
'            For i = 1 To fgRef.Rows - 1
'                If fgRef.TextMatrix(i, 0) <> "" Then
'                    If Trim(fgRef.TextMatrix(i, 1)) = "" Or Trim(fgRef.TextMatrix(i, 2)) = "" Or Trim(fgRef.TextMatrix(i, 3)) = "" Or Trim(fgRef.TextMatrix(i, 4)) = "" Or Trim(fgRef.TextMatrix(i, 5)) = "" Then
'                        MsgBox "Faltan datos en la lista de Referencias", vbInformation, "Aviso"
'                        SSTab2.Tab = 1
'                        validaDatos = False
'                        Exit Function
'                    End If
'                End If
'            Next i
            
'            If ValidaDatosReferencia = False Then
'                'SSTab2.Tab = 2
'                ValidaDatos = False
'                Exit Function
'            End If
'
'    ElseIf fnTipoPermiso = 2 Then
'        If Trim(txtVerif.Text) = "" Then
'            MsgBox "Favor de ingresar la Validación respectiva", vbInformation, "Aviso"
'            txtVerif.SetFocus
'            'SSTab2.Tab = 1
'            ValidaDatos = False
'            Exit Function
'        End If
'    End If
'    ValidaDatos = True
End Function



Private Function CargaControles(ByVal TipoPermiso As Integer, ByVal pPermiteGrabar As Boolean, Optional ByVal pBloqueaTodo As Boolean = False) As Boolean
    If TipoPermiso = 1 Then
        Call HabilitaControles(False, False, False)
        CargaControles = True
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False, True, pPermiteGrabar)
        CargaControles = True
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True, False, True)
        CargaControles = True
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        CargaControles = False
    End If
    If pBloqueaTodo Then
        Call HabilitaControles(False, False, False)
        CargaControles = True
    End If
End Function


Private Function HabilitaControles(ByVal pbHabilitaA As Boolean, ByVal pbHabilitaB As Boolean, ByVal pbHabilitaGuardar As Boolean)

'       txtFechaEvaluacion2.Enabled = True
'    txtGiroNeg.Enabled = pbHabilitaA
'    spnExpEmpAnio.Enabled = pbHabilitaA
'    spnExpEmpMes.Enabled = pbHabilitaA
'    spnTiempoLocalAnio.Enabled = pbHabilitaA
'    spnTiempoLocalMes.Enabled = pbHabilitaA
'    OptCondLocal(1).Enabled = pbHabilitaA
'    OptCondLocal(2).Enabled = pbHabilitaA
'    OptCondLocal(3).Enabled = pbHabilitaA
'    OptCondLocal(4).Enabled = pbHabilitaA
'    txtCondLocalOtros.Enabled = pbHabilitaA
'    txtCuotaPagar.Enabled = pbHabilitaA
'    spnCuotas.Enabled = pbHabilitaA
'    txtUltEndeuda.Enabled = pbHabilitaA
'    txtFecUltEndeuda.Enabled = pbHabilitaA
'    'cboMontoSol.Enabled = pbHabilitaA
'    txtMontoSol.Enabled = pbHabilitaA
'    fgVentasSem.Enabled = pbHabilitaA
'    cmdAgregarVentSem.Enabled = pbHabilitaA
'    cmdQuitarVentSem.Enabled = pbHabilitaA
'    fgGastoNeg.Enabled = pbHabilitaA
'    cmdAgregarGastoNeg.Enabled = pbHabilitaA
'    cmdQuitarGastoNeg.Enabled = pbHabilitaA
'    fgGastoFam.Enabled = pbHabilitaA
'    cmdAgregarGastoFam.Enabled = pbHabilitaA
'    cmdQuitarGastoFam.Enabled = pbHabilitaA
'    fgOtrosIng.Enabled = pbHabilitaA
'    cmdAgregarOtrosIng.Enabled = pbHabilitaA
'    cmdQuitarOtrosIng.Enabled = pbHabilitaA
'    txtCalcMonto.Enabled = pbHabilitaA
'    txtCalcTEM.Enabled = pbHabilitaA
'    spnCalcCuotas.Enabled = pbHabilitaA
'    cmdCalcular.Enabled = pbHabilitaA
'    txtComent.Enabled = pbHabilitaA
'    fgRef.Enabled = pbHabilitaA
'    cmdAgregarRef.Enabled = pbHabilitaA
'    cmdQuitarRef.Enabled = pbHabilitaA
'
'    txtVerif.Enabled = pbHabilitaB
'    cmdGrabar.Enabled = pbHabilitaGuardar
'
'    If Mid(sCtaCod, 9, 1) = "2" Then
'        Me.txtMontoSol.BackColor = RGB(200, 255, 200)
'        Me.txtCuotaPagar.BackColor = RGB(200, 255, 200)
'
'        txtCalcMonto.BackColor = RGB(200, 255, 200)
'        lblMontoMax.BackColor = RGB(200, 255, 200)
'        lblCuotaEstima.BackColor = RGB(200, 255, 200)
'        lblCuotaUNM.BackColor = RGB(200, 255, 200)
'        lblCuotaExcedeFam.BackColor = RGB(200, 255, 200)
'    Set DCredito = Nothing
'    Else
'        Me.txtMontoSol.BackColor = &HFFFFFF
'        Me.txtCuotaPagar.BackColor = &HFFFFFF
'        txtCalcMonto.BackColor = &HFFFFFF
'
'        lblMontoMax.BackColor = &HFFFFFF
'        lblCuotaEstima.BackColor = &HFFFFFF
'        lblCuotaUNM.BackColor = &HFFFFFF
'        lblCuotaExcedeFam.BackColor = &HFFFFFF
'    End If
End Function

Private Function Registro()
    gsOpeCod = gCredRegistrarEvaluacionCred
    'txtNombreCliente2.Text = fsCliente
    'txtNombreCliente2.Enabled = False
    'txtGiroNeg2.Text = fsGiroNego
    'txtGiroNeg2.Enabled = False
End Function

Private Function Mantenimiento()
'    Dim lnFila As Integer
'    If fnTipoPermiso = 3 Then
'        gsOpeCod = gCredMantenimientoEvaluacionCred
'    Else
'        'gsOpeCod = gCredVerificacionEvaluacionCred
'    End If
'
'    txtGiroNeg.Text = rsCredEval!cGiroNeg
'    spnExpEmpAnio.valor = rsCredEval!nExpEmpAnio
'    spnExpEmpMes.valor = rsCredEval!nExpEmpMes
'    spnTiempoLocalAnio.valor = rsCredEval!nTiempoLocalAnio
'    spnTiempoLocalMes.valor = rsCredEval!nTiempoLocalMes
'    OptCondLocal(rsCredEval!nCondLocal).value = 1
'    txtCondLocalOtros.Text = rsCredEval!cCondLocalOtros
'    txtCuotaPagar.Text = Format(rsCredEval!cCuotaPagar, "#,##0.00")
'    spnCuotas.valor = rsCredEval!nCuotas
'    txtUltEndeuda.Text = Format(rsCredEval!cUltEndeuda, "#,##0.00")
'    If rsCredEval!cUltEndeuda = 0 Then
'        txtFechaEvaluacion2.Enabled = False
'    Else
'        If fnTipoPermiso = 3 Then
'            txtFechaEvaluacion2.Enabled = True
'        End If
'    End If
'    txtFecUltEndeuda.Text = Format(IIf(rsCredEval!cFecUltEndeuda = "01/01/1900", "__/__/____", rsCredEval!cFecUltEndeuda), "dd/mm/yyyy")
'    cboMontoSol.ListIndex = IndiceListaCombo(cboMontoSol, rsCredEval!nmoneda)
'    txtMontoSol.Text = Format(rsCredEval!nMontoSol, "#,##0.00")
'    lblVentaProm.Caption = Format(rsCredEval!nVentaProm, "#,##0.00")
'    lblCostoVenta.Caption = Format(rsCredEval!nCostoVenta, "#,##0.00")
'    'Call FormatearGrillas(fgGastoNeg)
'    Call LimpiaFlex(fgGastoNeg)
'    Do While Not rsDatGastoNeg.EOF
'        fgGastoNeg.AdicionaFila
'        lnFila = fgGastoNeg.row
'        fgGastoNeg.TextMatrix(lnFila, 1) = rsDatGastoNeg!cConcepto
'        fgGastoNeg.TextMatrix(lnFila, 2) = Format(rsDatGastoNeg!nmonto, "#,##0.00")
'        lblTotalGastoNeg.Caption = Format(ccur(IIf(lblTotalGastoNeg.Caption = "", 0, lblTotalGastoNeg.Caption)) + rsDatGastoNeg!nmonto, "#,##0.00")
'        rsDatGastoNeg.MoveNext
'    Loop
'    rsDatGastoNeg.Close
'    Set rsDatGastoNeg = Nothing
'    'Call FormatearGrillas(fgGastoFam)
'    Call LimpiaFlex(fgGastoFam)
'    Do While Not rsDatGastoFam.EOF
'        fgGastoFam.AdicionaFila
'        lnFila = fgGastoFam.row
'        fgGastoFam.TextMatrix(lnFila, 1) = rsDatGastoFam!cConcepto
'        fgGastoFam.TextMatrix(lnFila, 2) = Format(rsDatGastoFam!nmonto, "#,##0.00")
'        lblTotalGastoFam.Caption = Format(ccur(IIf(lblTotalGastoFam.Caption = "", 0, lblTotalGastoFam.Caption)) + rsDatGastoFam!nmonto, "#,##0.00")
'        rsDatGastoFam.MoveNext
'    Loop
'    rsDatGastoFam.Close
'    Set rsDatGastoFam = Nothing
'    'Call FormatearGrillas(fgOtrosIng)
'    Call LimpiaFlex(fgOtrosIng)
'    Do While Not rsDatOtrosIng.EOF
'        fgOtrosIng.AdicionaFila
'        lnFila = fgOtrosIng.row
'        fgOtrosIng.TextMatrix(lnFila, 1) = rsDatOtrosIng!cConcepto
'        fgOtrosIng.TextMatrix(lnFila, 2) = Format(rsDatOtrosIng!nmonto, "#,##0.00")
'        lblTotalOtrosIng.Caption = Format(ccur(IIf(lblTotalOtrosIng.Caption = "", 0, lblTotalOtrosIng.Caption)) + rsDatOtrosIng!nmonto, "#,##0.00")
'        rsDatOtrosIng.MoveNext
'    Loop
'    rsDatOtrosIng.Close
'    Set rsDatOtrosIng = Nothing
'    Call LimpiaFlex(fgVentasSem)
'    Do While Not rsDatVentaSem.EOF
'        fgVentasSem.AdicionaFila
'        lnFila = fgVentasSem.row
'        fgVentasSem.TextMatrix(lnFila, 1) = rsDatVentaSem!cProducto
'        fgVentasSem.TextMatrix(lnFila, 2) = Format(rsDatVentaSem!nVentaAlta, "#,##0.00")
'        fgVentasSem.TextMatrix(lnFila, 3) = rsDatVentaSem!nDiasAlta
'        fgVentasSem.TextMatrix(lnFila, 4) = Format(rsDatVentaSem!nVentaBaja, "#,##0.00")
'        fgVentasSem.TextMatrix(lnFila, 5) = rsDatVentaSem!nDiasBaja
'        fgVentasSem.TextMatrix(lnFila, 6) = Format(rsDatVentaSem!nTotalMes, "#,##0.00")
'        fgVentasSem.TextMatrix(lnFila, 7) = Format(rsDatVentaSem!nCosto, "#,##0.00")
'        fgVentasSem.TextMatrix(lnFila, 8) = Format(rsDatVentaSem!nParticip, "#,##0.00")
'        fgVentasSem.TextMatrix(lnFila, 9) = Format(rsDatVentaSem!nReal, "#,##0.00")
'        rsDatVentaSem.MoveNext
'    Loop
'    rsDatVentaSem.Close
'    Set rsDatVentaSem = Nothing
'
'    lblUtilNeta.Caption = Format(rsCredEval!nUtilNeta, "#,##0.00")
'    lblExcedenteFam.Caption = Format(rsCredEval!nExcedenteFam, "#,##0.00")
'
'    txtCalcMonto.Text = Format(rsCredEval!nMontoCalc, "#,##0.00")
'    txtCalcTEM.Text = Format(rsCredEval!nTEMCalc, "#,##0.00")
'    spnCalcCuotas.valor = rsCredEval!nCuotasCalc
'
'    lblMontoMax.Caption = Format(rsCredEval!nMontoMax, "#,##0.00")
'    lblCuotaEstima.Caption = Format(rsCredEval!nCuotaEstima, "#,##0.00")
'    lblCuotaUNM.Caption = Format(rsCredEval!nCuotaUNM, "#,##0.00")
'    lblCuotaExcedeFam.Caption = Format(rsCredEval!nCuotaExcedeFam, "#,##0.00")
'
'    txtComent.Text = rsCredEval!cComent
'    'Call FormatearGrillas(fgRef)
'    Call LimpiaFlex(fgRef)
'    Do While Not rsDatRef.EOF
'        fgRef.AdicionaFila
'        lnFila = fgRef.row
'        fgRef.TextMatrix(lnFila, 1) = rsDatRef!cNombre
'        fgRef.TextMatrix(lnFila, 2) = rsDatRef!cDNI
'        fgRef.TextMatrix(lnFila, 3) = rsDatRef!cTelef
'        fgRef.TextMatrix(lnFila, 4) = rsDatRef!cReferido
'        fgRef.TextMatrix(lnFila, 5) = rsDatRef!cDNIRef
'        rsDatRef.MoveNext
'    Loop
'    rsDatRef.Close
'    Set rsDatRef = Nothing
'
'    txtVerif.Text = rsCredEval!cVerif

End Function

Private Sub CargaControlesInicio()
    Call CargarFlexEdit
'   Call CargarConstantes
End Sub

Private Sub CargarFlexEdit()
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim nMonto As Double
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval

    Dim nFila As Integer
    Dim NumRegRS  As Integer
    Dim NumRegRSPasivos As Integer
    Dim NumRegRSEstGanPer As Integer

    CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(6, sCtaCod, , , , , , , , , , , rsFeDatActivosForm6Plantilla, rsFeDatPasivosForm6Plantilla, rsFeDatEstadoGanPerdForm6Plantilla)
    
    '---------------------------- Activos
    feActivos.Clear
    feActivos.FormaCabecera
    feActivos.rows = 2
        Call LimpiaFlex(feActivos)
        nFila = 0
        NumRegRS = rsFeDatActivosForm6Plantilla.RecordCount
        'ReDim lvPrincipalActivos(NumRegRS)
        ReDim lvPrincipalActivos(NumRegRS)
        
    Do While Not rsFeDatActivosForm6Plantilla.EOF
        feActivos.AdicionaFila
        lnFila = feActivos.row
        feActivos.TextMatrix(lnFila, 1) = rsFeDatActivosForm6Plantilla!Concepto
        feActivos.TextMatrix(lnFila, 2) = Format(rsFeDatActivosForm6Plantilla!PP, "#,#0.00")
        feActivos.TextMatrix(lnFila, 3) = Format(rsFeDatActivosForm6Plantilla!PE, "#,#0.00")
        feActivos.TextMatrix(lnFila, 4) = Format(rsFeDatActivosForm6Plantilla!Total, "#,#0.00")
        feActivos.TextMatrix(lnFila, 5) = rsFeDatActivosForm6Plantilla!nConsCod
        feActivos.TextMatrix(lnFila, 6) = rsFeDatActivosForm6Plantilla!nConsValor
                
        '----------------- llena matriz activos
        nFila = nFila + 1
        'nFila = rsFeDatActivosForm6Plantilla!nConsValor
        lvPrincipalActivos(nFila).cConcepto = rsFeDatActivosForm6Plantilla!Concepto
        lvPrincipalActivos(nFila).nImportePP = rsFeDatActivosForm6Plantilla!PP
        lvPrincipalActivos(nFila).nImportePE = rsFeDatActivosForm6Plantilla!PE
        lvPrincipalActivos(nFila).nConsCod = rsFeDatActivosForm6Plantilla!nConsCod
        lvPrincipalActivos(nFila).nConsValor = rsFeDatActivosForm6Plantilla!nConsValor
                
        Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
            Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
                'Me.feActivos.CellBackColor() = &HC0FFFF 'color amarillo claro
                Me.feActivos.BackColorRow &HC0FFFF, True 'color amarillo claro
                'Me.feActivos.CellBackColor() = QBColor(6) 'color amarillo claro
        End Select
        
        Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
            Case 1, 10, 17
                'Me.feActivos.CellForeColor() = QBColor(1) 'color azul
                'Me.feActivos.CellBackColor() = QBColor(8) 'gris
                 Me.feActivos.BackColorRow QBColor(8), True 'color amarillo claro
        End Select
        
        rsFeDatActivosForm6Plantilla.MoveNext
    Loop
    rsFeDatActivosForm6Plantilla.Close
    Set rsFeDatActivosForm6Plantilla = Nothing


    '----------------- Pasivos
    fePasivos.FormaCabecera
    fePasivos.rows = 2
        Call LimpiaFlex(fePasivos)
        
        nFila = 0
        NumRegRSPasivos = rsFeDatPasivosForm6Plantilla.RecordCount
        ReDim lvPrincipalPasivos(NumRegRSPasivos)
        
    Do While Not rsFeDatPasivosForm6Plantilla.EOF
        fePasivos.AdicionaFila
        lnFila = fePasivos.row
        fePasivos.TextMatrix(lnFila, 1) = rsFeDatPasivosForm6Plantilla!Concepto
        fePasivos.TextMatrix(lnFila, 2) = Format(rsFeDatPasivosForm6Plantilla!PP, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 3) = Format(rsFeDatPasivosForm6Plantilla!PE, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 4) = Format(rsFeDatPasivosForm6Plantilla!Total, "#,#0.00")
        fePasivos.TextMatrix(lnFila, 5) = rsFeDatPasivosForm6Plantilla!nConsCod
        fePasivos.TextMatrix(lnFila, 6) = rsFeDatPasivosForm6Plantilla!nConsValor

        '-----------------llena matriz pasivos
        nFila = nFila + 1
        lvPrincipalPasivos(nFila).cConcepto = rsFeDatPasivosForm6Plantilla!Concepto
        lvPrincipalPasivos(nFila).nImportePP = rsFeDatPasivosForm6Plantilla!PP
        lvPrincipalPasivos(nFila).nImportePE = rsFeDatPasivosForm6Plantilla!PE
        lvPrincipalPasivos(nFila).nConsCod = rsFeDatPasivosForm6Plantilla!nConsCod
        lvPrincipalPasivos(nFila).nConsValor = rsFeDatPasivosForm6Plantilla!nConsValor

        '-----------------pinta items que se ingresaran detalles en pasivos
        Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
            Case 2, 5, 6, 7, 9, 11
                'Me.fePasivos.CellBackColor() = &HC0FFFF
                Me.fePasivos.BackColorRow &HC0FFFF, True 'color amarillo claro
        End Select

        Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
            Case 1, 10, 16, 23, 24, 25
                'Me.fePasivos.CellBackColor() = QBColor(8) 'color gris
                 Me.fePasivos.BackColorRow QBColor(8), True 'color gris
        End Select

        rsFeDatPasivosForm6Plantilla.MoveNext
    Loop
    rsFeDatPasivosForm6Plantilla.Close
    Set rsFeDatPasivosForm6Plantilla = Nothing

    '----------------- ESTADO DE GANANCIA Y PERDIDAS
    feEstaGananPerd.FormaCabecera
    feEstaGananPerd.rows = 2
        Call LimpiaFlex(feEstaGananPerd)
        
        nFila = 0
        NumRegRSEstGanPer = rsFeDatEstadoGanPerdForm6Plantilla.RecordCount
        ReDim lvPrincipalEstGanPer(NumRegRSEstGanPer)

    Do While Not rsFeDatEstadoGanPerdForm6Plantilla.EOF
        feEstaGananPerd.AdicionaFila
        lnFila = feEstaGananPerd.row
        feEstaGananPerd.TextMatrix(lnFila, 1) = rsFeDatEstadoGanPerdForm6Plantilla!Concepto
        feEstaGananPerd.TextMatrix(lnFila, 2) = Format(rsFeDatEstadoGanPerdForm6Plantilla!nMonto, "#,#0.00")
        feEstaGananPerd.TextMatrix(lnFila, 3) = rsFeDatEstadoGanPerdForm6Plantilla!nConsCod
        feEstaGananPerd.TextMatrix(lnFila, 4) = rsFeDatEstadoGanPerdForm6Plantilla!nConsValor

        '-----------------llena matriz estado ganancias y perdidas
        nFila = nFila + 1
        lvPrincipalEstGanPer(nFila).cConcepto = rsFeDatEstadoGanPerdForm6Plantilla!Concepto
        lvPrincipalEstGanPer(nFila).nImportePP = rsFeDatEstadoGanPerdForm6Plantilla!nMonto
        lvPrincipalEstGanPer(nFila).nConsCod = rsFeDatEstadoGanPerdForm6Plantilla!nConsCod
        lvPrincipalEstGanPer(nFila).nConsValor = rsFeDatEstadoGanPerdForm6Plantilla!nConsValor

        '-----------------pinta items que se ingresaran detalles en pasivos
        Select Case CInt(feEstaGananPerd.TextMatrix(Me.feEstaGananPerd.row, 0))
            Case 3, 6, 9, 16, 19
                Me.feEstaGananPerd.CellBackColor() = QBColor(8)
        End Select
        
        rsFeDatEstadoGanPerdForm6Plantilla.MoveNext
    Loop
    rsFeDatEstadoGanPerdForm6Plantilla.Close
    Set rsFeDatEstadoGanPerdForm6Plantilla = Nothing

End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdLimpiar_Click()
    Call CargarFlexEdit
End Sub

Private Sub cmdRegistrar_Click()
    Dim oNCred As COMDCredito.DCOMFormatosEval
    Dim i, j As Integer
    Dim nId As String
    Dim rsBuscaFe As ADODB.Recordset
    Dim nSuma As Double
    Dim cCaptionOriginal As String
    Dim cPunto As String
    Set oNCred = New COMDCredito.DCOMFormatosEval
              
    cCaptionOriginal = Trim(Me.Caption)
    cPunto = "."

    ' valida si se ingresó la fecha de eval
    If Not IsDate(mskFecRegistro) Then
        MsgBox "Ingrese una fecha por favor...", vbOKOnly + vbInformation, "Atención"
        Exit Sub
    End If
    
    ' valida que la feha no exista
    'rsBuscaFe = oNCred.ValidaFechaCredFormEval(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"))
    If oNCred.ValidaFechaCredFormEval(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd")).RecordCount > 0 Then
        MsgBox "La fecha que ingresó ya fue registrada, por favor verifique.", vbOKOnly + vbInformation, "Atención"
        Exit Sub
    End If
    
    '-- verifica si activos tiene datos
        nSuma = 0
        If UBound(lvPrincipalActivos) > 0 Then
            For i = 1 To UBound(lvPrincipalActivos)
'                Me.Caption = Me.Caption + cPunto
                If CCur(Me.feActivos.TextMatrix(i, 4)) > 0 Then
                    nSuma = nSuma + CCur(Me.feActivos.TextMatrix(i, 4))
                    'Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), ccur(Me.feActivos.TextMatrix(i, 2)), ccur(Me.feActivos.TextMatrix(i, 3)), ccur(Me.feActivos.TextMatrix(i, 4)))
                End If
            Next i
        End If
        If nSuma = 0 Then
            MsgBox "Ingrese datos en los Activos...", vbOKOnly + vbInformation, "Atención"
            Exit Sub
        End If
        
        If CCur(Me.feActivos.TextMatrix(1, 4)) = 0 Then
            MsgBox "El monto de Activo Corriente no debe ser cero.", vbOKOnly, "Atención"
            Exit Sub
        End If
        If CCur(Me.feActivos.TextMatrix(17, 4)) = 0 Then
            MsgBox "El monto de Total Activo no debe ser cero.", vbOKOnly, "Atención"
            Exit Sub
        End If
        If CCur(Me.feActivos.TextMatrix(17, 4)) <> CCur(Me.fePasivos.TextMatrix(25, 4)) Then
            MsgBox "El Activo y Pasivo no cuadran, por favor verificar.", vbOKOnly, "Atención"
            Exit Sub
        End If
    
    '-- verifica si pasivos tiene datos
        nSuma = 0
        If UBound(lvPrincipalPasivos) > 0 Then
            For i = 1 To UBound(lvPrincipalPasivos)
 '               Me.Caption = Me.Caption + cPunto
                If CCur(Me.fePasivos.TextMatrix(i, 4)) > 0 Then
                    nSuma = nSuma + CCur(Me.fePasivos.TextMatrix(i, 4))
                    'Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), ccur(Me.feActivos.TextMatrix(i, 2)), ccur(Me.feActivos.TextMatrix(i, 3)), ccur(Me.feActivos.TextMatrix(i, 4)))
                End If
            Next i
        End If
        If nSuma = 0 Then
            MsgBox "Ingrese datos en los Pasivos...", vbOKOnly + vbInformation, "Atención"
            Exit Sub
        End If
        
        If CCur(Me.fePasivos.TextMatrix(1, 4)) = 0 Then
            MsgBox "El monto de Pasivo Corriente no debe ser cero.", vbOKOnly, "Atención"
            Exit Sub
        End If
        If CCur(Me.fePasivos.TextMatrix(23, 4)) = 0 Then
            MsgBox "El monto de Total Pasivo no debe ser cero.", vbOKOnly, "Atención"
            Exit Sub
        End If
    
    '-- verifica si est ganan y perd tiene datos
        nSuma = 0
        If UBound(lvPrincipalEstGanPer) > 0 Then
            For i = 1 To UBound(lvPrincipalEstGanPer)
  '              Me.Caption = Me.Caption + cPunto
                If CCur(Me.feEstaGananPerd.TextMatrix(i, 2)) > 0 Then
                    nSuma = nSuma + CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
                End If
            Next i
        End If
        If nSuma = 0 Then
            MsgBox "Ingrese datos en los Estados de Ganancias y Perdidas...", vbOKOnly + vbInformation, "Atención"
            Exit Sub
        End If
        
        Me.Caption = cCaptionOriginal
    
    If MsgBox("Los Datos ingresados se guardarán, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

        '-- ACTIVOS
        If UBound(lvPrincipalActivos) > 0 Then
            For i = 1 To UBound(lvPrincipalActivos)
                Me.Caption = Me.Caption + cPunto
                If CCur(Me.feActivos.TextMatrix(i, 4)) <> 0 Then
                    'If i = 1 Or i = 10 Or i = 17 Then
                    If i = 17 Then
                        Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CCur(Me.feActivos.TextMatrix(i, 2)), CCur(Me.feActivos.TextMatrix(i, 3)), CCur(Me.feActivos.TextMatrix(i, 4)))
                    Else
                        Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6Det(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CCur(Me.feActivos.TextMatrix(i, 2)), CCur(Me.feActivos.TextMatrix(i, 3)), CCur(Me.feActivos.TextMatrix(i, 4)))
                    End If
                End If
            Next i
        End If

        '-- activos det DETALLE formato6
        If UBound(lvPrincipalActivos) > 0 Then
            For i = 1 To UBound(lvPrincipalActivos)
                Me.Caption = Me.Caption + cPunto
                If lvPrincipalActivos(i).nDetPP > 0 Then
                    For j = 1 To UBound(lvPrincipalActivos(i).vPP)
                        Me.Caption = Me.Caption + cPunto
                        Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
                            sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, _
                            CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), 1, _
                            Format(lvPrincipalActivos(i).vPP(j).dFecha, "yyyyMMdd"), _
                            lvPrincipalActivos(i).vPP(j).CDescripcion, _
                            lvPrincipalActivos(i).vPP(j).nImporte, 0, "")
                    Next j
                End If

                If lvPrincipalActivos(i).nDetPE > 0 Then
                    For j = 1 To UBound(lvPrincipalActivos(i).vPE)
                        Me.Caption = Me.Caption + cPunto
                        Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
                            sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, _
                            CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), 2, _
                            Format(lvPrincipalActivos(i).vPE(j).dFecha, "yyyyMMdd"), _
                            lvPrincipalActivos(i).vPE(j).CDescripcion, 0, _
                            lvPrincipalActivos(i).vPE(j).nImporte, "")
                    Next j
                End If
            Next i
        End If

        '-- PASIVOS

        If UBound(lvPrincipalPasivos) > 0 Then
            For i = 1 To UBound(lvPrincipalPasivos)
                Me.Caption = Me.Caption + cPunto
                If CCur(Me.fePasivos.TextMatrix(i, 4)) <> 0 Then
                    'If i = 1 Or i = 10 Or i = 16 Then
                    'If i = 16 Or i = 17 Or i = 18 Then
                    If i = 23 Or i = 24 Or i = 25 Then
                        Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), _
                            chkAudit.value, CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), CCur(Me.fePasivos.TextMatrix(i, 2)), _
                            CCur(Me.fePasivos.TextMatrix(i, 3)), CCur(Me.fePasivos.TextMatrix(i, 4)))
                    Else
                        Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6Det(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), _
                            chkAudit.value, CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), _
                            CCur(Me.fePasivos.TextMatrix(i, 2)), CCur(Me.fePasivos.TextMatrix(i, 3)), CCur(Me.fePasivos.TextMatrix(i, 4)))
                    End If
                End If
            Next i
        End If
        
        '-- pasivos det DETALLE formato6
        If UBound(lvPrincipalPasivos) > 0 Then
            For i = 1 To UBound(lvPrincipalPasivos)
                Me.Caption = Me.Caption + cPunto
                If lvPrincipalPasivos(i).nDetPP > 0 Then
                    For j = 1 To UBound(lvPrincipalPasivos(i).vPP)
                        Me.Caption = Me.Caption + cPunto
                    
                        lcCodifi = IIf(CInt(Me.fePasivos.TextMatrix(i, 6)) = 109 Or CInt(Me.fePasivos.TextMatrix(i, 6)) = 201, Right(lvPrincipalPasivos(i).vPP(j).CDescripcion, 8), "")
                        
                        Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
                            sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, _
                            CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), 1, _
                            Format(lvPrincipalPasivos(i).vPP(j).dFecha, "yyyyMMdd"), _
                            lvPrincipalPasivos(i).vPP(j).CDescripcion, _
                            lvPrincipalPasivos(i).vPP(j).nImporte, 0, lcCodifi)
                    Next j
                End If

                If lvPrincipalPasivos(i).nDetPE > 0 Then
                    For j = 1 To UBound(lvPrincipalPasivos(i).vPE)
                        Me.Caption = Me.Caption + cPunto
                        
                        lcCodifi = IIf(CInt(Me.fePasivos.TextMatrix(i, 6)) = 109 Or CInt(Me.fePasivos.TextMatrix(i, 6)) = 201, Right(lvPrincipalPasivos(i).vPE(j).CDescripcion, 8), "")
                        
                        Call oNCred.AgregaCredFormEvalEstFinActivosPasivosFormato6DetDetalle( _
                            sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, _
                            CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), 2, _
                            Format(lvPrincipalPasivos(i).vPE(j).dFecha, "yyyyMMdd"), _
                            lvPrincipalPasivos(i).vPE(j).CDescripcion, 0, _
                            lvPrincipalPasivos(i).vPE(j).nImporte, lcCodifi)
                    Next j
                End If
            Next i
        End If
        
        
        '-- ESTADOS DE GANANCIAS Y PERDIDAS
        If UBound(lvPrincipalEstGanPer) > 0 Then
            For i = 1 To UBound(lvPrincipalEstGanPer)
                Me.Caption = Me.Caption + cPunto
                If Abs(CCur(Me.feEstaGananPerd.TextMatrix(i, 2))) > 0 Then
                    Call oNCred.AgregaCredFormEvalEstFinEstGanPerFormato6(sCtaCod, lnNumForm, Format(mskFecRegistro, "yyyyMMdd"), chkAudit.value, CInt(Me.feEstaGananPerd.TextMatrix(i, 3)), CInt(Me.feEstaGananPerd.TextMatrix(i, 4)), CCur(Me.feEstaGananPerd.TextMatrix(i, 2)))
                End If
            Next i
        End If

        Me.Caption = cCaptionOriginal
                
        MsgBox "Se guardaron los datos ingresados satisfactoriamente.", vbInformation, "Atención"

        ReDim lvPrincipalActivos(0)
        ReDim lvPrincipalPasivos(0)
        ReDim lvPrincipalEstGanPer(0)
        
        Unload Me

End Sub

Private Sub feActivos_EnterCell()
    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0)) 'celda que se activa el textbuscar
        Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
            Me.feActivos.ListaControles = "0-0-1-1-0-0-0"
        Case Else
            Me.feActivos.ListaControles = "0-0-0-0-0-0-0"
        End Select

    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0)) 'celda que  no se puede editar
        Case 1, 10, 17
            Me.feActivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
End Sub

Private Sub feActivos_OnCellChange(pnRow As Long, pnCol As Long)
    
    If IsNumeric(feActivos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        
        Select Case CInt(feActivos.TextMatrix(pnRow, 0))
            Case 15 'negativos
                If CCur(feActivos.TextMatrix(pnRow, pnCol)) > 0 Then
                    feActivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feActivos.TextMatrix(pnRow, pnCol))) * -1, "#,#0.00")   '"0.00"
                End If
            Case 5 ' posi o negativo
            
            Case Else
                If CCur(feActivos.TextMatrix(pnRow, pnCol)) < 0 Then
                    feActivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feActivos.TextMatrix(pnRow, pnCol))), "#,#0.00")  '"0.00"
                End If
        End Select
    Else
        feActivos.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    Call CalculaCeldas(1)
    Call CalculaCeldas(2)

End Sub

Private Sub feActivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)

Dim pnMonto As Double
'Dim lvDetalleEstFin() As tForEvalEstFinFormato6 'matriz para activos 'LUCV20171015, Comentó
Dim lvDetalleEstFin() As tFormEvalDetalleEstFinFormato6 'LUCV20171015, Agregó
Dim Index As Integer
Dim nTotal As Double

If mskFecRegistro.Text = "__/__/____" Then
    MsgBox "Ingrese una fecha de Estados Financieros.", vbOKOnly, "Atención"
    Exit Sub
End If
       
    If feActivos.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = CInt(feActivos.TextMatrix(feActivos.row, 0))
    nTotal = 0
    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
        Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
            Set oFrm6 = New frmCredFormEvalDetalleFormato6
            
            If feActivos.Col = 2 Then 'column P.P.
                If IsArray(lvPrincipalActivos(Index).vPP) Then
                    lvDetalleEstFin = lvPrincipalActivos(Index).vPP
                    nTotal = lvPrincipalActivos(Index).nImportePP
                Else
                    ReDim lvDetalleEstFin(0)
                End If
            End If
            
            If feActivos.Col = 3 Then 'column P.P.
                If IsArray(lvPrincipalActivos(Index).vPE) Then
                    lvDetalleEstFin = lvPrincipalActivos(Index).vPE
                    nTotal = lvPrincipalActivos(Index).nImportePE
                Else
                    ReDim lvDetalleEstFin(0)
                End If
            End If

            If oFrm6.Registrar(True, 1, lvPrincipalActivos(Index).cConcepto, lvDetalleEstFin, lvDetalleEstFin, nTotal, CInt(feActivos.TextMatrix(Me.feActivos.row, 5)), CInt(feActivos.TextMatrix(Me.feActivos.row, 6)), mskFecRegistro.Text) Then

                If feActivos.Col = 2 Then 'column P.P.
                    lvPrincipalActivos(Index).vPP = lvDetalleEstFin
                End If
                If feActivos.Col = 3 Then ' columna P.E.
                    lvPrincipalActivos(Index).vPE = lvDetalleEstFin
                End If
                
            End If
            
            If feActivos.Col = 2 Then
                
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.Col) = Format(nTotal, "#,#0.00")
                    
                If nTotal <> 0 Then
                    lvPrincipalActivos(Index).nImportePP = nTotal
                    lvPrincipalActivos(Index).nDetPP = 1
                Else
                    lvPrincipalActivos(Index).nImportePP = nTotal
                    lvPrincipalActivos(Index).nDetPP = 0
                End If
            End If

            If feActivos.Col = 3 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalActivos(Index).nImportePE = nTotal
                    lvPrincipalActivos(Index).nDetPE = 1
                Else
                    lvPrincipalActivos(Index).nImportePE = nTotal
                    lvPrincipalActivos(Index).nDetPE = 0
                End If
            End If

            Call CalculaCeldas(1)
            Call CalculaCeldas(2)
            
        End Select

End Sub

Private Sub CalculaCeldas(pnActPas As Integer)
    Dim m1, m2 As Double
    Dim s1, s2, s3, s4 As Double 'para pasivos y activos
    Dim s5, s6, s7, s8, s9 As Double 'para est gana y perdi

    Dim lnTotActivo1 As Double
    Dim lnTotActivo2 As Double
    Dim lnTotPasivo1 As Double
    Dim lnTotPasivo2 As Double
    Dim lnResulEjer1 As Double
    Dim lnResulEjer2 As Double
    Dim lnResulAcum1 As Double
    Dim lnResulAcum2 As Double
    Dim lnCapiAdici1 As Double
    Dim lnExceReval1 As Double
    Dim lnReservaLe1 As Double
    Dim lnCapiAdici2 As Double
    Dim lnExceReval2 As Double
    Dim lnReservaLe2 As Double

    s1 = 0: s2 = 0: s3 = 0: s4 = 0
    s5 = 0: s6 = 0: s7 = 0: s8 = 0: s9 = 0

    If pnActPas = 1 Then '-- activos
            ' valida que todos los registros sean numeros
            For i = 1 To feActivos.rows - 1
                If Not IsNumeric(Me.feActivos.TextMatrix(i, 2)) Then Me.feActivos.TextMatrix(i, 2) = "0.00"
                If Not IsNumeric(Me.feActivos.TextMatrix(i, 3)) Then Me.feActivos.TextMatrix(i, 3) = "0.00"
            Next i


            For i = 2 To 9
                s1 = s1 + CCur(Me.feActivos.TextMatrix(i, 2))
                s2 = s2 + CCur(Me.feActivos.TextMatrix(i, 3))
            Next i

            Me.feActivos.TextMatrix(1, 2) = Format(s1, "#,#0.00")
            Me.feActivos.TextMatrix(1, 3) = Format(s2, "#,#0.00")

            For i = 11 To 16
                s3 = s3 + CCur(Me.feActivos.TextMatrix(i, 2))
                s4 = s4 + CCur(Me.feActivos.TextMatrix(i, 3))
            Next i

            Me.feActivos.TextMatrix(10, 2) = Format(s3, "#,#0.00")
            Me.feActivos.TextMatrix(10, 3) = Format(s4, "#,#0.00")

            Me.feActivos.TextMatrix(17, 2) = Format(s1 + s3, "#,#0.00")
            Me.feActivos.TextMatrix(17, 3) = Format(s2 + s4, "#,#0.00")

            '-- TOTALIZA TOTAL PATRIMONIO EN PP y PE (TOT ACTIVO - TOT PASIVO)
            Me.fePasivos.TextMatrix(24, 2) = Format(CDbl(Me.feActivos.TextMatrix(17, 2)) - CDbl(Me.fePasivos.TextMatrix(23, 2)), "#,#0.00")
            Me.fePasivos.TextMatrix(24, 3) = Format(CDbl(Me.feActivos.TextMatrix(17, 3)) - CDbl(Me.fePasivos.TextMatrix(23, 3)), "#,#0.00")

            '-- TOTALIZA TOTAL PASIVO Y PATRIMONIO EN PP y PE (TOT PASIVO + TOT PATRIMONIO)
            Me.fePasivos.TextMatrix(25, 2) = Format(CDbl(Me.fePasivos.TextMatrix(23, 2)) + CDbl(Me.fePasivos.TextMatrix(24, 2)), "#,#0.00")
            Me.fePasivos.TextMatrix(25, 3) = Format(CDbl(Me.fePasivos.TextMatrix(23, 3)) + CDbl(Me.fePasivos.TextMatrix(24, 3)), "#,#0.00")

            For i = 1 To Me.feActivos.rows - 1

                m1 = CCur(Me.feActivos.TextMatrix(i, 2))
                m2 = CCur(Me.feActivos.TextMatrix(i, 3))

                Me.feActivos.TextMatrix(i, 4) = Format(m1 + m2, "#,#0.00")
                Me.feActivos.TextMatrix(i, 2) = Format(m1, "#,#0.00")
                Me.feActivos.TextMatrix(i, 3) = Format(m2, "#,#0.00")
            Next i
    End If

    s1 = 0: s2 = 0: s3 = 0: s4 = 0
    s5 = 0: s6 = 0: s7 = 0: s8 = 0: s9 = 0

    If pnActPas = 2 Then '-- pasivos

        lnTotActivo1 = CCur(Me.feActivos.TextMatrix(17, 2))
        lnTotActivo2 = CCur(Me.feActivos.TextMatrix(17, 3))
        lnTotPasivo1 = CCur(Me.fePasivos.TextMatrix(23, 2))
        lnTotPasivo2 = CCur(Me.fePasivos.TextMatrix(23, 3))
        
        'lnCapiAdici1-lnExceReval1-lnReservaLe1
        'lnCapiAdici2-lnExceReval2-lnReservaLe2
        
        lnCapiAdici1 = CCur(Me.fePasivos.TextMatrix(18, 2))
        lnCapiAdici2 = CCur(Me.fePasivos.TextMatrix(18, 3))
        lnExceReval1 = CCur(Me.fePasivos.TextMatrix(19, 2))
        lnExceReval2 = CCur(Me.fePasivos.TextMatrix(19, 3))
        lnReservaLe1 = CCur(Me.fePasivos.TextMatrix(20, 2))
        lnReservaLe2 = CCur(Me.fePasivos.TextMatrix(20, 3))
        
        lnResulEjer1 = CCur(Me.fePasivos.TextMatrix(21, 2))
        lnResulEjer2 = CCur(Me.fePasivos.TextMatrix(21, 3))
        lnResulAcum1 = CCur(Me.fePasivos.TextMatrix(22, 2))
        lnResulAcum2 = CCur(Me.fePasivos.TextMatrix(22, 3))

        Me.fePasivos.TextMatrix(17, 2) = lnTotActivo1 - lnTotPasivo1 - lnResulEjer1 - lnResulAcum1 - lnCapiAdici1 - lnExceReval1 - lnReservaLe1
        Me.fePasivos.TextMatrix(17, 3) = lnTotActivo2 - lnTotPasivo2 - lnResulEjer2 - lnResulAcum2 - lnCapiAdici2 - lnExceReval2 - lnReservaLe2

        ' valida que todos los registros sean numeros
        For i = 1 To fePasivos.rows - 1
            If Not IsNumeric(Me.fePasivos.TextMatrix(i, 2)) Then Me.fePasivos.TextMatrix(i, 2) = "0.00"
            If Not IsNumeric(Me.fePasivos.TextMatrix(i, 3)) Then Me.fePasivos.TextMatrix(i, 3) = "0.00"
        Next i
        '----------------------------------------------------------------
        For i = 2 To 9
            s1 = s1 + CCur(Me.fePasivos.TextMatrix(i, 2))
            s2 = s2 + CCur(Me.fePasivos.TextMatrix(i, 3))
        Next i
        '-- TOTALIZA PASIVO CORRIENTE EN PP y PE
        Me.fePasivos.TextMatrix(1, 2) = Format(s1, "#,#0.00")
        Me.fePasivos.TextMatrix(1, 3) = Format(s2, "#,#0.00")
        '----------------------------------------------------------------
        For i = 11 To 15
            s3 = s3 + CCur(Me.fePasivos.TextMatrix(i, 2))
            s4 = s4 + CCur(Me.fePasivos.TextMatrix(i, 3))
        Next i

         '-- TOTALIZA PASIVO NO CORRIENTE EN PP y PE
         Me.fePasivos.TextMatrix(10, 2) = Format(s3, "#,#0.00")
         Me.fePasivos.TextMatrix(10, 3) = Format(s4, "#,#0.00")
         '----------------------------------------------------------------
         For i = 17 To 22
             s5 = s5 + CCur(Me.fePasivos.TextMatrix(i, 2))
             s6 = s6 + CCur(Me.fePasivos.TextMatrix(i, 3))
         Next i

         '-- TOTALIZA PATRIMONIO EN PP y PE
         Me.fePasivos.TextMatrix(16, 2) = Format(s5, "#,#0.00")
         Me.fePasivos.TextMatrix(16, 3) = Format(s6, "#,#0.00")

         '----------------------------------------------------------------

         '-- TOTALIZA TOTAL PASIVO EN PP y PE (PAS CORR + PAS NO CORR)
         Me.fePasivos.TextMatrix(23, 2) = Format(s1 + s3, "#,#0.00")
         Me.fePasivos.TextMatrix(23, 3) = Format(s2 + s4, "#,#0.00")

         '-- TOTALIZA TOTAL PATRIMONIO EN PP y PE (TOT ACTIVO - TOT PASIVO)
'            Me.fePasivos.TextMatrix(24, 2) = Format(ccur(Me.feActivos.TextMatrix(17, 2)) - ccur(Me.fePasivos.TextMatrix(16, 2)), "#,#0.00")
'            Me.fePasivos.TextMatrix(24, 3) = Format(ccur(Me.feActivos.TextMatrix(17, 3)) - ccur(Me.fePasivos.TextMatrix(16, 3)), "#,#0.00")
         Me.fePasivos.TextMatrix(24, 2) = Format(s5, "#,#0.00")
         Me.fePasivos.TextMatrix(24, 3) = Format(s6, "#,#0.00")


         '-- TOTALIZA TOTAL PASIVO Y PATRIMONIO EN PP y PE (TOT PASIVO + TOT PATRIMONIO)
         Me.fePasivos.TextMatrix(25, 2) = Format(CCur(Me.fePasivos.TextMatrix(23, 2)) + CCur(Me.fePasivos.TextMatrix(24, 2)), "#,#0.00")
         Me.fePasivos.TextMatrix(25, 3) = Format(CCur(Me.fePasivos.TextMatrix(23, 3)) + CCur(Me.fePasivos.TextMatrix(24, 3)), "#,#0.00")

        '----------------------------------------------------------------
        For i = 1 To Me.fePasivos.rows - 1 '-- para columna TOTAL
            m1 = CCur(Me.fePasivos.TextMatrix(i, 2))
            m2 = CCur(Me.fePasivos.TextMatrix(i, 3))

            Me.fePasivos.TextMatrix(i, 4) = Format(m1 + m2, "#,#0.00")
            Me.fePasivos.TextMatrix(i, 2) = Format(m1, "#,#0.00")
            Me.fePasivos.TextMatrix(i, 3) = Format(m2, "#,#0.00")
        Next i
    End If

    s1 = 0: s2 = 0: s3 = 0: s4 = 0
    s5 = 0: s6 = 0: s7 = 0: s8 = 0: s9 = 0

    If pnActPas = 3 Then '-- EST DE GANACIAS Y PERDIDAS
            ' valida que todos los registros sean numeros
            For i = 1 To feEstaGananPerd.rows - 1
                If Not IsNumeric(Me.feEstaGananPerd.TextMatrix(i, 2)) Then Me.feEstaGananPerd.TextMatrix(i, 2) = "0.00"
            Next i
            '-------------------total ing brutos
            For i = 1 To 2
                s5 = s5 + CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
            Next i

            '-------------------utilidad bruta
            For i = 4 To 5
                s6 = s6 + CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
            Next i
            '-------------------utilidad operativa
            For i = 7 To 8
                s7 = s7 + CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
            Next i
            '-------------------utilidad antes de P. y D.E. e imp renta
            For i = 10 To 15
                If (i = 12 Or i = 14) Then
                    s8 = s8 - CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
                Else
                    s8 = s8 + CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
                End If
            Next i
            '-------------------utilidad(perdida del ejercicio)
            For i = 17 To 18
                s9 = s9 + CCur(Me.feEstaGananPerd.TextMatrix(i, 2))
            Next i

            Me.feEstaGananPerd.TextMatrix(3, 2) = Format(s5, "#,#0.00")
            Me.feEstaGananPerd.TextMatrix(6, 2) = Format(s5 - s6, "#,#0.00")
            Me.feEstaGananPerd.TextMatrix(9, 2) = Format(s5 - s6 - s7, "#,#0.00")
            Me.feEstaGananPerd.TextMatrix(16, 2) = Format((s5 - s6 - s7) + s8, "#,#0.00")
            Me.feEstaGananPerd.TextMatrix(19, 2) = Format(((s5 - s6 - s7) + s8) - s9, "#,#0.00")

'            Me.feEstaGananPerd.TextMatrix(3, 2) = Format(s5, "#,#0.00")
'            Me.feEstaGananPerd.TextMatrix(6, 2) = Format(s5 + s6, "#,#0.00")
'            Me.feEstaGananPerd.TextMatrix(9, 2) = Format(s5 + s6 + s7, "#,#0.00")
'            Me.feEstaGananPerd.TextMatrix(16, 2) = Format(s8 + (s5 + s6 + s7), "#,#0.00")
'            Me.feEstaGananPerd.TextMatrix(19, 2) = Format(s9 + s8 + (s5 + s6 + s7), "#,#0.00")

'            Me.feEstaGananPerd.TextMatrix(3, 2) = Format(s5, "#,#0.00")
'            Me.feEstaGananPerd.TextMatrix(6, 2) = Format(s5 - s6, "#,#0.00")
'            Me.feEstaGananPerd.TextMatrix(9, 2) = Format(s5 - s6 - s7, "#,#0.00")
'            Me.feEstaGananPerd.TextMatrix(16, 2) = Format(s8, "#,#0.00")
'            Me.feEstaGananPerd.TextMatrix(19, 2) = Format(s9, "#,#0.00")

    End If

End Sub

Private Sub feActivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
    
    Editar = Split(Me.feActivos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        
        SendKeys "{TAB}"
        
        Exit Sub
    End If
            
    Select Case pnRow
        Case 2, 4, 5, 6, 7, 8, 9, 11, 14, 16
            Cancel = False
            SendKeys "{TAB}"
            Exit Sub
    End Select
    
End Sub

Private Sub feEstaGananPerd_EnterCell()

'    Select Case CInt(Me.fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'celda que se activa el textbuscar
'        Case 2, 5, 6, 7, 9
'            Me.fePasivos.ListaControles = "0-0-1-1-0-0-0"
'        Case Else
'            Me.fePasivos.ListaControles = "0-0-0-0-0-0-0"
'        End Select

    Select Case CInt(feEstaGananPerd.TextMatrix(Me.feEstaGananPerd.row, 0)) 'celda que  o se puede editar
        Case 3, 6, 9, 16, 19
            Me.feEstaGananPerd.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feEstaGananPerd.ColumnasAEditar = "X-X-2-X-X"
        End Select

End Sub

Private Sub feEstaGananPerd_OnCellChange(pnRow As Long, pnCol As Long)
'   If IsNumeric(Me.feEstaGananPerd.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
'        Select Case CInt(feEstaGananPerd.TextMatrix(pnRow, 0))
'            Case 1, 2, 11, 13 'positivos
'                If feEstaGananPerd.TextMatrix(pnRow, pnCol) < 0 Then
'                    feEstaGananPerd.TextMatrix(pnRow, pnCol) = Format(Abs(val(feEstaGananPerd.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
'                End If
'            Case 4, 5, 7, 8, 12, 14 'negativos
'                If feEstaGananPerd.TextMatrix(pnRow, pnCol) > 0 Then
'                    feEstaGananPerd.TextMatrix(pnRow, pnCol) = Format(Abs(val(feEstaGananPerd.TextMatrix(pnRow, pnCol))) * -1, "#,#0.00") '"0.00"
'                End If
'        End Select
'    Else
'        feEstaGananPerd.TextMatrix(pnRow, pnCol) = "0.00"
'    End If
'    Call CalculaCeldas(3)
    
    
    If IsNumeric(feEstaGananPerd.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
'        Select Case CInt(feEstaGananPerd.TextMatrix(pnRow, 0))
'            Case 1, 2, 4, 5, 7, 8, 11, 13  'positivos
                If feEstaGananPerd.TextMatrix(pnRow, pnCol) < 0 Then
                    feEstaGananPerd.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feEstaGananPerd.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
                End If
'            Case 12, 14 'negativos
'                If feEstaGananPerd.TextMatrix(pnRow, pnCol) > 0 Then
'                    feEstaGananPerd.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feEstaGananPerd.TextMatrix(pnRow, pnCol))) * -1, "#,#0.00") '"0.00"
'                End If
'        End Select
    Else
        feEstaGananPerd.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    Call CalculaCeldas(3)

End Sub

Private Sub feEstaGananPerd_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'    Call CalculaCeldas(3)
End Sub

Private Sub feEstaGananPerd_RowColChange()
    If feEstaGananPerd.Col = 2 Then
        feEstaGananPerd.AvanceCeldas = Vertical
    Else
        feEstaGananPerd.AvanceCeldas = Horizontal
    End If

End Sub

Private Sub fePasivos_EnterCell()
    
    Select Case CInt(Me.fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'celda que se activa el textbuscar
        Case 2, 5, 6, 7, 9, 11
            Me.fePasivos.ListaControles = "0-0-1-1-0-0-0"
        Case Else
            Me.fePasivos.ListaControles = "0-0-0-0-0-0-0"
        End Select
        
    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'celda que no se puede editar
        Case 1, 10, 16, 17, 23, 24, 25
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
        
End Sub

Private Sub fePasivos_OnCellChange(pnRow As Long, pnCol As Long)
    
    If IsNumeric(fePasivos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos

        Select Case CInt(fePasivos.TextMatrix(pnRow, 0))
            Case 21, 22 ' ingresa positivos o negativos
                fePasivos.TextMatrix(pnRow, pnCol) = Format(CCur(fePasivos.TextMatrix(pnRow, pnCol)), "#,#0.00")   '"0.00"
            'Case 4, 5, 7, 8, 12, 14 'negativos
            Case Else
                If fePasivos.TextMatrix(pnRow, pnCol) < 0 Then
                    fePasivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(fePasivos.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
                End If
        End Select
    Else
        fePasivos.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    Call CalculaCeldas(1)
    Call CalculaCeldas(2)
End Sub

Private Sub fePasivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)

Dim pnMonto As Double
'Dim lvDetalleEstFinPasivos() As tForEvalEstFinFormato6 'matriz para activos 'LUCV20171015, Comentó
Dim lvDetalleEstFinPasivos() As tFormEvalDetalleEstFinFormato6 'LUCV20171015, Agregó
Dim Index As Integer
Dim nTotal As Double

If mskFecRegistro.Text = "__/__/____" Then
    MsgBox "Ingrese una fecha de Estados Financieros.", vbOKOnly, "Atención"
    Exit Sub
End If
       
    If Me.fePasivos.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = CInt(fePasivos.TextMatrix(fePasivos.row, 0))
    nTotal = 0
    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
        Case 2, 5, 6, 7
            Set oFrm6 = New frmCredFormEvalDetalleFormato6
            
            If fePasivos.Col = 2 Then 'column P.P.
                If IsArray(lvPrincipalPasivos(Index).vPP) Then
                    lvDetalleEstFinPasivos = lvPrincipalPasivos(Index).vPP
                    nTotal = lvPrincipalPasivos(Index).nImportePP
                Else
                    ReDim lvDetalleEstFinPasivos(0)
                End If
            End If
            
            If fePasivos.Col = 3 Then 'column P.E.
                If IsArray(lvPrincipalPasivos(Index).vPE) Then
                    lvDetalleEstFinPasivos = lvPrincipalPasivos(Index).vPE
                    nTotal = lvPrincipalPasivos(Index).nImportePE
                Else
                    ReDim lvDetalleEstFinPasivos(0)
                End If
            End If

            If oFrm6.Registrar(True, 1, lvPrincipalPasivos(Index).cConcepto, lvDetalleEstFinPasivos, lvDetalleEstFinPasivos, nTotal, CInt(fePasivos.TextMatrix(Me.fePasivos.row, 5)), CInt(fePasivos.TextMatrix(Me.fePasivos.row, 6)), mskFecRegistro.Text) Then
                If fePasivos.Col = 2 Then 'column P.P.
                    lvPrincipalPasivos(Index).vPP = lvDetalleEstFinPasivos
                End If
                If fePasivos.Col = 3 Then ' columna P.E.
                    lvPrincipalPasivos(Index).vPE = lvDetalleEstFinPasivos
                End If
            End If
            
            If fePasivos.Col = 2 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalPasivos(Index).nImportePP = nTotal
                    lvPrincipalPasivos(Index).nDetPP = 1
                Else
                    lvPrincipalPasivos(Index).nImportePP = nTotal
                    lvPrincipalPasivos(Index).nDetPP = 0
                End If
                
            End If
            
            If fePasivos.Col = 3 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalPasivos(Index).nImportePE = nTotal
                    lvPrincipalPasivos(Index).nDetPE = 1
                Else
                    lvPrincipalPasivos(Index).nImportePE = nTotal
                    lvPrincipalPasivos(Index).nDetPE = 0
                End If
                
            End If

            Call CalculaCeldas(2)

        Case 9, 11 'detalle de Ifis

            If fePasivos.Col = 2 Then 'column P.P.
                If IsArray(lvPrincipalPasivos(Index).vPP) Then
                    lvDetalleEstFinPasivos = lvPrincipalPasivos(Index).vPP
                    nTotal = lvPrincipalPasivos(Index).nImportePP
                Else
                    ReDim lvDetalleEstFinPasivos(0)
                End If
            End If
            
            If fePasivos.Col = 3 Then 'column P.E.
                If IsArray(lvPrincipalPasivos(Index).vPE) Then
                    lvDetalleEstFinPasivos = lvPrincipalPasivos(Index).vPE
                    nTotal = lvPrincipalPasivos(Index).nImportePE
                Else
                    ReDim lvDetalleEstFinPasivos(0)
                End If
            End If

            If frmCredFormEvalIfisDetalleFormato6.Registrar(True, 1, lvPrincipalPasivos(Index).cConcepto, lvDetalleEstFinPasivos, lvDetalleEstFinPasivos, nTotal, CInt(fePasivos.TextMatrix(Me.fePasivos.row, 5)), CInt(fePasivos.TextMatrix(Me.fePasivos.row, 6)), Me.mskFecRegistro.Text) Then
                If fePasivos.Col = 2 Then 'column P.P.
                    lvPrincipalPasivos(Index).vPP = lvDetalleEstFinPasivos
                End If
                If fePasivos.Col = 3 Then ' columna P.E.
                    lvPrincipalPasivos(Index).vPE = lvDetalleEstFinPasivos
                End If

            End If
            
            If fePasivos.Col = 2 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalPasivos(Index).nImportePP = nTotal
                    lvPrincipalPasivos(Index).nDetPP = 1
                Else
                    lvPrincipalPasivos(Index).nImportePP = nTotal
                    lvPrincipalPasivos(Index).nDetPP = 0
                End If
                Call CalculaCeldas(2)
            End If
            
            If fePasivos.Col = 3 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                If nTotal <> 0 Then
                    lvPrincipalPasivos(Index).nImportePE = nTotal
                    lvPrincipalPasivos(Index).nDetPE = 1
                Else
                    lvPrincipalPasivos(Index).nImportePE = nTotal
                    lvPrincipalPasivos(Index).nDetPE = 0
                End If
                Call CalculaCeldas(2)
            End If

            Call CalculaCeldas(2)
            Call CalculaCeldas(2)

        End Select
End Sub

Private Sub fePasivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
    
    Editar = Split(Me.fePasivos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    Select Case pnRow
        Case 2, 5, 6, 7, 9, 11
            Cancel = False
            SendKeys "{TAB}"
            Exit Sub
    End Select
    
End Sub

Private Sub Form_Load()
    DisableCloseButton Me
End Sub

Private Sub mskFecRegistro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.feActivos.SetFocus
End If
End Sub


