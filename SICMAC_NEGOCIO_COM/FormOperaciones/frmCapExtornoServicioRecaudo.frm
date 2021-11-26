VERSION 5.00
Begin VB.Form frmCapExtornoServicioRecaudo 
   Caption         =   "Extorno - Servicio Recaudo"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   Icon            =   "frmCapExtornoServicioRecaudo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   9765
      TabIndex        =   11
      Top             =   6195
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10920
      TabIndex        =   10
      Top             =   6195
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   225
      TabIndex        =   9
      Top             =   6195
      Width           =   1035
   End
   Begin VB.Frame fraMovimientos 
      Caption         =   "Movimientos"
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
      Height          =   4110
      Left            =   210
      TabIndex        =   8
      Top             =   1995
      Width           =   11745
      Begin SICMACT.FlexEdit grdListaRecaudos 
         Height          =   3435
         Left            =   210
         TabIndex        =   13
         Top             =   420
         Width           =   11310
         _extentx        =   19950
         _extenty        =   6059
         cols0           =   13
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Mov-Convenio-Operación-Cuenta-Empresa-Monto-ITF-Comisión Empresa-Comisión Cliente-Glosa-cMovNro-cCodCliente"
         encabezadosanchos=   "500-900-2200-2300-1800-2500-1200-1200-1700-1700-2500-0-0"
         font            =   "frmCapExtornoServicioRecaudo.frx":030A
         font            =   "frmCapExtornoServicioRecaudo.frx":0336
         font            =   "frmCapExtornoServicioRecaudo.frx":0362
         font            =   "frmCapExtornoServicioRecaudo.frx":038E
         font            =   "frmCapExtornoServicioRecaudo.frx":03BA
         fontfixed       =   "frmCapExtornoServicioRecaudo.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "-0---0-0---0-0--0-0"
         encabezadosalineacion=   "C-L-L-L-C-C-R-R-R-R-C-C-C"
         formatosedit    =   "-0---0-0---0-0--0-0"
         textarray0      =   "#"
         lbpuntero       =   -1
         colwidth0       =   495
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
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
      Height          =   1770
      Left            =   8925
      TabIndex        =   6
      Top             =   105
      Width           =   3015
      Begin VB.TextBox txtGlosa 
         Height          =   1275
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   2565
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Busqueda"
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
      Height          =   1770
      Left            =   210
      TabIndex        =   1
      Top             =   105
      Width           =   8610
      Begin VB.Frame pnlBusquedaMovimiento 
         Height          =   1305
         Left            =   2350
         TabIndex        =   3
         Top             =   240
         Width           =   6120
         Begin VB.Frame Frame3 
            Height          =   645
            Left            =   210
            TabIndex        =   15
            Top             =   315
            Width           =   4530
            Begin VB.TextBox txtConvenio 
               Height          =   330
               Left            =   1890
               TabIndex        =   0
               Top             =   210
               Width           =   2445
            End
            Begin VB.Label lblBuscar 
               Caption         =   "Nro Movimiento: "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   105
               TabIndex        =   16
               Top             =   250
               Width           =   1410
            End
         End
         Begin VB.CommandButton btnBuscarMovimiento 
            Caption         =   "Buscar"
            CausesValidation=   0   'False
            Height          =   375
            Left            =   4830
            TabIndex        =   5
            Top             =   525
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1305
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2145
         Begin VB.OptionButton rbNombreConvenio 
            Caption         =   "&Nombre Convenio"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton rbCodConvenio 
            Caption         =   "Có&digo Convenio"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   495
            Width           =   1935
         End
         Begin VB.OptionButton rbNumMovimiento 
            Caption         =   "Número &Movimiento"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   135
            Value           =   -1  'True
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmCapExtornoServicioRecaudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBuscarMovimiento_Click()

    Dim rs As Recordset
    Set rs = New Recordset
    
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    DoEvents
    
    If rbNumMovimiento.value Then
    
        Set rs = ClsServicioRecaudo.listaMovExtornoServicioRecaudo(Year(gdFecSis) & _
                                                              IIf(Month(gdFecSis) < 10, "0" & Month(gdFecSis), Month(gdFecSis)) & _
                                                               IIf(Day(gdFecSis) < 10, "0" & Day(gdFecSis), Day(gdFecSis)) & "%", Trim(txtConvenio.Text))
    ElseIf rbCodConvenio.value Then
    
        Set rs = ClsServicioRecaudo.listaMovExtornoServicioRecaudo(Year(gdFecSis) & _
                                                               IIf(Month(gdFecSis) < 10, "0" & Month(gdFecSis), Month(gdFecSis)) & _
                                                               IIf(Day(gdFecSis) < 10, "0" & Day(gdFecSis), Day(gdFecSis)) & "%", , _
                                                               ("%" & Trim(txtConvenio.Text) & "%"))
    
    ElseIf rbNombreConvenio.value Then
    
        Set rs = ClsServicioRecaudo.listaMovExtornoServicioRecaudo(Year(gdFecSis) & _
                                                               IIf(Month(gdFecSis) < 10, "0" & Month(gdFecSis), Month(gdFecSis)) & _
                                                               IIf(Day(gdFecSis) < 10, "0" & Day(gdFecSis), Day(gdFecSis)) & "%", , , _
                                                               ("%" & Trim(txtConvenio.Text) & "%"))
    
    End If
    
    

    grdListaRecaudos.Clear
    grdListaRecaudos.FormaCabecera
    grdListaRecaudos.Rows = 2
        
    If rs.RecordCount = 0 Then
        MsgBox "No existen movimientos a ser extornados", vbInformation, "Aviso"
        Exit Sub
    End If
        
    Do While Not (rs.BOF Or rs.EOF)
        
        grdListaRecaudos.AdicionaFila
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 1) = rs!nMovNro
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 2) = rs!cCodConvenio
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 3) = rs!Concepto
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 4) = rs!cCtaCod
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 5) = rs!cPersNombre
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 6) = Format(rs!Monto, "#,##0.00")
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 7) = Format(rs!ITF, "#,##0.00")
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 8) = Format(rs!comision, "#,##0.00")
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 9) = Format(rs!ComiCliente, "#,##0.00")
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 10) = rs!Glosa
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 11) = rs!cMovNro
        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 12) = rs!cCodCliente 'CTI1 ERS027-2019
        rs.MoveNext
        
    Loop
    
End Sub


'Private Sub btnBuscarConvenio_Click()
'
'    Dim rs As Recordset
'    Set rs = New Recordset
'
'    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
'
'    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
'
'    Set rs = ClsServicioRecaudo.listaMovExtornoServicioRecaudo(Year(gdFecSis) & _
'                                                               IIf(Month(gdFecSis) < 10, "0" & Month(gdFecSis), Month(gdFecSis)) & _
'                                                               IIf(Day(gdFecSis) < 10, "0" & Day(gdFecSis), Day(gdFecSis)) & "%", , _
'                                                               ("%" & Trim(txtConvenio.Text) & "%"))
'    grdListaRecaudos.Clear
'    grdListaRecaudos.FormaCabecera
'    grdListaRecaudos.Rows = 2
'
'    If rs.RecordCount = 0 Then
'        MsgBox "No existen movimientos a ser extornados", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    Do While Not (rs.BOF Or rs.EOF)
'        grdListaRecaudos.AdicionaFila
'
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 1) = rs!nMovNro
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 2) = rs!cCodConvenio
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 3) = rs!Concepto
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 4) = rs!cCtaCod
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 5) = rs!cPersNombre
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 6) = Format(rs!Monto, "#,##0.00")
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 7) = Format(rs!ITF, "#,##0.00")
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 8) = Format(rs!comision, "#,##0.00")
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 9) = Format(rs!ComiCliente, "#,##0.00")
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 10) = rs!Glosa
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 11) = rs!cMovNro
'        rs.MoveNext
'
'    Loop
'
'End Sub

'Private Sub btnBuscarNombreConvenio_Click()
'
'    Dim rs As Recordset
'    Set rs = New Recordset
'
'    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
'
'    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
'
'    Set rs = ClsServicioRecaudo.listaMovExtornoServicioRecaudo(Year(gdFecSis) & _
'                                                               IIf(Month(gdFecSis) < 10, "0" & Month(gdFecSis), Month(gdFecSis)) & _
'                                                               IIf(Day(gdFecSis) < 10, "0" & Day(gdFecSis), Day(gdFecSis)) & "%", , _
'                                                               ("%" & Trim(txtConvenio.Text) & "%"), Trim(txtNombreConvenio.Text))
'    grdListaRecaudos.Clear
'    grdListaRecaudos.FormaCabecera
'    grdListaRecaudos.Rows = 2
'
'    If rs.RecordCount = 0 Then
'        MsgBox "No existen movimientos a ser extornados", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    Do While Not (rs.BOF Or rs.EOF)
'        grdListaRecaudos.AdicionaFila
'
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 1) = rs!nMovNro
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 2) = rs!cCodConvenio
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 3) = rs!Concepto
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 4) = rs!cCtaCod
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 5) = rs!cPersNombre
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 6) = Format(rs!Monto, "#,##0.00")
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 7) = Format(rs!ITF, "#,##0.00")
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 8) = Format(rs!comision, "#,##0.00")
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 9) = Format(rs!ComiCliente, "#,##0.00")
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 10) = rs!Glosa
'        grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 11) = rs!cMovNro
'        rs.MoveNext
'
'    Loop
'
'End Sub

Private Sub cmdCancelar_Click()

    rbNumMovimiento.value = True
    txtConvenio.Text = ""
    grdListaRecaudos.Clear
    grdListaRecaudos.FormaCabecera
    grdListaRecaudos.Rows = 2
    
End Sub

Private Sub cmdExtornar_Click()
    
    Dim nMonto As Double ' Monto depositado por el cliente
    Dim nComiEmp As Double ' Comision de parte de la empresa
    Dim nITF As Double ' Itf cargado a la cuenta de la empresa
    Dim nmoneda As Integer ' moneda de la cuenta
    Dim cCuenta As String ' Cuenta del convenio
    Dim nMovNro, nMovNroUltimo As Double ' Numero de movimiento
    Dim cMovNro, cMovNro2 As String ' Numero de cMovNro
    Dim nFila As Double ' Fila
    Dim sGlosa As String ' Glosa
    Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento 'Para validar el estado de la cuenta
    Dim oDCOMCaptaMovimiento As COMDCaptaGenerales.DCOMCaptaMovimiento ' Para validar el saldo de la cuenta
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim loVistoElectronico As frmVistoElectronico
    Dim lbResultadoVisto As Boolean
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    'CTI1 TI-ERS027-2019
    Dim ClsServicioRecaudoWS As COMDCaptaServicios.DCOMSrvRecaudoWS
    Dim bTieneWS As Boolean
    Dim sCodConvenio As String
    Dim sCodCliente As String
    Dim nCodConvenioRecaudoWS As Integer
    Dim cUrlSimaynas As String
    Dim lsOpeCodExtornoCap As String 'CTI6 ERS0112020
    Dim lsCtaAhoExt As String 'CTI6 ERS0112020
    Dim lnMontoAhorroExt As Double 'CTI6 ERS0112020
    Dim lsOperacionDescExt As String 'CTI6 ERS0112020
    Dim lnITFAhoExt As Double 'CTI6 ERS0112020
    Dim lnMovNroAExt As Long 'CTI6 ERS0112020
    Dim lsClienteExt As String 'CTI6 ERS0112020
    'CTI1 TI-ERS027-2019
    Dim lsMovNroExt As String 'CTI6 ERS0112020
        
    On Error GoTo error
        
    If Trim(grdListaRecaudos.TextMatrix(1, 1)) = "" Then
    Exit Sub
    End If
    
    Set oCont = New COMNContabilidad.NCOMContFunciones
    Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oDCOMCaptaMovimiento = New COMDCaptaGenerales.DCOMCaptaMovimiento
              
    nMovNro = CDbl(grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 1))
    sCodConvenio = grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 2) 'CTI1 TI-ERS027-2019
    sCodCliente = grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 12) 'CTI1 TI-ERS027-2019
    cCuenta = grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 4)
    nMonto = grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 6)
    nITF = grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 7)
    nComiEmp = grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 8)
    nmoneda = Mid(cCuenta, 9, 1)
    
    sGlosa = Trim(txtGlosa.Text)
    nFila = grdListaRecaudos.row

    ' Validando Estado de la cuenta
    If Not objValidar.ValidaEstadoCuenta(cCuenta, True) Then
        Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
        Dim rsCuenta As ADODB.Recordset
        Dim sEstado As String
        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCuenta = clsMant.GetDatosCuentaAho(cCuenta)
        If Not rsCuenta.EOF And Not rsCuenta.BOF Then
            If rsCuenta!nPrdEstado = 1100 Then
                sEstado = "BLOQUEO TOTAL"
            ElseIf rsCuenta!nPrdEstado = 1200 Then
                sEstado = "BLOQUEO RETIRO"
            ElseIf rsCuenta!nPrdEstado = 1300 Then
                sEstado = "ANULADO"
            ElseIf rsCuenta!nPrdEstado = 1400 Then
                sEstado = "CANCELADO"
            Else
                sEstado = "NO ACTIVO"
            End If
            MsgBox "Esta operacion no puede realizarse debido a que la cuenta se encuentra en estado: " & sEstado, vbExclamation + vbDefaultButton1, "Aviso"
            Set objValidar = Nothing
            Set clsMant = Nothing
            Set rsCuenta = Nothing
            Exit Sub
        Else
            MsgBox "Esta operacion no puede realizarse debido a que la cuenta se encuentra en estado: NO ACTIVO", vbExclamation + vbDefaultButton1, "Aviso"
            Set objValidar = Nothing
            Set clsMant = Nothing
            Set rsCuenta = Nothing
        Exit Sub
        End If
    End If
    
    ' validando saldo de la cuenta
    'GIPO añadió el argumento True 04-03-2017
     If Not objValidar.ValidaSaldoCuenta(cCuenta, nMonto - nITF - nComiEmp, , , , , , True) Then
        MsgBox "La cuenta NO Tiene saldo suficiente para la operacion", vbExclamation + vbDefaultButton1, "Aviso"
        Set objValidar = Nothing
        Exit Sub
     End If
     
    '20141203 COMENTADO POR RIRO ************
    ' Validando ultimo movimiento
    'nMovNroUltimo = oDCOMCaptaMovimiento.devolverUltimoMovimientoDeposito(cCuenta, Format(gdFecSis, "yyyyMMdd"))
    'If nMovNro > 0 And nMovNro <> nMovNroUltimo Then
    '    MsgBox "Cuenta " & cCuenta & " posee movimientos despues del depósito de recaudo, por favor extorne antes esos movimientos.", vbInformation, "Aviso"
    '    Set oDCOMCaptaMovimiento = Nothing
    '    Exit Sub
    'End If
    'END RIRO *******************************

    ' *** RIRO SEGUN TI-ERS108-2013 ***
    Dim nMovNroOperacion As Long
    nMovNroOperacion = 0
    If grdListaRecaudos.row >= 1 And Len(Trim(grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 1))) > 0 Then
        nMovNroOperacion = CLng(grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 1))
    End If
    ' *** FIN RIRO ***

    Set loVistoElectronico = New frmVistoElectronico
    lbResultadoVisto = loVistoElectronico.Inicio(3, gExtornoDepositoRecaudo, , , nMovNroOperacion)
    
    If Not lbResultadoVisto Then
        Exit Sub
    End If
    
    If MsgBox("Desea extornar el movimiento " & grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 1) & "?", vbYesNo + vbExclamation + vbDefaultButton1, "Aviso") = vbNo Then
     Exit Sub
    End If
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    cMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Sleep 1000
    cMovNro2 = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    Set ClsServicioRecaudoWS = New COMDCaptaServicios.DCOMSrvRecaudoWS 'CTI1 ERS027-2019
    bTieneWS = ClsServicioRecaudoWS.VerificarConvenioRecaudoWebService(sCodConvenio)
    If bTieneWS = True Then
        cUrlSimaynas = Trim(LeeConstanteSist(708))
    End If 'CTI1 ERS027-2019
    
    Dim loContFunct As COMNContabilidad.NCOMContFunciones 'CTI6 ERS0112020
    Set loContFunct = New COMNContabilidad.NCOMContFunciones 'CTI6 ERS0112020
    Dim lsMovNroCapExtorno As String 'CTI6 ERS0112020
    lsMovNroCapExtorno = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) 'CTI6 ERS0112020
    
    'clsCap.ExtornoPagoServicioRecaudo nComiEmp, nmoneda, cCuenta, nITF, nMovNro, cMovNro, cMovNro2, nMonto, sGlosa, _
                                      sCodConvenio, sCodCliente, bTieneWS, cUrlSimaynas 'CTI1 ERS027-2019
    If clsCap.ExtornoPagoServicioRecaudo(nComiEmp, nmoneda, cCuenta, nITF, nMovNro, cMovNro, cMovNro2, nMonto, sGlosa, _
                                      sCodConvenio, sCodCliente, bTieneWS, cUrlSimaynas, gExtornoDepositoRecaudo, _
                                      lsMovNroCapExtorno, lsOpeCodExtornoCap, lsCtaAhoExt, _
                                      lnMontoAhorroExt, lsOperacionDescExt, lnITFAhoExt, "", lsClienteExt) Then 'CTI1 ERS027-2019
    
     'CTI6 ERS0112020
    If lsOpeCodExtornoCap <> "" Then
        If (lsOpeCodExtornoCap = gAhoCargoServicioRecaudo) Then
            Dim lsCadImp As String
            Dim ClsMov As COMDMov.DCOMMov, sCodUserBusExt As String, sMovNroBusExt As String
            Set ClsMov = New COMDMov.DCOMMov
            sMovNroBusExt = "": sCodUserBusExt = ""
            sMovNroBusExt = grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 11)
            sCodUserBusExt = Right(sMovNroBusExt, 4)
                    
        
            Dim lsFechaHoraGrabExt As String
            lsFechaHoraGrabExt = fgFechaHoraGrab(lsMovNroCapExtorno)
            Dim oImp As New COMNContabilidad.NCOMContImprimir
            Set oImp = New COMNContabilidad.NCOMContImprimir
            lsClienteExt = grdListaRecaudos.TextMatrix(grdListaRecaudos.row, 5)
            lsCadImp = oImp.nPrintReciboExtorCargoCta(gsNomAge, lsFechaHoraGrabExt, "", lsCtaAhoExt, lnITFAhoExt, _
            lsClienteExt, lsOperacionDescExt, lnMontoAhorroExt, 0, lnMovNroAExt, gsCodUser, "", "", sCodUserBusExt, gImpresora, gbImpTMU)
             Do
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsCadImp & Chr$(12)
                    Print #nFicSal, ""
                Close #nFicSal
            Loop While MsgBox("Reimprimir Recibo de Extorno del Abono a Cuenta") = vbYes
        End If
    End If
    'END
    
    grdListaRecaudos.EliminaFila grdListaRecaudos.row
    txtConvenio.Text = ""
    txtGlosa.Text = ""
    MsgBox "El extorno se llevo acabo correctamente", vbInformation, "Aviso"
    Else 'CTI1 ERS027-2019
    MsgBox "No se pudo extorna correctamente" & Chr(13) & _
           "Comuniquese con el área de TI", vbExclamation, "Aviso"
    End If 'CTI1 ERS027-2019
    Exit Sub
error:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub rbCodConvenio_Click()
    lblBuscar.Caption = "Cod. Convenio: "
    txtConvenio.Text = ""
    txtConvenio.SetFocus
End Sub

Private Sub rbNombreConvenio_Click()
    lblBuscar.Caption = "Nom. Convenio: "
    txtConvenio.Text = ""
    txtConvenio.SetFocus
End Sub

Private Sub rbNumMovimiento_Click()
    lblBuscar.Caption = "Nro Movimiento: "
    txtConvenio.Text = ""
    txtConvenio.SetFocus
End Sub

Private Sub txtConvenio_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        btnBuscarMovimiento_Click
    End If
End Sub
