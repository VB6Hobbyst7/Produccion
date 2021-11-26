VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpEspecialesExt 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6855
   ClientLeft      =   2820
   ClientTop       =   2445
   ClientWidth     =   7245
   Icon            =   "frmOpEspecialesExt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
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
      ForeColor       =   &H00000080&
      Height          =   2700
      Left            =   2520
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   2845
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
         TabIndex        =   24
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFC0&
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         ItemData        =   "frmOpEspecialesExt.frx":030A
         Left            =   240
         List            =   "frmOpEspecialesExt.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   45
      Width           =   1020
   End
   Begin SICMACT.TxtBuscar txtUser 
      Height          =   330
      Left            =   810
      TabIndex        =   18
      Top             =   90
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sTitulo         =   ""
   End
   Begin VB.Frame fraOpeAExtornar 
      Caption         =   "Operaciones a Extornar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3120
      Left            =   90
      TabIndex        =   15
      Top             =   465
      Width           =   7035
      Begin MSComctlLib.ListView lstExtorno 
         Height          =   2715
         Left            =   90
         TabIndex        =   16
         Top             =   255
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   4789
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1995
      Left            =   90
      TabIndex        =   2
      Top             =   3600
      Width           =   7035
      Begin VB.TextBox txtGlosa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   165
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   975
         Width           =   6795
      End
      Begin VB.TextBox txtNroDoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin SICMACT.TxtBuscar txtPers 
         Height          =   315
         Left            =   165
         TabIndex        =   5
         Top             =   435
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         Appearance      =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   315
         Left            =   4410
         TabIndex        =   6
         Top             =   1545
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblSimbolo 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6330
         TabIndex        =   14
         Top             =   1575
         Width           =   300
      End
      Begin VB.Label lblPersona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1995
         TabIndex        =   11
         Top             =   435
         Width           =   4935
      End
      Begin VB.Label lblGlosa 
         Caption         =   "Glosa"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   765
         Width           =   915
      End
      Begin VB.Label lblDoc 
         Caption         =   "Nro.Doc"
         ForeColor       =   &H80000007&
         Height          =   225
         Left            =   225
         TabIndex        =   9
         Top             =   1590
         Width           =   735
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   3675
         TabIndex        =   8
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label lblPers 
         Caption         =   "Persona"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   210
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   350
      Left            =   6045
      TabIndex        =   1
      Top             =   6450
      Width           =   1095
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Extornar"
      Height          =   350
      Left            =   4875
      TabIndex        =   0
      Top             =   6450
      Width           =   1095
   End
   Begin VB.Frame fraExt 
      Caption         =   "Glosa Extorno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   750
      Left            =   90
      TabIndex        =   12
      Top             =   5610
      Width           =   7035
      Begin VB.TextBox txtGlosaExtorno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   420
         Left            =   165
         MaxLength       =   150
         TabIndex        =   13
         Top             =   240
         Width           =   6780
      End
   End
   Begin VB.Label lblNomUser 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1890
      TabIndex        =   19
      Top             =   90
      Width           =   4140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   165
      TabIndex        =   17
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "frmOpEspecialesExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsOpeCod As CaptacOperacion
Dim lsCaption As String
Dim lsOpeCodExt As CaptacOperacion
Private nPeriodo As Integer 'RIRO20150611 ERS162-2014

Sub CargarLista(ByVal psCodUser As String)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    lstExtorno.ListItems.Clear
   'RECO20160414*******************************************************
    Dim lsCodOpeBusc As String
    lsCodOpeBusc = lsOpeCod
    Select Case lsOpeCod
        Case 290006
            lsCodOpeBusc = "300150"
        Case 290007
            lsCodOpeBusc = "300151"
    End Select
    
    Set rs = clsCapMov.GetOtrasOperaciones(gdFecSis, psCodUser, lsCodOpeBusc) 'gsCodUser
    'Set rs = clsCapMov.GetOtrasOperaciones(gdFecSis, psCodUser, lsOpeCod) 'gsCodUser
    'RECO FIN************************************************************
    If Not RSVacio(rs) Then
       Do While Not rs.EOF
          LLenaLista rs
          rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Set clsCapMov = Nothing
End Sub

Private Sub LLenaLista(myRs As ADODB.Recordset)
    Dim itmX As ListItem
    If Not IsNull(myRs!cOpeDesc) Then
        Set itmX = lstExtorno.ListItems.Add(, , Trim(myRs!cOpeDesc))
       End If
    If Not IsNull(myRs!nMovImporte) Then
        itmX.SubItems(1) = Format$(myRs!nMovImporte, "#,##0.00")
       End If
    If Not IsNull(myRs!dFecTran) Then
        itmX.SubItems(2) = Format$(myRs!dFecTran, "dd/mm/yyyy hh:mm.ss")
    End If
    If Not IsNull(myRs!nMovNro) Then
        itmX.SubItems(3) = Trim(myRs!nMovNro)
    End If
    If Not IsNull(myRs!cOpecod) Then
        itmX.SubItems(4) = Trim(myRs!cOpecod)
    End If
    itmX.SubItems(5) = Trim(Str(myRs!nmoneda))
    
    If myRs!nmoneda = Moneda.gMonedaNacional Then
        itmX.ForeColor = 0
    Else
        itmX.ForeColor = &H8000&
        itmX.ListSubItems(1).ForeColor = &H8000&
        itmX.ListSubItems(2).ForeColor = &H8000&
        itmX.ListSubItems(3).ForeColor = &H8000&
        itmX.ListSubItems(4).ForeColor = &H8000&
        itmX.ListSubItems(5).ForeColor = &H8000&
    End If
    
End Sub
'*************************************************************************************
'********CTI3 (ferimoro)  10102018
Sub limpExt()
frmMotExtorno.Visible = False
Me.cmbMotivos.ListIndex = -1
Me.txtDetExtorno.Text = ""
fraDetalle.Enabled = True
txtUser.Enabled = True
cmdProcesar.Enabled = True
cmdAplicar.Enabled = True
End Sub
Private Sub cmdExtContinuar_Click()
    On Error GoTo Error
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsMovNro As String
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Dim lbReimp As Boolean
    Dim sBoleta As String, nResult As Integer 'RIRO 20150602 ERS162-2014
    Dim nMovNroExt As Long
    
    '***CTI3 (FERIMORO)   02102018
    Dim DatosExtorna(1) As String
    
    If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
        MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
        Exit Sub
    End If
    
    '***CTI3 (ferimoro)    02102018
    frmMotExtorno.Visible = False
    DatosExtorna(0) = cmbMotivos.Text
    DatosExtorna(1) = txtDetExtorno.Text
    txtGlosa.Text = txtDetExtorno.Text
    'JUEZ 20130906 *****************************************************************
    Dim lbResultadoVisto As Boolean
    Dim loVistoElectronico As frmVistoElectronico
    Set loVistoElectronico = New frmVistoElectronico
    
    Select Case lsOpeCodExt
        Case "300151", gComiDiversasAhoCredGasto, gComiDiversasAhoCredCom, gComiDiversasAhoCredComVoucher, gOpeTransferenciaCargo, gOpeTransferenciaDeposito
            lbResultadoVisto = loVistoElectronico.Inicio(3, lsOpeCod)
            If Not lbResultadoVisto Then
                Call limpExt
                Exit Sub
            End If
    End Select
    'END JUEZ **********************************************************************
    
    'RIRO 20150604 ERS162-2014 ******************
    If InStr(1, "300532,300533", lstExtorno.SelectedItem.ListSubItems(4)) > 0 Then
        nResult = clsCapMov.VerificaEstadoTrama(lstExtorno.SelectedItem.ListSubItems(3))
        If nResult <= 0 Then
            MsgBox "No es posible efectuar el extorno porque la trama de Utilidades fue dado de baja", vbExclamation, "Aviso"
            Set clsCapMov = Nothing
            Call limpExt 'CTI3
            Exit Sub
        End If
'        lbResultadoVisto = loVistoElectronico.Inicio(13, lsOpeCod)
'        If Not lbResultadoVisto Then
'            Exit Sub
'        End If
    End If
    'END RIRO ***********************************
    
    lsMovNro = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
     
    If MsgBox("Desea Extornar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        'RECO20160414*************************************************************
        Dim oDCOMCaptaMovimiento As New COMDCaptaGenerales.DCOMCaptaMovimiento
        Dim oSegSep As New COMNCaptaGenerales.NCOMSeguros
        Dim oRS As New ADODB.Recordset
        Dim lsNumCertif As String
        'RECO FIN*****************************************************************
        If InStr(1, "300532,300533", lstExtorno.SelectedItem.ListSubItems(4)) > 0 Then
            nMovNroExt = clsCapMov.ExtornoGrabarPagoUtilidades(lsMovNro, lstExtorno.SelectedItem.ListSubItems(4), lsOpeCod, _
                                                  lstExtorno.SelectedItem.ListSubItems(3), Trim(txtGlosa.Text), _
                                                  CDbl(txtMonto.Text), Trim(txtNroDoc.Text), gsNomAge, 0, "990201", _
                                                  sBoleta, gbImpTMU, nPeriodo, lblPersona.Caption, DatosExtorna)
        
        ElseIf lstExtorno.SelectedItem.ListSubItems(4) = gOtrOpePagoRecaudoCajeroCorresponsal Then 'RIRO20160111
            If Not clsCapMov.ExtornoPagoCC(lsMovNro, lstExtorno.SelectedItem.ListSubItems(3), lsOpeCod, Trim(txtGlosaExtorno.Text), gsNomAge, gbImpTMU, sBoleta, DatosExtorna) Then
                MsgBox "Se presentó un incoveniente durante el proceso de extorno", vbInformation, "Aviso"
                Call limpExt
                Exit Sub
            End If
        Else
            Dim psBoletaExtorno As String
            psBoletaExtorno = ""
            nMovNroExt = clsCapMov.OtrasOperacionesExtorno(lsMovNro, lstExtorno.SelectedItem.ListSubItems(3), lsOpeCod, Me.txtGlosaExtorno.Text, DatosExtorna, CStr(lstExtorno.SelectedItem.ListSubItems(4)), psBoletaExtorno, gsNomAge, gbImpTMU)
            'RECO20160414 *****************************************************
            If lsOpeCod = "290006" Then
                Set oRS = oSegSep.SepelioObtieneDatosExtorno(lstExtorno.SelectedItem.ListSubItems(3))
                
                If Not (oRS.EOF And oRS.BOF) Then
                    lsNumCertif = oSegSep.ActualizaEstadoSeguroSepelio(oRS!cNumCertificado, oRS!dFecAfiliacion, oRS!nMovNroReg, oRS!nSegEstado)
                End If
                'Call oDCOMCaptaMovimiento.AgregaSegSepelioAfiliacionHis(lsNumCertif, "", gdFecSis, lsMovNro, nMovNroExt, "", gsCodAge, 4)
                Call oDCOMCaptaMovimiento.SepelioActualizaEstadoHis(lstExtorno.SelectedItem.ListSubItems(3), 503) 'APRI20171025 ERS028-2017  4 -> 503
            End If
            If lsOpeCod = "290007" Then
                lsNumCertif = oSegSep.ActualizaEstadoSeguroSepelio("", "", lstExtorno.SelectedItem.ListSubItems(3), 503) 'APRI20171025 ERS028-2017  4 -> 503
                'Call oDCOMCaptaMovimiento.AgregaSegSepelioAfiliacionHis(lsNumCertif, "", gdFecSis, lsMovNro, nMovNroExt, "", gsCodAge)
                Call oDCOMCaptaMovimiento.SepelioActualizaEstadoHis(lstExtorno.SelectedItem.ListSubItems(3), 503) 'APRI20171025 ERS028-2017  4 -> 503
            End If
            'RECO FIN ***********************************************************
        End If
        
        'clsCapMov.OtrasOperacionesExtorno lsMovNro, lstExtorno.SelectedItem.ListSubItems(3), lsOpeCod, Me.txtGlosaExtorno.Text
        'EJRS EXTORNA REGISTRO DE PENDIENTES DE DEVOLUCION CLIENTES CON CONVENIO
        'RIRO 20150602 ADD "lstExtorno.SelectedItem.ListSubItems(4)"
      
        If lstExtorno.SelectedItem.ListSubItems(4) = "300503" Then
            Dim oCred As COMDCredito.DCOMCredito
            Set oCred = New COMDCredito.DCOMCredito
            oCred.ExtornaRegDevConvOpe lstExtorno.SelectedItem.ListSubItems(3)
            Set oCred = Nothing
        End If
        lbReimp = True
        Dim lsBoleta As String
        Dim nFicSal As Integer
        
        'JUEZ 20130417 *******************************************************
        
        If Left(lstExtorno.SelectedItem.ListSubItems(4), 4) = "3009" Then
            Dim oCredBol As COMNCredito.NCOMCredDoc
            Set oCredBol = New COMNCredito.NCOMCredDoc
            'CTI7 OPEv2*************************************************
            Dim sSubTituloBoleta As String
            If CStr(lstExtorno.SelectedItem.ListSubItems(4)) = gComiDiversasAhoCredComVoucher Then
                sSubTituloBoleta = "Pago comision-Voucher"
            Else
                sSubTituloBoleta = "Pago comision"
            End If
            '************************************************************
            lsBoleta = oCredBol.ImprimeBoletaComision("EXTORNO COMISION", Left(sSubTituloBoleta, 36), "", Str(Me.txtMonto.value), "", "", "________" & gMonedaNacional, False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU)
            Set oCredBol = Nothing
       
        ElseIf InStr(1, "300532,300533", lstExtorno.SelectedItem.ListSubItems(4)) > 0 Or _
               lstExtorno.SelectedItem.ListSubItems(4) = gOtrOpePagoRecaudoCajeroCorresponsal Then

            lsBoleta = sBoleta

        Else
            'INICIO ORCR20140714
            Dim lsCabecera As String
            lsCabecera = "EXTORNO"
        
            If lstExtorno.SelectedItem.ListSubItems(4) = "300528" Then
                lsCabecera = "EXTORNO DEV. SOBR. CHEQUE"
            End If
        
            Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
            Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
            
            'lsBoleta = oBol.ImprimeBoleta("EXTORNO", Me.lblPersona.Caption, lsOpeCod, Str(Me.txtMonto.value), "", "00000000" & Me.lstExtorno.SelectedItem.ListSubItems(5), "", 0, "0", "", 0, 0, False, False, , , , False, , , , gdFecSis, gsNomAge, gsCodUser, sLpt)
            lsBoleta = oBol.ImprimeBoleta(lsCabecera, Me.lblPersona.Caption, lsOpeCod, Str(Me.txtMonto.value), "", "________" & Me.lstExtorno.SelectedItem.ListSubItems(5), "", 0, "0", "", 0, 0, False, False, , , , False, , , , gdFecSis, gsNomAge, gsCodUser, sLpt)
            Set oBol = Nothing
            'FIN ORCR20140714
        End If
        'END JUEZ ************************************************************

        While lbReimp
            If lsBoleta <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea; lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Print #nFicSal, ""
                Close #nFicSal
                If MsgBox("Desea ReImprimir ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then lbReimp = False
            End If
        Wend
        
        If lsOpeCodExt = gOpeTransferenciaCargo Or lsOpeCodExt = gOpeTransferenciaDeposito Then
            While lbReimp
                If psBoletaExtorno <> "" Then
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea; psBoletaExtorno & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                        Print #nFicSal, ""
                    Close #nFicSal
                    If MsgBox("Desea ReImprimir ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then lbReimp = False
                End If
            Wend
        End If
        Me.txtGlosa.Text = ""
        Me.txtPers.Text = ""
        Me.txtGlosaExtorno.Text = ""
        Me.lblPersona.Caption = ""
        Me.txtMonto.Text = "0"
        Me.txtNroDoc.Text = ""
        lsOpeCodExt = 0 'JUEZ 20130906
        
        CargarLista txtUser
    End If
    Set clsCapMov = Nothing
    Set clsCont = Nothing
    Set oCred = Nothing
    Exit Sub
Error:
         MsgBox err.Description, vbInformation, "Aviso"
End Sub
Private Sub cmdAplicar_Click()
    
    If Me.txtUser = "" Then
        MsgBox "Ingrese el usuario", vbInformation, "Aviso"
        Exit Sub
    End If
    If Me.txtPers.Text = "" Then
        MsgBox "Debe elegir una operación ha extornar.", vbInformation, "Aviso"
        Me.lstExtorno.SetFocus
        Exit Sub
    End If
'    If Trim(Me.txtGlosaExtorno.Text) = "" Then
'        MsgBox "Debe ingresar un comentario por el extorno de la operación.", vbInformation, "Aviso"
'        Me.txtGlosaExtorno.SetFocus
'        Exit Sub
'    End If
frmMotExtorno.Visible = True
Me.txtDetExtorno.Text = ""
fraDetalle.Enabled = False
txtUser.Enabled = False
cmdProcesar.Enabled = False
cmdAplicar.Enabled = False
cmbMotivos.SetFocus
'******************************
End Sub
Private Sub cmdProcesar_Click()
If txtUser = "" Then
    MsgBox "Por favor seleccione el usuario", vbInformation, "aviso"
    Exit Sub
End If
Limpiar
CargarLista txtUser
End Sub

Private Sub cmdsalir_Click()
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
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    Me.lstExtorno.ColumnHeaders.Add , , "Nombre", 3000
    Me.lstExtorno.ColumnHeaders.Add , , "Monto", 1500, lvwColumnRight
    Me.lstExtorno.ColumnHeaders.Add , , "Fecha ", 2000, lvwColumnCenter
    Me.lstExtorno.ColumnHeaders.Add , , "Número", 2000, lvwColumnRight
    Me.lstExtorno.ColumnHeaders.Add , , "Código", 2000, lvwColumnRight
    Me.lstExtorno.ColumnHeaders.Add , , "Moneda", 0, lvwColumnRight
    lstExtorno.View = lvwReport
    Caption = lsCaption
    
    Dim oGen As COMDConstSistema.DCOMGeneral
     
    Set oGen = New COMDConstSistema.DCOMGeneral
    txtUser.psRaiz = "USUARIOS "
    txtUser.rs = oGen.GetUserAreaAgenciaResumenIngEgre("026", gsCodAge)
    Set oGen = Nothing
    Call CargaControles
End Sub
Sub Limpiar()
    lstExtorno.ListItems.Clear
    Me.txtGlosa = ""
    Me.txtGlosaExtorno = ""
    Me.txtMonto = 0#
    Me.txtPers = ""
    Me.lblDoc = ""
    Me.lblPers = ""
    Me.lblPersona = ""
    lsOpeCodExt = 0 'JUEZ 20130906
End Sub

Private Sub lstExtorno_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = clsCapMov.GetOtrasOperacionesDet(Item.ListSubItems(3))
    
    If Not (rs.EOF And rs.BOF) Then
        Me.txtPers.Text = rs!cperscod
        Me.lblPersona.Caption = rs!cPersNombre
        Me.txtGlosa.Text = rs!cMovDesc
        Me.txtNroDoc.Text = rs!cNroDoc
        Me.txtMonto.Text = Format$(rs!nMovImporte, "#,##0.00")
        nPeriodo = rs!nPeriodo 'RIRO20150611 ERS162-2014
        If Item.ListSubItems(5) = Moneda.gMonedaNacional Then
            Me.txtMonto.psSoles True
            Me.lblSimbolo.Caption = "S/."
        Else
            Me.txtMonto.psSoles False
            Me.lblSimbolo.Caption = "U$."
        End If
        lsOpeCodExt = rs!cOpecod 'JUEZ 20130906
    End If
    Set clsCapMov = Nothing
    
End Sub

Private Sub lstExtorno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAplicar.SetFocus
    End If
End Sub

Public Sub Ini(psOpeCod As CaptacOperacion, psCaption As String)
    lsOpeCod = psOpeCod
    lsCaption = psCaption
    lsOpeCodExt = 0 'JUEZ 20130906
    Me.Show 1
End Sub

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub

Private Sub txtGlosaExtorno_GotFocus()
    txtGlosaExtorno.SelStart = 0
    txtGlosaExtorno.SelLength = Len(txtGlosaExtorno.Text)
End Sub

Private Sub txtGlosaExtorno_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtUser_EmiteDatos()
Me.lblNomUser = Me.txtUser.psDescripcion
If lblNomUser <> "" Then
    Me.cmdProcesar.SetFocus
End If
End Sub

