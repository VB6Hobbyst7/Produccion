VERSION 5.00
Begin VB.Form frmCajaChicaHabilitacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3735
   ClientLeft      =   1530
   ClientTop       =   1980
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaChicaHabilitacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMontoAsignado 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   2760
      Width           =   1395
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "&Aprobar"
      Height          =   390
      Left            =   6000
      TabIndex        =   23
      Top             =   3240
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7320
      TabIndex        =   3
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
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
      Height          =   705
      Left            =   105
      TabIndex        =   18
      Top             =   150
      Width           =   8385
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   1200
         TabIndex        =   0
         Top             =   225
         Width           =   1095
         _ExtentX        =   1535
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7260
         TabIndex        =   22
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   210
         Left            =   6915
         TabIndex        =   21
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2310
         TabIndex        =   20
         Top             =   225
         Width           =   4500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia : "
         Height          =   210
         Left            =   90
         TabIndex        =   19
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.Frame fraEncargado 
      Caption         =   "Encargado"
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
      Height          =   1770
      Left            =   90
      TabIndex        =   4
      Top             =   885
      Width           =   8385
      Begin Sicmact.TxtBuscar TxtBuscarPersCod 
         Height          =   360
         Left            =   975
         TabIndex        =   1
         Top             =   195
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   423
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   7
         TipoBusPers     =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   17
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   1005
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   6495
         TabIndex        =   15
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   645
         Width           =   690
      End
      Begin VB.Label txtNomPer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   975
         TabIndex        =   13
         Top             =   585
         Width           =   5445
      End
      Begin VB.Label txtLEPer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6960
         TabIndex        =   12
         Top             =   570
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   11
         Top             =   1410
         Width           =   435
      End
      Begin VB.Label lblDescArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1395
         TabIndex        =   10
         Top             =   1365
         Width           =   2640
      End
      Begin VB.Label txtDirPer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   975
         TabIndex        =   9
         Top             =   960
         Width           =   7305
      End
      Begin VB.Label lblCodArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   975
         TabIndex        =   8
         Top             =   1365
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   4170
         TabIndex        =   7
         Top             =   1395
         Width           =   645
      End
      Begin VB.Label lblDescAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5385
         TabIndex        =   6
         Top             =   1365
         Width           =   2895
      End
      Begin VB.Label lblCodAge 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4920
         TabIndex        =   5
         Top             =   1365
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Guardar"
      Height          =   390
      Left            =   4680
      TabIndex        =   2
      Top             =   3240
      Width           =   1275
   End
   Begin Sicmact.Usuario Usu 
      Left            =   1080
      Top             =   2880
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Monto Asignado :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5520
      TabIndex        =   25
      Top             =   2760
      Width           =   1440
   End
End
Attribute VB_Name = "frmCajaChicaHabilitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCajaCH As nCajaChica
Dim oArendir As NARendir
Dim lbHabilitacion As Boolean

Dim fnMovNroCajaChica As Long '**Agregado por ELRO el 20120529, según OYP-RFC047-2012
Dim fsCtaFondofijo As String '**Agregado por ELRO el 20120529, según OYP-RFC047-2012
Dim fsCtaFondoFijoF As String '**Agregado por ELRO el 20120529, según OYP-RFC047-2012
Dim fsCtaDesembolso As String '**Agregado por ELRO el 20120529, según OYP-RFC047-2012
Dim fsopecod As String '**Agregado por ELRO el 20120529, según OYP-RFC047-2012

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(ByVal pbHabilitacion As Boolean)
lbHabilitacion = pbHabilitacion
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
Dim oCon As NContFunciones
Dim lsMovNro As String
Dim lsTexto As String
Dim oContImp As NContImprimir
'***Agregado por ELRO el 20120601, según OYP-RFC047-2012
Dim lnPorTopEgr As Currency
Dim lsMsg As String
'****Fin Agregado por ELRO******************************

If Valida = False Then Exit Sub

'***Modificado por ELRO el 20120601, según OYP-RFC047-2012
If lbHabilitacion And fnMovNroCajaChica = 0 Then
    lsMsg = "¿Desea Aperturar la Caja Chica del Area/Agencia Seleccionada?"
ElseIf lbHabilitacion And fnMovNroCajaChica > 0 Then
    lsMsg = "¿Desea modificar la Apertura de al Caja Chica del Area/Agencia Seleccionada?"
ElseIf lbHabilitacion = False Then
    lsMsg = "¿Desea realizar el Mantenimiento Caja Chica del Area/Agencia Seleccionada?"
End If
'If MsgBox("Desea realizar la Habilitación de Caja Chica de Area Seleccionada??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'****Fin Modificado por ELRO******************************
If MsgBox(lsMsg, vbYesNo + vbQuestion, "Aviso") = vbYes Then
    
    Set oCon = New NContFunciones
    Set oContImp = New NContImprimir
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    '***Modificado por ELRO el 20120529, según OYP-RFC047-2012
    
    lnPorTopEgr = oCajaCH.devolverTopeEgreso
    
    'If oCajaCH.GrabaHabNuevaCH(gsFormatoFecha, lsMovNro, gsOpeCod, txtMovdesc, Mid(txtBuscarAreaCH, 1, 3), _
    '    Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), CCur(txtMontoAsignado), TxtBuscarPersCod, _
    '    CCur(txtMontoAsignado), CCur(txtMontoTope)) = 0 Then
    If fnMovNroCajaChica = 0 Then
        If oCajaCH.GrabaHabNuevaCH(gsFormatoFecha, lsMovNro, _
                                   gsOpeCod, "Caja Chica", _
                                   Mid(txtBuscarAreaCH, 1, 3), _
                                   Mid(txtBuscarAreaCH, 4, 2), _
                                   Val(lblNroProcCH), _
                                   CCur(txtMontoAsignado), _
                                   TxtBuscarPersCod, _
                                   CCur(txtMontoAsignado), _
                                   lnPorTopEgr) > 0 Then
        
            'lsTexto = oContImp.ImprimeReciboIngresoEgreso(lsMovNro, gdFecSis, txtMovDesc, gsNomCmac, gsOpeCod, _
            '            TxtBuscarPersCod, CCur(txtMontoAsignado), gnColPage, , , , True, txtBuscarAreaCH & "-" & Val(lblNroProcCH), _
            '           lblCajaChicaDesc, , , , True)
            'lsTexto = oContImp.ImprimeReciboIngresoEgreso(lsMovNro, gdFecSis, "Caja Chica", gsNomCmac, gsOpeCod, _
            '            TxtBuscarPersCod, CCur(txtMontoAsignado), gnColPage, , , , True, txtBuscarAreaCH & "-" & Val(lblNroProcCH), _
            '            lblCajaChicaDesc, , , , True)
        '***Fin Modificado por ELRO*******************************
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Aperturo Caja Chica del  Area : " & lblCajaChicaDesc & " |Encargado : " & txtNomPer _
            & "|Monto : " & txtMontoAsignado.Text
            Set objPista = Nothing
            '*******
            
            'EnviaPrevio lsTexto, Me.Caption, gnLinPage '***Comentado por ELRO el 20120605, según OYP-RFC047-2012
            If MsgBox("Desea habilitar otra Area como Caja chica??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                txtBuscarAreaCH.rs = oArendir.EmiteCajasChicasNuevas
                TxtBuscarPersCod.Text = ""
                lblCodArea = ""
                lblCodAge = ""
                lblDescAge = ""
                lblDescArea = ""
                txtDirPer = ""
                txtLEPer = ""
                txtNomPer = ""
                lblCajaChicaDesc = ""
                txtMontoAsignado = "0.00"
                '***Comentado por ELRO el 20120605, según OYP-RFC047-2012
                'txtMontoTope = "0.00"
                'txtSaldo = "0.00"
                'txtSaldoAnt = "0.00"
                'txtMontoDesemb = "0.00"
                'txtMovDesc = ""
                '***Fin Comentado por ELRO*******************************
            Else
                Set oCon = Nothing
                Set oContImp = Nothing
                Unload Me
            End If
        End If
    Else
        Dim nMovNroEditar As Long
        nMovNroEditar = oCajaCH.GrabaEdicionAperturaCajaChica(fnMovNroCajaChica, _
                                                             CCur(txtMontoAsignado), _
                                                             TxtBuscarPersCod, _
                                                             Mid(txtBuscarAreaCH, 1, 3), _
                                                             Mid(txtBuscarAreaCH, 4, 2), _
                                                             lnPorTopEgr, _
                                                             Val(lblNroProcCH))
        If nMovNroEditar = fnMovNroCajaChica Then
            MsgBox "Se modificó correctamente la Caja Chica del Área/Agencia seleccionada."
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", Me.Caption & " Se Modifico Caja Chica del  Area : " & lblCajaChicaDesc & " |Encargado : " & txtNomPer _
            & "|Monto : " & txtMontoAsignado.Text
            Set objPista = Nothing
            '*******
        End If
    End If
End If
End Sub
Function Valida() As Boolean
Valida = True
If Len(Trim(txtBuscarAreaCH)) = 0 Then
    MsgBox "N° de caja chica no ingresado", vbInformation, "Aviso"
    If txtBuscarAreaCH.Enabled Then txtBuscarAreaCH.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(Me.TxtBuscarPersCod)) = 0 Then
    MsgBox "Persona encargada no ingresado", vbInformation, "Aviso"
    If TxtBuscarPersCod.Enabled And Me.fraEncargado.Enabled Then TxtBuscarPersCod.SetFocus
    Valida = False
    Exit Function
End If
If Val(txtMontoAsignado) = 0 Then
    MsgBox "Monto Asignado no ingresado", vbInformation, "Aviso"
    txtMontoAsignado.SetFocus
    Valida = False
    Exit Function
End If
'***Comentado por ELRO el 20120601, según OYP-RFC047-2012
'If Val(txtMontoTope) = 0 Then
'    MsgBox "Monto Tope no ingresado", vbInformation, "Aviso"
'    txtMontoTope.SetFocus
'    Valida = False
'    Exit Function
'End If

'If CCur(txtMontoTope) > CCur(txtMontoAsignado) Then
'    MsgBox "Monto tope debe ser mayor que monto Asignado", vbInformation, "Aviso"
'    txtMontoTope.SetFocus
'    Valida = False
'    Exit Function
'End If
'***Fin Comentado por ELRO*******************************

If lbHabilitacion Then
    '***Comentado por ELRO el 20120601, según OYP-RFC047-2012
    'If Len(Trim(txtMovDesc)) = 0 Then
    '    MsgBox "Descripción de operación no ingresada", vbInformation, "Aviso"
    '    txtMovDesc.SetFocus
    '    Valida = False
    '    Exit Function
    'End If
    '***Fin Comentado por ELRO*******************************
End If


End Function

Private Sub cmdAprobar_Click()
Dim oNCajaChica As nCajaChica
Set oNCajaChica = New nCajaChica
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim oDOperacion As New DOperacion
Dim rs As ADODB.Recordset
Dim rsAut As ADODB.Recordset

Dim lsNroDoc As String
Dim lsNroVoucher As String
Dim lsFechaDoc As String
Dim lsDocumento As String
Dim lsGlosa As String
Dim lsMovNro As String
Dim lsCadBol As String
Dim lnImporte As Currency
Dim lsCtaContDebeITF As String
Dim lsCtaContHaberITF As String
Dim lsMontoITF As Double
Dim lnMovNroApro As Long

If Valida = False Then Exit Sub

If fsopecod = CStr(gCHApropacionApeMN) Or fsopecod = CStr(gCHApropacionApeME) Then
    If CInt(lblNroProcCH) > 1 Then
        fsCtaFondofijo = oDOperacion.EmiteOpeCta(fsopecod, "D", "1")
    Else
        fsCtaFondofijo = oDOperacion.EmiteOpeCta(fsopecod, "D")
    End If
Else
    fsCtaFondofijo = oDOperacion.EmiteOpeCta(fsopecod, "D")
End If

lsNroDoc = ""
lsNroVoucher = ""
lsFechaDoc = ""
lsDocumento = ""
lsGlosa = ""
lnImporte = CCur(txtMontoAsignado)
Set rsAut = oNCajaChica.devolverDatosCajaChicaSinAprobar_2(fnMovNroCajaChica, Mid(Me.txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2))

If MsgBox("¿Desea realizar la Aprobación de la Apertura de Caja Chica?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    'edpyme - sgte reglon para fondos fijos por agencia
    Dim lsSubCta As String
    lsSubCta = oNContFunciones.GetFiltroObjetos(ObjCMACAgenciaArea, fsCtaFondofijo, Trim(txtBuscarAreaCH), False)
 
    If lsSubCta <> "" Then
       If Mid(txtBuscarAreaCH, 1, 3) = "042" Then 'Recuperaciones
              fsCtaFondoFijoF = fsCtaFondofijo & lsSubCta
              
       ElseIf Mid(txtBuscarAreaCH, 1, 3) = "043" Then  'Secretaria
              fsCtaFondoFijoF = fsCtaFondofijo & lsSubCta
              
       ElseIf Mid(txtBuscarAreaCH, 1, 3) = "023" Then  'Logistica
              fsCtaFondoFijoF = fsCtaFondofijo & lsSubCta
       Else
              fsCtaFondoFijoF = fsCtaFondofijo & lsSubCta
       End If
       
    Else
        MsgBox "Sub Cuenta Contable no definida.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lsCtaContDebeITF = oDOperacion.EmiteOpeCta(fsopecod, "D", 2)
    lsCtaContHaberITF = oDOperacion.EmiteOpeCta(fsopecod, "H", 2)
    
    lsMontoITF = fgTruncar(lnImporte * gnImpITF, 2)

    Call oNCajaChica.GrabaDesembolsoCH(lsMovNro, fnMovNroCajaChica, _
                                        TxtBuscarPersCod, _
                                        gsFormatoFecha, rs, fsopecod, _
                                        "Aprobación Apertura Caja Chica", _
                                        lnImporte, _
                                        Mid(Me.txtBuscarAreaCH, 1, 3), _
                                        Mid(Me.txtBuscarAreaCH, 4, 2), _
                                        Val(lblNroProcCH), "", _
                                        lsNroDoc, lsFechaDoc, _
                                        lsNroVoucher, rsAut, _
                                        "", 0, fsCtaFondoFijoF, _
                                        fsCtaDesembolso, 0#, "", _
                                        gbBitCentral, , lsCtaContDebeITF, _
                                        lsCtaContHaberITF, lsMontoITF, lnMovNroApro)
    If lnMovNroApro > 0 Then
        MsgBox "Se realizó correctamente la Aprobación de la Apertura de Caja Chica.", vbInformation, "Aviso"
        TxtBuscarPersCod.Text = ""
        TxtBuscarPersCod.psCodigoPersona = ""
        txtBuscarAreaCH.Text = ""
        txtBuscarAreaCH.psCodigoPersona = ""
        lblNroProcCH = ""
        lblCodArea = ""
        lblDescArea = ""
        lblCodAge = ""
        lblDescAge = ""
        txtDirPer = ""
        txtLEPer = ""
        txtNomPer = ""
        lblCajaChicaDesc = ""
        txtMontoAsignado = "0.00"
        '***Comentado por ELRO el 20120605, según OYP-RFC047-2012
        'txtMontoTope = "0.00"
        'txtSaldo = "0.00"
        'txtSaldoAnt = "0.00"
        'txtMontoDesemb = "0.00"
        '***Fin Comentado por ELRO*******************************
        
        'ARLO 20161221
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & "Aprobo Caja Chica"
        Set objPista = Nothing
        '*******************
    End If
End If
Set oNContFunciones = Nothing
Set oDOperacion = Nothing
Set rsAut = Nothing

End Sub

'***Comentado por ELRO el 20120605, según OYP-RFC047-2012
'Private Sub cmdCancelar_Click()
''***Comentado por ELRO el 20120605, según OYP-RFC047-2012
''cmdEditar.Caption = "&Editar"
''***Fin Comentado por ELRO*******************************
'txtBuscarAreaCH.Enabled = True
'cmdCancelar.Enabled = False
'txtMontoAsignado.Enabled = False
'txtMontoTope.Enabled = False
'fraEncargado.Enabled = False
'CargaDatosCajaChica Mid(Me.txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2), Val(Me.lblNroProcCH)
'End Sub
'***Fin Comentado por ELRO*******************************
'***Comentado por ELRO el 20120605, según OYP-RFC047-2012
'Private Sub cmdEditar_Click()
'Dim rs As ADODB.Recordset
'If cmdEditar.Caption = "&Editar" Then
'    cmdEditar.Caption = "&Guardar"
'    cmdCancelar.Enabled = True
'    txtBuscarAreaCH.Enabled = False
'    txtMontoAsignado.Enabled = True
'    txtMontoTope.Enabled = True
'    fraEncargado.Enabled = True
'    txtMontoAsignado.SetFocus
'Else
'    If Valida = False Then Exit Sub
'
'    If MsgBox("Desea Actualizar datos de la caja chica seleccionada??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'        Set rs = oCajaCH.GetCHRendSinAutorizacion(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
'        If Not rs.EOF Then
'
'        End If
'        If oCajaCH.ActualizaCH(gsFormatoFecha, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), _
'                    Trim(TxtBuscarPersCod), CCur(txtMontoAsignado), CCur(txtMontoTope)) = 0 Then
'
'            cmdCancelar.value = True
'        End If
'    End If
'End If
'End Sub
'***Fin Comentado por ELRO*******************************
'***Comentado por ELRO el 20120605, según OYP-RFC047-2012
'Private Sub cmdImprimir_Click()
'Dim oConImp As NContImprimir
'Dim lsTexto As String
'Dim oCon As DConecta
'Dim sSql As String
'Dim prs  As ADODB.Recordset
'Set oCon = New DConecta
'Set oConImp = New NContImprimir
'oCon.AbreConexion
'
'    sSql = "Select m.nMovNro, m.cMovNro, m.cMovDesc, c.nMovMonto from mov m join MovCont C ON c.nMovNro = m.nMovNro join movcajachica mc ON mc.nMovNro = m.nMovNro where cAreaCod = '" & Left(txtBuscarAreaCH, 3) & "' and cAgeCod = '" & Mid(txtBuscarAreaCH, 4, 2) & "' and mc.nMovNro = (SELECT Max(nMovNro) FROM MovCajachica mc1 where mc1.cAreaCod = mc.cAreaCod and mc1.cAgeCod = mc.cAgeCod and nProcNro <= " & Me.lblNroProcCH & " and cProcTpo = 1  ) "
'    Set prs = oCon.CargaRecordSet(sSql)
'
'    If prs.EOF Then
'        MsgBox "No existen datos para imprimir", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    lsTexto = oConImp.ImprimeReciboIngresoEgreso(prs!cMovNro, GetFechaMov(prs!cMovNro, True), prs!cMovDesc, gsNomCmac, gsOpeCod, _
'              Me.TxtBuscarPersCod, CCur(prs!nMovMonto), gnColPage, , , False, True, txtBuscarAreaCH & "-" & Val(lblNroProcCH), _
'              lblCajaChicaDesc, , , , True)
'    RSClose prs
'    EnviaPrevio lsTexto, Me.Caption, gnLinPage
'oCon.CierraConexion
'Set oCon = Nothing
'End Sub
'***Fin Comentado por ELRO*******************************

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set oArendir = New NARendir
Set oCajaCH = New nCajaChica
Dim oDActualizaDatosArea As DActualizaDatosArea 'Agregado por ELRO el 20120529, según OYP-RFC047-2012
Set oDActualizaDatosArea = New DActualizaDatosArea 'Agregado por ELRO el 20120529, según OYP-RFC047-2012
Dim cCodCarJefCon As String 'Agregado por ELRO el 20120529, según OYP-RFC047-2012

CentraForm Me
Me.Caption = gsOpeDesc

cCodCarJefCon = oArendir.devolverCodigoJefeContabilidad

If gsCodCargo <> cCodCarJefCon Then

    If lbHabilitacion Then
        '***Comentado por ELRO el 20120605, según OYP-RFC047-2012
        'cmdEditar.Visible = False
        'cmdCancelar.Visible = False
        '***Fin Comentado por ELRO*******************************
        fraEncargado.Enabled = True
        cmdAceptar.Visible = True
        '***Comentado por ELRO el 20120529, según OYP-RFC047-2012
        'cmdImprimir.Visible = False
        'txtMovDesc.Visible = True
        txtBuscarAreaCH.lbUltimaInstancia = False
        '***Modificado por ELRO el 20120529, según OYP-RFC047-2012
        'txtBuscarAreaCH.psRaiz = "NUEVAS CAJAS CHICAS"
        'txtBuscarAreaCH.rs = oArendir.EmiteCajasChicasNuevas
        txtBuscarAreaCH.psRaiz = "Areas - Agencias"
        txtBuscarAreaCH.rs = oDActualizaDatosArea.GetAgenciasAreas
        Set oDActualizaDatosArea = Nothing
        '***Fin Modificado por ELRO*******************************
    Else
        '***Comentado por ELRO el 20120605, según OYP-RFC047-2012
        'cmdEditar.Visible = True
        'cmdCancelar.Visible = True
        '***Fin Comentado por ELRO*******************************
        '***Modificado por ELRO el 201200605, según OYP-RFC047-2012
        'fraEncargado.Enabled = False
        'cmdAceptar.Visible = False
        fraEncargado.Enabled = True
        cmdAceptar.Visible = True
        '***Fin Modificado por ELRO*******************************
        '***Comentado por ELRO el 20120529, según OYP-RFC047-2012
        'txtMovDesc.Visible = False
        '***Fin Comentado por ELRO*******************************
        txtBuscarAreaCH.psRaiz = "CAJAS CHICAS"
        txtBuscarAreaCH.rs = oArendir.EmiteCajasChicas
        '***Modificado por ELRO el 201200605, según OYP-RFC047-2012
        'txtMontoAsignado.Enabled = False
        txtMontoAsignado.Enabled = True
        '***Fin Modificado por ELRO*******************************
        '***Comentado por ELRO el 20120604, según OYP-RFC047-2012
        'txtMontoTope.Enabled = False
        '***Fin Comentado por ELRO*******************************
    End If
    '***Agregado por ELRO el 20120529, según OYP-RFC047-2012
    cmdAprobar.Enabled = False
    '***Fin Agregado por ELRO*******************************
Else
    Dim oDOperacion As DOperacion
    Set oDOperacion = New DOperacion
    txtBuscarAreaCH.psRaiz = "Areas - Agencias"
    txtBuscarAreaCH.rs = oDActualizaDatosArea.GetAgenciasAreas
    fraEncargado.Enabled = False
    cmdAceptar.Enabled = False
    txtMontoAsignado.Locked = True
    Set oDActualizaDatosArea = Nothing
    fsopecod = IIf(Mid(gsOpeCod, 3, 1) = "1", CStr(gCHApropacionApeMN), CStr(gCHApropacionApeME))
    fsCtaFondofijo = oDOperacion.EmiteOpeCta(fsopecod, "D")
    fsCtaDesembolso = oDOperacion.EmiteOpeCta(fsopecod, "H")
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oArendir = Nothing
Set oCajaCH = Nothing
fnMovNroCajaChica = 0
End Sub



Private Sub txtBuscarAreaCH_EmiteDatos()
Dim lnAreaAge As Integer '***Agregado por ELRO el 20120529, según OYP-RFC047-2012
Dim lnAprobacion As Integer '***Agregado por ELRO el 20120529, según OYP-RFC047-2012

If txtBuscarAreaCH = "" Then Exit Sub

'***Agregado por ELRO el 20120529, según OYP-RFC047-2012
If lbHabilitacion Then
    lnAreaAge = oCajaCH.verificarAreasAgenciasCajaChica(Mid(txtBuscarAreaCH, 1, 3), IIf(Mid(txtBuscarAreaCH, 4, 2) = "", "01", Mid(txtBuscarAreaCH, 4, 2)))
    lnAprobacion = oCajaCH.verificarAprobacionCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
    
    If lnAreaAge = 0 Then
        MsgBox "El Área/Agencia selecionada no debe Aperturarse Caja Chica." & Chr(13) & "Consultar Reglamento de Caja Chica en la Intranet."
        txtBuscarAreaCH = ""
        Exit Sub
        txtBuscarAreaCH.SetFocus
    End If
    
    If lnAprobacion > 0 Then
        MsgBox "El Área/Agencia seleccionada ya fue Aperturada y Aprobada su Caja Chica."
        txtBuscarAreaCH = ""
        Exit Sub
        txtBuscarAreaCH.SetFocus
    End If
End If
'***Fin Agregado por ELRO*******************************

lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
lblNroProcCH = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
If lbHabilitacion Then
    '***Modificado por ELRO el 20120531, según OYP-RFC047-2012
    'If Val(lblNroProcCH) = 0 Then
    '    lblNroProcCH = Val(lblNroProcCH) + 1
    'End If
    If Val(lblNroProcCH) = 1 Then
        Call CargaDatosCajaChicaSinAprobar(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
    ElseIf Val(lblNroProcCH) = 0 Then
        If fsopecod = CStr(gCHApropacionApeMN) Or fsopecod = CStr(gCHApropacionApeME) Then
            MsgBox "Esta Área/Agencia seleccionada aún no esta Aperurada su Caja Chica.", vbInformation, "Titulo"
            Exit Sub
        Else
            lblNroProcCH = Val(lblNroProcCH) + 1
        End If
    '***Fin Modificado por ELRO*******************************
    End If
    '***Modificado por ELRO el 20120531, según OYP-RFC047-2012
    'TxtBuscarPersCod.SetFocus
    If fraEncargado.Enabled Then
        TxtBuscarPersCod.SetFocus
    End If
    '***Fin Modificado por ELRO*******************************
Else
    '***Modificado por ELRO el 201206051, según OYP-RFC047-2012
    'CargaDatosCajaChica Mid(Me.txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2), Val(Me.lblNroProcCH)
    If Val(lblNroProcCH) > 1 Then
        Call CargaDatosCajaChicaMantenimiento(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH))
'    ElseIf Val(lblNroProcCH) = 1 And fnMovNroCajaChica = 0 Then
'        MsgBox "El Área/Agencia seleccionada tiene N° de Proceso 1, " & Chr(13) & "por tal motivo no se puede realizar el mantenimiento.", vbInformation, "Aviso"
'        txtBuscarAreaCH = ""
'        txtBuscarAreaCH.psCodigoPersona = ""
'        lblCajaChicaDesc = ""
'        lblNroProcCH = ""
    ElseIf Val(lblNroProcCH) = 1 And fnMovNroCajaChica >= 0 Then
        MsgBox "El Mantenimiento de una Caja Chica se puede realizar antes del proceso de Autorización de Reembolso.", vbInformation, "Aviso"
        txtBuscarAreaCH = ""
        txtBuscarAreaCH.psCodigoPersona = ""
        lblCajaChicaDesc = ""
        lblNroProcCH = ""
    End If
    '***Fin Modificado por ELRO*******************************
    '***Comentado por ELRO el 20120529, según OYP-RFC047-2012
    'If cmdEditar.Visible Then cmdEditar.SetFocus
    '***Fin Comentado por ELRO*******************************
End If

End Sub

Private Sub txtBuscarAreaCH_Validate(Cancel As Boolean)
If txtBuscarAreaCH = "" Then
    Cancel = True
End If
End Sub

Private Sub TxtBuscarPersCod_EmiteDatos()
If TxtBuscarPersCod.Text = "" Then Exit Sub
usu.DatosPers TxtBuscarPersCod.Text
If usu.PersCod = "" Then
    MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
Else
    If (Len(txtBuscarAreaCH.Text) = 3 And Mid(txtBuscarAreaCH.Text, 1, 3) <> Trim(usu.cAreaCodAct)) Or (Len(txtBuscarAreaCH.Text) = 5 And Right(txtBuscarAreaCH.Text, 2) <> Trim(usu.CodAgeAct)) Then
        MsgBox "Persona no pertenece al área Seleccionada", vbInformation, "Aviso"
        TxtBuscarPersCod.Text = ""
        usu.DatosPers TxtBuscarPersCod.Text
        'Exit Sub
    End If
    txtMontoAsignado.SetFocus
End If
AsignaValores
End Sub
Private Sub AsignaValores()
    TxtBuscarPersCod.Text = usu.PersCod
    lblCodArea = usu.AreaCod
    lblCodAge = usu.CodAgeAct
    lblDescAge = usu.DescAgeAct
    lblDescArea = usu.AreaNom
    txtDirPer = usu.DireccionUser
    txtLEPer = IIf(usu.NroDNIUser = "", usu.NroRucUser, usu.NroDNIUser)
    txtNomPer = PstaNombre(usu.UserNom)
End Sub
Sub CargaDatosCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs = oCajaCH.GetRSDatosCajaChica(psAreaCh, psAgeCh, pnProcCH)
If Not rs.EOF And Not rs.BOF Then
    TxtBuscarPersCod = rs!cPersCod
    usu.DatosPers rs!cPersCod
    AsignaValores
    '***Comentado por ELRO el 20120605, según OYP-RFC047-2012
    'txtSaldo = Format(rs!nSaldo, "#,#0.00")
    'txtSaldoAnt = Format(rs!nSaldoAnt, "#,#0.00")
    'txtMontoDesemb = Format(rs!nMontoDesem, "#,#0.00")
    'txtMontoTope = Format(rs!nTopeEgresos, "#,#0.00")
    '***Fin Comentado por ELRO*******************************
    txtMontoAsignado = Format(rs!nMontoAsig, "#,#0.00")
End If
rs.Close
Set rs = Nothing
End Sub

'***Agregado por ELRO el 20120531, según OYP-RFC047-2012
Private Sub CargaDatosCajaChicaSinAprobar(ByVal psAreaCh As String, _
                                          ByVal psAgeCh As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs = oCajaCH.devolverDatosCajaChicaSinAprobar(psAreaCh, psAgeCh)
If Not rs.EOF And Not rs.BOF Then
    TxtBuscarPersCod = rs!cPersCod
    usu.DatosPers rs!cPersCod
    AsignaValores
    txtMontoAsignado = Format(rs!nMontoAsig, "#,#0.00")
    fnMovNroCajaChica = rs!nMovNroCajaChica
Else
    TxtBuscarPersCod = ""
    txtNomPer = ""
    txtLEPer = ""
    txtDirPer = ""
    lblCodArea = ""
    lblDescArea = ""
    lblCodAge = ""
    lblDescAge = ""
    txtMontoAsignado = ""
    fnMovNroCajaChica = 0
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub CargaDatosCajaChicaMantenimiento(ByVal psAreaCh As String, _
                                             ByVal psAgeCh As String, _
                                             ByVal pnProcNro As Integer)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs = oCajaCH.devolverDatosCajaChicaMantenimiento(psAreaCh, psAgeCh, pnProcNro)
If Not rs.EOF And Not rs.BOF Then
    TxtBuscarPersCod = rs!cPersCod
    usu.DatosPers rs!cPersCod
    AsignaValores
    txtMontoAsignado = Format(rs!nMontoAsig, "#,#0.00")
    fnMovNroCajaChica = rs!nMovNroCajaChica
Else
    MsgBox "El Mantenimiento de una Caja Chica se puede realizar antes del proceso de Autorización de Reembolso.", vbInformation, "Aviso"
    lblCajaChicaDesc = ""
    lblNroProcCH = ""
    txtBuscarAreaCH = ""
    txtBuscarAreaCH.psCodigoPersona = ""
    TxtBuscarPersCod = ""
    TxtBuscarPersCod.psCodigoPersona = ""
    txtNomPer = ""
    txtLEPer = ""
    txtDirPer = ""
    lblCodArea = ""
    lblDescArea = ""
    lblCodAge = ""
    lblDescAge = ""
    txtMontoAsignado = ""
    fnMovNroCajaChica = 0
End If
rs.Close
Set rs = Nothing
End Sub
'***Fin Agregado por ELRO*******************************

Private Sub txtMontoAsignado_GotFocus()
fEnfoque txtMontoAsignado
End Sub
'***Comentado por ELRO el 20120605, según OYP-RFC047-2012
'Private Sub txtMontoAsignado_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosDecimales(txtMontoAsignado, KeyAscii)
'If KeyAscii = 13 Then
'    txtMontoTope.SetFocus
'End If
'End Sub
'***Fin Comentado por ELRO*******************************

Private Sub txtMontoAsignado_LostFocus()
If txtMontoAsignado = "" Then txtMontoAsignado = 0
txtMontoAsignado = Format(txtMontoAsignado, "#,#0.00")
End Sub

'***Comentado por ELRO el 20120605, según OYP-RFC047-2012
'Private Sub txtMontoTope_GotFocus()
'fEnfoque txtMontoTope
'End Sub

'Private Sub txtMontoTope_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosDecimales(txtMontoTope, KeyAscii)
'If KeyAscii = 13 Then
'    If CCur(txtMontoTope) > CCur(txtMontoAsignado) And Val(txtMontoAsignado) > 0 Then
'        MsgBox "Monto tope debe ser mayor que monto Asignado", vbInformation, "Aviso"
'        txtMontoTope.SetFocus
'        Exit Sub
'    End If
'    If lbHabilitacion Then
'        '***Comentado por ELRO el 20120601, según OYP-RFC047-2012
'        'txtMovDesc.SetFocus
'        '***Fin Comentado por ELRO*******************************
'    Else
'        '***Comentado por ELRO el 20120601, según OYP-RFC047-2012
'        'cmdEditar.SetFocus
'        '***Fin Comentado por ELRO*******************************
'    End If
'End If
'End Sub

'Private Sub txtMontoTope_LostFocus()
'If txtMontoTope = "" Then txtMontoTope = 0
'txtMontoTope = Format(txtMontoTope, "#,#0.00")
'End Sub

'Private Sub txtMontoTope_Validate(Cancel As Boolean)
'If CCur(txtMontoTope) >= CCur(txtMontoAsignado) And Val(txtMontoAsignado) > 0 Then
'    MsgBox "Monto tope debe ser menor que monto Asignado", vbInformation, "Aviso"
'    Cancel = True
'End If
'End Sub
'***Fin Comentado por ELRO*******************************

'***Comentado por ELRO el 20120601, según OYP-RFC047-2012
'Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    KeyAscii = 0
'    cmdAceptar.SetFocus
'End If
'End Sub
'***Fin Comentado por ELRO*******************************
