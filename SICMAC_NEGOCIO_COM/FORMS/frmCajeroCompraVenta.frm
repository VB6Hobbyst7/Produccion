VERSION 5.00
Begin VB.Form frmCajeroCompraVenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compra Venta Moneda Extranjera"
   ClientHeight    =   4200
   ClientLeft      =   1875
   ClientTop       =   2205
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5220
      TabIndex        =   4
      Top             =   3750
      Width           =   1155
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   4050
      TabIndex        =   3
      Top             =   3750
      Width           =   1185
   End
   Begin VB.Frame fraTipoCambio 
      Height          =   1065
      Left            =   60
      TabIndex        =   9
      Top             =   3060
      Width           =   3825
      Begin VB.TextBox TxtMontoPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1665
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin SICMACT.EditMoney txtImporte 
         Height          =   375
         Left            =   1665
         TabIndex        =   2
         Top             =   195
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         font            =   "frmCajeroCompraVenta.frx":0000
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin VB.Label lblTipoCambio 
         Caption         =   "Monto a Cambiar:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto a Pagar:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblsimbolosoles2 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3435
         TabIndex        =   12
         Top             =   690
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "$."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   11
         Top             =   285
         Width           =   180
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Datos de la Persona"
      Height          =   2745
      Left            =   60
      TabIndex        =   7
      Top             =   315
      Width           =   6315
      Begin VB.TextBox txtIdAut 
         Height          =   330
         Left            =   4470
         TabIndex        =   18
         Top             =   225
         Visible         =   0   'False
         Width           =   1380
      End
      Begin SICMACT.FlexEdit fgDocs 
         Height          =   1230
         Left            =   1515
         TabIndex        =   1
         Top             =   1335
         Width           =   4095
         _extentx        =   7223
         _extenty        =   2170
         cols0           =   4
         highlight       =   2
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "N°-Documento-N° Doc.-Tipo"
         encabezadosanchos=   "450-1500-1800-0"
         font            =   "frmCajeroCompraVenta.frx":0024
         font            =   "frmCajeroCompraVenta.frx":004C
         font            =   "frmCajeroCompraVenta.frx":0074
         font            =   "frmCajeroCompraVenta.frx":009C
         font            =   "frmCajeroCompraVenta.frx":00C4
         fontfixed       =   "frmCajeroCompraVenta.frx":00EC
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X"
         listacontroles  =   "0-0-0-0"
         encabezadosalineacion=   "C-L-L-L"
         formatosedit    =   "0-0-0-0"
         textarray0      =   "N°"
         lbformatocol    =   -1  'True
         lbpuntero       =   -1  'True
         lbordenacol     =   -1  'True
         colwidth0       =   450
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin SICMACT.TxtBuscar txtBuscaPers 
         Height          =   330
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   1815
         _extentx        =   3201
         _extenty        =   582
         appearance      =   1
         appearance      =   1
         font            =   "frmCajeroCompraVenta.frx":0112
         appearance      =   1
         tipobusqueda    =   3
         stitulo         =   ""
         tipobuspers     =   1
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Id Autorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3150
         TabIndex        =   19
         Top             =   315
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Documentos :"
         Height          =   210
         Left            =   195
         TabIndex        =   8
         Top             =   1335
         Width           =   1200
      End
      Begin VB.Label lblPersDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   6
         Top             =   930
         Width           =   5670
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   585
         Width           =   5670
      End
   End
   Begin VB.Label lblTpoCambioDia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5250
      TabIndex        =   17
      Top             =   3210
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cambio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4050
      TabIndex        =   16
      Top             =   3270
      Width           =   1080
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "COMPRA MONEDA EXTRANJERA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   1410
      TabIndex        =   15
      Top             =   75
      Width           =   4140
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   3975
      Top             =   3195
      Width           =   2430
   End
End
Attribute VB_Name = "frmCajeroCompraVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPers As comdpersona.UCOMPersona
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsAgencia As String
Dim lbSalir As Boolean
Dim lsDocumento  As String
Dim nPersoneria As COMDConstantes.PersPersoneria


'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String


Private Sub CmdGuardar_Click()
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim lsMovNro As String
Dim oGen  As COMNContabilidad.NCOMContFunciones
Set oGen = New COMNContabilidad.NCOMContFunciones
Set oCajero = New COMNCajaGeneral.NCOMCajero
If ValidaInterfaz = False Then Exit Sub


If MsgBox("Desea grabar la Operación de Compra/Venta??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    
    Dim sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
    Dim nMontoLavDinero As Double
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
    Dim nMonto As Double
    Dim bLavDinero As Boolean
    Dim oVarPublicas As New COMFunciones.FCOMVarPublicas
    Dim lsBoleta As String
    Dim nFicSal As Integer
    bLavDinero = False
    nMonto = txtImporte.value
    'Realiza la Validación para el Lavado de Dinero
    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
    If Not clsExo.EsPersonaExoneradaLavadoDinero(txtBuscaPers) Then
        Set clsExo = Nothing
        sPersLavDinero = ""
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nMontoLavDinero = clsLav.GetCapParametro(COMDConstantes.gMonOpeLavDineroME)
        Set clsLav = Nothing
        
        If nMonto >= nMontoLavDinero Then
            If nPersoneria = COMDConstantes.gPersonaNat Then
                sPersLavDinero = frmMovLavDinero.Inicia(Trim(txtBuscaPers.Text), Trim(lblPersNombre), Trim(lblPersDireccion), Trim(fgDocs.TextMatrix(1, 2)), , , nMonto, " ", gsOpeDesc)
                sPersLavDinero = gVarPublicas.gReaPersLavDinero
                sReaPersLavDinero = gVarPublicas.gReaPersLavDinero
                sBenPersLavDinero = gVarPublicas.gBenPersLavDinero
                
            Else
                sPersLavDinero = frmMovLavDinero.Inicia(Trim(txtBuscaPers.Text), Trim(lblPersNombre), Trim(lblPersDireccion), Trim(fgDocs.TextMatrix(1, 2)), True, True, nMonto, " ", gsOpeDesc)
                sPersLavDinero = gVarPublicas.gReaPersLavDinero
                sReaPersLavDinero = gVarPublicas.gReaPersLavDinero
                sBenPersLavDinero = gVarPublicas.gBenPersLavDinero
                
            End If
            
            If sPersLavDinero = "" Then Exit Sub
            bLavDinero = True
        End If
    Else
        Set clsExo = Nothing
    End If
    lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If oCajero.GrabaCompraVenta(gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, nMonto, CDbl(lblTpoCambioDia), txtBuscaPers, bLavDinero, sBenPersLavDinero) = 0 Then
        
        Dim oImp As COMNContabilidad.NCOMContImprimir
        Dim lsTexto As String
        Dim lbReimp As Boolean
        Set oImp = New COMNContabilidad.NCOMContImprimir
        
        
        lsBoleta = oImp.ImprimeBoletaCompraVenta(lblTitulo, "", lblPersNombre, lblPersDireccion, lsDocumento, _
                    CCur(lblTpoCambioDia), gsOpeCod, CCur(txtImporte), CCur(TxtMontoPagar), gsNomAge, lsMovNro, sLpt, gsCodCMAC, gsNomCmac)
        lbReimp = True
        Do While lbReimp
             If Trim(lsBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoleta
                    Print #nFicSal, ""
                Close #nFicSal
             End If
           
        
            If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbReimp = False
            End If
        Loop
             
        Set oImp = Nothing
        txtBuscaPers = ""
        lblPersDireccion = ""
        lblPersNombre = ""
        fgDocs.Clear
        fgDocs.FormaCabecera
        fgDocs.Rows = 2
        txtImporte = 0
        TxtMontoPagar = "0.00"
        txtBuscaPers.SetFocus
    End If
    
    gVarPublicas.LimpiaVarLavDinero
    
    Set oGen = Nothing
    Set oCajero = Nothing
    
End If

End Sub

Function ValidaInterfaz() As Boolean
ValidaInterfaz = True
If txtBuscaPers = "" Then
    MsgBox "Persona no Ingresada", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtBuscaPers.SetFocus
    Exit Function
End If
If Val(txtImporte) = 0 Then
    MsgBox "Importe de Operación no Ingresado", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtImporte.SetFocus
    Exit Function
End If
If Val(TxtMontoPagar) = 0 Then
    MsgBox "Monto a Pagar no válido para Operación", vbInformation, "Aviso"
    ValidaInterfaz = False
    Exit Function
End If



End Function
Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim oOpe As COMNCajaGeneral.NCOMCajaCtaIF
Dim oTipCambio As COMDConstSistema.NCOMTipoCambio

Set oTipCambio = New COMDConstSistema.NCOMTipoCambio
gnTipCambioC = oTipCambio.EmiteTipoCambio(gdFecSis, TCCompra)
gnTipCambioV = oTipCambio.EmiteTipoCambio(gdFecSis, TCVenta)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Set oTipCambio = Nothing

Set oOpe = New COMNCajaGeneral.NCOMCajaCtaIF

Me.Caption = gsOpeDesc
txtImporte.psSoles False
lbSalir = False
lblTpoCambioDia = Format(IIf(gsOpeCod = gOpeCajeroMECompra, gnTipCambioC, gnTipCambioV), "#,#0.0000")
If Val(lblTpoCambioDia) = 0 Then
    MsgBox "Tipo de Cambio no ha sido Ingresado. Por favor Ingrese Tipo Cambio del Día", vbInformation, "Aviso"
    lbSalir = True
    Exit Sub
End If
Select Case gsOpeCod
    Case COMDConstSistema.gOpeCajeroMECompra
        lblTitulo = "COMPRA MONEDA EXTRANJERA NORMAL"
        Me.lblMonto = "Monto a Pagar"
    Case COMDConstSistema.gOpeCajeroMEVenta
        lblTitulo = "VENTA MONEDA EXTRANJERA NORMAL"
        Me.lblMonto = "Monto a Recibir"
End Select
TxtMontoPagar = "0.00"

'falta definir el objeto area agencia con que va a trabajar
lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , gsCodAge, ObjCMACAgenciaArea)
lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , gsCodAge, ObjCMACAgenciaArea)

End Sub

Private Sub txtBuscaPers_EmiteDatos()
lblPersNombre = txtBuscaPers.psDescripcion
lblPersDireccion = txtBuscaPers.sPersDireccion
nPersoneria = txtBuscaPers.PersPersoneria
fgDocs.Clear
fgDocs.FormaCabecera
fgDocs.Rows = 2
lsDocumento = ""
If txtBuscaPers <> "" Then
    lsDocumento = txtBuscaPers.sPersNroDoc
    Set fgDocs.Recordset = txtBuscaPers.rsDocPers
End If
fgDocs.RowHeight(-1) = 230
fgDocs.RowHeight(0) = 280
txtImporte.SetFocus
End Sub

Private Sub txtIdAut_KeyPress(KeyAscii As Integer)
Dim nMoneda As Integer
Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
 nMoneda = COMDConstantes.gMonedaExtranjera
        Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
Gtitular = Trim(txtBuscaPers.Text)
If Gtitular = "" Then
    MsgBox "Debe seleccionar persona que hace la operación. ", vbOKOnly + vbInformation, "Atención"
    Exit Sub
End If
If KeyAscii = 13 And Trim(txtIdAut.Text) <> "" Then
      Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        Set rs = oCapAut.SAA(Left(gsOpeCod, 5) & "1", Vusuario, "", Gtitular, CInt(nMoneda), CLng(txtIdAut.Text))
      Set oCapAut = Nothing
     If rs.State = 1 Then
       If rs.RecordCount > 0 Then
        txtImporte.Text = rs!nMontoAprobado
       Else
          MsgBox "No Existe este Id de Autorización para esta operación." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
          txtIdAut.Text = ""
       End If
       
     End If
 End If
 
 If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 13 Or KeyAscii = 8) Then
      KeyAscii = 0
 End If
End Sub

Private Sub txtImporte_Change()
  TxtMontoPagar = Format(Val(txtImporte.value) * Val(lblTpoCambioDia), "#,#0.00")
End Sub

Private Sub txtImporte_GotFocus()
With txtImporte
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdGuardar.SetFocus
End If
End Sub

Private Function Cargousu(ByVal NomUser As String) As String
 Dim rs As New ADODB.Recordset
 Dim oCons As COMDConstSistema.DCOMUAcceso
 Set oCons = New COMDConstSistema.DCOMUAcceso
 
 Set rs = oCons.Cargousu(NomUser)
  If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
  End If
 Set rs = Nothing
 'rs.Close
 Set oCons = Nothing
End Function
