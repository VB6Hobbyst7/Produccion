VERSION 5.00
Begin VB.Form frmCajeroCompraVentaEsp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compra y Venta ME - Tipo Cambio Especial"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPersona 
      Caption         =   "Datos de la Persona"
      Height          =   2745
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   6315
      Begin VB.TextBox txtIdAut 
         Height          =   330
         Left            =   4455
         TabIndex        =   18
         Top             =   225
         Visible         =   0   'False
         Width           =   1380
      End
      Begin SICMACT.FlexEdit fgDocs 
         Height          =   1230
         Left            =   1515
         TabIndex        =   10
         Top             =   1335
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2170
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Documento-N° Doc.-Tipo"
         EncabezadosAnchos=   "450-1500-1800-0"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.TxtBuscar txtBuscaPers 
         Height          =   330
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
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
         Left            =   3135
         TabIndex        =   19
         Top             =   315
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   14
         Top             =   585
         Width           =   5670
      End
      Begin VB.Label lblPersDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   13
         Top             =   930
         Width           =   5670
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Documentos :"
         Height          =   210
         Left            =   195
         TabIndex        =   12
         Top             =   1335
         Width           =   1200
      End
   End
   Begin VB.Frame fraTipoCambio 
      Height          =   1065
      Left            =   0
      TabIndex        =   2
      Top             =   2985
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
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin SICMACT.EditMoney txtImporte 
         Height          =   375
         Left            =   1665
         TabIndex        =   4
         Top             =   195
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
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
         TabIndex        =   8
         Top             =   285
         Width           =   180
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
         TabIndex        =   7
         Top             =   690
         Width           =   285
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
         TabIndex        =   6
         Top             =   660
         Width           =   1575
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
         TabIndex        =   5
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   3990
      TabIndex        =   1
      Top             =   3675
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5160
      TabIndex        =   0
      Top             =   3675
      Width           =   1155
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
      Left            =   1350
      TabIndex        =   17
      Top             =   0
      Width           =   4140
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
      Left            =   3990
      TabIndex        =   16
      Top             =   3195
      Width           =   1080
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
      Left            =   5190
      TabIndex        =   15
      Top             =   3135
      Width           =   1140
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   3915
      Top             =   3120
      Width           =   2430
   End
End
Attribute VB_Name = "frmCajeroCompraVentaEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oPers As COMDPersona.UCOMPersona
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsAgencia As String
Dim lbSalir As Boolean
Dim lsDocumento  As String

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
Dim lsBoleta As String
Dim nFicSal As Integer

Set oGen = New COMNContabilidad.NCOMContFunciones
Set oCajero = New COMNCajaGeneral.NCOMCajero

If ValidaInterfaz = False Then Exit Sub



If MsgBox("Desea grabar la Operación de Compra/Venta??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    Dim sPersLavDinero As String
    Dim nMontoLavDinero As Double
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
    Dim nMonto As Double
    Dim bLavDinero As Boolean
    
    bLavDinero = False
    nMonto = txtImporte.value
    'Realiza la Validación para el Lavado de Dinero
    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
    If Not clsExo.EsPersonaExoneradaLavadoDinero(txtBuscaPers) Then
        Set clsExo = Nothing
        sPersLavDinero = ""
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
        Set clsLav = Nothing
        
        If nMonto >= nMontoLavDinero Then
            sPersLavDinero = frmMovLavDinero.Inicia(Trim(txtBuscaPers.Text), Trim(lblPersNombre), Trim(lblPersDireccion), Trim(fgDocs.TextMatrix(1, 2)), , , nMonto, "000000002", gsOpeDesc)
            If sPersLavDinero = "" Then Exit Sub
            bLavDinero = True
        End If
    Else
        Set clsExo = Nothing
    End If
    lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If oCajero.GrabaCompraVenta(gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, CCur(txtImporte), CCur(lblTpoCambioDia), txtBuscaPers, bLavDinero) = 0 Then
        
        Dim oImp As COMNContabilidad.NCOMContImprimir
        Dim lsTexto As String
        Dim lbReimp As Boolean
        Set oImp = New COMNContabilidad.NCOMContImprimir
                
        lbReimp = True
        lsBoleta = oImp.ImprimeBoletaCompraVenta(lblTitulo, "", lblPersNombre, lblPersDireccion, lsDocumento, _
                    CCur(lblTpoCambioDia), gsOpeCod, CCur(txtImporte), CCur(TxtMontoPagar), gsNomAge, lsMovNro, sLpt, gsCodCMAC, , gbImpTMU)
        
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
        TxtMontoPagar = "0.0000"
        txtBuscaPers.SetFocus
    End If
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
Dim oOpe As COMDConstSistema.DCOMOperacion
Dim oTipCambio As COMDConstSistema.NCOMTipoCambio

Set oTipCambio = New COMDConstSistema.NCOMTipoCambio
gnTipCambioC = oTipCambio.EmiteTipoCambio(gdFecSis, TCCompraEsp)
gnTipCambioV = oTipCambio.EmiteTipoCambio(gdFecSis, TCVentaEsp)
Set oTipCambio = Nothing
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Set oOpe = New COMDConstSistema.DCOMOperacion

Me.Caption = gsOpeDesc
txtImporte.psSoles False
lbSalir = False
lblTpoCambioDia = Format(IIf(gsOpeCod = COMDConstSistema.gOpeCajeroMECompraEsp, gnTipCambioC, gnTipCambioV), "#,#0.0000")

If Val(lblTpoCambioDia) = 0 Then
    MsgBox "Tipo de Cambio Especial no ha sido Ingresado. Por favor Ingrese Tipo Cambio Especial del Día", vbInformation, "Aviso"
    lbSalir = True
    Exit Sub
End If

Select Case gsOpeCod
    Case COMDConstSistema.gOpeCajeroMECompraEsp
        lblTitulo = "COMPRA MONEDA EXT. ESPECIAL"
        Me.lblMonto = "Monto a Pagar"
    Case COMDConstSistema.gOpeCajeroMEVentaEsp
        lblTitulo = "VENTA MONEDA EXT. ESPECIAL"
        Me.lblMonto = "Monto a Recibir"
End Select

TxtMontoPagar = "0.0000"

'falta definir el objeto area agencia con que va a trabajar
'lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , gsCodAge, ObjCMACAgenciaArea)
'lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , gsCodAge, ObjCMACAgenciaArea)

Set oOpe = Nothing

End Sub
Private Sub txtBuscaPers_EmiteDatos()
lblPersNombre = txtBuscaPers.psDescripcion
lblPersDireccion = txtBuscaPers.sPersDireccion
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
 nMoneda = COMDConstantes.gMonedaExtranjera
        Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
Gtitular = Trim(txtBuscaPers.Text)
If Gtitular = "" Then
    MsgBox "Debe seleccionar persona que hace la operación. ", vbOKOnly + vbInformation, "Atención"
    Exit Sub
End If
If KeyAscii = 13 And Trim(txtIdAut.Text) <> "" Then
      Dim ocapaut As COMDCaptaGenerales.COMDCaptAutorizacion
      Set ocapaut = New COMDCaptaGenerales.COMDCaptAutorizacion
            Set rs = ocapaut.SAA(Left(gsOpeCod, 5) & "4", Vusuario, "", Gtitular, CInt(nMoneda), CLng(txtIdAut.Text))
      Set ocapaut = Nothing
      
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
  'TxtMontoPagar = Format(Val(txtImporte) * CCur(lblTpoCambioDia), "#,#0.0000")
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
