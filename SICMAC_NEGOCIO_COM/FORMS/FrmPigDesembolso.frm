VERSION 5.00
Begin VB.Form FrmPigDesembolso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Credito Pignoraticios - Desembolso"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "FrmPigDesembolso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraContenedor 
      Enabled         =   0   'False
      Height          =   795
      Index           =   2
      Left            =   105
      TabIndex        =   4
      Top             =   7350
      Width           =   7440
      Begin VB.TextBox txtInteres 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         TabIndex        =   7
         Top             =   150
         Width           =   1170
      End
      Begin VB.TextBox txtCostoCustodia 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   1515
         TabIndex        =   6
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txtCostoTasacion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1515
         TabIndex        =   5
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Neto a Recibir :"
         Height          =   255
         Index           =   7
         Left            =   5760
         TabIndex        =   12
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Impuesto :"
         Height          =   255
         Index           =   3
         Left            =   3075
         TabIndex        =   11
         Top             =   435
         Width           =   855
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Interes :"
         Height          =   255
         Index           =   2
         Left            =   3075
         TabIndex        =   10
         Top             =   165
         Width           =   780
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Costo Custodia :"
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   9
         Top             =   450
         Width           =   1260
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Costo Tasación :"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.Frame fraContenedor 
      Height          =   2985
      Index           =   1
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   7710
      Begin VB.Frame FraCredito 
         Height          =   735
         Left            =   105
         TabIndex        =   14
         Top             =   120
         Width           =   7530
         Begin VB.CommandButton cmdBuscar 
            Height          =   375
            Left            =   6900
            Picture         =   "FrmPigDesembolso.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Buscar ..."
            Top             =   225
            Width           =   420
         End
         Begin SICMACT.ActXCodCta AXCodCta 
            Height          =   435
            Left            =   135
            TabIndex        =   15
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   767
            Texto           =   "Crédito"
            EnabledCta      =   -1  'True
         End
      End
      Begin SICMACT.ActXPigDesCon axDesPigCon 
         Height          =   2070
         Left            =   90
         TabIndex        =   13
         Top             =   855
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3651
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6615
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5460
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblTasacion 
      Caption         =   "Label1"
      Height          =   270
      Left            =   300
      TabIndex        =   17
      Top             =   3165
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmPigDesembolso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnMontoTasacion As Currency
Dim lsPersCod As String
Dim lsPersNombre As String
Dim sPersCod As String, sNombre As String, sDireccion As String, sDocId As String, sTipoCuenta As String, sOperacion As String, sCuenta As String

Private Sub Limpiar()
   Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
   axDesPigCon.Limpiar
 End Sub

Private Function IniciaLavDinero() As String
Dim i As Long
Dim nRelacion As CaptacRelacPersona
Dim nMonto As Double
Dim oPersona As NCapMantenimiento
Dim oDatos As DPigContrato
Dim rsPersPigno As Recordset
'Dim sPersCod As String, sNombre As String, sDireccion As String, sDocId As String, sTipoCuenta As String, sOperacion As String, sCuenta As String

Set oDatos = New DPigContrato
    Set rsPersPigno = oDatos.dClientePigno(AXCodCta.NroCuenta)
    sCuenta = AXCodCta.NroCuenta
    sPersCod = rsPersPigno("cPersCod")
    sNombre = rsPersPigno("cPersnombre")
    sDireccion = rsPersPigno("cPersDireccDomicilio")
    sDocId = rsPersPigno("cPersIdNro")
    sTipoCuenta = ""
    sOperacion = ""
    Set oDatos = Nothing

    nMonto = axDesPigCon.neto1
    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, "", sOperacion, False, sTipoCuenta)
       
End Function

Private Function EsExoneradaLavadoDinero() As Boolean
Dim bExito As Boolean
Dim clsExo As NCapServicios
bExito = True

    Set clsExo = New NCapServicios
    
    If Not clsExo.EsPersonaExoneradaLavadoDinero(sPersCod) Then bExito = False

    Set clsExo = Nothing
    EsExoneradaLavadoDinero = bExito
    
End Function

Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As nPigValida
Dim oDatos As DPigContrato
Dim lsmensaje As String

    Set lrValida = New ADODB.Recordset
    Set loValContrato = New nPigValida
        Set lrValida = loValContrato.nValidaDesembolsoCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If

    Set loValContrato = Nothing
    If lrValida Is Nothing Then
        Limpiar
        Set lrValida = Nothing
        Exit Sub
        
    End If
    lbOk = fgMuestraPig_AXPigDesCon(psNroContrato, Me.axDesPigCon, False)
    Set lrValida = Nothing
        
    fraCredito.Enabled = False
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
        
    Set oDatos = New DPigContrato
    lblTasacion = oDatos.dObtieneValorTasacion(psNroContrato)
    
    Set oDatos = Nothing
    AXCodCta.Enabled = False
 
 Exit Sub
ControlError:
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsEstados As String
Dim loPersContrato As DPigContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As UProdPersona

On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    If Not loPers Is Nothing Then
        lsPersCod = loPers.sPersCod
        lsPersNombre = loPers.sPersNombre
    Else
        Exit Sub
    End If
    
Set loPers = Nothing

lsEstados = gPigEstRegis

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New DPigContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New UProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    fraCredito.Enabled = True
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub

Private Sub cmdGrabar_Click()

Dim loContFunct As NContFunciones
Dim loGrabarDesem As NPigContrato
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String

Dim lnSaldoCap As Currency, lnCapital As Currency, lnComision As Currency
Dim lnInteresComp As Currency, lnImpuesto As Currency
Dim lnMontoEntregar As Currency
Dim oImpre As NPigImpre
Dim oPrevio As Previo.clsPrevio
Dim lsCadImprimir As String
Dim lnPlazo As Integer
Dim lnMovNro As Long, nMonto As Long

lsCuenta = AXCodCta.NroCuenta
lnSaldoCap = Me.axDesPigCon.prestamo1 'Prestamo
lnComision = Me.axDesPigCon.comision1      'Comision
lnMontoEntregar = lnSaldoCap - lnComision
lsPersNombre = axDesPigCon.listaClientes.ListItems(1).ListSubItems(1)


lnPlazo = Me.axDesPigCon.Piezas ' es el plazo

lnImpuesto = 0

If MsgBox(" Grabar Desembolso de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
  ' **********************************************************************
  'Realiza la Validación para el Lavado de Dinero
    Dim clsLav As nCapDefinicion
    Dim nMontoLavDinero As Double, nTC As Double
    Dim sPersLavDinero As String
    
    sPersLavDinero = ""
    Set clsLav = New nCapDefinicion
    If clsLav.EsOperacionEfectivo(Trim(gPigOpeDesembolsoEFE)) Then
        If Not EsExoneradaLavadoDinero() Then

            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            
            Dim clsTC As nTipoCambio
            Set clsTC = New nTipoCambio
            nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set clsTC = Nothing
            
            nMonto = axDesPigCon.neto1
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                sPersLavDinero = IniciaLavDinero()
                If sPersLavDinero = "" Then Exit Sub
            End If
        End If
    End If
  ' *******************************
    
        Set loContFunct = New NContFunciones
         lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarDesem = New NPigContrato
        lnMovNro = loGrabarDesem.nDesembolsoCredPignoraticio(lsCuenta, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, lnComision, lnImpuesto, lblTasacion, lnPlazo, sPersLavDinero, sPersCod)
        Set loGrabarDesem = Nothing
            'IMPRESION DE LAVADO DE DINERO
        If sPersLavDinero <> "" Then
          Dim oBoleta As NCapImpBoleta
          Set oBoleta = New NCapImpBoleta
           Do
               oBoleta.ImprimeBoletaLavadoDinero gsNomCmac, gsNomAge, gdFecSis, sCuenta, sNombre, sDocId, sDireccion, _
                        sNombre, sDocId, sDireccion, sNombre, sDocId, sDireccion, sOperacion, nMonto, sLpt, , , Trim(Left("", 15))
            Loop Until MsgBox("¿Desea reimprimir Boleta de Lavado de Dinero?", vbQuestion + vbYesNo, "Aviso") = vbNo
            Set oBoleta = Nothing
       End If
      
        Set oImpre = New NPigImpre
        Call oImpre.ImpreReciboDesembolsoContrato(gsInstCmac, gsNomAge, lsFechaHoraGrab, lsCuenta, lsPersNombre, lnSaldoCap, _
                    lnComision, lnMontoEntregar, gsCodUser, lnMovNro, sLpt, "", gsCodCMAC, gcEmpresaRUC)
        
        Do While MsgBox("Desea Reimprimir Comprobante de Desembolso? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes
            Call oImpre.ImpreReciboDesembolsoContrato(gsInstCmac, gsNomAge, lsFechaHoraGrab, lsCuenta, lsPersNombre, lnSaldoCap, _
                lnComision, lnMontoEntregar, gsCodUser, lnMovNro, sLpt, "", gsCodCMAC, gcEmpresaRUC)
        Loop
        
      Set oImpre = Nothing
        Limpiar
        fraCredito.Enabled = True
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
     
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Limpiar
    AXCodCta.Enabled = True
    Me.Icon = LoadPicture(App.path & "\bmps\cm.ico")
End Sub

