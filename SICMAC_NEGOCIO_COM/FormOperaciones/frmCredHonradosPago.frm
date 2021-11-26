VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredHonradosPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devolución por Honramiento de Créditos"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   Icon            =   "frmCredHonradosPago.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstDatos 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Registro de Créditos Honrados"
      TabPicture(0)   =   "frmCredHonradosPago.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPagos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fraPagos 
         Caption         =   "Datos del Pago"
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
         Height          =   1335
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   9735
         Begin VB.TextBox txtMontoHonrado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1320
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   6840
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8400
            TabIndex        =   11
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblMontoPagar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6480
            TabIndex        =   21
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label lblITF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3600
            TabIndex        =   20
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label7 
            Caption         =   "Pago Total:"
            Height          =   255
            Left            =   5400
            TabIndex        =   18
            Top             =   400
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "I.T.F.:"
            Height          =   255
            Left            =   3000
            TabIndex        =   17
            Top             =   400
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Monto a Pagar:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   400
            Width           =   1215
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos del Crédito"
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
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9735
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   4080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1440
            Width           =   600
         End
         Begin VB.CommandButton cmdBuscaCuenta 
            Caption         =   "..."
            Height          =   375
            Left            =   3960
            TabIndex        =   2
            Top             =   360
            Width           =   375
         End
         Begin SICMACT.ActXCodCta ActXCodCta 
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            Texto           =   "Crédito:"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblSaldoPendiente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   7800
            TabIndex        =   15
            Top             =   1440
            Width           =   1680
         End
         Begin VB.Label lblMontoHonrado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4680
            TabIndex        =   14
            Top             =   1440
            Width           =   1665
         End
         Begin VB.Label Label4 
            Caption         =   "Saldo Pendiente: "
            Height          =   255
            Left            =   6480
            TabIndex        =   13
            Top             =   1515
            Width           =   1215
         End
         Begin VB.Label lblTitular 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   10
            Top             =   960
            Width           =   8640
         End
         Begin VB.Label Label1 
            Caption         =   "Titular: "
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "DOI: "
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   1515
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Monto Honrado: "
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   1515
            Width           =   1215
         End
         Begin VB.Label lblDOI 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   6
            Top             =   1440
            Width           =   1800
         End
      End
   End
End
Attribute VB_Name = "frmCredHonradosPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredHonradosPago
'***     Descripcion:      Realizar los Pagos Por Honramiento
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     10/10/2013 08:30:00 AM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Private fnITF As Double
Private fnRedondeoITF As Double
Private fnMoneda As Integer
Private fsPersCod As String
Private fnMovRef As Long
'Lavado de Dinero
Private fbExoneradaLavado As Boolean
'Variables agregadas para el uso de los Componentes
Private bOperacionEfectivo As Boolean
Private nMontoLavDinero As Double
Private nTC As Double

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CargaDatos (Trim(ActXCodCta.NroCuenta))
End If
End Sub

Private Sub cmdBuscaCuenta_Click()
Dim oPersona As COMDPersona.UCOMPersona
Dim loPersCreditos As COMDCredito.DCOMCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

LimpiarDatos

On Error GoTo ControlError

Set oPersona = Nothing
Set oPersona = New COMDPersona.UCOMPersona
Set oPersona = frmBuscaPersona.inicio
If oPersona Is Nothing Then Exit Sub

fsPersCod = oPersona.sPersCod
If Trim(oPersona.sPersCod) <> "" Then
    Set loPersCreditos = New COMDCredito.DCOMCredito
    Set lrCreditos = loPersCreditos.CreditosHonradosAPagar(oPersona.sPersCod)
    Set loPersCreditos = Nothing
End If

If Not (lrCreditos.EOF And lrCreditos.BOF) Then
Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.inicio(oPersona.sPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        ActXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        ActXCodCta.Enabled = False
        Call ActXCodCta_KeyPress(13)
    End If
Else
    MsgBox "Persona No cuenta con créditos Cancelados", vbInformation, "Aviso"
End If
Set loCuentas = Nothing
Exit Sub
ControlError:
MsgBox "Error: " & err.Description, vbCritical, "Error"
End Sub

Private Sub cmdCancelar_Click()
LimpiarDatos
End Sub

Private Sub cmdGuardar_Click()
Dim oCredito As New COMNCredito.NCOMCredito
Dim lnMovNro As Long
Dim sImpreBoleta As String
Dim vPrevio As New previo.clsprevio
Dim loLavDinero As New frmMovLavDinero
Dim sPersLavDinero As String
Dim nmoneda As Integer
Dim nMonto As Double
Dim oPersonaSPR As New UPersona_Cli
Dim oPersonaU As New COMDPersona.UCOMPersona
Dim nTipoConBN As Integer
Dim sConPersona As String
Dim pbClienteReforzado As Boolean
Dim rsAgeParam As Recordset
Dim objCred As New COMNCredito.NCOMCredito
Dim fnCondicion As Integer
Dim regPersonaRealizaPago As Boolean

nMonto = CDbl(txtMontoHonrado.Text)
nmoneda = CLng(Mid(ActXCodCta.NroCuenta, 9, 1))

On Error GoTo ErrorPago
If ValidaDatos Then
    If MsgBox("Desea Grabar el Pago?", vbInformation + vbYesNo, "Aviso") = vbYes Then
               
        sPersLavDinero = ""
        If bOperacionEfectivo Then
            If Not fbExoneradaLavado Then
                If CDbl(txtMontoHonrado.Text) >= Round(nMontoLavDinero * IIf(Mid(ActXCodCta.NroCuenta, 9, 1) = 1, 1, nTC), 2) Then
                     Call IniciaLavDinero(loLavDinero)
                     sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, ActXCodCta.NroCuenta, Me.Caption, True, "", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                     If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                End If
            End If
        End If
        
        If loLavDinero.OrdPersLavDinero = "Exit" Then
            regPersonaRealizaPago = False
            pbClienteReforzado = False
            fnCondicion = 0
    
            oPersonaSPR.RecuperaPersona fsPersCod
                                
            If oPersonaSPR.Personeria = 1 Then
                If oPersonaSPR.Nacionalidad <> "04028" Then
                    sConPersona = "Extranjera"
                    fnCondicion = 1
                    pbClienteReforzado = True
                ElseIf oPersonaSPR.Residencia <> 1 Then
                    sConPersona = "No Residente"
                    fnCondicion = 2
                    pbClienteReforzado = True
                ElseIf oPersonaSPR.RPeps = 1 Then
                    sConPersona = "PEPS"
                    fnCondicion = 4
                    pbClienteReforzado = True
                ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                    If nTipoConBN = 1 Or nTipoConBN = 3 Then
                        sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                        fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                        pbClienteReforzado = True
                    End If
                End If
            Else
                If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                    If nTipoConBN = 1 Or nTipoConBN = 3 Then
                        sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                        fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                        pbClienteReforzado = True
                    End If
                End If
            End If
            
            If pbClienteReforzado Then
                MsgBox "El Cliente: " & Trim(lblTitular.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                frmPersRealizaOpeGeneral.Inicia Me.Caption & " (Persona " & sConPersona & ")", gCredPagHonramiento
                regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                
                If Not regPersonaRealizaPago Then
                    MsgBox "Se va a proceder a Anular el Pago", vbInformation, "Aviso"
                    Exit Sub
                End If
            Else
                fnCondicion = 0
                pbClienteReforzado = False
                
                Set rsAgeParam = objCred.obtieneCredPagoCuotasAgeParam(gsCodAge)
                Set objCred = Nothing
                
                If Mid(ActXCodCta.NroCuenta, 9, 1) = 2 Then
                    nMonto = Round(nMonto * nTC, 2)
                End If
            
                If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                    If nMonto >= rsAgeParam!nMontoMin And nMonto <= rsAgeParam!nMontoMax Then
                        frmPersRealizaOpeGeneral.Inicia Me.Caption, gCredPagHonramiento
                        regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                        If Not regPersonaRealizaPago Then
                            MsgBox "Se va a proceder a Anular el Pago", vbInformation, "Aviso"
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
            
            Set oPersonaU = Nothing
            Set oPersonaSPR = Nothing
            Set objCred = Nothing
        End If
        
        Call oCredito.PagoCreditoHonrado(Trim(ActXCodCta.NroCuenta), CDbl(txtMontoHonrado.Text), CDbl(lblITF.Caption), _
                                        gdFecSis, lnMovNro, fnMovRef, gsCodAge, gsCodUser, CDbl(Me.lblMontoHonrado.Caption), sImpreBoleta, Trim(lblTitular.Caption), _
                                        gsNomAge, gsInstCmac, gsCodCMAC, gbImpTMU)
        
        If lnMovNro > 0 Then
            Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , lnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224

            If regPersonaRealizaPago Then
                frmPersRealizaOpeGeneral.InsertaPersonasOperacion lnMovNro, Trim(ActXCodCta.NroCuenta), fnCondicion
                regPersonaRealizaPago = False
            End If
            
            vPrevio.PrintSpool sLpt, sImpreBoleta
            Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
                vPrevio.PrintSpool sLpt, sImpreBoleta
            Loop
            Set vPrevio = Nothing
            
            MsgBox "Pago realizado Satisfactoriamente", vbInformation, "Aviso"
            LimpiarDatos
            'INICIO JHCU ENCUESTA 16-10-2019
            Encuestas gsCodUser, gsCodAge, "ERS0292019", "100218"
            'FIN
        End If
        Set oCredito = Nothing
    End If
End If

Exit Sub
ErrorPago:
     MsgBox err.Description, vbCritical, "Error"
     LimpiarDatos
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CargaMoneda
txtMontoHonrado.Enabled = False
cmdGuardar.Enabled = False
LimpiarDatos
End Sub

Private Sub txtMontoHonrado_GotFocus()
    fEnfoque txtMontoHonrado
End Sub

Private Sub txtMontoHonrado_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoHonrado, KeyAscii, 15)
cmdGuardar.Enabled = False
lblITF.Caption = "0.00"
LblMontoPagar.Caption = "0.00"
If KeyAscii <> 13 Then Exit Sub

If Not IsNumeric(txtMontoHonrado.Text) Then
    MsgBox "Ingrese un monto válido", vbInformation, "Aviso"
    lblITF.Caption = "0.00"
    LblMontoPagar.Caption = "0.00"
    txtMontoHonrado.Text = "0.00"
    cmdGuardar.Enabled = False
    Exit Sub
End If

If IsNumeric(Trim(lblSaldoPendiente.Caption)) Then
    If CDbl(Trim(lblSaldoPendiente.Caption)) < CDbl(Trim(txtMontoHonrado.Text)) Then
        MsgBox "El monto Ingresado no debe ser mayor al Saldo Pendiente.", vbInformation, "Aviso"
        txtMontoHonrado.Text = lblSaldoPendiente.Caption
        cmdGuardar.Enabled = False
        Exit Sub
    End If
End If

Dim oITF As COMDConstSistema.FCOMITF
Set oITF = New COMDConstSistema.FCOMITF
oITF.fgITFParametros

fnITF = oITF.CortaDosITF(CDbl(txtMontoHonrado.Text) * oITF.gnITFPorcent)

lblITF.Caption = Format(fnITF, "###," & String(15, "#") & "#0.00")
fnRedondeoITF = fgDiferenciaRedondeoITF(CCur(fnITF))
If fnRedondeoITF > 0 Then
    lblITF.Caption = Format(CCur(fnITF) - fnRedondeoITF, "###," & String(15, "#") & "#0.00")
End If

LblMontoPagar.Caption = Format(CDbl(txtMontoHonrado.Text) + CDbl(lblITF.Caption), "###," & String(15, "#") & "#0.00")

Set oITF = Nothing
cmdGuardar.Enabled = True
If KeyAscii = 13 Then
    cmdGuardar.SetFocus
End If
End Sub

Private Sub txtMontoHonrado_LostFocus()
     If Len(Trim(txtMontoHonrado.Text)) = 0 Then
         txtMontoHonrado.Text = "0.00"
    End If
    txtMontoHonrado.Text = Format(txtMontoHonrado.Text, "###," & String(15, "#") & "#0.00")
End Sub

Private Sub CargaDatos(ByVal psCtaCod As String)
Dim oCredito As New COMNCredito.NCOMCredito
Dim rsCredito As ADODB.Recordset
Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios
Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsTC As New COMDConstSistema.NCOMTipoCambio
 
If oCredito.EsCreditoHonrado(psCtaCod) Then
    
    Set rsCredito = oCredito.ObtenerCreditosHonrados(psCtaCod)
    
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        lblTitular.Caption = Trim(rsCredito!cPersNombre)
        lblDOI.Caption = Trim(rsCredito!cPersIDnro)
        lblMontoHonrado.Caption = Format(rsCredito!nMontoHonrado, "###," & String(15, "#") & "#0.00")
        lblSaldoPendiente.Caption = Format(CDbl(rsCredito!nMontoHonrado) - CDbl(rsCredito!nMontoDevuelto), "###," & String(15, "#") & "#0.00")
        txtMontoHonrado.MaxLength = Len(lblSaldoPendiente.Caption) + 1
        fnMoneda = CInt(Mid(psCtaCod, 9, 1))
        fnMovRef = CLng(Trim(rsCredito!nMovNro))
        If fnMoneda = 2 Then
            txtMontoHonrado.ForeColor = &H289556
            cmbMoneda.Text = "$."
        Else
            cmbMoneda.Text = "S/."
            txtMontoHonrado.ForeColor = vbBlue
        End If
        cmbMoneda.Enabled = False
        txtMontoHonrado.Enabled = True
        txtMontoHonrado.SetFocus
    Else
        lblTitular.Caption = ""
        lblDOI.Caption = ""
        cmbMoneda.Enabled = True
        txtMontoHonrado.Enabled = False
    End If
    
    fbExoneradaLavado = clsExo.EsPersonaExoneradaLavadoDinero(fsPersCod)
    bOperacionEfectivo = clsLav.EsOperacionEfectivo(Trim(gCredPagHonramiento))
    nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    
    Set clsTC = Nothing
    Set clsExo = Nothing
    Set clsLav = Nothing
    Set oCredito = Nothing
    txtMontoHonrado.SetFocus
Else
    MsgBox "Credito no Esta Honrado.", vbInformation, "Aviso"
    LimpiarDatos
    Exit Sub
End If
End Sub


Private Sub CargaMoneda()
cmbMoneda.Clear
cmbMoneda.AddItem "S/."
cmbMoneda.AddItem "$."
End Sub

Private Sub LimpiarDatos()
ActXCodCta.NroCuenta = ""
ActXCodCta.Enabled = True
ActXCodCta.CMAC = gsCodCMAC
ActXCodCta.Age = gsCodAge
ActXCodCta.EnabledCMAC = False
lblTitular.Caption = ""
lblDOI.Caption = ""
cmbMoneda.Enabled = True
fnMoneda = 0
txtMontoHonrado.Text = ""
fsPersCod = ""
txtMontoHonrado.Enabled = False
cmdGuardar.Enabled = False
lblMontoHonrado.Caption = ""
lblSaldoPendiente.Caption = ""
fnMovRef = 0
lblITF.Caption = ""
LblMontoPagar.Caption = ""
fsPersCod = ""

End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = True

If Trim(txtMontoHonrado.Text) = "" Or Trim(txtMontoHonrado.Text) = "0.00" Then
    MsgBox "Ingrese el Monto A Pagar", vbInformation, "Aviso"
    ValidaDatos = False
    txtMontoHonrado.SetFocus
    Exit Function
End If

If IsNumeric(Trim(txtMontoHonrado.Text)) = False Then
    MsgBox "Valor Inválido del Monto A Pagar", vbInformation, "Aviso"
    ValidaDatos = False
    txtMontoHonrado.SetFocus
    Exit Function
End If

End Function

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
poLavDinero.TitPersLavDinero = fsPersCod
poLavDinero.OrdPersLavDinero = fsPersCod
End Sub
