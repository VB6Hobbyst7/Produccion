VERSION 5.00
Begin VB.Form frmPenalidadEcoTaxi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Penalidad por Convenio - Créditos EcoTaxi"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   Icon            =   "frmPenalidadEcoTaxi.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      TabIndex        =   37
      Top             =   5070
      Width           =   1050
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   80
      TabIndex        =   23
      Top             =   5070
      Width           =   1050
   End
   Begin VB.CommandButton btnGrabar 
      Caption         =   "&Grabar"
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
      Height          =   345
      Left            =   5880
      TabIndex        =   22
      Top             =   5070
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Caption         =   "Penalidad"
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
      Height          =   3015
      Left            =   50
      TabIndex        =   11
      Top             =   2000
      Width           =   8055
      Begin VB.Frame fraRetraso 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   120
         TabIndex        =   28
         Top             =   525
         Width           =   5895
         Begin VB.TextBox txtDiasRetraso 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   30
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Monto x Día S/.:"
            Height          =   255
            Left            =   2880
            TabIndex        =   32
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label txtMontoxDia 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4200
            TabIndex        =   31
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label9 
            Caption         =   "Días de Retraso:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.TextBox txtComentarios 
         Height          =   645
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1125
         Width           =   6375
      End
      Begin VB.ComboBox ddlMotivoPenalidad 
         Height          =   315
         ItemData        =   "frmPenalidadEcoTaxi.frx":030A
         Left            =   1560
         List            =   "frmPenalidadEcoTaxi.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   3825
      End
      Begin VB.Label txtMontoFavorInstitucion 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Top             =   2620
         Width           =   1020
      End
      Begin VB.Label txtMontoFavorCliente 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   2260
         Width           =   1020
      End
      Begin VB.Label txtMontoPenalidad 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         TabIndex        =   19
         Top             =   1900
         Width           =   1020
      End
      Begin VB.Label Label12 
         Caption         =   "A favor de la Institución S/.:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "A favor del Cliente S/.:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Monto Penalidad S/.:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Comentarios:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Motivo:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   260
         Width           =   615
      End
   End
   Begin VB.Frame fraDatosCredito 
      Caption         =   "Crédito"
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
      Height          =   1935
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   840
         Left            =   5880
         TabIndex        =   24
         Top             =   120
         Width           =   2115
         Begin VB.ListBox LstCred 
            Height          =   450
            Left            =   75
            TabIndex        =   25
            Top             =   225
            Width           =   1980
         End
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   900
      End
      Begin SICMACT.ActXCodCta ActxCtaCod 
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   714
         Texto           =   "Cuenta N°:"
         EnabledCta      =   -1  'True
      End
      Begin VB.Label Label14 
         Caption         =   "Cta Recaudo:"
         Height          =   255
         Left            =   4920
         TabIndex        =   36
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label txtCtaCodRecaudo 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6000
         TabIndex        =   35
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblPersCodConce 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPersCodTit 
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   840
         TabIndex        =   33
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Placa:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   615
      End
      Begin VB.Label txtPlaca 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label txtConcesionarioCtaCod 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6000
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Cta Afectada:"
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label txtConcesionarioNom 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Afectado:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label txtProductoNom 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "             EcoTaxi"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label txtClienteNom 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPenalidadEcoTaxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fnPenalidadEcoTaxixAtrasoEntregaxDia As Double
Dim fnPenalidadEcoTaxixVehiculoDistintoAcordado As Double

Private Sub Form_Load()
    CentraForm Me
    CargarCombos
    LimpiaPantalla
    CargarConstantes
End Sub
Private Sub LimpiaPantalla()
    LimpiaControles Me, True
    InicializaCombos Me
    Me.ActXCtaCod.NroCuenta = ""
    Me.ActXCtaCod.CMAC = gsCodCMAC
    Me.ActXCtaCod.Age = gsCodAge
    Me.ActXCtaCod.Prod = "517"
    Me.txtClienteNom.Caption = ""
    Me.txtCtaCodRecaudo.Caption = ""
    Me.lblPersCodTit.Caption = ""
    Me.txtConcesionarioNom.Caption = ""
    Me.lblPersCodConce.Caption = ""
    Me.txtConcesionarioCtaCod.Caption = ""
    Me.txtProductoNom.Caption = "             EcoTaxi"
    Me.txtDiasRetraso.Text = "0"
    Me.txtMontoxDia.Caption = "0.00"
    Me.txtcomentarios.Text = ""
    Me.txtMontoPenalidad.Caption = "0.00"
    Me.txtMontoFavorCliente.Caption = "0.00"
    Me.txtMontoFavorInstitucion.Caption = "0.00"
End Sub
Private Sub btnBuscar_Click()
    Dim oCredito As COMDCredito.DCOMCredito
    Dim R As ADODB.Recordset
    Dim oPers As COMDPersona.UCOMPersona

    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigVenc, gColocEstVigNorm, gColocEstRefMor, gColocEstRefVenc, gColocEstRefNorm))
        Do While Not R.EOF
            If Mid(R!cCtaCod, 6, 3) = "517" And R!cAgeCodAct = Right(gsCodAge, 2) Then
                LstCred.AddItem R!cCtaCod
            End If
            R.MoveNext
        Loop
        R.Close
        
        If LstCred.ListCount = 0 Then
            MsgBox "El Cliente No Tiene Creditos EcoTaxi Vigentes en esta Agencia", vbInformation, "Aviso"
            btnBuscar.SetFocus
        ElseIf LstCred.ListCount = 1 Then
            LstCred.ListIndex = 0
            Me.ActXCtaCod.NroCuenta = LstCred.Text
            ActXCtaCod_KeyPress (13)
        End If
    End If
    Set R = Nothing
    Set oCredito = Nothing
    Set oPers = Nothing
End Sub
Private Sub LstCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            Me.ActXCtaCod.NroCuenta = LstCred.Text
            Me.ActXCtaCod.SetFocusCuenta
        End If
    End If
End Sub
Private Sub CargarCombos()
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Set oCons = New COMDConstantes.DCOMConstantes
    Set rs = oCons.RecuperaConstantes(9024) 'Lista Motivos Penalidades
    Call Llenar_Combo_con_Recordset(rs, ddlMotivoPenalidad)
End Sub
Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim oCred As COMDCredito.DCOMCredito
    Dim rsCred As ADODB.Recordset
    
    Set oCred = New COMDCredito.DCOMCredito
    Set rsCred = oCred.RecuperaDatosCredEcoTaxixPenalidad(psCtaCod)

    CargaDatos = False
    If Not RSVacio(rsCred) Then
        Me.txtPlaca.Caption = rsCred!cPlaca
        Me.txtClienteNom.Caption = rsCred!cPersNombreTitular
        Me.txtCtaCodRecaudo.Caption = rsCred!cCtaCodAhorroRecaudo
        Me.lblPersCodTit.Caption = rsCred!cPersCodTitular
        Me.txtConcesionarioNom.Caption = rsCred!cPersNombreConcesionario
        Me.txtConcesionarioCtaCod.Caption = rsCred!cCtaCodAhorroConcesionario
        Me.lblPersCodConce.Caption = rsCred!cPersCodConcesionario
        CargaDatos = True
    End If
End Function
Private Sub btnGrabar_Click()
    Dim oCapta As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oBase As COMDCredito.DCOMCredActBD
    Dim oDCred As COMDCredito.DCOMCredito
    Dim oITF As COMDConstSistema.FCOMITF
    Dim rsCapta As ADODB.Recordset, rsPenalidad As ADODB.Recordset
    Dim sMovNro As String
    Dim bTransac As Boolean
    Dim pMatDatosAhoAbo As Variant
    Dim nITFAbono As Double, nRedondeoITF As Double
    Dim nMovNro As Long
    
    Dim lnMotivoPenalidad As Integer
    Dim lsPersCodConce As String, lsCtaCodAhoConce As String, lsOperacionDebitoConcesionario As String
    Dim lsPersCodTit As String, lsCtaCodAhoTit As String
    Dim lnMontoCargoConce As Double, lnITFCargoConce As Double, lnITFRedondeoConce As Double
    Dim lnMontoDepositoTit As Double, lnITFDepositoTit As Double, lnITFRedondeoTit As Double
    Dim lbRedondeoITFConce As Boolean
    Dim lbRedondeoITFTit As Boolean
    
    Dim lnDiasRetraso As Integer
    Dim lnMontoxDia As Double

    If validaDatos = False Then Exit Sub
    
    If MsgBox("¿Esta seguro de registrar la presente Penalidad EcoTaxi?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If

On Error GoTo ErrorGrabar
    ReDim pMatDatosAhoAbo(14)
    pMatDatosAhoAbo(0) = "" 'Cuenta de Ahorros
    pMatDatosAhoAbo(1) = "0.00" 'Monto de Apertura
    pMatDatosAhoAbo(2) = "0.00" 'Interes Ganado de Abono
    pMatDatosAhoAbo(3) = "0.00" 'Interes Ganado de Retiro Gastos
    pMatDatosAhoAbo(4) = "0.00" 'Interes Ganado de Retiro Cancelaciones
    pMatDatosAhoAbo(5) = "0.00" 'Monto de Abono
    pMatDatosAhoAbo(6) = "0.00" 'Monto de Retiro de Gastos
    pMatDatosAhoAbo(7) = "0.00" 'Monto de Retiro de Cancelaciones
    pMatDatosAhoAbo(8) = "0.00" 'Saldo Disponible Abono
    pMatDatosAhoAbo(9) = "0.00" 'Saldo Contable Abono
    pMatDatosAhoAbo(10) = "0.00" 'Saldo Disponible Retiro de Gastos
    pMatDatosAhoAbo(11) = "0.00" 'Saldo Contable Retiro de Gastos
    pMatDatosAhoAbo(12) = "0.00" 'Saldo Disponible Retiro de Cancelaciones
    pMatDatosAhoAbo(13) = "0.00" 'Saldo Contable Retiro de Cancelaciones
    
    Set oBase = New COMDCredito.DCOMCredActBD
    Set oDCred = New COMDCredito.DCOMCredito
    Set oCapta = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsCapta = New ADODB.Recordset

    lnMotivoPenalidad = CInt(Trim(Right(ddlMotivoPenalidad.Text, 5)))
    
    Set rsPenalidad = oDCred.DamePenalidadxMotivo(Me.ActXCtaCod.NroCuenta, lnMotivoPenalidad)
    If Not RSVacio(rsPenalidad) Then
        MsgBox "La Penalidad por el presente motivo para este Crédito ya fue registrada el " & Format(rsPenalidad!dRegistroPenalidad, gsFormatoFechaHoraView), vbExclamation, "Aviso"
        Exit Sub
    End If
    
    lsPersCodConce = Me.lblPersCodConce.Caption
    lsCtaCodAhoConce = Trim(Me.txtConcesionarioCtaCod)
    lnMontoCargoConce = CDbl(Me.txtMontoPenalidad.Caption)
    lnITFCargoConce = DameMontoITF(lnMontoCargoConce, lbRedondeoITFConce, lnITFRedondeoConce)

    lsPersCodTit = Me.lblPersCodTit.Caption
    lsCtaCodAhoTit = Trim(Me.txtCtaCodRecaudo.Caption)
    lnMontoDepositoTit = CDbl(Me.txtMontoFavorCliente.Caption)
    lnITFDepositoTit = DameMontoITF(lnMontoDepositoTit, lbRedondeoITFTit, lnITFRedondeoTit)

    lnDiasRetraso = CInt(IIf(IsNumeric(Me.txtDiasRetraso.Text), Me.txtDiasRetraso.Text, 0))
    lnMontoxDia = CDbl(IIf(IsNumeric(Me.txtMontoxDia.Caption), Me.txtMontoxDia.Caption, 0))
        
    Set rsCapta = oCapta.GetDatosCuentaAho(lsCtaCodAhoConce)
    If rsCapta!nSaldo < (lnMontoCargoConce + lnITFCargoConce) Then
        MsgBox "El Saldo de la Cta del Concesionario es menor de la Penalidad a cobrar", vbInformation, "Aviso"
        Exit Sub
    End If

    Select Case lnMotivoPenalidad
        Case 1
            lsOperacionDebitoConcesionario = gAhoCargoCtaxPenalidadEcotaxixDiasRetrasoEntrega
        Case 2
            lsOperacionDebitoConcesionario = gAhoCargoCtaxPenalidadEcotaxixVehiculoDistintoAcordado
        Case Else
            MsgBox "Operación desconocida, el proceso no continuará", vbInformation, "Aviso"
            Exit Sub
    End Select

    bTransac = False
    Call oBase.dBeginTrans
    bTransac = True

    sMovNro = oBase.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oBase.InsertaMov(sMovNro, gOtrOpePenalidadEcoTaxi, "Penalidad EcoTaxi", gMovEstContabMovContable, gMovFlagVigente)
    nMovNro = oBase.GetnMovNro(sMovNro)

    'Retiramos de la Cuenta de Ahorro del Concesionario
    oBase.CapCargoCuentaAho pMatDatosAhoAbo, lsCtaCodAhoConce, lnMontoCargoConce, lsOperacionDebitoConcesionario, sMovNro, "Cargo a Cta x Penalidad EcoTaxi", gdFecSis, , , , , , , , , , , , , , True, lnITFCargoConce, False, gITFCobroCargo
    If lbRedondeoITFConce Then
        Call oBase.InsertaMovRedondeoITF(sMovNro, 1, lnITFCargoConce + lnITFRedondeoConce, lnITFCargoConce)
    End If

    Select Case lnMotivoPenalidad
        Case 1
            'Depositamos toda la Penalidad a favor del cliente
            oBase.CapAbonoCuentaAho pMatDatosAhoAbo, lsCtaCodAhoTit, lnMontoDepositoTit, gAhoDepCtaxPenalidadEcotaxixDiasRetrasoEntrega, sMovNro, "Depósito a Cta x Penalidad EcoTaxi", , , , , , , gdFecSis, , True, lnITFDepositoTit, False, gITFCobroCargo
            If lbRedondeoITFTit Then
                Call oBase.InsertaMovRedondeoITF(sMovNro, 2, lnITFDepositoTit + lnITFRedondeoTit, lnITFDepositoTit)
            End If
        Case 2
            'En la config del Asiento se hara la jugada de la referencia de la cta de la caja
    End Select
    
    'Detalle Penalidad
    Call oBase.dRegistrarPenalidad(nMovNro, lnMotivoPenalidad, Me.ActXCtaCod.NroCuenta, lsCtaCodAhoTit, lsPersCodTit, lsCtaCodAhoConce, lsPersCodConce, lnDiasRetraso, lnMontoxDia, Trim(Me.txtcomentarios.Text), CDbl(Me.txtMontoPenalidad.Caption), lnMontoDepositoTit, CDbl(Me.txtMontoFavorInstitucion.Caption))

    Call oBase.dCommitTrans
    bTransac = False

    MsgBox "Se ha realizado el cobro respectivo de la Penalidad al Concesionario", vbInformation, "Aviso"
    If MsgBox("¿Desea registrar otra Penalidad EcoTaxi?", vbYesNo + vbInformation, "Aviso") = vbYes Then
        btnCancelar_Click
    Else
        Unload Me
    End If
    
    Set oCapta = Nothing
    Set oBase = Nothing
    Set oDCred = Nothing
    Set oITF = Nothing
    
    Exit Sub
ErrorGrabar:
    If bTransac Then
        Call oBase.dRollbackTrans
        Set oBase = Nothing
    End If
    err.Raise err.Number, "Error En Proceso", err.Description
End Sub
Private Sub btnCancelar_Click()
    LimpiaPantalla
    HabilitarCtrlsBusqueda (True)
    Me.btnGrabar.Enabled = False
    Me.ActXCtaCod.SetFocus
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub CargarConstantes()
    Dim oConsSist As COMDConstSistema.NCOMConstSistema
    Set oConsSist = New COMDConstSistema.NCOMConstSistema
    
    fnPenalidadEcoTaxixAtrasoEntregaxDia = oConsSist.LeeConstSistema(gConstSistMontoPenalidadEcoTaxixAtrasoEntregaxDia)
    fnPenalidadEcoTaxixVehiculoDistintoAcordado = oConsSist.LeeConstSistema(gConstSistMontoPenalidadEcoTaxixVehiculoDistintoAcordado)
End Sub
Private Sub ActXCtaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatos(ActXCtaCod.NroCuenta) Then
           MsgBox "No se encontraron datos con este Nro. de Cuenta", vbInformation, "Aviso"
        Else
            HabilitarCtrlsBusqueda (False)
            ddlMotivoPenalidad.SetFocus
            Me.btnGrabar.Enabled = True
        End If
    End If
End Sub
Private Sub ddlMotivoPenalidad_Click()
    If Me.ddlMotivoPenalidad.ListIndex <> -1 Then
        Select Case CInt(Trim(Right(ddlMotivoPenalidad.Text, 5)))
            Case 1
                Me.fraRetraso.Visible = True
                Me.txtMontoxDia.Caption = Format(fnPenalidadEcoTaxixAtrasoEntregaxDia, "##,##0.00")
                Me.txtDiasRetraso.Text = "0"
                Me.txtMontoPenalidad.Caption = Format(CDbl(txtMontoxDia.Caption) * CDbl(Me.txtDiasRetraso.Text), "##,##0.00")
                Me.txtMontoFavorCliente.Caption = Me.txtMontoPenalidad.Caption
                Me.txtMontoFavorInstitucion.Caption = "0.00"
            Case 2
                Me.fraRetraso.Visible = False
                Me.txtDiasRetraso.Text = "0"
                Me.txtMontoxDia.Caption = "0.00"
                
                Me.txtMontoPenalidad.Caption = Format(fnPenalidadEcoTaxixVehiculoDistintoAcordado, "##,##0.00")
                Me.txtMontoFavorCliente.Caption = "0.00"
                Me.txtMontoFavorInstitucion.Caption = Me.txtMontoPenalidad.Caption
        End Select
    End If
End Sub
Private Sub ddlMotivoPenalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case CInt(Trim(Right(Me.ddlMotivoPenalidad, 10)))
            Case 1
                Me.txtDiasRetraso.SetFocus
            Case 2
                Me.txtcomentarios.SetFocus
        End Select
    End If
End Sub
Private Sub txtDiasRetraso_Change()
    Me.txtMontoPenalidad.Caption = Format(CDbl(IIf(txtMontoxDia.Caption = "", "0", txtMontoxDia.Caption)) * CDbl(IIf(Me.txtDiasRetraso.Text = "", 0, Me.txtDiasRetraso.Text)), "##,##0.00")
    Me.txtMontoFavorCliente.Caption = Me.txtMontoPenalidad.Caption
End Sub
Private Sub txtDiasRetraso_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        Me.txtcomentarios.SetFocus
    End If
End Sub
Private Function validaDatos() As Boolean
    If Me.ddlMotivoPenalidad.ListIndex = -1 Then
        validaDatos = False
        MsgBox "Ud. debe seleccionar el Motivo de la Penalidad", vbInformation, "Aviso"
        Me.ddlMotivoPenalidad.SetFocus
        Exit Function
    End If
    If CDbl(Me.txtMontoPenalidad.Caption) = 0 Then
        validaDatos = False
        MsgBox "El Monto de la Penalidad debe ser mayor a cero", vbInformation, "Aviso"
        Exit Function
    End If
    If CInt(Trim(Right(ddlMotivoPenalidad.Text, 5))) = 1 And CInt(Me.txtDiasRetraso.Text) = 0 Then
        validaDatos = False
        MsgBox "Ingrese los dias de retraso de la entrega del Vehículo", vbInformation, "Aviso"
        If Me.txtDiasRetraso.Enabled Then Me.txtDiasRetraso.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(ddlMotivoPenalidad.Text, 5))) = 1 And CInt(Me.txtMontoFavorCliente.Caption) = 0 Then
        validaDatos = False
        MsgBox "Monto a favor del Cliente incorrecto, debe ser mayor a cero", vbInformation, "Aviso"
        If Me.txtDiasRetraso.Enabled Then Me.txtDiasRetraso.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(ddlMotivoPenalidad.Text, 5))) = 2 And CInt(Me.txtMontoFavorInstitucion.Caption) = 0 Then
        validaDatos = False
        MsgBox "Monto a favor de la Institución incorrecto, debe ser mayor a cero", vbInformation, "Aviso"
        Exit Function
    End If
    If Len(Trim(Me.txtcomentarios.Text)) = 0 Then
        validaDatos = False
        MsgBox "Ud. debe ingresar un comentario sobre la presente Penalidad", vbInformation, "Aviso"
        Me.txtcomentarios.SetFocus
        Exit Function
    End If
    If Len(Trim(Me.txtConcesionarioNom.Caption)) = 0 Then
        validaDatos = False
        MsgBox "El Crédito EcoTaxi NO tiene configurado al Concesionario", vbInformation, "Aviso"
        Exit Function
    End If
    If Len(Trim(Me.txtConcesionarioCtaCod.Caption)) <> 18 Then
        validaDatos = False
        MsgBox "El Crédito EcoTaxi NO tiene configurado la Cuenta de Ahorro del Concesionario", vbInformation, "Aviso"
        Exit Function
    End If
    If Len(Trim(Me.txtCtaCodRecaudo.Caption)) <> 18 Then
        validaDatos = False
        MsgBox "El Crédito EcoTaxi NO tiene configurado la Cuenta de Recaudo", vbInformation, "Aviso"
        Exit Function
    End If
    validaDatos = True
End Function
Private Sub HabilitarCtrlsBusqueda(ByVal pbHabilita As Boolean)
    fraDatosCredito.Enabled = pbHabilita
End Sub
Private Function DameMontoITF(ByVal pnMonto As Double, ByRef pbITFRedondeo As Boolean, ByRef pnRedondeoITF As Double) As Double
    Dim oITF As COMDConstSistema.FCOMITF
    Dim lnMonto As Double, lnITF As Double, lnRedondeoITF As Double
    
    Set oITF = New COMDConstSistema.FCOMITF
    oITF.fgITFParametros

    lnMonto = pnMonto
    lnITF = lnMonto * oITF.gnITFPorcent
    lnITF = oITF.CortaDosITF(lnITF)
    lnRedondeoITF = fgDiferenciaRedondeoITF(lnITF)
    If lnRedondeoITF > 0 Then
        lnITF = lnITF - lnRedondeoITF
    End If
       
    pbITFRedondeo = False
    If lnITF + lnRedondeoITF > 0 Then
        pbITFRedondeo = True
    End If
    
    pnRedondeoITF = lnRedondeoITF
    DameMontoITF = lnITF
    Set oITF = Nothing
End Function

