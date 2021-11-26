VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSegTarjDepPagoBanco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depósito por Comisión Seguro Tarjetas"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   Icon            =   "frmSegTarjDepPagoBanco.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      TabIndex        =   24
      Top             =   5520
      Width           =   1170
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   0
      Top             =   5520
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      TabIndex        =   2
      Top             =   5520
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   1
      Top             =   5520
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Destino"
      TabPicture(0)   =   "frmSegTarjDepPagoBanco.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraOperacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraCtaIF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraPeriodo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame FraPeriodo 
         Caption         =   " Periodo "
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
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   3855
         Begin VB.CommandButton cmdSeleccionar 
            Caption         =   "Seleccionar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2520
            TabIndex        =   19
            Top             =   240
            Width           =   1170
         End
         Begin VB.TextBox txtAño 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmSegTarjDepPagoBanco.frx":0326
            Left            =   120
            List            =   "frmSegTarjDepPagoBanco.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblMes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label lblAño 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1560
            TabIndex        =   20
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame FraCtaIF 
         Caption         =   " Cuenta Institución Financiera "
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
         TabIndex        =   11
         Top             =   1200
         Width           =   6255
         Begin Sicmact.TxtBuscar txtBuscarBanco 
            Height          =   330
            Left            =   120
            TabIndex        =   12
            Top             =   375
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   582
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
            sTitulo         =   ""
         End
         Begin VB.Label lblDescCtaBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            TabIndex        =   15
            Top             =   750
            Width           =   6015
         End
         Begin VB.Label lblDescBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2205
            TabIndex        =   14
            Top             =   375
            Width           =   3930
         End
         Begin VB.Label lblCtaBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            TabIndex        =   13
            Top             =   375
            Width           =   2010
         End
      End
      Begin VB.Frame FraOperacion 
         Caption         =   " Datos de la operación "
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
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   6255
         Begin Sicmact.EditMoney txtMonto 
            Height          =   300
            Left            =   4800
            TabIndex        =   23
            Top             =   1800
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMovDesc 
            Height          =   720
            Left            =   1080
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   840
            Width           =   5010
         End
         Begin VB.Label lblsMovNro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   5160
            TabIndex        =   26
            Top             =   360
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblnMovNro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4200
            TabIndex        =   25
            Top             =   360
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label labelMonto 
            Caption         =   "Monto :"
            Height          =   285
            Left            =   4200
            TabIndex        =   22
            Top             =   1840
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Glosa :"
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   705
         End
         Begin VB.Label LabelMontoCalc 
            Caption         =   "Monto Calculado:"
            Height          =   405
            Left            =   120
            TabIndex        =   8
            Top             =   1710
            Width           =   825
         End
         Begin VB.Label lblFecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1080
            TabIndex        =   7
            Top             =   345
            Width           =   1290
         End
         Begin VB.Label lblMontoCalc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            Top             =   1800
            Width           =   1290
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha :"
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   705
         End
      End
   End
End
Attribute VB_Name = "frmSegTarjDepPagoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmSegTarjDepPagoBanco
'** Descripción : Formulario para registrar el depósito por comisión de seguros de tarjetas y
'**               su extorno creado segun TI-ERS029-2013
'** Creación : JUEZ, 20140711 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim nOperacion As OpeCGOpeBancos
Dim oOpe As DOperacion
Dim oDSeg As DSeguros
Dim oNSeg As NSeguros

Dim rs As ADODB.Recordset
Dim lsCtaContBanco As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Public Sub Inicio(ByVal pnOperacion As OpeCGOpeBancos)
nOperacion = pnOperacion

If nOperacion = gOpeCGOpeBancosDepComSegTarjMN Then
    Me.Caption = "Depósito por Comisión Seguro Tarjetas"
    txtBuscarBanco.psRaiz = "Cuentas de Bancos"
    Set oOpe = New DOperacion
    txtBuscarBanco.rs = oOpe.GetOpeObj(pnOperacion, "2")
    Set oOpe = Nothing
    HabilitaControles (False)
    CargaComboConstante 1010, CboMes
    HabilitaControlesExtorno False
Else
    Me.Caption = "Extorno MN: Extorno Depósito por Comisión Seguro Tarjetas"
    Set oDSeg = New DSeguros
        Set rs = oDSeg.RecuperaSegTarjetaOpeBancosExtorno(Format(gdFecSis, "yyyymmdd"), 2)
    Set oDSeg = Nothing
    If rs.EOF Then
        MsgBox "No se encontraron operaciones de Depósito por Comisión de Seguro de Tarjetas para extornar", vbInformation, "Aviso"
        Exit Sub
    Else
        lblAño.Caption = rs!cAnio
        lblMes.Caption = rs!cMES
        lblCtaBanco.Caption = rs!cCtaBanco
        lblDescBanco.Caption = rs!cBancoDesc
        lblDescCtaBanco.Caption = rs!cCtaBancoDesc
        lblFecha.Caption = Format(rs!dFechaPago, "dd/mm/yyyy")
        txtMovDesc.Text = rs!cMovDesc
        txtMovDesc.Enabled = False
        lblMontoCalc.Caption = Format(rs!nMonto, "#0.00")
        LabelMontoCalc.Caption = "Monto:"
        LabelMontoCalc.Top = 1840
        lblsMovNro.Caption = rs!cMovNro
        lblnMovNro.Caption = rs!nMovNro
        labelMonto.Visible = False
        txtMonto.Visible = False
    End If
    HabilitaControlesExtorno True
End If

Me.Show 1
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
fraPeriodo.Enabled = Not pbHabilita
FraCtaIF.Enabled = pbHabilita
fraOperacion.Enabled = pbHabilita
cmdAceptar.Enabled = pbHabilita
End Sub

Private Sub HabilitaControlesExtorno(ByVal pbHabilitaExt As Boolean)
lblMes.Visible = pbHabilitaExt
lblAño.Visible = pbHabilitaExt
CboMes.Visible = Not pbHabilitaExt
txtAño.Visible = Not pbHabilitaExt
cmdSeleccionar.Visible = Not pbHabilitaExt
lblCtaBanco.Visible = pbHabilitaExt
txtBuscarBanco.Visible = Not pbHabilitaExt
cmdExtornar.Visible = pbHabilitaExt
cmdAceptar.Visible = Not pbHabilitaExt
cmdCerrar.Visible = pbHabilitaExt
cmdCancelar.Visible = Not pbHabilitaExt
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAño.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim oCtasIF As NCajaCtaIF
Dim oNSeg As NSeguros
Dim oImp As NContImprimir

Dim lsPersCodIf As String
Dim lsTpoIf As String
Dim lsCtaBanco As String
Dim lsSubCuentaIF As String
Dim lsPeriodo As String
Dim lsMovNro As String
Dim lsImpre As String

If Trim(txtMovDesc.Text) = "" Then
    MsgBox "Debe ingresar la glosa", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If

If CDbl(txtMonto.Text) = 0 Then
    If MsgBox("El monto de la operación es 0. ¿Desea continuar?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
End If

Set oCtasIF = New NCajaCtaIF

lsPersCodIf = Mid(txtBuscarBanco.Text, 4, 13)
lsTpoIf = Mid(txtBuscarBanco.Text, 1, 2)
lsCtaBanco = Mid(txtBuscarBanco.Text, 18, Len(txtBuscarBanco.Text))
lsSubCuentaIF = oCtasIF.SubCuentaIF(lsPersCodIf)

lsPeriodo = txtAño.Text + IIf(Len(Trim(Right(CboMes.Text, 2))) = 1, "0", "") & Trim(Right(CboMes.Text, 2))

If MsgBox("¿Esta seguro de realizar la operación?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub

Set oNSeg = New NSeguros
Call oNSeg.GrabarDepositoComisionSeguroTarjeta(gdFecSis, gsCodAge, gsCodUser, nOperacion, lsPeriodo, lblFecha.Caption, Trim(txtMovDesc.Text), _
                                               CDbl(txtMonto.Text), lsCtaContBanco, lsPersCodIf, lsTpoIf, lsCtaBanco, lsMovNro)
Set oNSeg = Nothing

MsgBox "La operación fue realizada", vbInformation, "Aviso"
Set oImp = New NContImprimir
    lsImpre = oImp.ImprimeAsientoContable(lsMovNro, gnLinPage, gnColPage, "DEPÓSITO POR COMISIÓN SEGURO DE TARJETAS")
    EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, "DEPÓSITO POR COMISIÓN SEGURO DE TARJETAS", gnLinPage, False
Set oImp = Nothing
        
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo la Operación "
        Set objPista = Nothing
        '****
Unload Me

End Sub

Private Sub cmdCancelar_Click()
    CboMes.ListIndex = -1
    txtAño.Text = ""
    txtBuscarBanco.Text = ""
    lblDescBanco.Caption = ""
    lblDescCtaBanco.Caption = ""
    lblFecha.Caption = ""
    txtMovDesc.Text = ""
    lblMontoCalc.Caption = ""
    txtMonto.Text = 0
    HabilitaControles (False)
    If nOperacion = gOpeCGOpeBancosDepComSegTarjMN Then CboMes.SetFocus
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdExtornar_Click()
Dim oNContFun As NContFunciones
Dim oNCaja As nCajaGeneral
Dim oImp As NContImprimir

Dim lsMovNroExt As String
Dim lbEliminaMov As Boolean
Dim lsImpre As String

On Error GoTo ExtornarErr
    
    Set oNContFun = New NContFunciones
    lsMovNroExt = oNContFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oNContFun = Nothing
    
    If MsgBox("Se va a extornar la operación, Desea continuar?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    
    Dim oFun As New NContFunciones
    lbEliminaMov = oFun.PermiteModificarAsiento(lblsMovNro.Caption, False)
    
    Set oNCaja = New nCajaGeneral
    If oNCaja.GrabaExtornoMov(gdFecSis, gdFecSis, lsMovNroExt, CLng(lblnMovNro.Caption), nOperacion, txtMovDesc.Text, CCur(lblMontoCalc.Caption), lsMovNroExt, lbEliminaMov, , , , , gbBitCentral) = 0 Then
        If Not lbEliminaMov Then
            Set oImp = New NContImprimir
                lsImpre = oImp.ImprimeAsientoContable(lsMovNroExt, gnLinPage, gnColPage, "EXTORNO DEPOSITO POR COMISION SEGURO DE TARJETAS")
                EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, "EXTORNO DEPOSITO POR COMISION SEGURO DE TARJETAS", gnLinPage, False
            Set oImp = Nothing
            
        End If
        MsgBox "El extorno fue realizado", vbInformation, "Aviso"
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo el Extorno "
        Set objPista = Nothing
        '****
        Unload Me
        Exit Sub
    End If
    cmdExtornar.Enabled = True

Exit Sub
ExtornarErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSeleccionar_Click()
If Not ValidaSeleccionar Then Exit Sub
Dim lnPorcComision As Double
Dim lnMonto As Double

    HabilitaControles (True)
    txtBuscarBanco.SetFocus
    lblFecha.Caption = gdFecSis
    Set oDSeg = New DSeguros
    Set rs = oDSeg.ObtenerSegTarjetaParametros(102)
    Set oDSeg = Nothing
    lnPorcComision = rs!nParamValor
    Set oDSeg = New DSeguros
    Set rs = oDSeg.RecuperaSegTarjetaOpeBancos(txtAño.Text, IIf(Len(Trim(Right(CboMes.Text, 2))) = 1, "0", "") & Trim(Right(CboMes.Text, 2)), 1)
    Set oDSeg = Nothing
    lnMonto = rs!nMonto * (lnPorcComision / 100)
    lblMontoCalc.Caption = Format(lnMonto, "#,##0.00")
    txtMonto.Text = Format(lnMonto, "#,##0.00")
End Sub

Private Function ValidaSeleccionar() As Boolean
ValidaSeleccionar = False
If Trim(txtAño.Text) = "" Then
    MsgBox "Ingrese correctamente el año", vbInformation, "Aviso"
    txtAño.SetFocus
    Exit Function
End If
If Val(txtAño.Text) < 1900 Or Val(txtAño.Text) > 9972 Then
    MsgBox "Ingrese correctamente el año", vbInformation, "Aviso"
    txtAño.SetFocus
    Exit Function
End If
If Trim(CboMes.Text) = "" Then
    MsgBox "Ingrese correctamente el mes", vbInformation, "Aviso"
    CboMes.SetFocus
    Exit Function
End If
'JUEZ 20150103 ************************************************************
Dim nAnioMesRep As Long
Dim nAnioMesSis As Long
nAnioMesRep = CLng(txtAño.Text & IIf(Len(Trim(Right(CboMes.Text, 2))) = 1, "0", "") & Trim(Right(CboMes.Text, 2)))
nAnioMesSis = CLng(Right(gdFecSis, 4) & Mid(gdFecSis, 4, 2))
'If CInt(Trim(Right(cboMes.Text, 2))) >= CInt(Mid(gdFecSis, 4, 2)) Or val(txtAño.Text) > Right(gdFecSis, 4) Then
If nAnioMesRep >= nAnioMesSis Then
'END JUEZ *****************************************************************
    MsgBox "El periodo seleccionado debe ser anterior al mes actual", vbInformation, "Aviso"
    Exit Function
End If

Set oDSeg = New DSeguros
    If Not oDSeg.RecuperaSegTarjetaOpeBancos(txtAño.Text, IIf(Len(Trim(Right(CboMes.Text, 2))) = 1, "0", "") & Trim(Right(CboMes.Text, 2)), 2).EOF Then
        MsgBox "El Depósito por Comisión de Seguro de Tarjeta para este periodo ya fue registrado", vbInformation, "Aviso"
        Exit Function
    End If
    If oDSeg.RecuperaSegTarjetaOpeBancos(txtAño.Text, IIf(Len(Trim(Right(CboMes.Text, 2))) = 1, "0", "") & Trim(Right(CboMes.Text, 2)), 1).EOF Then
        MsgBox "No se puede generar el Depósito por Comisión porque no se ha registrado el Retiro de Pago por Seguro de Tarjetas para el periodo seleccionado", vbInformation, "Aviso"
        Exit Function
    End If
Set oDSeg = Nothing

ValidaSeleccionar = True
End Function

Private Sub txtBuscarBanco_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Dim oOpe As DOperacion
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF

lblDescBanco = oCtaIf.NombreIF(Mid(txtBuscarBanco, 4, 13))
lblDescCtaBanco = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscarBanco, 18, 10)) + " " + txtBuscarBanco.psDescripcion
lsCtaContBanco = oOpe.EmiteOpeCta(nOperacion, "D", , txtBuscarBanco.Text, ObjEntidadesFinancieras)
    If lsCtaContBanco = "" Then
        MsgBox "Cuentas Contables no determinadas Correctamente" & Chr(13) & "consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
        txtBuscarBanco.Text = ""
        lblDescBanco.Caption = ""
        lblDescCtaBanco.Caption = ""
        Exit Sub
    End If
txtMovDesc.SetFocus
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        txtMonto.SetFocus
    End If
End Sub
