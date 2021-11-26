VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaGenHabilitacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4635
   ClientLeft      =   1635
   ClientTop       =   2640
   ClientWidth     =   7965
   Icon            =   "frmCajaGenHabilitacion.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEfectivo 
      Caption         =   "&Efectivo"
      Height          =   375
      Left            =   75
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6570
      TabIndex        =   15
      Top             =   4200
      Width           =   1320
   End
   Begin TabDlg.SSTab StabHab 
      Height          =   4080
      Left            =   15
      TabIndex        =   7
      Top             =   15
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   7197
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Habilitaciones"
      TabPicture(0)   =   "frmCajaGenHabilitacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Transporte"
      TabPicture(1)   =   "frmCajaGenHabilitacion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   3510
         Left            =   135
         TabIndex        =   24
         Top             =   420
         Width           =   7605
         Begin VB.TextBox txtMovDesc 
            Height          =   780
            Left            =   90
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2565
            Width           =   4455
         End
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   5715
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   3000
            Width           =   1740
         End
         Begin VB.Frame frameOrigen 
            Caption         =   "Origen"
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
            Height          =   990
            Left            =   75
            TabIndex        =   30
            Top             =   525
            Width           =   7350
            Begin Sicmact.TxtBuscar txtCtaContOrigen 
               Height          =   330
               Left            =   1410
               TabIndex        =   1
               Top             =   195
               Width           =   1485
               _ExtentX        =   2619
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
            Begin Sicmact.TxtBuscar txtAreaAgeOrig 
               Height          =   330
               Left            =   1410
               TabIndex        =   2
               Top             =   600
               Width           =   1485
               _ExtentX        =   2619
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
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta Origen:"
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   270
               Width           =   1065
            End
            Begin VB.Label Label1 
               Caption         =   "Area Agencia :"
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   630
               Width           =   1215
            End
            Begin VB.Label lblCtaContOrigDesc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2940
               TabIndex        =   32
               Top             =   210
               Width           =   4275
            End
            Begin VB.Label lblAreaAgeOrig 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2925
               TabIndex        =   31
               Top             =   600
               Width           =   4275
            End
         End
         Begin VB.Frame frameDestino 
            Caption         =   "Destino"
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
            Height          =   975
            Left            =   75
            TabIndex        =   25
            Top             =   1515
            Width           =   7380
            Begin Sicmact.TxtBuscar txtCtaContDest 
               Height          =   330
               Left            =   1380
               TabIndex        =   3
               Top             =   195
               Width           =   1485
               _ExtentX        =   2619
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
            Begin Sicmact.TxtBuscar TxtAreaAgeDest 
               Height          =   330
               Left            =   1380
               TabIndex        =   4
               Top             =   585
               Width           =   1485
               _ExtentX        =   2619
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
            Begin VB.Label Label6 
               Caption         =   "Area - Agencia"
               Height          =   195
               Left            =   105
               TabIndex        =   29
               Top             =   615
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cuenta Destino :"
               Height          =   255
               Left            =   105
               TabIndex        =   28
               Top             =   255
               Width           =   1215
            End
            Begin VB.Label lblCtacontDestDesc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2910
               TabIndex        =   27
               Top             =   195
               Width           =   4275
            End
            Begin VB.Label lblAreaAgeDest 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2895
               TabIndex        =   26
               Top             =   585
               Width           =   4275
            End
         End
         Begin MSMask.MaskEdBox txtFechaMov 
            Height          =   330
            Left            =   6210
            TabIndex        =   0
            Top             =   180
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblTitulo 
            Alignment       =   2  'Center
            Caption         =   "HABILITACION ENTRE AGENCIAS "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   360
            Left            =   315
            TabIndex        =   37
            Top             =   135
            Width           =   5190
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4995
            TabIndex        =   36
            Top             =   3075
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   5595
            TabIndex        =   35
            Top             =   225
            Width           =   540
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   390
            Left            =   4875
            Top             =   2985
            Width           =   2610
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos principales"
         Height          =   2970
         Left            =   -74700
         TabIndex        =   17
         Top             =   735
         Width           =   7305
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1125
            Width           =   4545
         End
         Begin VB.TextBox txtNetoServ 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5265
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   1605
            Width           =   1590
         End
         Begin VB.TextBox txtMontoServ 
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
            Height          =   345
            Left            =   5265
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   2490
            Width           =   1590
         End
         Begin VB.TextBox txtNroComp 
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   5325
            MaxLength       =   15
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cboTipoTrans 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   720
            Width           =   2790
         End
         Begin VB.Line Line1 
            X1              =   3840
            X2              =   7155
            Y1              =   2385
            Y2              =   2385
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Neto Servicio :"
            Height          =   210
            Left            =   3960
            TabIndex        =   23
            Top             =   1665
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Impuesto :"
            Height          =   210
            Left            =   3960
            TabIndex        =   22
            Top             =   2055
            Width           =   735
         End
         Begin VB.Label lblImpuesto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5265
            TabIndex        =   12
            Top             =   1995
            Width           =   1590
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Servicio :"
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
            Left            =   3975
            TabIndex        =   21
            Top             =   2565
            Width           =   1200
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "N° :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   4905
            TabIndex        =   20
            Top             =   285
            Width           =   330
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Empresa :"
            Height          =   210
            Left            =   120
            TabIndex        =   19
            Top             =   1170
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tipo :"
            Height          =   210
            Left            =   120
            TabIndex        =   18
            Top             =   735
            Width           =   390
         End
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5235
      TabIndex        =   14
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "frmCajaGenHabilitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oOpe As DOperacion
Dim lsAreCodCaja As String
Dim lbSalir As Boolean
Dim objPista As COMManejador.Pista
Dim oTransp As NCajaGenTransp

Private Sub cboEmpresa_Click()
txtNetoServ = Format(oTransp.GetMontosServicio(Right(cboEmpresa, 4), CCur(txtMonto), NetoServicio, Mid(gsOpeCod, 3, 1), gnTipCambio), "#,#0.00")
lblImpuesto = Format(oTransp.GetMontosServicio(Right(cboEmpresa, 4), CCur(txtMonto), IGV, Mid(gsOpeCod, 3, 1), gnTipCambio), "#,#0.00")
txtMontoServ = Format(oTransp.GetMontosServicio(Right(cboEmpresa, 4), CCur(txtMonto), MontoServicio, Mid(gsOpeCod, 3, 1), gnTipCambio), "#,#0.00")
If Val(txtNetoServ) = 0 And Right(cboTipoTrans, 1) <> CGTipoTransportePropio Then
    txtNetoServ.Locked = False
Else
    txtNetoServ.Locked = True
End If
End Sub

Private Sub cboEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNetoServ.SetFocus
End If
End Sub

Private Sub cboTipoTrans_Click()
If cboTipoTrans = "" Then Exit Sub
    CargaCombo cboEmpresa, oTransp.GetTranspTipo(Right(cboTipoTrans, 1))
    Select Case Right(cboTipoTrans, 1)
        Case CGTipoTransportePropio, CGTipoTransporteAlquilado
            txtNroComp = ""
            txtNroComp.Enabled = False
        Case Else
            txtNroComp.Enabled = True
    End Select
    If cboEmpresa.ListCount > 0 Then
        cboEmpresa.ListIndex = 0
    End If
End Sub
Private Sub cboTipoTrans_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboEmpresa.SetFocus
End If
End Sub

Private Sub cmdAceptar_Click()
Dim oCaja As nCajaGeneral
Dim oCon As NContFunciones
Dim lnSaldoDisp As Currency

Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset
Dim lsTipoDocTransp As TpoDoc
Dim lsTipoTransp As CGTipoTransporte
Dim lsNroCompTransp As String
Dim lsCodTransp As String
Dim lsDesTransp As String
Dim lnTotalServicio As Currency
Dim lsMovNro As String
Dim lsNroCompSal As String
Dim lsSubCta  As String

Set rsBill = New ADODB.Recordset
Set rsMon = New ADODB.Recordset
Set oCaja = New nCajaGeneral
Set oCon = New NContFunciones
On Error GoTo ErrHabilita

If Valida = False Then Exit Sub
'If ValidaTransp = False Then Exit Sub

lsTipoDocTransp = -1: lsTipoTransp = -1
lsNroCompTransp = "": lsCodTransp = ""
lsDesTransp = "": lnTotalServicio = 0
Select Case gsOpeCod
    Case gOpeBoveAgeHabEntreAgeMN, gOpeBoveAgeHabEntreAgeME
        If TxtAreaAgeDest = txtAreaAgeOrig Then
            MsgBox "Agencia de Destino no puede ser la misma que de Origen", vbInformation, "Aviso"
            If TxtAreaAgeDest.Enabled Then TxtAreaAgeDest.SetFocus
            Exit Sub
        End If
    Case Else
End Select
'lsTipoDocTransp = TpoDocCompServTranspValores
'lsTipoTransp = Val(Right(cboTipoTrans, 1))
'lsNroCompTransp = Trim(txtNroComp)
'lsCodTransp = Trim(Right(cboEmpresa, 4))
'lsDesTransp = Trim(Left(cboEmpresa, 50))
'lnTotalServicio = CCur(txtMontoServ)

Select Case gsOpeCod
    Case gOpeBoveCGHabAgeMN, gOpeBoveCGHabAgeME
        lnSaldoDisp = oCon.GetSaldoCtaCont(txtFechaMov, txtCtaContOrigen)
        If CCur(txtMonto) > lnSaldoDisp Then
            MsgBox "No posee saldo suficiente para realizar la operación" & vbCrLf & "Saldo Disponible : " & Format(lnSaldoDisp, "#,#0.00"), vbInformation, "Aviso"
            Exit Sub
        End If
    Case gOpeBoveAgeHabAgeACGMN, gOpeBoveAgeHabAgeACGME, gOpeBoveAgeHabCajeroMN, gOpeBoveAgeHabCajeroME
        'MsgBox "Todavia falta definir el cuadre de las operaciones de Agencia", vbCritical
        'Exit Sub
End Select

frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, CCur(txtMonto), Mid(gsOpeCod, 3, 1), False
If frmCajaGenEfectivo.lbOk Then
     Set rsBill = frmCajaGenEfectivo.rsBilletes
     Set rsMon = frmCajaGenEfectivo.rsMonedas
Else
    Set frmCajaGenEfectivo = Nothing
    Exit Sub
End If
Set frmCajaGenEfectivo = Nothing
If rsBill Is Nothing And rsMon Is Nothing Then
    MsgBox "Error en Ingreso de Billetaje", vbInformation, "Aviso"
    Exit Sub
End If

If MsgBox("¿ Desea Grabar la Habilitación respectiva ?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(txtFechaMov, gsCodAge, gsCodUser)
    If oCaja.GrabaHabEfectivo(lsMovNro, rsBill, rsMon, gsOpeCod, txtMovDesc, _
                              txtCtaContDest, txtCtaContOrigen, CCur(txtMonto), Trim(txtAreaAgeOrig), _
                              Trim(TxtAreaAgeDest), lsNroCompTransp, Format(txtFechaMov, gsFormatoFecha), lsNroCompSal, Format(txtFechaMov, gsFormatoFecha), _
                              lsCodTransp, lnTotalServicio, , lsAreCodCaja) = 0 Then

'If oCaja.GrabaHabEfectivoNew(lsMovNro, rsBill, rsMon, gsOpeCod, txtMovDesc, _
'                              txtCtaContDest, txtCtaContOrigen, CCur(txtMonto), Trim(txtAreaAgeOrig), _
'                              Trim(TxtAreaAgeDest), lsNroCompTransp, Format(txtFechaMov, gsFormatoFecha), lsNroCompSal, Format(txtFechaMov, gsFormatoFecha), _
'                              lsCodTransp, lnTotalServicio, , lsAreCodCaja) = 0 Then
        
        Select Case gsOpeCod
            Case gOpeBoveCGHabAgeMN, gOpeBoveCGHabAgeME
                ImprimeAsientoContable lsMovNro, , , , True, False, Me.txtMovDesc, , , , , True, 3
            Case gOpeBoveCGConfHabAgeBoveMN, gOpeBoveCGConfHabAgeBoveME
                ImprimeAsientoContable lsMovNro, , , , True, True, txtMovDesc, , , , , True, 2
            Case Else
                Dim oContImp As NContImprimir
                Dim lsTexto As String
                Set oContImp = New NContImprimir
                
                lsTexto = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, lsMovNro, gsNomCmac, 2)
                EnviaPrevio lsTexto, Me.Caption, gnLinPage, False
                Set oContImp = Nothing
                'ImprimeAsientoContable lsMovNro, , , , True, False, Me.txtMovdesc, , , , , True
                
        End Select
        Set oCaja = Nothing
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Habilitacion de Boveda a Agencias"
        If MsgBox("Desea Registrar otra habilitación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            txtMovDesc = ""
            txtMonto = "0.00"
            lsTipoDocTransp = -1
            lsTipoTransp = -1
            lsNroCompTransp = ""
            lsCodTransp = ""
            lsDesTransp = ""
            lnTotalServicio = 0
            txtMontoServ = "0.00"
            txtNroComp = ""
            cboEmpresa.ListIndex = -1
            cboTipoTrans.ListIndex = -1
            lblImpuesto = "0.00"
            txtNetoServ = "0.00"
            StabHab.Tab = 0
            If txtFechaMov.Enabled Then txtFechaMov.SetFocus
        Else
            Unload Me
        End If
    End If
End If
Exit Sub
ErrHabilita:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub
Function Valida() As Boolean
Valida = False
If txtCtaContOrigen = "" Then
    MsgBox "Cuenta de Origen no Seleccionada", vbInformation, "Aviso"
    If txtCtaContOrigen.Enabled Then txtCtaContOrigen.SetFocus
    Valida = False
    StabHab.Tab = 0
    Exit Function
End If
If txtAreaAgeOrig = "" Then
    MsgBox "Area o Agencia de Origen no Seleccionada", vbInformation, "Aviso"
    If txtAreaAgeOrig.Enabled Then txtAreaAgeOrig.SetFocus
    Valida = False
    StabHab.Tab = 0
    Exit Function
End If
If Me.txtCtaContDest = "" Then
    MsgBox "Cuenta de Destino no Seleccionada", vbInformation, "Aviso"
    If txtCtaContDest.Enabled Then txtCtaContDest.SetFocus
    Valida = False
    StabHab.Tab = 0
    Exit Function
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción del Movimiento no Ingresado", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Valida = False
    StabHab.Tab = 0
    Exit Function
End If
If Val(txtMonto) = 0 Then
    MsgBox "Ingrese Monto de Operación", vbInformation, "Aviso"
    txtMonto.SetFocus
    Valida = False
    StabHab.Tab = 0
    Exit Function
End If
If Not ValidaFechaContab(Me.txtFechaMov, gdFecSis) Then
    txtFechaMov.SetFocus
    Exit Function
End If
Valida = True
End Function

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim oGen As DGeneral
Set oGen = New DGeneral
StabHab.Tab = 0
CentraForm Me
lbSalir = False
txtFechaMov = gdFecSis
Me.Caption = gsOpeDesc
Set oOpe = New DOperacion
Set rs = New ADODB.Recordset
Set oTransp = New NCajaGenTransp
Set objPista = New COMManejador.Pista

CambiaTamañoCombo Me.cboTipoTrans
CargaCombo cboTipoTrans, oGen.GetConstante(gCGTipoTransporte)

Select Case gsOpeCod
    Case gOpeBoveCGHabAgeMN, gOpeBoveCGHabAgeME
        lbltitulo = "HABILITACION A AGENCIAS"
        txtCtaContOrigen.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "H")
        txtCtaContDest.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")
        txtAreaAgeOrig.rs = GetObjetosOpeCta(gsOpeCod, "1", txtCtaContOrigen, "")
        TxtAreaAgeDest.rs = GetObjetosOpeCta(gsOpeCod, "2", txtCtaContDest, "")
        lsAreCodCaja = ""
        StabHab.TabVisible(1) = False
        
    'Solo hasta que este el Centralizado de Negocio
    Case gOpeBoveCGConfHabAgeBoveMN, gOpeBoveCGConfHabAgeBoveME
        lbltitulo = "HABILITACION DE AGENCIAS"
        txtCtaContOrigen.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "H")
        txtCtaContDest.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")
        txtAreaAgeOrig.rs = GetObjetosOpeCta(gsOpeCod, "1", txtCtaContOrigen, "")
        TxtAreaAgeDest.rs = GetObjetosOpeCta(gsOpeCod, "0", txtCtaContDest, "")
        lsAreCodCaja = ""
        StabHab.TabVisible(1) = False
    '-------------------------------
    
    Case gOpeBoveAgeHabAgeACGMN, gOpeBoveAgeHabAgeACGME
        lbltitulo = "HABILITACION A CAJA GENERAL"
        txtCtaContOrigen.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "H")
        txtCtaContDest.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")
        Set rs = oOpe.GetOpeObj(gsOpeCod, "0")
        If Not rs.EOF And Not rs.EOF Then
            lsAreCodCaja = rs!Codigo
        Else
            MsgBox "No se ha definido el Objeto para Caja General", vbInformation, "Aviso"
            lbSalir = True
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
        txtAreaAgeOrig.rs = GetObjetosOpeCta(gsOpeCod, "1", txtCtaContOrigen, "", , Right(gsCodAge, 2))
        TxtAreaAgeDest.rs = GetObjetosOpeCta(gsOpeCod, "2", txtCtaContDest, "", , Right(gsCodAge, 2))
    Case gOpeBoveAgeHabEntreAgeMN, gOpeBoveAgeHabEntreAgeME
        lbltitulo = "HABILITACION ENTRE AGENCIAS"
        txtCtaContOrigen.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "H")
        txtCtaContDest.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")
        txtAreaAgeOrig.rs = GetObjetosOpeCta(gsOpeCod, "1", txtCtaContOrigen, "", , gsCodAge)
        TxtAreaAgeDest.rs = GetObjetosOpeCta(gsOpeCod, "2", txtCtaContDest, "")
        lsAreCodCaja = ""
End Select
txtMonto.BackColor = IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, vbWhite, &HC0FFC0)
Set oGen = Nothing

End Sub
Private Sub TxtAreaAgeDest_EmiteDatos()
lblAreaAgeDest = TxtAreaAgeDest.psDescripcion
If lblAreaAgeDest <> "" Then
    If txtMovDesc.Visible Then
        txtMovDesc.SetFocus
    End If
End If
End Sub
Private Sub txtAreaAgeOrig_EmiteDatos()
lblAreaAgeOrig = txtAreaAgeOrig.psDescripcion
If lblAreaAgeOrig <> "" And txtAreaAgeOrig.Enabled Then
    txtMovDesc.SetFocus
End If
End Sub
Private Sub txtCtaContDest_EmiteDatos()
lblCtacontDestDesc = txtCtaContDest.psDescripcion
End Sub
Private Sub txtCtaContOrigen_EmiteDatos()
lblCtaContOrigDesc = txtCtaContOrigen.psDescripcion
End Sub

Private Sub txtFechamov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtCtaContOrigen.Enabled Then
        txtCtaContOrigen.SetFocus
    ElseIf txtAreaAgeOrig.Enabled Then
            txtAreaAgeOrig.SetFocus
        ElseIf txtCtaContDest.Enabled Then
            txtCtaContDest.SetFocus
            ElseIf TxtAreaAgeDest.Enabled Then
                    TxtAreaAgeDest.SetFocus
                ElseIf Me.txtMovDesc.Enabled Then
                    txtMovDesc.SetFocus
                ElseIf Me.txtMonto.Enabled Then
                        txtMonto.SetFocus
                End If
End If
End Sub

Private Sub txtmonto_GotFocus()
fEnfoque txtMonto
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 20, 2)
If KeyAscii = 13 Then
    CmdAceptar.SetFocus
End If
End Sub
Private Sub txtMonto_LostFocus()
 If txtMonto = "" Then txtMonto = 0
txtMonto = Format(txtMonto, "#,#0.00")
End Sub

Private Sub txtMontoServ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdAceptar.SetFocus
End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Me.txtMonto.SetFocus
End If
End Sub

Private Sub txtNetoServ_GotFocus()
fEnfoque txtNetoServ
End Sub

Private Sub txtNetoServ_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtNetoServ, KeyAscii, 15, 2)
If KeyAscii = 13 Then
    If txtNetoServ.Locked = False Then
        txtMontoServ = Format(txtNetoServ, "#,#0.00")
    End If
    txtMontoServ.SetFocus
End If
End Sub
Private Sub txtNetoServ_LostFocus()
    txtNetoServ = Format(txtNetoServ, "#,#0.00")
End Sub


Private Sub txtNroComp_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    cboTipoTrans.SetFocus
End If
End Sub
Function ValidaTransp() As Boolean
ValidaTransp = True
If cboTipoTrans = "" Then
    MsgBox "Tipo de Transporte no Seleccionado", vbInformation, "Aviso"
    ValidaTransp = False
    StabHab.Tab = 1
    cboTipoTrans.SetFocus
    Exit Function
End If
If Right(cboTipoTrans, 1) = CGTipoTransporteBlindado Then
    If Len(Trim(txtNroComp)) = 0 Then
        MsgBox "Nro de Comprobante de Transporte no Ingresado", vbInformation, "Aviso"
        ValidaTransp = False
        StabHab.Tab = 1
        txtNroComp.SetFocus
        Exit Function
    End If
Else
    If Len(Trim(txtNroComp)) = 0 And txtNroComp.Enabled Then
        If MsgBox("Nro de Comprobante no Ingresado" & vbCrLf & "Desea Continuar??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            ValidaTransp = False
            StabHab.Tab = 1
            If txtNroComp.Enabled Then txtNroComp.SetFocus
            Exit Function
        End If
    End If
End If
If cboEmpresa = "" Then
    MsgBox "Empresa de Transporte de Valores no Seleccionada", vbInformation, "Aviso"
    ValidaTransp = False
    cboEmpresa.SetFocus
    StabHab.Tab = 1
    Exit Function
End If
If Right(cboTipoTrans, 1) <> CGTipoTransportePropio Then
    If Val(txtNetoServ) = 0 Then
        MsgBox "Monto Neto de Servicio no Ingresado", vbInformation, "Aviso"
        ValidaTransp = False
        txtNetoServ.SetFocus
        StabHab.Tab = 1
        Exit Function
    End If
    If Val(txtMontoServ) = 0 Then
        MsgBox "Monto total de Servicio no ha sido Calculado", vbExclamation, "Aviso"
        ValidaTransp = False
        txtNetoServ.SetFocus
        StabHab.Tab = 1
        Exit Function
    End If
    If CCur(txtMontoServ) > CCur(txtMonto) Then
        MsgBox "Monto de Servicio no puede ser mayor que el Transportado", vbExclamation, "Aviso"
        ValidaTransp = False
        StabHab.Tab = 1
        txtNetoServ.SetFocus
        Exit Function
    End If
End If
End Function

