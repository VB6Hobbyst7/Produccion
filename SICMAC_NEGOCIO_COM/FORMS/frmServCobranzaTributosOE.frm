VERSION 5.00
Begin VB.Form frmServCobranzaTributosOE 
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   9
      Top             =   3945
      Width           =   1110
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3570
      TabIndex        =   8
      Top             =   3945
      Width           =   1110
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4770
      TabIndex        =   7
      Top             =   3945
      Width           =   1110
   End
   Begin VB.Frame fraseleccion 
      Height          =   1035
      Left            =   2130
      TabIndex        =   3
      Top             =   60
      Width           =   3765
      Begin VB.CommandButton cmdbuscar 
         Caption         =   "&Buscar"
         Height          =   360
         Left            =   2565
         TabIndex        =   5
         Top             =   435
         Width           =   1020
      End
      Begin VB.TextBox txtderechos 
         Height          =   360
         Left            =   180
         TabIndex        =   4
         Top             =   420
         Width           =   2145
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Tipo de Pago "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1020
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   1950
      Begin VB.OptionButton optBuscar 
         Caption         =   "Otros Pagos Varios"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   6
         Top             =   720
         Width           =   1680
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "Recibo(Tributos)"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   495
         Width           =   1635
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "Derechos"
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   225
         Width           =   1680
      End
   End
   Begin VB.Frame frmderecho 
      Caption         =   "Cobranza de Derechos(Tupa)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2115
      Left            =   165
      TabIndex        =   27
      Top             =   1185
      Width           =   5715
      Begin VB.TextBox lbltotalD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1875
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   660
         Width           =   1065
      End
      Begin VB.TextBox lblimpo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Left            =   1875
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox txtobserva 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1710
         TabIndex        =   35
         Top             =   1605
         Width           =   3390
      End
      Begin VB.TextBox txtNombres 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1710
         TabIndex        =   34
         Top             =   1155
         Width           =   3390
      End
      Begin VB.Label lblCom 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4590
         TabIndex        =   33
         Top             =   825
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblObservacion 
         Caption         =   "Observacion"
         Height          =   210
         Left            =   600
         TabIndex        =   32
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label lblnombre1 
         Caption         =   "Nombre"
         Height          =   240
         Left            =   600
         TabIndex        =   31
         Top             =   1245
         Width           =   600
      End
      Begin VB.Label lbltotal1 
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   735
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000017&
         DrawMode        =   6  'Mask Pen Not
         Index           =   3
         X1              =   1560
         X2              =   3030
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000017&
         DrawMode        =   6  'Mask Pen Not
         Index           =   2
         X1              =   1560
         X2              =   3030
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Label lblImporte 
         Caption         =   "Importe"
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   29
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lblcomision1 
         Caption         =   "Comision"
         Height          =   285
         Left            =   4710
         TabIndex        =   28
         Top             =   900
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame frmcobro 
      Caption         =   "Cobranza de Tributos/Recibos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2760
      Left            =   165
      TabIndex        =   10
      Top             =   1155
      Width           =   5700
      Begin VB.Frame FrmDeuda 
         BackColor       =   &H80000004&
         Caption         =   "Deuda "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2040
         Left            =   1200
         TabIndex        =   12
         Top             =   660
         Width           =   2595
         Begin VB.Label lblCostas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2940
            TabIndex        =   37
            Top             =   1380
            Width           =   345
         End
         Begin VB.Label lblgastadm 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2940
            TabIndex        =   36
            Top             =   990
            Width           =   330
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000017&
            DrawMode        =   6  'Mask Pen Not
            Index           =   1
            X1              =   930
            X2              =   2400
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label lblvalorgastos 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1185
            TabIndex        =   24
            Top             =   540
            Width           =   1125
         End
         Begin VB.Label lblvalorInt 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1185
            TabIndex        =   23
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label lblvalorajus 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1185
            TabIndex        =   22
            Top             =   870
            Width           =   1125
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000017&
            DrawMode        =   6  'Mask Pen Not
            Index           =   0
            X1              =   930
            X2              =   2400
            Y1              =   1530
            Y2              =   1530
         End
         Begin VB.Label lblTotal 
            Caption         =   "Total "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   330
            TabIndex        =   21
            Top             =   1650
            Width           =   585
         End
         Begin VB.Label lblvalorImp 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1185
            TabIndex        =   20
            Top             =   210
            Width           =   1125
         End
         Begin VB.Label lblvalorcom 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2610
            TabIndex        =   19
            Top             =   1380
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblvalorTot 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
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
            Height          =   285
            Left            =   1200
            TabIndex        =   18
            Top             =   1590
            Width           =   1125
         End
         Begin VB.Label lblComision 
            Caption         =   "Comision"
            Height          =   285
            Left            =   2745
            TabIndex        =   17
            Top             =   840
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lblImporte 
            Caption         =   "Importe"
            Height          =   285
            Index           =   0
            Left            =   285
            TabIndex        =   16
            Top             =   255
            Width           =   630
         End
         Begin VB.Label lblGastos 
            Caption         =   "Gastos  "
            Height          =   285
            Index           =   1
            Left            =   300
            TabIndex        =   15
            Top             =   555
            Width           =   750
         End
         Begin VB.Label lblAjustes 
            Caption         =   "Ajustes "
            Height          =   285
            Index           =   2
            Left            =   300
            TabIndex        =   14
            Top             =   870
            Width           =   735
         End
         Begin VB.Label lblImporte 
            Caption         =   "Int/Otros"
            Height          =   285
            Index           =   1
            Left            =   300
            TabIndex        =   13
            Top             =   1215
            Width           =   765
         End
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1695
         TabIndex        =   11
         Top             =   330
         Width           =   3480
      End
      Begin VB.Label txtcodcont 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   915
         TabIndex        =   26
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H80000004&
         Caption         =   "Nombre "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   135
         TabIndex        =   25
         Top             =   405
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmServCobranzaTributosOE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Function GetValorComision() As Double
Dim rsPar As Recordset

Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion

Set rsPar = oCap.GetTarifaParametro(COMDConstantes.gServCobSATTReciboDerecho, COMDConstantes.gMonedaNacional, COMDConstantes.gCostoComServSATRecibosDerechos)
Set oCap = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetValorComision = 0
Else
    GetValorComision = rsPar("nParValor")
End If
rsPar.Close
Set rsPar = Nothing
End Function

Private Sub cmdBuscar_Click()
Dim sDato As String
Dim ntipo As Integer
If optBuscar(0).value Then 'Derechos
    sDato = Trim(txtderechos.Text)
    sDato = Replace(sDato, "-", "", 1, , vbTextCompare)
    ntipo = 0
ElseIf optBuscar(1).value Then 'Recibos/Tributos
    sDato = Trim(txtderechos.Text)
  ntipo = 1
ElseIf optBuscar(2).value Then 'Otros Pagos
'    sDato = Trim(txtPapeleta.Text)
    ntipo = 2
End If
'Verifica que se enviee informacion
If sDato = "" Then
    MsgBox "Debe ingresar el dato a buscar", vbInformation, "Aviso"
    Exit Sub
End If


Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
Set oServ = New COMNCaptaServicios.NCOMCaptaServicios

Dim rsSat As Recordset
Dim Total As Double

Set rsSat = oServ.GetServSATTributos(sDato, ntipo)
Set oServ = Nothing
If rsSat.EOF And rsSat.BOF Then
    MsgBox "Dato NO Encontrado o NO Posee Infracciones", vbInformation, "Aviso"
Else
    If optBuscar(0).value Then 'Derechos
        Me.lblimpo.Text = Format$(rsSat("nvalDerecho"), "#,##0.00")
        lblCom.Caption = 0 'Format$(GetValorComision(), "#,##0.00")
        lbltotalD = CStr(Format$(CDbl(lblCom) + CDbl(lblimpo), "#,##0.00"))
        txtNombres.SetFocus
        Me.lblimpo.Enabled = True
        
    End If
    If optBuscar(1).value Then 'Recibos
        txtNombre.Text = rsSat("cnombre")
        txtcodcont.Caption = rsSat("ccontrib")
        lblvalorImp.Caption = Format$(rsSat("nvaldeuda"), "#,##0.00")
        lblvalorgastos.Caption = Format$(rsSat("nvalderemis"), "#,##0.00")
        lblvalorajus.Caption = Format$(rsSat("nvalajuste"), "#,##0.00")
        lblgastadm.Caption = Format$(rsSat("nValGastos"), "#,##0.00")
        lblCostas.Caption = Format$(rsSat("nValCostas"), "#,##0.00")
        lblvalorInt.Caption = CStr(Format$(CDbl(rsSat("nvalintmor")) + CDbl(rsSat("nvalgastos")) + CDbl(rsSat("nvalCostas")), "#,##0.00"))
        lblvalorTot = 0
        lblvalorTot = CStr(Format$(CDbl(lblvalorImp) + CDbl(lblvalorgastos) + CDbl(lblvalorajus) + _
                                                   CDbl(lblvalorInt), "#,##0.00"))
    End If
    'Validacion de controles desactivados
    cmdBuscar.Enabled = False
    fraBusqueda.Enabled = False
    cmdCancelar.Enabled = True
End If
rsSat.Close
Set rsSat = Nothing
End Sub

Private Sub cmdCancelar_Click()
 'Validacion de controles desactivados
    cmdBuscar.Enabled = True
    fraBusqueda.Enabled = True
    cmdCancelar.Enabled = False
    LimpiaControles
    txtderechos.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim sCuenta As String, sCliente As String
Dim nMonto As Double, nMontoComision As Double

Dim CLSSERV As COMNCaptaServicios.NCOMCaptaServicios
Set CLSSERV = New COMNCaptaServicios.NCOMCaptaServicios

If (IsNull(Me.lblvalorImp) Or Me.lblvalorImp = 0) And (IsNull(Me.lblimpo) Or Me.lblimpo = 0) Then
  If (IsNull(txtderechos.Text) Or txtderechos.Text = "") Then
      MsgBox "Falta Ingresar Valor del Recibo o Derecho", , "AVISO DE ERROR"
  Else
      MsgBox "Dato Mal Ingresados en el Recibo o Derecho", , "AVISO DE ERROR"
  End If
  txtderechos.SetFocus
  txtderechos.Text = ""
Else
   If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then

    Dim clsMov As COMNContabilidad.NCOMContFunciones
    
    Dim sMovNro As String
    Dim L As MSComctlLib.ListItem
    On Error GoTo ErrGraba

        Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set clsMov = Nothing
         If Me.optBuscar(0).value = True Then 'Derecho
            sCuenta = Trim(txtderechos.Text)
            sCliente = Trim(txtNombres.Text)
            nMonto = lblimpo.Text
            nMontoComision = 0 'lblCom.Caption
            CLSSERV.CapCobranzaServicios sMovNro, COMDConstantes.gServCobSATTReciboDerechoOficEsp, sCuenta, sCliente, nMonto, _
                              gsNomCmac, gsNomAge, sLpt, nMontoComision, COMDConstantes.gMonedaNacional, 0, 0, 0, gsCodCMAC
        ElseIf optBuscar(1).value = True Then 'Recibo
            sCuenta = Trim(txtderechos.Text)
            sCliente = txtcodcont.Caption
            nMonto = CDbl(lblvalorImp) + CDbl(lblvalorajus) + CDbl(lblvalorInt)
            nMontoComision = CDbl(lblvalorcom.Caption)
            CLSSERV.CapCobranzaServicios sMovNro, COMDConstantes.gServCobSATTReciboDerecho, sCuenta, sCliente, nMonto, _
                              gsNomCmac, gsNomAge, sLpt, nMontoComision, , lblgastadm, lblCostas, lblvalorgastos, gsCodCMAC
        ElseIf optBuscar(2).value = True Then 'Pagos Varios
        End If
     LimpiaControles
     txtderechos.SetFocus
    End If
  End If
Set CLSSERV = Nothing
Exit Sub
ErrGraba:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
optBuscar(0).value = True
LimpiaControles
ValorIncial
Me.Caption = "Pagos de Servicios Varios -  OFICINAS ESPECIALES"
End Sub

Private Sub ValorIncial()
    If Me.optBuscar(0).value = True Then 'Derechos
        lblCom = 0 'Format$(GetValorComision(), "#,##0.00")
        lbltotalD = lblCom
    ElseIf Me.optBuscar(1).value = True Then 'Tributos
        lblvalorcom = 0 'Format$(GetValorComision(), "#,##0.00")
        lblvalorTot = lblvalorcom
    ElseIf Me.optBuscar(2).value = True Then 'Otros Pagos Varios
    End If
End Sub

Private Sub LimpiaControles()
txtderechos.Text = ""
If Me.optBuscar(0).value = True Then  'Derechos
    lblimpo.Text = "0.00"
    lblCom.Caption = "0.00"
    lbltotalD.Text = "0.00"
    txtNombres.Text = ""
    txtobserva.Text = ""
    cmdBuscar.Enabled = True
    'txtderechos.SetFocus
ElseIf Me.optBuscar(1).value = True Then  'Recibos
  txtcodcont.Caption = ""
  txtNombre.Text = ""
  lblvalorImp.Caption = "0.00"
  lblvalorgastos.Caption = "0.00"
  lblvalorajus.Caption = "0.00"
  lblvalorInt.Caption = "0.00"
  lblvalorcom.Caption = "0.00"
  lblvalorTot.Caption = "0.00"
  cmdBuscar.Enabled = True
    txtderechos.SetFocus
ElseIf Me.optBuscar(2).value = True Then 'Pagos Varios
End If
End Sub
Private Sub lblimpo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    lbltotalD.Text = Format$(lblimpo.Text, "#,##0.00")
    lblimpo.Text = Format$(lblimpo.Text, "#,##0.00")
    txtNombres.SetFocus
End If
End Sub

Private Sub optBuscar_Click(Index As Integer)
If optBuscar(0).value = True Then
   LimpiaControles
   Me.frmcobro.Visible = False
   Me.frmderecho.Visible = True
ElseIf optBuscar(1).value = True Then
    LimpiaControles
    Me.frmcobro.Visible = True
    Me.frmderecho.Visible = False
End If
ValorIncial
End Sub
Private Sub txtderechos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
End If
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtobserva.SetFocus
End If
End Sub
Private Sub txtobserva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub
