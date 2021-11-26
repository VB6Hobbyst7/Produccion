VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmColPDesemb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolso Pignoraticio"
   ClientHeight    =   8385
   ClientLeft      =   1440
   ClientTop       =   1980
   ClientWidth     =   8175
   HelpContextID   =   210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTabDatos 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Desembolsos"
      TabPicture(0)   =   "frmColPDesemb.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCampRetenPrend(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdGrabar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSalir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fr_CampPrendReferido"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Cuentas de ahorro"
      TabPicture(1)   =   "frmColPDesemb.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      Begin VB.Frame fr_CampPrendReferido 
         Caption         =   "Referidos"
         Height          =   735
         Left            =   120
         TabIndex        =   39
         Top             =   6840
         Width           =   7695
         Begin VB.TextBox txtCampaPrendTrabajador 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   42
            Top             =   240
            Width           =   5295
         End
         Begin VB.TextBox txtCampaPrendUser 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdCampPrendBuscar 
            Caption         =   "..."
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtCamPredPersCodReferido 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6480
            TabIndex        =   43
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6675
         Left            =   -74760
         TabIndex        =   30
         Top             =   360
         Width           =   7560
         Begin VB.CommandButton CmdAperturar 
            Caption         =   "&Aperturar"
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
            Height          =   420
            Left            =   240
            TabIndex        =   33
            Top             =   3360
            Width           =   1440
         End
         Begin TabDlg.SSTab SSTCtasAho 
            Height          =   2655
            Left            =   120
            TabIndex        =   34
            Top             =   495
            Width           =   7320
            _ExtentX        =   12912
            _ExtentY        =   4683
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Cuentas"
            TabPicture(0)   =   "frmColPDesemb.frx":0038
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "FECtaAhoDesemb"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Clientes"
            TabPicture(1)   =   "frmColPDesemb.frx":0054
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FEPersCtaAho"
            Tab(1).ControlCount=   1
            Begin SICMACT.FlexEdit FEPersCtaAho 
               Height          =   2055
               Left            =   -74880
               TabIndex        =   35
               Top             =   480
               Width           =   7080
               _ExtentX        =   12488
               _ExtentY        =   3625
               Cols0           =   3
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "-Cliente-Relacion"
               EncabezadosAnchos=   "350-5000-1200"
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
               ColumnasAEditar =   "X-X-X"
               TextStyleFixed  =   4
               ListaControles  =   "0-0-0"
               BackColor       =   13816486
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-C"
               FormatosEdit    =   "0-0-0"
               lbUltimaInstancia=   -1  'True
               lbBuscaDuplicadoText=   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
               CellBackColor   =   13816486
            End
            Begin SICMACT.FlexEdit FECtaAhoDesemb 
               Height          =   1935
               Left            =   120
               TabIndex        =   36
               Top             =   480
               Width           =   7095
               _ExtentX        =   12330
               _ExtentY        =   3413
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "-Cuenta-Agencia-ITF"
               EncabezadosAnchos=   "350-2000-3000-1500"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
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
               TextStyleFixed  =   4
               ListaControles  =   "0-0-0-0"
               BackColor       =   16777215
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-C-C"
               FormatosEdit    =   "0-0-0-0"
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
               CellBackColor   =   16777215
            End
         End
         Begin VB.CommandButton CmdSeleccionar 
            Caption         =   "&Seleccionar"
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
            Height          =   420
            Left            =   1680
            TabIndex        =   31
            Top             =   3360
            Width           =   1440
         End
         Begin VB.CommandButton CmdDeseleccionar 
            Caption         =   "&Deseleccionar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1680
            TabIndex        =   32
            Top             =   3360
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label21 
            Caption         =   "ABONO A CUENTA :"
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
            Left            =   1425
            TabIndex        =   38
            Top             =   225
            Width           =   1875
         End
         Begin VB.Label LblTipoAbono 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   3300
            TabIndex        =   37
            Top             =   225
            Width           =   1290
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   29
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6720
         TabIndex        =   28
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   27
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Frame lblCampRetenPrend 
         Height          =   6420
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7725
         Begin VB.CommandButton cmdBuscar 
            Height          =   345
            Left            =   7110
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Buscar ..."
            Top             =   180
            Width           =   420
         End
         Begin VB.Frame fraContenedor 
            Enabled         =   0   'False
            Height          =   2055
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   4200
            Width           =   7480
            Begin VB.Frame FraAbono 
               Caption         =   "Abono"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   120
               TabIndex        =   7
               Top             =   960
               Width           =   5535
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Cuenta :"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   13
                  Top             =   270
                  Width           =   600
               End
               Begin VB.Label LblCtaAbo 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
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
                  Height          =   300
                  Left            =   705
                  TabIndex        =   12
                  Top             =   255
                  Width           =   2055
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo :"
                  Height          =   195
                  Left            =   2865
                  TabIndex        =   11
                  Top             =   285
                  Width           =   405
               End
               Begin VB.Label LblTipoCta 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   270
                  Left            =   3345
                  TabIndex        =   10
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Age.:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   9
                  Top             =   600
                  Width           =   375
               End
               Begin VB.Label LblAgeAbono 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   270
                  Left            =   705
                  TabIndex        =   8
                  Top             =   555
                  Width           =   2055
               End
            End
            Begin VB.TextBox txtImpuesto 
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
               Height          =   255
               Left            =   3930
               TabIndex        =   6
               Top             =   480
               Width           =   1170
            End
            Begin VB.TextBox txtInteres 
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
               Height          =   285
               Left            =   3945
               TabIndex        =   5
               Top             =   165
               Width           =   1170
            End
            Begin VB.TextBox txtCostoCustodia 
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
               Height          =   255
               Left            =   1515
               TabIndex        =   4
               Top             =   465
               Width           =   1215
            End
            Begin VB.TextBox txtCostoTasacion 
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
               Height          =   270
               Left            =   1515
               TabIndex        =   3
               Top             =   165
               Width           =   1215
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Total a Pagar :"
               Height          =   255
               Index           =   1
               Left            =   6015
               TabIndex        =   23
               Top             =   1335
               Width           =   1170
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "ITF"
               Height          =   255
               Index           =   0
               Left            =   6000
               TabIndex        =   22
               Top             =   750
               Width           =   1170
            End
            Begin VB.Label LblTotalPagar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   6000
               TabIndex        =   21
               Top             =   1575
               Width           =   1395
            End
            Begin VB.Label LblITF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   6000
               TabIndex        =   20
               Top             =   990
               Width           =   1395
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Neto a Recibir :"
               Height          =   255
               Index           =   7
               Left            =   6000
               TabIndex        =   19
               Top             =   180
               Width           =   1170
            End
            Begin VB.Label lblNetoRecibir 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   6000
               TabIndex        =   18
               Top             =   420
               Width           =   1395
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Impuesto :"
               Height          =   255
               Index           =   3
               Left            =   3075
               TabIndex        =   17
               Top             =   435
               Width           =   855
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Interes :"
               Height          =   255
               Index           =   2
               Left            =   3075
               TabIndex        =   16
               Top             =   165
               Width           =   780
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Costo Custodia :"
               Height          =   240
               Index           =   4
               Left            =   180
               TabIndex        =   15
               Top             =   450
               Width           =   1260
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Costo Tasación :"
               Height          =   255
               Index           =   5
               Left            =   165
               TabIndex        =   14
               Top             =   195
               Width           =   1245
            End
         End
         Begin SICMACT.ActXCodCta AXCodCta 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            Texto           =   "Crédito"
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
         End
         Begin SICMACT.ActXColPDesCon AXDesCon 
            Height          =   3495
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   6165
         End
         Begin VB.Label lblCampaPrend 
            Caption         =   "campana"
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   3840
            TabIndex        =   44
            Top             =   120
            Width           =   3135
         End
      End
   End
End
Attribute VB_Name = "frmColPDesemb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'DESCRIPCION: REALIZA EL DESEMBOLSO DE CREDITO PIGNORATICIO
'             EN EFECTIVO Ó CON ABONO A CUENTA EXISTENTE O NUEVA
'
'Archivo: frmColPDesemb.frm
'PEAC: 20090430
'*************************************************************

Option Explicit

Private RPersCtaAho() As ADODB.Recordset

Private pRSRela As ADODB.Recordset
Private pnTasa As Double
Private pnPersoneria As Integer
Private pnTipoCuenta As Integer
Private pnTipoTasa As Integer
Private pbDocumento As Boolean
Private psNroDoc As String
Private psCodIF As String

Dim loItf1 As Double, loItf2 As Double
Dim vFecVenc As Date
Dim lnOpeCod As Long
Dim lsPersCod As String
Dim sOperacion As String
Dim lbDesembCC As Boolean
Dim lbCuentaNueva As Boolean
Dim rsRel As ADODB.Recordset

Dim lcCodPersTitu As String
Dim nContPersCtaAho As Integer
'Dim RsPersCtaAho As Variant
Dim lcCodCli As String, lcNomCli As String, lcDocNat As String
Dim lcDocJur As String, lcCliDirec As String

Dim sCtaAho As String
Dim pnFilaSelecCtaAho As Integer
'Dim vbCuentaNueva As Boolean

Dim lcNumCtaAho As String
Dim MatTitulares As Variant
Dim nProgAhorros As Integer
Dim nMontoAbonar As Double
Dim nPlazoAbonar As Integer
Dim sPromotorAho As String
Dim nRedondeoITF As Double
'Dim ventana As Integer 'MADM 20090928
Dim pnCampana As Integer 'JOEP20210915 campana prendario


'Inicializa las variables del formulario
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtCostoTasacion.Text = ""
    txtCostoCustodia.Text = ""
    Txtinteres.Text = ""
    txtImpuesto.Text = ""
    lblNetoRecibir.Caption = ""
    
    ReDim RPersCtaAho(0)
    ReDim MatCargoAutom(0)
        
    LimpiaFlex FECtaAhoDesemb
    LimpiaFlex FEPersCtaAho
    'JOEP20210915 campana prendario
    lblCampaPrend.Visible = False
    lblCampaPrend.Caption = ""
    fr_CampPrendReferido.Enabled = False
    txtCampaPrendUser.Text = ""
    txtCampaPrendTrabajador.Text = ""
    pnCampana = 0
    'JOEP20210915 campana prendario
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim loValMontoPrestamo As Double
Dim loValOtrosCostos As Double
Dim lsmensaje As String

'''***
Dim oCredito As COMDCredito.DCOMCredito
Dim rsDesemb As ADODB.Recordset
Dim oCaptacion As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCuentas As ADODB.Recordset, RTemp As ADODB.Recordset
'Dim lcCodPersTitu As String
Dim lvAgencias As Variant
Dim vAgencias As Variant

Dim lvCtaAhorro As Variant
Dim nNumCtasAhorro As Integer
Dim i As Integer
'Dim RPersCtaAho As Variant
'Dim lcNumCtaAho As String
Dim nMontoITF As Double
'********
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas

Dim Rf As ADODB.Recordset
'----- MADM
Dim oCreditoN As COMNCredito.NCOMCredito 'WIOR 20140203
Dim sPersCodTitular As String 'JUEZ 20140217
    
    'JUEZ 20140217 ****************************************
    Dim oPers  As COMDPersona.UCOMPersona
    Dim oDCredito As COMDCredito.DCOMCredito
    Set oDCredito = New COMDCredito.DCOMCredito
    sPersCodTitular = oDCredito.RecuperaTitularCredito(psNroContrato)
    
    Set oPers = New COMDPersona.UCOMPersona
    If oPers.fgVerificaEmpleado(sPersCodTitular) Or oPers.fgVerificaEmpleadoVincualdo(sPersCodTitular) Then
    'END JUEZ *********************************************
        'WIOR 20140203 *********************
        Set oCreditoN = New COMNCredito.NCOMCredito
        If Not oCreditoN.ExisteAsignaSaldo(psNroContrato, 2) Then
            MsgBox "El crédito aún no tiene saldo asignado, verificar con el Departamente de Administración de Créditos.", vbInformation, "Aviso"
            Exit Sub
        End If
        Set oCreditoN = Nothing
        'WIOR FIN **************************
    End If
    Set oPers = Nothing
        
'On Error GoTo ControlError
    gITF.fgITFParamAsume (Mid(psNroContrato, 4, 2)), Mid(psNroContrato, 6, 3)
    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaDesembolsoCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If

    Set loValContrato = Nothing
    
    lcCodPersTitu = lrValida!cperscod
    
    If lrValida Is Nothing Then ' Hubo un Error
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    'Muestra Datos
    lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)
        If Not lbOk Then 'MACM 20210319 VALIDA REPUESTA DE SOL DE TASA PREFERERNCIAL
        AXCodCta.Enabled = True
        AXCodCta.SetFocusCuenta
        cmdCancelar_Click
        Exit Sub
    End If
        
    Me.Txtinteres = Format(lrValida!nInteres, "#0.00")
    Me.txtImpuesto = Format(lrValida!nImpuesto, "#0.00")
    Me.txtCostoTasacion = Format(lrValida!nTasacion, "#0.00")
    Me.txtCostoCustodia = Format(lrValida!nCustodia, "#0.00")
    'PEAC 20070813
    vFecVenc = Format(lrValida!dVenc, "dd/mm/yyyy")


    '''*** Llena los datos del cliente
        lcCodCli = lrValida!cperscod
        lcNomCli = lrValida!cPersNombre
        lcDocNat = IIf(IsNull(lrValida!cDNI), "", lrValida!cDNI)
        lcDocJur = IIf(IsNull(lrValida!cRUC), "", lrValida!cRUC)
        lcCliDirec = IIf(IsNull(lrValida!cPersDireccDomicilio), "", lrValida!cPersDireccDomicilio)
    '''*** FIN

    Set lrValida = Nothing
    
    loValMontoPrestamo = CCur(AXDesCon.SaldoCapital)
    loValOtrosCostos = CCur(Me.txtCostoTasacion) + CCur(txtCostoCustodia.Text) + CCur(Txtinteres.Text) + CCur(txtImpuesto.Text)
        
    '*** PEAC 20071206 - en el neto a recibir no estara el interes adelantado ********************
    'Me.lblNetoRecibir.Caption = Format(CCur(AXDesCon.SaldoCapital) - (CCur(Me.txtCostoTasacion) + CCur(txtCostoCustodia.Text) + CCur(txtInteres.Text) + CCur(txtImpuesto.Text)), "#0.00")
    Me.lblNetoRecibir.Caption = Format(CCur(AXDesCon.SaldoCapital) - (CCur(Me.txtCostoTasacion) + CCur(txtCostoCustodia.Text) + CCur(txtImpuesto.Text)), "#0.00")
    '*********************************************************************************************
    
    'CROB20180609 ERS076-2017
    If loValMontoPrestamo > 3500 And SSTabDatos.TabVisible(1) = False Then
        MsgBox ("El monto del prestamo S/." & Format(loValMontoPrestamo, "#,##0.00") & " Soles" & Chr(10) & "debera usar la opción abono cuenta."), vbInformation, "Aviso"
        cmdCancelar_Click
        Exit Sub
    End If
    'CROB20180609 ERS076-2017
    
    'JOEP20210927 campana prendario
    Call CampPrendVerificaCampanas(psNroContrato, gdFecSis, 0)
    Call HabilitaReferido(psNroContrato)
    'JOEP20210927 campana prendario
    
    ' **************  ITF ***************
    If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
            loItf1 = Format(gITF.fgITFCalculaImpuesto(Val(Me.lblNetoRecibir.Caption)), "#0.00")
            'loItf2 = Format(gITF.fgITFCalculaImpuesto(loValOtrosCostos), "#0.00")
            'Me.LblITF = Format(gITF.fgITFCalculaImpuesto(CDbl(Me.lblNetoRecibir)), "#0.00")
            'Me.LblITF = Format(loItf1 + loItf2, "#0.00")
            Me.lblITF = Format(loItf1, "#0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
            If nRedondeoITF > 0 Then
               Me.lblITF = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
               loItf1 = lblITF.Caption
            End If
            '*** END BRGO
            Me.LblTotalPagar = Format(CDbl(Me.lblNetoRecibir) - CDbl(Me.lblITF), "#0.00")
        Else
            loItf1 = Format(gITF.fgITFCalculaImpuesto(loValMontoPrestamo), "#0.00")
            loItf2 = Format(gITF.fgITFCalculaImpuesto(loValOtrosCostos), "#0.00")
            'Me.LblITF = Format(gITF.fgITFCalculaImpuesto(CDbl(Me.lblNetoRecibir)), "#0.00")
            Me.lblITF = Format(loItf1 + loItf2, "#0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
            If nRedondeoITF > 0 Then
               Me.lblITF = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
               If loItf1 >= nRedondeoITF Then
                    loItf1 = loItf1 - nRedondeoITF
               Else
                    loItf2 = loItf2 - nRedondeoITF
               End If
            End If
            '*** END BRGO
            Me.LblTotalPagar = Me.lblNetoRecibir
        End If
    Else
        Me.lblITF = Format(0, "#0.00")
        Me.LblTotalPagar = Me.lblNetoRecibir
    End If
    ' **************  ITF ***************
    
    cmdgrabar.Enabled = True
    cmdgrabar.SetFocus
    
    
''----------------------------------------------------------------------------

    '''*** Llena las cuentas de ahorro para el abono
    If lbDesembCC Then '' si es desembolso en cta de ahorro
        Dim oCtaAho As COMNColoCPig.NCOMColPValida
        Set oCtaAho = New COMNColoCPig.NCOMColPValida
        
        Call oCtaAho.CargarDatosCtaAho(psNroContrato, gdFecSis, lcCodCli, rsCuentas, vAgencias, nContPersCtaAho, RPersCtaAho, nMontoITF, _
                                    sOperacion)
                            
        LimpiaFlex FECtaAhoDesemb
        If rsCuentas.RecordCount > 0 Then rsCuentas.MoveFirst
        
        If Not (rsCuentas.EOF And rsCuentas.BOF) Then
            Me.CmdAperturar.Enabled = True
            Me.CmdDeseleccionar.Enabled = True
            Me.CmdSeleccionar.Enabled = True
        
            Do While Not rsCuentas.EOF
                CmdSeleccionar.Enabled = True
                FECtaAhoDesemb.AdicionaFila , , True
                FECtaAhoDesemb.TextMatrix(rsCuentas.Bookmark, 1) = rsCuentas!cCtaCod
                FECtaAhoDesemb.TextMatrix(rsCuentas.Bookmark, 2) = vAgencias(rsCuentas.Bookmark)
                FECtaAhoDesemb.TextMatrix(rsCuentas.Bookmark, 3) = IIf(rsCuentas!Exonerada = "", "Afecta", "Exonerada")
                rsCuentas.MoveNext
            Loop
            FECtaAhoDesemb.row = 1

            LimpiaFlex FEPersCtaAho
            If nContPersCtaAho > 0 Then
                RPersCtaAho(0).MoveFirst
                Call CargaClientesCtaAho(RPersCtaAho(0)!cCtaCod)
            End If
        Else
            LimpiaFlex FEPersCtaAho
            Me.CmdAperturar.Enabled = True
        End If
    End If

''----------------------------------------------------------------------------
         '************ firma madm 20090928
       'If ventana = 0 Then
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(lcCodCli, BusquedaCodigo)
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(lcCodCli), Mid(lcCodCli, 4, 2), False, False) 'MOD BY JATO 20210324 true --> false
        End If
         End If
         Set Rf = Nothing
      '  End If
      '  ventana = 0
    '************ firma madm
 
    AXCodCta.Enabled = False

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
'ALPA 20090420 Comentado para compilacion sin cambios dejado por CAPI
'By capi 16032009
'    If (AXCodCta.NroCuenta) <> "" Then
'        cmdCtasAhorros.Visible = True
'        cmdCtasAhorros.SetFocus
'    End If
    '
End Sub

Private Sub CmdAperturar_Click()
    If CmdDeseleccionar.Visible Then
        Call CmdDeseleccionar_Click
    End If
    lbCuentaNueva = True
    
'    frmCapAperturas.IniciaDesembAbonoCta gCapAhorros, gAhoApeEfec, "", CInt(Mid(Me.AXCodCta.NroCuenta, 9, 1)), _
'         lcCodPersTitu, lcNomCli, lblNetoRecibir, Nothing, _
'        0, 0, 0, 0, False, "", "", False, _
'        lcCodCli, lcNomCli, , nProgAhorros, nMontoAbonar, nPlazoAbonar, sPromotorAho
        
    frmCapAperturas.IniciaDesembAbonoCta gCapAhorros, gAhoApeEfec, "", CInt(Mid(Me.AXCodCta.NroCuenta, 9, 1)), _
         lcCodPersTitu, lcNomCli, lblNetoRecibir, pRSRela, _
        pnTasa, pnPersoneria, pnTipoCuenta, pnTipoTasa, pbDocumento, psNroDoc, psCodIF, False, _
         , , MatTitulares, nProgAhorros, nMontoAbonar, nPlazoAbonar, sPromotorAho
        

        
    Me.LblTipoAbono.Caption = "NUEVA"
    Me.LblCtaAbo.Caption = "NUEVA"
    Me.LblCtaAbo.Alignment = 2
    Me.LblTipoCta.Caption = "Propia"
    Me.LblAgeAbono.Caption = gsNomAge

End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona

'By capi 06032009
'Dim lsPersCod As String, lsPersNombre As String
Dim lsPersNombre As String

Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstRegis

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    cmdgrabar.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub
'ALPA 20090420 Comentado para compilacion sin cambios dejado por CAPI
'Private Sub cmdCtasAhorros_Click()
'    frmCredDesembAbonoCta.DesembolsoPigAbonoCuenta Str(gColPOpeDesembolsoAboCta), lsPersCod, AXCodCta.NroCuenta
'End Sub

Private Sub cmdCtasAhorros_Click()

End Sub

Private Sub CmdDeseleccionar_Click()
    sCtaAho = ""
    CmdSeleccionar.Visible = True
    CmdDeseleccionar.Visible = False
    If pnFilaSelecCtaAho <> -1 Then
        FECtaAhoDesemb.row = pnFilaSelecCtaAho
        Call FECtaAhoDesemb.BackColorRow(vbWhite)
    End If
    
    pnFilaSelecCtaAho = -1
    
    LblCtaAbo.Caption = ""
    LblTipoCta.Caption = ""
    LblAgeAbono.Caption = ""
    LblTipoAbono.Caption = ""
    
End Sub

Private Sub cmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarDesem As COMNColoCPig.NCOMColPContrato
Dim loColImp As COMNColoCPig.NCOMColPImpre
Dim clsprevio As New previo.clsprevio
Dim lsCadImp As String
Dim opt As Integer
Dim OptBt2 As Integer
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsFechaHoraPrend As String
Dim lsCuenta As String
Dim nFicSal As Integer
Dim psCtaAhoNew As String
Dim oITF  As COMDConstSistema.FCOMITF
Dim oMov As COMDMov.DCOMMov
Dim oRegRefCampPren As COMDColocPig.DCOMColPContrato

Dim lnValorItfAbono As Double
'Dim lnValorItfGasto As Double
'Dim lnValorItfCancelacion As Double

Dim lnMovNro As Long
Dim MatDatosAho(14) As String
Dim lcTextImp As String
Dim lnSaldoCap As Currency, lnInteresComp As Currency, lnImpuesto As Currency
Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
Dim lnMontoEntregar As Currency
Set loColImp = New COMNColoCPig.NCOMColPImpre

psCtaAhoNew = ""
lnValorItfAbono = 0: 'lnValorItfGasto = 0: lnValorItfCancelacion = 0

lsCuenta = AXCodCta.NroCuenta
lnSaldoCap = Me.AXDesCon.SaldoCapital
lnInteresComp = CCur(Me.Txtinteres.Text)
lnImpuesto = CCur(Me.txtImpuesto.Text)
lnCostoCustodia = CCur(Me.txtCostoCustodia.Text)
lnCostoTasacion = CCur(Me.txtCostoTasacion.Text)
lnMontoEntregar = CCur(Me.LblTotalPagar.Caption)

If lbDesembCC And LblTipoAbono.Caption = "" Then
    MsgBox "Por favor seleccione una Cuenta de Ahorro afecta a ITF...", vbInformation, "Atención"
    Exit Sub
End If

'CROB20180528 begin
    Dim oDNiv As COMDColocPig.DCOMColPContrato
    Set oDNiv = New COMDColocPig.DCOMColPContrato

    If oDNiv.ExisteAprobacionCredNivelesPendientesPigno(lsCuenta) Then
        MsgBox ("Ud. No puede finalizar la operación hasta que" & Chr(10) & "Apruebe el Supervisor de Operaciones."), vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Not oDNiv.ValidarNivelAprobacionPigno(lsCuenta) Then
        MsgBox ("Solicitud de nivel de aprobación fue rechazada"), vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set oDNiv = Nothing
'CROB20180528 end

'EJVG20120322 Verifica actualización Persona
Dim oPersona As New COMNPersona.NCOMPersona
If oPersona.NecesitaActualizarDatos(lcCodPersTitu, gdFecSis) Then
     MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Titular: " & lcNomCli, vbInformation, "Aviso"
     Dim foPersona As New frmPersona
     If Not foPersona.realizarMantenimiento(lcCodPersTitu) Then
         MsgBox "No se ha realizado la actualización de los datos de " & lcNomCli & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
         Exit Sub
     End If
End If
'WIOR 20121009**********************************************************
Dim oDPersona As COMDPersona.DCOMPersona
Dim rsPersonaCred As ADODB.Recordset
Dim rsPersona As ADODB.Recordset
Dim Cont As Integer
Set oDPersona = New COMDPersona.DCOMPersona

Set rsPersonaCred = oDPersona.ObtenerPersCuentaRelac(Trim(AXCodCta.NroCuenta), gColRelPersTitular)

If rsPersonaCred.RecordCount > 0 Then
    If Not (rsPersonaCred.EOF And rsPersonaCred.BOF) Then
        For Cont = 0 To rsPersonaCred.RecordCount - 1
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cperscod))
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                    MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.Inicio(Trim(rsPersonaCred!cperscod), PersonaActualiza)
                    End If
                End If
            End If
            Set rsPersona = Nothing
            rsPersonaCred.MoveNext
        Next Cont
    End If
End If
'WIOR FIN ***************************************************************
    

If MsgBox(" Grabar Desembolso de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdgrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing

        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        lsFechaHoraPrend = fgFechaHoraPrend(lsMovNro)

        Set loGrabarDesem = New COMNColoCPig.NCOMColPContrato

        'Grabar Desembolso Pignoraticio
        Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(lsCuenta) Then
          Dim sPersLavDinero As String
          Dim nMontoLavDinero As Double, nTC As Double
          Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion, nMoneda As Integer, nMonto As Double
           'MARG ERS073***
          Dim bLavDinero As Boolean
          Dim loLavDinero As frmMovLavDinero
          Set loLavDinero = New frmMovLavDinero
          'END MARG******

            nMonto = CDbl(LblTotalPagar.Caption)
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            nMoneda = gMonedaNacional
            If nMoneda = gMonedaNacional Then
                'Modificar
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                'By Capi 18022008 no aplica a desembolsos segun manifestado por riesgos
                'sPersLavDinero = IniciaLavDinero()
                 sPersLavDinero = ""
                'If sPersLavDinero = "" Then Exit Sub
                'MARG ERS073***
                loLavDinero.TitPersLavDinero = Trim(lcCodPersTitu)
                sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, lsCuenta, Mid(Me.Caption, 1), True, "", , , , , gMonedaNacional, , gnTipoREU, gnMontoAcumulado, gsOrigen, , gsOpeCod)
                bLavDinero = True
                 If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                'END MARG******
            End If
         Else
            Set clsExo = Nothing
         End If
                        
        '*** PEAC 20071206 - no se graba el interes solo se muestra ***********************************
            'Call loGrabarDesem.nDesembolsoCredPignoraticio(lsCuenta, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, lnInteresComp, lnImpuesto, lnCostoTasacion, lnCostoCustodia, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, loItf1, loItf2, CCur(lblNetoRecibir.Caption), False)
            
        '*** PEAC 20090505
            'Call loGrabarDesem.nDesembolsoCredPignoraticio(lsCuenta, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, 0, lnImpuesto, lnCostoTasacion, lnCostoCustodia, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, loItf1, loItf2, CCur(lblNetoRecibir.Caption), False)
            
            'RECO20140311 ERS160-2013*********************************************************************
            'CONTROL DE SISTEMA LAVADO DE ACTIVOS
            Dim lbResultadoVisto As Boolean
            Dim loVistoElectronico As frmVistoElectronico
            Dim bVisto As Boolean
            Set loVistoElectronico = New frmVistoElectronico
            If lnSaldoCap > 5000 Then
                bVisto = True
            End If
            If loGrabarDesem.ObtieneMontoDesembCreDPigMes(lsPersCod, Format(gdFecSis, "yyyyMM")) + lnSaldoCap > 5000 Then
                bVisto = True
            End If
            If bVisto = True Then
                loVistoElectronico.VistoMovNro = lsMovNro
                lbResultadoVisto = loVistoElectronico.Inicio(10, sOperacion)
                If Not lbResultadoVisto Then
                    MsgBox "Operación cancelada por el usuario", vbInformation, "Aviso"
                    Exit Sub
                End If
            End If
            'RECO FIN*************************************************************************************
            Call loGrabarDesem.nDesembolsoCredPignoEfectivoAbono(lsCuenta, lbDesembCC, lbCuentaNueva, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, 0, lnImpuesto, lnCostoTasacion, lnCostoCustodia, IIf(lbDesembCC, False, gITF.gbITFAplica), gITF.gbITFAsumidocreditos, loItf1, loItf2, CCur(lblNetoRecibir.Caption), lnMovNro, False)
            '*** BRGO 20110914 *********************
            If gITF.gbITFAplica Then
                Set oMov = New COMDMov.DCOMMov
                Call oMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Me.lblITF.Caption) + nRedondeoITF, CCur(Me.lblITF.Caption))
                Set oMov = Nothing
            End If
            '*** END BRGO *******************************
''********** aqui graba el abono en cta

    If lbDesembCC Then

        ReDim MatDatosAhoNew(4)

        MatDatosAhoNew(0) = nProgAhorros
        MatDatosAhoNew(1) = nMontoAbonar
        MatDatosAhoNew(2) = nPlazoAbonar
        MatDatosAhoNew(3) = sPromotorAho
            
        Set oITF = New COMDConstSistema.FCOMITF
        oITF.fgITFParametros
        oITF.fgITFParamAsume Mid(lsCuenta, 4, 2)
    
        lnValorItfAbono = CDbl(lblITF.Caption)
        'lnValorItfGasto = oITF.fgITFDesembolso(0) 'pnMontoGastos
        'lnValorItfCancelacion = oITF.fgITFDesembolso(0) 'pnMontoCancel
        
        Dim pMatGastos() As String
        Dim pMatCargoAutom() As String
        
        If lbCuentaNueva Then
            Call loGrabarDesem.DesembolsoPignoConAbonoCta(lnMovNro, MatDatosAho, lsCuenta, lnSaldoCap, _
                lnMontoEntregar, pMatGastos, pMatCargoAutom, gdFecSis, Mid(lsCuenta, 4, 2), gsCodUser, True, True, , , sCtaAho, _
                 pRSRela, pnTasa, pnPersoneria, pnTipoCuenta, pnTipoTasa, pbDocumento, psNroDoc, psCodIF, , , , , _
                oITF.gbITFAplica, lnValorItfAbono, , , False, gITFCobroCargoPigno, , , , , , psCtaAhoNew, MatDatosAhoNew, , , lsMovNro)
        Else

            Call loGrabarDesem.DesembolsoPignoConAbonoCta(lnMovNro, MatDatosAho, lsCuenta, lnSaldoCap, _
                lnMontoEntregar, pMatGastos, pMatCargoAutom, gdFecSis, Mid(lsCuenta, 4, 2), gsCodUser, True, False, , , sCtaAho, , , , , , , , , _
                , , , , _
                oITF.gbITFAplica, lnValorItfAbono, , , False, gITFCobroCargoPigno, , , , , , , , , , lsMovNro)
                'ALPA20130827 se Agregó lsMovNro
        End If
    End If
    Set oITF = Nothing
    
    'JOEP20210915 Campana Prendario
    If lblCampaPrend.Visible = True Then
        Set oRegRefCampPren = New COMDColocPig.DCOMColPContrato
        Call oRegRefCampPren.CampPrenRegCampCred(lsCuenta, pnCampana, "Desembolso", AXDesCon.TasaEfectivaMensual, 0, 0, 0, lsMovNro, 1, 1)
        Set oRegRefCampPren = Nothing
    End If
    If fr_CampPrendReferido.Enabled = True Then
        Set oRegRefCampPren = New COMDColocPig.DCOMColPContrato
        Call oRegRefCampPren.CampPrenDesbRegReferido(lsCuenta, txtCamPredPersCodReferido.Text, lsMovNro, 1)
        Set oRegRefCampPren = Nothing
    End If
'JOEP20210915 Campana Prendario

'MARG ERS073*****************************************
    If bLavDinero Then
        Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , lnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4)
        Set loLavDinero = Nothing
    End If
'END MARG********************************************
'********************************************************

        If Trim(Me.LblTipoAbono.Caption) = "NUEVA" Then
           'IMPRIME REGISTRO DE FISMAS
           Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
           Dim lsCadImpFirmas As String
           Dim lsCadImpCartilla As String
           Dim sTipoCuenta As String
           If pnTipoCuenta = 0 Then
                sTipoCuenta = "INDIVIDUAL"
           ElseIf pnTipoCuenta = 1 Then
                sTipoCuenta = "MALCOMUNADA"
           ElseIf pnTipoCuenta = 2 Then
                'sTipoCuenta = "INDISTINTA"
                sTipoCuenta = "SOLIDARIA" 'APRI20190109 ERS077-2018
           End If
           Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                clsMant.IniciaImpresora gImpresora
                lsCadImpFirmas = clsMant.GeneraRegistroFirmas(psCtaAhoNew, sTipoCuenta, gdFecSis, False, pRSRela, gsNomAge, gdFecSis, gsCodUser)
           Set clsMant = Nothing
           Set rsRel = Nothing
           
           'IMPRIME CARTILLA EN WORD
           Dim lnTasaE As Double
           lnTasaE = Round(((1 + (pnTasa / 100 / 12) / 30) ^ 360 - 1) * 100, 2)
            ImpreCartillaAhoCorriente MatTitulares, psCtaAhoNew, lnTasaE, lnSaldoCap, nProgAhorros

        End If

'********************************************************
        '*** IMPRIME REGISTRO DE FIRMAS - AHORRO
        If Trim(LblTipoAbono.Caption) = "NUEVA" Then
            MsgBox "Coloque Papel Continuo Tamaño Carta, Para la Impresion del Registros de Firmas", vbInformation, "Aviso"
            clsprevio.PrintSpool sLpt, oImpresora.gPrnCondensadaON & lsCadImpFirmas & oImpresora.gPrnCondensadaOFF, False, gnLinPage   'ARCV 01-11-2006
        End If

'********************************************************

        lsCadImp = loColImp.nPrintReciboDesembolso(vFecVenc, lsCuenta, lnSaldoCap, lsFechaHoraPrend, _
                       lnMontoEntregar, lnInteresComp, gsNomAge, gsCodUser, CDbl(lblITF.Caption), gImpresora, lbDesembCC, sCtaAho, psCtaAhoNew, gbImpTMU)

'*********************************************************
        lcTextImp = "Desea Imprimir las boletas de desembolsos"
        Do
        OptBt2 = MsgBox(lcTextImp, vbInformation + vbYesNo, "Aviso")
        If vbYes = OptBt2 Then
        lcTextImp = "Desea Reimprimir las boletas de desembolsos"
        MsgBox "Cambie de Papel para imprimir las boletas de desembolsos", vbExclamation, "Aviso"
        nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
                Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
                Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
                Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
                Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
                Print #nFicSal, lsCadImp & Chr$(12)
                Print #nFicSal, ""
                Close #nFicSal
        End If
        Loop Until OptBt2 = vbNo
            
'*********************************************************
        'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", sOperacion
        'FIN
        Set loGrabarDesem = Nothing
        Set loColImp = Nothing
        Limpiar
        Me.lblNetoRecibir = "0.00"
        Me.lblITF = "0.00"
        Me.LblTotalPagar = "0.00"

        AXCodCta.Enabled = True
        AXCodCta.SetFocus
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Function IniciaLavDinero() As String
Dim i As Long
Dim nRelacion As CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nMonto As Double, nPersoneria As Integer
Dim sCuenta As String
'For i = 1 To grdCliente.Rows - 1
    'nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    nPersoneria = gPersonaNat
    If nPersoneria = gPersonaNat Then
        'If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            sPersCod = AXDesCon.listaClientes.ListItems(1).Text
            sNombre = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            sDireccion = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            sDocId = AXDesCon.listaClientes.ListItems(1).SubItems(7)
         '   Exit For
       ' End If
    Else
        'If nRelacion = gCapRelPersTitular Then
            sPersCod = AXDesCon.listaClientes.ListItems(1).Text
            sNombre = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            sDireccion = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            sDocId = AXDesCon.listaClientes.ListItems(1).SubItems(9)
          '  Exit For
        'End If
    End If
'Next i
nMonto = CDbl(LblTotalPagar.Caption)
sCuenta = AXCodCta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nmonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, gColPOpeDesembolsoEFE, , gMonedaNacional)
'End If
End Function

'Finaliza el formulario actual
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSeleccionar_Click()
    'By Capi 11012007 para que valide que no sea una cuenta exonerada de ITF
    If Trim(FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 3)) = "Exonerada" Then
        MsgBox "No Puede Elegir Una Cuenta Exonerada ITF...", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Trim(FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 1)) <> "" Then
        sCtaAho = FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 1)
        pnFilaSelecCtaAho = FECtaAhoDesemb.row
        CmdSeleccionar.Visible = False
        CmdDeseleccionar.Visible = True
        Call FECtaAhoDesemb.BackColorRow(vbYellow)
        LblCtaAbo.Caption = sCtaAho
        LblTipoCta.Caption = "Propia"
        LblAgeAbono.Caption = FECtaAhoDesemb.TextMatrix(FECtaAhoDesemb.row, 2)
        LblTipoAbono.Caption = "EXISTENTE"
        lbCuentaNueva = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
            SendKeys "{Enter}"
        End If
    ElseIf KeyCode = 13 And Trim(AXCodCta.EnabledCta) And AXCodCta.Age <> "" And Trim(AXCodCta.Cuenta) = "" Then
                AXCodCta.SetFocusCuenta
                 Exit Sub
        
    End If
End Sub

'Inicializa el formulario actual
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    'ventana = 0
    Limpiar
End Sub

Public Sub DesembolsoEfectivo(ByVal sCodOpe As String)
    sOperacion = sCodOpe
    SSTabDatos.TabVisible(1) = False
    
    lbDesembCC = False
    
'    vbCuentaNueva = False
'    vbDesembCheque = False

    Me.Caption = "Desembolso en Efectivo"
    Me.Show 1
    
End Sub

Public Sub DesembolsoCargoCuenta(ByVal sCodOpe As String)
    SSTabDatos.TabVisible(1) = True
    sOperacion = sCodOpe
    
    lbDesembCC = True
    
'    vbCuentaNueva = False
'    vbDesembCheque = False

    Me.Caption = "Desembolso con Abono a Cuenta"
    Me.Show 1
End Sub

Private Sub CargaClientesCtaAho(ByVal psCtaCod As String)
Dim i As Integer
Dim J As Integer
Dim sPersCod As String
     LimpiaFlex FEPersCtaAho
     sPersCod = ""
     J = 0
    For i = 0 To nContPersCtaAho - 1
        RPersCtaAho(i).MoveFirst
        If RPersCtaAho(i)!cCtaCod = psCtaCod Then
            Do While Not RPersCtaAho(i).EOF
                FEPersCtaAho.AdicionaFila
                If sPersCod <> RPersCtaAho(i)!cperscod Then
                    J = J + 1
                    FEPersCtaAho.TextMatrix(J, 1) = PstaNombre(RPersCtaAho(i)!Nombre)
                    FEPersCtaAho.TextMatrix(J, 2) = Trim(Mid(RPersCtaAho(i)!Relacion, 1, 30)) 'Trim(RPersCtaAho(i)!Relacion)
                    
                End If
                sPersCod = RPersCtaAho(i)!cperscod
                RPersCtaAho(i).MoveNext
            Loop
            Exit For
        End If
    Next i
End Sub


'JOEP20210915 campana prendario
Private Sub cmdCampPrendBuscar_Click()
Dim oPers As COMDPersona.UCOMPersona
Dim lcPersNombre As String
Dim lcPersCod As String
Dim lcEstados As String
Dim oPersContrato As COMDColocPig.DCOMColPContrato
Dim rsContratos As New ADODB.Recordset
Dim oCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

txtCampaPrendUser.Text = ""
txtCampaPrendTrabajador.Text = ""
txtCamPredPersCodReferido.Text = ""

Set oPers = New COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If oPers Is Nothing Then Exit Sub
    lcPersCod = oPers.sPersCod
    lcPersNombre = oPers.sPersNombre
Set oPers = Nothing

If Trim(lcPersCod) <> "" Then
    Set oPersContrato = New COMDColocPig.DCOMColPContrato
        Set rsContratos = oPersContrato.CampPrendarioDesbVerfCampaReferido(lcPersCod, gsCodUser)
    Set oPersContrato = Nothing

    If Not (rsContratos.BOF And rsContratos.EOF) Then
        If rsContratos!nPase <> 0 Then
            txtCampaPrendUser.Text = rsContratos!cUser
            txtCampaPrendTrabajador.Text = rsContratos!cPersNombre
            txtCamPredPersCodReferido.Text = rsContratos!cperscod
        Else
            MsgBox rsContratos!cMsgbox, vbInformation, "Aviso"
            txtCampaPrendUser.Text = ""
            txtCampaPrendTrabajador.Text = ""
            txtCamPredPersCodReferido.Text = ""
        End If
    Else
        txtCampaPrendUser.Text = ""
        txtCampaPrendTrabajador.Text = ""
        txtCamPredPersCodReferido.Text = ""
    End If
End If

RSClose rsContratos
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
'JOEP20210915 campana prendario

Private Sub CampPrendVerificaCampanas(ByVal pcCtaCod As String, ByVal pdFechaSis As String, ByVal pnModulo As Integer)
    Dim oCampPrend As COMDColocPig.DCOMColPContrato
    Dim rsCampPrend As ADODB.Recordset
    Set oCampPrend = New COMDColocPig.DCOMColPContrato
    
    Set rsCampPrend = oCampPrend.CampPrendarioDesbCampa(pcCtaCod, pdFechaSis, pnModulo)
    txtCampaPrendUser.Text = ""
    txtCampaPrendTrabajador.Text = ""
    If Not (rsCampPrend.BOF And rsCampPrend.EOF) Then
        lblCampaPrend.Visible = True
        lblCampaPrend.Caption = rsCampPrend!cResultado
        pnCampana = rsCampPrend!nCampana
    Else
        lblCampaPrend.Visible = False
        lblCampaPrend.Caption = ""
        pnCampana = 0
    End If
    Set oCampPrend = Nothing
    RSClose oCampPrend
End Sub

Private Sub HabilitaReferido(ByVal pcCtaCod As String)
    Dim oRef As COMDColocPig.DCOMColPContrato
    Dim rsRef As ADODB.Recordset
    Set oRef = New COMDColocPig.DCOMColPContrato
    Set rsRef = oRef.CampPrendarioHabilitaReferido(pcCtaCod)
    If Not (rsRef.BOF And rsRef.EOF) Then
        If rsRef!nPase = 0 Then
            fr_CampPrendReferido.Enabled = True
        Else
            fr_CampPrendReferido.Enabled = False
        End If
    Else
        fr_CampPrendReferido.Enabled = False
    End If
    
    Set oRef = Nothing
    RSClose rsRef
End Sub

