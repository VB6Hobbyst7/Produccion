VERSION 5.00
Begin VB.Form frmCapCargos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7560
   ClientLeft      =   2925
   ClientTop       =   1290
   ClientWidth     =   9045
   Icon            =   "frmCapCargos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDocumentoTrans 
      Caption         =   "Documento"
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
      ForeColor       =   &H00000080&
      Height          =   2670
      Left            =   45
      TabIndex        =   55
      Top             =   4410
      Visible         =   0   'False
      Width           =   4920
      Begin VB.TextBox txtTitular 
         Height          =   315
         Left            =   3015
         TabIndex        =   66
         Top             =   1080
         Width           =   1770
      End
      Begin VB.TextBox txtGlosaTrans 
         Height          =   735
         Left            =   1125
         TabIndex        =   65
         Top             =   1845
         Width           =   3660
      End
      Begin VB.TextBox txtCuentaTrans 
         Height          =   315
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   62
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox ckMismoTitular 
         Caption         =   "Mismo Titular"
         Height          =   315
         Left            =   1125
         TabIndex        =   60
         Top             =   1080
         Width           =   1365
      End
      Begin VB.ComboBox cbPlazaTrans 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   720
         Width           =   1815
      End
      Begin SICMACT.TxtBuscar txtBancoTrans 
         Height          =   315
         Left            =   1125
         TabIndex        =   56
         Top             =   315
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
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
      Begin VB.Label lblGlosa 
         Caption         =   "Glosa:"
         Height          =   330
         Left            =   135
         TabIndex        =   64
         Top             =   1890
         Width           =   735
      End
      Begin VB.Label lblNombreBancoTrans 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   735
         Left            =   3015
         TabIndex        =   63
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label Label4 
         Caption         =   "CCI:"
         Height          =   240
         Left            =   135
         TabIndex        =   61
         Top             =   1485
         Width           =   915
      End
      Begin VB.Label lblPlaza 
         Caption         =   "Plaza:"
         Height          =   285
         Left            =   135
         TabIndex        =   59
         Top             =   765
         Width           =   645
      End
      Begin VB.Label lblNombreBanco 
         Caption         =   "Banco:"
         Height          =   285
         Left            =   135
         TabIndex        =   57
         Top             =   315
         Width           =   690
      End
   End
   Begin VB.PictureBox pctNotaAbono 
      Height          =   300
      Left            =   1950
      Picture         =   "frmCapCargos.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   29
      Top             =   7125
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pctCheque 
      Height          =   345
      Left            =   1680
      Picture         =   "frmCapCargos.frx":088C
      ScaleHeight     =   285
      ScaleWidth      =   120
      TabIndex        =   28
      Top             =   7125
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
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
      Height          =   1275
      Left            =   45
      TabIndex        =   16
      Top             =   0
      Width           =   8910
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   90
         TabIndex        =   0
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Frame fraDatos 
         Height          =   585
         Left            =   45
         TabIndex        =   17
         Top             =   630
         Width           =   8800
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lblApertura 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   900
            TabIndex        =   20
            Top             =   195
            Width           =   1965
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Contacto :"
            Height          =   195
            Left            =   3435
            TabIndex        =   19
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lblUltContacto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4740
            TabIndex        =   18
            Top             =   195
            Width           =   1995
         End
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   3960
      End
   End
   Begin VB.Frame fraCliente 
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
      Height          =   3195
      Left            =   45
      TabIndex        =   9
      Top             =   1215
      Width           =   8925
      Begin VB.CommandButton cmdVerRegla 
         Caption         =   "Ver Regla"
         Height          =   315
         Left            =   5760
         TabIndex        =   54
         Top             =   2025
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdMostrarFirma 
         Caption         =   "Mostrar Firma"
         Height          =   315
         Left            =   7350
         TabIndex        =   48
         Top             =   2025
         Width           =   1440
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1755
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3096
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Direccion-ID-Firma Oblig-Grupo-Presente"
         EncabezadosAnchos=   "250-1500-3200-1500-0-0-0-1000-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-8"
         ListaControles  =   "0-0-0-0-0-0-0-0-4"
         EncabezadosAlineacion=   "C-L-L-L-C-C-C-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   255
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin VB.Label LblTituloExoneracion 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Exoneración :"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   2835
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblExoneracion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXONERADA POR  ..."
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
         Left            =   1620
         TabIndex        =   46
         Top             =   2790
         Visible         =   0   'False
         Width           =   7170
      End
      Begin VB.Label lblMinFirmas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   5160
         TabIndex        =   45
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mínimo Firmas :"
         Height          =   195
         Left            =   4680
         TabIndex        =   44
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblAlias 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1620
         TabIndex        =   43
         Top             =   2430
         Width           =   7185
      End
      Begin VB.Label Label3 
         Caption         =   "Alias de la Cuenta:"
         Height          =   225
         Left            =   180
         TabIndex        =   42
         Top             =   2490
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   2123
         Width           =   960
      End
      Begin VB.Label lblTipoCuenta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1620
         TabIndex        =   14
         Top             =   2070
         Width           =   1800
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "# Firmas :"
         Height          =   195
         Left            =   3480
         TabIndex        =   13
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label lblFirmas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   4080
         TabIndex        =   12
         Top             =   2040
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   11
      Top             =   7125
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7950
      TabIndex        =   10
      Top             =   7125
      Width           =   1000
   End
   Begin VB.Frame fraMonto 
      Caption         =   "Monto"
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
      Height          =   2670
      Left            =   5040
      TabIndex        =   25
      Top             =   4410
      Width           =   3930
      Begin VB.CheckBox chkTransfEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   68
         Top             =   2280
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.ComboBox cboMedioRetiro 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   285
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.CheckBox chkVBComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   51
         Top             =   720
         Width           =   705
      End
      Begin VB.CheckBox chkITFEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   37
         Top             =   1485
         Width           =   705
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1020
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         ForeColor       =   192
         Text            =   "0"
      End
      Begin VB.Label lblComisionTransf 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   2235
         TabIndex        =   67
         Top             =   2220
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMedioRetiro 
         AutoSize        =   -1  'True
         Caption         =   "Medio de Retiro :"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblComision 
         AutoSize        =   -1  'True
         Caption         =   "Comision :"
         Height          =   195
         Left            =   495
         TabIndex        =   50
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblMonComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   2235
         TabIndex        =   49
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1440
         TabIndex        =   41
         Top             =   1830
         Width           =   1890
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   2235
         TabIndex        =   40
         Top             =   1425
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   495
         TabIndex        =   39
         Top             =   1890
         Width           =   450
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   495
         TabIndex        =   38
         Top             =   1485
         Width           =   330
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3420
         TabIndex        =   27
         Top             =   1035
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   495
         TabIndex        =   26
         Top             =   1065
         Width           =   540
      End
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
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
      Height          =   2280
      Left            =   45
      TabIndex        =   23
      Top             =   4410
      Width           =   4920
      Begin VB.TextBox txtCtaBanco 
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   35
         Top             =   330
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboMonedaBanco 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   705
         Visible         =   0   'False
         Width           =   1050
      End
      Begin SICMACT.TxtBuscar txtBanco 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   330
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
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
         psRaiz          =   "BANCOS"
         sTitulo         =   ""
      End
      Begin VB.TextBox txtGlosa 
         Height          =   690
         Left            =   1095
         TabIndex        =   6
         Top             =   1080
         Width           =   3600
      End
      Begin VB.TextBox txtOrdenPago 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   330
         Width           =   1155
      End
      Begin VB.CommandButton cmdDocumento 
         Height          =   350
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   315
         Width           =   475
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2565
      End
      Begin VB.Label lblCtaBanco 
         Caption         =   "Cta Banco :"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblEtqBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   2520
         TabIndex        =   33
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblBanco 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   1110
         TabIndex        =   32
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblOrdenPago 
         AutoSize        =   -1  'True
         Caption         =   "Orden Pago :"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   393
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6885
      TabIndex        =   8
      Top             =   7125
      Width           =   1000
   End
End
Attribute VB_Name = "frmCapCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PbPreIngresado As Boolean
Private pnMovNro As Long

Public nProducto As COMDConstantes.Producto
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim nMoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim bDocumento As Boolean
Dim nDocumento As COMDConstantes.tpoDoc
Dim dFechaValorizacion As Date
Dim sPersCodCMAC As String
Dim sNombreCMAC As String, sTipoCuenta As String
Dim nPersoneria As COMDConstantes.PersPersoneria
Dim pbOrdPag As Boolean
Dim sOperacion As String
Dim nSaldoCuenta As Double
Dim nSaldoRetiro As Double

'************************** MADM 20100927 ************************
Dim nEstadoT As COMDConstantes.CaptacEstado
'************************** MADM *********************************

'Variables para la impresion de la boleta de Lavado de Dinero
Dim sPersCod As String, sDocId As String, sDireccion As String
Dim sPersCodRea As String, sNombreRea As String, sDocIdRea As String, sDireccionRea As String
Dim sNombre As String

'Variable para la autorización de retiros y Cancelaciones
Dim sMovNroAut As String
'Variables para el ITF
Dim lbITFCtaExonerada As Boolean

'Variable para obtener el SubProduto de Ahorros
Dim lnTpoPrograma As Integer
Dim lsDescTpoPrograma As String
'Variable para obtener el numero de tarjeta Add By Gitu 04-05-2010
Dim sNumTarj As String
Dim lsTieneTarj As String
Dim sCuenta As String
Dim cGetValorOpe As String 'MADM 20101112
Dim nComisionVB As Double
Dim nRedondeoITF As Double 'BRGO 20110914
Dim fnRetiroPersRealiza As Boolean 'WIOR 20121005
'Funcion de Impresion de Boletas

' *** Agregado Por RIRO el 20130501, Proyecto Ahorro - Poderes ***
Dim bProcesoNuevo As Boolean
Dim strReglas As String
' *** Fin RIRO ***
Dim bInstFinanc As Boolean 'JUEZ 20140414
'JUEZ 20141017 Nuevos Parámetros **************
Dim bValidaCantRet As Boolean
Dim nParCantRetLib As Integer
Dim nComiMaxOpe As Double
Dim nParDiasVerifRegSueldo As Integer
Dim nParUltRemunBrutas As Integer
'END JUEZ *************************************

'***ADD BY MARG ERS065-2017***
Dim loVistoElectronico As frmVistoElectronico
Dim nRespuesta As Integer
'END MARG ***********************
Dim aPersonasInvol() As String, cVioFirma As String, bFirmasPendientes As Boolean, bFirmaObligatoria As Boolean, bPresente As Boolean 'ande 20170914
'end ande
Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    'Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

'********************************************
Private Function EsTitularExoneradoLavDinero() As Boolean
Dim bExito As Boolean
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim sPersCod As String
bExito = True
For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    If nRelacion = gCapRelPersTitular Then
        sPersCod = grdCliente.TextMatrix(i, 1)
        
        Exit For
    End If
Next i
EsTitularExoneradoLavDinero = bExito
End Function

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nMonto As Double
Dim sCuenta As String
Dim oDatos As COMDPersona.DCOMPersonas
Dim rsPersona As New ADODB.Recordset

For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    If nPersoneria = gPersonaNat Then
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.TitPersLavDineroDir = grdCliente.TextMatrix(i, 4)
            poLavDinero.TitPersLavDineroDoc = grdCliente.TextMatrix(i, 5)
            Exit For
        End If
    Else
        'By Capi 08072008
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.TitPersLavDineroDir = grdCliente.TextMatrix(i, 4)
            poLavDinero.TitPersLavDineroDoc = grdCliente.TextMatrix(i, 5)
            'Exit For
        End If
        '
        If nRelacion = gCapRelPersRepTitular Then
            poLavDinero.ReaPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.ReaPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.ReaPersLavDineroDir = grdCliente.TextMatrix(i, 4)
            poLavDinero.ReaPersLavDineroDoc = grdCliente.TextMatrix(i, 5)
            Exit For
        End If
    End If
Next i
nMonto = txtMonto.value
sCuenta = txtCuenta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPerscod, sNombre, sDireccion, sDocId, False, False, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
'    If txtCuenta.Prod = Producto.gCapCTS Then
'        IniciaLavDinero = frmMovLavDinero.Inicia(sPerscod, sNombre, sDireccion, sDocId, False, False, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'    Else
'        IniciaLavDinero = frmMovLavDinero.Inicia(sPerscod, sNombre, sDireccion, sDocId, False, False, nMonto, sCuenta, sOperacion, True, sTipoCuenta)
'    End If
'End If
End Sub

Sub MuestraDataTarj(ByVal sNumTar As String, ByVal sClaveTar As String, ByVal sCaption As String)
Dim lnResult As ResultVerificacionTarjeta
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim nEstado As Integer
Set clsGen = New COMDConstSistema.DCOMGeneral
        
Select Case clsGen.ValidaTarjeta(sNumTar, sClaveTar)
    Case gClaveValida
            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
            Dim rsTarj As New ADODB.Recordset
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
            If rsTarj.EOF And rsTarj.BOF Then
                MsgBox "Tarjeta no posee ninguna relación con cuentas de activas   o Tarjeta no activa.", vbInformation, "Aviso"
                Me.Caption = sCaption
                Set clsGen = Nothing
                Exit Sub
            Else
                nEstado = rsTarj("nEstado")
                If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
                    If nEstado = gCapTarjEstBloqueada Then
                        MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                    ElseIf nEstado = gCapTarjEstCancelada Then
                        MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                    End If
                    Me.Caption = sCaption
                    Set clsGen = Nothing
                    Exit Sub
                End If
                Dim rsPers As New ADODB.Recordset
                Dim sCta As String, sProducto As String, sMoneda As String
                Dim clsCuenta As UCapCuenta
                Set rsPers = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
                Set clsMant = Nothing
                If Not (rsPers.EOF And rsPers.EOF) Then
                    Do While Not rsPers.EOF
                        sCta = rsPers("cCtaCod")
                        sProducto = rsPers("Producto")
                        sMoneda = Trim(rsPers("Moneda"))
                        frmCapMantenimientoCtas.lstCuentas.AddItem sCta & space(2) & sProducto & space(2) & sMoneda
                        rsPers.MoveNext
                    Loop
                    Set clsCuenta = New UCapCuenta
                    Set clsCuenta = frmCapMantenimientoCtas.Inicia
                    If clsCuenta.sCtaCod <> "" Then
                        txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                        txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
                        txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                        txtCuenta.SetFocusCuenta
                        Call txtCuenta_KeyPress(13)
                    End If
                    Set clsCuenta = Nothing
                Else
                    MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
                End If
                rsPers.Close
                Set rsPers = Nothing
            End If
            Set rsTarj = Nothing
            Set clsMant = Nothing
    Case gTarjNoRegistrada
        'ppoa Modificacion
'        If Not WriteToLcd("Espere Por Favor") Then
'            FinalizaPinPad
'            MsgBox "No se Realizó Envío", vbInformation, "Aviso"
'            Set clsGen = Nothing
'            Exit Sub
'        End If
        MsgBox "Tarjeta no Registrada", vbInformation, "Aviso"
    Case gClaveNOValida
        'ppoa Modificacion
 '       If Not WriteToLcd("Clave Incorrecta") Then
 '           MsgBox "No se Realizó Envío", vbInformation, "Aviso"
 '           Set clsGen = Nothing
 '           Exit Sub
 '       End If
        MsgBox "Clave Incorrecta", vbInformation, "Aviso"
End Select
Set clsGen = Nothing
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'JUEZ 20141017
Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset, rsV As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
'----- MADM

'----- GITU
'Dim loVistoElectronico As New frmVistoElectronico 'COMMENT BY MARG ERS065-2017
Dim lbVistoVal As Boolean
'----- END GITU
Set loVistoElectronico = New frmVistoElectronico 'ADD BY MARG ERS065-2017
Dim rsPar As ADODB.Recordset 'JUEZ 20141017
Dim nCantOpeCta As Integer 'JUEZ 20141017

Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
Set clsCap = Nothing
If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        If bDocumento And nDocumento = TpoDocOrdenPago Then
            If Not rsCta("bOrdPag") Then
                rsCta.Close
                Set rsCta = Nothing
                MsgBox "Cuenta NO fue aperturada con ORDEN DE PAGO", vbInformation, "Aviso"
                txtCuenta.Cuenta = ""
                If gnCodOpeTarj <> 1 Then
                    txtCuenta.SetFocus
                End If
                Exit Sub
            End If
        End If
        strReglas = IIf(IsNull(rsCta!cReglas), "", rsCta!cReglas) 'Agregado por RIRO el 20130501, Proyecto Ahorro - Poderes
        'ITF INICIO
        lbITFCtaExonerada = fgITFVerificaExoneracion(sCuenta)
        fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
        'grdCliente.lbEditarFlex = flase ' Modificado Por RIRO el 20130501, Proyecto Ahorro - Poderes / se cambio de false a true
        grdCliente.lbEditarFlex = True
        If sPersCodCMAC = "" Then
            If nProducto = gCapAhorros Then
                If gbITFAsumidoAho Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.value = 0
                End If
            ElseIf nProducto = gCapPlazoFijo Then
                If gbITFAsumidoPF Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.Enabled = True
                    Me.chkITFEfectivo.value = 1
                End If
            ElseIf nProducto = gCapCTS Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
            End If
        Else
            chkITFEfectivo.Enabled = False
            If nProducto = gCapAhorros Then
                If gbITFAsumidoAho Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.value = 0
                End If
            ElseIf nProducto = gCapPlazoFijo Then
                If gbITFAsumidoPF Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.value = 1
                End If
            ElseIf nProducto = gCapCTS Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
            End If
        End If
        'ITF FIN
        
        nSaldoCuenta = rsCta("nSaldoDisp")
        nEstado = rsCta("nPrdEstado")
        nPersoneria = rsCta("nPersoneria")
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm:ss")
        nMoneda = CLng(Mid(sCuenta, 9, 1))
        ' MADM 20101115
        nEstadoT = rsCta("nPrdEstado")
        ' END MADM
        'JUEZ 20141017 ******************************************************
        lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        'Set rsPar = clsDef.GetCapParametroNew(nProducto, lnTpoPrograma)
        Set rsPar = clsDef.GetCapParametroNew(nProducto, lnTpoPrograma, sCuenta) 'APRI20190109 ERS077-2018
        If nProducto = gCapAhorros Then
            nParCantRetLib = rsPar!nCantOpeVentRet
        Else
            nParCantRetLib = rsPar!nCantOpeVentRet 'APRI20190109 ERS077-2018
            nParDiasVerifRegSueldo = rsPar!nDiasVerifUltRegSueldo
            nParUltRemunBrutas = rsPar!nUltRemunBrutas
        End If
        Set rsPar = Nothing
        'END JUEZ ***********************************************************
        
        If nMoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            txtMonto.BackColor = &HC0FFFF
            'lblMon.Caption = "S/."
            lblMon.Caption = gcPEN_SIMBOLO 'APRI20191022 SEGURENCIA CALIDAD
        Else
            sMoneda = "MONEDA EXTRANJERA"
            txtMonto.BackColor = &HC0FFC0
            lblMon.Caption = "US$"
        End If
        
        lblITF.BackColor = txtMonto.BackColor
        lblTotal.BackColor = txtMonto.BackColor
        
        Select Case nProducto
            Case gCapAhorros
                nComiMaxOpe = 0 'JUEZ 20150105
                'JUEZ 20141017 ******************************************************
                If bValidaCantRet Then
                    'Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                        nCantOpeCta = clsMant.ObtenerCantidadOperaciones(sCuenta, gCapMovRetiro, gdFecSis)
                    'Set clsMant = Nothing
                    
'                    If nCantOpeCta >= nParCantRetLib Then
'                            nCantOpeCta = nCantOpeCta + 1 'JIPR20191018 mejora
'                            nComiMaxOpe = clsMant.ObtenerCapValorComision(sCuenta, nCantOpeCta, 1, 1) 'JIPR20191018 mejora
'                            If MsgBox("La operación solicitada genera un cargo de " & IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 1, gcPEN_SIMBOLO & " ", IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 2, "$. ", "Eu.")) & Format$(nComiMaxOpe, "#,##0.00") & ", por exceso de operaciones de Retiros. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub 'JIPR20191018 mejora
'                        'nComiMaxOpe = clsDef.GetCapParametro(2155)
'                        'nComiMaxOpe = clsDef.GetCapParametro(IIf(Mid(sCuenta, 9, 1) = gMonedaNacional, 2155, 2156)) 'JUEZ 20150105
'                        'APRI20190109 ERS077-2018
'                        'END APRI
'
'
'                    End If
                                        'ADD JHCU
                    If nCantOpeCta >= nParCantRetLib Then
                        nCantOpeCta = nCantOpeCta + 1 'JIPR20191018 mejora
                        nComiMaxOpe = clsMant.ObtenerCapValorComision(sCuenta, nCantOpeCta, 1, 1) 'JIPR20191018 mejora
                        Dim clsCC As COMNCaptaGenerales.NCOMCaptaMovimiento
                        Set clsCC = New COMNCaptaGenerales.NCOMCaptaMovimiento
                        Dim bEsCtaCC As Boolean
                        Dim xcCodOpe As String
                        xcCodOpe = nOperacion
                        bEsCtaCC = clsCC.ValidaCuentaCC(sCuenta, xcCodOpe, 0, "", 0, 2)
                        Set clsCC = Nothing
                        If Not bEsCtaCC Then
                           If MsgBox("La operación solicitada genera un cargo de " & IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 1, gcPEN_SIMBOLO & " ", IIf(Mid(txtCuenta.NroCuenta, 9, 1) = 2, "$. ", "Eu.")) & Format$(nComiMaxOpe, "#,##0.00") & ", por exceso de operaciones de Retiros. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub 'JIPR20191018 mejora
                        End If
                    End If
                    'END JHCU
                End If
                'END JUEZ ***********************************************************
                If rsCta("bOrdPag") Then
                    lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
                    pbOrdPag = True
                Else
                    'JUEZ 20140429 ****************************
                    If lnTpoPrograma = 1 Then
                        lblMensaje = "AHORRO ÑAÑITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 5 Then
                        lblMensaje = "AHORRO CUENTA SOÑADA" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 6 Then
                        lblMensaje = "CAJA SUELDO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 7 Then
                        lblMensaje = "AHORRO ECOTAXI" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 8 Then
                        lblMensaje = "AHORRO CUENTA CONVENIO" & Chr$(13) & sMoneda
                    Else
                        lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
                    End If
                    'END JUEZ *********************************
                    pbOrdPag = False
                End If
                lblUltContacto = Format$(rsCta("dUltContacto"), "dd mmm yyyy hh:mm:ss")
                Me.lblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.lblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
                
                If lbITFCtaExonerada Then
                    Dim nTipoExo As String, sDescripcion As String
                    nTipoExo = fgITFTipoExoneracion(sCuenta, sDescripcion)
                    LblTituloExoneracion.Visible = True
                    lblExoneracion.Visible = True
                    lblExoneracion.Caption = sDescripcion
                End If
                
               If nProducto = gCapAhorros Then
                    Dim oCons As COMDConstantes.DCOMConstantes
                    Set oCons = New COMDConstantes.DCOMConstantes
                    lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
                    lsDescTpoPrograma = Trim(oCons.DameDescripcionConstante(2030, lnTpoPrograma))
                    Set oCons = Nothing
               End If
               
             
            Case gCapPlazoFijo
                lblUltContacto = rsCta("nPlazo")
                Me.lblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.lblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
            
            Case gCapCTS
                'APRI20190109 ERS077-2018
                nComiMaxOpe = 0
                If bValidaCantRet Then
                    nCantOpeCta = clsMant.ObtenerCantidadOperaciones(sCuenta, gCapMovRetiro, gdFecSis)
                    If nCantOpeCta >= nParCantRetLib Then
                        If MsgBox("Se ha realizado el número máximo de retiros, se cargará una comisión. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
                        nCantOpeCta = nCantOpeCta + 1
                        nComiMaxOpe = clsMant.ObtenerCapValorComision(sCuenta, nCantOpeCta, 1, 1)
                    End If
                End If
                'END APRI
                lblUltContacto = rsCta("cInstitucion")
                Dim nDiasTranscurridos As Long
                'Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
                Dim nSaldoMinRet As Double
                    
                Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
                nSaldoMinRet = clsDef.GetSaldoMinimoPersoneria(nProducto, nMoneda, nPersoneria, False)
                Set clsDef = Nothing
                If rsCta("nSaldoDisp") - rsCta("nSaldRetiro") >= nSaldoMinRet Then
                    txtMonto.Text = Format(rsCta("nSaldRetiro"), "#,##0.00")
                Else
                    txtMonto.Text = IIf((rsCta("nSaldRetiro") - nSaldoMinRet) > 0, Format(rsCta("nSaldRetiro") - nSaldoMinRet, "### ###,##0.00"), "0.00")
                End If
                nSaldoRetiro = rsCta("nSaldRetiro")
                'JUEZ 20130731 ***********************************************************
                txtMonto.Text = Format(IIf(rsCta("nSaldoDisp") - rsCta("nSaldRetiro") < 0, rsCta("nSaldoDisp"), rsCta("nSaldRetiro")), "#,##0.00")
                nSaldoRetiro = IIf(rsCta("nSaldoDisp") - rsCta("nSaldRetiro") < 0, rsCta("nSaldoDisp"), rsCta("nSaldRetiro"))
                'END JUEZ ****************************************************************
        End Select
        
        '***Agregado por ELRO el 20130722, según TI-ERS079-2013****
        If nOperacion = gAhoRetEfec Or nOperacion = gAhoRetOP Or nOperacion = gCTSRetEfec Then
             cargarMediosRetiros
        End If
         '***Fin Agregado por ELRO el 20130722, según TI-ERS079-2013
         
         ' RIRO20131210 ERS137
        If nOperacion = gAhoRetTransf Or nOperacion = gCTSRetTransf Then
            lblMedioRetiro.Visible = True
            cboMedioRetiro.Visible = True
            cargarMediosRetiros
            cboMedioRetiro.Text = "TRANSFERENCIA BANCO                                                                                                    3"
            cboMedioRetiro.Enabled = False
            CalculaComision
            'txtComision.Visible = True
        End If
        ' FIN RIRO
                    
        'Add By Gitu 23-08-2011 para cobro de comision por operacion sin tarjeta
        If sNumTarj = "" And (nOperacion = "200301" Or nOperacion = "200303" Or nOperacion = "200310" Or nOperacion = "200401" Or nOperacion = "200403" Or nOperacion = "220301" Or nOperacion = "220302" Or nOperacion = "220401" Or nOperacion = "220402") Then
            cGetValorOpe = ""
            If nMoneda = gMonedaNacional Then
                cGetValorOpe = GetMontoDescuento(2117, 1, 1)
            Else
                cGetValorOpe = GetMontoDescuento(2118, 1, 2)
            End If
            
            'If Mid(sCuenta, 6, 3) = "232" Or Mid(sCuenta, 6, 3) = "234" Then
            If (Mid(sCuenta, 6, 3) = "232" And lnTpoPrograma <> 1) Or Mid(sCuenta, 6, 3) = "234" Then 'JUEZ 20140425 Para no cobrar comisión a cuentas de ahorro ñañito
                Set rsV = clsMant.ValidaTarjetizacion(sCuenta, lsTieneTarj)
                
'                If lsTieneTarj = "SI" And rsV.RecordCount > 0 Then 'comentado por GIPO 20180803
                If rsV.RecordCount > 0 Then
                'COMMENT BY MARG ERS065-2017******************************
'                    If MsgBox("El Cliente posee tarjeta se cobrará una comision, desea continuar con la operacion?", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then
'                        Set loVistoElectronico = New frmVistoElectronico
'
'                        lbVistoVal = loVistoElectronico.Inicio(5, nOperacion)
'
'                        If Not lbVistoVal Then
'                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones, se cobrara comision por esta operacion", vbInformation, "Mensaje del Sistema"
'                            Exit Sub
'                        End If
'
'                        'loVistoElectronico.RegistraVistoElectronico (0)
'                    Else
'                        cGetValorOpe = "0.00"
'                        Exit Sub
'                    End If
                'END MARG *******************************************************
                    
                    '***ADD BY MARG ERS065-2017**************************
                    Dim tipoCta As Integer
                    
                    tipoCta = rsCta("nPrdCtaTpo")
                    If tipoCta = 0 Or tipoCta = 2 Then
                    
                        'GIPO 20180723
'                        If lsTieneTarj = "SI" Then
'                            Mensaje = "El Cliente posee tarjeta, para continuar se necesita el VB del Jefe o coordinador de Operaciones. ¿Desea Continuar?"
'                        Else
'                            Mensaje = "El Cliente NO posee tarjeta activa, por lo tanto se necesita el VB del Jefe o coordinador de Operaciones. ¿Desea Continuar?"
'                        End If
'
'                        If MsgBox(Mensaje, vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then
'
                            Dim rsCli As New ADODB.Recordset
                            Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
                            Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
                            Dim bExitoSol As Integer
                        
                            Set rsCli = clsCli.GetPersonaCuenta(sCuenta, gCapRelPersTitular)
                            nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, sCuenta, rsCli!cperscod, CStr(nOperacion))
                        
                            If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                                 MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                 Exit Sub
                            End If
                            If nRespuesta = 2 Then '2:Tiene visto aceptado
                                MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                            End If
                            If nRespuesta = 3 Then '3:Tiene visto rechazado
                               If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                    Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, sCuenta, rsCli!cperscod, CStr(nOperacion))
                                    Exit Sub
                                Else
                                    Exit Sub
                                End If
                            End If
                            If nRespuesta = 4 Then '4:Se permite registrar la solicitud
                                'GIPO 20180723 cambios en Retiros sin Tarjeta
                                Dim mensaje As String
                                If lsTieneTarj = "SI" Then
                                    mensaje = "El Cliente posee tarjeta. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                                Else
                                    mensaje = "El Cliente NO posee tarjeta activa. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                                End If
                            
                                If MsgBox(mensaje, vbInformation + vbYesNo, "Aviso") = vbYes Then
                            
                                    bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, sCuenta, rsCli!cperscod, CStr(nOperacion))
                                    If bExitoSol > 0 Then
                                        MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                    End If
                                    Exit Sub
                                Else
                                     cGetValorOpe = "0.00"
                                    Exit Sub
                                End If
                            End If
                            lbVistoVal = loVistoElectronico.Inicio(5, nOperacion)
                            If Not lbVistoVal Then
                                MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                                Exit Sub
                            End If
'                        Else
'                            cGetValorOpe = "0.00"
'                            Exit Sub
'                        End If
                    End If
                    '***END MARG*******************************************
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    'If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? se le cobrara una comision", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'comment by marg ers065-2017
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        'Set loVistoElectronico = New frmVistoElectronico 'COMMENT BY MARG ERS065-2017
                
                        lbVistoVal = loVistoElectronico.Inicio(5, nOperacion)
                    
                        If Not lbVistoVal Then
                            'MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones, se cobrara comision por esta operacion", vbInformation, "Mensaje del Sistema" 'COMMENT BY MARG ERS065-2017
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                        
                        'loVistoElectronico.RegistraVistoElectronico (0) 'COMMENT BY MARG ERS065
                    Else
                        cGetValorOpe = "0.00"
                        Exit Sub
                    End If
                Else
                    cGetValorOpe = "0.00"
                End If
            Else
                cGetValorOpe = "0.00"
            End If
            
            If cGetValorOpe <> "0.00" Then
                '***ADD BY MARG 065-2017***
                If nOperacion = 200301 Then
                    lblMonComision = Format(cGetValorOpe, "#,##0.00")
                    lblComision.Visible = False
                    lblMonComision.Visible = False
                    chkVBComision.Visible = False
                End If
                '***END MARG****************
                If nOperacion <> 200301 And Not nOperacion = 200310 Then 'add by marg ers065-2017
                    lblMonComision = Format(cGetValorOpe, "#,##0.00")
                    lblComision.Visible = True
                    lblMonComision.Visible = True
                    chkVBComision.Visible = False
                End If 'add by marg ers065-2017
            End If
            
        End If
        'End Gitu
        
        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
        sTipoCuenta = lblTipoCuenta
        nTipoCuenta = rsCta("nPrdCtaTpo")
        lblFirmas = Format$(rsCta("nFirmas"), "#0")
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        
        sPersona = ""
        
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales 'DCapMantenimiento
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
                
        'ande 20170914
        Dim i As Integer, bTieneFirma As Boolean, MsgSinFirma As String
        Dim bHaySinFirma As Boolean
        i = 0
        bHaySinFirma = False
        MsgSinFirma = "Los siguientes clientes no cuentan con firmas, se recomienda actualizarlos: " & Chr$(13)
        'end ande 20170918
                
        Do While Not rsRel.EOF
            If rsRel("cPersCod") = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                grdCliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & space(50) & Trim(rsRel("nPrdPersRelac"))
                grdCliente.TextMatrix(nRow, 4) = rsRel("Direccion") & ""
                grdCliente.TextMatrix(nRow, 5) = rsRel("ID N°")
                'grdCliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("cobligatorio")) Or rsRel("cobligatorio") = "N", "NO", IIf(rsRel("cobligatorio") = "S", "SI", "OPCIONAL"))
                
                ' ***** Agregado por RIRO el 20130501, Proyecto Ahorro - Poderes *****
                 
                If rsRel("cGrupo") <> "" Then
                    bProcesoNuevo = True
                    grdCliente.TextMatrix(nRow, 7) = rsRel("cGrupo")
                Else
                    bProcesoNuevo = False
                    grdCliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("cobligatorio")) Or rsRel("cobligatorio") = "N", "NO", IIf(rsRel("cobligatorio") = "S", "SI", "OPCIONAL"))
                End If
                ' ***** Fin RIRO *****
                  'ande 20170914
                
                If rsRel("cGrupo") <> "PJ" Or IsNull(rsRel("cGrupo")) Then
                    bFirmaObligatoria = True
                    Call MuestraFirma(rsRel("cPersCod"), , True, bTieneFirma)
                    If bTieneFirma Then
                        ReDim Preserve aPersonasInvol(i)
                        aPersonasInvol(i) = rsRel("cPersCod") & "," & UCase(PstaNombre(rsRel("Nombre"))) & "," & Trim(rsRel("nPrdPersRelac")) & ",NO"
                        i = i + 1
                        bFirmasPendientes = True
                    Else
                        MsgSinFirma = MsgSinFirma & Chr$(13) & "- " & UCase(PstaNombre(rsRel("Nombre")))
                        bHaySinFirma = True
                    End If
                End If
                'end ande
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
         If bHaySinFirma = True Then
            MsgBox MsgSinFirma, vbInformation, "Aviso" 'ande 20170918
        End If
        
        'JUEZ 20140414 ****************************************
        If nOperacion = gAhoRetEfec Or nOperacion = gAhoRetOP Or nOperacion = gAhoRetFondoFijo Or nOperacion = gAhoRetEmiChq Then
            'Dim i As Integer 'ande 20170919
            For i = 1 To grdCliente.Rows - 1
                If Trim(Left(grdCliente.TextMatrix(i, 3), 10)) = "TITULAR" Then
                    Dim oDInstFinan As COMDPersona.DCOMInstFinac
                    Set oDInstFinan = New COMDPersona.DCOMInstFinac
                    bInstFinanc = oDInstFinan.VerificaEsInstFinanc(Trim(grdCliente.TextMatrix(i, 1)))
                    Set oDInstFinan = Nothing
                    txtMonto_Change
                End If
            Next
        End If
        'END JUEZ *********************************************
        
        '********* firma
        ' COMENTADO POR RIRO
            'Set lafirma = New frmPersonaFirma
            'Set ClsPersona = New COMDPersona.DCOMPersonas
            '
            'Set Rf = ClsPersona.BuscaCliente(grdCliente.TextMatrix(nRow, 1), BusquedaCodigo)
            '
            'If Not Rf.BOF And Not Rf.EOF Then
            '   If Rf!nPersPersoneria = 1 Then
            '   Call frmPersonaFirma.Inicio(Trim(grdCliente.TextMatrix(nRow, 1)), Mid(grdCliente.TextMatrix(nRow, 1), 4, 2), False, True)
            '   End If
            'End If
            'Set Rf = Nothing

         '*******************


        ' ***** Agregado Por RIRO el 20130501, Proyecto Ahorro - Poderes *****
        
        If bProcesoNuevo Then
        
            Label5.Visible = False
            Label5.Left = 5160
            Label5.Top = 2100
            Label5.Height = 195
            Label5.Width = 1215
            
            Label8.Visible = False
            Label8.Left = 3600
            Label8.Top = 2100
            Label8.Height = 195
            Label8.Width = 765
            
            lblFirmas.Visible = False
            lblFirmas.Left = 4400
            lblFirmas.Top = 2040
            lblFirmas.Height = 300
            lblFirmas.Width = 465
            
            lblMinFirmas.Visible = False
            lblMinFirmas.Left = 6360
            lblMinFirmas.Top = 2040
            lblMinFirmas.Height = 300
            lblMinFirmas.Width = 465
            
            lblFirmas.Visible = False
            lblMinFirmas.Visible = False
            cmdVerRegla.Visible = True
            
            grdCliente.ColWidth(1) = 1300
            grdCliente.ColWidth(2) = 3600
            grdCliente.ColWidth(3) = 1200
            grdCliente.ColWidth(6) = 0
            grdCliente.ColWidth(7) = 900
            grdCliente.ColWidth(8) = 1000
                
        Else
        
            Label5.Visible = True
            Label5.Left = 5160
            Label5.Top = 2100
            Label5.Height = 195
            Label5.Width = 1215
            
            Label8.Visible = True
            Label8.Left = 3600
            Label8.Top = 2100
            Label8.Height = 195
            Label8.Width = 765
            
            lblFirmas.Visible = True
            lblFirmas.Left = 4400
            lblFirmas.Top = 2040
            lblFirmas.Height = 300
            lblFirmas.Width = 465
            
            lblMinFirmas.Visible = True
            lblMinFirmas.Left = 6360
            lblMinFirmas.Top = 2040
            lblMinFirmas.Height = 300
            lblMinFirmas.Width = 465
            
            lblFirmas.Visible = True ' Modificado por RIRO el 20130501
            lblMinFirmas.Visible = True
            
            cmdVerRegla.Visible = False
            grdCliente.ColWidth(1) = 1700
            grdCliente.ColWidth(2) = 3800
            grdCliente.ColWidth(3) = 1500
            grdCliente.ColWidth(7) = 0
            grdCliente.ColWidth(6) = 1200
            grdCliente.ColWidth(8) = 0
            
            MsgBox "Se recomienda actualizar los grupos y reglas de la cuenta", vbExclamation, "Aviso"
            Set lafirma = New frmPersonaFirma
            Set ClsPersona = New COMDPersona.DCOMPersonas
            
            Set Rf = ClsPersona.BuscaCliente(grdCliente.TextMatrix(nRow, 1), BusquedaCodigo)
            
            'ande 20170919
'            If Not Rf.BOF And Not Rf.EOF Then
'              If Rf!nPersPersoneria = 1 Then
'              Call frmPersonaFirma.Inicio(Trim(grdCliente.TextMatrix(nRow, 1)), Mid(grdCliente.TextMatrix(nRow, 1), 4, 2), False, True)
'              End If
'            End If
'            Set Rf = Nothing
'end ande
        End If
        
        ' ***** Fin RIRO *****

        rsRel.Close
        Set rsRel = Nothing
        fraCliente.Enabled = True
        fraDocumento.Enabled = True
        fraDocumentoTrans.Enabled = True
        fraMonto.Enabled = True
        
        'CTI4 ERS0112020
        If nOperacion = "200310" Then
            lblMonComision = Format(GetMontoComisionEmisionCheque(Mid(txtCuenta.NroCuenta, 9, 1)), "#,##0.00")
            Me.lblTotal.Caption = Format(lblMonComision, "#,##0.00")
            Me.cboMonedaBanco.ListIndex = Mid(txtCuenta.NroCuenta, 9, 1) - 1
            Me.cboMonedaBanco.Enabled = False
        End If
        'CTI4 End
        
        If gnCodOpeTarj <> 1 Then
            If cboDocumento.Visible Then
                cboDocumento.Enabled = True
            ElseIf txtOrdenPago.Visible Then
                txtOrdenPago.SetFocus
                
            Else
                If txtGlosa.Visible Then txtGlosa.SetFocus
            End If
        End If
        
        fraCuenta.Enabled = False
        
        cmdGrabar.Enabled = True
       
        cmdCancelar.Enabled = True
    End If
    
'    MuestraFirmas sCuenta
    
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
Set clsMant = Nothing
End Sub

Private Sub LimpiaControles()
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
txtGlosa = ""

If bDocumento Then

End If
txtMonto.BackColor = &HC0FFFF
lblITF.BackColor = txtMonto.BackColor
lblTotal.BackColor = txtMonto.BackColor

If Not nOperacion = 200310 Then 'CTI4 ERS0112020
    lblComision.Visible = False
    lblMonComision.Visible = False
    chkVBComision.Visible = False
End If
lblMonComision = "0.00"
txtMonto.value = 0

'lblMon.Caption = "S/."
lblMon.Caption = gcPEN_SIMBOLO 'APRI20191022 SEGURENCIA CALIDAD
lblMensaje = ""
cmdGrabar.Enabled = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = ""
txtCuenta.Cuenta = ""
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
lblApertura = ""
lblUltContacto = ""
lblFirmas = ""
lblTipoCuenta = ""
fraCliente.Enabled = False
fraDatos.Enabled = False
fraDocumento.Enabled = False
fraMonto.Enabled = False
fraCuenta.Enabled = True
txtOrdenPago.Text = ""

txtCuenta.SetFocus

txtBanco.Text = ""
txtCtaBanco.Text = ""
Me.cboMonedaBanco.ListIndex = -1
Me.lblBanco.Caption = ""

lblAlias.Caption = ""
lblMinFirmas.Caption = ""

nSaldoCuenta = 0
sMovNroAut = ""
If nProducto = Producto.gCapAhorros Then
        Label3.Visible = True
        'Label5.Visible = True
           
        lblAlias.Visible = True
        lblMinFirmas.Visible = True
ElseIf nProducto = Producto.gCapCTS Then
        Label3.Visible = False
        Label5.Visible = False
        lblAlias.Visible = False
        lblMinFirmas.Visible = False
End If
lblExoneracion.Visible = False
LblTituloExoneracion.Visible = False
nRedondeoITF = 0
sNumTarj = ""
'***Agregado por ELRO el 20130724, según TI-ERS079-2013****
If cboMedioRetiro.Visible Then
    cargarMediosRetiros
End If
'***Fin Agregado por ELRO el 20130724, según TI-ERS079-2013

' RIRO20131212 ERS137
If fraDocumentoTrans.Visible And fraDocumentoTrans.Enabled Then
    txtBancoTrans.Text = ""
    lblNombreBanco.Caption = ""
    cbPlazaTrans.ListIndex = 0
    ckMismoTitular.value = False
    txtTitular.Text = ""
    txtCuentaTrans.Text = ""
    txtGlosaTrans.Text = ""
    fraDocumentoTrans.Enabled = False
End If
' FIN RIRO
bInstFinanc = False 'JUEZ 20140414

Call RestoreValidacionFirma 'ande 20170914 restaurando validación de firmas
End Sub

Public Sub Inicia(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, _
        ByVal sDescOperacion As String, Optional sCodCmac As String = "", _
        Optional sNomCmac As String, Optional lcCtaCod As String, Optional pnMonto As Double, Optional lnMovNro As Long)

nProducto = nProd
nOperacion = nOpe
sPersCodCMAC = sCodCmac
sNombreCMAC = sNomCmac
sOperacion = sDescOperacion

Select Case nProd
    Case gCapAhorros
        lblEtqUltCnt = "Ult. Contacto :"
        lblUltContacto.Width = 2000
        txtCuenta.Prod = Trim(Str(gCapAhorros))
        If sPersCodCMAC = "" Then
            Me.Caption = "Captaciones - Cargo - Ahorros " & sDescOperacion
        Else
            Me.Caption = "Captaciones - Cargo - Ahorros " & sDescOperacion & " - " & sNombreCMAC
        End If
        Label3.Visible = True
        Label5.Visible = False  'RIRO SE CAMBIO A FALSE

        lblAlias.Visible = True
        lblMinFirmas.Visible = False  'RIRO CAMBIO A FALSE
        
        grdCliente.ColWidth(6) = 0 ' RIRO SE COMENTO
        
        ' SE AGREGO POR RIRO
        Label8.Visible = False
        lblFirmas.Visible = False
        
    Case gCapPlazoFijo
    
        lblEtqUltCnt = "Plazo :"
        lblUltContacto.Width = 1000
        txtCuenta.Prod = Trim(Str(gCapPlazoFijo))
        Me.Caption = "Captaciones - Cargo - Plazo Fijo " & sDescOperacion
        Label3.Visible = True
        Label5.Visible = False ' RIRO SE CAMBIO A FALSE
        
        lblAlias.Visible = True
        lblMinFirmas.Visible = False ' RIRO SE CAMBIO A FALSE
        
        grdCliente.ColWidth(6) = 1200
               
    Case gCapCTS
        lblEtqUltCnt = "Institución :"
        lblUltContacto.Width = 4250
        lblUltContacto.Left = lblUltContacto.Left - 275
        txtCuenta.Prod = Trim(Str(gCapCTS))
        If sPersCodCMAC = "" Then
            Me.Caption = "Captaciones - Cargo - CTS " & sDescOperacion
        Else
            Me.Caption = "Captaciones - Cargo - CTS " & sDescOperacion & " - " & sNombreCMAC
        End If
        fraMonto.Visible = True
        Label3.Visible = False
        Label5.Visible = False
        
        lblAlias.Visible = False
        lblMinFirmas.Visible = False
        grdCliente.ColWidth(6) = 0
End Select

'***Agregado por ELRO el 20130722, según TI-ERS079-2013****
If nOperacion = gAhoRetEfec Or nOperacion = gAhoRetOP Or nOperacion = gCTSRetEfec Then
    lblMedioRetiro.Visible = True
    cboMedioRetiro.Visible = True
End If
'***Fin Agregado por ELRO el 20130722, según TI-ERS079-2013

'Verifica si la operacion necesita algun documento
Dim clsOpe As COMDConstSistema.DCOMOperacion 'DOperacion
Dim rsDoc As ADODB.Recordset
Set clsOpe = New COMDConstSistema.DCOMOperacion
Set rsDoc = clsOpe.CargaOpeDoc(Trim(nOperacion))
Set clsOpe = Nothing

If Not (rsDoc.EOF And rsDoc.BOF) Then
    nDocumento = rsDoc("nDocTpo")
    If nDocumento = TpoDocOrdenPago Then
        lblDocumento.Visible = False
        cboDocumento.Visible = False
        cmdDocumento.Visible = False
        lblOrdenPago.Visible = True
        txtOrdenPago.Visible = True
        
        If nOperacion = gAhoRetOPCanje Then
            txtBanco.Visible = True
            lblBanco.Visible = True
            lblEtqBanco.Visible = True
            Dim clsBanco As COMNCajaGeneral.NCOMCajaCtaIF 'NCajaCtaIF
            Dim rsBanco As New ADODB.Recordset
            Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
                Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "_1%", MuestraInstituciones, "1")
            Set clsBanco = Nothing
            txtBanco.rs = rsBanco
             
        Else
            txtBanco.Visible = False
            lblBanco.Visible = False
            lblEtqBanco.Visible = False
            
        End If
    End If
    fraDocumento.Caption = Trim(rsDoc("cDocDesc"))
    bDocumento = True
Else
    ' RIRO20131212 ERS137
    'If gAhoRetTransf = nOperacion Or gAhoRetEmiChq = nOperacion Or "220302" = nOperacion Or "210202" = nOperacion Then
     If gAhoRetEmiChq = nOperacion Or "210202" = nOperacion Then
    
        Dim OCon As COMDConstantes.DCOMConstantes 'DConstante
        Set OCon = New COMDConstantes.DCOMConstantes
        Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
            Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "0[123]%", MuestraInstituciones)
        Set clsBanco = Nothing
        
        txtBanco.rs = rsBanco
        
        CargaCombo Me.cboMonedaBanco, OCon.RecuperaConstantes(gMoneda)
        Set OCon = Nothing
        
        txtBanco.Visible = True
        lblBanco.Visible = True
        lblEtqBanco.Visible = True
        
        lblDocumento.Visible = True
        cmdDocumento.Visible = True
        
        cboMonedaBanco.Visible = True
        txtCtaBanco.Visible = True
        lblCtaBanco.Visible = True
    
    ' RIRO20131212 ERS137
    ElseIf nOperacion = gAhoRetTransf Or nOperacion = gCTSRetTransf Then
                   
        'Cargando Bancos
        Dim oCons As COMDConstantes.DCOMConstantes
        Set oCons = New COMDConstantes.DCOMConstantes
        Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
        Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "0[123]%", MuestraInstituciones)
        Set clsBanco = Nothing
        
        Dim oConstante As COMDConstSistema.DCOMGeneral
        Dim rsConstante As ADODB.Recordset
        Set oConstante = New COMDConstSistema.DCOMGeneral
        txtBancoTrans.rs = rsBanco
        
        'Cargando Plaza
        Set rsConstante = oConstante.GetConstante("10032", , "'20[^0]'")
        CargaCombo cbPlazaTrans, rsConstante
        cbPlazaTrans.ListIndex = 0
        Set rsConstante = Nothing
        
        fraDocumento.Visible = False
        fraDocumentoTrans.Visible = True
                
        fraDocumentoTrans.Left = fraDocumento.Left
        fraDocumentoTrans.Top = fraDocumento.Top
        fraMonto.Height = fraDocumentoTrans.Height
        
        cboMedioRetiro.Visible = True
        lblMedioRetiro.Visible = True
        lblComisionTransf.Visible = True
        chkTransfEfectivo.Visible = True
        lblComisionTransf.Top = 660
        chkTransfEfectivo.Top = 720
        
    Else
        txtBanco.Visible = False
        lblBanco.Visible = False
        lblEtqBanco.Visible = False
    
        lblDocumento.Visible = False
        cmdDocumento.Visible = False
        
    End If
    
    lblDocumento.Visible = False
    cboDocumento.Visible = False
    cmdDocumento.Visible = False

    bDocumento = False
    lblOrdenPago.Visible = False
    txtOrdenPago.Visible = False
    
End If
rsDoc.Close
Set rsDoc = Nothing
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledProd = False
txtCuenta.EnabledCMAC = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraCliente.Enabled = False
fraDocumento.Enabled = False
fraMonto.Enabled = False

sMovNroAut = ""

PbPreIngresado = False
If lcCtaCod <> "" Then
    txtCuenta.NroCuenta = lcCtaCod
    txtMonto.value = pnMonto
    PbPreIngresado = True
    pnMovNro = lnMovNro
End If

    'madm 20101112 ----------------------------
    If nOperacion = "200341" Then
        cGetValorOpe = ""
        cGetValorOpe = GetMontoDescuento(2114, 1)
        txtMonto.value = Format(cGetValorOpe, "#,##0.00")
    End If
    
    If nOperacion = "200304" Then
        lblComision.Visible = True
        lblMonComision.Visible = True
        chkVBComision.Visible = False
        cGetValorOpe = ""
        cGetValorOpe = GetMontoDescuento(2113, 1)
        lblMonComision = Format(cGetValorOpe, "#,##0.00")
        Me.lblTotal.Caption = Format(lblMonComision, "#,##0.00")
    ElseIf nOperacion = "200310" Then 'CTI4 ERS0112020
        lblComision.Visible = True
        lblMonComision.Visible = True
        chkVBComision.Visible = True
        Me.lblTotal.Caption = Format(lblMonComision, "#,##0.00")
    
    ' RIRO20131212 ERS137
    ElseIf nOperacion = gAhoRetTransf Or nOperacion = gCTSRetTransf Then
        lblComision.Visible = True
        lblMonComision.Visible = True
        chkVBComision.Visible = False
    ' END IF
    Else
        lblComision.Visible = False
        lblMonComision.Visible = False
        chkVBComision.Visible = False
    End If
    '------------------------------------------
    'END MADM
    
    bInstFinanc = False 'JUEZ 20140414
    'JUEZ 20141017 Verificar si operación valida cantidad de retiros en mes ****
    Dim oCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set oCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    bValidaCantRet = oCapDef.ValidaCantOperaciones(nOperacion, nProducto, gCapMovRetiro)
    Set oCapDef = Nothing
    'END JUEZ ********************************************************************
    
    'ADD By GITU para el uso de las operaciones con tarjeta
    If gnCodOpeTarj = 1 And (gsOpeCod = "200301" Or gsOpeCod = "200303" Or gsOpeCod = "200310" Or gsOpeCod = "220301" Or gsOpeCod = "220302") Then
    
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, nProducto)
        If sCuenta <> "123456789" Then
            If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If sCuenta <> "" Then
                txtCuenta.NroCuenta = sCuenta
                
                'lblComision.Visible = True
                
                'txtCuenta.SetFocusCuenta
                ObtieneDatosCuenta sCuenta
                'Me.Show 1 'comment by marg ers065
                
                '***ADD BY MARG ERS 065***
                 If nOperacion = 200301 Or nOperacion = 200310 Then 'retiro efectivo/retiro Emision Cheque
                    Me.Show
                    Call Form_KeyDown(121, 0)
                    Me.Visible = False
                    Me.Show 1
                Else
                    Me.Show 1
                End If
                '***END MARG **************
            End If
        Else
            lblComision.Visible = True
            lblMonComision.Visible = True
            chkVBComision.Visible = True
           'Me.Show 1 'comment by marg ers065
           
            '***ADD BY MARG ERS 065***
             If nOperacion = 200301 Or nOperacion = 200310 Then  'retiro efectivo/retiro Emision Cheque
                Me.Show
                Call Form_KeyDown(121, 0)
                Me.Visible = False
                Me.Show 1
            Else
                Me.Show 1
            End If
            '***END MARG **************
        End If
    Else

        'Me.Show 1 'comment by marg ers065
        '***ADD BY MARG ERS 065***
        If nOperacion = 200301 Or nOperacion = 200310 Or nOperacion = 200303 Then  'retiro efectivo/retiro Emision Cheque
            Me.Show
            Call Form_KeyDown(121, 0)
            Me.Visible = False
            Me.Show
        Else
            Me.Show 1
        End If
        '***END MARG **************
    End If
    'End GITU
    

End Sub

'MADM 20101112
Private Function GetMontoDescuento(pnTipoDescuento As CaptacParametro, Optional pnCntPag As Integer = 0, _
                                   Optional pnMoneda As Integer = 1) As Double
Dim oParam As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As New ADODB.Recordset

Set oParam = New COMNCaptaGenerales.NCOMCaptaDefinicion
'Modi By Gitu 29-08-2011
    Set rsPar = oParam.GetTarifaParametro(nOperacion, pnMoneda, pnTipoDescuento)
'End Gitu
Set oParam = Nothing

If rsPar.EOF And rsPar.BOF Then
    GetMontoDescuento = 0
Else
    GetMontoDescuento = rsPar("nParValor") * pnCntPag
End If
rsPar.Close
Set rsPar = Nothing
End Function
'END MADM

Private Sub cboDocumento_Click()
If nDocumento = TpoDocNotaCargo Then
    If cboDocumento.Text = "<Nuevo>" Then
        cmdDocumento.Enabled = True
        txtMonto.Text = "0.00"
    Else
        cmdDocumento.Enabled = False
        Dim nMonto As Double
        nMonto = CDbl(Trim(Right(cboDocumento.Text, 15)))
        txtMonto.Text = Format$(nMonto, "#,##0.00")
    End If
End If
End Sub

Private Sub cboDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub


Private Sub cboMedioRetiro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMonto.Enabled = True
    txtMonto.SetFocus
End If
End Sub

Private Sub cboMonedaBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtGlosa.SetFocus
    End If
    
End Sub

Private Sub cbPlazaTrans_Click()
    If ckMismoTitular.Visible And ckMismoTitular.Enabled Then ckMismoTitular.SetFocus
    CalculaComision
End Sub

Private Sub chkVBComision_Click()

Dim nMonto As Double
nMonto = txtMonto.value
    'GITU 20110829
    If chkVBComision.value = 1 And chkITFEfectivo.value = 1 Then
        If (nOperacion = "200301" Or nOperacion = "200303" Or nOperacion = "200310" _
            Or nOperacion = "220301" Or nOperacion = "220302") Then
            lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption) + lblMonComision, "#,##0.00")
            'lblTotal.Caption = Format(nMonto + lblMonComision, "#,##0.00")
            Exit Sub
        End If
        lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
    ElseIf chkVBComision.value = 1 And chkITFEfectivo.value = 0 Then
        If (nOperacion = "200301" Or nOperacion = "200303" Or nOperacion = "200310" _
            Or nOperacion = "220301" Or nOperacion = "220302") Then
            lblTotal.Caption = Format(nMonto + lblMonComision, "#,##0.00")
            Exit Sub
        End If
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    ElseIf chkVBComision.value = 0 And chkITFEfectivo.value = 1 Then
        If (nOperacion = "200301" Or nOperacion = "200303" Or nOperacion = "200310" _
            Or nOperacion = "220301" Or nOperacion = "220302") Then
            lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
            Exit Sub
        End If
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    
    ElseIf nOperacion = gAhoRetTransf Or nOperacion = gCTSRetTransf Then 'RIRO20140210 ERS137
    Else
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    End If
    
    'lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
    'End GITU
End Sub

Private Sub ckMismoTitular_Click()
    If ckMismoTitular.value Then
        txtTitular.Text = ""
        txtTitular.Visible = False
        lblITF.Caption = "0"
    Else
        txtTitular.Visible = True
        txtTitular.SetFocus
    End If
    txtMonto_Change
End Sub

Private Sub cmdCancelar_Click()
LimpiaControles
End Sub

Private Sub cmdGrabar_Click()
'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
 'fin Comprobacion si es RFIII

'***Agregado por ELRO el 20130722, según TI-ERS079-2013****
If cboMedioRetiro.Visible = True Then
    If Trim(cboMedioRetiro) = "" Then
        MsgBox "Debe seleccionar el medio de retiro.", vbInformation, "Aviso"
        cboMedioRetiro.SetFocus
        Exit Sub
    End If
End If
'***Fin Agregado por ELRO el 20130722, según TI-ERS079-2013
'WIOR 20130301 **************************
Dim fbPersonaReaAhorros As Boolean
Dim fnCondicion As Integer
Dim nI As Integer
nI = 0
'WIOR FIN *******************************
    Dim sNroDoc As String, sCodIF As String
    Dim nMonto As Double
    Dim sCuenta As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Dim sMovNro, sMovNroTransf As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim lbProcesaAutorizacion As Boolean
    Dim lsmensaje As String
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    Dim nMontoTemp As Double  'MADM 20101115
    Dim loLavDinero As frmMovLavDinero
    Set loLavDinero = New frmMovLavDinero
    
    'ande 20170914
    Dim Msg As String, aSplitPersonasInvol() As String, iVioFirma As Integer, nCantVerFirma As Integer
    nCantVerFirma = 0
    'end ande

    Dim objPersona As COMDPersona.DCOMPersonas 'JACA 20110512
    Set objPersona = New COMDPersona.DCOMPersonas 'JACA 20110512
    
    Dim loMov As COMDMov.DCOMMov 'BRGO 20110914
    Set loMov = New COMDMov.DCOMMov 'BRGO 20110914
    
    Dim lbResultadoVisto As Boolean 'RECO 20131022 ERS141
    'Dim loVistoElectronico As frmVistoElectronico 'RECO 20131022 ERS141 'COMMENT BY MARG ERS065-2017
    'Set loVistoElectronico = New frmVistoElectronico 'RECO 20131022 ERS141 'COMMENT BY MARG ERS065-2017
    Dim loCaptaGen As COMNCaptaGenerales.NCOMCaptaGenerales  'RECO 20131022 ERS141
    Set loCaptaGen = New COMNCaptaGenerales.NCOMCaptaGenerales  'RECO 20131022 ERS141
    
     'ANDE 20180419 ERS021-2018 campaña mundialito
    Dim cperscod As String
    Dim nTitularCod As Integer, nTipoPersona As Integer, bParticipaCamp As Boolean, cTextoDatos As String
    Dim ix As Integer
    'Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    For ix = 1 To grdCliente.Rows - 1
        nTitularCod = Val(Right(Trim(grdCliente.TextMatrix(ix, 3)), 2))
        If nTitularCod = 10 Then
            'nTipoPersona = grdCliente.TextMatrix(ix, 1)
            cperscod = grdCliente.TextMatrix(ix, 1)
            nTipoPersona = loCaptaGen.getVerificarPersonaNatJur(cperscod)
        End If
    Next ix
    'end ande
    
    ' ***** Agregado Por RIRO el 20130501, Proyecto Ahorro - Poderes *****
        
    If bProcesoNuevo = True Then
    
        If ValidarReglasPersonas = False Then
            MsgBox "Las personas seleccionadas no tienen suficientes poderes para realizar el retiro", vbInformation
            Exit Sub
        End If
        
                'Validar Mayoria de Edad por RIRO 20130501, Proyecto de Ahorros - Poderes
        Dim oPersonaTemp As COMNPersona.NCOMPersona
        Dim iTemp, nMenorEdad As Integer
        
        Set oPersonaTemp = New COMNPersona.NCOMPersona
        
        For iTemp = 1 To grdCliente.Rows - 1
            If oPersonaTemp.validarPersonaMayorEdad(grdCliente.TextMatrix(iTemp, 1), Format(gdFecSis, "dd/mm/yyyy")) = False _
               And grdCliente.TextMatrix(iTemp, 7) <> "PJ" Then
                nMenorEdad = nMenorEdad + 1
            End If
        Next
        
        If nMenorEdad > 0 Then
            If MsgBox("Uno de los intervinientes en la cuenta es menor de edad, SOLO podrá disponer de los fondos con autorización del Juez " & vbNewLine & "Desea continuar?", vbInformation + vbYesNo, "AVISO") = vbYes Then
                
                Dim VistoElectronico As frmVistoElectronico
                Dim ResultadoVisto As Boolean
                Set VistoElectronico = New frmVistoElectronico
                ResultadoVisto = False
                ResultadoVisto = VistoElectronico.Inicio(3, nOperacion)
                If Not ResultadoVisto Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
        
    End If
    
    ' ***** Fin RIRO *****
    'ande 20170907
    If bFirmaObligatoria Then
        bPresente = False
        If bFirmasPendientes Then
                        
            'If lblTipoCuenta.Caption = "INDISTINTA" Then
            If lblTipoCuenta.Caption = "SOLIDARIA" Then 'APRI20190109 ERS077-2018
                Dim cUserSeleccionados As String
                Dim cPersonasNoVioFirma As String
                
                For nI = 0 To grdCliente.Rows - 1
                    If grdCliente.TextMatrix(nI, 8) = "." Then
                        cUserSeleccionados = cUserSeleccionados & "," & grdCliente.TextMatrix(nI, 1)
                    End If
                Next nI
            End If
            
            If aPersonasInvol(0) <> "" Then
                If cVioFirma = "" Then
                    'If lblTipoCuenta.Caption = "INDISTINTA" Then
                    If lblTipoCuenta.Caption = "SOLIDARIA" Then 'APRI20190109 ERS077-2018
                        For nI = 0 To UBound(aPersonasInvol)
                            aSplitPersonasInvol = Split(aPersonasInvol(nI), ",")
                            If InStr(1, cUserSeleccionados, aSplitPersonasInvol(0)) > 0 Then
                                cPersonasNoVioFirma = cPersonasNoVioFirma & "," & aSplitPersonasInvol(0)
                                bPresente = True
                            End If
                        Next nI
                    Else
                        For nI = 0 To UBound(aPersonasInvol)
                            aSplitPersonasInvol = Split(aPersonasInvol(nI), ",")
                            cPersonasNoVioFirma = cPersonasNoVioFirma & "," & aSplitPersonasInvol(0)
                            bPresente = True
                        Next nI
                    End If
                Else
                    For nI = 0 To UBound(aPersonasInvol)
                        aSplitPersonasInvol = Split(aPersonasInvol(nI), ",")
                        'If lblTipoCuenta.Caption <> "INDISTINTA" Then
                        If lblTipoCuenta.Caption = "SOLIDARIA" Then 'APRI20190109 ERS077-2018
                            If InStr(1, cVioFirma, aSplitPersonasInvol(0)) = 0 Then
                                cPersonasNoVioFirma = cPersonasNoVioFirma & "," & aSplitPersonasInvol(0)
                                bPresente = True
                            End If
                        Else
                            If InStr(1, cUserSeleccionados, aSplitPersonasInvol(0)) > 0 Then
                                If InStr(1, cVioFirma, aSplitPersonasInvol(0)) = 0 Then
                                    cPersonasNoVioFirma = cPersonasNoVioFirma & "," & aSplitPersonasInvol(0)
                                    bPresente = True
                                End If
                            End If
                        End If
                    Next nI
                End If
                If bPresente = True Then
                    
                    For nI = 0 To UBound(aPersonasInvol)
                        aSplitPersonasInvol = Split(aPersonasInvol(nI), ",")
                        If InStr(1, cPersonasNoVioFirma, aSplitPersonasInvol(0)) > 0 Then
                            nCantVerFirma = nCantVerFirma + 1
                            Msg = Msg & Chr$(13) & " - " & aSplitPersonasInvol(1)
                        End If
                    Next nI
                
                    If nCantVerFirma = 1 Then
                        Msg = "Ud. necesariamente debe verificar la firma del cliente." & Chr$(13) & "Falta ver la firma de:" & Chr$(13) & Msg
                    ElseIf nCantVerFirma >= 1 Then
                        Msg = "Ud. necesariamente debe verificar las firmas de todos los clientes." & Chr$(13) & "Falta ver las firmas de:" & Chr$(13) & Msg
                    End If
                
                    MsgBox Msg, vbOKOnly + vbInformation, "Aviso"
                    Exit Sub
                End If
            Else
                If MsgBox("La cuenta indica que necesita de firmas para el proceso de retiro, pero los clientes no cuentan con firmas en nuestros registros. ¿Desea continuar?", vbYesNo + vbExclamation, "Aviso") = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    'end ande
    
    sCuenta = txtCuenta.NroCuenta
    nMonto = txtMonto.value
    nComisionVB = CDbl(lblMonComision) 'Add By GITU 29-08-2011
    
    If nMonto = 0 Then
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        If txtMonto.Enabled Then txtMonto.SetFocus
        Exit Sub
    End If

    'Validar que no se realice retiros de Ctas de Aho Pandero/Panderito/Destino
    If nProducto = gCapAhorros Then
        'By capi 19012009 para que no permita retiros de ahorro ñañito.
        'If lnTpoPrograma = 2 Or lnTpoPrograma = 3 Or lnTpoPrograma = 4 Then
        'If lnTpoPrograma = 1 Or lnTpoPrograma = 2 Or lnTpoPrograma = 3 Or lnTpoPrograma = 4 Then
        If lnTpoPrograma = 2 Or lnTpoPrograma = 3 Or lnTpoPrograma = 4 Then
            MsgBox "No se Puede realizar un Retiro de una Cuenta de " & lsDescTpoPrograma, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    '--------------------------------------------------------------------------
    If nProducto = gCapAhorros And nOperacion < 200310 Then
        Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion, nMontoMinRet As Double
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        If pbOrdPag Then
            'nMontoMinRet = clsDef.GetMontoMinimoRetOPPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag)
            nMontoMinRet = clsDef.GetMontoMinimoRetOPPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag, sCuenta) 'APRI20190109 ERS077-2018
            If nMontoMinRet > nMonto Then
                'MsgBox "El Monto de Retiro Cta con Ord. Pago es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, "S/. ", "US$. ") & CStr(nMontoMinRet), vbOKOnly + vbInformation, "Aviso"
                MsgBox "El Monto de Retiro Cta con Ord. Pago es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, gcPEN_SIMBOLO, "US$. ") & Format(nMontoMinRet, "#,##0.00"), vbOKOnly + vbInformation, "Aviso"  'APRI20191022 SUGERENCIA CALIDAD
                txtMonto.SetFocus 'APRI20191022 SUGERENCIA CALIDAD
                txtMonto.value = nMontoMinRet 'APRI20191022 SUGERENCIA CALIDAD
                Exit Sub
            End If
        Else
            'nMontoMinRet = clsDef.GetMontoMinimoRetPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag)
            nMontoMinRet = clsDef.GetMontoMinimoRetPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag, sCuenta) 'APRI20190109 ERS077-2018
            If nMontoMinRet > nMonto Then
                'MsgBox "El Monto de Retiro es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, "S/. ", "US$. ") & CStr(nMontoMinRet), vbOKOnly + vbInformation, "Aviso"
                MsgBox "El Monto de Retiro es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, gcPEN_SIMBOLO, "US$. ") & Format(nMontoMinRet, "#,##0.00"), vbOKOnly + vbInformation, "Aviso" 'APRI20191022 SUGERENCIA CALIDAD
                txtMonto.SetFocus 'APRI20191022 SUGERENCIA CALIDAD
                txtMonto.value = nMontoMinRet 'APRI20191022 SUGERENCIA CALIDAD
                Exit Sub
            End If
        End If
        Set clsDef = Nothing
    End If
    
    'JUEZ 20130731 **********************************************
    If nProducto = gCapCTS Then
        If ValidaSaldoAutorizadoRetiroCTS = False Then Exit Sub
        
        'JUEZ 20141017 *******************************************
        Dim cDCapMov As New COMDCaptaGenerales.DCOMCaptaMovimiento
        Dim R As ADODB.Recordset
        Dim dFecha As Date
        Set R = cDCapMov.ObtenerFecUltimaActSueldosCTS(sCuenta)
        If R.BOF Or R.EOF Then
            MsgBox "No se encontraron registros de sueldos del titular de la cuenta. Debe registrar el total de los " & nParUltRemunBrutas & " últimos sueldos para proceder"
            Exit Sub
        Else
            dFecha = R!FechaAct
            If DateDiff("d", dFecha, gdFecSis) > nParDiasVerifRegSueldo Then
                MsgBox "La última actualización ha caducado. Favor actualice su registro de Sueldos"
                Exit Sub
            End If
        End If
        R.Close
        Set R = Nothing
        Set cDCapMov = Nothing
        'END JUEZ ************************************************
    End If
    'END JUEZ ***************************************************
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    clsCap.IniciaImpresora gImpresora

    'Valida documento
    If bDocumento Then
        If nDocumento = TpoDocOrdenPago Then
            sNroDoc = Trim(txtOrdenPago)
            If sNroDoc = "" Then
                MsgBox "Debe digitar un N° de Orden de Pago Válido", vbInformation, "Aviso"
                Exit Sub
            End If
            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
            Dim rsOP As New ADODB.Recordset
            Dim nEstadoOP As COMDConstantes.CaptacOrdPagoEstado
            Dim bOPExiste As Boolean
            If clsCap.EsOrdenPagoEmitida(sCuenta, CLng(sNroDoc)) Then
                Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                Set rsOP = clsMant.GetDatosOrdenPago(sCuenta, CLng(sNroDoc))  'DocRecOP
                Set clsMant = Nothing
                If Not (rsOP.EOF And rsOP.BOF) Then
                    bOPExiste = True
                    nEstadoOP = rsOP("nEstado")
                    If nEstadoOP = gCapOPEstAnulada Or nEstadoOP = gCapOPEstCobrada Or nEstadoOP = gCapOPEstExtraviada Then
                        MsgBox "Orden de Pago N° " & sNroDoc & " " & rsOP("cDescripcion"), vbInformation, "Aviso"
                        rsOP.Close
                        Set rsOP = Nothing
                        txtOrdenPago.SetFocus
                        Exit Sub
                    ElseIf rsOP("nEstado") = gCapOPEstCertifiCada Then
                        If nMonto <> rsOP("nMonto") Then
                            MsgBox "Orden de Pago Certificada. Monto No Coincide con Monto de Certificación", vbInformation, "Aviso"
                            txtMonto.Text = Format$(rsOP("nMonto"), "#,##0.00")
                            rsOP.Close
                            Set rsOP = Nothing
                            Exit Sub
                        Else
                            If nOperacion = gAhoRetOP Then
                                nOperacion = gAhoRetOPCert
                            ElseIf nOperacion = gAhoRetOPCanje Then
                                nOperacion = gAhoRetOPCertCanje
                            End If
                        End If
                    End If
                Else
                    bOPExiste = False
                End If
                rsOP.Close
                Set rsOP = Nothing
                If txtBanco.Visible Then
                    If txtBanco.Text = "" Then
                        MsgBox "Debe Seleccionar un Banco.", vbInformation, "Aviso"
                        Exit Sub
                    Else
                        sCodIF = txtBanco.Text
                    End If
                Else
                    sCodIF = ""
                End If
            Else
                MsgBox "Orden de Pago No ha sido emitida para esta cuenta", vbInformation, "Aviso"
                Exit Sub
            End If
        ElseIf nDocumento = TpoDocNotaCargo Then
            sNroDoc = Trim(Left(cboDocumento.Text, 8))
            If InStr(1, sNroDoc, "<Nuevo>", vbTextCompare) > 0 Then
                MsgBox "Debe seleccionar un documento (" & fraDocumento.Caption & ") válido para la operacion.", vbInformation, "Aviso"
                cboDocumento.SetFocus
                Exit Sub
            End If
            sCodIF = ""
        End If
    End If

'Valida Saldo de la Cuenta a Retirar
    If nOperacion <> gAhoRetOPCert And nOperacion <> gAhoRetOPCertCanje Then
        Dim nITF As Double
        nITF = 0
        If gbITFAplica And lbITFCtaExonerada = False Then
            If (nProducto = gCapAhorros And gbITFAsumidoAho = False) Or (nProducto = gCapPlazoFijo And gbITFAsumidoPF = False) Then
                If chkITFEfectivo.value = vbUnchecked Then
                    nITF = CDbl(lblITF.Caption)
                End If
            End If
        End If
    
        Dim nComixMov As Double
        Dim nComixRet As Double
        'By Capi 05032008
        Dim lbCuentaRRHH As Boolean
        Dim loCptGen As COMDCaptaGenerales.DCOMCaptaGenerales
        
        'Comentado by JACA 20111021****************************************
            'If nOperacion = gCMACOAAhoRetEfec Then
            '    nComixMov = 0
            '    nComixRet = 0
            'ElseIf nOperacion = gCMACOAAhoRetOP Then
            '    nComixMov = 0
            '    nComixRet = 0
            'Else
        'JACA END**************************************************
        
            If nProducto = gCapAhorros Or nProducto = gCapCTS Then 'APRI20190109 ERS077-2018
                '*****  VERIFICAR MAX MOVIMIENTOS Y CALCULAR COMISION  AVMM 03-06-2006 *****
                'nComixMov = Round(CalcularComisionxMaxOpeRet(), 2)
                nComixMov = Round(nComiMaxOpe, 2) 'JUEZ 20141017
                '***************************************************************************
            
                '******                 CALCULAR COMISION  AVMM-03-2006                *****
                nComixRet = Round(CalcularComisionRetOtraAge(), 2)
                '***************************************************************************
                'By Capi 05032008
                
               'JIPR20200317 SUGERENCIA GELU SE COMENTÓ
               'Set loCptGen = New COMDCaptaGenerales.DCOMCaptaGenerales
               'lbCuentaRRHH = loCptGen.ObtenerSiEsCuentaRRHH(sCuenta)
               'If lbCuentaRRHH Then
               'nComixMov = 0
               'nComixRet = 0
               'End If
                
                'Comentado by JACA 20111025****************************
                    'If lnTpoPrograma = 6 Or lnTpoPrograma = 5 Then
                    '    nComixMov = 0
                    '    nComixRet = 0
                    'End If
                    
                    'If nOperacion <> gAhoRetEfec And nOperacion <> gAhoRetOP Then 'Valida cobro de comisión solo a retiros en efectivo y con OP
                    '    nComixRet = 0
                    'End If
                'JACA END****************************************************

            Else
               'nComixMov = 0 Comentado by JACA 20111025
               nComixRet = 0
            End If
        'End If
        
        
        'MADM 20101115
       If (nOperacion = "200304" Or nOperacion = "200301" Or nOperacion = "200303" Or nOperacion = "200310" _
            Or nOperacion = "220301" Or nOperacion = "220302") Then ' Mody By Gitu 29-08-2011
            'nMontoTemp = nMonto + nITF + lblMonComision 'Comentado CTI4 ERS0112020
            nMontoTemp = nMonto + nITF + IIf(chkVBComision.value = 1 And nOperacion = "200310", 0, lblMonComision) 'CTI4 ERS0112020
       Else
            nMontoTemp = nMonto + nITF
        End If
       'END MADM
        
        If Not clsCap.ValidaSaldoCuenta(sCuenta, nMontoTemp, , nComixRet, nComixMov, , IIf(chkTransfEfectivo.value = 1, 0, lblComisionTransf.Caption)) Then
            If nOperacion = gAhoRetOPCanje Then
                'Por ahora en esta validacion no se da un trato especial esperar cambios
                '            If MsgBox("Cuenta NO posee SALDO SUFICIENTE. ¿Desea registrarla como Devuelta?", vbInformation, "Aviso") = vbYes Then
                '                'Registrar Orden de Pago Devuelta
                '
                '            End If
            ElseIf nOperacion = gAhoRetOP Then
                Dim nMaxSobregiro As Long, nSobregiro As Long
                Dim nMontoDescuento As Double, nSaldoMinimo As Double
                Dim nBloqueoParcial As Double
                Dim clsGen As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
                Dim nEstado As COMDConstantes.CaptacEstado
                Dim oMov As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
                Dim sGlosa As String
                Dim nSaldoDisponible As Double
        
                Set clsGen = New COMNCaptaGenerales.NCOMCaptaDefinicion
                nMaxSobregiro = clsGen.GetCapParametro(gNumVecesMinRechOP)
                If nMoneda = gMonedaNacional Then
                    nMontoDescuento = clsGen.GetCapParametro(gMonDctoMNRechOP)
                Else
                    nMontoDescuento = clsGen.GetCapParametro(gMonDctoMERechOP)
                End If
                                   
                Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
                Set rsOP = clsMant.GetDatosCuenta(sCuenta)
                Set clsMant = Nothing
                nSobregiro = rsOP("nSobregiro")
                nEstado = rsOP("nPrdEstado")
                nPersoneria = rsOP("nPersoneria")
                nSaldoDisponible = rsOP("nSaldoDisp")
                nBloqueoParcial = rsOP("nBloqueoParcial")
'                ' MADM 20101115
'                nEstadoT = rsOP("nPrdEstado")
'                ' END MADM
                nSaldoMinimo = clsGen.GetSaldoMinimoPersoneria(gCapAhorros, nMoneda, nPersoneria, True)
                Set clsGen = Nothing
            
                nPersoneria = rsOP("nPersoneria")
            
                'Determinar el monto de descuento Real de acuerdo a su saldo
                If nSaldoDisponible - nSaldoMinimo <= 0 Then
                    nMontoDescuento = 0
                ElseIf nSaldoDisponible - nMontoDescuento - nSaldoMinimo - nBloqueoParcial < 0 Then
                        nMontoDescuento = nSaldoDisponible - nSaldoMinimo - nBloqueoParcial
                End If
        
                Set clsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
                sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Set clsMov = Nothing
            
                Set oMov = New COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
                oMov.IniciaImpresora gImpresora
            
                If nSobregiro + 1 >= nMaxSobregiro Then 'El ultimo sobregiro, se descuenta y luego se cancela
                    MsgBox "NO POSSE SALDO SUFICIENTE..." & Chr$(13) _
                       & "La Orden de Pago N° " & sNroDoc & " ha sido Sobregirada " & nMaxSobregiro & " VECES!." & Chr$(13) _
                       & "Se procederá a bloquear la cuenta y hacer el descuento por Orden de Pago Rechazada.", vbInformation, "Aviso"
                
                    'Hacer el descuento
                    sGlosa = "OP Rechazada " & sNroDoc & ". Cuenta " & sCuenta
                
                    oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nMoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF, psRegla:=ObtenerRegla
                    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"

                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    clsMant.ActualizaSobregiro sCuenta, nSobregiro + 1, sMovNro, nEstado, False
                    clsMant.BloqueCuentaTotal sCuenta, gCapMotBlqTotOrdenPagoRechazada, sGlosa, sMovNro
                    Set clsMant = Nothing
                
                'COMENTADO HASTA DEFINIR PROCESO 0609-2006-AVMM
    '                oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nMoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF

    '                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
    '                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
    '                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
    '
    '                Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    '                clsMant.ActualizaSobregiro sCuenta, nMaxSobregiro, sMovNro, nEstado, True
    '                Set clsMant = Nothing
    '                'Cancelar la cuenta
    '                oMov.GetSaldoCancelacion sCuenta, gdFecSis, gsCodAge, 0  'x Midificar GetSaldoCancelacion Raul
    '
    '                Set clsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
    '                    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    '                Set clsMov = Nothing
    '
    '                oMov.CapCancelaCuentaAho sCuenta, sMovNro, sGlosa, gAhoCancSobregiroOP, gCapEstCancelada, gsNomAge, sLpt, , gsCodCMAC, False, , , , , lsmensaje, lsBoleta, lsBoletaITF
    '
    '                Set oMov = Nothing
    '
    '                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
    '                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
    '                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
                                
                    cmdCancelar_Click
                    Exit Sub
            
                ElseIf nSobregiro + 1 = nMaxSobregiro - 1 Then 'Una menos que la ultima se descuenta y se bloquea
                    MsgBox "NO POSEE SALDO SUFICIENTE" & Chr$(13) _
                      & "La Cuenta " & sCuenta & " ha sido Sobregirada " & nSobregiro + 1 & " VECES!." & Chr$(13) _
                      & "Se procederá a bloquear la cuenta y hacer el descuento por Orden de Pago Rechazada.", vbInformation, "Aviso"
                    'Hacer el descuento
                    sGlosa = "OP Rechazada " & sNroDoc & ". Cuenta " & sCuenta

                    oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nMoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF, psRegla:=ObtenerRegla
                    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"

                    'Se actualiza el sobregiro y se bloquea la cuenta
                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    clsMant.ActualizaSobregiro sCuenta, nSobregiro + 1, sMovNro, nEstado, False
                    clsMant.BloqueCuentaTotal sCuenta, gCapMotBlqTotOrdenPagoRechazada, sGlosa, sMovNro
                    Set clsMant = Nothing
                    'MIOL 20120917, SEGUN RQ12272 ********************************************
                    Dim oDatosCliente As COMNPersona.NCOMPersona
                    Set oDatosCliente = New COMNPersona.NCOMPersona
                    Dim rsDatosCliente As Recordset
                    Set rsDatosCliente = oDatosCliente.MostrarClientexCuenta(sCuenta)
                    Call oDatosCliente.insClienteBloqueadoxCuenta(sMovNro, rsDatosCliente!cperscod, rsDatosCliente!Nombre, rsDatosCliente!cCtaCod, 1, Format(gdFecSis, "yyyymmdd"))
                    Set rsDatosCliente = Nothing
                    'END MIOL ****************************************************************
                    
                    cmdCancelar_Click
                    Exit Sub
                Else 'Se descuenta noma
                
                    MsgBox "NO POSEE SALDO SUFICIENTE" & Chr$(13) _
                       & "La Cuenta " & sCuenta & " ha sido Sobregirada " & nSobregiro + 1 & " VECES!." & Chr$(13) _
                       & "Se procederá a hacer el descuento por Orden de Pago Rechazada.", vbInformation, "Aviso"
            
                    'Hacer el descuento
                    'oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nMoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF

                    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
                
                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    clsMant.ActualizaSobregiro sCuenta, nSobregiro + 1, sMovNro, nEstado, False
                    Set clsMant = Nothing
                
                    cmdCancelar_Click
                    Exit Sub
                End If
                Set oMov = Nothing
            
            Else
                MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
                If txtMonto.Enabled Then txtMonto.SetFocus
                Exit Sub
            End If
        End If
    End If

    'Valida que la transaccion no se pueda realizar porque la cuenta no posee firmas
    If Not gbRetiroSinFirma Then
        If Not clsCap.CtaConFirmas(txtCuenta.NroCuenta) Then
            MsgBox "No puede retirar, porque la cuenta no cuenta con las firmas de las personas relacionadas a ella.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

    'Valida la operacion de Retiro por Transferencias
    If nOperacion = gAhoRetTransf Or nOperacion = gCTSRetTransf Then
        ' RIRO20131212 ERS137
        'If Trim(txtBanco.Text) = "" Then
        If Trim(txtBancoTrans.Text) = "" Then
            MsgBox "Debe seleccionar el Banco a Transferir", vbInformation, "Aviso"
            If txtBancoTrans.Enabled Then txtBancoTrans.SetFocus
            Exit Sub
        End If
        
        'RIRO20131212 ERS137 - Comentado
        
        'If cboMonedaBanco.Text = "" Then
        '    MsgBox "Debe seleccionar la Moneda de la Cuenta de Banco a Transferir", vbInformation, "Aviso"
        '    cboMonedaBanco.SetFocus
        '    Exit Sub
        'End If
        
        ' RIRO20131212 ERS137
        'If Trim(txtCtaBanco.Text) = "" Then
        If Trim(txtBancoTrans.Text) = "" Then
            MsgBox "Debe digitar la Cuenta de Banco a Transferir", vbInformation, "Aviso"
            txtCtaBanco.SetFocus
            Exit Sub
        End If
    End If

    '----------- Verificar Autorizacion -- AVMM -- 18/04/2004 -------
   
    
    
    
    If VerificarAutorizacion = False Then Exit Sub
    '----------------------------------------------------------------

    '----------- Verificar Creditos -- CAAU -- 15-12-2006----
    
    If nProducto = gCapAhorros Then
        Dim K As Integer
        Dim clsCapD As COMDCaptaGenerales.DCOMCaptaGenerales
        Set clsCapD = New COMDCaptaGenerales.DCOMCaptaGenerales
        Dim oCred As COMDCredito.DCOMCredito
        Dim lsMsgCred As String
        Set oCred = New COMDCredito.DCOMCredito
    
        For K = 1 To Me.grdCliente.Rows - 1
            '10= Titular
            If Right(Me.grdCliente.TextMatrix(K, 3), 2) = "10" Then
                If oCred.VerificarClienteCredMorosos(Me.grdCliente.TextMatrix(K, 1)) Then
                    lsMsgCred = "Cliente posee pagos de Creditos Pendientes..."
                End If
            End If
        Next K
        Set clsCapD = Nothing
    Else
        lsMsgCred = ""
    End If
    '---------------------------------------------------------
    'JACA 20110512 *****VERIFICA SI LAS PERSONAS CUENTAN CON OCUPACION E INGRESO PROMEDIO
        Dim rsPersVerifica As Recordset
        Dim i As Integer
        Set rsPersVerifica = New Recordset
        For i = 1 To grdCliente.Rows - 1
            Call VerSiClienteActualizoAutorizoSusDatos(grdCliente.TextMatrix(i, 1), nOperacion) 'FRHU ERS077-2015 20151204
            Set rsPersVerifica = objPersona.ObtenerDatosPersona(Me.grdCliente.TextMatrix(i, 1))
            If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
                If MsgBox("Necesita Registrar la Ocupacion e Ingreso Promedio de: " + Me.grdCliente.TextMatrix(i, 2), vbYesNo) = vbYes Then
                    'frmPersona.Inicio Me.grdCliente.TextMatrix(i, 1), PersonaActualiza
                    frmPersOcupIngreProm.Inicio Me.grdCliente.TextMatrix(i, 1), Me.grdCliente.TextMatrix(i, 2), rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
                End If
            End If
        Next i
    'JACA END***************************************************************************
    
    'Add By GITU 07/10/2011 para corregir la operacion sin tarjeta para que no cobre comision en un Retiro normal
    If Not lblComision.Visible Then
        nComisionVB = 0
    End If
    'End Gitu
    'WIOR 20121029 Clientes Observados ******************************************
    If nOperacion = gAhoRetEfec Or nOperacion = gAhoRetOP Or nOperacion = gCTSRetEfec Then
        Dim oDPersona As COMDPersona.DCOMPersona
        Dim rsPersona As ADODB.Recordset
        Dim sCodPersona As String
        Dim Cont As Integer
        
        Set oDPersona = New COMDPersona.DCOMPersona
        
        For Cont = 0 To grdCliente.Rows - 2
            If Trim(Right(grdCliente.TextMatrix(Cont + 1, 3), 5)) = gCapRelPersTitular Then
                sCodPersona = Trim(grdCliente.TextMatrix(Cont + 1, 1))
                Set rsPersona = oDPersona.ObtenerUltimaVisita(sCodPersona)
                If rsPersona.RecordCount > 0 Then
                    If Not (rsPersona.EOF And rsPersona.BOF) Then
                        If Trim(rsPersona!sUsual) = "3" Then
                            MsgBox Trim(grdCliente.TextMatrix(Cont + 1, 2)) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                            Call frmPersona.Inicio(sCodPersona, PersonaActualiza)
                        End If
                    End If
                End If
                Set rsPersona = Nothing
            End If
        Next Cont
    End If
    'WIOR FIN ***************************************************************
    
      'AMDO 20130702 TI-ERS063-2013 ****************************************************
        If nOperacion = gAhoRetEfec Or nOperacion = gCTSRetEfec Then
            Dim oDPersonaAct As COMDPersona.DCOMPersona
            Dim conta As Integer
            Dim sPersCod As String
            
            Set oDPersonaAct = New COMDPersona.DCOMPersona
            For conta = 0 To grdCliente.Rows - 2
            sPersCod = Trim(grdCliente.TextMatrix(conta + 1, 1))
                            If oDPersonaAct.VerificaExisteSolicitudDatos(sPersCod) Then
                                MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & grdCliente.TextMatrix(conta + 1, 2)) & "." & Chr(10), vbInformation, "Aviso"
                                Call frmActInfContacto.Inicio(sPersCod)
                            End If
            Next conta
        End If
    'AMDO FIN ********************************************************************************
    
    If MsgBox(lsMsgCred & "¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Dim sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
        Dim nSaldo As Double, nPorcDisp As Double
        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
        Dim nMontoLavDinero As Double, nTC As Double

        'Realiza la Validación para el Lavado de Dinero
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
            Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
            If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
                Set clsExo = Nothing
                sPersLavDinero = ""
                nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
                Set clsLav = Nothing

                If nMoneda = gMonedaNacional Then
                    Dim clsTC As COMDConstSistema.NCOMTipoCambio 'nTipoCambio
                    Set clsTC = New COMDConstSistema.NCOMTipoCambio
                    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set clsTC = Nothing
                Else
                    nTC = 1
                End If
                If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                    'By Capi 1402208
                    Call IniciaLavDinero(loLavDinero)
                    'ALPA
                    'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, sTipoCuenta, , , , , nmoneda)
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, sTipoCuenta, , , , , nMoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                    'End
                End If
            Else
                Set clsExo = Nothing
            End If
        'Else
        '    Set clsLav = Nothing
        'End If
        'WIOR 20130301 Personas Sujetas a Procedimiento Reforzado *************************************
        fbPersonaReaAhorros = False
        If (loLavDinero.OrdPersLavDinero = "Exit") _
                And (nOperacion = gAhoRetEfec Or nOperacion = gAhoRetOP Or nOperacion = gAhoRetEmiChq Or nOperacion = gAhoRetTransf _
                 Or nOperacion = gCTSRetEfec Or nOperacion = gCTSRetTransf) Then
                
                Dim oPersonaSPR As UPersona_Cli
                Dim oPersonaU As COMDPersona.UCOMPersona
                Dim nTipoConBN As Integer
                Dim sConPersona As String
                Dim pbClienteReforzado As Boolean
                Dim rsAgeParam As Recordset
                Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
                Dim lnMontoX As Double, lnTC As Double
                Dim ObjTc As COMDConstSistema.NCOMTipoCambio
                
                
                Set oPersonaU = New COMDPersona.UCOMPersona
                Set oPersonaSPR = New UPersona_Cli
                
                fbPersonaReaAhorros = False
                pbClienteReforzado = False
                fnCondicion = 0
                
                For nI = 0 To grdCliente.Rows - 2
                    oPersonaSPR.RecuperaPersona Trim(grdCliente.TextMatrix(nI + 1, 1))
                                        
                    If oPersonaSPR.Personeria = 1 Then
                        If oPersonaSPR.Nacionalidad <> "04028" Then
                            sConPersona = "Extranjera"
                            fnCondicion = 1
                            pbClienteReforzado = True
                            Exit For
                        ElseIf oPersonaSPR.Residencia <> 1 Then
                            sConPersona = "No Residente"
                            fnCondicion = 2
                            pbClienteReforzado = True
                            Exit For
                        ElseIf oPersonaSPR.RPeps = 1 Then
                            sConPersona = "PEPS"
                            fnCondicion = 4
                            pbClienteReforzado = True
                            Exit For
                        ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                            If nTipoConBN = 1 Or nTipoConBN = 3 Then
                                sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                                fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                                pbClienteReforzado = True
                                Exit For
                            End If
                        End If
                    Else
                        If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                            If nTipoConBN = 1 Or nTipoConBN = 3 Then
                                sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                                fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                                pbClienteReforzado = True
                                Exit For
                            End If
                        End If
                    End If
                Next nI
                
                If pbClienteReforzado Then
                    MsgBox "El Cliente: " & Trim(grdCliente.TextMatrix(nI + 1, 2)) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                    frmPersRealizaOpeGeneral.Inicia sOperacion & " (Persona " & sConPersona & ")", nOperacion
                    fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                    
                    If Not fbPersonaReaAhorros Then
                        MsgBox "Se va a proceder a Anular la Operacion ", vbInformation, "Aviso"
                        cmdGrabar.Enabled = True
                        Exit Sub
                    End If
                Else
                    fnCondicion = 0
                    lnMontoX = nMonto
                    pbClienteReforzado = False
                    
                    Set ObjTc = New COMDConstSistema.NCOMTipoCambio
                    lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set ObjTc = Nothing
                
                
                    Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    Set rsAgeParam = objCap.getCapAbonoAgeParam(gsCodAge)
                    Set objCap = Nothing
                    
                    If Mid(Trim(txtCuenta.NroCuenta), 9, 1) = 1 Then
                        lnMontoX = Round(lnMontoX / lnTC, 2)
                    End If
                
                    If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                        If lnMontoX >= rsAgeParam!nMontoMin And lnMontoX <= rsAgeParam!nMontoMax Then
                            frmPersRealizaOpeGeneral.Inicia sOperacion, nOperacion
                            fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                            If Not fbPersonaReaAhorros Then
                                MsgBox "Se va a proceder a Anular la Operacion", vbInformation, "Aviso"
                                cmdGrabar.Enabled = True
                                Exit Sub
                            End If
                        End If
                    End If
                    
                End If
        End If
        'WIOR FIN ***************************************************************
        'WIOR 20130301 COMENTO INICIO **********************************************************************
        ''WIOR 20121005 Retiro Orden Pago y Efectivo(Personas Juridicas)**********************************
        'If Trim(loLavDinero.OrdPersLavDinero) = "Exit" And (nOperacion = gAhoRetOP Or nOperacion = gAhoRetEfec) Then
        '    Dim ContPersoneria As Integer
        '    ContPersoneria = 0
        '    If nOperacion = gAhoRetEfec Then
        '        For Cont = 0 To grdCliente.Rows - 2
        '            If Trim(Right(grdCliente.TextMatrix(Cont + 1, 3), 5)) = gColRelPersTitularCred Then
        '            sCodPersona = Trim(grdCliente.TextMatrix(Cont + 1, 1))
        '                Call oDPersona.RecuperaPersona(sCodPersona)
        '                If CInt(oDPersona.Personeria) > 1 Then
        '                    ContPersoneria = ContPersoneria + 1
        '                End If
        '            End If
        '        Next Cont
        '        If ContPersoneria > 0 Then
        '            frmPersRealizaOperacion.Inicia "Retiro en Efectivo", gPersRealizaRetiroEfec
        '            fnRetiroPersRealiza = frmPersRealizaOperacion.PersRegistrar
        '        Else
        '            fnRetiroPersRealiza = False
        '            ContPersoneria = 0
        '        End If
        '    ElseIf nOperacion = gAhoRetOP Then
        '        frmPersRealizaOperacion.Inicia "Orden de Pago", gPersRealizaRetiroOrdPag
        '        fnRetiroPersRealiza = frmPersRealizaOperacion.PersRegistrar
        '    End If
        '
        '    If nOperacion = gAhoRetEfec And ContPersoneria > 0 Then
        '        If Not fnRetiroPersRealiza Then
        '            MsgBox "Se va a proceder a Anular el Retiro de la Cuenta."
        '            Exit Sub
        '        End If
        '    Else
        '        If nOperacion = gAhoRetOP Then
        '            If Not fnRetiroPersRealiza Then
        '                MsgBox "Se va a proceder a Anular el Retiro de la Cuenta."
        '                Exit Sub
        '            End If
        '        End If
        '    End If
        'End If
        ''WIOR FIN ***************************************************************************************
        'WIOR 20130301 COMENTO FIN ******************************************************************************
        Set clsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Sleep (1000)
        sMovNroTransf = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set clsMov = Nothing
    '    On Error GoTo ErrGraba
        '*************RECO 20131022 ERS141**********************************
        If loCaptaGen.RecuperaSubTipoProducto(sCuenta)!nTpoPrograma = 7 Then
            lbResultadoVisto = loVistoElectronico.Inicio(3, "200301")
            If Not lbResultadoVisto Then
                cmdGrabar.Enabled = True
                Exit Sub
            End If
        End If
        '**************END RECO*************************************
        'ANDE 20180419 ERS021-2018 participanción en campaña mundialto
        Dim nPuntosRef As Integer, nCondicion As Integer, nPTotalAcumulado As Integer
        If nOperacion = gAhoRetEfec Then
            Dim nOpeTipo As Integer
            nOpeTipo = 3 '3:RETIRO
            nMoneda = Mid(sCuenta, 9, 1)
            If nMoneda = gMonedaNacional Then
                'Dim oCampanhaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
                Call loCaptaGen.VerificarParticipacionCampMundial(cperscod, sCuenta, nOperacion, nOpeTipo, nMoneda, nMonto, nTipoPersona, bParticipaCamp, sMovNro, , gdFecSis, lnTpoPrograma, txtCuenta.Age, nPuntosRef, nCondicion, , nPTotalAcumulado)
                If bParticipaCamp Then
                    cTextoDatos = "#" & IIf(bParticipaCamp, "1", "0") & "." & CStr(nPuntosRef) & "$" & CStr(nCondicion) & "_" & CStr(nPTotalAcumulado) & "&"
                    lsmensaje = cTextoDatos
                End If
            End If
        End If
        'end ande
        
        Select Case nProducto

            Case gCapAhorros
                If bDocumento Then
                    If nDocumento = TpoDocOrdenPago Then
                        If nOperacion = gAhoRetOPCert Or nOperacion = gAhoRetOPCertCanje Then
                            nSaldo = clsCap.CapCargoAhoOPCertifcada(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, sNroDoc, sCodIF, , , gsNomAge, sLpt, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, lsBoleta, , , gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, ObtenerRegla)
                        Else
                            '***Agregado por ELRO el 20130723, según TI-ERS079-2013****
                            'nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, sCodIF, bOPExiste, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , , , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                            If nOperacion = gAhoRetEfec Or nOperacion = gAhoRetOP Then
                                nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, sCodIF, bOPExiste, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , , , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, , , , CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla)
                            Else
                                nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, sCodIF, bOPExiste, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , , , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, psRegla:=ObtenerRegla)
                            End If
                            '***Fin Agregado por ELRO el 20130723, según TI-ERS079-2013
                            
                        End If
                    ElseIf nDocumento = TpoDocNotaCargo Then
                        nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, sCodIF, bOPExiste, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , , , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, psRegla:=ObtenerRegla)
                    End If
                Else
                    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    If nOperacion = gAhoRetEmiChq Then 'CTI4 ERS0112020
                        'nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , Right(Me.txtBanco.Text, 13), , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , Me.txtCtaBanco.Text, , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, sNumTarj, nComisionVB, IIf(Me.chkVBComision.value = 0, gComisionEmisionChequeCargoCta, gComisionEmisionChequeEfectivo), psRegla:=ObtenerRegla)

                        nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , Right(Me.txtBanco.Text, 13), , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , Me.txtCtaBanco.Text, , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, sNumTarj, , , , ObtenerRegla, , , , , , , , , nComisionVB, IIf(Me.chkVBComision.value = 0, gComisionEmisionChequeCargoCta, gComisionEmisionChequeEfectivo), sMovNroTransf)

                    Else
                        If Me.cboMonedaBanco.Text <> "" Then
                            nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , Right(Me.txtBanco.Text, 13), , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , Me.txtCtaBanco.Text, CInt(Right(Me.cboMonedaBanco.Text, 3)), gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, psRegla:=ObtenerRegla)
                            
                        Else
                            '***Agregado por ELRO el 20130723, según TI-ERS079-2013****
                            If nOperacion = gAhoRetEfec Or nOperacion = gAhoRetOP Then
                                nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , Right(Me.txtBanco.Text, 13), , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , Me.txtCtaBanco.Text, , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, sNumTarj, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"), CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla)
                            '*** RIRO20131212 ERS137 ***
                            ElseIf nOperacion = gAhoRetTransf Then
                                nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, getGlosa, , , Right(Me.txtBancoTrans.Text, 13), , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , Me.txtCuentaTrans.Text, , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, , , , , ObtenerRegla, CDbl(lblComisionTransf.Caption), sMovNroTransf, getTitular, chkTransfEfectivo.value)
                            '*** FIN RIRO ***
                            Else
                                nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , Right(Me.txtBanco.Text, 13), , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , Me.txtCtaBanco.Text, , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, sNumTarj, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"), psRegla:=ObtenerRegla)
                            End If
                            '***Fin Agregado por ELRO el 20130723, según TI-ERS079-2013
                        End If
                        'If nSaldo = -69 Then frmCapAutorizacion.Inicio
                    End If
                End If

                '-------------------------------------------------
                
                
                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"

            Case gCapCTS
                If nOperacion = gCTSRetEfec Or nOperacion = gCMACOACTSRetEfec Then
                    '***Modificado por ELRO el 20130724, según TI-ERS079-2013****
                    'nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"))
                    If nOperacion = gCTSRetEfec Then
                        'nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla)
                        'nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, , , , , , , sNumTarj) 'RECO20141127
                        nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, , , , , , , sNumTarj, , , , nComixRet, nComixMov) 'APRI20190109 ERS077-2018
                    Else
                        nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), psRegla:=ObtenerRegla)
                    End If
                    '***Fin Modificado por ELRO el 20130724, según TI-ERS079-2013
                ElseIf nOperacion = gCTSRetTransf Then
                    'nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, getGlosa, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, Right(Me.txtBancoTrans.Text, 13), Me.txtCuentaTrans.Text, , 0, 0, 0, 0, 0, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), , ObtenerRegla, lblComisionTransf.Caption, sMovNroTransf, getTitular, chkTransfEfectivo.value)     'RIRO20131212 ERS137 - Agregado
                    nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, getGlosa, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, Right(Me.txtBancoTrans.Text, 13), Me.txtCuentaTrans.Text, , 0, 0, 0, 0, 0, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), , ObtenerRegla, lblComisionTransf.Caption, sMovNroTransf, getTitular, chkTransfEfectivo.value, , , , , , , nComixRet, nComixMov) 'APRI20190109 ERS077-2018
                    'nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, Right(Me.txtBanco.Text, 13), Me.txtCtaBanco.Text, CInt(Right(Me.cboMonedaBanco.Text, 3)), 0, 0, 0, 0, 0, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"), psRegla:=ObtenerRegla) 'RIRO20131212 ERS137 - Comentado
    '            ElseIf nOperacion = "220303" Then
    '                nMonto = CDbl(Me.TxtMonto2.Text)
    '                nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , CDbl(Val(Me.lblSalIntaD.Caption)), CDbl(Val(Me.lblIntIntaD.Caption)), CDbl(Val(Me.lblIntDisD.Caption)))
                End If
                
                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
        End Select
        
        'APRI20190109 ERS077-2018
        If (nOperacion = gAhoRetEfec Or nOperacion = gCTSRetEfec) And sNumTarj <> "" Then
            Dim cDCap As New COMDCaptaGenerales.DCOMCaptaGenerales
            Dim rs As ADODB.Recordset
            Dim bComunica As Boolean
            Dim cFechaApli As String
            nI = 0
            For nI = 1 To grdCliente.Rows - 1
                Set rs = cDCap.AplicaComunicacionCaptaciones(grdCliente.TextMatrix(nI, 1), "")
                bComunica = rs!bComun
                If Not bComunica Then
                   Exit For
                End If
            Next nI
             
            If Not bComunica Then
                Set rs = cDCap.ObtenerFechaAplicacionTarifario
                cFechaApli = rs!dFecApli
            
                MsgBox "En Cumplimiento al Reglamento de Gestión de Conducta de Mercado del Sistema Financiero Res. SBS N° 3274-2017 y sus modificatorias, tenemos el agrado de comunicarle que a partir del " & cFechaApli & " entró en vigencia las nuevas condiciones de nuestros productos pasivos.", vbInformation, "COMUNICACIÓN POR CAMBIOS CONTRACTUALES"
                        
                ImpreCartaNotificacionTarifario "", sCuenta, gdFecSis
                
            End If
            rs.Close
            Set rs = Nothing
            Set cDCap = Nothing
        End If
        'END APRI
        
        ' MADM 20101115
        If nOperacion = "200304" Then
            Call CargoAutomatico(sCuenta, 1) 'MADM 20101115
        End If
        ' MADM 20101115
        
        If Me.chkVBComision.value = 1 And Not nOperacion = gAhoRetEmiChq Then
            Call CobroComisionVBEfec(sCuenta, nComisionVB, "300628", gnMovNro)  'GITU 03-10-2011
        End If
        
        If gnMovNro > 0 Then
            'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
             Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
        End If
        'JACA 20110512***********************************************************
        
        '*****BRGO 20110914 *****************************************************
        If gbITFAplica = True And CCur(lblITF.Caption) > 0 Then
            Call loMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption)) 'BRGO 20110914
        End If
        Set loMov = Nothing
        '*** End BRGO *****************
        
                'Dim objPersona As COMDPersona.DCOMPersonas
                Dim rsPersOcu As Recordset
                Dim nAcumulado As Currency
                Dim nMontoPersOcupacion As Currency
                 
                Set rsPersOcu = New Recordset
                'Set objPersona = New COMDPersona.DCOMPersonas
                
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
                                
                Set rsPersOcu = objPersona.ObtenerDatosPersona(Me.grdCliente.TextMatrix(1, 1))
                nAcumulado = objPersona.ObtenerPersAcumuladoMontoOpe(nTC, Mid(Format(gdFecSis, "yyyymmdd"), 1, 6), rsPersOcu!cperscod)
                nMontoPersOcupacion = objPersona.ObtenerParamPersAgeOcupacionMonto(Mid(rsPersOcu!cperscod, 4, 2), CInt(Mid(rsPersOcu!cPersCIIU, 2, 2)))
            
                If nAcumulado >= nMontoPersOcupacion Then
                    If Not objPersona.ObtenerPersonaAgeOcupDatos_Verificar(rsPersOcu!cperscod, gdFecSis) Then
                        objPersona.insertarPersonaAgeOcupacionDatos gnMovNro, rsPersOcu!cperscod, IIf(nMoneda = gMonedaNacional, lblTotal, lblTotal * nTC), nAcumulado, gdFecSis, sMovNro
                    End If
                End If
               
        
    'JACA END*****************************************************************
    'WIOR 20130301 COMENTO INICIO **********************************************
    ''WIOR 20121005 ************************************
    'If fnRetiroPersRealiza Then
    '    frmPersRealizaOperacion.InsertaPersonaRealizaOperacion gnMovNro, sCuenta, frmPersRealizaOperacion.PersTipoCliente, _
    '    frmPersRealizaOperacion.PersCod, frmPersRealizaOperacion.PersTipoDOI, frmPersRealizaOperacion.PersDOI, frmPersRealizaOperacion.PersNombre, _
    '    frmPersRealizaOperacion.TipoOperacion
    '
    '    fnRetiroPersRealiza = False
    'End If
    ''WIOR FIN *****************************************
    'WIOR 20130301 COMENTO FIN ****************************************************
    
    'WIOR 20130301 ************************************************************
    If fbPersonaReaAhorros And gnMovNro > 0 Then
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuenta), fnCondicion
        fbPersonaReaAhorros = False
    End If
    'WIOR FIN *****************************************************************
    
    '***ADD BY MARG ERS065-2017***
    Dim oMovOperacion As COMDMov.DCOMMov
    Dim nMovNroOperacion As Long
    Dim rsCli As New ADODB.Recordset
    Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set oMovOperacion = New COMDMov.DCOMMov
    nMovNroOperacion = oMovOperacion.GetnMovNro(sMovNro)
    
    loVistoElectronico.RegistraVistoElectronico nMovNroOperacion, , gsCodUser, nMovNroOperacion
    
    If nRespuesta = 2 Then
        Set rsCli = clsCli.GetPersonaCuenta(sCuenta, gCapRelPersTitular)
        oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, sCuenta, rsCli!cperscod, nMovNroOperacion, CStr(nOperacion)
    End If
    Set oMovOperacion = Nothing
    nRespuesta = 0
    'INICIO JHCU ENCUESTA 16-10-2019
    Dim CodOpeEnc As Integer
    'CodOpeEnc = nOperacion
    Encuestas gsCodUser, gsCodAge, "ERS0292019", nOperacion
    'FIN
    '***END MARG********************
    
        Set clsLav = Nothing
        Set clsCap = Nothing
        Set loLavDinero = Nothing
        sNumTarj = ""
        gVarPublicas.LimpiaVarLavDinero
        cmdCancelar_Click
    End If
    sNumTarj = ""
    Unload Me
    Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub cmdMostrarFirma_Click()
With grdCliente
    If .TextMatrix(.row, 1) = "" Then Exit Sub
    Call frmPersonaFirma.Inicio(Trim(.TextMatrix(.row, 1)), Trim(txtCuenta.Age), True)
    
     'ande 20170914
    If bFirmasPendientes = True Then
        If gbTieneFirma = True Then
            Dim i As Integer
            Dim aSplitFirma() As String
                        
            For i = 0 To UBound(aPersonasInvol)
                aSplitFirma = Split(aPersonasInvol(i), ",")
                            
                If (Trim(.TextMatrix(.row, 1)) = aSplitFirma(0)) Then
                    If InStr(1, cVioFirma, aSplitFirma(0)) = 0 Then
                        cVioFirma = cVioFirma & "," & aSplitFirma(0)
                    End If
                End If
            Next i
        End If
    End If
    'end ande
    
End With
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVerRegla_Click()
    If strReglas <> "" Then
        Call frmCapVerReglas.Inicia(strReglas)
    Else
        MsgBox "Cuenta no tiene reglas definidas", vbInformation, "Aviso"
    End If
End Sub

'Cargos
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim sNumTar As String
    Dim sClaveTar As String
    Dim nErr As Integer
    Dim sCaption As String
    Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
    Dim nEstado As COMDConstantes.CaptacTarjetaEstado
    Dim nCOM As Integer
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    sCaption = Me.Caption
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        Set clsGen = New COMDConstSistema.DCOMGeneral
            bRetSinTarjeta = clsGen.GetPermisoEspecialUsuario(gCapPermEspRetSinTarj, gsCodUser, gsDominio)
        Set clsGen = Nothing
        sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta)
        If Val(Mid(sCuenta, 6, 3)) <> nProducto Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
'Fin Deshacer el descomentado
    
If KeyCode = vbKeyF11 And txtCuenta.Enabled = True Then 'F11
'        Dim nPuerto As TipoPuertoSerial
        Dim sMaquina As String
        sMaquina = GetComputerName
        sCaption = Me.Caption
        Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
        MsgBox "Pase la tarjeta.", vbInformation, "AVISO"
        'ppoa Modificacion
        'sNumTar = GetNumTarjeta
        sNumTar = ""
        sNumTar = GetNumTarjeta_ACS

        If Len(sNumTar) <> 16 Then
            MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
            Me.Caption = sCaption
            'Set clsGen = Nothing
            Exit Sub
        End If

        Me.Caption = "Ingrese la Clave de la Tarjeta."
        MsgBox "Ingrese la Clave de la Tarjeta.", vbInformation, "AVISO"
        'ppoa Modificacion
        'sClaveTar = GetClaveTarjeta
        Select Case GetClaveTarjeta_ACS(sNumTar, 1)
            Case gClaveValida
                    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    Dim rsTarj As ADODB.Recordset
                    Set rsTarj = New ADODB.Recordset
                    Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
                    Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
                    Set rsTarj = ObjTarj.Get_Datos_Tarj(sNumTar)

                    If rsTarj.EOF And rsTarj.BOF Then
                        MsgBox "Tarjeta no posee ninguna relación con cuentas de activas   o Tarjeta no activa.", vbInformation, "Aviso"
                        Me.Caption = sCaption
                        Set ObjTarj = Nothing
                        Exit Sub
                    Else
                        nEstado = rsTarj("nEstado")
                        If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
                            If nEstado = gCapTarjEstBloqueada Then
                                MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                            ElseIf nEstado = gCapTarjEstCancelada Then
                                MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                            End If
                            Set ObjTarj = Nothing
                            Me.Caption = sCaption
                            Exit Sub
                        End If
                        Dim rsPers As New ADODB.Recordset
                        Dim sCta As String, sProducto As String, sMoneda As String
                        Dim clsCuenta As UCapCuenta
                        Set rsPers = clsMant.GetCuentasPersona(rsTarj("cPersCod"), nProducto)
                        Set clsMant = Nothing
                        If Not (rsPers.EOF And rsPers.EOF) Then
                            Do While Not rsPers.EOF
                                sCta = rsPers("cCtaCod")
                                sProducto = rsPers("cDescripcion")
                                sMoneda = Trim(rsPers("cMoneda"))
                                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & space(2) & sProducto & space(2) & sMoneda
                                rsPers.MoveNext
                            Loop
                            Set clsCuenta = New UCapCuenta
                            Set clsCuenta = frmCapMantenimientoCtas.Inicia
                            If clsCuenta.sCtaCod <> "" Then
                                txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                                txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
                                txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                                txtCuenta.SetFocusCuenta
                                Call txtCuenta_KeyPress(13)
                                'SendKeys "{Enter}"
                            End If
                            Set clsCuenta = Nothing
                        Else
                            MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
                        End If
                        rsPers.Close
                        Set rsPers = Nothing
                    End If
                    Set rsTarj = Nothing
                    Set clsMant = Nothing

            Case gClaveNOValida
                MsgBox "Clave Incorrecta", vbInformation, "Aviso"
            Case Else

        End Select
        Me.Caption = "Captaciones - Cargo - Ahorros " & sOperacion
    End If
'Fin Deshacer el descomentado

    '**DAOR 20081125, Para tarjetas ***********************
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        Dim sVisibleVBSinTarj As Boolean 'CTI4 ERS0112020
        sVisibleVBSinTarj = True
        If nOperacion = 200310 Then sVisibleVBSinTarj = False
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, nProducto, sVisibleVBSinTarj)
        If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
            ObtieneDatosCuenta sCuenta
        End If
    End If
    '*******************************************************
        bPresente = False ' ande 20170918
End Sub
Public Function ObtieneDatosTarjeta(ByVal psCodTarj As String) As Boolean
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsTarj As New ADODB.Recordset
Dim sTarjeta As String, sPersona As String
Dim nEstado As COMDConstantes.CaptacTarjetaEstado

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsTarj = clsMant.GetTarjetaCuentas(psCodTarj)
If rsTarj.EOF And rsTarj.BOF Then
    MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
    ObtieneDatosTarjeta = False
Else
    ObtieneDatosTarjeta = True
End If
Set rsTarj = Nothing
Set clsMant = Nothing
End Function

Private Sub Form_Load()
    'txtGlosa.SetFocus
     fnRetiroPersRealiza = False 'WIOR 20121005
     
    ' *** AGREGADO POR RIRO
    Label5.Left = 5160
    Label5.Top = 2100
    Label5.Height = 195
    Label5.Width = 1215
    
    Label8.Left = 3600
    Label8.Top = 2100
    Label8.Height = 195
    Label8.Width = 765
    
    lblFirmas.Left = 4400
    lblFirmas.Top = 2040
    lblFirmas.Height = 300
    lblFirmas.Width = 465
    
    lblMinFirmas.Left = 6360
    lblMinFirmas.Top = 2040
    lblMinFirmas.Height = 300
    lblMinFirmas.Width = 465
    'END RIRO
        Call RestoreValidacionFirma 'ande 20170914 restaurando validación de firmas
    End Sub

Private Sub grdCliente_DblClick()
'ande 20170919
'Dim sPersCod As String
'
'sPersCod = grdCliente.TextMatrix(grdCliente.row, 1)
'If sPersCod = "" Then Exit Sub
'MuestraFirma sPersCod, gsCodAge
'end ande
End Sub

Private Sub grdCliente_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If pnCol = 8 And (Trim(grdCliente.TextMatrix(pnRow, 7)) = "PJ" Or _
                      Trim(grdCliente.TextMatrix(pnRow, 7)) = "AP") Then
        grdCliente.TextMatrix(grdCliente.row, 8) = False
    End If
End Sub

Private Sub grdCliente_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim lsColumnas() As String
    lsColumnas = Split(grdCliente.ColumnasAEditar, "-")
    
    If lsColumnas(grdCliente.Col) = "X" Then
        Cancel = False
        MsgBox "No es posible editar este campo", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub

Private Sub txtBanco_EmiteDatos()
lblBanco = Trim(txtBanco.psDescripcion)
If lblBanco <> "" Then
    If Me.cboMonedaBanco.Visible And Me.cboMonedaBanco.Enabled Then
        cboMonedaBanco.SetFocus
    Else
        txtGlosa.SetFocus
    End If
End If
End Sub

Private Sub txtBancoTrans_EmiteDatos()
    Dim oNCajaCtaIF As New clases.NCajaCtaIF
    Dim oDOperacion As New clases.DOperacion
    
    If txtBancoTrans.Text <> "" Then
       lblNombreBancoTrans.Caption = oNCajaCtaIF.NombreIF(Mid(txtBancoTrans.Text, 4, 13))
       cbPlazaTrans.SetFocus
    End If
    CalculaComision
    chkITFEfectivo_Click
    Set oNCajaCtaIF = Nothing
    Set oDOperacion = Nothing
End Sub

Private Sub txtCtaBanco_GotFocus()
    txtCtaBanco.SelStart = 0
    txtCtaBanco.SelLength = 50
End Sub

Private Sub txtCtaBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtBanco.Enabled And txtBanco.Visible Then Me.txtBanco.SetFocus
    End If
End Sub

'COMENTADO NDX ERS0112020
Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sNumTarj = ""
        sCta = txtCuenta.NroCuenta
        ObtieneDatosCuenta sCta
        frmSegSepelioAfiliacion.Inicio sCta
         'COMENTADO POR APRI-PARA EL ERS DE BONIFICACION DE AHORROS
'        '********APRI 20161031
'        If nOperacion = "200301" Or nOperacion = "200302" Then
'            BusquedaProducto "", sCta, nOperacion
'        End If
'        '********END APRI
    End If
End Sub

Private Sub txtCuentaTrans_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosaTrans.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    '***Modificado por ELRO el 20130722, según TI-ERS079-2013****
    'txtMonto.Enabled = True
    'txtMonto.SetFocus
    If cboMedioRetiro.Visible Then
        cboMedioRetiro.SetFocus
    Else
        txtMonto.Enabled = True
        txtMonto.SetFocus
    End If
    '***Fin Modificado por ELRO el 20130722, según TI-ERS079-2013
End If
End Sub

Private Sub txtglosaTrans_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If cboMedioRetiro.Visible And cboMedioRetiro.Enabled Then
        cboMedioRetiro.SetFocus
    Else
        txtMonto.Enabled = True
        txtMonto.SetFocus
    End If
End If
End Sub

Private Sub txtMonto_Change()
    If gbITFAplica And nProducto <> gCapCTS Then       'Filtra para CTS
        If txtMonto.value > gnITFMontoMin Then
            If Not ckMismoTitular.value = 1 Then 'RIRO20131212137
                If Not lbITFCtaExonerada Then
                    'If nOperacion = gAhoRetTransf Or nOperacion = gAhoDepTransf Or nOperacion = gAhoDepPlanRRHH Or nOperacion = gAhoDepPlanRRHHAdelantoSueldos Then
                    If nOperacion = gAhoDepPlanRRHH Or nOperacion = gAhoDepPlanRRHHAdelantoSueldos Then
                        Me.lblITF.Caption = Format(0, "#,##0.00")
                    ElseIf nProducto = gCapAhorros And (nOperacion <> gAhoDepChq Or nOperacion <> gCMACOAAhoDepChq Or nOperacion <> gCMACOTAhoDepChq) Then
                        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                    Else
                        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                    End If
                    '*** BRGO 20110908 ************************************************
                    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
                    If nRedondeoITF > 0 Then
                        Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
                    End If
                    '*** END BRGO
                    If bInstFinanc Then lblITF.Caption = "0.00" 'JUEZ 20140414
                Else
                    Me.lblITF.Caption = "0.00"
                End If
            Else
                Me.lblITF.Caption = "0.00"
            End If
            
            If nOperacion = gAhoRetOPCanje Or nOperacion = gAhoRetOPCertCanje Or nOperacion = gAhoRetFondoFijoCanje Then
                Me.lblTotal.Caption = Format(0, "#,##0.00")
            ElseIf nOperacion = gAhoDepChq Then
                Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption), "#,##0.00")
            Else
                If nProducto = gCapAhorros And gbITFAsumidoAho Then
                    Me.lblITF.Caption = "0.00"
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                    Me.lblITF.Caption = "0.00"
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                Else
                    If chkITFEfectivo.value = 1 Then
                        Me.lblTotal.Caption = Format(txtMonto.value + CDbl(Me.lblITF.Caption), "#,##0.00")
                    Else
                        Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                    End If
                End If
            End If
            If bInstFinanc Then lblITF.Caption = "0.00" 'JUEZ 20140414
        End If
    Else
        Me.lblITF.Caption = Format(0, "#,##0.00")
        
        If nProducto = gCTSDepChq Then
            Me.lblTotal.Caption = Format(0, "#,##0.00")
        Else
            Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
        End If
    End If
    
    If txtMonto.value = 0 Then
        Me.lblITF.Caption = "0.00"
        Me.lblTotal.Caption = "0.00"
    End If
    
    ' Agregado Por RIRO el 20130501, Proyecto Ahorro
    If Trim(txtMonto.Text) = "." Then
        txtMonto.Text = 0
    End If
    
    chkITFEfectivo_Click
    chkVBComision_Click
    
End Sub

Private Sub chkITFEfectivo_Click()
Dim nMonto As Double
nMonto = txtMonto.value

If chkITFEfectivo.value = 1 Then
    'MADM 20101112
    If nOperacion = "200304" Then
        lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption) + lblMonComision, "#,##0.00")
        Exit Sub
    End If
    'END MADM
    'Mody By GITU 29-08-2011
    If chkVBComision.value = 1 Then
        lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption) + lblMonComision, "#,##0.00")
    Else
        lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
    End If
    'End GITU
    'RIRO20131212 ERS137
    If chkTransfEfectivo.value = 0 Then
        lblTotal.Caption = Format(CDbl(lblTotal.Caption) + CDbl(lblComisionTransf), "#,##0.00")
    End If
    'END RIRO
        
Else
    If nProducto = gCapAhorros And gbITFAsumidoAho Then
        'MADM 20101112
        If nOperacion = "200304" Then
            lblTotal.Caption = Format(nMonto + lblMonComision, "#,##0.00")
            Exit Sub
        End If
        'END MADM
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
        'MADM 20101112
        If nOperacion = "200304" Then
            lblTotal.Caption = Format(nMonto + lblMonComision, "#,##0.00")
            Exit Sub
        End If
        'END MADM
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    Else
        'MADM 20101112
        If nOperacion = "200304" Then
            lblTotal.Caption = Format(nMonto + lblMonComision, "#,##0.00")
            Exit Sub
        End If
        'END MADM
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    End If
    'Mody By GITU 29-08-2011
    If chkVBComision.value = 1 Then
        lblTotal.Caption = Format(nMonto + lblMonComision, "#,##0.00")
    End If
    'End GITU
    
    'RIRO20131212 ERS137
    If chkTransfEfectivo.value = 0 Then
        lblTotal.Caption = Format(CDbl(lblTotal.Caption) + CDbl(lblComisionTransf), "#,##0.00")
    End If
    'END RIRO
    
End If
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
txtMonto.SelStart = 0
txtMonto.SelLength = Len(txtMonto.Text)
End Sub


Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And CDbl(txtMonto.Text) > 0 Then
      cmdGrabar.Enabled = True
      CalculaComision
      txtMonto_Change  'RIRO20131212 ERS137
      Call cmdGrabar_Click
End If
End Sub

Private Sub txtOrdenPago_GotFocus()
With txtOrdenPago
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtOrdenPago_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Trim(txtOrdenPago.Text) <> "" And IsNumeric(txtOrdenPago) Then
        Dim oMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
        Dim rsOP As New ADODB.Recordset
        Dim nEstadoOP As COMDConstantes.CaptacOrdPagoEstado
        Dim sCuenta As String
        sCuenta = txtCuenta.NroCuenta
        Set oMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsOP = oMant.GetDatosOrdenPago(sCuenta, CLng(Val(Trim(txtOrdenPago.Text))))
        Set oMant = Nothing
        If Not (rsOP.EOF And rsOP.BOF) Then
            nEstadoOP = rsOP("nEstado")
            If rsOP("nEstado") = gCapOPEstCertifiCada Then
                txtMonto.Text = Format$(rsOP("nMonto"), "#,##0.00")
            End If
        End If
        rsOP.Close
        Set rsOP = Nothing
    Else
        MsgBox "Debe ingresar un número válido para la Orden de Pago", vbOKOnly + vbExclamation, App.Title
        Exit Sub
    End If
    If txtBanco.Visible Then
        txtBanco.SetFocus
    Else
        txtGlosa.SetFocus
    End If
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub


Private Sub MuestraFirmas(ByVal sCuenta As String)
Dim i As Integer
Dim sPersona As String
    
If nPersoneria <> PersPersoneria.gPersonaNat Then
    For i = 1 To Me.grdCliente.Rows - 1
        If Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersRepSuplente Or Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersRepTitular Then
            sPersona = grdCliente.TextMatrix(i, 1)
            MuestraFirma sPersona, gsCodAge
        End If
    Next i
Else
    For i = 1 To Me.grdCliente.Rows - 1
        If Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersTitular Then
            sPersona = grdCliente.TextMatrix(i, 1)
            MuestraFirma sPersona, gsCodAge
        End If
    Next i
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
rs.Close
Set rs = Nothing
Set oCons = Nothing
End Function
 
Private Function VerificarAutorizacion() As Boolean
Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim oCapAutN  As COMNCaptaGenerales.NCOMCaptAutorizacion
Dim oPers As COMDPersona.UCOMAcceso
Dim rs As New ADODB.Recordset
Dim lnMonTopD As Double
Dim lnMonTopS As Double
Dim lsmensaje As String
Dim gsGrupo As String
Dim sCuenta As String, sNivel As String
Dim lbEstadoApr As Boolean
Dim nMonto As Double
Dim nMoneda As Moneda

sCuenta = txtCuenta.NroCuenta
nMonto = txtMonto.value
nMoneda = CLng(Mid(sCuenta, 9, 1))
'Obtiene los grupos al cual pertenece el usuario
Set oPers = New COMDPersona.UCOMAcceso
    gsGrupo = oPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
Set oPers = Nothing
 
'Verificar Montos
Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    'Set rs = ocapaut.ObtenerMontoTopNivAutRetCan(gsGrupo, "1", gsCodAge) RIRO20141105 ERS122
    Set rs = oCapAut.ObtenerMontoTopNivAutRetCan(gsGrupo, "1", gsCodAge, gsCodPersUser)
Set oCapAut = Nothing
 
If Not (rs.EOF And rs.BOF) Then
    lnMonTopD = rs("nTopDol")
    lnMonTopS = rs("nTopSol")
    sNivel = rs("cNivCod")
Else
    MsgBox "Usuario no Autorizado para realizar Operacion", vbInformation, "Aviso"
    VerificarAutorizacion = False
    Exit Function
End If

If nMoneda = gMonedaNacional Then
    If nMonto <= lnMonTopS Then
        VerificarAutorizacion = True
        Exit Function
    End If
Else
    If nMonto <= lnMonTopD Then
        VerificarAutorizacion = True
        Exit Function
    End If
End If
   
Set oCapAutN = New COMNCaptaGenerales.NCOMCaptAutorizacion
If sMovNroAut = "" Then 'Si es nueva, registra
    oCapAutN.NuevaSolicitudAutorizacion sCuenta, "1", nMonto, gdFecSis, gsCodAge, gsCodUser, nMoneda, gOpeAutorizacionRetiro, sNivel, sMovNroAut
    MsgBox "Solicitud Registrada, comunique a su Admnistrador para la Aprobación..." & Chr$(10) & _
        " No salir de esta operación mientras se realice el proceso..." & Chr$(10) & _
        " Porque sino se procedera a grabar otra Solicitud...", vbInformation, "Aviso"
    VerificarAutorizacion = False
Else
    'Valida el estado de la Solicitud
    If Not oCapAutN.VerificarAutorizacion(sCuenta, "1", nMonto, sMovNroAut, lsmensaje) Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        VerificarAutorizacion = False
    Else
        VerificarAutorizacion = True
    End If
End If
Set oCapAutN = Nothing
End Function

'******                 CALCULAR COMISION  AVMM-03-2006                *****
Public Function CalcularComisionRetOtraAge() As Double
    Dim nComixRet As Double
    Dim nPorComixRet As Double
    Dim nMontoLimiRetAge As Double
    Dim nMontoComixRetAge As Double 'BRGO 20110127
    Dim oCons As COMDConstantes.DCOMAgencias
    Dim clsGen As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim clsTC As COMDConstSistema.NCOMTipoCambio 'nTipoCambio
    
    Set oCons = New COMDConstantes.DCOMAgencias
     
   
        '*** Verificar Ubicacion de la Cuenta ***
            
        If oCons.VerficaZonaAgencia(gsCodAge, Mid(txtCuenta.NroCuenta, 4, 2)) Then
        
            'Set clsGen = New COMNCaptaGenerales.NCOMCaptaDefinicion Comentado by JACA 20111021
            ' If nMoneda = gMonedaNacional Then
            '       nMontoLimiRetAge = clsGen.GetCapParametro(gMontoLimiMNRetOtraAge)
            ' Else
            '     nMontoLimiRetAge = clsGen.GetCapParametro(gMontoLimiMERetOtraAge)
            '     nComixRetAgeME = clsGen.GetCapParametro(2047)
            ' End If
            ' ' *** Verificar Monto de Retiro en Otra Agencia ***
            ' If CDbl(txtMonto.Text) >= nMontoLimiRetAge Then
            ' ' *** Calcular Comisión ***
            '     nPorComixRet = clsGen.GetCapParametro(gComisionRetOtraAge)
            '     nComixRet = CDbl(txtMonto.Text) * (Val(nPorComixRet) / 100)
            ' Else
            '     nComixRet = 0
            ' End If
                
            '***** BRGO 20110128 *********************
                'Comentado by JACA 20111021*****************************************************
                 'If nmoneda = gMonedaNacional Then
                 '    nMontoComixRetAge = clsGen.GetCapParametro(2046) 'Monto comisión en soles
                 'Else
                 '    nMontoComixRetAge = clsGen.GetCapParametro(2047) 'Monto comisión en dólares
                 'End If
                 
                 'Valida si cuenta no pertenece a las Agencias Requena o Aguaytía
                 'If (Mid(txtCuenta.NroCuenta, 4, 2) <> 12 And Mid(txtCuenta.NroCuenta, 4, 2) <> 13) And (gsCodAge <> 12 And gsCodAge <> 13) Then
                 '    nComixRet = nMontoComixRetAge
                 'Else
                 '    If gsCodAge = 12 Or gsCodAge = 13 Then
                 '        nPorComixRet = clsGen.GetCapParametro(gComisionRetOtraAge)
                 '        nComixRet = CDbl(txtMonto.Text) * (val(nPorComixRet) / 100)
                 '        If nComixRet < nMontoComixRetAge Then
                 '            nComixRet = nMontoComixRetAge
                 '        End If
                 '    Else
                 '       nComixRet = nMontoComixRetAge
                 '    End If
                 'End If
                 'Set clsGen = Nothing
                'JACA END**********************************************************************
          '***END BRGO
          'JACA 20111021*******************************************************
                
                Dim objCap As New COMNCaptaGenerales.NCOMCaptaGenerales
                Dim rs As New Recordset
                
                'Obtenemos los parametros de la comision
                'Set rs = objCap.obtenerValTarifaOpeEnOtrasAge(gsCodAge, lnTpoPrograma, CInt(Mid(Me.txtCuenta.NroCuenta, 9, 1)))
                'Set rs = objCap.obtenerValTarifaOpeEnOtrasAge(Me.TxtCuenta.NroCuenta) 'APRI20190109 ERS077-2018
                Set rs = objCap.obtenerValTarifaOpeEnOtrasAge(Me.txtCuenta.NroCuenta, 1)  'APRI20210403
                If Not (rs.EOF And rs.BOF) Then
                    
                    If rs!nComision > 0 Then
                        'si la comision es menor que el monto minimo
                        If CDbl(txtMonto.Text) * (rs!nComision / 100) < rs!nMontoMin Then
                            CalcularComisionRetOtraAge = rs!nMontoMin
                        'si la comision es mayor que el monto maximo
                        ElseIf CDbl(txtMonto.Text) * (rs!nComision / 100) > rs!nMontoMax Then
                            CalcularComisionRetOtraAge = rs!nMontoMax
                        'si la comision esta entre en monto min y max
                        Else
                            CalcularComisionRetOtraAge = CDbl(txtMonto.Text) * (rs!nComision / 100)
                        End If
                    Else
                        CalcularComisionRetOtraAge = 0
                    End If
                    
                Else
                     CalcularComisionRetOtraAge = 0
                End If
                'JACA END************************************************************
        Else
            'nComixRet = 0 comentado by JACA 20111021
            CalcularComisionRetOtraAge = 0
        End If
    Set oCons = Nothing
    'CalcularComisionRetOtraAge = nComixRet Comentado by JACA 20111021
End Function

'******   VERIFICAR MAX MOVIMIENTOS Y CALCULAR COMISION  AVMM 03-06-2006  *****
Public Function CalcularComisionxMaxOpeRet() As Double
    Dim nNroMaxOpe As Double
    Dim nNroOpe As Double
    Dim nMontoxMaxOpe As Double
    Dim nMontoTope As Double
    Dim sFecha As String
    Dim oCapG As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim clsGen As COMNCaptaGenerales.NCOMCaptaDefinicion
    
    Dim rsComision As Recordset 'JACA 20111025
    Dim oCapN As New COMNCaptaGenerales.NCOMCaptaGenerales 'JACA 20111025
    Dim clsTC As COMDConstSistema.NCOMTipoCambio 'JACA 20111025
    Dim nTC As Double 'JACA 20111025

    
    Set oCapG = New COMDCaptaGenerales.DCOMCaptaGenerales
    
        Set clsGen = New COMNCaptaGenerales.NCOMCaptaDefinicion
            
            Set rsComision = New Recordset
            Set rsComision = oCapN.obtenerComisionxNroOperaciones(gsOpeCod, gsCodAge, lnTpoPrograma)
            
            If Not (rsComision.EOF And rsComision.BOF) Then
                    nNroMaxOpe = rsComision!nOpeMax
            Else
                CalcularComisionxMaxOpeRet = 0
                Exit Function
            End If
            ' *** Obtener Nro de Operaciones ***
            'sFecha = Mid(Format(gdFecSis, "yyyymmdd"), 1, 6)'Comentado By JACA 20111025
            nNroOpe = oCapN.obtenerNroMovimientos(txtCuenta.NroCuenta, gdFecSis)
            
            ' *** Verificar Nro de Operaciones ***
            If nNroOpe >= nNroMaxOpe Then
            
            ' *** Calcular Monto de Comisión ***
                 
                 'Comentado by JACA 20111025******************************************
                    'If nmoneda = gMonedaNacional Then
                    '    nMontoTope = clsGen.GetCapParametro(2097)
                    '    If nSaldoCuenta <= nMontoTope Then
                    '       nMontoxMaxOpe = clsGen.GetCapParametro(gMontoMNx31Ope)
                    '    Else
                    '       nMontoxMaxOpe = 0
                    '    End If
                    'Else
                    '    nMontoTope = clsGen.GetCapParametro(2098)
                    '    If nSaldoCuenta <= nMontoTope Then
                    '       nMontoxMaxOpe = clsGen.GetCapParametro(gMontoMEx31Ope)
                    '    Else
                    '       nMontoxMaxOpe = 0
                    '    End If
                    'End If
                    
                 'JACA END*************************************************************
                    
                 'JACA 20111025********************************************************
                        Set clsTC = New COMDConstSistema.NCOMTipoCambio
                        nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
 
                        If nMoneda = gMonedaNacional Then
                            CalcularComisionxMaxOpeRet = rsComision!nMontoMN
                        ElseIf nMoneda = gMonedaExtranjera Then
                            If nTC <> 0 Then
                                CalcularComisionxMaxOpeRet = rsComision!nMontoMN / nTC
                            Else
                                CalcularComisionxMaxOpeRet = 0
                            End If
                        End If
                 'JACA END************************************************************
                 
            Else
                'nMontoxMaxOpe = 0
                CalcularComisionxMaxOpeRet = 0 'JACA 20111025
            End If
            
        Set clsGen = Nothing
        
    Set oCapG = Nothing
    'CalcularComisionxMaxOpeRet = nMontoxMaxOpe Comentado by JACA 20111025
End Function

Sub Finaliza_Verifone5000()
        If Not GmyPSerial Is Nothing Then
            GmyPSerial.Disconnect
            Set GmyPSerial = Nothing
        End If
End Sub

Private Sub CargoAutomatico(psCuenta As String, pnCntPag As Integer)
    Dim sMensajeCola As String
    Dim bExito As Boolean
    Dim nMonto As Double
    Dim bAplicar As Boolean
    Dim lsmensaje As String
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    bAplicar = True
    
    bExito = False
    If bAplicar Then
        If nProducto = gCapAhorros And nPersoneria <> gPersonaJurCFLCMAC Then 'Ahorros y que no sean CMACs
            If nEstadoT <> gCapEstAnulada And nEstadoT <> gCapEstCancelada Then
                nMoneda = CLng(Mid(psCuenta, 9, 1))
                If nMoneda = gMonedaNacional Then
                        nMonto = GetMontoDescuento(2113, pnCntPag)
                Else
                        nMonto = GetMontoDescuento(2112, pnCntPag)
                End If
                If nMonto > 0 Then
                    Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento  'NCapMovimientos
                    Dim sMovNro As String
                    Dim oMov As COMNContabilidad.NCOMContFunciones  'NContFunciones
                    Dim nFlag As Double, nITF As Currency
                    
                    nITF = fgITFCalculaImpuesto(nMonto)
                    Set oMov = New COMNContabilidad.NCOMContFunciones
                    sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Set oMov = Nothing
                    
                    Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    'oCap.IniciaImpresora gImpresora
                    nFlag = oCap.CapCargoCuentaAho(psCuenta, nMonto, gAhoRetComOPCanje, sMovNro, "Descuento Comisión Orden Pago", , , , , , , , , gsNomAge, sLpt, , , , , gsCodCMAC, , gsCodAge, , , nITF, , , , , lsmensaje, lsBoleta, lsBoletaITF)
                    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
                    Set oCap = Nothing
                Else
                    bExito = False
                End If
            End If
        End If
    End If
End Sub

Private Sub CobroComisionVBEfec(psCuenta As String, pnMonComVB As Double, psCodOpe As String, pnMovNroOpe As Long)
Dim sMensajeCola As String
Dim bExito As Boolean
Dim nMonto As Double
Dim bAplicar As Boolean
Dim lsmensaje As String
Dim lsBoleta As String

    bAplicar = True
    
    bExito = False
    If bAplicar Then
        If nProducto = gCapAhorros And nPersoneria <> gPersonaJurCFLCMAC Then 'Ahorros y que no sean CMACs
            nMoneda = CLng(Mid(psCuenta, 9, 1))
'            If nmoneda = gMonedaNacional Then
'                nMonto = GetMontoDescuento(2113, pnCntPag)
'            Else
'                nMonto = GetMontoDescuento(2112, pnCntPag)
'            End If
            If pnMonComVB > 0 Then
                Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento  'NCapMovimientos
                Dim sMovNro As String
                Dim oMov As COMNContabilidad.NCOMContFunciones  'NContFunciones
                Dim nFlag As Double, nITF As Currency
                
                Set oMov = New COMNContabilidad.NCOMContFunciones
                sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Set oMov = Nothing
                
                Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                'oCap.IniciaImpresora gImpresora
                nFlag = oCap.CapCobroComisionEfec(psCuenta, pnMonComVB, psCodOpe, sMovNro, "Comision Ope. Sin Tarjeta Efec.", , , gsNomAge, sLpt, nMoneda, gsCodCMAC, , gsCodAge, , lsmensaje, lsBoleta, gbImpTMU, pnMovNroOpe)
                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                Set oCap = Nothing
            Else
                bExito = False
            End If
        End If
    End If
End Sub
'***Agregado por ELRO el 20130724, según TI-ERS079-2013****
Private Sub cargarMediosRetiros()
Dim ODCOMConstantes As COMDConstantes.DCOMConstantes
Set ODCOMConstantes = New COMDConstantes.DCOMConstantes
Dim rsMedio As ADODB.Recordset
Set rsMedio = New ADODB.Recordset

Set rsMedio = ODCOMConstantes.devolverMediosRetiros
cboMedioRetiro.Clear

If Not (rsMedio.BOF And rsMedio.EOF) Then
 Do While Not rsMedio.EOF
     cboMedioRetiro.AddItem rsMedio!cConsDescripcion & space(100) & rsMedio!nConsValor
     rsMedio.MoveNext
 Loop
End If
Set rsMedio = Nothing
Set ODCOMConstantes = Nothing
End Sub
'***Fin Agregado por ELRO el 20130724, según TI-ERS079-2013

'JUEZ 20130731 ***********************************************************************
Public Function ValidaSaldoAutorizadoRetiroCTS() As Boolean
    ValidaSaldoAutorizadoRetiroCTS = False
    If CDbl(txtMonto.Text) > nSaldoRetiro Then
        'MsgBox "El saldo autorizado de retiro máximo para esta cuenta sólo es de " & IIf(Mid(txtCuenta.NroCuenta, 9, 1), "S/. ", "$ ") & Format(nSaldoRetiro, "#,##0.00"), vbInformation, "Aviso"
        MsgBox "Cuenta no posee saldo suficiente para realizar el retiro", vbInformation, "Aviso"
        txtMonto.Text = Format(nSaldoRetiro, "#,##0.00")
        Exit Function
    End If
    ValidaSaldoAutorizadoRetiroCTS = True
End Function
'END JUEZ ****************************************************************************

'Validar si Personas cumplen con las Reglas, Agregado por RIRO 20/11/2012
Private Function ValidarReglasPersonas() As Boolean
 Dim sReglas() As String
    Dim sGrupos() As String
    Dim sTemporal As String
    Dim v1, v2 As Variant
    Dim bAprobado As Boolean
    Dim intRegla, i, J As Integer
    
    If Trim(strReglas) = "" Then
        ValidarReglasPersonas = False
        Exit Function
    End If
    sReglas = Split(strReglas, "-")
    For i = 1 To grdCliente.Rows - 1
        If grdCliente.TextMatrix(i, 8) = "." Then
            If J = 0 Then
               sTemporal = sTemporal & grdCliente.TextMatrix(i, 7)
            Else
               sTemporal = sTemporal & "," & grdCliente.TextMatrix(i, 7)
            End If
            J = J + 1
        End If
    Next
    If Trim(sTemporal) = "" Then
        ValidarReglasPersonas = False
        Exit Function
    End If
    sGrupos = Split(sTemporal, ",")
    For Each v1 In sReglas
        bAprobado = True
        For Each v2 In sGrupos
            If InStr(CStr(v1), CStr(v2)) = 0 Then
                bAprobado = False
                Exit For
            End If
        Next
        If bAprobado Then
            If UBound(sGrupos) = UBound(Split(CStr(v1), "+")) Then
                Exit For
            Else
                bAprobado = False
            End If
        End If
    Next
    ValidarReglasPersonas = bAprobado
End Function

' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

Private Function ObtenerRegla() As String

    Dim nLetraMin, nMedio, i, J As Integer
    Dim sRegla As String
    Dim nReglas() As Integer
    
    nLetraMin = 65
    nMedio = 90
    J = 0
    ReDim Preserve nReglas(0)
    For i = 1 To grdCliente.Rows - 1
        If Trim(grdCliente.TextMatrix(i, 8)) = "." Then
            If Trim(grdCliente.TextMatrix(i, 7)) <> "AP" And Trim(grdCliente.TextMatrix(i, 7)) <> "N/A" Then
                If Len(Trim(grdCliente.TextMatrix(i, 7))) > 0 Then
                    ReDim Preserve nReglas(J)
                    nReglas(J) = CInt(AscW(grdCliente.TextMatrix(i, 7)))
                    J = J + 1
                End If
            End If
        End If
    Next
    
    nLetraMin = 0
    nMedio = 90
    
    If J > 0 Then
        For i = 0 To UBound(nReglas)
            nMedio = 90
            For J = 0 To UBound(nReglas)
                If nReglas(J) > nLetraMin And nReglas(J) <= nMedio Then
                    nMedio = nReglas(J)
                End If
            Next
            nLetraMin = nMedio
            sRegla = sRegla & "+" & ChrW(nMedio)
        Next
        sRegla = Mid(sRegla, 2, Len(sRegla) - 1)
    Else
        sRegla = ""
    End If
    
    ObtenerRegla = sRegla

End Function

' *** FIN RIRO

' RIRO20131212 ERS137
Private Sub CalculaComision()

Dim idBanco As String
Dim nPlaza As Integer
Dim nMoneda As Integer
Dim nTipo As Integer
Dim nMonto As Double
Dim nComision As Double
Dim oDefinicion As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oDefinicion = New COMNCaptaGenerales.NCOMCaptaDefinicion

If fraDocumentoTrans.Enabled And fraDocumentoTrans.Visible Then
    
    If nOperacion = gAhoRetTransf Or nOperacion = gAhoCancTransfAbCtaBco Or _
       nOperacion = gPFRetIntAboCtaBanco Or nOperacion = gPFCancTransf Or _
       nOperacion = gCTSRetTransf Or nOperacion = gCTSCancTransfBco Then
        
        idBanco = Mid(txtBancoTrans.Text, 4, 13)
        nPlaza = CDbl(Val(Trim(Right(cbPlazaTrans.Text, 8))))
        nMoneda = Val(Mid(txtCuenta.NroCuenta, 9, 1))
        nTipo = 102 ' Emision
        nMonto = txtMonto.value
        'nComision = oDefinicion.getCalculaComision(idBanco, nPlaza, nMoneda, nTipo, nMonto, gdFecSis) 'Comentado CTI4 ERS0112020
        If Not nOperacion = gAhoRetTransf Then nComision = oDefinicion.getCalculaComision(idBanco, nPlaza, nMoneda, nTipo, nMonto, gdFecSis) 'CTI4 ERS0112020
        
    End If
    
End If

'lblComisionTransf.Caption = Format(Round(nComision, 2), "#0.00") 'Comentado CTI4 ERS0112020
'CTI4 ERS0112020
If nOperacion = gAhoRetTransf Then
    lblComisionTransf.Caption = Format(Round(GetMontoComisionTrAboCtaBco(txtMonto.value, Val(Mid(txtCuenta.NroCuenta, 9, 1))), 2), "#0.00")
Else
    lblComisionTransf.Caption = Format(Round(nComision, 2), "#0.00")
End If
'CTI4 end

End Sub

Private Function getTitular() As String

    Dim sTitular As String
    Dim nI As Integer
    
    If txtTitular.Visible And txtTitular.Enabled Then
        sTitular = Trim(txtTitular.Text)
        
    Else
        For nI = 1 To grdCliente.Rows - 1
            If Val(Trim(Right(grdCliente.TextMatrix(nI, 3), 5))) = 10 Then
                sTitular = Trim(grdCliente.TextMatrix(nI, 2))
                nI = grdCliente.Rows
            End If
        Next
    End If
    
    getTitular = sTitular
    
End Function

Private Function getGlosa() As String
    Dim sGlosa As String
    sGlosa = "Bco Destino: " & lblNombreBancoTrans.Caption & ", Titular: " & getTitular & ", " & Trim(txtGlosaTrans.Text)
    getGlosa = UCase(sGlosa)
End Function

Private Sub txtTitular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCuentaTrans.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtMonto_LostFocus()
    CalculaComision
    chkITFEfectivo_Click
End Sub

Private Sub txtBancoTrans_Click(psCodigo As String, psDescripcion As String)
    CalculaComision
    chkITFEfectivo_Click
End Sub

Private Sub chkTransfEfectivo_Click()
    CalculaComision
    chkITFEfectivo_Click
End Sub

' END RIRO
'ande 20170914 limpiando la validación de firmas
Private Sub RestoreValidacionFirma()
    ReDim Preserve aPersonasInvol(0) 'default ubound(1)
    cVioFirma = ""
    aPersonasInvol(0) = ""
    gbTieneFirma = False
    bFirmasPendientes = False
    bFirmaObligatoria = False
End Sub
'end ande
'CTI4 ERS0112020
Private Function GetMontoComisionEmisionCheque(ByVal pnMoneda As Moneda) As Double
Dim oParam As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As New ADODB.Recordset

Set oParam = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set rsPar = oParam.GetParametrosComision(IIf(pnMoneda = gMonedaNacional, "CGA2009", "CGA2027"), "1", "A")
Set oParam = Nothing

If rsPar.EOF And rsPar.BOF Then
    GetMontoComisionEmisionCheque = 0
Else
    GetMontoComisionEmisionCheque = rsPar("nParMonto")
End If
rsPar.Close
Set rsPar = Nothing
End Function
Private Function GetMontoComisionTrAboCtaBco(ByVal pnMonto As Double, ByVal pnMoneda As Moneda) As Double
Dim oParam As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsTC As COMDConstSistema.NCOMTipoCambio
Dim rsPar As New ADODB.Recordset
Dim sParCod As String
Dim nPrimerRango, nSegundoRango, nTC As Double

Set oParam = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set clsTC = New COMDConstSistema.NCOMTipoCambio

If pnMonto = 0 Then GetMontoComisionTrAboCtaBco = 0: Exit Function

nPrimerRango = 50000: nSegundoRango = 150000
If pnMonto <= nPrimerRango Then
    sParCod = "CGA2006"
ElseIf pnMonto > nPrimerRango And pnMonto <= nSegundoRango Then
    sParCod = "CGA2007"
ElseIf pnMonto > nSegundoRango Then
    sParCod = "CGA2008"
End If

nTC = clsTC.EmiteTipoCambio(gdFecSis, TCCompra)

Set rsPar = oParam.GetParametrosComision(sParCod, "1", "A")
Set oParam = Nothing

If rsPar.EOF And rsPar.BOF Then
    GetMontoComisionTrAboCtaBco = 0
Else
    GetMontoComisionTrAboCtaBco = rsPar("nParMonto") / (IIf(pnMoneda = gMonedaNacional, 1, nTC))
End If
rsPar.Close
Set rsPar = Nothing
End Function
'CTI4 End
