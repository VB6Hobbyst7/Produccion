VERSION 5.00
Begin VB.Form frmCapCancelacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmCapCancelacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8880
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
      Height          =   3045
      Left            =   60
      TabIndex        =   49
      Top             =   4725
      Visible         =   0   'False
      Width           =   4560
      Begin VB.ComboBox cbPlazaTrans 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   720
         Width           =   1725
      End
      Begin VB.CheckBox ckMismoTitular 
         Caption         =   "Mismo Titular"
         Height          =   315
         Left            =   945
         TabIndex        =   53
         Top             =   1080
         Width           =   1365
      End
      Begin VB.TextBox txtCuentaTrans 
         Height          =   315
         Left            =   945
         MaxLength       =   20
         TabIndex        =   52
         Top             =   1440
         Width           =   1725
      End
      Begin VB.TextBox txtGlosaTrans 
         Height          =   735
         Left            =   945
         TabIndex        =   51
         Top             =   1965
         Width           =   3525
      End
      Begin VB.TextBox txtTitular 
         Height          =   315
         Left            =   2700
         TabIndex        =   50
         Top             =   1080
         Width           =   1770
      End
      Begin SICMACT.TxtBuscar txtBancoTrans 
         Height          =   315
         Left            =   945
         TabIndex        =   55
         Top             =   315
         Width           =   1725
         _ExtentX        =   3043
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
      Begin VB.Label lblNombreBanco 
         Caption         =   "Banco:"
         Height          =   285
         Left            =   135
         TabIndex        =   60
         Top             =   315
         Width           =   690
      End
      Begin VB.Label lblPlaza 
         Caption         =   "Plaza:"
         Height          =   285
         Left            =   135
         TabIndex        =   59
         Top             =   765
         Width           =   645
      End
      Begin VB.Label Label10 
         Caption         =   "CCI:"
         Height          =   240
         Left            =   120
         TabIndex        =   58
         Top             =   1485
         Width           =   915
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
         Left            =   2700
         TabIndex        =   57
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label lblGlosa 
         Caption         =   "Glosa:"
         Height          =   330
         Left            =   135
         TabIndex        =   56
         Top             =   2010
         Width           =   735
      End
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
      Height          =   1350
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   8745
      Begin VB.Frame fraDatos 
         Height          =   585
         Left            =   105
         TabIndex        =   15
         Top             =   660
         Width           =   7680
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   135
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   195
            Width           =   1965
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Contacto :"
            Height          =   195
            Left            =   2970
            TabIndex        =   17
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
            Left            =   4215
            TabIndex        =   16
            Top             =   195
            Width           =   1995
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
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
         TabIndex        =   20
         Top             =   240
         Width           =   3960
      End
   End
   Begin VB.PictureBox pctNotaAbono 
      Height          =   300
      Left            =   3360
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   28
      Top             =   7890
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pctCheque 
      Height          =   345
      Left            =   2940
      ScaleHeight     =   285
      ScaleWidth      =   120
      TabIndex        =   27
      Top             =   7860
      Visible         =   0   'False
      Width           =   180
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
      Height          =   2955
      Left            =   60
      TabIndex        =   21
      Top             =   4740
      Width           =   4530
      Begin VB.CommandButton cmdDocumento 
         Height          =   350
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   315
         Width           =   475
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   333
         Width           =   2610
      End
      Begin VB.TextBox txtGlosa 
         Height          =   1215
         Left            =   705
         TabIndex        =   4
         Top             =   840
         Width           =   3705
      End
      Begin VB.Label lblDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   390
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Clientes"
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
      Height          =   3375
      Left            =   60
      TabIndex        =   9
      Top             =   1365
      Width           =   8745
      Begin VB.CommandButton cmdMostrarFirma 
         Caption         =   "Mostrar Firma"
         Height          =   330
         Left            =   7320
         TabIndex        =   64
         Top             =   2120
         Width           =   1245
      End
      Begin VB.CommandButton cmdVerRegla 
         Caption         =   "Ver Reglas"
         Height          =   330
         Left            =   6150
         TabIndex        =   48
         Top             =   2120
         Width           =   1125
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1845
         Left            =   105
         TabIndex        =   1
         Top             =   225
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3254
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Direccion-ID-Firma Oblig-Grupo-Presente"
         EncabezadosAnchos=   "250-1700-3800-1500-0-0-1200-1200-1200"
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
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-4"
         EncabezadosAlineacion=   "C-L-L-L-C-C-C-C-C"
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Exoneración :"
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   3000
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
         Left            =   1590
         TabIndex        =   37
         Top             =   2940
         Visible         =   0   'False
         Width           =   7065
      End
      Begin VB.Label LblMinFirmas 
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
         Left            =   5490
         TabIndex        =   36
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mínimo Firmas :"
         Height          =   195
         Left            =   4320
         TabIndex        =   35
         Top             =   2220
         Width           =   1110
      End
      Begin VB.Label LblAlias 
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
         Left            =   1575
         TabIndex        =   34
         Top             =   2535
         Width           =   7065
      End
      Begin VB.Label Label3 
         Caption         =   "Alias de la Cuenta:"
         Height          =   225
         Left            =   165
         TabIndex        =   33
         Top             =   2595
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2220
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
         Left            =   1575
         TabIndex        =   12
         Top             =   2160
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "# Firmas :"
         Height          =   195
         Left            =   3120
         TabIndex        =   11
         Top             =   2220
         Width           =   690
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
         Left            =   3825
         TabIndex        =   10
         Top             =   2160
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Top             =   7860
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6750
      TabIndex        =   6
      Top             =   7830
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   7830
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
      Height          =   3045
      Left            =   4650
      TabIndex        =   24
      Top             =   4725
      Width           =   4155
      Begin VB.CheckBox chkTransfEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   63
         Top             =   3285
         Width           =   705
      End
      Begin VB.ComboBox cboMedioRetiro 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkVBComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   44
         Top             =   1393
         Visible         =   0   'False
         Width           =   735
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   975
         Width           =   1920
         _ExtentX        =   3387
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
         Caption         =   "Comision: "
         Height          =   255
         Left            =   645
         TabIndex        =   62
         Top             =   3270
         Width           =   735
      End
      Begin VB.Label lblMonComisionTransf 
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
         Left            =   2400
         TabIndex        =   61
         Top             =   3225
         Width           =   1095
      End
      Begin VB.Label lblMedioRetiro 
         AutoSize        =   -1  'True
         Caption         =   "Medio de Retiro :"
         Height          =   195
         Left            =   225
         TabIndex        =   47
         Top             =   1395
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblEtqComi 
         AutoSize        =   -1  'True
         Caption         =   "Comision :"
         Height          =   195
         Left            =   720
         TabIndex        =   45
         Top             =   1785
         Visible         =   0   'False
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
         Left            =   2400
         TabIndex        =   43
         Top             =   1725
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblInteres 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1560
         TabIndex        =   42
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label LblCapital 
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
         Left            =   1560
         TabIndex        =   41
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Interes :"
         Height          =   195
         Left            =   720
         TabIndex        =   40
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Capital :"
         Height          =   195
         Left            =   720
         TabIndex        =   39
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   720
         TabIndex        =   32
         Top             =   2130
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   600
         TabIndex        =   31
         Top             =   2550
         Width           =   450
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
         Left            =   2400
         TabIndex        =   30
         Top             =   2085
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
         ForeColor       =   &H000040C0&
         Height          =   300
         Left            =   1590
         TabIndex        =   29
         Top             =   2460
         Width           =   1905
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/."
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
         Left            =   3600
         TabIndex        =   26
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   720
         TabIndex        =   25
         Top             =   1080
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmCapCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As COMDConstantes.Producto
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim nMoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim bDocumento As Boolean
Dim nDocumento As COMDConstantes.tpoDoc
Dim nPersoneria As COMDConstantes.PersPersoneria
Dim sOperacion As String
Dim sMovNroAut  As String
Dim lbITFCtaExonerada As Boolean
Dim gsPersCod As String

Dim lnTpoPrograma As Integer
Dim sNumTarj As String
Dim sCuenta As String
Dim cGetValorOpe As String
Dim nComisionVB As Double
Dim nRedondeoITF As Double 'BRGO 20110914

'Agregado Por RIRO el 20130501, Proyecto Ahorro - Poderes
Dim bProcesoNuevo As Boolean
Dim strReglas As String

' RIRO20131212 ERS137
Dim nMontoCancelacion As Double
Dim bInstFinanc As Boolean 'JUEZ 20140414

Dim lbVistoVal As Boolean 'Add by GITU 04-10-2016
'***ADD BY MARG 20171222 - ERS065-2017 ---SUBIDO DESDE LA 60***
Dim loVistoElectronico As frmVistoElectronico
Dim nRespuesta As Integer
'END MARG ***********************

'Funcion de ImpresionnMontoCancelacion de Boletas
Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, oImpresora.gPrnSaltoLinea & sBoleta & oImpresora.gPrnSaltoLinea
    Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub 'Funcion de Impresion de Boletas

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
            poLavDinero.TitPersLavDineroDir = grdCliente.TextMatrix(i, 3)
            poLavDinero.TitPersLavDineroDoc = grdCliente.TextMatrix(i, 4)
            Exit For
        End If
    Else
        'By Capi 08072008
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.TitPersLavDineroDir = grdCliente.TextMatrix(i, 3)
            poLavDinero.TitPersLavDineroDoc = grdCliente.TextMatrix(i, 4)
            'Exit For
        End If
        '
        If nRelacion = gCapRelPersRepTitular Then
            poLavDinero.ReaPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.ReaPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.ReaPersLavDineroDir = grdCliente.TextMatrix(i, 3)
            poLavDinero.ReaPersLavDineroDoc = grdCliente.TextMatrix(i, 4)
            Exit For
        End If
    End If
Next i
nMonto = TxtMonto.value
sCuenta = txtCuenta.NroCuenta
'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, False, nMonto, sCuenta, sOperacion, , Me.lblTipoCuenta.Caption)
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
Dim lnCapital As Double
Dim lnInteres As Double
'By Capi 11112008
Dim lnRetencion As Double
'
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
'----- MADM
Dim nPendienteMC As Double 'APRI20190109 ERS077-2019
'Dim loVistoElectronico As New frmVistoElectronico 'COMMENT BY MARG 20171222 - ERS065-2017 ----SUBIDO DESDE LA 60
Set loVistoElectronico = New frmVistoElectronico 'ADD BY MARG 20171222 - ERS065-2017 ----SUBIDO DESDE LA 60
Dim lbVistoVal As Boolean
Dim lsTieneTarj As String

Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset, rsV As ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
If sMsg = "" Then
    If clsCap.TieneChequesValorizacion(sCuenta) Then
        MsgBox "La cuenta posee cheques en valorización.", vbInformation, "Aviso"
        Set clsCap = Nothing
        Exit Sub
    End If
    Set clsCap = Nothing
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    strReglas = rsCta!cReglas 'Agregado Por RIRO el 20130501,
    If Not (rsCta.EOF And rsCta.BOF) Then
        If nProducto = gCapAhorros Then
            If rsCta("nBloqueoParcial") > 0 Then
                MsgBox "La cuenta posee un Saldo Parcial Bloqueado.", vbInformation, "Aviso"
                Set clsMant = Nothing
                rsCta.Close
                Set rsCta = Nothing
                Exit Sub
            End If
            'APRI20190109 ERS077-2018
            nPendienteMC = IIf(IsNull(rsCta("nPendienteMC")), 0, rsCta("nPendienteMC"))
            If nPendienteMC > 0 Then
                MsgBox "Cuenta no puede ser Cancelada...Tiene pago pendiente por Mantenimiento de Cuenta", vbInformation, "Aviso"
                Exit Sub
            End If
            'END APRI
        End If
        '-- AVMM -- 16-06-2006 -- validar fecha de cancelacion
        If CDate(Format$(rsCta("dApertura"), "dd mmm yyyy ")) = CDate(gdFecSis) Then
            MsgBox "Cuenta no puede ser Cancelada el mismo día de la Apertura", vbInformation, "Aviso"
            Exit Sub
        End If
        'By Capi 11112008 caj aut
        lnRetencion = IIf(IsNull(rsCta("nRetencion")), 0, rsCta("nRetencion"))
        If lnRetencion > 0 Then
            MsgBox "Cuenta no puede ser Cancelada...Tiene Movimientos Pendientes Cajeros Automaticos", vbInformation, "Aviso"
            Exit Sub
        End If
        '
        nEstado = rsCta("nPrdEstado")
        nPersoneria = rsCta("nPersoneria")
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm:ss")
        If nProducto = gCapAhorros Then
            lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        End If
        
        'ITF INICIO
        lbITFCtaExonerada = fgITFVerificaExoneracion(sCuenta)
        fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
        'ITF FIN
        
        nMoneda = CLng(Mid(sCuenta, 9, 1))
        If nMoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            TxtMonto.BackColor = &HC0FFFF
            lblMon.Caption = "S/."
        Else
            sMoneda = "MONEDA EXTRANJERA"
            TxtMonto.BackColor = &HC0FFC0
            lblMon.Caption = "US$"
        End If
                
        Me.lblITF.BackColor = TxtMonto.BackColor
        Me.lblTotal.BackColor = TxtMonto.BackColor
        
        Select Case nProducto
            Case gCapAhorros
                lblMensaje.Tag = IIf(rsCta("bOrdPag"), 1, 0)
                If rsCta("bOrdPag") Then
                    lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
                Else
                    If lnTpoPrograma = 1 Then
                        lblMensaje = "AHORRO ÑAÑITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 2 Then
                        lblMensaje = "AHORROS PANDERITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 3 Then
                        '*** PEAC 20090722
                        'lblMensaje = "AHORROS PANDERO" & Chr$(13) & sMoneda
                        lblMensaje = "AHORROS POCO A POCO AHORRO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 4 Then
                        lblMensaje = "AHORROS DESTINO" & Chr$(13) & sMoneda
                    Else
                        lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
                    End If
                End If
                
                lblUltContacto = Format$(rsCta("dUltContacto"), "dd mmm yyyy hh:mm:ss")
                Me.LblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.LblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
                
                If lbITFCtaExonerada Then
                    Dim nTipoExo As Integer
                    Dim sDescripcion As String
                    sDescripcion = ""
                    nTipoExo = fgITFTipoExoneracion(sCuenta, sDescripcion)
                    lblExoneracion.Visible = True
                    lblExoneracion.Caption = sDescripcion
                End If
                
            Case gCapPlazoFijo
                lblUltContacto = rsCta("nPlazo")
                Me.LblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.LblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
                
            Case gCapCTS
                lblUltContacto = rsCta("cInstitucion")
                
        End Select
        '***Agregado por ELRO el 20130724, según TI-ERS079-2013****
        If nOperacion = gAhoCancAct Or nOperacion = gCTSCancEfec Then
            cargarMediosRetiros
        End If
        '***Fin Agregado por ELRO el 20130724, según TI-ERS079-2013
        
         ' RIRO20131210 ERS137
        If nOperacion = gAhoCancTransfAbCtaBco Or nOperacion = gCTSCancTransfBco Then
            lblMedioRetiro.Visible = True
            cboMedioRetiro.Visible = True
            cargarMediosRetiros
            cboMedioRetiro.Text = "TRANSFERENCIA BANCO                                                                                                    3"
            cboMedioRetiro.Enabled = False
            fraDocumentoTrans.Enabled = True
            CalculaComision
            'txtComision.Visible = True
        End If
        ' FIN RIRO
        
        'Add By Gitu 23-08-2011 para cobro de comision por operacion sin tarjeta
        
        If sNumTarj = "" Then
            cGetValorOpe = ""
            If nMoneda = gMonedaNacional Then
                cGetValorOpe = GetMontoDescuento(2117, 1, 1)
            Else
                cGetValorOpe = GetMontoDescuento(2118, 1, 2)
            End If
            
            'If Mid(sCuenta, 6, 3) = "232" Or Mid(sCuenta, 6, 3) = "234" Then 'Comentado x JUEZ 20140425 Para no cobrar comisión para CTS
            If Mid(sCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then 'JUEZ 20140425 Para no cobrar comisión a cuenta de ahorros ñañito
                Set rsV = clsMant.ValidaTarjetizacion(sCuenta, lsTieneTarj)
                
'                If lsTieneTarj = "SI" And rsV.RecordCount > 0 Then  'comentado por GIPO 20180803 MEMO 1809-2018-GM-DI/CMACM
                If rsV.RecordCount > 0 Then
                'COMMENT BY MARG 20171222 - ERS065-2017******************************
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
'                        loVistoElectronico.RegistraVistoElectronico (0)
'                    Else
'                        cGetValorOpe = "0.00"
'                        Exit Sub
'                    End If
                'END MARG *******************************************************
                    '***ADD BY MARG 20171222 - ERS065-2017 ----SUBIDO DESDE LA 60 **************************
                    Dim tipoCta As Integer
'                    Dim Mensaje As String
                    tipoCta = rsCta("nPrdCtaTpo")
                    If tipoCta = 0 Or tipoCta = 2 Then 'individual o indistinta
                    
                        'GIPO 20180723 cambios en Retiros sin Tarjeta
'                        If lsTieneTarj = "SI" Then
'                            Mensaje = "El Cliente posee tarjeta, para continuar se necesita el VB del Jefe o coordinador de Operaciones. ¿Desea Continuar?"
'                        Else
'                            Mensaje = "El Cliente NO posee tarjeta activa, por lo tanto se necesita el VB del Jefe o coordinador de Operaciones. ¿Desea Continuar?"
'                        End If
'
'                        If MsgBox("El Cliente posee tarjeta, para continuar se necesita el VB del Jefe o coordinador de Operaciones. ¿Desea Continuar?", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then
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
                               'GIPO 20180723
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
                    'If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? se le cobrara una comision", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'comment by marg 20171222 - ers065-2017
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg 20121222 - ers065-2017 ----SUBIDO DESDE LA 60
                        'Set loVistoElectronico = New frmVistoElectronico 'COMMENT BY MARG ERS065-2017
                
                        lbVistoVal = loVistoElectronico.Inicio(5, nOperacion)
                    
                        If Not lbVistoVal Then
                            'MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones, se cobrara comision por esta operacion", vbInformation, "Mensaje del Sistema" 'COMMENT BY MARG 20171222 - ERS065-2017
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones,", vbInformation, "Mensaje del Sistema" 'ADD BY MARG 20171222 - ERS065-2017 ----SUBIDO DESDE LA 60
                            Exit Sub
                        End If
                        
                        'loVistoElectronico.RegistraVistoElectronico (0) 'COMMENT BY MARG ERS065
                    Else
                        cGetValorOpe = "0.00"
                        Exit Sub
                    End If
                Else 'JUEZ 20140408
                    cGetValorOpe = "0.00"
                End If
                
                If cGetValorOpe <> "0.00" Then
                    '***ADD BY MARG 20171222 ERS065-2017 - ----SUBIDO DESDE LA 60***
                    If nOperacion = 200401 Then
                        lblMonComision = Format(cGetValorOpe, "#,##0.00")
                        lblEtqComi.Visible = False
                        lblMonComision.Visible = False
                        'chkVBComision.Visible = False
                    End If
                    '***END MARG****************
                    If nOperacion <> 200401 Then 'add by marg ers065-2017
                        lblMonComision = Format(cGetValorOpe, "#,##0.00")
                        
                        lblEtqComi.Visible = True
                        lblMonComision.Visible = True
                        'chkVBComision.Visible = False
                    End If 'add by marg ers065-2017
                End If
            Else 'JUEZ 20141006
                cGetValorOpe = "0.00"
            End If
            
            lblMonComision = Format(cGetValorOpe, "#,##0.00")
        End If
        'End Gitu
        
        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
        nTipoCuenta = rsCta("nPrdCtaTpo")
        lblFirmas = Format$(rsCta("nFirmas"), "#0")
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        
        Do While Not rsRel.EOF
            If rsRel("cPersCod") = gsCodPersUser Then
                rsRel.Close
                Set rsRel = Nothing
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                gsPersCod = rsRel("cPersCod")
                grdCliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & space(50) & Trim(rsRel("nPrdPersRelac"))
                grdCliente.TextMatrix(nRow, 4) = rsRel("Direccion")
                grdCliente.TextMatrix(nRow, 5) = rsRel("ID N°")
                grdCliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("cobligatorio")) Or rsRel("cobligatorio") = "N", "NO", IIf(rsRel("cobligatorio") = "S", "SI", "OPCIONAL"))
                sPersona = rsRel("cPersCod")
                
   ' ***** Agregado Por RIRO el 20130501 *****
                
                If rsRel("cGrupo") <> "" Then
                
                        bProcesoNuevo = True
                        Label5.Visible = False
                        Label8.Visible = False
                        lblFirmas.Visible = False
                        'cmdMostrarFirma.Visible = False
                        LblMinFirmas.Visible = False
                        
                        cmdVerRegla.Visible = True
                        grdCliente.ColWidth(1) = 1400
                        grdCliente.ColWidth(2) = 3400
                        grdCliente.ColWidth(3) = 1200
                        grdCliente.ColWidth(6) = 0
                        grdCliente.ColWidth(7) = 800
                        grdCliente.ColWidth(8) = 900
                        grdCliente.TextMatrix(nRow, 7) = rsRel("cGrupo")
                        
                Else
                        bProcesoNuevo = False
                        Label5.Visible = True
                        Label8.Visible = True
                        lblFirmas.Visible = True
                        'cmdMostrarFirma.Visible = True
                        LblMinFirmas.Visible = True
                        
                        cmdVerRegla.Visible = False
                        grdCliente.ColWidth(1) = 1700
                        grdCliente.ColWidth(2) = 3800
                        grdCliente.ColWidth(3) = 1500
                        grdCliente.ColWidth(6) = 1200
                        grdCliente.ColWidth(7) = 0
                        grdCliente.ColWidth(8) = 0
                    grdCliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("cobligatorio")) Or rsRel("cobligatorio") = "N", "NO", IIf(rsRel("cobligatorio") = "S", "SI", "OPCIONAL"))
                End If
                
                ' ***** Fin RIRO *****
                
            End If
            rsRel.MoveNext
        Loop
        
        'JUEZ 20140414 ****************************************
        If nOperacion = gAhoCancAct Then
            Dim i As Integer
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
        
         '********* firma madm
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(grdCliente.TextMatrix(nRow, 1), BusquedaCodigo)
         
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(grdCliente.TextMatrix(nRow, 1)), Mid(grdCliente.TextMatrix(nRow, 1), 4, 2), False, True)
            End If
         End If
         Set Rf = Nothing

         '*******************
         
        rsRel.Close
        Set rsRel = Nothing
        FraCliente.Enabled = True
        fraDocumento.Enabled = True
        fraMonto.Enabled = True
        If cboDocumento.Visible Then
            cboDocumento.Enabled = True
        Else
            'txtGlosa.SetFocus
        End If
        cmdGrabar.Enabled = True
        cmdCancelar.Enabled = True
        fraCuenta.Enabled = False
        
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento  'NCapMovimientos
            TxtMonto.Text = Format$(clsCap.GetSaldoCancelacion(sCuenta, gdFecSis, gsCodAge, , lnCapital, lnInteres), "#,##0.00")
        Set clsCap = Nothing
        
        'JUEZ 20141006 *********************************************
        'If Me.lblTotal.Caption < 0 And cGetValorOpe <> "0.00"
        If (Me.lblTotal.Caption < 0 And cGetValorOpe <> "0.00") Or (Me.TxtMonto.Text = 0 And cGetValorOpe <> "0.00") Then 'JUEZ 20150130
            MsgBox "Cuenta no posee suficiente saldo para el cobro de comisión por operación sin tarjeta", vbInformation, "Aviso"
            cGetValorOpe = 0
            lblMonComision.Caption = "0.00"
            lblEtqComi.Visible = False
            lblMonComision.Visible = False
            txtMonto_Change
        End If
        'END JUEZ **************************************************
         
'        If nProducto = gCapAhorros Then
'             If Not ValidaCreditosPendientes(sCuenta) Then
'                MsgBox "La cuenta ha sido bloqueada para retiro, pues posee créditos pendientes de pago.", vbInformation, "Operacion"
'                LimpiaControles
'                txtCuenta.SetFocus
'                Exit Sub
'            End If
'        End If
                  
        If nOperacion = gAhoCancTransfAbCtaBco Or nOperacion = gCTSCancTransfBco Then chkTransfEfectivo_Click ' RIRO20131210 ERS137
         
        lblCapital.Caption = Format(lnCapital, "#,##0.00")
        lblInteres.Caption = Format(lnInteres, "#,##0.00")
        TxtMonto.Enabled = False
        MuestraFirmas sCuenta
    End If
Else
    Set clsCap = Nothing
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
Set clsMant = Nothing

End Sub

Private Sub LimpiaControles()
lbITFCtaExonerada = False
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
lblMonComision.BackColor = &HC0FFFF
lblMonComision.Caption = "0.00"
chkVBComision.value = 0
txtGlosa = ""
TxtMonto.value = 0
If bDocumento Then
    cboDocumento.Clear
    cboDocumento.AddItem "<Nuevo>"
    If nDocumento = TpoDocNotaCargo Then
    End If
    cboDocumento.ListIndex = 0
End If
TxtMonto.BackColor = &HC0FFFF
lblITF.BackColor = &HC0FFFF
lblTotal.BackColor = &HC0FFFF

lblEtqComi.Visible = False
lblMonComision.Visible = False
chkVBComision.Visible = False
lblMonComision = "0.00"

lblMon.Caption = "S/."
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
lblCapital = ""
lblInteres.Caption = ""
FraCliente.Enabled = False
FraDatos.Enabled = False
fraDocumento = False
fraMonto.Enabled = False
fraCuenta.Enabled = True
         
        LblAlias.Caption = ""
        LblMinFirmas.Caption = ""

If nProducto = Producto.gCapAhorros Then
        Label3.Visible = True
        Label5.Visible = True
           
        LblAlias.Visible = True
        LblMinFirmas.Visible = True
ElseIf nProducto = Producto.gCapCTS Then
        Label3.Visible = False
        Label5.Visible = False
           
        LblAlias.Visible = False
        LblMinFirmas.Visible = False

End If
Me.lblExoneracion.Visible = False
sMovNroAut = ""
nRedondeoITF = 0
sNumTarj = ""
txtCuenta.SetFocus
'***Agregado por ELRO el 20130724, según TI-ERS079-2013****
If cboMedioRetiro.Visible Then
    cargarMediosRetiros
End If
'***Fin Agregado por ELRO el 20130724, según TI-ERS079-2013

' RIRO20131212 ERS137
If txtBancoTrans.Visible Then
    txtBancoTrans.Text = ""
    lblNombreBancoTrans.Caption = ""
    cbPlazaTrans.ListIndex = 0
    ckMismoTitular.value = 0
    txtCuentaTrans.Text = ""
    txtGlosaTrans.Text = ""
    txtTitular.Text = ""
    fraDocumentoTrans.Enabled = False
    nMontoCancelacion = -1
End If
' FIN RIRO
bInstFinanc = False 'JUEZ 20140414

End Sub

Public Sub Inicia(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, ByVal sOpeDesc As String)
    nProducto = nProd
    nOperacion = nOpe
    Select Case nProd
        Case gCapAhorros
            lblEtqUltCnt = "Ult. Contacto :"
            lblUltContacto.Width = 2000
            txtCuenta.Prod = Trim(Str(gCapAhorros))
            Me.Caption = "Captaciones - Abono - Ahorros " & sOpeDesc
        
            Label3.Visible = True
            Label5.Visible = True
        
            LblAlias.Visible = True
            LblMinFirmas.Visible = True
        
            grdCliente.ColWidth(6) = 1200
        Case gCapCTS
            Label3.Visible = False
            Label5.Visible = False
           
            LblAlias.Visible = False
            LblMinFirmas.Visible = False
    
            lblEtqUltCnt = "Institución :"
            lblUltContacto.Width = 4250
            lblUltContacto.Left = lblUltContacto.Left - 250
            txtCuenta.Prod = Trim(Str(gCapCTS))
            Me.Caption = "Captaciones - Abono - CTS " & sOpeDesc
            grdCliente.ColWidth(6) = 0
    End Select
    '***Agregado por ELRO el 20130724, según TI-ERS079-2013****
    If nOperacion = gAhoCancAct Or nOperacion = gCTSCancEfec Then
        lblMedioRetiro.Visible = True
        cboMedioRetiro.Visible = True
    End If
    '***Fin Agregado por ELRO el 20130724, según TI-ERS079-2013
    
    'RIRO20131226 ERS137
    nMontoCancelacion = -1
    If nOperacion = gAhoCancTransfAbCtaBco Or nOperacion = gCTSCancTransfBco Then
        MostrarControlesTransferencia
    End If
    'FIN RIRO
    
    'Verifica si la operacion necesita algun documento
    Dim clsOpe As COMDConstSistema.DCOMOperacion 'DOperacion
    Dim rsDoc As New ADODB.Recordset
    Set clsOpe = New COMDConstSistema.DCOMOperacion
    Set rsDoc = clsOpe.CargaOpeDoc(Trim(nOperacion))
    Set clsOpe = Nothing
    If Not (rsDoc.EOF And rsDoc.BOF) Then
        lblDocumento.Visible = True
        cboDocumento.Visible = True
        cmdDocumento.Visible = True
        cboDocumento.Clear
        cboDocumento.AddItem "<Nuevo>"
        nDocumento = rsDoc("nDocTpo")
        If nDocumento = TpoDocNotaCargo Then
            cmdDocumento.Picture = pctNotaAbono
        End If
        fraDocumento.Caption = Trim(rsDoc("cDocDesc"))
        cboDocumento.ListIndex = 0
        CambiaTamañoCombo cboDocumento, 300
        bDocumento = True
    Else
        lblDocumento.Visible = False
        cboDocumento.Visible = False
        cmdDocumento.Visible = False
        bDocumento = False
    End If
    rsDoc.Close
    Set rsDoc = Nothing
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.EnabledCMAC = False
    txtCuenta.EnabledProd = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    FraCliente.Enabled = False
    fraDocumento.Enabled = False
    fraMonto.Enabled = False
    sMovNroAut = ""
    
    'ADD By GITU para el uso de las operaciones con tarjeta
    If gnCodOpeTarj = 1 And (gsOpeCod = "200401" Or gsOpeCod = "200403" Or gsOpeCod = "220401" Or gsOpeCod = "220402") Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, nProducto)
        If sCuenta <> "123456789" Then
            If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If sCuenta <> "" Then
                txtCuenta.NroCuenta = sCuenta
                'txtCuenta.SetFocusCuenta
                ObtieneDatosCuenta sCuenta
            End If
            If sCuenta <> "" Then
                'Me.Show 1 'comment by marg ers065-2017
                '***ADD BY MARG 20171222 ERS 065 ----SUBIDO DESDE LA 60***
                 If nOperacion = 200401 Then 'cancelacion activa
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
            Me.lblEtqComi.Visible = True
            Me.chkVBComision.Visible = True
            Me.lblMonComision.Visible = True
            'Me.Show 1 'comment by marg ers065-2017
            '***ADD BY MARG 20171222 ERS 065 ----SUBIDO DESDE LA 60***
             If nOperacion = 200401 Then 'cancelacion activa
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
        'Me.Show 1 'comment by marg ers065-2017
        '***ADD BY MARG 20171222 - ERS 065 ----SUBIDO DESDE LA 60***
         If nOperacion = 200401 Then 'cancelacion activa
            Me.Show
            Call Form_KeyDown(121, 0)
            Me.Visible = False
            Me.Show 1
        Else
            Me.Show 1
        End If
        '***END MARG **************
    End If
    'End GITU
    bInstFinanc = False 'JUEZ 20140414
End Sub

Private Sub cboDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

'***Agregado por ELRO el 20130724, según TI-ERS079-2013****
Private Sub cboMedioRetiro_KeyPress(KeyAscii As Integer)
cmdGrabar.SetFocus
End Sub
'***Fin Agregado por ELRO el 20130724, según TI-ERS079-2013

'RIRO20131212 ERS137
Private Sub chkTransfEfectivo_Click()
    If nMontoCancelacion = -1 Then nMontoCancelacion = TxtMonto.value
'    If nOperacion = gAhoCancTransfAbCtaBco Or nOperacion = gCTSCancTransfBco Then
'        txtMonto.Text = Format(nMontoCancelacion + IIf(chkTransfEfectivo.value = 1, 0, -CDbl(lblMonComisionTransf.Caption)), "#,##0.00")
'    End If
    txtMonto_Change
End Sub
'END RIRO

Private Sub chkVBComision_Click()
Dim nMonto As Double
nMonto = TxtMonto.value
    'GITU 20110829
    If chkVBComision.value = 1 Then
        If (nOperacion = "200401" Or nOperacion = "200403") Then
            If gbITFAsumidoAho Then
                lblTotal.Caption = Format(nMonto, "#,##0.00")
            Else
                lblTotal.Caption = Format(nMonto - CDbl(lblITF.Caption), "#,##0.00")
            End If
            'lblTotal.Caption = Format(nMonto + lblMonComision, "#,##0.00")
            Exit Sub
        End If
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    Else
        If (nOperacion = "200401" Or nOperacion = "200403") Then
            If gbITFAsumidoAho Then
                lblTotal.Caption = Format(nMonto - lblMonComision, "#,##0.00")
            Else
                lblTotal.Caption = Format(nMonto - lblMonComision - CDbl(lblITF.Caption), "#,##0.00")
            End If
        Else
            If gbITFAsumidoAho Then
                lblTotal.Caption = Format(nMonto, "#,##0.00")
            Else
                lblTotal.Caption = Format(nMonto - CDbl(lblITF.Caption), "#,##0.00")
            End If
        End If
    End If
    'lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
    'End GITU
End Sub

Private Sub ckMismoTitular_Click()
    If ckMismoTitular.value Then
        txtTitular.Text = ""
        txtTitular.Visible = False
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

If cboMedioRetiro.Visible Then
    If Trim(cboMedioRetiro) = "" Then
        MsgBox "Debe seleccionar el medio de retiro.", vbInformation, "Aviso"
        cboMedioRetiro.SetFocus
        Exit Sub
    End If
End If
'WIOR 20130301 **************************
Dim fbPersonaReaAhorros As Boolean
Dim fnCondicion As Integer
Dim nI As Integer
nI = 0
'WIOR FIN *******************************
Dim sNroDoc As String, sCodIF As String
Dim nMonto As Double
Dim sCuenta As String, sGlosa As String
'Dim nI As Integer'WIOR 20130301comento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos

Dim oMov As COMDMov.DCOMMov
Set oMov = New COMDMov.DCOMMov

Dim lsmensaje As String
Dim lsBoleta As String
Dim lsBoletaITF As String

Dim nFicSal As Integer
'FRHU ERS077-2015 20151204
Dim i As Integer
For i = 1 To grdCliente.Rows - 1
    Call VerSiClienteActualizoAutorizoSusDatos(grdCliente.TextMatrix(i, 1), nOperacion)
Next i
'FIN FRHU ERS077-2015 20151204

nMonto = TxtMonto.value

If lblMonComision.Visible Then
    nComisionVB = CDbl(lblMonComision)
Else
    nComisionVB = 0
End If


 'ANDE 20180419 ERS021-2018 camapaña mundialito
    Dim cperscod As String
    Dim nTitularCod As Integer, nTipoPersona As Integer, bParticipaCamp As Boolean, cTextoDatos As String
    Dim ix As Integer
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    For ix = 1 To grdCliente.Rows - 1
        nTitularCod = Val(Right(Trim(grdCliente.TextMatrix(ix, 3)), 2))
        If nTitularCod = 10 Then
            'nTipoPersona = grdCliente.TextMatrix(ix, 1)
            cperscod = grdCliente.TextMatrix(ix, 1)
            nTipoPersona = oCaptaLN.getVerificarPersonaNatJur(cperscod)
        End If
    Next ix
    'end ande


' ***** Agregado Por RIRO el 20130501, Proyecto Ahorro - Poderes *****

    If bProcesoNuevo = True Then
        
        If ValidarReglasPersonas = False Then
            MsgBox "Personas no cumplen con ninguna regla", vbInformation
            Exit Sub
        End If
    
    'Validar Mayoria de Edad
    
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
            If MsgBox("Uno de los intervinientes en la cuenta es menor de edad, SOLO podrá cancelar la cuenta con una autorización del Juez " & vbNewLine & "Desea continuar?", vbInformation + vbYesNo, "AVISO") = vbYes Then
                
                'Dim loVistoElectronico As frmVistoElectronico 'COMMENT BY MARG ERS065-2017
                Dim lbResultadoVisto As Boolean
                'Set loVistoElectronico = New frmVistoElectronico 'COMMENT BY MARG ERS065-2017
                lbResultadoVisto = False
                lbResultadoVisto = loVistoElectronico.Inicio(3, nOperacion)
                If Not lbResultadoVisto Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End If
' *** fin riro

'Mody By GITU 2010-06-08 Para que permita cancelar las cuentas con monto cero
If nMonto = 0 Then
    If MsgBox("El Monto de la cancelacion es igual a cero ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        If TxtMonto.Enabled Then TxtMonto.SetFocus
        Exit Sub
    End If
End If
'End Gitu

sNroDoc = Trim(Left(cboDocumento.Text, 15))
If bDocumento Then
    If InStr(1, sNroDoc, "<Nuevo>", vbTextCompare) > 0 Then
        MsgBox "Debe seleccionar un documento (" & fraDocumento.Caption & ") válido para la operacion.", vbInformation, "Aviso"
        cboDocumento.SetFocus
        Exit Sub
    End If
    If nDocumento = TpoDocNotaCargo Then
        sCodIF = ""
    End If
End If

If Not gbRetiroSinFirma Then
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    If Not clsCap.CtaConFirmas(txtCuenta.NroCuenta) Then
        MsgBox "No puede cancelar, porque la cuenta no cuenta con las firmas de las personas relacionadas a ella.", vbInformation, "Aviso"
        Exit Sub
    End If
    Set clsCap = Nothing
End If
sCuenta = txtCuenta.NroCuenta

'-- AUTORIZACION -- AVMM -- 18/04/2006---------------------------------------------
 If VerificarAutorizacion = False Then Exit Sub
'----------------------------------------------------------------------------------
    'WIOR 20121009 Clientes Observados *************************************
    If nOperacion = gAhoCancAct Or nOperacion = gCTSCancEfec Then
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
    
    
 'APRI20190601 RFC1902040001
Dim nRegistraMotivo As Boolean
Dim nMotivo As Integer
Dim cGlosaMo As String
If nProducto = gCapCTS Then
    frmCapMovCancelacion.Inicio
    nRegistraMotivo = frmCapMovCancelacion.RegistraMotivo
    nMotivo = frmCapMovCancelacion.Motivo
    cGlosaMo = frmCapMovCancelacion.Glosa
    If Not nRegistraMotivo Then
        MsgBox "Debe resgistrar el motivo de la cancelación", vbInformation, "Mensaje del Sistema"
        Exit Sub
    End If
End If
'END APRI
    
If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
    Dim sMovNro As String, sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
    Dim sMovNroCom As String
    Dim ClsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim nSaldo As Double, nPorcDisp As Double
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
    Dim nMontoLavDinero As Double, nTC As Double
    Dim clsprevio As previo.clsprevio
    Dim loLavDinero As frmMovLavDinero
    
    Set loLavDinero = New frmMovLavDinero
      
    
    '--Add By GITU 26-09-2016 para mostrar el formulario de visto electronico para las cancelaciones de CTS
    If nProducto = gCapCTS Then
        If MsgBox("Comuníquese con el Supervisor de Operaciones para autorizar la operación previa confirmación de Cese del trabajador", vbQuestion + vbOKCancel, "Aviso") = vbOk Then
            
            For Cont = 0 To grdCliente.Rows - 2
                If Trim(Right(grdCliente.TextMatrix(Cont + 1, 3), 5)) = gCapRelPersTitular Then
                    sCodPersona = Trim(grdCliente.TextMatrix(Cont + 1, 1))
                End If
            Next Cont
            
            lbVistoVal = frmCapVBCancelaCTS.Inicia(sCuenta, gsCodAge, nMonto, gsCodUser, sCodPersona)
            
            If Not lbVistoVal Then
                MsgBox "Visto Incorrecto, se cancelara la operacion", vbInformation, "Mensaje del Sistema"
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    'End GITU
    
    Set clsprevio = New previo.clsprevio
     
    'Realiza la Validación para el Lavado de Dinero
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
    If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
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
                 'ALPA 20081009*******************************************************************************
                 'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, "", , , , , nmoneda)
                 sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, "", , , , , nMoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                 '********************************************************************************************
                 If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                'End

                
            End If
        Else
            Set clsExo = Nothing
        End If
        Set clsExo = Nothing
    Else
        Set clsLav = Nothing
    End If
    Set clsLav = Nothing
    'WIOR 20130301 Personas Sujetas a Procedimiento Reforzado *************************************
    fbPersonaReaAhorros = False
    If (loLavDinero.OrdPersLavDinero = "Exit" Or loLavDinero.OrdPersLavDinero = "") _
                And (nOperacion = gAhoCancAct Or nOperacion = gAhoCancTransfAct Or nOperacion = gCTSCancEfec Or nOperacion = gCTSCancTransf) Then
                
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
                    frmPersRealizaOpeGeneral.Inicia Me.Caption & " (Persona " & sConPersona & ")", nOperacion
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
                            frmPersRealizaOpeGeneral.Inicia Me.Caption, nOperacion
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
    
    Set ClsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
    sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Sleep (1000)
    sMovNroCom = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set ClsMov = Nothing
    On Error GoTo ErrGraba
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    clsCap.IniciaImpresora gImpresora
    sCuenta = txtCuenta.NroCuenta
    sGlosa = Trim(txtGlosa)
    
    'ANDE 20180426 ERS021-2018
    Dim nPuntosRef As Integer, nCondicion As Integer, nPTotalAcumulado As Integer
    If _
        nOperacion = gAhoCancAct Or nOperacion = gAhoCancTransfAct _
    Then
        Dim nOpeTipo As Integer
        nOpeTipo = 3 '3:RETIRO
        nMoneda = Mid(sCuenta, 9, 1)
        If nMoneda = gMonedaNacional Then
            'Dim oCampanhaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
            Call oCaptaLN.VerificarParticipacionCampMundial(cperscod, sCuenta, nOperacion, nOpeTipo, nMoneda, nMonto, nTipoPersona, bParticipaCamp, sMovNro, , gdFecSis, lnTpoPrograma, txtCuenta.Age, nPuntosRef, nCondicion, , nPTotalAcumulado)
            If bParticipaCamp Then
                cTextoDatos = "#" & IIf(bParticipaCamp, "1", "0") & "." & CStr(nPuntosRef) & "$" & CStr(nCondicion) & "_" & CStr(nPTotalAcumulado) & "&"
                lsmensaje = cTextoDatos
            End If
        End If
    End If
    'end ande
    
    Select Case nProducto
        Case gCapAhorros
            Dim clsCapD As COMDCaptaGenerales.DCOMCaptaGenerales
            Dim oCred As COMDCredito.DCOMCredito
            '********* AVMM-PLAZO FIJO.... **************
'            Set oCred = New COMDCredito.DCOMCredito
'            If oCred.VerificarClienteCreditos(gsPersCod) Then
'               Set clsCapD = New COMDCaptaGenerales.DCOMCaptaGenerales
'               clsCapD.BloqueoDesPlazoFijo Me.txtCuenta.NroCuenta, 1
'               MsgBox "Cliente posee Creditos Pendientes en Judicial...Se Bloqueara la Cuenta", vbInformation, "Aviso"
'               Exit Sub
'            Else
            '***Modificado por ELRO el 20130724, según TI-ERS079-2013****
            'clsCap.CapCancelaCuentaAho sCuenta, sMovNro, sGlosa, nOperacion, , gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628")
            If nOperacion = gAhoCancAct Then
                clsCap.CapCancelaCuentaAho sCuenta, sMovNro, sGlosa, nOperacion, , gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"), CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla
                'ALPA20130930********************************
                If gnMovNro = 0 Then
                    MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
                    Exit Sub
                End If
                '*********************************************
            'RIRO20131230 ERS137
            ElseIf nOperacion = gAhoCancTransfAbCtaBco Then
                clsCap.CapCancelaCuentaAho sCuenta, sMovNro, getGlosa, nOperacion, , gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"), CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, lblMonComisionTransf.Caption, Mid(txtBancoTrans.Text, 4, 13), sMovNroCom, txtCuentaTrans.Text, getTitular, chkTransfEfectivo.value
                'ALPA20130930********************************
                If gnMovNro = 0 Then
                    MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
                    Exit Sub
                End If
                '*********************************************
                
            Else
                clsCap.CapCancelaCuentaAho sCuenta, sMovNro, sGlosa, nOperacion, , gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"), psRegla:=ObtenerRegla
                'ALPA20130930********************************
                If gnMovNro = 0 Then
                    MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
                    Exit Sub
                End If
                '*********************************************
            End If
            '***Fin Modificado por ELRO el 20130724, según TI-ERS079-2013
'            End If
        Case gCapCTS
            If bDocumento Then
                clsCap.CapCancelaCuentaCTS sCuenta, sMovNro, sGlosa, nOperacion, sNroDoc, TpoDocNotaCargo, gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, psRegla:=ObtenerRegla
            Else
                '***Modificado por ELRO el 20130724, según TI-ERS079-2013****
                'clsCap.CapCancelaCuentaCTS sCuenta, sMovNro, sGlosa, nOperacion, , , gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro 'nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"
                If nOperacion = gCTSCancEfec Then
                    clsCap.CapCancelaCuentaCTS sCuenta, sMovNro, sGlosa, nOperacion, , , gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628")
                
                'RIRO20131226 ERS137
                ElseIf nOperacion = gCTSCancTransfBco Then
                    clsCap.CapCancelaCuentaCTS sCuenta, sMovNro, getGlosa, nOperacion, , , gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, , , lblMonComisionTransf.Caption, Mid(txtBancoTrans.Text, 4, 13), sMovNroCom, txtCuentaTrans.Text, getTitular, chkTransfEfectivo.value, Trim(lblNombreBancoTrans.Caption)

                Else
                    ' RIRO20131102 Se agregó parámetro "ObtenerRegla"
                    clsCap.CapCancelaCuentaCTS sCuenta, sMovNro, getGlosa, nOperacion, , , gsNomAge, sLpt, sReaPersLavDinero, gsCodCMAC, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, , ObtenerRegla, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628")    'nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"
                End If
                '***Fin Modificado por ELRO el 20130724, según TI-ERS079-2013
            End If
    End Select
    If gnMovNro > 0 Then
        'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
        Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
    End If
    Set clsCap = Nothing
    Set loLavDinero = Nothing
    'WIOR 20130301 ************************************************************
    If fbPersonaReaAhorros And gnMovNro > 0 Then
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuenta), fnCondicion
        fbPersonaReaAhorros = False
    End If
    'WIOR FIN *****************************************************************
        '***ADD BY MARG 20171222 ERS065-2017 ----SUBIDO DESDE LA 60***
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
        '***END MARG********************
    '*****BRGO 20110914 *****************************************************
    If gbITFAplica = True And CCur(lblITF.Caption) > 0 Then
       Call oMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption)) 'BRGO 20110914
    End If
    Set oMov = Nothing
    '*** End BRGO *****************
    
      'APRI20190601 RFC1902040001
        If nProducto = gCapCTS And nRegistraMotivo Then
         frmCapMovCancelacion.GuardarMotivoCancelacion sCuenta, nMotivo, cGlosaMo
        End If
    'END APRI
    'APRI20190109 ERS077-2018
    If (nOperacion = gAhoCancAct And sNumTarj <> "") Or nOperacion = gCTSCancEfec Then
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
    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
    'FIN
    cmdCancelar_Click
End If
Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub
End Sub
'FRHU 20141105 TIC1411050012
Private Sub cmdMostrarFirma_Click()
    With grdCliente
        If .TextMatrix(.row, 1) = "" Then Exit Sub
        'Call frmPersonaFirma.Inicio(Trim(.TextMatrix(.row, 1)), Trim(txtCuenta.Age), True) 'FRHU 20141115 MEMO-2766-2014
        Call frmPersonaFirma.Inicio(Trim(.TextMatrix(.row, 1)), Trim(txtCuenta.Age), True, , , txtCuenta.NroCuenta) 'FRHU 20141115 MEMO-2766-2014
    End With
End Sub
'FIN FRHU 20141105
Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub Finaliza_Verifone5000()
        If Not GmyPSerial Is Nothing Then
            GmyPSerial.Disconnect
            Set GmyPSerial = Nothing
        End If
End Sub

Private Sub cmdVerRegla_Click()
    If strReglas <> "" Then
        Call frmCapVerReglas.Inicia(strReglas)
    Else
        MsgBox "Cuenta no tiene reglas definidas", vbInformation
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
    'sCaption = Me.Caption
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Set clsGen = New COMDConstSistema.DCOMGeneral
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        bRetSinTarjeta = clsGen.GetPermisoEspecialUsuario(gCapPermEspRetSinTarj, gsCodUser, gsDominio)
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
    
    
If KeyCode = vbKeyF11 And txtCuenta.Enabled = True Then 'F11
        'Dim nPuerto As COMDConstantes.TipoPuertoSerial
        Dim sNumTar As String
        Dim sClaveTar As String
        Dim nErr As Integer
        Dim nEstado As COMDConstantes.CaptacTarjetaEstado
        Dim sCaption As String
        Dim sMaquina As String
        sMaquina = GetComputerName
        sCaption = Me.Caption
        Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
        MsgBox "Pase la tarjeta.", vbInformation, "AVISO"
        sNumTar = ""
        sNumTar = GetNumTarjeta_ACS
        If Len(sNumTar) <> 16 Then
            MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
            Me.Caption = sCaption
            Set clsGen = Nothing
            Exit Sub
        End If
        Me.Caption = "Ingrese la Clave de la Tarjeta."
        MsgBox "Ingrese la Clave de la Tarjeta.", vbInformation, "AVISO"
        Dim lnResult As ResultVerificacionTarjeta
        
        'Set clsGen = New COMDConstSistema.DCOMGeneral
        'Select Case clsGen.ValidaTarjeta(sNumTar, sClaveTar)
        Select Case GetClaveTarjeta_ACS(sNumTar, 1)
            Case gClaveValida
                    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
                    Dim rsTarj As New ADODB.Recordset
                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    
                    'Dim rsTarj As ADODB.Recordset
                    Set rsTarj = New ADODB.Recordset
                    Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
                    Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
                    Set rsTarj = ObjTarj.Get_Datos_Tarj(sNumTar)
                    
                    'Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
                    If rsTarj.EOF And rsTarj.BOF Then
                        MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
                        Me.Caption = sCaption
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
                            Exit Sub
                        End If
                        
                        Dim rsPers As New ADODB.Recordset
                        Dim sCta As String, sProducto As String, sMoneda As String
                        Dim clsCuenta As UCapCuenta
                        
                       Set rsPers = clsMant.GetCuentasPersona(rsTarj("cPersCod"), nProducto)
                        'Set rsPers = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
                        
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
                'ppoa Modificacion
                MsgBox "Clave Incorrecta", vbInformation, "Aviso"
        End Select
                                                                       
        Set clsGen = Nothing
        Me.Caption = "Captaciones - Cargo - Ahorros " & sOperacion
End If

    '**DAOR 20081125, Para tarjetas ***********************
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, nProducto)
        If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    '*******************************************************
    
Set clsGen = Nothing
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

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub grdCliente_DblClick()
    'FRHU 20141106 Observación
'    Dim R As ADODB.Recordset
'    Dim ssql As String
'    Dim clsFirma As COMDCaptaGenerales.DCOMCaptaMovimiento 'DCapMovimientos
'    Set clsFirma = New COMDCaptaGenerales.DCOMCaptaMovimiento
'
'    If Me.grdCliente.TextMatrix(grdCliente.row, 1) = "" Then Exit Sub
'
'    Set R = New ADODB.Recordset
'    Set R = clsFirma.GetFirma(Me.grdCliente.TextMatrix(grdCliente.row, 1))
'    If R.BOF Or R.EOF Then
'       Set R = Nothing
'       MsgBox "La visualización del DNI no esta Disponible", vbOKOnly + vbInformation, "AVISO"
'       Exit Sub
'    End If
'
'    If R.RecordCount > 0 Then
'       If IsNull(R!iPersFirma) = True Then
'         MsgBox "El cliente no posse Firmas", vbInformation, "Aviso"
'         Exit Sub
'       End If
'       ' Call frmMuestraFirma.IDBFirma.CargarFirma(R)
'       frmMuestraFirma.psCodCli = Me.grdCliente.TextMatrix(grdCliente.row, 1)
'       Set frmMuestraFirma.rs = R
'    End If
'    Set clsFirma = Nothing
'    frmMuestraFirma.Show 1
    
    Dim sPersCod As String

    sPersCod = grdCliente.TextMatrix(grdCliente.row, 1)
    If sPersCod = "" Then Exit Sub
    MuestraFirma sPersCod, gsCodAge
    'FIN FRHU 20141106
End Sub

Private Sub grdCliente_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If pnCol = 8 And Trim(grdCliente.TextMatrix(pnRow, 7)) = "PJ" Then
        grdCliente.TextMatrix(grdCliente.row, 8) = False
    End If
End Sub

Private Sub grdCliente_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdCliente.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = txtCuenta.NroCuenta
        ObtieneDatosCuenta sCta
        'frmSegSepelioAfiliacion.Inicio sCta
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
    '***Modificado por ELRO el 20130724, según TI-ERS079-2013****
    'cmdGrabar.SetFocus
    If cboMedioRetiro.Visible Then
        cboMedioRetiro.SetFocus
    Else
        cmdGrabar.SetFocus
    End If
    '***Fin Modificado por ELRO el 20130724, según TI-ERS079-2013
End If
End Sub

Private Sub txtglosaTrans_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        If cmdGrabar.Enabled And cmdGrabar.Visible Then
            cmdGrabar.SetFocus
        End If
    End If
End Sub

Private Sub txtMonto_Change()
    If gbITFAplica And nProducto <> gCapCTS Then       'Filtra para CTS
        If TxtMonto.value > gnITFMontoMin Then
            If Not lbITFCtaExonerada And ckMismoTitular.value = 0 Then 'RIRO20131212 ERS137 "ckMismoTitular"
                'If nOperacion = gAhoCancTransfInact Or nOperacion = gAhoCancTransfAct Or nOperacion = gAhoRetTransf Or nOperacion = gAhoDepTransf Or nOperacion = gAhoDepPlanRRHH Or nOperacion = gAhoDepPlanRRHHAdelantoSueldos Then
                If nOperacion = gAhoDepPlanRRHH Or nOperacion = gAhoDepPlanRRHHAdelantoSueldos Then
                    Me.lblITF.Caption = Format(0, "#,##0.00")
                ElseIf nProducto = gCapAhorros And (nOperacion <> gAhoDepChq Or nOperacion <> gCMACOAAhoDepChq Or nOperacion <> gCMACOTAhoDepChq) Then
                    Me.lblITF.Caption = Format(fgITFCalculaImpuesto(TxtMonto.value), "#,##0.00")
                Else
                    Me.lblITF.Caption = Format(fgITFCalculaImpuesto(TxtMonto.value), "#,##0.00")
                End If

                '*** BRGO 20110908 ************************************************
                nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
                If nRedondeoITF > 0 Then
                    Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
                End If
                '*** END BRGO
            Else
                Me.lblITF.Caption = "0.00"
            End If
            If nOperacion = gAhoRetOPCanje Or nOperacion = gAhoRetOPCertCanje Or nOperacion = gAhoRetFondoFijoCanje Then
                Me.lblTotal.Caption = Format(0, "#,##0.00")
            ElseIf nOperacion = gAhoDepChq Then
                Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption), "#,##0.00")
            Else
                If nProducto = gCapAhorros And gbITFAsumidoAho Then
                    Me.lblTotal.Caption = Format(TxtMonto.value - CCur(Me.lblMonComision.Caption), "#,##0.00")
                ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                    Me.lblTotal.Caption = Format(TxtMonto.value, "#,##0.00")
                Else
                    Me.lblTotal.Caption = Format(TxtMonto.value - CCur(Me.lblITF.Caption) - CCur(Me.lblMonComision.Caption), "#,##0.00")
                End If
            End If
            'If bInstFinanc Then lblITF.Caption = "0.00" 'JUEZ 20140414
        End If
    Else
        Me.lblITF.Caption = Format(0, "#,##0.00")
        
        If nProducto = gCTSDepChq Then
            Me.lblTotal.Caption = Format(0, "#,##0.00")
        Else
            Me.lblTotal.Caption = Format(Me.TxtMonto.value - CCur(Me.lblMonComision.Caption), "#,##0.00")
        End If
    End If
    
    If TxtMonto.value = 0 Then
        Me.lblITF.Caption = "0.00"
        Me.lblTotal.Caption = "0.00"
    End If
    
    If chkVBComision.Visible Then
        chkVBComision_Click
    End If
    
    If chkTransfEfectivo.value = 0 Then
        Me.lblTotal.Caption = Format(CDbl(Me.lblTotal.Caption) - CDbl(lblMonComisionTransf.Caption), "#,##0.00")
    End If
    'JUEZ 20141006 *************************************
    If bInstFinanc Then
        lblTotal.Caption = Format(lblTotal.Caption + CCur(Me.lblITF.Caption), "#,##0.00")
        lblITF.Caption = "0.00"
    End If
    'END JUEZ ******************************************
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub MuestraFirmas(ByVal sCuenta As String)
    Dim i As Integer
    Dim sPersona As String
    
    If nPersoneria <> PersPersoneria.gPersonaNat Then
        For i = 1 To Me.grdCliente.Rows - 1
            If Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersRepSuplente Or Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersRepTitular Then
                sPersona = grdCliente.TextMatrix(i, 1)
                'MuestraFirma sPersona
            End If
        Next i
    Else
        For i = 1 To Me.grdCliente.Rows - 1
            If Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersTitular Then
                sPersona = grdCliente.TextMatrix(i, 1)
                'MuestraFirma sPersona
            End If
        Next i
    End If
End Sub

Private Function Cargousu(ByVal NomUser As String) As String
    Dim cCon As COMDConstSistema.DCOMUAcceso
    Dim rs As New ADODB.Recordset
    
    Set cCon = New COMDConstSistema.DCOMUAcceso
    Set rs = cCon.Cargousu(NomUser)
    
    If Not (rs.EOF And rs.BOF) Then
        Cargousu = rs(0)
    End If
    Set cCon = Nothing
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
nMonto = TxtMonto.value
nMoneda = CLng(Mid(sCuenta, 9, 1))
'Obtiene los grupos al cual pertenece el usuario
Set oPers = New COMDPersona.UCOMAcceso
    gsGrupo = oPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
Set oPers = Nothing
 
'Verificar Montos
Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    'Set rs = ocapaut.ObtenerMontoTopNivAutRetCan(gsGrupo, "2", gsCodAge)
    Set rs = oCapAut.ObtenerMontoTopNivAutRetCan(gsGrupo, "2", gsCodAge, gsCodPersUser) 'RIRO20141106 ERS159
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
    oCapAutN.NuevaSolicitudAutorizacion sCuenta, "2", nMonto, gdFecSis, gsCodAge, gsCodUser, nMoneda, gOpeAutorizacionRetiro, sNivel, sMovNroAut
    MsgBox "Solicitud Registrada, comunique a su Admnistrador para la Aprobación..." & Chr$(10) & _
        " No salir de esta operación mientras se realice el proceso..." & Chr$(10) & _
        " Porque sino se procedera a grabar otra Solicitud...", vbInformation, "Aviso"
    VerificarAutorizacion = False
Else
    'Valida el estado de la Solicitud
    If Not oCapAutN.VerificarAutorizacion(sCuenta, "2", nMonto, sMovNroAut, lsmensaje) Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        VerificarAutorizacion = False
    Else
        VerificarAutorizacion = True
    End If
End If
Set oCapAutN = Nothing
End Function


Private Function ValidaCreditosPendientes(psCtaCod As String) As Boolean
    Dim rsCred As ADODB.Recordset
    Dim lnI As Long
    Dim bPendiente As Boolean
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCapMant As COMDCaptaGenerales.DCOMCaptaGenerales 'DCapMantenimiento
    Set clsCapMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim clsCont As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Dim lsMovNro As String
    
    If nOperacion = gPFCancEfec Or nOperacion = gPFCancTransf Then
        If clsCapMov.BuscaCreditosPendientesPago(psCtaCod) Then
            bPendiente = False
            For lnI = 1 To Me.grdCliente.Rows - 1
                If Trim(Right(grdCliente.TextMatrix(lnI, 3), 5)) = gCapRelPersTitular Then
                    
                    Set rsCred = New ADODB.Recordset
                    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    Set rsCred = clsCapMov.GetCreditosPendientes(Me.grdCliente.TextMatrix(lnI, 1), gdFecSis)
                    Set clsCapMov = Nothing
                    If Not rsCred Is Nothing Then
                        If Not (rsCred.EOF And rsCred.BOF) Then
                            bPendiente = True
                            'frmCredPendPago.Inicia rsCred
                        End If
                    End If
                    Set rsCred = Nothing
                End If
            Next
            
            If bPendiente Then
                lsMovNro = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                'Bloquea la cuenta
                    clsCapMant.NuevoBloqueoRetiro psCtaCod, gCapMotBlqRetGarantia, "POR CREDITO PENDIENTES DE PAGO", lsMovNro
                    clsCapMant.ActualizaEstadoCuenta psCtaCod, gCapEstBloqRetiro
                Set clsCapMant = Nothing
                Set clsCont = Nothing
                ValidaCreditosPendientes = False
                Exit Function
            End If
        End If
    End If

    ValidaCreditosPendientes = True
End Function

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

' Agregado Por RIRO el 20130501, Proyecto de ahorros
Private Function ValidarReglasPersonas() As Boolean
    Dim arreReglas() As String
    Dim intRegla As Integer
    Dim i As Integer
    arreReglas = Split(strReglas, "-")
   
    
    Dim bAprobado As Boolean
'    Dim intRegla As Integer
    
    
    
    For i = 1 To grdCliente.Rows - 1
        If grdCliente.TextMatrix(i, 8) = "." Then
            'If cont = 0 Then
            intRegla = intRegla + AscW(grdCliente.TextMatrix(i, 7))
            'Else
             '   regla = regla & "+" & grdCliente.TextMatrix(i, 7)
            'End If
        End If
    Next
    
    
    Dim V As Variant
    Dim vRF As Variant
    Dim arrReglasFijadas() As String
    Dim intReglasFijadas As Integer
    
    For Each V In arreReglas
        arrReglasFijadas = Split(V, "+")
        intReglasFijadas = 0
        For Each vRF In arrReglasFijadas
            intReglasFijadas = intReglasFijadas + AscW(vRF)
        Next
        
        If intRegla = intReglasFijadas Then
            ValidarReglasPersonas = True
            Exit Function
        End If
    Next
    
    ValidarReglasPersonas = False
    
    
End Function

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

' RIRO20131226 ERS137
Private Sub MostrarControlesTransferencia()

    Dim clsBanco As COMNCajaGeneral.NCOMCajaCtaIF
    Dim rsBanco As New ADODB.Recordset
        
    Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
    Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "01%", MuestraInstituciones)
    Set clsBanco = Nothing
    txtBancoTrans.rs = rsBanco
    
    Dim rsConstante As ADODB.Recordset
    Dim oConstante As COMDConstSistema.DCOMGeneral
    Set oConstante = New COMDConstSistema.DCOMGeneral
    Set rsConstante = oConstante.GetConstante("10032", , "'20[^0]'")
    CargaCombo cbPlazaTrans, rsConstante
    cbPlazaTrans.ListIndex = 0
    Set rsConstante = Nothing
    Set oConstante = Nothing
    fraDocumento.Visible = False
    fraDocumentoTrans.Visible = True
    lblMonComisionTransf.Visible = True
    lblComisionTransf.Visible = True
    lblMonComisionTransf.Top = 1725
    chkTransfEfectivo.Top = 1790
    lblComisionTransf.Top = 1785
    cboMedioRetiro.Visible = True
    lblMedioRetiro.Visible = True
    lblMonComision.Visible = False
    Label16.Caption = "Total Transf:"
End Sub

Private Sub txtTitular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCuentaTrans.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

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
        nPlaza = CDbl(Trim(Right(cbPlazaTrans.Text, 8)))
        nMoneda = Val(Mid(txtCuenta.NroCuenta, 9, 1))
        nTipo = 102 ' Emision
        nMonto = TxtMonto.value
        nComision = oDefinicion.getCalculaComision(idBanco, nPlaza, nMoneda, nTipo, nMonto, gdFecSis)
    End If
End If

lblMonComisionTransf.Caption = Format(Round(nComision, 2), "#0.00")

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
    sGlosa = "Banco destino: " & lblNombreBancoTrans.Caption & ", Titular: " & getTitular & ", " & Trim(txtGlosaTrans.Text)
    getGlosa = UCase(sGlosa)
End Function

Private Sub txtBancoTrans_EmiteDatos()
    lblNombreBancoTrans.Caption = Trim(txtBancoTrans.psDescripcion)
    If lblNombreBancoTrans.Caption <> "" Then
        cbPlazaTrans.SetFocus
    End If
    CalculaComision
    
End Sub

Private Sub txtBancoTrans_Click(psCodigo As String, psDescripcion As String)
CalculaComision
txtMonto_Change
End Sub

Private Sub cbPlazaTrans_Click()
    If ckMismoTitular.Visible And ckMismoTitular.Enabled Then ckMismoTitular.SetFocus
    CalculaComision
    txtMonto_Change
End Sub

Private Sub txtMonto_LostFocus()
   CalculaComision
   txtMonto_Change
End Sub

' FIN RIRO
