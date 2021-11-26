VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCapDepositoLote 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "frmCapDepositoLote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   10080
      TabIndex        =   87
      Top             =   2520
      Width           =   4920
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   360
         Width           =   2565
      End
      Begin VB.CommandButton cmdDocumento 
         Height          =   350
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   315
         Width           =   475
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
         TabIndex        =   92
         Top             =   330
         Width           =   1155
      End
      Begin VB.ComboBox cboMonedaBanco 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   705
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox txtCtaBanco 
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   88
         Top             =   330
         Visible         =   0   'False
         Width           =   1815
      End
      Begin SICMACT.TxtBuscar txtBanco 
         Height          =   315
         Left            =   3000
         TabIndex        =   90
         Top             =   330
         Width           =   1755
         _extentx        =   3096
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmCapDepositoLote.frx":030A
         psraiz          =   "BANCOS"
         appearance      =   1
         stitulo         =   ""
      End
      Begin VB.Label lblDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   393
         Width           =   915
      End
      Begin VB.Label lblOrdenPago 
         AutoSize        =   -1  'True
         Caption         =   "Orden Pago :"
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   98
         Top             =   1080
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
         TabIndex        =   97
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblEtqBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   2520
         TabIndex        =   96
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblCtaBanco 
         Caption         =   "Cta Banco :"
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   360
         Visible         =   0   'False
         Width           =   945
      End
   End
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
      Left            =   10080
      TabIndex        =   74
      Top             =   4920
      Visible         =   0   'False
      Width           =   4920
      Begin VB.ComboBox cbPlazaTrans 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox ckMismoTitular 
         Caption         =   "Mismo Titular"
         Height          =   315
         Left            =   1125
         TabIndex        =   78
         Top             =   1080
         Width           =   1365
      End
      Begin VB.TextBox txtCuentaTrans 
         Height          =   315
         Left            =   1125
         TabIndex        =   77
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtGlosaTrans 
         Height          =   735
         Left            =   1125
         TabIndex        =   76
         Top             =   1845
         Width           =   3660
      End
      Begin VB.TextBox txtTitular 
         Height          =   315
         Left            =   3015
         TabIndex        =   75
         Top             =   1080
         Width           =   1770
      End
      Begin SICMACT.TxtBuscar txtBancoTrans 
         Height          =   315
         Left            =   1125
         TabIndex        =   80
         Top             =   315
         Width           =   1815
         _extentx        =   3201
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmCapDepositoLote.frx":0336
         appearance      =   1
         stitulo         =   ""
      End
      Begin VB.Label lblNombreBanco 
         Caption         =   "Banco:"
         Height          =   285
         Left            =   135
         TabIndex        =   85
         Top             =   315
         Width           =   690
      End
      Begin VB.Label lblPlaza 
         Caption         =   "Plaza:"
         Height          =   285
         Left            =   135
         TabIndex        =   84
         Top             =   765
         Width           =   645
      End
      Begin VB.Label Label12 
         Caption         =   "Nro Cuenta:"
         Height          =   240
         Left            =   135
         TabIndex        =   83
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
         Left            =   3015
         TabIndex        =   82
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label lblGlosa 
         Caption         =   "Glosa:"
         Height          =   330
         Left            =   135
         TabIndex        =   81
         Top             =   1890
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   10080
      TabIndex        =   58
      Top             =   7800
      Width           =   3930
      Begin VB.CheckBox chkITFEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   62
         Top             =   1485
         Width           =   705
      End
      Begin VB.CheckBox chkVBComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   61
         Top             =   720
         Width           =   705
      End
      Begin VB.ComboBox cboMedioRetiro 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.CheckBox chkTransfEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   59
         Top             =   2280
         Visible         =   0   'False
         Width           =   705
      End
      Begin SICMACT.EditMoney EditMoney1 
         Height          =   315
         Left            =   1440
         TabIndex        =   63
         Top             =   1020
         Width           =   1890
         _extentx        =   3334
         _extenty        =   556
         font            =   "frmCapDepositoLote.frx":0362
         backcolor       =   12648447
         forecolor       =   192
         text            =   "0"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   495
         TabIndex        =   73
         Top             =   1065
         Width           =   540
      End
      Begin VB.Label Label9 
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
         Left            =   3390
         TabIndex        =   72
         Top             =   1035
         Width           =   315
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   495
         TabIndex        =   71
         Top             =   1485
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   495
         TabIndex        =   70
         Top             =   1890
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
         Left            =   2235
         TabIndex        =   69
         Top             =   1425
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
         TabIndex        =   68
         Top             =   1830
         Width           =   1890
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
         TabIndex        =   67
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label lblComision 
         AutoSize        =   -1  'True
         Caption         =   "Comision :"
         Height          =   195
         Left            =   495
         TabIndex        =   66
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblMedioRetiro 
         AutoSize        =   -1  'True
         Caption         =   "Medio de Retiro :"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   64
         Top             =   2220
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10500
      Left            =   45
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   45
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   18521
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Deposito en Lote"
      TabPicture(0)   =   "frmCapDepositoLote.frx":038E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dlgArchivo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraMonto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraGlosa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSalir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdGrabar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraCuenta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraTipo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraTranferecia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancelar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "pbProgres"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraCliente"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraCheque"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.Frame fraCheque 
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
         ForeColor       =   &H80000002&
         Height          =   2115
         Left            =   135
         TabIndex        =   99
         Top             =   7920
         Visible         =   0   'False
         Width           =   4410
         Begin VB.CommandButton cmdCheque 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCapDepositoLote.frx":03AA
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   315
            Width           =   475
         End
         Begin VB.Label lblChequeIFI 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   108
            Top             =   690
            Width           =   3465
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nro Doc :"
            Height          =   195
            Left            =   60
            TabIndex        =   107
            Top             =   330
            Width           =   690
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Banco :"
            Height          =   195
            Left            =   60
            TabIndex        =   106
            Top             =   690
            Width           =   555
         End
         Begin VB.Label lblChequeNro 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   105
            Top             =   315
            Width           =   1575
         End
         Begin VB.Label lblChequeMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   2880
            TabIndex        =   104
            Top             =   1080
            Width           =   1365
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   2220
            TabIndex        =   103
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lblChequeNroOpe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   102
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nro Ope :"
            Height          =   195
            Left            =   60
            TabIndex        =   101
            Top             =   1080
            Width           =   690
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Datos de Autorización"
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
         Height          =   2835
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   9570
         Begin VB.CommandButton cmdMostrarFirma 
            Caption         =   "Mostrar Firma"
            Height          =   315
            Left            =   7920
            TabIndex        =   44
            Top             =   2025
            Width           =   1440
         End
         Begin VB.CommandButton cmdVerRegla 
            Caption         =   "Ver Regla"
            Height          =   315
            Left            =   6360
            TabIndex        =   43
            Top             =   2025
            Visible         =   0   'False
            Width           =   1440
         End
         Begin SICMACT.FlexEdit grdCliente 
            Height          =   1755
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   9375
            _extentx        =   16536
            _extenty        =   3096
            cols0           =   9
            highlight       =   1
            allowuserresizing=   3
            visiblepopmenu  =   -1
            encabezadosnombres=   "#-Codigo-Nombre-Relacion-Direccion-ID-Firma Oblig-Grupo-Presente"
            encabezadosanchos=   "250-1500-3200-1500-0-0-0-1000-1000"
            font            =   "frmCapDepositoLote.frx":07EC
            font            =   "frmCapDepositoLote.frx":0814
            font            =   "frmCapDepositoLote.frx":083C
            font            =   "frmCapDepositoLote.frx":0864
            font            =   "frmCapDepositoLote.frx":088C
            fontfixed       =   "frmCapDepositoLote.frx":08B4
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X-X-X-8"
            listacontroles  =   "0-0-0-0-0-0-0-0-4"
            encabezadosalineacion=   "C-L-L-L-C-C-C-L-C"
            formatosedit    =   "0-0-0-0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbflexduplicados=   0
            colwidth0       =   255
            rowheight0      =   300
            tipobuspersona  =   1
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
            TabIndex        =   52
            Top             =   2040
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "# Firmas :"
            Height          =   195
            Left            =   3480
            TabIndex        =   51
            Top             =   2040
            Visible         =   0   'False
            Width           =   525
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
            TabIndex        =   50
            Top             =   2070
            Width           =   1800
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   180
            TabIndex        =   49
            Top             =   2123
            Width           =   960
         End
         Begin VB.Label Label4 
            Caption         =   "Alias de la Cuenta:"
            Height          =   225
            Left            =   180
            TabIndex        =   48
            Top             =   2490
            Width           =   1425
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
            TabIndex        =   47
            Top             =   2430
            Width           =   7755
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Mínimo Firmas :"
            Height          =   195
            Left            =   4680
            TabIndex        =   46
            Top             =   2040
            Visible         =   0   'False
            Width           =   375
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
            Left            =   5280
            TabIndex        =   45
            Top             =   2040
            Visible         =   0   'False
            Width           =   465
         End
      End
      Begin ComctlLib.ProgressBar pbProgres 
         Height          =   195
         Left            =   2295
         TabIndex        =   38
         Top             =   10170
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancela&r"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   7
         Top             =   10080
         Width           =   1095
      End
      Begin VB.Frame fraTranferecia 
         Caption         =   "Transferencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   2115
         Left            =   135
         TabIndex        =   23
         Top             =   7920
         Visible         =   0   'False
         Width           =   4410
         Begin VB.ComboBox cboTransferMoneda 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   195
            Width           =   1575
         End
         Begin VB.CommandButton cmdTranfer 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCapDepositoLote.frx":08DA
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   555
            Width           =   475
         End
         Begin VB.TextBox txtTransferGlosa 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   840
            MaxLength       =   255
            TabIndex        =   24
            Top             =   1290
            Width           =   3465
         End
         Begin VB.Label lbltransferBco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   37
            Top             =   930
            Width           =   3465
         End
         Begin VB.Label lbltransferN 
            AutoSize        =   -1  'True
            Caption         =   "Nro Doc :"
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   570
            Width           =   690
         End
         Begin VB.Label lbltransferBcol 
            AutoSize        =   -1  'True
            Caption         =   "Banco :"
            Height          =   195
            Left            =   60
            TabIndex        =   35
            Top             =   930
            Width           =   555
         End
         Begin VB.Label lblTrasferND 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   34
            Top             =   555
            Width           =   1575
         End
         Begin VB.Label lblTransferMoneda 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   60
            TabIndex        =   33
            Top             =   225
            Width           =   585
         End
         Begin VB.Label lblTransferGlosa 
            AutoSize        =   -1  'True
            Caption         =   "Glosa :"
            Height          =   195
            Left            =   60
            TabIndex        =   32
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lblTTCC 
            Caption         =   "TCC"
            Height          =   285
            Left            =   3120
            TabIndex        =   31
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label11 
            Caption         =   "TCV"
            Height          =   285
            Left            =   3120
            TabIndex        =   30
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblTTCCD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3570
            TabIndex        =   29
            Top             =   165
            Width           =   735
         End
         Begin VB.Label lblTTCVD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3570
            TabIndex        =   28
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblMonTra 
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
            Left            =   2925
            TabIndex        =   27
            Top             =   1680
            Width           =   1365
         End
         Begin VB.Label lblSimTra 
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
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   2370
            TabIndex        =   26
            Top             =   1680
            Width           =   300
         End
         Begin VB.Label lblEtiMonTra 
            AutoSize        =   -1  'True
            Caption         =   "Monto Transacción"
            Height          =   195
            Left            =   870
            TabIndex        =   25
            Top             =   1710
            Width           =   1380
         End
      End
      Begin VB.Frame fraTipo 
         Caption         =   "Datos Generales"
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
         Height          =   1080
         Left            =   135
         TabIndex        =   19
         Top             =   405
         Width           =   9570
         Begin VB.TextBox txtCuentaInstitucion 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3690
            TabIndex        =   41
            Top             =   720
            Width           =   1995
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   270
            Width           =   1455
         End
         Begin VB.TextBox lblInst 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   3690
         End
         Begin SICMACT.TxtBuscar txtInstitucion 
            Height          =   315
            Left            =   3720
            TabIndex        =   54
            Top             =   240
            Width           =   1995
            _extentx        =   3519
            _extenty        =   556
            appearance      =   1
            appearance      =   1
            font            =   "frmCapDepositoLote.frx":0D1C
            appearance      =   1
            tipobusqueda    =   3
            stitulo         =   ""
            tipobuspers     =   1
            enabledtext     =   0
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
            Left            =   5760
            TabIndex        =   53
            Top             =   600
            Width           =   3705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
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
            Left            =   2640
            TabIndex        =   40
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
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
            Left            =   180
            TabIndex        =   22
            Top             =   315
            Width           =   810
         End
         Begin VB.Label lblInstEtq 
            AutoSize        =   -1  'True
            Caption         =   "Institución :"
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
            Left            =   2610
            TabIndex        =   21
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.Frame fraCuenta 
         Caption         =   "Datos Cuenta"
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
         Height          =   3480
         Left            =   135
         TabIndex        =   17
         Top             =   4440
         Width           =   9570
         Begin VB.CommandButton cmdFormato 
            Caption         =   "&Formato"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   135
            TabIndex        =   1
            Top             =   2970
            Width           =   915
         End
         Begin VB.TextBox txtArchivo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   2970
            Width           =   4245
         End
         Begin VB.CommandButton cmdCargar 
            Caption         =   "&Cargar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6795
            TabIndex        =   4
            Top             =   2970
            Width           =   840
         End
         Begin VB.CommandButton cmdBusca 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6210
            TabIndex        =   3
            Top             =   2970
            Width           =   495
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8595
            TabIndex        =   6
            Top             =   2970
            Width           =   855
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7695
            TabIndex        =   5
            Top             =   2970
            Width           =   855
         End
         Begin SICMACT.FlexEdit grdCuenta 
            Height          =   2655
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   9330
            _extentx        =   16457
            _extenty        =   4683
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            encabezadosnombres=   "#-Cod Cliente-Cuenta-Nombre-DOI-Monto-campo1"
            encabezadosanchos=   "500-1500-1800-4000-1200-1200-0"
            font            =   "frmCapDepositoLote.frx":0D44
            font            =   "frmCapDepositoLote.frx":0D6C
            font            =   "frmCapDepositoLote.frx":0D94
            font            =   "frmCapDepositoLote.frx":0DBC
            font            =   "frmCapDepositoLote.frx":0DE4
            fontfixed       =   "frmCapDepositoLote.frx":0E0C
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-1-X-X-X-5-X"
            textstylefixed  =   4
            listacontroles  =   "0-1-0-0-0-0-0"
            encabezadosalineacion=   "C-L-L-L-L-C-C"
            formatosedit    =   "0-0-0-0-0-2-2"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbflexduplicados=   0
            lbbuscaduplicadotext=   -1
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Archivo :"
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
            Left            =   1170
            TabIndex        =   18
            Top             =   3015
            Width           =   780
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         Top             =   10080
         Width           =   960
      End
      Begin VB.CommandButton cmdSalir 
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
         Height          =   315
         Left            =   8730
         TabIndex        =   9
         Top             =   10080
         Width           =   960
      End
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
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
         Height          =   2115
         Left            =   4590
         TabIndex        =   15
         Top             =   7920
         Width           =   2430
         Begin RichTextLib.RichTextBox txtGlosa 
            Height          =   1770
            Left            =   90
            TabIndex        =   16
            Top             =   225
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   3122
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmCapDepositoLote.frx":0E32
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
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
         ForeColor       =   &H8000000D&
         Height          =   2115
         Left            =   7080
         TabIndex        =   13
         Top             =   7920
         Width           =   2655
         Begin SICMACT.EditMoney txtMonto 
            Height          =   315
            Left            =   720
            TabIndex        =   57
            Top             =   360
            Width           =   1365
            _extentx        =   2408
            _extenty        =   556
            font            =   "frmCapDepositoLote.frx":0EAC
            appearance      =   0
            backcolor       =   12648447
            text            =   "0.00"
            enabled         =   -1
         End
         Begin VB.Label lblMon 
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
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   2160
            TabIndex        =   39
            Top             =   360
            Width           =   300
         End
         Begin VB.Label Label1 
            Caption         =   "Monto:"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   14
            Top             =   405
            Width           =   555
         End
      End
      Begin MSComDlg.CommonDialog dlgArchivo 
         Left            =   1575
         Top             =   9900
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCapDepositoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************************************************************************
'* NOMBRE         : "frmCapDepositoLote"
'* DESCRIPCION    : Formulario creado para el pago en lote.
'* CREACION       : RIRO, 20140430 10:00 AM
'*********************************************************************************************************************************************************
Option Explicit

Private fnProducto As Producto
Private fnOpeCod As CaptacOperacion
Private fsDescOperacion As String
Private fnMoneda As COMDConstantes.Moneda
Private bCargaLote As Boolean
Private nNroDeposito As Integer
Private fnMovNroRVD As Long
Private lnMovNroTransfer As Long
Private nmoneda As COMDConstantes.Moneda

'***MARG ERS065-2017*************************************
Private PbPreIngresado As Boolean
Private pnMovNro As Long
Dim sCuenta As String

Dim strReglas As String
Dim bDocumento As Boolean
Dim nDocumento As COMDConstantes.tpoDoc
Dim sPersCodCMAC As String
Dim sNombreCMAC As String
Dim sOperacion As String
Public nProducto As COMDConstantes.Producto
Dim nSaldoCuenta As Double
Dim nPersoneria As COMDConstantes.PersPersoneria
'Dim nmoneda As COMDConstantes.Moneda
Dim nEstadoT As COMDConstantes.CaptacEstado
Dim lnTpoPrograma As Integer
Dim nParCantRetLib As Integer
Dim nParDiasVerifRegSueldo As Integer
Dim nParUltRemunBrutas As Integer
Dim nComiMaxOpe As Double
Dim bValidaCantRet As Boolean
Dim pbOrdPag As Boolean
Dim lsDescTpoPrograma As String
Dim nSaldoRetiro As Double
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim sNumTarj As String
Dim nRedondeoITF As Double '
Dim cGetValorOpe As String
Dim nComisionVB As Double
Dim lsTieneTarj As String
Dim sTipoCuenta As String
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim bProcesoNuevo As Boolean
Dim bFirmaObligatoria As Boolean
Dim bPresente As Boolean
Dim aPersonasInvol() As String
Dim bFirmasPendientes As Boolean
Dim bInstFinanc As Boolean
Dim sMovNroAut As String
Dim cVioFirma As String
Dim lbITFCtaExonerada As Boolean

Private bRetiroExitoso As Boolean
Private bDepositoExitoso As Boolean
'***END MARG**********************************************
Dim oDocRec As UDocRec 'CTI620210606

'CTI6 20210606***
Private Sub cmdCheque_Click()
    Dim oForm As New frmChequeBusqueda
    Dim lnOperacion As TipoOperacionCheque
    Dim oDocRecTmp As UDocRec

    On Error GoTo ErrCargaDocumento
    If gsOpeCod = gAhoDepositoHaberesEnLoteChq Then
        lnOperacion = AHO_DepositoHaberesLote
    Else
        lnOperacion = Ninguno
    End If

    Set oDocRecTmp = oForm.Iniciar(nmoneda, lnOperacion)
    Set oForm = Nothing
    If Len(Trim(oDocRecTmp.fsNroDoc)) = 0 Then Exit Sub
    
    Set oDocRec = oDocRecTmp
    Call setDatosCheque
    
    Exit Sub
ErrCargaDocumento:
    MsgBox "Ha sucedido un error al cargar los datos del Documento", vbCritical, "Aviso"
End Sub

Private Sub setDatosCheque()
    lblChequeNro.Caption = oDocRec.fsNroDoc
    lblChequeIFI.Caption = oDocRec.fsPersNombre
    lblChequeNroOpe.Caption = oDocRec.fnNroCliLote
    lblChequeMonto.Caption = Format(oDocRec.fnMonto, gsFormatoNumeroView)
    txtGlosa.Text = oDocRec.fsGlosa
End Sub

Private Function ValidaSeleccionCheque() As Boolean
    ValidaSeleccionCheque = True
    If oDocRec Is Nothing Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
    If Len(Trim(oDocRec.fsNroDoc)) = 0 Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
End Function
'END CTI6 *******

Private Sub cmdsalir_Click()
    If MsgBox("¿Deseas salir de la formulario?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub cboMoneda_Click()
    nmoneda = CLng(Right(cboMoneda.Text, 1))
    If nmoneda = gMonedaNacional Then
        txtMonto.BackColor = &HC0FFFF
        lblMon.Caption = "S/."
    ElseIf nmoneda = gMonedaExtranjera Then
        txtMonto.BackColor = &HC0FFC0
        lblMon.Caption = "US$"
    End If
    If fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, Trim(Right(cboMoneda.Text, 5)))
        SetDatosTransferencia "", "", "", 0, -1, ""
        nNroDeposito = 0
        fnMovNroRVD = 0
        lnMovNroTransfer = 0
    End If
    
    LimpiarCheque
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtInstitucion.SetFocus
    End If
End Sub

Private Sub cboTransferMoneda_Click()
    If nmoneda = gMonedaNacional Then
        lblMonTra.BackColor = &HC0FFFF
        lblSimTra.Caption = "S/."
    ElseIf nmoneda = gMonedaExtranjera Then
        lblMonTra.BackColor = &HC0FFC0
        lblSimTra.Caption = "US$"
    End If
End Sub

Private Sub cmdAgregar_Click()
    If bCargaLote Then
        If MsgBox("Al usar esta opción, se limpiará el Grid, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            LimpiarGrdCuenta
            txtArchivo.Text = ""
            bCargaLote = False
            txtMonto.Text = "0.00"
            grdCuenta.ColumnasAEditar = "X-1-X-X-X-5"
        Else
            Exit Sub
        End If
    End If
    grdCuenta.lbEditarFlex = True
    grdCuenta.Col = 1
    grdCuenta_RowColChange
    grdCuenta.AdicionaFila
    grdCuenta.SetFocus
    SendKeys "{Enter}"
End Sub

Private Sub cmdBusca_Click()
   On Error GoTo error_handler
    
    txtArchivo.Text = Empty
    
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        txtArchivo.Text = dlgArchivo.FileName
        If fnProducto = gCapCTS Then
            txtArchivo.Locked = True
            cmdAgregar.Enabled = False
        End If
    Else
        txtArchivo.Text = "NO SE ABRIO NINGUN ARCHIVO"
        Exit Sub
    End If
    cmdCargar.Enabled = True
    
     Exit Sub
error_handler:
    
    If err.Number = 32755 Then
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        MsgBox "Error al momento de seleccionar el archivo", vbCritical, "Aviso"
    End If
End Sub

Private Sub cmdCancelar_Click()
Limpiar
End Sub

Private Sub cmdCargar_Click()

    Dim lsNroDoc As String
    Dim lsPersCod As String
    Dim lsNombre As String
    Dim lsMonto As String
    Dim lnPersoneria As Integer
    Dim lnOP As Integer
    Dim lnTasaCli As Double
    Dim lnMonApeCli As Double
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, fila As Integer
    Dim psArchivoAGrabar As String
    Dim sCad As String
    Dim nFila As Long
    Dim lsNomArch As String
    Dim lsDire As String
    Dim lsTipDOI As String
    Dim X As Integer
    Dim Y As Integer, Z As Integer
    Dim bMayorEdad As Boolean, bFormato As Boolean
    Dim psArchivoAGrabarMenores As String
    Dim psArchivoAGrabarPersJurid As String
    
    Dim oBookMenores As Object
    Dim oSheetMenores As Object
    
    Dim oBookPersJurid As Object
    Dim oSheetPersJurid As Object
    
    Dim sClientes As String
    Dim sMontos As String
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
        
    Dim oPer As New COMDPersona.DCOMPersonas
    Dim rsPers As New ADODB.Recordset
    Dim rsPersTmp As New ADODB.Recordset
        
    If txtArchivo.Text = "" Then
        MsgBox "No selecciono ningun archivo", vbExclamation, "Aviso"
        Set oPer = Nothing
        Set rsPers = Nothing
        Set rsPersTmp = Nothing
        Exit Sub
    End If
    If MsgBox("¿Esta operación puede tardar minutos, esta seguro de continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    If grdCuenta.Rows >= 2 And Len(Trim(grdCuenta.TextMatrix(1, 1))) > 0 Then
        If MsgBox("Al cargar la trama, se limpiaran los registros del Grid, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        Else
            grdCuenta.Clear
            grdCuenta.Rows = 2
            grdCuenta.FormaCabecera
        End If
    End If
    grdCuenta.lbEditarFlex = False
    grdCuenta.ColumnasAEditar = "X-X-X-X-X-X"
    pbProgres.Max = 10
    pbProgres.Min = 1
    pbProgres.value = 1
    pbProgres.Visible = True
    DoEvents
    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Open(txtArchivo.Text)
    psArchivoAGrabar = App.Path & "\SPOOLER\NoCumpleValidacion_" & Format(gdFecSis, "yyyymmdd") & ".xls"
    
    grdCuenta.SetFocus
                
    cmdEliminar.Enabled = True
    X = 1
    Y = 1: Z = 1
    
    If Dir(psArchivoAGrabar) <> "" Then
        Kill psArchivoAGrabar
    End If
    pbProgres.value = 2
    DoEvents
       Set oExcel = CreateObject("Excel.Application")
       Set oBook = oExcel.Workbooks.Add
       Set oSheet = oBook.Worksheets(1)
    pbProgres.value = 3
    DoEvents
        bFormato = True
        bCargaLote = True
        
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 1))) <> "ITEM" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 2))) <> "TIPO DOC (1=DNI 2=RUC)" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 3))) <> "NRO DOC" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 4))) <> "MONTO" Then bFormato = False
        If bFormato = False Then
            MsgBox "El archivo seleccionado no tiene el formato adecuado para la carga en lote, verifíquelo e inténtelo de nuevo", vbInformation, "Aviso"
            If Not objExcel Is Nothing Then
                objExcel.Workbooks.Close
                Set objExcel = Nothing
            End If
            If Not oExcel Is Nothing Then
                oExcel.Workbooks.Close
                Set oExcel = Nothing
            End If
            pbProgres.Visible = False
            Exit Sub
        End If
        fila = 2
        pbProgres.value = 4
        DoEvents
        With xLibro
            With .Sheets(1)
            Do While Len(Trim(.Cells(fila, 2))) > 0
                lsNroDoc = Trim(.Cells(fila, 3))
                lsMonto = Trim(.Cells(fila, 4))
                sClientes = sClientes & Trim(lsNroDoc) & ","
                sMontos = sMontos & Trim(lsMonto) & ","
                fila = fila + 1
            Loop
            End With
        End With
        pbProgres.value = 7
        DoEvents
        If Len(sClientes) > 2 Then
            sClientes = Mid(sClientes, 1, Len(sClientes) - 1)
        End If
        If Len(sMontos) > 2 Then
            sMontos = Mid(sMontos, 1, Len(sMontos) - 1)
        End If
        If Len(sClientes) = 0 Or Len(sMontos) = 0 Then
            MsgBox "La trama seleccionada no contiene datos para la carga", vbInformation, "Aviso"
            If Not objExcel Is Nothing Then
                objExcel.Workbooks.Close
                Set objExcel = Nothing
            End If
            If Not oExcel Is Nothing Then
                oExcel.Workbooks.Close
                Set oExcel = Nothing
            End If
            pbProgres.Visible = False
            Exit Sub
        End If
        Set rsPers = oPer.ValidaTramaDeposito(sClientes, txtInstitucion.Text, Trim(Right(cboMoneda.Text, 5)))
        pbProgres.value = 8
        DoEvents
        If (MostrarErrores(rsPers)) Then
            grdCuenta.Clear
            grdCuenta.Rows = 2
            grdCuenta.FormaCabecera
        Else
            Set rsPers = Nothing
            Set rsPers = oPer.ListarTramaDeposito(sClientes, sMontos, txtInstitucion.Text, Trim(Right(cboMoneda.Text, 5)))
            If Not rsPers Is Nothing Then
                grdCuenta.rsFlex = rsPers
            End If
        End If
        pbProgres.value = 9
        DoEvents
        txtMonto.Text = Format(grdCuenta.SumaRow(5), "#,##0.00")
        pbProgres.value = 10
        DoEvents
        pbProgres.Visible = False
        objExcel.Quit
        Set objExcel = Nothing
        Set xLibro = Nothing
        Set oBook = Nothing
        oExcel.Quit
        Set oExcel = Nothing
End Sub

Private Sub CmdEliminar_Click()
    Dim nFila As Long
    nFila = grdCuenta.row
    If bCargaLote Then Exit Sub
    If MsgBox("¿Desea eliminar la fila seleccionada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdCuenta.EliminaFila nFila
    End If
End Sub

Private Sub cmdFormato_Click()
    
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim nFila, i As Double
        
On Error GoTo error_handler
        
    dlgArchivo.FileName = Empty
    dlgArchivo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx| Archivos de Excel (*.xls)|*.xls"
    dlgArchivo.FileName = "DepositoHaberesLote" & Format(Now, "yyyyMMddhhnnss") & ".xlsx"
    dlgArchivo.ShowSave
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    If fs.FileExists(dlgArchivo.FileName) Then
        MsgBox "El archivo '" & dlgArchivo.FileTitle & "' ya existe, debe asignarle un nombre diferente", vbExclamation, ""
        Exit Sub
    End If
    If fnProducto = gCapAhorros Then
        lsArchivo = App.Path & "\FormatoCarta\FormatoDepositoLoteAhorro.xlsx"
        lsNomHoja = "DepositoLoteAhorro"
    End If
    lsArchivo1 = dlgArchivo.FileName
    If fs.FileExists(lsArchivo) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(lsArchivo)
    Else
        MsgBox "No Existe Plantilla en la Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
    For Each xlHoja1 In xlsLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
        End If
    Next
    xlHoja1.SaveAs lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    MsgBox "Se exportó el formato para la carga de archivos", vbInformation, "Aviso"
               
Exit Sub
    
error_handler:
    If err.Number = 32755 Then
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Error al momento de generar el archivo", vbCritical, "Aviso"
    End If
End Sub

Private Function ValidacionDeposito() As String

    Dim oPer As New COMDPersona.DCOMPersonas
    Dim sClientes As String, sMensaje As String
    Dim rsPers As ADODB.Recordset
    Dim i As Integer
    
    ' registros en la grilla
    If grdCuenta.Rows = 2 And Len(Trim(grdCuenta.TextMatrix(1, 1))) = 0 Then
        sMensaje = "Debe al menos ingresar un registro en la grilla" & vbNewLine
    End If
    
    ' seleccion de institucion
    If Len(Trim(txtInstitucion.Text)) = 0 Then
        sMensaje = sMensaje & "Debe seleccionar una institucion" & vbNewLine
    End If
    
    ' texto en la glosa
    If Len(Trim(Replace(txtGlosa.Text, vbNewLine, ""))) = 0 Then
        sMensaje = sMensaje & "Debe ingresar un valor en la glosa" & vbNewLine
    End If
    
    ' registros del voucher
    If fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        If nNroDeposito <> grdCuenta.Rows - 1 Then
            sMensaje = sMensaje & "El número de registros ingresados no coincide con el numero de depósitos del voucher" & vbNewLine
        End If
    End If
    ' Monto Depositado
    If CCur(lblMonTra.Caption) <> CCur(txtMonto.Text) And fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        sMensaje = sMensaje & "El monto de depósito es diferente al monto del voucher" & vbNewLine
    End If
    
    ' registros del cheque
    If fnOpeCod = gAhoDepositoHaberesEnLoteChq And oDocRec.fnNroCliLote <> (grdCuenta.Rows - 1) Then
        sMensaje = sMensaje & "El número de registros ingresados no coincide con el numero de depósitos del cheque" & vbNewLine
    End If
    ' Monto Depositado
    If fnOpeCod = gAhoDepositoHaberesEnLoteChq And CCur(lblChequeMonto.Caption) <> CCur(txtMonto.Text) Then
        sMensaje = sMensaje & "El monto de depósito es diferente al monto del cheque" & vbNewLine
    End If
    
    ' Validando las cuentas en funcion a los DOI de los clientes
    For i = 1 To grdCuenta.Rows - 1
        sClientes = sClientes & Trim(grdCuenta.TextMatrix(i, 4)) & ","
    Next
    If Len(sClientes) > 2 Then
        sClientes = Mid(sClientes, 1, Len(sClientes) - 1)
    End If
    Set rsPers = oPer.ValidaTramaDeposito(sClientes, txtInstitucion.Text, Trim(Right(cboMoneda.Text, 5)))
    If Not rsPers Is Nothing Then
        If Not rsPers.EOF And Not rsPers.BOF Then
            sMensaje = sMensaje & "Revisar la titularidad del cliente, la moneda seleccionada y su vinculación con la institución" & vbNewLine
        End If
    End If
    ValidacionDeposito = sMensaje
    Exit Function
End Function


Private Sub cmdGrabar_Click()
'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
'fin Comprobacion si es RFIII

If gsOpeCod = gAhoDepositoHaberesEnLoteChq And Not ValidaSeleccionCheque() Then
    MsgBox "Ud. debe seleccionar el cheque a utilizar.", vbExclamation, "Aviso"
    EnfocaControl cmdCheque
    Exit Sub
End If

'***add by marg ers065-2017***
If Me.txtCuentaInstitucion.Text <> "" And bRetiroExitoso = False Then
    Call GrabarRetiro
End If
If (Me.txtCuentaInstitucion.Text <> "" And bRetiroExitoso = True And bDepositoExitoso = False) Or (Me.txtCuentaInstitucion.Text = "") Then
    Call GrabarDeposito
End If
'***end marg*******************


  '************ Registrar actividad de opertaciones especiales - ANDE 2017-12-18
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim RVerOpe As ADODB.Recordset
    Dim nEstadoActividad As Integer
    nEstadoActividad = oCaptaLN.RegistrarActividad(fnOpeCod, gsCodUser, gdFecSis)
    
    If nEstadoActividad = 1 Then
        MsgBox "He detectado un problema; su operación no fue afectada, pero por favor comunciar a TI-Desarrollo.", vbError, "Error"
    ElseIf nEstadoActividad = 2 Then
        MsgBox "Ha usado el total de operaciones permitidas para el día de hoy. Si desea realizar más operaciones, comuníquese con el área de Operaciones.", vbInformation + vbOKOnly, "Aviso"
        Unload Me
    End If
    ' END ANDE ******************************************************************

End Sub

'***ADD BY MARG ERS065-2017**************
Private Sub ReadaptarFormulario()
'***sin retiro***
If Me.txtCuentaInstitucion.Text = "" Then
    fraTipo.Left = 135
    fraTipo.Top = 405
    fraTipo.Width = 9570
    fraTipo.Height = 1080
    
    lblInst.Left = 5760
    lblInst.Top = 240
    lblInst.Width = 3690
    lblInst.Height = 315
    
    fraCuenta.Left = 135
    fraCuenta.Top = 1530
    fraCuenta.Width = 9570
    fraCuenta.Height = 3840
    
    fraTranferecia.Left = 135
    fraTranferecia.Top = 5400
    fraTranferecia.Width = 4410
    fraTranferecia.Height = 2115
    
    fraCheque.Left = 135
    fraCheque.Top = 5400
    fraCheque.Width = 4410
    fraCheque.Height = 2115
    
    fraGlosa.Left = 4590
    fraGlosa.Top = 5400
    fraGlosa.Width = 2430
    fraGlosa.Height = 2115
    
    fraMonto.Left = 7065
    fraMonto.Top = 5400
    fraMonto.Width = 2655
    fraMonto.Height = 2115
    
    cmdCancelar.Left = 45
    cmdCancelar.Top = 7560
    cmdCancelar.Width = 1095
    cmdCancelar.Height = 315
    
    pbProgres.Left = 2295
    pbProgres.Top = 7650
    pbProgres.Width = 5055
    pbProgres.Height = 195
    
    
    cmdGrabar.Left = 7650
    cmdGrabar.Top = 7560
    cmdGrabar.Width = 960
    cmdGrabar.Height = 315
    
    
    cmdSalir.Left = 8730
    cmdSalir.Top = 7560
    cmdSalir.Width = 960
    cmdSalir.Height = 315
    
    SSTab1.Left = 45
    SSTab1.Top = 45
    SSTab1.Width = 9825
    SSTab1.Height = 7980
    
    Me.Width = 9990
    Me.Height = 8475

    Label3.Visible = False
    txtCuentaInstitucion.Visible = False
    fraCliente.Visible = False
End If

'***con retiro***
If Me.txtCuentaInstitucion.Text <> "" Then

    Dim oCuenta As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim orsCuenta As ADODB.Recordset
    Dim lnPersoneria As COMDConstantes.PersPersoneria
    Dim lnTipoCuenta As COMDConstantes.ProductoCuentaTipo
    Set oCuenta = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set orsCuenta = New ADODB.Recordset
    Set orsCuenta = oCuenta.GetDatosCuenta(Me.txtCuentaInstitucion.Text)
    If Not (orsCuenta.EOF And orsCuenta.BOF) Then
        lnPersoneria = orsCuenta("nPersoneria")
        lnTipoCuenta = orsCuenta("nPrdCtaTpo")
    End If
    
    If lnPersoneria <> gPersonaNat Or lnTipoCuenta = gPrdCtaTpoMancom Or lnTipoCuenta = gPrdCtaTpoIndist Then 'persona juridica o cuenta mancomunado o indistintinta
        fraTipo.Left = 135
        fraTipo.Top = 405
        fraTipo.Width = 9570
        fraTipo.Height = 1080
        
        lblInst.Left = 5760
        lblInst.Top = 240
        lblInst.Width = 3690
        lblInst.Height = 315
        
        fraCuenta.Left = 135
        fraCuenta.Top = 4440
        fraCuenta.Width = 9570
        fraCuenta.Height = 3480
        
        fraTranferecia.Left = 135
        fraTranferecia.Top = 5400
        fraTranferecia.Width = 4410
        fraTranferecia.Height = 2115
        
        fraCheque.Left = 135
        fraCheque.Top = 5400
        fraCheque.Width = 4410
        fraCheque.Height = 2115
        
        fraGlosa.Left = 4590
        fraGlosa.Top = 7920
        fraGlosa.Width = 2430
        fraGlosa.Height = 2115
        
        fraMonto.Left = 7080
        fraMonto.Top = 7920
        fraMonto.Width = 2655
        fraMonto.Height = 2115
        
        cmdCancelar.Left = 135
        cmdCancelar.Top = 10080
        cmdCancelar.Width = 1095
        cmdCancelar.Height = 315
        
        pbProgres.Left = 2295
        pbProgres.Top = 7650
        pbProgres.Width = 5055
        pbProgres.Height = 195
        
        cmdGrabar.Left = 7560
        cmdGrabar.Top = 10080
        cmdGrabar.Width = 960
        cmdGrabar.Height = 315
        
        cmdSalir.Left = 8730
        cmdSalir.Top = 10080
        cmdSalir.Width = 960
        cmdSalir.Height = 315
        
        SSTab1.Left = 45
        SSTab1.Top = 45
        SSTab1.Width = 9825
        SSTab1.Height = 10500
        
        Me.Width = 10035
        Me.Height = 11010
        
        Label3.Visible = True
        txtCuentaInstitucion.Visible = True
        fraCliente.Visible = True
    Else
        fraTipo.Left = 135
        fraTipo.Top = 405
        fraTipo.Width = 9570
        fraTipo.Height = 1080
        
        lblInst.Left = 5760
        lblInst.Top = 240
        lblInst.Width = 3690
        lblInst.Height = 315
        
        fraCuenta.Left = 135
        fraCuenta.Top = 1530
        fraCuenta.Width = 9570
        fraCuenta.Height = 3840
        
        fraTranferecia.Left = 135
        fraTranferecia.Top = 5400
        fraTranferecia.Width = 4410
        fraTranferecia.Height = 2115
        
        fraCheque.Left = 135
        fraCheque.Top = 5400
        fraCheque.Width = 4410
        fraCheque.Height = 2115
        
        fraGlosa.Left = 4590
        fraGlosa.Top = 5400
        fraGlosa.Width = 2430
        fraGlosa.Height = 2115
        
        fraMonto.Left = 7065
        fraMonto.Top = 5400
        fraMonto.Width = 2655
        fraMonto.Height = 2115
        
        cmdCancelar.Left = 45
        cmdCancelar.Top = 7560
        cmdCancelar.Width = 1095
        cmdCancelar.Height = 315
        
        pbProgres.Left = 2295
        pbProgres.Top = 7650
        pbProgres.Width = 5055
        pbProgres.Height = 195
        
        
        cmdGrabar.Left = 7650
        cmdGrabar.Top = 7560
        cmdGrabar.Width = 960
        cmdGrabar.Height = 315
        
        
        cmdSalir.Left = 8730
        cmdSalir.Top = 7560
        cmdSalir.Width = 960
        cmdSalir.Height = 315
        
        SSTab1.Left = 45
        SSTab1.Top = 45
        SSTab1.Width = 9825
        SSTab1.Height = 7980
        
        Me.Width = 9990
        Me.Height = 8475
    
        Label3.Visible = True
        txtCuentaInstitucion.Visible = True
        fraCliente.Visible = False
    End If
End If

End Sub

Private Sub GrabarDeposito()
Dim nMontoCargo As Double
Dim sCuenta As String, sGlosa As String
Dim lsmensaje As String
Dim lsBoleta As String
Dim lsBoletaITF As String
Dim nFicSal As Integer
Dim Autid As Long
Dim bResult As Boolean
Dim sCuentas() As String
Dim pvDocRec As Variant

On Error GoTo ErrGraba

lsmensaje = Trim(ValidacionDeposito)
If Len(lsmensaje) > 0 Then
    MsgBox "Se presentaron observaciones: " & vbNewLine & lsmensaje, vbInformation, "Aviso"
    Exit Sub
End If

If fnOpeCod = gAhoDepositoHaberesEnLoteChq Then
    ReDim pvDocRec(5)
    pvDocRec(0) = oDocRec.fsIFTpo
    pvDocRec(1) = oDocRec.fsIFCta
    pvDocRec(2) = oDocRec.fsPersCod
    pvDocRec(3) = oDocRec.fnTpoDoc
    pvDocRec(4) = oDocRec.fsNroDoc
End If

If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then

    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMovNro As String, sMovNroV As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim rsCtaAbo As ADODB.Recordset

    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Sleep (1000)
    sMovNroV = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set rsCtaAbo = grdCuenta.GetRsNew()
    sGlosa = Replace(Trim(txtGlosa.Text), vbNewLine, " ")

    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios, sPersLavDinero As String
    Dim nMontoLavDinero As Double, nTC As Double, sReaPersLavDinero As String, sBenPersLavDinero As String

    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
    Set clsExo = Nothing
    Set clsLav = Nothing

    nMontoCargo = CDbl(txtMonto.Text)

    If fnOpeCod = gAhoDepositoHaberesEnLoteEfec Then
        bResult = clsCap.CapAbonoLoteCtaSueldo(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(gnTipCambioC), CDbl(gnTipCambioV), gbITFAplica, 0, gbITFAsumidoAho, 0, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , , , gnMovNro, fnOpeCod)
    ElseIf fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        bResult = clsCap.CapAbonoLoteCtaSueldo(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(gnTipCambioC), CDbl(gnTipCambioV), gbITFAplica, 0, gbITFAsumidoAho, 0, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , , , gnMovNro, fnOpeCod, , , , , , lnMovNroTransfer, fnMovNroRVD, sMovNroV, nNroDeposito)
    ElseIf fnOpeCod = gAhoDepositoHaberesEnLoteChq Then
        bResult = clsCap.CapAbonoLoteCtaSueldo(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, sPersLavDinero, CDbl(gnTipCambioC), CDbl(gnTipCambioV), gbITFAplica, 0, gbITFAsumidoAho, 0, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , , , gnMovNro, fnOpeCod, , , , , , lnMovNroTransfer, fnMovNroRVD, sMovNroV, nNroDeposito, pvDocRec)
    End If
    
    If bResult Then
     If gnMovNro > 0 Then
         Call frmMovLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sBenPersLavDinero, , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen) 'JACA 20110224
     End If

      If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
      End If

      Do
        If Trim(lsBoleta) <> "" Then
           nFicSal = FreeFile
           Open sLpt For Output As nFicSal
              Print #nFicSal, lsBoleta & Chr$(12)
              Print #nFicSal, ""
           Close #nFicSal
        End If
      Loop Until MsgBox("¿Desea reimprimir Boleta de Depósito en lote? ", vbQuestion + vbYesNo, Me.Caption) = vbNo

      bDepositoExitoso = True 'ADD BY MARG ERS065-2017
      cmdCancelar_Click
      MsgBox "La operación se realizó correctamente", vbInformation, "Aviso"

      'INICIO JHCU ENCUESTA 16-10-2019
      sOperacion = fnOpeCod
      Encuestas gsCodUser, gsCodAge, "ERS0292019", sOperacion
      'FIN
    Else
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
End If
 Set clsCap = Nothing
 gVarPublicas.LimpiaVarLavDinero

Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub
End Sub
'***END MARG*******************************


Private Sub cmdTranfer_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oForm As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lsDetalle As String
    Dim lnTransferSaldo As Currency
    Dim fsPersCodTransfer As String
    If cboTransferMoneda.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMoneda.Visible And cboTransferMoneda.Enabled Then cboTransferMoneda.SetFocus
        Exit Sub
    End If
    If gsOpeCod = gAhoDepositoHaberesEnLoteTransf Then
        lnTipMot = 9
    End If
    fnMovNroRVD = 0
    Set oForm = New frmCapRegVouDepBus
    SetDatosTransferencia "", "", "", 0, -1, "" 'Limpiamos datos y variables globales
    oForm.iniciarFormulario Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    If IsNumeric(Trim(lsDetalle)) Then
        nNroDeposito = CInt(lsDetalle)
    Else
        nNroDeposito = 0
    End If
    SetDatosTransferencia lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
    Me.grdCuenta.row = 1
    Set oForm = Nothing
    Exit Sub
End Sub
Private Sub SetDatosTransferencia(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosa.Text = psGlosa
    lbltransferBco.Caption = psInstit
    lblTrasferND.Caption = psDoc
    If psDetalle <> "" Then

    End If
    
    If pnMovNroTransfer <> -1 Then
        txtTransferGlosa.SetFocus
    End If
    
    txtTransferGlosa.Locked = True
    txtMonto.Enabled = False
    lblMonTra = Format(pnTransferSaldo, "#,##0.00")
    
    Set rsPersona = Nothing
    Set oPersona = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub

Private Sub Form_Load()
Me.Caption = fsDescOperacion
Me.txtMonto.Enabled = False
bCargaLote = False
lblTTCCD.Caption = Format(gnTipCambioC, "#,#0.0000")
lblTTCVD.Caption = Format(gnTipCambioV, "#,#0.0000")
End Sub

Private Sub grdCuenta_OnCellChange(pnRow As Long, pnCol As Long)
    txtMonto.Text = Format(grdCuenta.SumaRow(5), "#,##0.00")
End Sub

Private Sub grdCuenta_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)

Dim sCta As String
Dim sRelac As String
Dim sEstado As String
Dim i As Long
Dim bDuplicadoCuenta As Boolean

If psDataCod = "" Then
    grdCuenta.EliminaFila pnRow
    Exit Sub
End If

Dim ClsPersona As New COMDPersona.DCOMPersonas
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsPers As New COMDPersona.UCOMPersona
Dim rsCuenta As New ADODB.Recordset
Dim rsPersona As New ADODB.Recordset
Dim clsCuenta As New UCapCuenta

Set rsCuenta = clsCap.GetCuentasPersona(psDataCod, gCapAhorros, True, , Trim(Right(cboMoneda.Text, 5)), , , "6", True)
Set rsPersona = ClsPersona.BuscaCliente(psDataCod, BusquedaCodigo)

If Not rsCuenta Is Nothing Then
    If Not (rsCuenta.EOF And rsCuenta.EOF) Then
        Do While Not rsCuenta.EOF
            sCta = rsCuenta("cCtaCod")
            sRelac = rsCuenta("cRelacion")
            sEstado = Trim(rsCuenta("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & space(2) & sRelac & space(2) & sEstado
            rsCuenta.MoveNext
        Loop
    Else
        grdCuenta.EliminaFila pnRow
        MsgBox "Persona no posee cuenta sueldo", vbInformation, "Aviso"
        rsCuenta.Close
        Set rsCuenta = Nothing
        Set clsPers = Nothing
        Set ClsPersona = Nothing
        Set clsCap = Nothing
        Exit Sub
    End If
End If
rsCuenta.Close
grdCuenta.TextMatrix(pnRow, 2) = ""
Set rsCuenta = Nothing
Set clsCuenta = New UCapCuenta
Set clsCuenta = frmCapMantenimientoCtas.Inicia
If Not clsCuenta Is Nothing Then
    If clsCuenta.sCtaCod <> "" Then
        If pbEsDuplicado Then
            For i = 1 To grdCuenta.Rows - 1
                If clsCuenta.sCtaCod = grdCuenta.TextMatrix(i, 2) Then
                    MsgBox "El registro seleccionado es duplicado.", vbInformation, "Aviso"
                    grdCuenta.EliminaFila pnRow
                    Exit Sub
                End If
            Next
        End If
        grdCuenta.TextMatrix(grdCuenta.row, 1) = rsPersona!cperscod
        grdCuenta.TextMatrix(grdCuenta.row, 2) = clsCuenta.sCtaCod
        grdCuenta.TextMatrix(grdCuenta.row, 3) = rsPersona!cPersNombre
        grdCuenta.TextMatrix(grdCuenta.row, 4) = rsPersona!cPersIDnroDNI
        grdCuenta.Col = 4
        SendKeys "{F2}"
    Else
        grdCuenta.EliminaFila pnRow
    End If
Else
    grdCuenta.EliminaFila pnRow
End If
Set clsCuenta = Nothing
            
End Sub

Private Sub grdCuenta_OnRowDelete()
    txtMonto.Text = Format$(grdCuenta.SumaRow(5), "#,##0.00")
End Sub

Private Sub grdCuenta_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdCuenta.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub

Private Sub grdCuenta_RowColChange()
    Dim nRow As Long
    Dim nCol As Long
    
    nRow = grdCuenta.row
    nCol = grdCuenta.Col
    If bCargaLote Then
        If nCol = 1 Then
            grdCuenta.lbEditarFlex = False
            Me.KeyPreview = True
        Else
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = True
        End If
    Else
        If nCol = 1 Then
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = False
        Else
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = True
        End If
    End If
    If Not IsNumeric(Trim(grdCuenta.TextMatrix(nRow, 1))) Then
        grdCuenta.TextMatrix(nRow, 1) = ""
    End If
End Sub


Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        If cmdGrabar.Enabled Then cmdGrabar.SetFocus
    End If
End Sub
Private Sub txtInstitucion_EmiteDatos()
    If txtInstitucion.Text <> "" Then
        '***ADD BY MARG ERS065-2017**************************
        If fnOpeCod = gAhoDepositoHaberesEnLoteEfec Then
            Call frmCapCuentasInstitucionVer.Inicio(txtInstitucion.Text)
            If txtCuentaInstitucion.Text <> "" Then
                Call ReadaptarFormulario
                IniciaRetiro gCapAhorros, 200301, "RETIRO EFECTIVO"
                ObtieneDatosCuenta txtCuentaInstitucion.Text
            End If
            lblInst.Text = txtInstitucion.psDescripcion
            If cmdAgregar.Enabled Then cmdAgregar.SetFocus
        End If
        'END MARG**********************************************
         If fnOpeCod <> gAhoDepositoHaberesEnLoteEfec Then 'add by marg ers065-2017
            If txtInstitucion.PersPersoneria < 2 Then
                lblInst.Text = ""
                txtInstitucion.Text = ""
                MsgBox "La institución seleccionada debe tener personería Jurídica", vbInformation, "Aviso"
            Else
                lblInst.Text = txtInstitucion.psDescripcion
                If cmdAgregar.Enabled Then cmdAgregar.SetFocus
            End If
        End If 'add by marg ers065-2017
    Else
        lblInst.Text = ""
    End If
End Sub
Public Sub iniciarFormulario(ByVal pnProducto As Producto, ByVal pnOpeCod As CaptacOperacion, _
                             Optional psDescOperacion As String = "")
fnProducto = pnProducto
fnOpeCod = pnOpeCod
fsDescOperacion = psDescOperacion

'/*Verificar cantidad de operaciones disponibles ANDE 20171218*/
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim bProsigue As Boolean
    Dim cMsgValid As String
    bProsigue = oCaptaLN.OperacionPermitida(gsCodUser, gdFecSis, fnOpeCod, cMsgValid)
    If bProsigue = False Then
        MsgBox cMsgValid, vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
'/*end ande*/

Dim clsGen As New COMDConstSistema.DCOMGeneral
Dim rsConst As New ADODB.Recordset
Set rsConst = clsGen.GetConstante(gMoneda)

Select Case fnProducto

    Case gCapAhorros
    
        If fnOpeCod = gAhoDepositoHaberesEnLoteEfec Then
            fraTranferecia.Visible = False
        ElseIf fnOpeCod = gAhoDepositoHaberesEnLoteTransf Then
            fraTranferecia.Visible = True
        ElseIf fnOpeCod = gAhoDepositoHaberesEnLoteChq Then
            fraCheque.Visible = True
        End If

    Case gCapCTS
            
End Select

If gsOpeCod = gAhoDepositoHaberesEnLoteTransf Then
    Set rsConst = clsGen.GetConstante(gMoneda)
    CargaCombo cboTransferMoneda, rsConst
End If

Set rsConst = clsGen.GetConstante(gMoneda)
CargaCombo cboMoneda, rsConst
If cboMoneda.ListCount > 0 Then
    cboMoneda.ListIndex = 0
End If

'***add by marg ers065-2017***
    Call ReadaptarFormulario
'***end marg*******************

Me.Show 1
End Sub

Private Function MostrarErrores(ByVal rsErrores As ADODB.Recordset) As Boolean

    Dim oBook As Object
    Dim oSheet As Object
    Dim sDireccion As String
    Dim oExcel As Object
    Dim bResult As Boolean
    Dim i As Long
    
    On Error GoTo error_handler
    
    bResult = False
    If Not rsErrores Is Nothing Then
        Set oExcel = CreateObject("Excel.Application")
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Worksheets(1)
        sDireccion = App.Path & "\SPOOLER\Observaciones_" & Format(CDate(gdFecSis), "yyyyMMdd") & ".xls"
        If Dir(sDireccion) <> "" Then
            Kill sDireccion
        End If
        oSheet.Range("A1:F1").Font.Bold = True
        oSheet.Columns("B:B").NumberFormat = "@"
        oSheet.Columns("C:C").NumberFormat = "@"
        oSheet.Range("A1").value = "#"
        oSheet.Columns("A:A").ColumnWidth = 7
        oSheet.Columns("B:B").ColumnWidth = 21
        oSheet.Columns("C:C").ColumnWidth = 15
        oSheet.Columns("D:D").ColumnWidth = 51
        oSheet.Columns("F:F").ColumnWidth = 80
        oSheet.Range("B1").value = "NRO CUENTA"
        oSheet.Range("C1").value = "COD CLIENTE"
        oSheet.Range("D1").value = "NOMBRE"
        oSheet.Range("E1").value = "DOI"
        oSheet.Range("F1").value = "OBSERVACIONES"
        i = 2
        Do While Not rsErrores.EOF And Not rsErrores.BOF
            oSheet.Range("A" & i).value = i - 1
            oSheet.Range("B" & i).value = rsErrores!cCtaCod
            oSheet.Range("C" & i).value = rsErrores!cperscod
            oSheet.Range("D" & i).value = rsErrores!cPersNombre
            oSheet.Range("E" & i).value = rsErrores!cPersDoi
            If rsErrores!nRegistrado = 0 Then
                oSheet.Range("F" & i).value = "Persona no registrada"
            ElseIf rsErrores!nTitular = 0 Then
                oSheet.Range("F" & i).value = "Cliente no es titular de la cuenta sueldo"
            ElseIf rsErrores!nMonedaSelec = 0 Then
                oSheet.Range("F" & i).value = "La cuenta sueldo no es de la moneda seleccionada"
            ElseIf rsErrores!nVinculacion = 0 Then
                oSheet.Range("F" & i).value = "La cuenta sueldo no está vinculada a la empresa seleccionada"
            ElseIf rsErrores!nCantCuentas > 1 Then
                oSheet.Range("F" & i).value = "Cliente tiene mas de una cuenta sueldo con la institucion y moneda seleccionados"
            Else
                oSheet.Range("F" & i).value = "La persona presenta observaciones"
            End If
            i = i + 1
            rsErrores.MoveNext
            bResult = True
        Loop
        oBook.SaveAs sDireccion
        oExcel.Quit
        Set oExcel = Nothing
        Set oBook = Nothing
        If i > 2 Then
            If MsgBox("¿Algunos registros de la trama no cumplen con las validaciones respectivas, deseas exportar el detalle a Excel?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                MostrarErrores = bResult
                Exit Function
            End If
            Dim m_Excel As New Excel.Application
            m_Excel.Workbooks.Open (sDireccion)
            m_Excel.Visible = True
        End If
    End If
    
    MostrarErrores = bResult
    Exit Function
    
error_handler:
        oExcel.Quit
        Set oExcel = Nothing
        Set oBook = Nothing
        MsgBox "Error al momento de generar el archivo", vbCritical, "Aviso"
        MostrarErrores = False
        
End Function

Private Sub LimpiarGrdCuenta()
    grdCuenta.Clear
    grdCuenta.Rows = 2
    grdCuenta.FormaCabecera
End Sub

Private Sub Limpiar()
LimpiaControlesRetiro
LimpiarGrdCuenta
LimpiarTransferencia
LimpiarCheque
LimpiarControles
'***add bymarg ers065-217***
ResetearVariables
ReadaptarFormulario
'***end marg*****************
End Sub
'***add by marg ers065-2017***
Private Sub ResetearVariables()
    bRetiroExitoso = False
    bDepositoExitoso = False
End Sub
Private Sub SetearControles()
Me.cboMedioRetiro.ListIndex = 0
End Sub
'***end marg******************

Private Sub LimpiarTransferencia()
    SetDatosTransferencia "", "", "", 0, -1, "" 'Limpiamos datos y variables globales
End Sub

Private Sub LimpiarCheque()
    Set oDocRec = New UDocRec
    Call setDatosCheque
End Sub

Private Sub LimpiarControles()
    If cboMoneda.ListCount > 0 Then
        cboMoneda.ListIndex = 0
    End If
    txtArchivo.Text = ""
    txtMonto.Text = "0.00"
    txtInstitucion.Text = ""
    Me.txtCuentaInstitucion.Text = "" 'add by marg ers065-2017
    lblInst.Text = ""
    txtGlosa.Text = ""
    grdCuenta.ColumnasAEditar = "X-1-X-X-X-5"
    bCargaLote = False
    cboMoneda.SetFocus
End Sub

'***MARG ERS065-2017 metodos reutilizados de frmCapCargos***********************************************************************************************

Public Sub IniciaRetiro(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, _
        ByVal sDescOperacion As String, Optional sCodCmac As String = "", _
        Optional sNomCmac As String, Optional lcCtaCod As String, Optional pnMonto As Double, Optional lnMovNro As Long)

nProducto = nProd
nOperacion = nOpe
sPersCodCMAC = sCodCmac
sNombreCMAC = sNomCmac
sOperacion = sDescOperacion

Select Case nProd
    Case gCapAhorros
        'lblEtqUltCnt = "Ult. Contacto :" 'COMMENT BY MARG ERS065-2017
        'lblUltContacto.Width = 2000 'COMMENT BY MARG ERS065-2017
        'txtCuenta.Prod = Trim(Str(gCapAhorros)) 'COMMENT BY MARG ERS065-2017
        If sPersCodCMAC = "" Then
            'Me.Caption = "Captaciones - Cargo - Ahorros " & sDescOperacion 'COMMENT BY MARG ERS065-2017
        Else
            'Me.Caption = "Captaciones - Cargo - Ahorros " & sDescOperacion & " - " & sNombreCMAC 'COMMENT BY MARG ERS065-2017
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
    
        'lblEtqUltCnt = "Plazo :" 'COMMENT BY MARG ERS065-2017
        'lblUltContacto.Width = 1000 'COMMENT BY MARG ERS065-2017
        'txtCuenta.Prod = Trim(Str(gCapPlazoFijo)) 'COMMENT BY MARG ERS065-2017
        'Me.Caption = "Captaciones - Cargo - Plazo Fijo " & sDescOperacion 'COMMENT BY MARG ERS065-2017
        Label3.Visible = True
        Label5.Visible = False ' RIRO SE CAMBIO A FALSE
        
        lblAlias.Visible = True
        lblMinFirmas.Visible = False ' RIRO SE CAMBIO A FALSE
        
        grdCliente.ColWidth(6) = 1200
               
    Case gCapCTS
        'lblEtqUltCnt = "Institución :" 'COMMENT BY MARG ERS065-2017
        'lblUltContacto.Width = 4250 'COMMENT BY MARG ERS065-2017
        'lblUltContacto.Left = lblUltContacto.Left - 275 'COMMENT BY MARG ERS065-2017
        'txtCuenta.Prod = Trim(Str(gCapCTS)) 'COMMENT BY MARG ERS065-2017
        If sPersCodCMAC = "" Then
            'Me.Caption = "Captaciones - Cargo - CTS " & sDescOperacion 'COMMENT BY MARG ERS065-2017
        Else
            'Me.Caption = "Captaciones - Cargo - CTS " & sDescOperacion & " - " & sNombreCMAC 'COMMENT BY MARG ERS065-2017
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
        'lblDocumento.Visible = False 'COMMENT BY MARG ERS065-2017
        'cboDocumento.Visible = False 'COMMENT BY MARG ERS065-2017
        'cmdDocumento.Visible = False 'COMMENT BY MARG ERS065-2017
        'lblOrdenPago.Visible = True 'COMMENT BY MARG ERS065-2017
        'txtOrdenPago.Visible = True 'COMMENT BY MARG ERS065-2017
        
        If nOperacion = gAhoRetOPCanje Then
            'txtBanco.Visible = True 'COMMENT BY MARG ERS065-2017
            'lblBanco.Visible = True 'COMMENT BY MARG ERS065-2017
            'lblEtqBanco.Visible = True 'COMMENT BY MARG ERS065-2017
            Dim clsBanco As COMNCajaGeneral.NCOMCajaCtaIF 'NCajaCtaIF
            Dim rsBanco As New ADODB.Recordset
            Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
                Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "_1%", MuestraInstituciones, "1")
            Set clsBanco = Nothing
            'txtBanco.rs = rsBanco 'COMMENT BY MARG ERS065-2017
             
        Else
            'txtBanco.Visible = False 'COMMENT BY MARG ERS065-2017
            'lblBanco.Visible = False 'COMMENT BY MARG ERS065-2017
            'lblEtqBanco.Visible = False 'COMMENT BY MARG ERS065-2017
            
        End If
    End If
    'fraDocumento.Caption = Trim(rsDoc("cDocDesc")) 'COMMENT BY MARG ERS065-2017
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
        
        'txtBanco.rs = rsBanco 'COMMENT BY MARG ERS065-2017
        
        'CargaCombo Me.cboMonedaBanco, oCon.RecuperaConstantes(gMoneda) 'COMMENT BY MARG ERS065-2017
        Set OCon = Nothing
        
        'txtBanco.Visible = True 'COMMENT BY MARG ERS065-2017
        'lblBanco.Visible = True 'COMMENT BY MARG ERS065-2017
        'lblEtqBanco.Visible = True 'COMMENT BY MARG ERS065-2017
        
        'lblDocumento.Visible = True 'COMMENT BY MARG ERS065-2017
        'cmdDocumento.Visible = True 'COMMENT BY MARG ERS065-2017
        
        'cboMonedaBanco.Visible = True 'COMMENT BY MARG ERS065-2017
        'txtCtaBanco.Visible = True 'COMMENT BY MARG ERS065-2017
        'lblCtaBanco.Visible = True 'COMMENT BY MARG ERS065-2017
    
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
        'txtBancoTrans.rs = rsBanco 'COMMENT BY MARG ERS065-2017
        
        'Cargando Plaza
        Set rsConstante = oConstante.GetConstante("10032", , "'20[^0]'")
        'CargaCombo cbPlazaTrans, rsConstante 'COMMENT BY MARG ERS065-2017
        'cbPlazaTrans.ListIndex = 0 'COMMENT BY MARG ERS065-2017
        Set rsConstante = Nothing
        
        'fraDocumento.Visible = False 'COMMENT BY MARG ERS065-2017
        'fraDocumentoTrans.Visible = True 'COMMENT BY MARG ERS065-2017
                
        'fraDocumentoTrans.Left = fraDocumento.Left 'COMMENT BY MARG ERS065-2017
        'fraDocumentoTrans.Top = fraDocumento.Top 'COMMENT BY MARG ERS065-2017
        'fraMonto.Height = fraDocumentoTrans.Height 'COMMENT BY MARG ERS065-2017
        
        cboMedioRetiro.Visible = True
        lblMedioRetiro.Visible = True
        lblComisionTransf.Visible = True
        chkTransfEfectivo.Visible = True
        lblComisionTransf.Top = 660
        chkTransfEfectivo.Top = 720
        
    Else
        'txtBanco.Visible = False 'COMMENT BY MARG ERS065-2017
        'lblBanco.Visible = False 'COMMENT BY MARG ERS065-2017
        'lblEtqBanco.Visible = False 'COMMENT BY MARG ERS065-2017
    
        'lblDocumento.Visible = False 'COMMENT BY MARG ERS065-2017
        'cmdDocumento.Visible = False 'COMMENT BY MARG ERS065-2017
        
    End If
    
    'lblDocumento.Visible = False 'COMMENT BY MARG ERS065-2017
    'cboDocumento.Visible = False 'COMMENT BY MARG ERS065-2017
    'cmdDocumento.Visible = False 'COMMENT BY MARG ERS065-2017

    bDocumento = False
    'lblOrdenPago.Visible = False 'COMMENT BY MARG ERS065-2017
    'txtOrdenPago.Visible = False 'COMMENT BY MARG ERS065-2017
    
End If
rsDoc.Close
Set rsDoc = Nothing
'txtCuenta.CMAC = gsCodCMAC 'COMMENT BY MARG ERS065-2017
'txtCuenta.EnabledProd = False 'COMMENT BY MARG ERS065-2017
'txtCuenta.EnabledCMAC = False 'COMMENT BY MARG ERS065-2017
'cmdGrabar.Enabled = False 'COMMENT BY MARG ERS065-2017
'cmdCancelar.Enabled = False 'COMMENT BY MARG ERS065-2017
fraCliente.Enabled = False
'fraDocumento.Enabled = False 'COMMENT BY MARG ERS065-2017
'fraMonto.Enabled = False 'COMMENT BY MARG ERS065-2017

sMovNroAut = ""

PbPreIngresado = False
If lcCtaCod <> "" Then
    'txtCuenta.NroCuenta = lcCtaCod 'COMMENT BY MARG ERS065-2017
    Me.txtCuentaInstitucion.Text = lcCtaCod 'ADD BY MARG ERS065-2017
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
                'txtCuenta.NroCuenta = sCuenta 'COMMENT BY MARG ERS065-2017
                Me.txtCuentaInstitucion.Text = sCuenta 'ADD BY MARG ERS065-2017
                lblComision.Visible = True
                
                'txtCuenta.SetFocusCuenta
                ObtieneDatosCuenta sCuenta
                
                'Me.Show 1 'COMMENT BY MARG ERS065-2017
            End If
        Else
            lblComision.Visible = True
            lblMonComision.Visible = True
            chkVBComision.Visible = True
            
            'Me.Show 1 'COMMENT BY MARG ERS065-2017
        End If
    Else
        'Me.Show 1 'COMMENT BY MARG ERS065-2017
    End If
    'End GITU
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
Dim loVistoElectronico As New frmVistoElectronico
Dim lbVistoVal As Boolean
'----- END GITU
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
                'txtCuenta.Cuenta = "" 'COMMENT BY MARG ERS065-2017
                Me.txtCuentaInstitucion.Text = "" 'ADD BY MARG ERS065-2017
                If gnCodOpeTarj <> 1 Then
                    'txtCuenta.SetFocus 'COMMENT BY MARG ERS065-2017
                    Me.txtCuentaInstitucion.SetFocus 'ADD BY MARG ERS065-2017
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
                   'Me.chkITFEfectivo.value = 0
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
        'lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm:ss") 'COMMENT BY MARG ERS065-2017
        nmoneda = CLng(Mid(sCuenta, 9, 1))
        ' MADM 20101115
        nEstadoT = rsCta("nPrdEstado")
        ' END MADM
        'JUEZ 20141017 ******************************************************
        lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set rsPar = clsDef.GetCapParametroNew(nProducto, lnTpoPrograma)
        If nProducto = gCapAhorros Then
            nParCantRetLib = rsPar!nCantOpeVentRet
        Else
            nParDiasVerifRegSueldo = rsPar!nDiasVerifUltRegSueldo
            nParUltRemunBrutas = rsPar!nUltRemunBrutas
        End If
        Set rsPar = Nothing
        'END JUEZ ***********************************************************
        
        If nmoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            txtMonto.BackColor = &HC0FFFF
            lblMon.Caption = "S/."
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
                    
                    If nCantOpeCta >= nParCantRetLib Then
                        If MsgBox("Se ha realizado el número máximo de retiros, se cargará una comisión. Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
                        'nComiMaxOpe = clsDef.GetCapParametro(2155)
                        nComiMaxOpe = clsDef.GetCapParametro(IIf(Mid(sCuenta, 9, 1) = gMonedaNacional, 2155, 2156)) 'JUEZ 20150105
                    End If
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
                'lblUltContacto = Format$(rsCta("dUltContacto"), "dd mmm yyyy hh:mm:ss") 'COMMENT BY MARG ERS065-2017
                Me.lblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.lblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
                
                If lbITFCtaExonerada Then
                    Dim nTipoExo As String, sDescripcion As String
                    nTipoExo = fgITFTipoExoneracion(sCuenta, sDescripcion)
                    'LblTituloExoneracion.Visible = True 'COMMENT BY MARG ERS065-2017
                    'lblExoneracion.Visible = True 'COMMENT BY MARG ERS065-2017
                    'lblExoneracion.Caption = sDescripcion 'COMMENT BY MARG ERS065-2017
                End If
                
               If nProducto = gCapAhorros Then
                    Dim oCons As COMDConstantes.DCOMConstantes
                    Set oCons = New COMDConstantes.DCOMConstantes
                    lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
                    lsDescTpoPrograma = Trim(oCons.DameDescripcionConstante(2030, lnTpoPrograma))
                    Set oCons = Nothing
               End If
               
             
            Case gCapPlazoFijo
                'lblUltContacto = rsCta("nPlazo") 'COMMENT BY MARG ERS065-2017
                Me.lblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.lblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
            
            Case gCapCTS
                'lblUltContacto = rsCta("cInstitucion") 'COMMENT BY MARG ERS065-2017
                Dim nDiasTranscurridos As Long
                'Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
                Dim nSaldoMinRet As Double
                    
                Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
                'nSaldoMinRet = clsDef.GetSaldoMinimoPersoneria(nProducto, nmoneda, nPersoneria, False)
                nSaldoMinRet = clsDef.GetSaldoMinimoPersoneria(nProducto, nmoneda, nPersoneria, False, , sCuenta) 'APRI20190109 ERS077-2018
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
            If nmoneda = gMonedaNacional Then
                cGetValorOpe = GetMontoDescuento(2117, 1, 1)
            Else
                cGetValorOpe = GetMontoDescuento(2118, 1, 2)
            End If
            
            'If Mid(sCuenta, 6, 3) = "232" Or Mid(sCuenta, 6, 3) = "234" Then
            If (Mid(sCuenta, 6, 3) = "232" And lnTpoPrograma <> 1) Or Mid(sCuenta, 6, 3) = "234" Then 'JUEZ 20140425 Para no cobrar comisión a cuentas de ahorro ñañito
                Set rsV = clsMant.ValidaTarjetizacion(sCuenta, lsTieneTarj)
                
                If lsTieneTarj = "SI" And rsV.RecordCount > 0 Then
                    'If MsgBox("El Cliente posee tarjeta se cobrará una comision, desea continuar con la operacion?", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'comment by marg ers065-2017
                    If MsgBox("El Cliente posee tarjeta, desea continuar con la operacion?", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers065-2017
                        Set loVistoElectronico = New frmVistoElectronico
                
                        lbVistoVal = loVistoElectronico.Inicio(5, nOperacion)
                    
                        If Not lbVistoVal Then
                            'MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones, se cobrara comision por esta operacion", vbInformation, "Mensaje del Sistema" 'comment by marg ers065-2017
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'comment by marg ers065-2017
                            Exit Sub
                        End If
                        
                        loVistoElectronico.RegistraVistoElectronico (0)
                    Else
                        cGetValorOpe = "0.00"
                        Exit Sub
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    'If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? se le cobrara una comision", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'comment by marg ers065-2017
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers065-2017
                        Set loVistoElectronico = New frmVistoElectronico
                
                        lbVistoVal = loVistoElectronico.Inicio(5, nOperacion)
                    
                        If Not lbVistoVal Then
                            'MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones, se cobrara comision por esta operacion", vbInformation, "Mensaje del Sistema" 'comment by marg ers065-2017
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'add by marg ers065-2017
                            Exit Sub
                        End If
                        
                        loVistoElectronico.RegistraVistoElectronico (0)
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
                lblMonComision = Format(cGetValorOpe, "#,##0.00")
                
                lblComision.Visible = True
                lblMonComision.Visible = True
                chkVBComision.Visible = False
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
        'fraDocumento.Enabled = True 'COMMENT BY MARG ERS065-2017
        'fraDocumentoTrans.Enabled = True 'COMMENT BY MARG ERS065-2017
        'fraMonto.Enabled = True 'COMMENT BY MARG ERS065-2017
        
        If gnCodOpeTarj <> 1 Then
            'If cboDocumento.Visible Then 'COMMENT BY MARG ERS065-2017
                'cboDocumento.Enabled = True 'COMMENT BY MARG ERS065-2017
            'ElseIf txtOrdenPago.Visible Then 'COMMENT BY MARG ERS065-2017
                'txtOrdenPago.SetFocus 'COMMENT BY MARG ERS065-2017
                
            'Else 'COMMENT BY MARG ERS065-2017
                'If txtGlosa.Visible Then txtGlosa.SetFocus 'COMMENT BY MARG ERS065-2017
            'End If 'COMMENT BY MARG ERS065-2017
        End If
        
        'fraCuenta.Enabled = False 'COMMENT BY MARG ERS065-2017
        
        'cmdGrabar.Enabled = True 'COMMENT BY MARG ERS065-2017
       
        'cmdCancelar.Enabled = True 'COMMENT BY MARG ERS065-2017
    End If
    
'    MuestraFirmas sCuenta
    
Else
    MsgBox sMsg, vbInformation, "Operacion"
    'txtCuenta.SetFocus 'COMMENT BY MARG ERS065-2017
End If
Set clsMant = Nothing
End Sub

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

Private Sub LimpiaControlesRetiro()
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
txtGlosa = ""

'If bDocumento Then
'
'End If
'txtMonto.BackColor = &HC0FFFF
'lblITF.BackColor = txtMonto.BackColor
'lblTotal.BackColor = txtMonto.BackColor

'lblComision.Visible = False
'lblMonComision.Visible = False
'chkVBComision.Visible = False
'lblMonComision = "0.00"
'txtMonto.value = 0

'lblMon.Caption = "S/."
'lblMensaje = ""
'cmdGrabar.Enabled = False
'txtCuenta.CMAC = gsCodCMAC
'txtCuenta.Age = ""
'txtCuenta.Cuenta = ""
'cmdGrabar.Enabled = False
'cmdCancelar.Enabled = False
'lblApertura = ""
'lblUltContacto = ""
lblFirmas = ""
lblTipoCuenta = ""
fraCliente.Enabled = False
'fraDatos.Enabled = False
'fraDocumento.Enabled = False
'fraMonto.Enabled = False
'fraCuenta.Enabled = True
txtOrdenPago.Text = ""
'
'txtCuenta.SetFocus

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
'lblExoneracion.Visible = False
'LblTituloExoneracion.Visible = False
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

Private Sub RestoreValidacionFirma()
    ReDim Preserve aPersonasInvol(0) 'default ubound(1)
    cVioFirma = ""
    aPersonasInvol(0) = ""
    gbTieneFirma = False
    bFirmasPendientes = False
    bFirmaObligatoria = False
End Sub

Private Sub GrabarRetiro()
'***Agregado por ELRO el 20130722, según TI-ERS079-2013****
If cboMedioRetiro.Visible = True Then
    If Trim(cboMedioRetiro) = "" Then
        'MsgBox "Debe seleccionar el medio de retiro.", vbInformation, "Aviso" 'comment by marg
        'cboMedioRetiro.SetFocus 'comment by marg
        SetearControles
        'Exit Sub 'comment by marg
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
    Dim loVistoElectronico As frmVistoElectronico 'RECO 20131022 ERS141
    Set loVistoElectronico = New frmVistoElectronico 'RECO 20131022 ERS141
    Dim loCaptaGen As COMNCaptaGenerales.NCOMCaptaGenerales  'RECO 20131022 ERS141
    Set loCaptaGen = New COMNCaptaGenerales.NCOMCaptaGenerales  'RECO 20131022 ERS141
        
    ' ***** Agregado Por RIRO el 20130501, Proyecto Ahorro - Poderes *****
        
    If bProcesoNuevo = True Then
    
        If validarReglasPersonas = False Then
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
    
    'sCuenta = txtCuenta.NroCuenta 'comment by marg
    sCuenta = Me.txtCuentaInstitucion.Text
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
                MsgBox "El Monto de Retiro Cta con Ord. Pago es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, "S/. ", "US$. ") & CStr(nMontoMinRet), vbOKOnly + vbInformation, "Aviso"
                Exit Sub
            End If
        Else
            'nMontoMinRet = clsDef.GetMontoMinimoRetPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag)
            nMontoMinRet = clsDef.GetMontoMinimoRetPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag, sCuenta) 'APRI20190109 ERS077-2018
            If nMontoMinRet > nMonto Then
                MsgBox "El Monto de Retiro es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, "S/. ", "US$. ") & CStr(nMontoMinRet), vbOKOnly + vbInformation, "Aviso"
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
                '***comment by marg****************
'''                If txtBanco.Visible Then
'''                    If txtBanco.Text = "" Then
'''                        MsgBox "Debe Seleccionar un Banco.", vbInformation, "Aviso"
'''                        Exit Sub
'''                    Else
'''                        sCodIF = txtBanco.Text
'''                    End If
'''                Else
'''                    sCodIF = ""
'''                End If
                '***end marg*************
            Else
                MsgBox "Orden de Pago No ha sido emitida para esta cuenta", vbInformation, "Aviso"
                Exit Sub
            End If
        ElseIf nDocumento = TpoDocNotaCargo Then
            sNroDoc = Trim(Left(cboDocumento.Text, 8))
            If InStr(1, sNroDoc, "<Nuevo>", vbTextCompare) > 0 Then
                MsgBox "Debe seleccionar un documento (" & fraDocumento.Caption & ") válido para la operacion.", vbInformation, "Aviso"
                'cboDocumento.SetFocus 'comment by marg
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
                
                
                Set loCptGen = New COMDCaptaGenerales.DCOMCaptaGenerales
                lbCuentaRRHH = loCptGen.ObtenerSiEsCuentaRRHH(sCuenta)
                If lbCuentaRRHH Then
                    nComixMov = 0
                    nComixRet = 0
                End If
                
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
            nMontoTemp = nMonto + nITF + lblMonComision
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
                If nmoneda = gMonedaNacional Then
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
                'nSaldoMinimo = clsGen.GetSaldoMinimoPersoneria(gCapAhorros, nmoneda, nPersoneria, True)
                nSaldoMinimo = clsGen.GetSaldoMinimoPersoneria(gCapAhorros, nmoneda, nPersoneria, True, , sCuenta) 'APRI20190109 ERS077-2018
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
                
                    oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nmoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF, psRegla:=ObtenerRegla
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
                                
                    'cmdCancelar_Click 'COMMENT BY MARG
                    Exit Sub
            
                ElseIf nSobregiro + 1 = nMaxSobregiro - 1 Then 'Una menos que la ultima se descuenta y se bloquea
                    MsgBox "NO POSEE SALDO SUFICIENTE" & Chr$(13) _
                      & "La Cuenta " & sCuenta & " ha sido Sobregirada " & nSobregiro + 1 & " VECES!." & Chr$(13) _
                      & "Se procederá a bloquear la cuenta y hacer el descuento por Orden de Pago Rechazada.", vbInformation, "Aviso"
                    'Hacer el descuento
                    sGlosa = "OP Rechazada " & sNroDoc & ". Cuenta " & sCuenta

                    oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nmoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF, psRegla:=ObtenerRegla
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
                    
                    'cmdCancelar_Click 'COMMENT BY MARG
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
                
                    'cmdCancelar_Click 'COMMENT BY MARG
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
        'If Not clsCap.CtaConFirmas(txtCuenta.NroCuenta) Then 'comment by marg
        If Not clsCap.CtaConFirmas(Me.txtCuentaInstitucion.Text) Then 'add by marg
            MsgBox "No puede retirar, porque la cuenta no cuenta con las firmas de las personas relacionadas a ella.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

    'Valida la operacion de Retiro por Transferencias
    If nOperacion = gAhoRetTransf Or nOperacion = gCTSRetTransf Then
        ' RIRO20131212 ERS137
        'If Trim(txtBanco.Text) = "" Then
        
        '***comment by marg********
'''        If Trim(txtBancoTrans.Text) = "" Then
'''            MsgBox "Debe seleccionar el Banco a Transferir", vbInformation, "Aviso"
'''            If txtBancoTrans.Enabled Then txtBancoTrans.SetFocus
'''            Exit Sub
'''        End If
        '***end marg******************
        
        'RIRO20131212 ERS137 - Comentado
        
        'If cboMonedaBanco.Text = "" Then
        '    MsgBox "Debe seleccionar la Moneda de la Cuenta de Banco a Transferir", vbInformation, "Aviso"
        '    cboMonedaBanco.SetFocus
        '    Exit Sub
        'End If
        
        ' RIRO20131212 ERS137
        'If Trim(txtCtaBanco.Text) = "" Then
        
        '***comment by marg*********
'''        If Trim(txtBancoTrans.Text) = "" Then
'''            MsgBox "Debe digitar la Cuenta de Banco a Transferir", vbInformation, "Aviso"
'''            txtCtaBanco.SetFocus
'''            Exit Sub
'''        End If
        '***end marg*****************
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

                If nmoneda = gMonedaNacional Then
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
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, sTipoCuenta, , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
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
                    
                    'If Mid(Trim(txtCuenta.NroCuenta), 9, 1) = 1 Then 'comment by marg ers065-2017
                    If Mid(Trim(Me.txtCuentaInstitucion.Text), 9, 1) = 1 Then 'add by marg ers065-2017
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
                        nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, , , , , , , sNumTarj) 'RECO20141127
                    Else
                        nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), psRegla:=ObtenerRegla)
                    End If
                    '***Fin Modificado por ELRO el 20130724, según TI-ERS079-2013
                ElseIf nOperacion = gCTSRetTransf Then
                    nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, getGlosa, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, Right(Me.txtBancoTrans.Text, 13), Me.txtCuentaTrans.Text, , 0, 0, 0, 0, 0, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300629", "300628"), , ObtenerRegla, lblComisionTransf.Caption, sMovNroTransf, getTitular, chkTransfEfectivo.value)     'RIRO20131212 ERS137 - Agregado
                    'nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, Right(Me.txtBanco.Text, 13), Me.txtCtaBanco.Text, CInt(Right(Me.cboMonedaBanco.Text, 3)), 0, 0, 0, 0, 0, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, nComisionVB, IIf(Me.chkVBComision.value = 0, "300627", "300628"), psRegla:=ObtenerRegla) 'RIRO20131212 ERS137 - Comentado
    '            ElseIf nOperacion = "220303" Then
    '                nMonto = CDbl(Me.TxtMonto2.Text)
    '                nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , CDbl(Val(Me.lblSalIntaD.Caption)), CDbl(Val(Me.lblIntIntaD.Caption)), CDbl(Val(Me.lblIntDisD.Caption)))
                End If
                
                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
        End Select
        
        ' MADM 20101115
        If nOperacion = "200304" Then
            Call CargoAutomatico(sCuenta, 1) 'MADM 20101115
        End If
        ' MADM 20101115
        
        If Me.chkVBComision.value = 1 Then
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
                        objPersona.insertarPersonaAgeOcupacionDatos gnMovNro, rsPersOcu!cperscod, IIf(nmoneda = gMonedaNacional, lblTotal, lblTotal * nTC), nAcumulado, gdFecSis, sMovNro
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
        Set clsLav = Nothing
        Set clsCap = Nothing
        Set loLavDinero = Nothing
        sNumTarj = ""
        gVarPublicas.LimpiaVarLavDinero
        'cmdCancelar_Click 'COMMENT BY MARG
        bRetiroExitoso = True 'ADD BY MARG ERS065-2017
    End If
    sNumTarj = ""
    'Unload Me 'COMMENT BY MARG ERS065-2017
    
    Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub
End Sub

Private Function validarReglasPersonas() As Boolean
 Dim sReglas() As String
    Dim sGrupos() As String
    Dim sTemporal As String
    Dim v1, v2 As Variant
    Dim bAprobado As Boolean
    Dim intRegla, i, J As Integer
    
    If Trim(strReglas) = "" Then
        validarReglasPersonas = False
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
        validarReglasPersonas = False
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
    validarReglasPersonas = bAprobado
End Function

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
            
        'If oCons.VerficaZonaAgencia(gsCodAge, Mid(txtCuenta.NroCuenta, 4, 2)) Then 'comment by marg
        If oCons.VerficaZonaAgencia(gsCodAge, Mid(Me.txtCuentaInstitucion.Text, 4, 2)) Then 'add by marg
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
                'Set rs = objCap.obtenerValTarifaOpeEnOtrasAge(gsCodAge, lnTpoPrograma, CInt(Mid(Me.txtCuenta.NroCuenta, 9, 1))) 'comment by marg
                'Set rs = objCap.obtenerValTarifaOpeEnOtrasAge(gsCodAge, lnTpoPrograma, CInt(Mid(Me.txtCuentaInstitucion.Text, 9, 1))) 'add by marg
                Set rs = objCap.obtenerValTarifaOpeEnOtrasAge(Me.txtCuentaInstitucion.Text) 'APRI20190109 ERS077-2018
                
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
                nmoneda = CLng(Mid(psCuenta, 9, 1))
                If nmoneda = gMonedaNacional Then
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
            nmoneda = CLng(Mid(psCuenta, 9, 1))
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
                nFlag = oCap.CapCobroComisionEfec(psCuenta, pnMonComVB, psCodOpe, sMovNro, "Comision Ope. Sin Tarjeta Efec.", , , gsNomAge, sLpt, nmoneda, gsCodCMAC, , gsCodAge, , lsmensaje, lsBoleta, gbImpTMU, pnMovNroOpe)
                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                Set oCap = Nothing
            Else
                bExito = False
            End If
        End If
    End If
End Sub

Private Function VerificarAutorizacion() As Boolean
Dim ocapaut As COMDCaptaGenerales.COMDCaptAutorizacion
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
Dim nmoneda As Moneda

'sCuenta = txtCuenta.NroCuenta 'comment by marg
sCuenta = Me.txtCuentaInstitucion.Text 'add by marg
nMonto = txtMonto.value
nmoneda = CLng(Mid(sCuenta, 9, 1))
'Obtiene los grupos al cual pertenece el usuario
Set oPers = New COMDPersona.UCOMAcceso
    gsGrupo = oPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
Set oPers = Nothing
 
'Verificar Montos
Set ocapaut = New COMDCaptaGenerales.COMDCaptAutorizacion
    'Set rs = ocapaut.ObtenerMontoTopNivAutRetCan(gsGrupo, "1", gsCodAge) RIRO20141105 ERS122
    Set rs = ocapaut.ObtenerMontoTopNivAutRetCan(gsGrupo, "1", gsCodAge, gsCodPersUser)
Set ocapaut = Nothing
 
If Not (rs.EOF And rs.BOF) Then
    lnMonTopD = rs("nTopDol")
    lnMonTopS = rs("nTopSol")
    sNivel = rs("cNivCod")
Else
    MsgBox "Usuario no Autorizado para realizar Operacion", vbInformation, "Aviso"
    VerificarAutorizacion = False
    Exit Function
End If

If nmoneda = gMonedaNacional Then
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
    oCapAutN.NuevaSolicitudAutorizacion sCuenta, "1", nMonto, gdFecSis, gsCodAge, gsCodUser, nmoneda, gOpeAutorizacionRetiro, sNivel, sMovNroAut
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

Private Sub CalculaComision()

Dim idBanco As String
Dim nPlaza As Integer
Dim nmoneda As Integer
Dim nTipo As Integer
Dim nMonto As Double
Dim nComision As Double
Dim oDefinicion As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oDefinicion = New COMNCaptaGenerales.NCOMCaptaDefinicion

'***COMMENT BY MARG***********************
'''If fraDocumentoTrans.Enabled And fraDocumentoTrans.Visible Then
'''
'''    If nOperacion = gAhoRetTransf Or nOperacion = gAhoCancTransfAbCtaBco Or _
'''       nOperacion = gPFRetIntAboCtaBanco Or nOperacion = gPFCancTransf Or _
'''       nOperacion = gCTSRetTransf Or nOperacion = gCTSCancTransfBco Then
'''
'''        idBanco = Mid(txtBancoTrans.Text, 4, 13)
'''        nPlaza = CDbl(Val(Trim(Right(cbPlazaTrans.Text, 8))))
'''        nmoneda = Val(Mid(txtCuenta.NroCuenta, 9, 1))
'''        nTipo = 102 ' Emision
'''        nMonto = txtMonto.value
'''        nComision = oDefinicion.getCalculaComision(idBanco, nPlaza, nmoneda, nTipo, nMonto, gdFecSis)
'''
'''    End If
'''
'''End If
'***END MARG**********************************
lblComisionTransf.Caption = Format(Round(nComision, 2), "#0.00")

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
'sCuenta = txtCuenta.NroCuenta 'comment by marg
sCuenta = Me.txtCuentaInstitucion.Text 'add by marg
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

Private Function getGlosa() As String
    Dim sGlosa As String
    sGlosa = "Bco Destino: " & lblNombreBancoTrans.Caption & ", Titular: " & getTitular & ", " & Trim(txtGlosaTrans.Text)
    getGlosa = UCase(sGlosa)
End Function

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

Private Sub cmdMostrarFirma_Click()
With grdCliente
    If .TextMatrix(.row, 1) = "" Then Exit Sub
    'Call frmPersonaFirma.Inicio(Trim(.TextMatrix(.row, 1)), Trim(txtCuenta.Age), True) 'comment by marg
    Call frmPersonaFirma.Inicio(Trim(.TextMatrix(.row, 1)), Mid(Trim(Me.txtCuentaInstitucion.Text), 4, 2), True) 'add by marg
    
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

Private Sub cmdVerRegla_Click()
    If strReglas <> "" Then
        Call frmCapVerReglas.Inicia(strReglas)
    Else
        MsgBox "Cuenta no tiene reglas definidas", vbInformation, "Aviso"
    End If
End Sub

'END MARG **********************************************************************************************************
