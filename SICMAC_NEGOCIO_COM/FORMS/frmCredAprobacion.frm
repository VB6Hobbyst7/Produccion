VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredAprobacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Crédito"
   ClientHeight    =   9915
   ClientLeft      =   3510
   ClientTop       =   2220
   ClientWidth     =   7695
   Icon            =   "frmCredAprobacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8475
      Left            =   75
      TabIndex        =   13
      Top             =   720
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   14949
      _Version        =   393216
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   697
      TabCaption(0)   =   "Datos &Cliente"
      TabPicture(0)   =   "frmCredAprobacion.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Solicitud"
      TabPicture(1)   =   "frmCredAprobacion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   " Sugerencia"
      TabPicture(2)   =   "frmCredAprobacion.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraEmpVinculados"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Aprobacion"
      TabPicture(3)   =   "frmCredAprobacion.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "btnVer"
      Tab(3).Control(1)=   "Frame22"
      Tab(3).Control(2)=   "FraNiv1"
      Tab(3).Control(3)=   "Frame5"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "&I y E"
      TabPicture(4)   =   "frmCredAprobacion.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame17"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Balance"
      TabPicture(5)   =   "frmCredAprobacion.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame18"
      Tab(5).Control(1)=   "Frame15"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Ga&rantias"
      TabPicture(6)   =   "frmCredAprobacion.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame20"
      Tab(6).Control(1)=   "Frame23"
      Tab(6).Control(2)=   "Frame24"
      Tab(6).ControlCount=   3
      Begin VB.Frame Frame24 
         Height          =   735
         Left            =   -74880
         TabIndex        =   244
         Top             =   6360
         Visible         =   0   'False
         Width           =   6975
         Begin VB.TextBox txtMontoMivivienda 
            Height          =   375
            Left            =   2760
            TabIndex        =   245
            Text            =   "0.00"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Valor de Venta (MIVIVIENDA)"
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
            Left            =   120
            TabIndex        =   246
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.CommandButton btnVer 
         Enabled         =   0   'False
         Height          =   390
         Left            =   -70360
         Picture         =   "frmCredAprobacion.frx":03CE
         Style           =   1  'Graphical
         TabIndex        =   243
         Top             =   3645
         Width           =   450
      End
      Begin VB.Frame Frame23 
         Caption         =   "Garantia Bien Futuro"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   -74880
         TabIndex        =   240
         Top             =   5160
         Width           =   6975
         Begin VB.CheckBox ckPolizaBF 
            Caption         =   "Poliza de Bien Futuro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   241
            Top             =   360
            Width           =   2295
         End
         Begin MSMask.MaskEdBox txtFechaPBF 
            Height          =   330
            Left            =   3120
            TabIndex        =   242
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame Frame22 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   -69220
         TabIndex        =   237
         Top             =   720
         Width           =   1605
         Begin VB.TextBox txtDiasVoucher 
            Height          =   285
            Left            =   960
            MaxLength       =   2
            TabIndex        =   238
            Top             =   120
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label lblDiasVoucher 
            AutoSize        =   -1  'True
            Caption         =   "D.Voucher:"
            Height          =   195
            Left            =   120
            TabIndex        =   239
            Top             =   180
            Visible         =   0   'False
            Width           =   810
         End
      End
      Begin VB.Frame FraNiv1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   207
         Top             =   600
         Width           =   7335
         Begin VB.CommandButton cmdEnvioEstCta 
            Caption         =   "Envio Estado Cta."
            Enabled         =   0   'False
            Height          =   360
            Left            =   5760
            TabIndex        =   256
            Top             =   600
            Width           =   1485
         End
         Begin VB.CheckBox chkTasa 
            Caption         =   "[Exoneración Tasa]"
            Height          =   255
            Left            =   5400
            TabIndex        =   249
            Top             =   2760
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox ckcPreferencial 
            Caption         =   "Preferencial"
            Height          =   255
            Left            =   5400
            TabIndex        =   248
            Top             =   3400
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox CmbDestino 
            Height          =   315
            ItemData        =   "frmCredAprobacion.frx":0787
            Left            =   1280
            List            =   "frmCredAprobacion.frx":0789
            Style           =   2  'Dropdown List
            TabIndex        =   220
            Top             =   240
            Width           =   1700
         End
         Begin VB.ComboBox CmbCondicion 
            Height          =   315
            Left            =   3990
            Style           =   2  'Dropdown List
            TabIndex        =   219
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CmbApoderado 
            Height          =   315
            Left            =   1280
            Style           =   2  'Dropdown List
            TabIndex        =   218
            Top             =   1035
            Width           =   5940
         End
         Begin VB.ComboBox CmbCondicionOtra 
            Height          =   315
            Left            =   1280
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   645
            Width           =   1700
         End
         Begin VB.ComboBox cmbFuentes 
            Height          =   315
            Left            =   5820
            Style           =   2  'Dropdown List
            TabIndex        =   215
            Top             =   1035
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Frame frmCtaAho 
            Caption         =   "Cuenta de Ahorros para Desembolsos(Terceros)"
            Enabled         =   0   'False
            Height          =   735
            Left            =   135
            TabIndex        =   211
            Top             =   2805
            Width           =   5175
            Begin VB.CommandButton cmdCtaAho 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   3915
               TabIndex        =   212
               Top             =   240
               Width           =   435
            End
            Begin SICMACT.ActXCodCta actxCtaAho 
               Height          =   435
               Left            =   120
               TabIndex        =   213
               Top             =   240
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   767
               Texto           =   "Cuenta N°"
               EnabledCMAC     =   -1  'True
               EnabledCta      =   -1  'True
               EnabledProd     =   -1  'True
               EnabledAge      =   -1  'True
            End
            Begin VB.Label lblTercero 
               Height          =   375
               Left            =   5520
               TabIndex        =   214
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.ComboBox CmbCampana 
            Height          =   315
            Left            =   1280
            Style           =   2  'Dropdown List
            TabIndex        =   210
            Top             =   1800
            Width           =   4240
         End
         Begin VB.CommandButton cmdVinculados 
            Caption         =   "Vinculados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5415
            TabIndex        =   208
            Top             =   3045
            Width           =   1455
         End
         Begin SICMACT.TxtBuscar txtBuscarLinea 
            Height          =   345
            Left            =   1260
            TabIndex        =   221
            Top             =   2175
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   609
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
            ForeColor       =   -2147483631
         End
         Begin VB.TextBox txtTasInt 
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
            Height          =   300
            Left            =   6240
            TabIndex        =   209
            Top             =   1800
            Width           =   960
         End
         Begin VB.TextBox txtInteresTasa 
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
            Height          =   300
            Left            =   6240
            TabIndex        =   250
            Top             =   1800
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.CommandButton cmdLineas 
            Height          =   315
            Left            =   6855
            Picture         =   "frmCredAprobacion.frx":078B
            Style           =   1  'Graphical
            TabIndex        =   216
            ToolTipText     =   "Buscar Lineas de Credito"
            Top             =   1780
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label lblFormEval 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   1280
            TabIndex        =   253
            Top             =   1440
            Width           =   5960
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Destino Crédito:"
            Height          =   195
            Left            =   130
            TabIndex        =   231
            Top             =   300
            Width           =   1125
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            Height          =   195
            Left            =   3190
            TabIndex        =   230
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Otra Condición:"
            Height          =   195
            Left            =   160
            TabIndex        =   229
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label LblCredNormalRefinan 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   3075
            TabIndex        =   228
            Top             =   645
            Width           =   2580
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            Caption         =   "Linea Crédito:"
            Height          =   195
            Left            =   270
            TabIndex        =   227
            Top             =   2250
            Width           =   975
         End
         Begin VB.Label lblLineaDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
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
            Height          =   495
            Left            =   3040
            TabIndex        =   226
            Top             =   2190
            Width           =   4185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Formato Eval.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   225
            Top             =   1485
            Width           =   1020
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Apoderado:"
            Height          =   195
            Left            =   420
            TabIndex        =   224
            Top             =   1090
            Width           =   825
         End
         Begin VB.Label Label77 
            Caption         =   "Campaña:"
            Height          =   255
            Left            =   540
            TabIndex        =   223
            Top             =   1880
            Width           =   855
         End
         Begin VB.Label Label78 
            Caption         =   "Tasa Int.:"
            Height          =   255
            Left            =   5535
            TabIndex        =   222
            Top             =   1875
            Width           =   735
         End
      End
      Begin VB.Frame fraEmpVinculados 
         Caption         =   "Empresas vinculadas - Producto EcoTaxi"
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
         Height          =   1845
         Left            =   240
         TabIndex        =   205
         Top             =   6240
         Width           =   6840
         Begin SICMACT.FlexEdit grdEmpVinculados 
            Height          =   1560
            Left            =   120
            TabIndex        =   206
            Top             =   240
            Width           =   6600
            _ExtentX        =   11642
            _ExtentY        =   2752
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "-Codigo-Empresa-Relacion-Monto Abono-Cuenta Abono"
            EncabezadosAnchos=   "400-0-4000-1500-1200-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-R-L"
            FormatosEdit    =   "0-0-0-0-2-0"
            lbUltimaInstancia=   -1  'True
            lbPuntero       =   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483635
         End
      End
      Begin VB.Frame Frame5 
         Height          =   7875
         Left            =   -74920
         TabIndex        =   71
         Top             =   450
         Width           =   7485
         Begin VB.Frame Frame8 
            Height          =   3930
            Left            =   180
            TabIndex        =   72
            Top             =   3840
            Width           =   7140
            Begin VB.CheckBox chkExoneraMYPE 
               Caption         =   "Exonerar Seguro MYPE"
               Height          =   255
               Left            =   3120
               TabIndex        =   257
               Top             =   3480
               Width           =   2055
            End
            Begin VB.Frame Frame12 
               Caption         =   "Datos Principales"
               Height          =   1640
               Left            =   120
               TabIndex        =   179
               Top             =   200
               Width           =   2955
               Begin VB.TextBox TxtMontoApr 
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
                  ForeColor       =   &H8000000D&
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   182
                  Top             =   200
                  Width           =   1290
               End
               Begin MSMask.MaskEdBox TxtFecDesemb 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   180
                  Top             =   540
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin Spinner.uSpinner SpnCuotasApr 
                  Height          =   300
                  Left            =   1560
                  TabIndex        =   181
                  Top             =   900
                  Width           =   630
                  _ExtentX        =   1111
                  _ExtentY        =   529
                  Max             =   350
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontName        =   "MS Sans Serif"
                  FontSize        =   8.25
               End
               Begin Spinner.uSpinner SpnPlazo 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   183
                  Top             =   1250
                  Width           =   630
                  _ExtentX        =   1111
                  _ExtentY        =   556
                  Max             =   3000
                  MaxLength       =   4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontName        =   "MS Sans Serif"
                  FontSize        =   8.25
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  Caption         =   "Monto:"
                  Height          =   195
                  Left            =   1040
                  TabIndex        =   187
                  Top             =   240
                  Width           =   495
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  Caption         =   "Cuotas:"
                  Height          =   195
                  Left            =   975
                  TabIndex        =   186
                  Top             =   980
                  Width           =   540
               End
               Begin VB.Label Label36 
                  AutoSize        =   -1  'True
                  Caption         =   "Plazo:"
                  Height          =   195
                  Left            =   1065
                  TabIndex        =   185
                  Top             =   1320
                  Width           =   435
               End
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Desembolso:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   184
                  Top             =   600
                  Width           =   1410
               End
            End
            Begin VB.Frame fraGracia 
               Height          =   975
               Left            =   5640
               TabIndex        =   173
               Top             =   840
               Width           =   1420
               Begin VB.TextBox TxtTasaGracia 
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
                  ForeColor       =   &H8000000D&
                  Height          =   285
                  Left            =   120
                  TabIndex        =   188
                  Top             =   1410
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.CommandButton CmdTiposGracia 
                  Caption         =   "--->"
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
                  Left            =   840
                  TabIndex        =   177
                  ToolTipText     =   "Definir Distribucion de Interes Gracia"
                  Top             =   1380
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.TextBox TxtDiasGracia 
                  Alignment       =   2  'Center
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
                  ForeColor       =   &H8000000D&
                  Height          =   285
                  Left            =   60
                  TabIndex        =   176
                  Top             =   340
                  Width           =   660
               End
               Begin VB.OptionButton optTipoGracia 
                  Caption         =   "Gracia en Cuotas"
                  Enabled         =   0   'False
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   175
                  Top             =   1095
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.OptionButton optTipoGracia 
                  Caption         =   "Capitalizar"
                  Enabled         =   0   'False
                  Height          =   375
                  Index           =   0
                  Left            =   120
                  TabIndex        =   174
                  Top             =   1020
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Día(s)"
                  ForeColor       =   &H80000011&
                  Height          =   195
                  Left            =   795
                  TabIndex        =   189
                  Top             =   400
                  Width           =   435
               End
               Begin VB.Label Label38 
                  AutoSize        =   -1  'True
                  Caption         =   "Periodo Gracia:"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   60
                  TabIndex        =   178
                  Top             =   140
                  Width           =   1095
               End
            End
            Begin VB.Frame fraMiViv 
               Caption         =   "Gastos Seguro Desgravamen"
               Height          =   705
               Left            =   120
               TabIndex        =   91
               Top             =   2880
               Visible         =   0   'False
               Width           =   2685
               Begin VB.OptionButton OptTipoGasto 
                  Caption         =   "&Fijo"
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   92
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   540
               End
               Begin VB.OptionButton OptTipoGasto 
                  Caption         =   "&Variable"
                  Height          =   195
                  Index           =   1
                  Left            =   1620
                  TabIndex        =   93
                  Top             =   270
                  Width           =   975
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "Tipo Cuota"
               ForeColor       =   &H80000008&
               Height          =   645
               Left            =   5640
               TabIndex        =   73
               Top             =   200
               Width           =   1420
               Begin VB.OptionButton OptTCuota 
                  Caption         =   "&Fija"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   2
                  Top             =   300
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton OptTCuota 
                  Caption         =   "C&reciente"
                  ForeColor       =   &H80000011&
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   76
                  Top             =   345
                  Visible         =   0   'False
                  Width           =   1080
               End
               Begin VB.OptionButton OptTCuota 
                  Caption         =   "D&ecreciente"
                  ForeColor       =   &H80000011&
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   75
                  Top             =   345
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.OptionButton OptTCuota 
                  Caption         =   "&Cuota Libre"
                  ForeColor       =   &H80000011&
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   74
                  Top             =   345
                  Visible         =   0   'False
                  Width           =   1185
               End
            End
            Begin VB.Frame Frame11 
               Caption         =   "Calendario"
               Height          =   765
               Left            =   140
               TabIndex        =   83
               Top             =   2880
               Visible         =   0   'False
               Width           =   2625
               Begin VB.OptionButton OptTipoCalend 
                  Caption         =   "&Fijo"
                  Enabled         =   0   'False
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   5
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   540
               End
               Begin VB.OptionButton OptTipoCalend 
                  Caption         =   "&Dinamico"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000011&
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   84
                  Top             =   490
                  Width           =   975
               End
            End
            Begin VB.Frame fratipodes 
               Caption         =   "Desembolso"
               Height          =   740
               Left            =   4320
               TabIndex        =   81
               Top             =   1920
               Width           =   2565
               Begin VB.OptionButton OptDesemb 
                  Caption         =   "&Parcial"
                  Height          =   225
                  Index           =   1
                  Left            =   1215
                  TabIndex        =   82
                  Top             =   315
                  Width           =   780
               End
               Begin VB.OptionButton OptDesemb 
                  Caption         =   "&Total"
                  Height          =   195
                  Index           =   0
                  Left            =   255
                  TabIndex        =   3
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   660
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "Tipo Pago"
               ForeColor       =   &H80000008&
               Height          =   1640
               Left            =   3120
               TabIndex        =   77
               Top             =   200
               Width           =   2475
               Begin MSMask.MaskEdBox txtFechaFija 
                  Height          =   315
                  Left            =   1260
                  TabIndex        =   192
                  ToolTipText     =   "Presione Enter"
                  Top             =   960
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.TextBox TxtDiaFijo2 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   645
                  MaxLength       =   2
                  TabIndex        =   170
                  Text            =   "00"
                  Top             =   2040
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.OptionButton OptTPer 
                  Caption         =   "&Periodo Fijo"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   4
                  Top             =   280
                  Value           =   -1  'True
                  Width           =   1215
               End
               Begin VB.OptionButton OptTPer 
                  Caption         =   "Fec&ha Fija"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   80
                  Top             =   540
                  Width           =   1080
               End
               Begin VB.TextBox TxtDiaFijo 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1260
                  MaxLength       =   2
                  TabIndex        =   79
                  Top             =   1320
                  Width           =   405
               End
               Begin VB.CheckBox ChkProxMes 
                  Caption         =   "Pro&x. Mes"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   120
                  TabIndex        =   78
                  Top             =   2040
                  Visible         =   0   'False
                  Width           =   1020
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "&Fecha Pago:"
                  Height          =   195
                  Left            =   320
                  TabIndex        =   204
                  Top             =   1005
                  Width           =   915
               End
               Begin VB.Label lblDia2 
                  AutoSize        =   -1  'True
                  Caption         =   "&Dia 2:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   172
                  Top             =   2100
                  Visible         =   0   'False
                  Width           =   420
               End
               Begin VB.Label LblDia 
                  AutoSize        =   -1  'True
                  Caption         =   "&Dia Fijo:"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   660
                  TabIndex        =   171
                  Top             =   1365
                  Width           =   570
               End
            End
            Begin VB.Frame FraTpoCalend 
               Caption         =   "Tipos de Calendario"
               Height          =   735
               Left            =   120
               TabIndex        =   88
               Top             =   1920
               Width           =   4140
               Begin VB.CommandButton cmdMIVIVIENDA 
                  Caption         =   "Ver Bono Buen Pagador"
                  Height          =   375
                  Left            =   2130
                  TabIndex        =   252
                  Top             =   220
                  Width           =   1935
               End
               Begin VB.CheckBox ChkCuotaCom 
                  Caption         =   "Calendario Con Cuota Comodin"
                  Height          =   240
                  Left            =   120
                  TabIndex        =   90
                  Top             =   735
                  Visible         =   0   'False
                  Width           =   2535
               End
               Begin VB.CheckBox ChkMiViv 
                  Caption         =   "Calendario Mi Vivienda"
                  Height          =   240
                  Left            =   120
                  TabIndex        =   89
                  Top             =   300
                  Width           =   1935
               End
            End
            Begin VB.Frame FraNiv2 
               BorderStyle     =   0  'None
               Caption         =   "Frame22"
               Height          =   1215
               Left            =   3120
               TabIndex        =   232
               Top             =   2520
               Width           =   3855
               Begin VB.CheckBox chkDesembBcoNac 
                  Caption         =   "&Desembolso en Banco de la Nación"
                  Height          =   240
                  Left            =   0
                  TabIndex        =   234
                  Top             =   720
                  Width           =   2985
               End
               Begin VB.CheckBox ChkBloqueo 
                  Caption         =   "Bloquear Desembolso"
                  Height          =   240
                  Left            =   0
                  TabIndex        =   236
                  Top             =   240
                  Width           =   1905
               End
               Begin VB.CheckBox ChkExoneracion 
                  Caption         =   "Exonerar Seguro Desgravamen"
                  Height          =   240
                  Left            =   0
                  TabIndex        =   235
                  Top             =   480
                  Width           =   2745
               End
               Begin VB.CommandButton cmbDesBcoNac 
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
                  Height          =   360
                  Left            =   3240
                  TabIndex        =   233
                  ToolTipText     =   "Desembolso en Banco de la Nación"
                  Top             =   600
                  Width           =   375
               End
            End
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Garantias"
         Height          =   4500
         Left            =   -74880
         TabIndex        =   163
         Top             =   600
         Width           =   6960
         Begin VB.CommandButton CmdGarantia 
            Caption         =   "Garantias"
            Height          =   330
            Left            =   4530
            TabIndex        =   168
            Top             =   2865
            Width           =   1515
         End
         Begin SICMACT.FlexEdit FEGarantCred 
            Height          =   2190
            Left            =   120
            TabIndex        =   164
            Top             =   240
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   3863
            Cols0           =   12
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Garantia-Gravament-Comercial-Realizacion-Disponible-Titular-Nro Docum-TipoDoc-cNumGarant-Legal-Poliza"
            EncabezadosAnchos=   "300-3800-1200-1200-1200-1200-3500-1200-0-1500-1800-1500"
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
            ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-R-R-R-R-L-L-L-C-L-L"
            FormatosEdit    =   "0-0-2-2-2-2-0-0-0-0-0-0"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "Total Garantias Gravament:"
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
            Height          =   195
            Left            =   120
            TabIndex        =   166
            Top             =   2880
            Width           =   2355
         End
         Begin VB.Label LblTotGarantiasGrav 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Left            =   3240
            TabIndex        =   165
            Top             =   2880
            Width           =   1200
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   6840
            Y1              =   2520
            Y2              =   2520
         End
      End
      Begin VB.Frame Frame15 
         Enabled         =   0   'False
         Height          =   4815
         Left            =   -74880
         TabIndex        =   122
         Top             =   480
         Width           =   6975
         Begin VB.Frame Frame19 
            Caption         =   "Ingresos y Egresos"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1650
            Left            =   40
            TabIndex        =   148
            Top             =   2880
            Width           =   6900
            Begin VB.TextBox TxtBalEgrFam 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00;(0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   5160
               MaxLength       =   13
               TabIndex        =   154
               Text            =   "0.00"
               Top             =   825
               Width           =   1530
            End
            Begin VB.TextBox TxtBalIngFam 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00;(0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   1875
               MaxLength       =   13
               TabIndex        =   153
               Text            =   "0.00"
               Top             =   840
               Width           =   1350
            End
            Begin VB.TextBox txtVentas 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00;(0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   1875
               MaxLength       =   13
               TabIndex        =   152
               Text            =   "0.00"
               Top             =   255
               Width           =   1350
            End
            Begin VB.TextBox txtrecuperacion 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0;(0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   1875
               MaxLength       =   13
               TabIndex        =   151
               Text            =   "0.00"
               Top             =   540
               Width           =   1350
            End
            Begin VB.TextBox txtcompras 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00;(0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   5160
               MaxLength       =   13
               TabIndex        =   150
               Text            =   "0.00"
               Top             =   240
               Width           =   1530
            End
            Begin VB.TextBox txtOtrosEgresos 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   5160
               MaxLength       =   13
               TabIndex        =   149
               Text            =   "0.00"
               Top             =   525
               Width           =   1530
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "Egresos Familiares"
               Height          =   195
               Left            =   3300
               TabIndex        =   162
               Top             =   870
               Width           =   1305
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               Caption         =   "Ingresos Familiares"
               Height          =   195
               Left            =   180
               TabIndex        =   161
               Top             =   870
               Width           =   1335
            End
            Begin VB.Line Line2 
               X1              =   105
               X2              =   6720
               Y1              =   1185
               Y2              =   1185
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
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
               Height          =   195
               Left            =   3645
               TabIndex        =   160
               Top             =   1305
               Width           =   555
            End
            Begin VB.Label LblSaldoIngEgr 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00;(0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
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
               Left            =   5175
               TabIndex        =   159
               Top             =   1245
               Width           =   1515
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "Ventas :"
               Height          =   195
               Left            =   195
               TabIndex        =   158
               Top             =   285
               Width           =   585
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               Caption         =   "Rec. de Ctas x Cobrar :"
               Height          =   195
               Left            =   195
               TabIndex        =   157
               Top             =   570
               Width           =   1650
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Costo de Ventas :"
               Height          =   195
               Left            =   3285
               TabIndex        =   156
               Top             =   255
               Width           =   1260
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "Otros Egresos :"
               Height          =   195
               Left            =   3285
               TabIndex        =   155
               Top             =   555
               Width           =   1080
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Balance de Situacion"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2340
            Left            =   40
            TabIndex        =   123
            Top             =   240
            Width           =   6900
            Begin VB.TextBox txtDisponible 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1605
               MaxLength       =   13
               TabIndex        =   130
               Text            =   "0.00"
               Top             =   900
               Width           =   1335
            End
            Begin VB.TextBox txtcuentas 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   1605
               MaxLength       =   13
               TabIndex        =   129
               Text            =   "0.00"
               Top             =   1185
               Width           =   1335
            End
            Begin VB.TextBox txtInventario 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   1605
               MaxLength       =   13
               TabIndex        =   128
               Text            =   "0.00"
               Top             =   1470
               Width           =   1335
            End
            Begin VB.TextBox txtactivofijo 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1830
               MaxLength       =   13
               TabIndex        =   127
               Text            =   "0.00"
               Top             =   1800
               Width           =   1320
            End
            Begin VB.TextBox txtPrestCmact 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   5100
               MaxLength       =   13
               TabIndex        =   126
               Text            =   "0.00"
               Top             =   1485
               Width           =   1485
            End
            Begin VB.TextBox txtOtrosPrest 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00;(0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   5100
               MaxLength       =   13
               TabIndex        =   125
               Text            =   "0.00"
               Top             =   1200
               Width           =   1485
            End
            Begin VB.TextBox txtProveedores 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00;(0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   5100
               MaxLength       =   13
               TabIndex        =   124
               Text            =   "0.00"
               Top             =   900
               Width           =   1485
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               Caption         =   "Activo :"
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
               Height          =   195
               Left            =   120
               TabIndex        =   147
               Top             =   345
               Width           =   675
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               Caption         =   "Activo Circulante :"
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
               Height          =   195
               Left            =   120
               TabIndex        =   146
               Top             =   615
               Width           =   1590
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               Caption         =   "Activo Fijo :"
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
               Height          =   195
               Left            =   120
               TabIndex        =   145
               Top             =   1860
               Width           =   1035
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "Pasivo y Patrimonio :"
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
               Height          =   195
               Left            =   3390
               TabIndex        =   144
               Top             =   315
               Width           =   1800
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               Caption         =   "Pasivo :"
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
               Height          =   195
               Left            =   3390
               TabIndex        =   143
               Top             =   615
               Width           =   705
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               Caption         =   "Patrimonio :"
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
               Height          =   195
               Left            =   3390
               TabIndex        =   142
               Top             =   1905
               Width           =   1020
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               Caption         =   "Disponible :"
               Height          =   195
               Left            =   120
               TabIndex        =   141
               Top             =   930
               Width           =   825
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               Caption         =   "Cuentas x Cobrar:"
               Height          =   195
               Left            =   120
               TabIndex        =   140
               Top             =   1215
               Width           =   1260
            End
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               Caption         =   "Inventario :"
               Height          =   195
               Left            =   120
               TabIndex        =   139
               Top             =   1500
               Width           =   795
            End
            Begin VB.Label lblActCirc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Left            =   1800
               TabIndex        =   138
               Top             =   600
               Width           =   1320
            End
            Begin VB.Label lblActivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Left            =   1800
               TabIndex        =   137
               Top             =   300
               Width           =   1320
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               Caption         =   "Prestamos CMACT"
               Height          =   195
               Left            =   3390
               TabIndex        =   136
               Top             =   1485
               Width           =   1335
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               Caption         =   "Otros Préstamos :"
               Height          =   195
               Left            =   3390
               TabIndex        =   135
               Top             =   1200
               Width           =   1245
            End
            Begin VB.Label Label72 
               AutoSize        =   -1  'True
               Caption         =   "Proveedores :"
               Height          =   195
               Left            =   3390
               TabIndex        =   134
               Top             =   945
               Width           =   990
            End
            Begin VB.Label lblPasivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Left            =   5400
               TabIndex        =   133
               Top             =   555
               Width           =   1320
            End
            Begin VB.Label lblPasPatrim 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Left            =   5400
               TabIndex        =   132
               Top             =   240
               Width           =   1320
            End
            Begin VB.Label lblPatrimonio 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
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
               Left            =   5355
               TabIndex        =   131
               Top             =   1830
               Width           =   1320
            End
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Frame18"
         Height          =   15
         Left            =   -73800
         TabIndex        =   121
         Top             =   480
         Width           =   615
      End
      Begin VB.Frame Frame17 
         Enabled         =   0   'False
         Height          =   4695
         Left            =   -74760
         TabIndex        =   100
         Top             =   480
         Width           =   6735
         Begin VB.TextBox TxtCargo 
            Height          =   285
            Left            =   1590
            TabIndex        =   117
            Top             =   2565
            Width           =   4560
         End
         Begin VB.Frame Frame13 
            Enabled         =   0   'False
            Height          =   1590
            Left            =   360
            TabIndex        =   104
            Top             =   360
            Width           =   5835
            Begin VB.TextBox TxtIngCon 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1440
               MaxLength       =   15
               TabIndex        =   108
               Text            =   "0.00"
               Top             =   495
               Width           =   1155
            End
            Begin VB.TextBox txtOtroIng 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1440
               MaxLength       =   15
               TabIndex        =   107
               Text            =   "0.00"
               Top             =   795
               Width           =   1155
            End
            Begin VB.TextBox txtIngFam 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   4455
               MaxLength       =   15
               TabIndex        =   106
               Text            =   "0.00"
               Top             =   210
               Width           =   1155
            End
            Begin VB.TextBox txtEgreFam 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   4455
               MaxLength       =   15
               TabIndex        =   105
               Text            =   "0.00"
               Top             =   525
               Width           =   1155
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               Caption         =   "Ingreso Cony:"
               Height          =   195
               Left            =   180
               TabIndex        =   116
               Top             =   570
               Width           =   975
            End
            Begin VB.Label LblIngresos 
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
               Height          =   255
               Left            =   1440
               TabIndex        =   115
               Top             =   225
               Width           =   1155
            End
            Begin VB.Line Line1 
               X1              =   195
               X2              =   5640
               Y1              =   1155
               Y2              =   1155
            End
            Begin VB.Label lblSaldo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00;(0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
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
               Left            =   4455
               TabIndex        =   114
               Top             =   1215
               Width           =   1200
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
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
               Height          =   195
               Left            =   2925
               TabIndex        =   113
               Top             =   1275
               Width           =   555
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               Caption         =   "Otros Ingresos:"
               Height          =   195
               Left            =   180
               TabIndex        =   112
               Top             =   870
               Width           =   1065
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               Caption         =   "Ingresos :"
               Height          =   195
               Left            =   210
               TabIndex        =   111
               Top             =   255
               Width           =   690
            End
            Begin VB.Label lblingreso 
               AutoSize        =   -1  'True
               Caption         =   "Ingreso Cliente :"
               Height          =   195
               Left            =   2925
               TabIndex        =   110
               Top             =   270
               Width           =   1140
            End
            Begin VB.Label lblEgreso 
               AutoSize        =   -1  'True
               Caption         =   "Egreso Familiar :"
               Height          =   195
               Left            =   2910
               TabIndex        =   109
               Top             =   585
               Width           =   1155
            End
         End
         Begin VB.Frame Frame14 
            Height          =   1530
            Left            =   120
            TabIndex        =   101
            Top             =   3000
            Width           =   6390
            Begin RichTextLib.RichTextBox Txtcomentarios 
               Height          =   1215
               Left            =   90
               TabIndex        =   102
               Top             =   225
               Width           =   6210
               _ExtentX        =   10954
               _ExtentY        =   2143
               _Version        =   393217
               Enabled         =   -1  'True
               MaxLength       =   300
               TextRTF         =   $"frmCredAprobacion.frx":09FD
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               Caption         =   "Comentario:"
               Height          =   195
               Left            =   75
               TabIndex        =   103
               Top             =   15
               Width           =   840
            End
         End
         Begin MSComCtl2.DTPicker DTPFecIni 
            Height          =   315
            Left            =   1605
            TabIndex        =   118
            Top             =   2130
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Format          =   134873089
            CurrentDate     =   37014
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Cargo : "
            Height          =   195
            Left            =   315
            TabIndex        =   120
            Top             =   2595
            Width           =   555
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Inicio :"
            Height          =   195
            Left            =   315
            TabIndex        =   119
            Top             =   2175
            Width           =   1185
         End
      End
      Begin VB.Frame Frame4 
         Height          =   5745
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   6840
         Begin VB.Frame Frame21 
            Caption         =   "Datos de  Credito"
            Height          =   1455
            Left            =   120
            TabIndex        =   197
            Top             =   240
            Width           =   6615
            Begin VB.Label lblInstitucionFinanciera 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1800
               TabIndex        =   203
               Top             =   960
               Width           =   4650
            End
            Begin VB.Label lblSubTipoCred 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1800
               TabIndex        =   202
               Top             =   600
               Width           =   4650
            End
            Begin VB.Label lblTipoCred 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1800
               TabIndex        =   201
               Top             =   240
               Width           =   4650
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Crédito:"
               Height          =   195
               Left            =   120
               TabIndex        =   200
               Top             =   360
               Width           =   900
            End
            Begin VB.Label label112 
               AutoSize        =   -1  'True
               Caption         =   "Sub Tipo:"
               Height          =   195
               Left            =   120
               TabIndex        =   199
               Top             =   720
               Width           =   690
            End
            Begin VB.Label Label123 
               Caption         =   "Insritución Financiera:"
               Height          =   255
               Left            =   120
               TabIndex        =   198
               Top             =   1080
               Width           =   1815
            End
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "# Con. Certicom :"
            Height          =   195
            Left            =   345
            TabIndex        =   191
            Top             =   3645
            Width           =   1230
         End
         Begin VB.Label lblNumConCer 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1575
            TabIndex        =   190
            Top             =   3600
            Width           =   675
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Int. Morat."
            Height          =   195
            Left            =   3585
            TabIndex        =   97
            Top             =   5340
            Width           =   720
         End
         Begin VB.Label LblIntMor 
            BackColor       =   &H8000000E&
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
            Height          =   285
            Left            =   5175
            TabIndex        =   96
            Top             =   5295
            Width           =   960
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Int. Compensat."
            Height          =   195
            Left            =   540
            TabIndex        =   95
            Top             =   5340
            Width           =   1110
         End
         Begin VB.Label LblIntComp 
            BackColor       =   &H8000000E&
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
            Height          =   285
            Left            =   1830
            TabIndex        =   94
            Top             =   5325
            Width           =   960
         End
         Begin VB.Label LblTasaGraciaSug 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5550
            TabIndex        =   87
            ToolTipText     =   "Tasa de Interes para Dias de Gracia"
            Top             =   2820
            Width           =   840
         End
         Begin VB.Label LblAnalista 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1575
            TabIndex        =   86
            Top             =   3195
            Width           =   4815
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Analista :"
            Height          =   195
            Left            =   375
            TabIndex        =   85
            Top             =   3255
            Width           =   645
         End
         Begin VB.Label LblTasaSug 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4800
            TabIndex        =   70
            ToolTipText     =   "Tasa de Interes Compensatorio"
            Top             =   2820
            Width           =   720
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Linea Credito :"
            Height          =   195
            Left            =   360
            TabIndex        =   69
            Top             =   2880
            Width           =   1020
         End
         Begin VB.Label LblLineaCredSug 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1575
            TabIndex        =   68
            Top             =   2835
            Width           =   3225
         End
         Begin VB.Label LblTipoGraciaSug 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1830
            TabIndex        =   67
            Top             =   4635
            Width           =   1485
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Gracia :"
            Height          =   195
            Left            =   540
            TabIndex        =   66
            Top             =   5010
            Width           =   1140
         End
         Begin VB.Label LblDiaFijoSug 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4830
            TabIndex        =   65
            Top             =   2475
            Width           =   675
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Dia Fijo :"
            Height          =   195
            Left            =   3585
            TabIndex        =   64
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label LblGraciaSug 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1575
            TabIndex        =   63
            Top             =   2475
            Width           =   675
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Dias Gracia :"
            Height          =   195
            Left            =   345
            TabIndex        =   62
            Top             =   2520
            Width           =   915
         End
         Begin VB.Label LblTipoCalendSug 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5175
            TabIndex        =   61
            Top             =   4650
            Width           =   1485
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Calendario :"
            Height          =   195
            Left            =   3540
            TabIndex        =   60
            Top             =   4335
            Width           =   1425
         End
         Begin VB.Label LblTipoDesembSug 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5175
            TabIndex        =   59
            Top             =   4245
            Width           =   1485
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Desembolso :"
            Height          =   195
            Left            =   3540
            TabIndex        =   58
            Top             =   3900
            Width           =   1545
         End
         Begin VB.Label LblTipoPeriodoSug 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1830
            TabIndex        =   57
            Top             =   4620
            Width           =   1485
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Periodo :"
            Height          =   195
            Left            =   540
            TabIndex        =   56
            Top             =   4635
            Width           =   1215
         End
         Begin VB.Label LblTipoCuotaSug 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1830
            TabIndex        =   55
            Top             =   4260
            Width           =   1485
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cuota :"
            Height          =   195
            Left            =   540
            TabIndex        =   54
            Top             =   4290
            Width           =   1095
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Caracteristicas de Credito :"
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
            Height          =   195
            Left            =   225
            TabIndex        =   53
            Top             =   3990
            Width           =   2325
         End
         Begin VB.Label LblCuotaSug 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4830
            TabIndex        =   52
            Top             =   1815
            Width           =   1020
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Cuota Sugerida :"
            Height          =   195
            Left            =   3585
            TabIndex        =   51
            Top             =   1860
            Width           =   1185
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   345
            TabIndex        =   50
            Top             =   1860
            Width           =   540
         End
         Begin VB.Label LblMontoSug 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1575
            TabIndex        =   49
            Top             =   1815
            Width           =   1155
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas :"
            Height          =   195
            Left            =   345
            TabIndex        =   48
            Top             =   2190
            Width           =   585
         End
         Begin VB.Label LblCuotasSug 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1575
            TabIndex        =   47
            Top             =   2145
            Width           =   675
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Plazo :"
            Height          =   195
            Left            =   3585
            TabIndex        =   46
            Top             =   2190
            Width           =   480
         End
         Begin VB.Label LblPlazoSug 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4830
            TabIndex        =   45
            Top             =   2145
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "dias"
            Height          =   195
            Left            =   5595
            TabIndex        =   44
            Top             =   2190
            Width           =   285
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4860
         Left            =   -74880
         TabIndex        =   29
         Top             =   675
         Width           =   6915
         Begin VB.Frame Frame3 
            Height          =   2565
            Left            =   120
            TabIndex        =   42
            Top             =   2160
            Width           =   6660
            Begin SICMACT.FlexEdit FERelacCred 
               Height          =   2160
               Left            =   90
               TabIndex        =   1
               Top             =   255
               Width           =   6480
               _extentx        =   11430
               _extenty        =   3810
               cols0           =   3
               highlight       =   1
               allowuserresizing=   3
               rowsizingmode   =   1
               encabezadosnombres=   "-Cliente-Relacion"
               encabezadosanchos=   "400-3800-1700"
               font            =   "frmCredAprobacion.frx":0A80
               font            =   "frmCredAprobacion.frx":0AA8
               font            =   "frmCredAprobacion.frx":0AD0
               font            =   "frmCredAprobacion.frx":0AF8
               font            =   "frmCredAprobacion.frx":0B20
               fontfixed       =   "frmCredAprobacion.frx":0B48
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               lbultimainstancia=   -1
               columnasaeditar =   "X-X-X"
               textstylefixed  =   4
               listacontroles  =   "0-0-0"
               encabezadosalineacion=   "C-C-C"
               formatosedit    =   "0-0-0"
               lbpuntero       =   -1
               colwidth0       =   405
               rowheight0      =   300
               forecolorfixed  =   -2147483635
            End
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Sub Producto :"
            Height          =   195
            Left            =   210
            TabIndex        =   196
            Top             =   990
            Width           =   1065
         End
         Begin VB.Label lblSubProducto 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   195
            Top             =   960
            Width           =   5010
         End
         Begin VB.Label LblCondicionOtra 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4440
            TabIndex        =   99
            Top             =   285
            Width           =   1425
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Otra Condicion :"
            Height          =   195
            Left            =   3120
            TabIndex        =   98
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "dias"
            Height          =   195
            Left            =   4620
            TabIndex        =   40
            Top             =   1650
            Width           =   285
         End
         Begin VB.Label LblPlazoSol 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3870
            TabIndex        =   39
            Top             =   1605
            Width           =   675
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Plazo :"
            Height          =   195
            Left            =   3285
            TabIndex        =   38
            Top             =   1650
            Width           =   480
         End
         Begin VB.Label LblCuotasSol 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   37
            Top             =   1605
            Width           =   675
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas :"
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   1650
            Width           =   585
         End
         Begin VB.Label LblMontoSol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   35
            Top             =   1275
            Width           =   1395
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   210
            TabIndex        =   34
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label LblProducto 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            Top             =   600
            Width           =   5010
         End
         Begin VB.Label LblCredTipoTit 
            AutoSize        =   -1  'True
            Caption         =   "Producto :"
            Height          =   195
            Left            =   210
            TabIndex        =   32
            Top             =   705
            Width           =   735
         End
         Begin VB.Label lblCondicion 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   31
            Top             =   280
            Width           =   1425
         End
         Begin VB.Label LblCredCond 
            AutoSize        =   -1  'True
            Caption         =   "Condicion :"
            Height          =   195
            Left            =   210
            TabIndex        =   30
            Top             =   360
            Width           =   1140
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4425
         Left            =   -74850
         TabIndex        =   15
         Top             =   660
         Width           =   7290
         Begin VB.CommandButton CmdCredVig 
            Caption         =   "Creditos Vigentes"
            Height          =   375
            Left            =   2190
            TabIndex        =   167
            ToolTipText     =   "Creditos Vigentes"
            Top             =   2940
            Width           =   2505
         End
         Begin VB.Label lblFormEvalDatCli 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   254
            Top             =   1320
            Width           =   4935
         End
         Begin VB.Label Label19 
            Caption         =   "Creditos en Agencia :"
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
            Height          =   210
            Left            =   240
            TabIndex        =   41
            Top             =   3015
            Width           =   1845
         End
         Begin VB.Label LblDireccion 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1800
            TabIndex        =   28
            Top             =   960
            Width           =   4935
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Direccion : "
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label LblDescripFte 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1800
            TabIndex        =   26
            Top             =   2340
            Visible         =   0   'False
            Width           =   4470
         End
         Begin VB.Label LblCodFte 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2400
            TabIndex        =   25
            Top             =   2010
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "&Descripcion :"
            Height          =   195
            Left            =   420
            TabIndex        =   24
            Top             =   2400
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Formato Evaluación:"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fuente de Ingreso :"
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
            Height          =   195
            Left            =   210
            TabIndex        =   22
            Top             =   2010
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label LblNomCli 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1800
            TabIndex        =   21
            Top             =   600
            Width           =   4935
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Apellidos y Nombres :"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label LblCodCli 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo :"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Solicitud"
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
         Height          =   195
         Left            =   -72720
         TabIndex        =   19
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Solicitud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   -72720
         TabIndex        =   18
         Top             =   480
         Width           =   750
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Credito"
      ForeColor       =   &H80000006&
      Height          =   720
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   7605
      Begin VB.CommandButton cmdVerEntidades 
         Caption         =   "Ver Entidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   247
         Top             =   240
         Width           =   1335
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   165
         TabIndex        =   0
         Top             =   255
         Width           =   3660
         _extentx        =   6456
         _extenty        =   741
         texto           =   "Credito"
         enabledcmac     =   -1
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
      Begin VB.CommandButton cmdbuscar 
         Caption         =   "E&xaminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         ToolTipText     =   "Buscar Credito"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   60
      TabIndex        =   14
      Top             =   9240
      Width           =   7605
      Begin VB.CommandButton cmdCheckList 
         Caption         =   "CheckList"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4720
         TabIndex        =   251
         Top             =   180
         Width           =   980
      End
      Begin VB.CommandButton cmdEvaluacion 
         Caption         =   "Evaluación"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3750
         TabIndex        =   194
         Top             =   180
         Width           =   980
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "Ca&ncelar"
         Height          =   345
         Left            =   5700
         TabIndex        =   193
         ToolTipText     =   "Limpiar la Pantalla"
         Top             =   180
         Width           =   900
      End
      Begin VB.CommandButton CmdGastos 
         Caption         =   "&Gastos"
         Enabled         =   0   'False
         Height          =   345
         Left            =   2850
         TabIndex        =   9
         ToolTipText     =   "Generar Plan Gastos"
         Top             =   180
         Width           =   900
      End
      Begin VB.CommandButton CmdDesembolsos 
         Caption         =   "&Desemb."
         Enabled         =   0   'False
         Height          =   345
         Left            =   1880
         TabIndex        =   8
         ToolTipText     =   "Ingresar los Desembolsos Parciales"
         Top             =   180
         Width           =   980
      End
      Begin VB.CommandButton CmdCalend 
         Caption         =   "Calen&dario"
         Enabled         =   0   'False
         Height          =   345
         Left            =   960
         TabIndex        =   7
         ToolTipText     =   "Mostrar el Calendario de Pagos"
         Top             =   180
         Width           =   910
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   6600
         TabIndex        =   10
         Top             =   180
         Width           =   900
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Grabar Datos de Sugerencia"
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.CheckBox ChkReglamento 
      Caption         =   "Aprobación con Exoneración de Reglamento"
      Height          =   240
      Left            =   120
      TabIndex        =   255
      Top             =   8880
      Width           =   3585
   End
   Begin VB.Label Label73 
      AutoSize        =   -1  'True
      Caption         =   "Monto :"
      Height          =   195
      Left            =   270
      TabIndex        =   169
      Top             =   2910
      Width           =   540
   End
End
Attribute VB_Name = "frmCredAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredAprobacion
'***     Descripcion:       Opcion que Realiza la Aprobacion del Credito Pasando
'                           el Credito a un Nuevo estado
'***     Creado por:        NSSE
'***     Maquina:           07SISTDES03
'***     Fecha-Tiempo:         12/07/2001 10:52:16 AM
'***     Ultima Modificacion: Fecha de Crecion
'*****************************************************************************************
Option Explicit

Private bGraciaGen As Boolean
Private bCalendGen As Boolean
Private bGastosGen As Boolean
Private bDesembParcialGenerado As Boolean
Private vnTipoGracia As Integer
Private MatGracia As Variant
Private MatCalend As Variant
Private MatCalend_2 As Variant
Private MatDesemb As Variant
Private MatGastos As Variant
Private nNumGastos As Integer
Private bCapitalInt As Boolean
Private bRefinanc As Boolean
Private bSustiDeudor As Boolean
Private vsLineaCredito As String
Private nUltNroCalend As Integer
Private MatCredVig As Variant
''
Private oPersona As UPersona_Cli
Private nIndice As Integer
Private nProcesoEjecutado As Integer '1 Nueva fte de Ingreso; 2 Editar fte de Ingreso ; 3 Consulta de Fte
Private nTipoFteIngreso As Integer

'-------------------------------------------
Private nMontoCuota As Double
Private cMonedaFteCod As String

Dim nIndtempo As Integer
'-------------------------------------------
Dim sLineaTmp As String

Dim nTasaComp As Double
Dim nTasaMor As Double
Dim nTasaGrac As Double

Dim nEstadoPoli As Integer '***PEAC 20080204

Dim nPlazo As Integer ' Add Gitu 20-08-2008

'Para tenerlo cargado en memoria los valores
Dim sMontosARef As Variant
Dim bQuincenal As Boolean
'Actualizacion de Filtros de Lineas de Credito
Dim bBuscarLineas As Boolean
Dim nTasCosEfeAnu As Double 'DAOR 20070419, Tasa Costo Efectivo Anual
'**DAOR 20080109 ***************************************
Private fnPersPersoneria As Integer
Private fnPersTipoDoc As Integer
Private fsPersNumeDoc As String
'*******************************************************

'Add by gitu 06-04-2009
Dim lsNumfuente As String
Dim lnPersTipFte As Integer
Dim ldPersEval As String

'By Capi 10042008
Dim sCliCtaAho As String
Dim objPista As COMManejador.Pista  '' *** PEAC 20080126
'ALPA 20100607 B2******************
Dim lsTpoProdCod As String
Dim lsTpoCredCod As String
'**********************************
'MADM 20110303 ********************
Dim lsTpoIDNroRUC As String
Dim lsTpoIDNroDNI As String
Dim objPersona As COMDPersona.UCOMPersona
'**********************************
Dim bLeasing As Boolean 'ALPA 20111212******************************
Private nCondCred As Integer 'MAVM 20110613
Dim nAgenciaCredEval As Integer '** JUEZ 20120907
Dim bAprobNivPendiente As Boolean '** JUEZ 20121217

Dim nIntGraInicial As Double 'MAVM 20130209 ***
Dim fbValidaExoSegDes As Boolean 'WIOR 20130418
Dim fbLimpiaExoSegDes As Boolean 'WIOR 20130710
Private fnCuotaBalon As Integer 'WIOR 20131111
'ALPA20130219*********************************
Dim nPolizaBF As Integer
Dim dFechaPBF As Date
'*********************************************
Private MatCalendSegDes As Variant 'WIOR 20140825
Dim lnCSP As Integer 'ALPA 20141126
Dim oRsVerEntidades As ADODB.Recordset 'ALPA20141021***
Dim lnCantidadVerEntidades As Integer 'ALPA20141021***
Dim bCantidadVerEntidadesCmac As Integer 'ALPA20141021***
Dim nLogicoVerEntidades As Integer 'ALPA20141201
Dim lnColocDestino As Integer 'ALPA20141201
Dim lnColocCondicion As Integer
Dim lnCuotaAprobada As Currency 'ALPA20141201
'ALPA 20150115*********************************************
Dim lnTasaInicial As Currency
Dim lnTasaFinal As Currency
Dim RLinea As ADODB.Recordset
Dim lnCampanaId As Integer
Dim lnLogicoBuscarDatos As Integer
Dim lnExoneracionTasa As Integer
Dim lnCliPreferencial As Integer
Dim rsExonera As ADODB.Recordset
Dim lbExoneraTasa As Integer 'ALPA 20150210
Dim bCheckList As Boolean 'RECO20150421 ERS010-2015
Private fbGrabar As Boolean 'WIOR 20150717
'WIOR 20151223 ***
Private fbMIVIVIENDA As Boolean
Private fArrMIVIVIENDA As Variant
'WIOR FIN ********
Dim fvGravamen() As tGarantiaGravamen 'EJVG20150513
Dim fbSolicitudPoliza As Boolean 'EJVG20150704
'WIOR 20160224 ***
Private fnTasaSegDes As Double
Private fnCantAfiliadosSegDes As Integer
'WIOR FIN *********
Dim fbSolicitudAutAmp As Boolean 'JUEZ 20160510
Dim fbEsAmpliado As Boolean 'EJVG20160530
'APRI20180404 ERS036-2017
Dim fbRegistraEnvio As Boolean
Dim Clientes(20, 2) As String
Dim Ini, counter As Integer
'END APRI

'JOEP20190131 CP
Dim lsCategoria As String
Dim lsTipoCredito As String
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
'JOEP20190131 CP
'*******************************************************************

Public Sub Aprobacion(Optional ByVal pbLeasing As Boolean = False)
    bLeasing = pbLeasing
    If bLeasing = True Then
        Me.Caption = "Aprobación de Arrendamiento Financiero"
        ActxCta.texto = "Operación"
        Frame7.Caption = "Operación"
    End If
    fbLimpiaExoSegDes = False 'WIOR 20130710
    Me.Show 1
End Sub

Private Function DameTipoCuota() As Integer
Dim i As Integer
    For i = 0 To 3
        If OptTCuota(i).value Then
            DameTipoCuota = i + 1
            Exit For
        End If
    Next i
End Function

Private Sub HabilitaCuotaLibre(ByVal pbHabilita As Boolean)
    OptTPer(0).value = Not pbHabilita
    OptTPer(0).Enabled = Not pbHabilita
    OptTPer(1).Enabled = Not pbHabilita
    OptTipoCalend(0).value = Not pbHabilita
    OptTipoCalend(0).Enabled = Not pbHabilita
    OptTipoCalend(1).Enabled = Not pbHabilita
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    'Dim R As ADODB.Recordset
    'Dim oCredito As COMDCredito.DCOMCredito
    Dim oNCredito As COMNCredito.NCOMCredito
    'Dim oCalend As COMDCredito.DCOMCalendario
    Dim RDes As ADODB.Recordset
    Dim RGar As ADODB.Recordset
    Dim nMontoGravado As Double
    'Dim oCal As COMDCredito.DCOMCalendario
    'Dim R2 As ADODB.Recordset
    Dim RCal As ADODB.Recordset
    Dim RDatos As ADODB.Recordset
    Dim RRel As ADODB.Recordset
    Dim rsLinea As ADODB.Recordset
    Dim rsFte As ADODB.Recordset
    Dim rsDesBcoNac As ADODB.Recordset 'DAOR 20070202
    Dim rsFuentes As ADODB.Recordset  'ARCV 30-12-2006
    Dim rsRelEmp As ADODB.Recordset 'BRGO 20111107
    Dim rsCredEval As ADODB.Recordset 'JUEZ 20120907
    Dim i As Integer
    nPolizaBF = 0
    'Dim oDGarantia As COMDCredito.DCOMGarantia 'MADM 20110505
    'Dim rsGarantReal As ADODB.Recordset 'MADM 20110505
    'Dim nCerrar As Integer
    'WIOR 20160618 ***
    Dim oDCredito As COMDCredito.DCOMCredito
    Dim rsSobreEnd As ADODB.Recordset
    'WIOR FIN ********
    
    'RECO20140226 ERS174-2013*************************
    Dim oDNivExo As COMDCredito.DCOMNivelAprobacion
    Set oDNivExo = New COMDCredito.DCOMNivelAprobacion
    Dim rsExoAut As ADODB.Recordset
    'RECO FIN*****************************************
    
    'ALPA 20111212********************************************
    Dim nLogicoLeasingDesembolso As Integer
    Dim oDLeasing As COMDCredito.DCOMleasing
    Dim oRS As ADODB.Recordset
    Dim nNroProveedores As Integer
    Dim rsPolizaSolicitud As New ADODB.Recordset 'EJVG20150616

    nLogicoLeasingDesembolso = 0
    Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
    If objProducto.GetResultadoCondicionCatalogo("N0000072", ActxCta.Prod) Then
    'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
        nLogicoLeasingDesembolso = 1
        Set oDLeasing = New COMDCredito.DCOMleasing
        Set oRS = New ADODB.Recordset
        Set oRS = oDLeasing.ObtenerAprobacionLeasing(ActxCta.NroCuenta, nNroProveedores)
        Set oRS = Nothing
        If nNroProveedores > 0 Then
            nLogicoLeasingDesembolso = 2
        End If
    End If
     nLogicoLeasingDesembolso = 2
    '*********************************************************

    '*** PEAC 20160822
    Dim oCredFormEval As COMDCredito.DCOMFormatosEval
    Set oCredFormEval = New COMDCredito.DCOMFormatosEval
    Dim oRSDatFormEval As ADODB.Recordset
    Dim lsFormEval As String
    Set oRSDatFormEval = oCredFormEval.ObtieneDatosComunes(ActxCta.NroCuenta)
    lsFormEval = oRSDatFormEval!cFormato
    oRSDatFormEval.Close
    Set oRSDatFormEval = Nothing
    Set oCredFormEval = Nothing
    Me.lblFormEval.Caption = Trim(UCase(lsFormEval))
    lblFormEvalDatCli.Caption = Trim(UCase(lsFormEval))
    'FIN PEAC

'CTI3 ERS032020*********************************************************
    'WIOR 20160618 ***
'    Set oDCredito = New COMDCredito.DCOMCredito
'    Set rsSobreEnd = oDCredito.SobreEndObtenerCodigosRegXCta(psCtaCod)
'    If Not (rsSobreEnd.EOF And rsSobreEnd.BOF) Then
'        Dim sCodigosPotSobreEnd As String
'        Dim sCodigosSobreEnd As String
'        sCodigosPotSobreEnd = ""
'        sCodigosSobreEnd = ""
'
'        sCodigosPotSobreEnd = IIf(CInt(rsSobreEnd!nCodigo1) = 1, "1,", "") & _
'                                IIf(CInt(rsSobreEnd!nCodigo2) = 1, "2,", "") & _
'                                IIf(CInt(rsSobreEnd!nCodigo3) = 1, "3,", "") & _
'                                IIf(CInt(rsSobreEnd!nCodigo4) = 1, "4,", "") & _
'                                IIf(CInt(rsSobreEnd!nCodigo5) = 1, "5,", "")
'        If Len(sCodigosPotSobreEnd) > 0 Then
'            sCodigosPotSobreEnd = Mid(sCodigosPotSobreEnd, 1, Len(sCodigosPotSobreEnd) - 1)
'            If Len(sCodigosPotSobreEnd) > 1 Then
'                sCodigosPotSobreEnd = Mid(sCodigosPotSobreEnd, 1, Len(sCodigosPotSobreEnd) - 2) & " y " & Mid(sCodigosPotSobreEnd, Len(sCodigosPotSobreEnd), 2)
'            End If
'        End If
'
'        sCodigosSobreEnd = IIf(CInt(rsSobreEnd!nCodigo1) = 2, "1,", "") & _
'                                IIf(CInt(rsSobreEnd!nCodigo2) = 2, "2,", "") & _
'                                IIf(CInt(rsSobreEnd!nCodigo3) = 2, "3,", "") & _
'                                IIf(CInt(rsSobreEnd!nCodigo4) = 2, "4,", "") & _
'                                IIf(CInt(rsSobreEnd!nCodigo5) = 2, "5,", "")
'        If Len(sCodigosSobreEnd) > 0 Then
'            sCodigosSobreEnd = Mid(sCodigosSobreEnd, 1, Len(sCodigosSobreEnd) - 1)
'            If Len(sCodigosSobreEnd) > 1 Then
'                sCodigosSobreEnd = Mid(sCodigosSobreEnd, 1, Len(sCodigosSobreEnd) - 2) & " y " & Mid(sCodigosSobreEnd, Len(sCodigosSobreEnd), 2)
'            End If
'        End If
'
'        If CInt(rsSobreEnd!nEstado) = 0 Then
'            CargaDatos = False
'            MsgBox "El crédito se encuentra bloqueado temporalmente por Sobreendeudamiento." & Chr(10) & Chr(10) & "Favor de solicitar su Desbloqueo.", vbInformation, "Aviso"
'            Exit Function
'        ElseIf CInt(rsSobreEnd!nEstado) = 1 Then
'            CargaDatos = False
'            MsgBox "Crédito ya fue desbloqueado por Potencial Sobreendeudamiento(Codigo(s): " & sCodigosPotSobreEnd & ")" & _
'            ", faltando aún el desbloqueo por SobreEndeudamiento(Codigo(s): " & sCodigosSobreEnd & ")." & Chr(10) & Chr(10) & "Favor de solicitar su Desbloqueo respetivo.", vbInformation, "Aviso"
'            Exit Function
'        ElseIf CInt(rsSobreEnd!nEstado) = 2 Then
'            CargaDatos = False
'            MsgBox "Crédito ya fue desbloqueado por Sobreendeudamiento(Codigo(s): " & sCodigosSobreEnd & ")" & _
'            ", faltando aún el desbloqueo por Potencial SobreEndeudamiento(Codigo(s): " & sCodigosPotSobreEnd & ")." & Chr(10) & Chr(10) & "Favor de solicitar su Desbloqueo respetivo.", vbInformation, "Aviso"
'            Exit Function
'        End If
'    End If
'    Set rsSobreEnd = Nothing
    'WIOR FIN ********
'************************************************************************************************
    On Error GoTo ErrorCargaDatos
    fbSolicitudPoliza = False 'EJVG20150703
    fbSolicitudAutAmp = False 'JUEZ 20160510
    MatCredVig = ""
    Set oNCredito = New COMNCredito.NCOMCredito
    If nLogicoLeasingDesembolso = 0 Or nLogicoLeasingDesembolso = 2 Then
    Call oNCredito.CargarDatosAprobacion(psCtaCod, gdFecSis, Mid(ActxCta.NroCuenta, 6, 3), _
                                        Mid(ActxCta.NroCuenta, 9, 1), RDatos, RCal, bRefinanc, bSustiDeudor, _
                                         RGar, nMontoGravado, RRel, RDes, sLineaTmp, rsLinea, rsFte, sMontosARef, _
                                         bQuincenal, , rsFuentes, rsDesBcoNac, , , rsRelEmp, rsCredEval, rsPolizaSolicitud, fbEsAmpliado) 'DAOR 20070202, se aumento rsDesBcoNac
    
    Set oNCredito = Nothing
    Set rsExonera = Nothing   'ALPA20141202
    'nCerrar = 0
    lnLogicoBuscarDatos = 0 'APA20150116
    nLogicoVerEntidades = 0 'ALPA 20141231
    'Set oCredito = New COMDCredito.DCOMCredito
    'Set R = oCredito.RecuperaDatosAprobacion(psCtaCod)
    
    If Not RDatos.BOF And Not RDatos.EOF Then
        'RECO20160328 ERS013-2016*********************************
        If UCase(RDatos!cAnalistaCod) = UCase(gsCodUser) Then
            CargaDatos = False
            MsgBox "Usted no puede Aprobar un crédito que haya Solicitado.", vbInformation, "Alerta SICMACM"
            Exit Function
        End If
        'RECO FIN*************************************************
        'EJVG20160428 ***
        If RGar.RecordCount = 0 Then
            CargaDatos = False
            MsgBox "El crédito no cuenta con Registro de Cobertura, comuníqueselo al Analista del Crédito.", vbInformation, "Aviso"
            Exit Function
        End If
        'END EJVG *******
        'DAOR 20070202 ***************
        If Not rsDesBcoNac.BOF And Not rsDesBcoNac.EOF Then
            If rsDesBcoNac!nEstado = 1 Then ' Si ya se generó el archivo para el banco de la nación
                CargaDatos = False
                MsgBox "El Credito es para desembolso en Banco de la Nación y ya forma parte del archivo que será enviado al Banco de la Nación", vbExclamation, "Aviso"
                Exit Function
            End If
        End If
        '*****************************
        CargaDatos = True
        'RECO20160711 ERS002-2016 ****************************************
        Dim oCredito As New COMDCredito.DCOMCredito
        'Call oCredito.RegistraAutorizacionesRequeridas(Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss"), gsCodUser, gsCodAge, ActxCta.NroCuenta) 'FRHU 20160829 Se quito.
        If oCredito.verificarExisteAutorizaciones(ActxCta.NroCuenta) Then
            MsgBox "El crédito tiene una autorización/exoneración pendiente", vbInformation, "Alerta"
            Call frmCredNewNivAutorizaVer.Consultar(ActxCta.NroCuenta)
            CargaDatos = False
            Exit Function
        End If
        Set oCredito = Nothing
        'RECO FIN*********************************************************
        
        'JUEZ 20160510 ****************************************************************
        Dim oDCredAmp As COMDCredito.DCOMCredito
        Dim RsAutAmp As ADODB.Recordset
        Set oDCredAmp = New COMDCredito.DCOMCredito
            Set RsAutAmp = oDCredAmp.RecuperaSolicitudAutorizacionAmpliacion(psCtaCod)
            If Not (RsAutAmp.EOF And RsAutAmp.BOF) Then
            If RsAutAmp!nEstado = 0 Then
                MsgBox "El crédito fue solicitado requiriendo Autorización de Ampliación Excepcional, la cual aun está pendiente en Riesgos.", vbInformation, "Aviso"
                CargaDatos = False
                fbSolicitudAutAmp = True
                Exit Function
            ElseIf RsAutAmp!nEstado = 2 Then
                MsgBox "El crédito no puede ser sugerido, pues su solicitud de Autorización de Ampliación Excepcional fue observada por la Gerencia de Riesgos. El crédito deberá ser solicitado como refinanciado.", vbInformation, "Aviso"
                CargaDatos = False
                fbSolicitudAutAmp = True
                Exit Function
            End If
        End If
        Set oDCredAmp = Nothing
        'END JUEZ *********************************************************************
        'JUEZ 20121217 ************************************************************************************
        bAprobNivPendiente = False
        If gnAgenciaCredEval = 1 Then
            Dim oDNiv As COMDCredito.DCOMNivelAprobacion
            Set oDNiv = New COMDCredito.DCOMNivelAprobacion
            'RECO20160415 *****************************************
            'Dim oCliPre As New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
            'Set oCliPre = New COMNCredito.NCOMCredito  'COMENTADO POR ARLO 20170722
            Dim bClientPref As Boolean
            
            'bClientPref = oCliPre.ValidarClientePreferencial(LblCodCli.Caption)
            'bClientPref = oCliPre.ValidarClientePreferencial(RDatos!cPersCod) 'COMENTADO POR ARLO 20170722
            bClientPref = False 'ARLO 20170722
            'RECO FIN *********************************************
                If oDNiv.ExisteAprobacionCredNivelesPendientes(psCtaCod, IIf(bClientPref = True, 2, 1)) Then 'RECO20160515
                'If oDNiv.ExisteAprobacionCredNivelesPendientes(psCtaCod) Then 'RECO20160515
                    bAprobNivPendiente = True
                    If MsgBox("El crédito no puede ser aprobado porque aun no ha recibido los V°B° de todos los niveles." & vbCrLf & _
                                  "                 ¿Desea revisar el historial de Aprobaciones del crédito?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                        frmCredNewNivAprHist.InicioCredito (psCtaCod)
                    End If
                    CargaDatos = False
                    Exit Function
                End If
            Set oDNiv = Nothing
        End If
        'END JUEZ *****************************************************************************************
        'EJVG20150616 ***
        If Not rsPolizaSolicitud.EOF Then
            MsgBox "El crédito no puede ser aprobado porque tiene pendiente una Solicitud de Póliza del día " & Format(rsPolizaSolicitud!dSolicitud, gsFormatoFechaHoraViewAMPM) & "." & _
                Chr(13) & Chr(13) & "Por favor coordinar con el Supervisor de Operaciones para la emisión de Póliza", vbInformation, "Aviso"
            RSClose rsPolizaSolicitud
            CargaDatos = False
            fbSolicitudPoliza = True
            Exit Function
        End If
        RSClose rsPolizaSolicitud
        'END EJVG *******
        
        txtMontoMivivienda.Text = Format(RDatos!nMontoMiVivienda, "#0.00") 'ALPA20140611
        nEstadoPoli = RDatos!nEstadoPoliza '***PEAC 20080204
        nUltNroCalend = RDatos!nMaxCalend
        vsLineaCredito = RDatos!cLineaCredCod
        vnTipoGracia = RDatos!nTipoGracia
        'ALPA20141127******************
        If RDatos!cTpoCredCod = "853" Then
            lnCSP = RDatos!nCuotaPolizaMivivienda
        Else
            lnCSP = -1
        End If
        '*************************************
        
        'Set oCal = New COMDCredito.DCOMCalendario
        'Set R2 = oCal.RecuperaCalendarioPagos(psCtaCod)
        'Set oCal = Nothing
        'ReDim MatGracia(R2.RecordCount)
        ReDim MatGracia(RCal.RecordCount)
        Do While Not RCal.EOF
            MatGracia(RCal.Bookmark - 1) = Format(RCal!nIntGracia, "#0.00")
            RCal.MoveNext
        Loop
        'R2.Close
                
        'Set oNCredito = New COMNCredito.NCOMCredito
        'bRefinanc = oNCredito.EsRefinanciado(ActxCta.NroCuenta)
        'bSustiDeudor = oNCredito.EsSustiDeudor(ActxCta.NroCuenta)

        If bRefinanc Then
           If bSustiDeudor Then
              LblCredNormalRefinan = "CREDITO SUSTITUTOR DE DEUDOR"
           Else
              LblCredNormalRefinan = "CREDITO REFINANCIADO"
           End If
        Else
           LblCredNormalRefinan = "CREDITO NORMAL"
        End If
        'Set oNCredito = Nothing
        
        If RDatos!cTipoGasto = "V" Then
            OptTipoGasto(1).value = True
        Else
            OptTipoGasto(0).value = True
        End If
        
        If bRefinanc Then
            CmdGastos.Enabled = False
        End If
        
        If IsNull(RDatos!bRefCapInt) Then
            bCapitalInt = False
        Else
            bCapitalInt = IIf(RDatos!bRefCapInt, True, False)
        End If
        ChkCuotaCom.value = IIf(IsNull(RDatos!bCuotaCom), 0, RDatos!bCuotaCom)
        ChkMiViv.value = IIf(IsNull(RDatos!bMiVivienda), 0, RDatos!bMiVivienda)
        LblCodCli.Caption = RDatos!cPersCod
        LblNomCli.Caption = RDatos!cPersNombre
        LblDireccion.Caption = RDatos!cPersDireccDomicilio
        lnColocCondicion = RDatos!nColocCondicion 'ALPA 20141230
        'RECO20140226 ERS174-2013*************************
        Set rsExoAut = New ADODB.Recordset
        Set rsExoAut = oDNivExo.ObtieneDatosExoAutCred(Me.ActxCta.NroCuenta)
        If Not (rsExoAut.EOF And rsExoAut.BOF) Then
            ChkReglamento.value = 1
        Else
            ChkReglamento.value = 0
        End If
        'RECO FIN*****************************************
        'madm 201100303 ***********************************
        lsTpoIDNroDNI = RDatos!Dni
        lsTpoIDNroRUC = RDatos!Ruc
        '*************************************************
        
        'JUEZ 20120907 ****************************************************
        Set oNCredito = New COMNCredito.NCOMCredito
        nAgenciaCredEval = IIf(oNCredito.ObtieneAgenciaCredEval(gsCodAge) = 1 And Not rsCredEval.EOF, 1, 0)
        Set oNCredito = Nothing
        If nAgenciaCredEval = 0 Then
            'madm 20100513 --------------------------------------------------------------------------------------
            'If Mid(ActxCta.NroCuenta, 6, 3) <> "302" Then
            If Mid(ActxCta.NroCuenta, 6, 3) <> "703" Then
                LblCodFte.Caption = RDatos!cPersFIPersCod
                LblDescripFte.Caption = RDatos!cFteIngreso
            
            ' CMACICA_CSTS - 12/11/2003 -------------------------------------------------------------------------
                nTipoFteIngreso = RDatos!nPersTipFte
            End If
            'end madm
        End If
        'END JUEZ *********************************************************
        lnColocDestino = RDatos!nColocDestino 'ALPA20141215***************
        Set oPersona = Nothing
        Set oPersona = New UPersona_Cli
        Call oPersona.RecuperaPersona(Trim(LblCodCli.Caption), ActxCta.NroCuenta)
        If oPersona.PersCodigo = "" Then
            MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        End If
        
        Call LimpiaFlex(FEGarantCred)
        'Set RGar = oCredito.RecuperaGarantiasCredito(ActxCta.NroCuenta)
        Do While Not RGar.EOF
            FEGarantCred.AdicionaFila
            FEGarantCred.RowHeight(RGar.Bookmark) = 280
            FEGarantCred.TextMatrix(RGar.Bookmark, 1) = RGar!cTpoGarDescripcion
            FEGarantCred.TextMatrix(RGar.Bookmark, 2) = Format(RGar!nGravado, "#,#0.00")
            FEGarantCred.TextMatrix(RGar.Bookmark, 3) = Format(RGar!nTasacion, "#,#0.00")
            'FEGarantCred.TextMatrix(RGar.Bookmark, 4) = Format(RGar!nRealizacion, "#,#0.00")
            FEGarantCred.TextMatrix(RGar.Bookmark, 4) = Format(RGar!nValorGarantia, "#,#0.00") 'EJVG20150518
            'FEGarantCred.TextMatrix(RGar.Bookmark, 5) = Format(RGar!nPorGravar, "#,#0.00")
            FEGarantCred.TextMatrix(RGar.Bookmark, 5) = Format(RGar!nDisponible, "#,#0.00") 'EJVG20150518
            FEGarantCred.TextMatrix(RGar.Bookmark, 6) = Trim(RGar!cPersNombre)
            FEGarantCred.TextMatrix(RGar.Bookmark, 7) = Trim(RGar!cNroDoc)
            FEGarantCred.TextMatrix(RGar.Bookmark, 8) = Trim(RGar!cTpoDoc)
            FEGarantCred.TextMatrix(RGar.Bookmark, 9) = Trim(RGar!cNumGarant)
            FEGarantCred.TextMatrix(RGar.Bookmark, 10) = RGar!cVerificaLegal
            FEGarantCred.TextMatrix(RGar.Bookmark, 11) = RGar!cVerificaPoliza
            
            'If FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Conforme" Then
             If nPolizaBF = 0 Then
                If RGar!nPolizaBF = 1 Then
                    nPolizaBF = 1
                    dFechaPBF = Format(RGar!dFechaPBF)
                Else
                    nPolizaBF = 0
                    dFechaPBF = "01/01/1900"
                End If
             End If
            'End If
            
            'END MADM
            RGar.MoveNext
        Loop

        LblTotGarantiasGrav = Format(nMontoGravado, "#0.00")
        'ALPA20140219************************
        ckPolizaBF.value = nPolizaBF
        txtFechaPBF.Text = Format(dFechaPBF, "DD/MM/YYYY")
          
        'Tipo de Desembolso
        If RDatos!nTipoDesembolso = gColocTiposDesembolsoTotal Then
            OptDesemb(0).value = True
        End If
        If RDatos!nTipoDesembolso = gColocTiposDesembolsoParcial Then
            OptDesemb(1).value = True
        End If
        
        lblCondicion.Caption = RDatos!cCondicionCred
        
        ' CMACICA_CSTS - 05/11/2003 -----------------------------------------------------------------------
        LblCondicionOtra.Caption = RDatos!cCondicionOtraCred
        ' -------------------------------------------------------------------------------------------------
        'LblCredTipo.Caption = RDatos!cTipoCredDescrip
        LblMontoSol.Caption = Format(RDatos!nMontoSol, "#0.00")
        LblCuotasSol.Caption = RDatos!nCuotasSol
        LblPlazoSol.Caption = RDatos!nPlazoSol
        LblMontoSug.Caption = Format(RDatos!nMonto, "#0.00")
        LblCuotaSug.Caption = Format(RDatos!nCuotaSug, "#0.00")
        LblCuotasSug.Caption = RDatos!nCuotas
        LblPlazoSug.Caption = RDatos!nPlazo
        LblGraciaSug.Caption = RDatos!nPeriodoGracia
        LblDiaFijoSug.Caption = RDatos!nPeriodoFechaFija
        LblLineaCredSug.Caption = RDatos!cLineaCred
        lblNumConCer.Caption = IIf(IsNull(RDatos!nNumConCer), "", RDatos!nNumConCer)
        TxtTasaGracia.Text = Format(RDatos!nTasaGracia, "#0.00")
        'Add By Gitu 11-04-2009
        txtTasInt.Text = RDatos!nTasaInteres
        '**********************
        'sLineaTmp = RDatos!cLineaCred
        Call CargarLineaForAprobacion(rsLinea)
        
        'ARCV 24-02-2007
        nTasaGrac = Format(RDatos!nTasaComp, "#0.0000")
        nTasaComp = Format(RDatos!nTasaComp, "#0.0000")
        TxtTasaGracia.Text = Format(RDatos!nTasaComp, "#0.0000")
        '----------
        LblTasaSug.Caption = Format(RDatos!nTasaComp, "#0.0000")
        LblTasaGraciaSug.Caption = Format(IIf(IsNull(RDatos!nTasaGracia), 0, RDatos!nTasaGracia), "#0.0000")
        LblAnalista.Caption = PstaNombre(RDatos!cAnalista)
        LblIntComp.Caption = Format(IIf(IsNull(RDatos!nTasaIntComp), 0, RDatos!nTasaIntComp), "#0.0000")
        LblIntMor.Caption = Format(IIf(IsNull(RDatos!nTasaIntMora), 0, RDatos!nTasaIntMora), "#0.0000")
        TxtMontoApr.Text = Format(IIf(IsNull(RDatos!nMonto), 0, RDatos!nMonto), "#0.00")
        TxtDiaFijo.Text = Format(IIf(IsNull(RDatos!nPeriodoFechaFija), 0, RDatos!nPeriodoFechaFija), "#00")
        
        If RDatos!nColocCalendCod = gColocCalendCodCL Then
            LblTipoCuotaSug.Caption = "CUOTA LIBRE"
            LblTipoPeriodoSug.Caption = "NINGUNO"
            LblTipoGraciaSug.Caption = "NINGUNO"
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCC Or RDatos!nColocCalendCod = gColocCalendCodFFCCPG Or RDatos!nColocCalendCod = gColocCalendCodPFCC Or RDatos!nColocCalendCod = gColocCalendCodPFCCPG Then
            LblTipoCuotaSug.Caption = "CUOTA CRECIENTE"
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCF Or RDatos!nColocCalendCod = gColocCalendCodFFCFPG Or RDatos!nColocCalendCod = gColocCalendCodPFCF Or RDatos!nColocCalendCod = gColocCalendCodPFCFPG Then
            LblTipoCuotaSug.Caption = "CUOTA FIJA"
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCD Or RDatos!nColocCalendCod = gColocCalendCodFFCDPG Or RDatos!nColocCalendCod = gColocCalendCodPFCD Or RDatos!nColocCalendCod = gColocCalendCodPFCDPG Then
            LblTipoCuotaSug.Caption = "CUOTA DECRECIENTE"
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCC Or RDatos!nColocCalendCod = gColocCalendCodFFCCPG Or RDatos!nColocCalendCod = gColocCalendCodFFCD Or RDatos!nColocCalendCod = gColocCalendCodFFCCPG _
          Or RDatos!nColocCalendCod = gColocCalendCodFFCCPG Or RDatos!nColocCalendCod = gColocCalendCodFFCD Or RDatos!nColocCalendCod = gColocCalendCodFFCDPG Or RDatos!nColocCalendCod = gColocCalendCodFFCF Or RDatos!nColocCalendCod = gColocCalendCodFFCFPG Then
            LblTipoPeriodoSug.Caption = "FECHA FIJA"
        End If
        If RDatos!nColocCalendCod = gColocCalendCodPFCC Or RDatos!nColocCalendCod = gColocCalendCodPFCCPG Or RDatos!nColocCalendCod = gColocCalendCodPFCD Or RDatos!nColocCalendCod = gColocCalendCodPFCCPG _
          Or RDatos!nColocCalendCod = gColocCalendCodPFCCPG Or RDatos!nColocCalendCod = gColocCalendCodPFCD Or RDatos!nColocCalendCod = gColocCalendCodPFCDPG Or RDatos!nColocCalendCod = gColocCalendCodPFCF Or RDatos!nColocCalendCod = gColocCalendCodPFCFPG Then
            LblTipoPeriodoSug.Caption = "PERIODO FIJO"
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCCPG Or RDatos!nColocCalendCod = gColocCalendCodFFCDPG Or RDatos!nColocCalendCod = gColocCalendCodFFCFPG Or RDatos!nColocCalendCod = gColocCalendCodPFCCPG _
           Or RDatos!nColocCalendCod = gColocCalendCodPFCDPG Or RDatos!nColocCalendCod = gColocCalendCodPFCFPG Then
           Select Case RDatos!nTipoGracia
            Case gColocTiposGraciaConfigurable
                LblTipoGraciaSug.Caption = "CONFIGURABLE"
            Case gColocTiposGraciaExonerada
                LblTipoGraciaSug.Caption = "EXONERADA"
            Case gColocTiposGraciaPrimeraCuota
                LblTipoGraciaSug.Caption = "PRIMERA CUOTA"
            Case gColocTiposGraciaProrateada
                LblTipoGraciaSug.Caption = "PRORATEADA"
            Case gColocTiposGraciaUltimaCuota
                LblTipoGraciaSug.Caption = "ULTIMA CUOTA"
            'Nuevos Tipos de Gracia
            Case gColocTiposGraciaCapitalizada
                LblTipoGraciaSug.Caption = "CAPITALIZADA"
            Case gColocTiposGraciaEnCuotas
                LblTipoGraciaSug.Caption = "EN CUOTAS"
           End Select
        Else
            LblTipoGraciaSug.Caption = "NINGUNO"
        End If
        
        'Tipo de Calendario
        If RDatos!nCalendDinamico = 0 Then
            LblTipoCalendSug.Caption = "FIJO"
        Else
            LblTipoCalendSug.Caption = "DINAMICO"
        End If
        'Destino del Credito
        lsTpoProdCod = RDatos!cTpoProdCod 'FRHU20140221 -RQ14005
        CmbDestino.ListIndex = IndiceListaCombo(CmbDestino, RDatos!nColocDestino)
        'Condicion del Credito
        nCondCred = RDatos!nColocCondicion 'MAVM 20110613
        CmbCondicion.ListIndex = IndiceListaCombo(CmbCondicion, RDatos!nColocCondicion)
        'Campaña del Credito Add GITU 24-03-2009
        'CmbCampana.ListIndex = IndiceListaCombo(CmbCampana, RDatos!IdCampana)
        CmbCampana.ListIndex = IndiceListaCombo(CmbCampana, IIf(IsNull(RDatos!idCampana), 0, RDatos!idCampana))
        
        ' CMACICA_CSTS - 05/11/2003 -----------------------------------------------------------------------
        'Condicion 2 del Credito
        CmbCondicionOtra.ListIndex = IndiceListaCombo(CmbCondicionOtra, RDatos!nColocCondicion2)
        ' -------------------------------------------------------------------------------------------------
        
        If gnAgenciaCredEval = 1 Then CmbApoderado.ListIndex = IndiceListaCombo(CmbApoderado, RDatos!cCodApoderado) 'JUEZ 20121218
        
        If nIndtempo = -99 Then
            CmbApoderado.ListIndex = -1
            CmbApoderado.Enabled = True
        Else
            
        End If
        
        SpnCuotasApr.valor = RDatos!nCuotas
        TxtFecDesemb.Text = Format(gdFecSis, "dd/mm/yyyy")
        SpnPlazo.valor = RDatos!nPlazo
        TxtDiasGracia.Text = IIf(IsNull(RDatos!nPeriodoGracia), 0, RDatos!nPeriodoGracia)
        
        'MAVM 25102010 ***
        If SpnPlazo.valor <> "0" Then
            txtFechaFija.Text = CDate(txtFechaFija.Text + CDate(TxtDiasGracia.Text))
        End If
        '***
        Set oNCredito = New COMNCredito.NCOMCredito
        'Asignar Tipo de Cuota
        If RDatos!nColocCalendCod = gColocCalendCodCL Then
            OptTCuota(3).value = True
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCF Or RDatos!nColocCalendCod = gColocCalendCodFFCFPG _
            Or RDatos!nColocCalendCod = gColocCalendCodPFCF Or RDatos!nColocCalendCod = gColocCalendCodPFCFPG Then
            OptTCuota(0).value = True
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCC Or RDatos!nColocCalendCod = gColocCalendCodFFCCPG _
            Or RDatos!nColocCalendCod = gColocCalendCodPFCC Or RDatos!nColocCalendCod = gColocCalendCodPFCCPG Then
            OptTCuota(1).value = True
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCD Or RDatos!nColocCalendCod = gColocCalendCodFFCDPG _
            Or RDatos!nColocCalendCod = gColocCalendCodPFCD Or RDatos!nColocCalendCod = gColocCalendCodPFCDPG Then
            OptTCuota(2).value = True
        End If
        
        'Tipo de Periodo
        If RDatos!nColocCalendCod = gColocCalendCodPFCC Or RDatos!nColocCalendCod = gColocCalendCodPFCCPG _
            Or RDatos!nColocCalendCod = gColocCalendCodPFCD Or RDatos!nColocCalendCod = gColocCalendCodPFCDPG _
            Or RDatos!nColocCalendCod = gColocCalendCodPFCF Or RDatos!nColocCalendCod = gColocCalendCodPFCFPG Then
            OptTPer(0).value = True
'***********COMENTADO POR APRI20180404 ERS036-2017(chkCalFijo no se usaba)**************
'            If chkCalFijo.value = 1 Then
'                txtFechaFija.Text = CDate(TxtFecDesemb.Text) + RDatos!nPeriodoGracia + RDatos!Plazo
'            End If
'************END APRI*******************************************
        End If
        If RDatos!nColocCalendCod = gColocCalendCodFFCC Or RDatos!nColocCalendCod = gColocCalendCodFFCCPG _
            Or RDatos!nColocCalendCod = gColocCalendCodFFCD Or RDatos!nColocCalendCod = gColocCalendCodFFCDPG _
            Or RDatos!nColocCalendCod = gColocCalendCodFFCF Or RDatos!nColocCalendCod = gColocCalendCodFFCFPG Then
            OptTPer(1).value = True
            TxtDiaFijo.Text = Format(IIf(IsNull(RDatos!nPeriodoFechaFija), "00", RDatos!nPeriodoFechaFija), "00")
'***********COMENTADO POR APRI20180404 ERS036-2017(chkCalFijo no se usaba)**************
'            If chkCalFijo.value = 1 Then
'                txtFechaFija.Text = CDate(TxtFecDesemb.Text) + RDatos!nPeriodoGracia + 30
'            End If
'************END APRI*******************************************
            ChkProxMes.value = IIf(IIf(IsNull(RDatos!nProxMes), False, RDatos!nProxMes), 1, 0)
        End If
        
        'MAVM 25102010 ***
        TxtDiasGracia.Text = IIf(IsNull(RDatos!nPeriodoGracia), 0, RDatos!nPeriodoGracia)
        If TxtDiasGracia.Text <> "0" Then
            TxtTasaGracia.Text = Format(RDatos!nTasaComp, "#0.0000")
        End If
        TxtDiasGracia.Enabled = False
        If (TxtDiaFijo.Text <> "00") And (TxtDiaFijo.Text <> "0") Then
            ChkProxMes.Enabled = False
            If ChkProxMes.value = 0 Then
                txtFechaFija.Text = CDate(TxtDiaFijo.Text & "/" & Month(gdFecSis) & "/" & Year(gdFecSis)) + CDate(TxtDiasGracia.Text)
            Else
                If TxtDiasGracia.Text > "0" Then
                    txtFechaFija.Text = Mid(CDate(TxtFecDesemb.Text + CDate(30) + TxtDiasGracia.Text), 1, 10)
                Else
                    If Not Trim(ValidaFecha((TxtDiaFijo.Text & Mid(DateAdd("m", 1, TxtFecDesemb.Text), 3, 8)))) = "" Then
                        MsgBox Trim(ValidaFecha(txtFechaFija.Text)), vbInformation, "Aviso"
                    Else
                        txtFechaFija.Text = CDate(TxtDiaFijo.Text & Mid(DateAdd("m", 1, TxtFecDesemb.Text), 3, 8))
                    End If
                End If
            End If
        End If
        '***
        
        'Tipo de Desembolso
        '----
        
        'Tipo Calendario
        If RDatos!nCalendDinamico = 0 Then
            OptTipoCalend(0).value = True
        End If
        If RDatos!nCalendDinamico = 1 Then
            OptTipoCalend(1).value = True
        End If
        
        'Tipo de Desembolso
        If RDatos!nTipoDesembolso = gColocTiposDesembolsoTotal Then
            LblTipoDesembSug.Caption = "TOTAL"
            MatDesemb = ""
        Else
        '    Set oCalend = New COMDCredito.DCOMCalendario
            LblTipoDesembSug.Caption = "PARCIAL"
        '    Set RDes = oCalend.RecuperaCalendarioDesemb(ActxCta.NroCuenta)
        '    Set oCalend = Nothing
            If RDes.RecordCount > 0 Then
                ReDim MatDesemb(RDes.RecordCount, 2)
            End If
            Do While Not RDes.EOF
                MatDesemb(RDes.Bookmark - 1, 0) = Format(RDes!dVenc, "dd/mm/yyyy")
                MatDesemb(RDes.Bookmark - 1, 1) = Format(RDes!nCapital, "#0.00")
                RDes.MoveNext
            Loop
            'RDes.Close
        End If
        
        'RDatos.Close
        '**** CUSCO *****
        If RDatos!bMiVivienda = 1 Then
            ChkMiViv.Visible = True
            'fraMiViv.Visible = True 'LUCV20180601, Comentó según ERS022-2018
        Else
            ChkMiViv.Visible = False
            'fraMiViv.Visible = False 'LUCV20180601, Comentó según ERS022-2018
        End If
        
        LimpiaFlex FERelacCred
        'Set R = oCredito.RecuperaRelacPers(ActxCta.NroCuenta)
        Do While Not RRel.EOF
            FERelacCred.AdicionaFila
            FERelacCred.TextMatrix(RRel.Bookmark, 1) = RRel!cPersNombre
            FERelacCred.TextMatrix(RRel.Bookmark, 2) = RRel!cConsDescripcion
            'APRI20180404 ERS036-2017
            Ini = Ini + 1
            'ReDim Preserve Clientes(ini)
            Clientes(counter, 0) = RRel!cPersCod
            Clientes(counter, 1) = RRel!nConsValor
            counter = counter + 1
            'END APRI
            RRel.MoveNext
        Loop
        'R.Close
        'Set R = Nothing
        
        
        If bRefinanc Then
            TxtMontoApr.Enabled = False
            'cmdLineas.Enabled = False 'DAOR 20070409 'madm 20110413
        End If
        '****************
        
        'Para manejar 2 Dias Fijos
        TxtDiaFijo2.Text = Format(IIf(IsNull(RDatos!nPeriodoFechaFija2), 0, RDatos!nPeriodoFechaFija2), "#00")
        'Para los nuevos Tipos de Gracia
        If vnTipoGracia = gColocTiposGraciaCapitalizada Then
            optTipoGracia(0).value = True
            optTipoGracia(0).Enabled = True 'MAVM 20130507
            Call optTipoGracia_Click(0)
        End If
        'chkIncremenK.value = RDatos!bIncremGraciaCap
        If vnTipoGracia = gColocTiposGraciaEnCuotas Then
            optTipoGracia(1).value = True
        End If
        '**************************
        '06-05-2005
        bBuscarLineas = True
        '*********************
        
        'DAOR 20070202, datos para desembolso en banco de la nación**
        If Not rsDesBcoNac.EOF Then
            chkDesembBcoNac.value = 1
            frmCredDesembBcoNac.txtCodigo.Text = rsDesBcoNac!cClave
            frmCredDesembBcoNac.cmbAgencia.ListIndex = IndiceListaCombo(frmCredDesembBcoNac.cmbAgencia, rsDesBcoNac!cCodAgeBcoNac)
        End If
        '************************************************************
        
        'JUEZ 20121218 **********************************************
        If gnAgenciaCredEval = 1 Then
            ChkBloqueo.value = IIf(RDatos!bBloqueo = True, 1, 0)
            ChkReglamento.value = RDatos!nExoneraReg
            ChkExoneracion.value = RDatos!nExoSeguroDes
            actxCtaAho.NroCuenta = RDatos!cCtaCodAho
            chkExoneraMYPE.value = RDatos!nExoSegMYPE 'APRI20180821 ERS061-2018
        End If
        'END JUEZ ***************************************************
        
        '**DAOR 20080109 ********************************************
        fnPersPersoneria = RDatos!nPersPersoneria
        fnPersTipoDoc = RDatos!nPersIdTpo
        fsPersNumeDoc = RDatos!cPersIDnro
        '************************************************************
        'ALPA 20100605************************************************
        LblProducto.Caption = RDatos!cTipoProdDescrip
        lblSubProducto.Caption = RDatos!cSTipoProdDescrip
        lblTipoCred.Caption = RDatos!cTipoCredDescrip
        lblSubTipoCred.Caption = RDatos!cSTipoCredDescrip
        lsTpoProdCod = RDatos!cTpoProdCod
        lsTpoCredCod = RDatos!cTpoCredCod
        
        'JOEP20190131 CP
        lsCategoria = ""
        lsTipoCredito = ""
        lsCategoria = RDatos!Categoria
        lsTipoCredito = RDatos!TipoCredito
        'JOEP20190131 CP
        
        If Mid(RDatos!cTpoCredCod, 1, 2) = Mid(gColCredCorpo, 1, 2) Then
            Label123.Visible = True
            lblInstitucionFinanciera.Visible = True
            lblInstitucionFinanciera.Caption = RDatos!cConsInstCorpDesc
        Else
            Label123.Visible = False
            lblInstitucionFinanciera.Visible = False
        End If
        '*************************************************************
        '***BRGO 20111104 ***************************************
        If RDatos!cTpoProdCod = "517" Then
            Dim rsRel As ADODB.Recordset
            Dim clsGen As COMDConstSistema.DCOMGeneral
            Set clsGen = New COMDConstSistema.DCOMGeneral
            
            Me.fraEmpVinculados.Visible = True
            Set rsRel = clsGen.GetConstante(gColRelacPersInfoGas)
            grdEmpVinculados.CargaCombo rsRel
            Set rsRel = Nothing
            Set clsGen = Nothing
            If Not rsRelEmp.EOF And Not rsRelEmp.BOF Then
                Set grdEmpVinculados.Recordset = rsRelEmp
                For i = 1 To Me.grdEmpVinculados.rows - 1
                    Me.grdEmpVinculados.TextMatrix(i, 4) = Format(Me.grdEmpVinculados.TextMatrix(i, 4), "#,000.00")
                Next
            End If
        Else
            Me.fraEmpVinculados.Visible = False
        End If
        '***END BRGO *****************************************************
        fbLimpiaExoSegDes = False 'WIOR 20130902
        'WIOR 20130419*******************************************************
        'Dim oDCredito As COMDCredito.DCOMCredito
        Dim rsSegDes  As ADODB.Recordset
        Dim CantExo As Integer
        
        Set oDCredito = New COMDCredito.DCOMCredito
        CantExo = 0
        Set rsSegDes = oDCredito.ListaPersonaExonerasSegDesg(psCtaCod, "%")
        fbValidaExoSegDes = True
    
        If Not (rsSegDes.EOF And rsSegDes.BOF) Then
            For i = 1 To rsSegDes.RecordCount
                If CBool(rsSegDes!bExonera) Then
                    CantExo = CantExo + 1
                End If
                rsSegDes.MoveNext
            Next i
     
        End If
        
        If CantExo > 0 Then
            If rsSegDes.RecordCount = CantExo Then
                If ChkExoneracion.value = 1 Then
                    fbValidaExoSegDes = False
                End If
                ChkExoneracion.value = 1
            ElseIf (rsSegDes.RecordCount - CantExo) > 0 Then
                If ChkExoneracion.value = 2 Then
                    fbValidaExoSegDes = False
                End If
                ChkExoneracion.value = 2
            End If
        Else
            If ChkExoneracion.value = 0 Then
                fbValidaExoSegDes = False
            End If
            ChkExoneracion.value = 0
        End If
        
        Set oDCredito = Nothing
        Set rsSegDes = Nothing
        'WIOR FIN ***********************************************************
        ''FRHU 20140521 ERS015-2014 Observacion
        'Dim oAmpliado As New COMDCredito.DCOMAmpliacion
        'If oAmpliado.ValidaCreditoaAmpliar(ActxCta.NroCuenta) Then
        If fbEsAmpliado Then 'EJVG20160531
            MatCredVig = frmCredVigentes.Inicio(LblCodCli.Caption, LblNomCli.Caption, ActxCta.NroCuenta, MatCredVig)
        End If
        'Set oAmpliado = Nothing
        ''FIN FRHU 20140521 Observacion
        'ALPA 20150115**********************************
        Call CargarDatosProductoCrediticio
        txtTasInt.Text = RDatos!nTasaInteres
        Call MostrarLineas
        
        Dim oCredExo As COMNCredito.NCOMCredito
        Set oCredExo = New COMNCredito.NCOMCredito
            If oCredExo.ValidaExoneracion(psCtaCod, "TIP0009") Then
                txtTasInt.Text = RDatos!nTasaExononeracion
                chkTasa.value = 1
                txtInteresTasa.Text = RDatos!nTasaInteres
                
            Else
                txtTasInt.Text = RDatos!nTasaInteres
                chkTasa.value = 0
                txtInteresTasa.Text = 0
                
            End If
        Set oCredExo = Nothing
        ckcPreferencial.value = 0 'IIf(RDatos!bPreferencial = 1, 1, 0)
        lnExoneracionTasa = 0 'IIf(RDatos!bExononeracionTasa = 1, 1, 0)
        lnCliPreferencial = 0 'IIf(RDatos!bPreferencial = 1, 1, 0)
        
        lnLogicoBuscarDatos = 1
        '**********************************************
        'WIOR 20151223 ***
        fbMIVIVIENDA = oNCredito.EsCredMIVIVENDA(lsTpoProdCod, lsTpoCredCod, 2)
        TxtMontoApr.Enabled = True
        FraTpoCalend.Enabled = False
        ChkMiViv.Enabled = False
        ChkMiViv.value = 0
        cmdMIVIVIENDA.Enabled = False

        If fbMIVIVIENDA Then
            TxtMontoApr.Enabled = False
            FraTpoCalend.Enabled = True
            ChkMiViv.Enabled = False
            ChkMiViv.value = 1
            cmdMIVIVIENDA.Enabled = True
            
            Set oDCredito = New COMDCredito.DCOMCredito
            Dim rsMiViv As ADODB.Recordset
            Set rsMiViv = oDCredito.ObtenerDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstSug)
            
            If Not (rsMiViv.EOF And rsMiViv.BOF) Then
                ReDim fArrMIVIVIENDA(11)
                fArrMIVIVIENDA(0) = CDbl(rsMiViv!nMontoVivienda)
                fArrMIVIVIENDA(1) = CDbl(rsMiViv!nCuotaInicial)
                fArrMIVIVIENDA(2) = CDbl(rsMiViv!nBonoOtorgado)
                fArrMIVIVIENDA(3) = CDbl(rsMiViv!nMOntoCred)
                fArrMIVIVIENDA(4) = CDbl(rsMiViv!nUIT)
                fArrMIVIVIENDA(5) = CLng(rsMiViv!nDesde)
                fArrMIVIVIENDA(6) = CLng(rsMiViv!nHasta)
                fArrMIVIVIENDA(7) = CDbl(rsMiViv!nBono)
                fArrMIVIVIENDA(8) = CDbl(rsMiViv!nMinCredUIT)
                fArrMIVIVIENDA(9) = 1
                fArrMIVIVIENDA(10) = CInt(rsMiViv!nPeriodoPerdBono)
                fArrMIVIVIENDA(11) = CDbl(rsMiViv!nGastoCierre) 'CTI5-ERS0012020
            End If
            
            Set oDCredito = Nothing
        End If
        'WIOR FIN ********
        Call ObtenerTasaSeguroDesg 'WIOR 20160224
    Else
        MsgBox "El Crédito no esta en estado Sugerido o ya fue Aprobado", vbInformation, "AVISO!"
        CargaDatos = False
        lnLogicoBuscarDatos = 0
    End If
    'ALPA 20111212****************************************
    Else
        MsgBox "El Crédito Leasing debe ser aprobado en el SAF", vbInformation, "AVISO!"
        CargaDatos = False
        lnLogicoBuscarDatos = 0
    End If
    '*****************************************************
    txtBuscarLinea.Enabled = False 'LUCV20180601, Comentó según ERS022-2018
    
    Exit Function

ErrorCargaDatos:
        MsgBox Err.Description, vbCritical, "Aviso"
End Function
'ALPA 20150115***************************************************
Private Sub CargarDatosProductoCrediticio()
Dim sCodigo As String
Dim sCtaCodOrigen As String
Dim oLineas As COMDCredito.DCOMLineaCredito
Set RLinea = New ADODB.Recordset
sCodigo = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
Set oLineas = New COMDCredito.DCOMLineaCredito
txtBuscarLinea.Text = ""
lblLineaDesc.Caption = ""
If Len(ActxCta.NroCuenta) = 18 Then 'WIOR 20150718
    Set RLinea = oLineas.RecuperaLineadeCreditoProductoCrediticio(lsTpoProdCod, Trim(Right(IIf(CmbCampana.Text = "", 0, CmbCampana.Text), 10)), Trim(Right((txtBuscarLinea.psDescripcion), 15)), sCodigo, lblLineaDesc, Mid(ActxCta.NroCuenta, 9, 1), CCur(IIf((TxtMontoApr.Text) = "", 0, TxtMontoApr.Text)), IIf(ckcPreferencial.value = 1, 1, 0))
    Set oLineas = Nothing
           If RLinea.RecordCount > 0 Then
              If txtBuscarLinea.Text = "" Then
                txtBuscarLinea.Text = "XXX"
              End If
              Call CargaDatosLinea
              If txtBuscarLinea.Text = "XXX" Then
                txtBuscarLinea.Text = ""
              End If
           Else
                lnTasaInicial = 0
                lnTasaFinal = 0
           End If
    Call MostrarLineas
End If 'WIOR 20150718
End Sub
Private Sub CargaDatosLinea()
ReDim MatCalend(0, 0)
ReDim MatrizCal(0, 0)
    
    If Trim(txtBuscarLinea.Text) = "" Then
        Exit Sub
    End If
    'ALPA 20150313**********************
    If RLinea.BOF Or RLinea.EOF Then
        lnTasaInicial = 0#
        lnTasaFinal = 0#
        Exit Sub
    End If
    '***********************************

    lnTasaInicial = RLinea!nTasaIni
    lnTasaFinal = RLinea!nTasaFin
    If RLinea!nTasaIni <> RLinea!nTasaFin Then
        txtTasInt.Visible = True
        txtTasInt.Enabled = True
        If txtTasInt.Text >= 0.001 And txtTasInt.Text < RLinea!nTasaFin Then
        Else
        txtTasInt.Text = Format(RLinea!nTasaFin, "#0.0000")
        txtInteresTasa.Text = Format(RLinea!nTasaFin, "#0.0000")
        End If
        'TxtTasaGracia.Text = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
        If Not IsNull(RLinea!nTasaGraciaFin) Then
            If TxtTasaGracia.Text >= 0.001 And TxtTasaGracia.Text < RLinea!nTasaGraciaFin Then
            Else
            TxtTasaGracia.Text = Format(RLinea!nTasaGraciaFin, "#0.0000")
            End If
            txtTasInt.ToolTipText = "Minima : " & Format(RLinea!nTasaIni, "#0.0000") & "  Maxima : " & Format(RLinea!nTasaFin, "#0.0000")
        Else
        TxtTasaGracia.Text = 0
        End If
    End If
    If RLinea!nTasaGraciaIni <> RLinea!nTasaGraciaFin Then
        'TxtTasaGracia.Visible = True 'LUCV20180601, Comentó según ERS022-2018
        TxtTasaGracia.Enabled = True
        If Not IsNull(RLinea!nTasaGraciaFin) Then
            If TxtTasaGracia.Text >= 0.001 And TxtTasaGracia.Text < RLinea!nTasaGraciaFin Then
            Else
            TxtTasaGracia.Text = Format(RLinea!nTasaGraciaFin, "#0.0000")
            End If
            'TxtTasaGracia.Text = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
            TxtTasaGracia.ToolTipText = "Minima : " & Format(RLinea!nTasaGraciaIni, "#0.0000") & "  Maxima : " & Format(RLinea!nTasaGraciaFin, "#0.0000")
        Else
            TxtTasaGracia = 0#
        End If
    End If
End Sub

'******************************************************************
' CMACICA_CSTS - 12/11/2003 - -----------------------------------------------------------
Private Sub CargaFuenteIngreso(ByVal pnIndice As Integer, ByRef poPersona As UPersona_Cli, _
                                ByVal prsFte As ADODB.Recordset)
    Set oPersona = poPersona
    nIndice = pnIndice
    nProcesoEjecutado = 3
    Call CargaDatosFteIngreso(pnIndice, poPersona, prsFte)
    Call HabilitaBalance(False)
    Call HabilitaIngresosEgresos(False)
    If nTipoFteIngreso = gPersFteIngresoTipoDependiente Then
        SSTab1.TabVisible(5) = False
        SSTab1.TabVisible(4) = True
    Else
        SSTab1.TabVisible(4) = False
        SSTab1.TabVisible(5) = True
    End If
End Sub

Private Sub CargaDatosFteIngreso(ByVal pnIndice As Integer, _
                                ByRef poPersona As UPersona_Cli, _
                                ByVal prsFte As ADODB.Recordset, _
                                Optional ByVal pnFteDetalle As Integer = -1)
Dim nUltFte As Integer

    If CInt(poPersona.ObtenerFteIngTipo(pnIndice)) = gPersFteIngresoTipoDependiente Then
        If poPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) = 0 Then
            Call poPersona.RecuperaFtesIngresoDependiente(pnIndice, prsFte)
        End If
        Call HabilitaBalance(False)
    Else
        If poPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) = 0 Then
            Call poPersona.RecuperaFtesIngresoIndependiente(pnIndice, prsFte)
        End If
        Call HabilitaBalance(True)
    End If

'    cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, poPersona.ObtenerFteIngMoneda(pnIndice))
    cMonedaFteCod = poPersona.ObtenerFteIngMoneda(pnIndice)
    'Carga Ingresos y Egresos
    DTPFecIni.value = CDate(Format(poPersona.ObtenerFteIngInicioFuente(pnIndice), "dd/mm/yyyy"))
    TxtCargo.Text = poPersona.ObtenerFteIngCargo(pnIndice)
    Txtcomentarios.Text = poPersona.ObtenerFteIngComentarios(pnIndice)

    If poPersona.ObtenerFteIngIngresoTipo(pnIndice) = gPersFteIngresoTipoDependiente Then
        If pnFteDetalle = -1 Then
            nUltFte = poPersona.ObtenerFteIngIngresoNumeroFteDep(pnIndice) - 1
        Else
            nUltFte = pnFteDetalle
        End If
        TxtIngCon.Text = Format(poPersona.ObtenerFteIngIngresoConyugue(pnIndice, nUltFte), "#0.00")
        txtIngFam.Text = Format(poPersona.ObtenerFteIngIngresoFam(pnIndice, nUltFte), "#0.00")
        txtOtroIng.Text = Format(poPersona.ObtenerFteIngIngresoOtros(pnIndice, nUltFte), "#0.00")
        'LblIngresos.Caption = Format(poPersona.ObtenerFteIngIngresos(pnIndice, nUltFte), "#0.00")
        txtEgreFam.Text = Format(poPersona.ObtenerFteIngGastoFam(pnIndice, nUltFte), "#0.00")
        'lblSaldo.Caption = CDbl(txtIngFam.Text) + CDbl(txtOtroIng.Text) + CDbl(LblIngresos.Caption) - CDbl(txtEgreFam.Text)
    Else
        If pnFteDetalle = -1 Then
            nUltFte = poPersona.ObtenerFteIngIngresoNumeroFteIndep(pnIndice) - 1
        Else
            nUltFte = pnFteDetalle
        End If
        'Carga el Balance
        txtDisponible.Text = Format(poPersona.ObtenerFteIngActivoDisp(pnIndice, nUltFte), "#0.00")
        txtcuentas.Text = Format(poPersona.ObtenerFteIngCtasxCob(pnIndice, nUltFte), "#0.00")
        txtInventario.Text = Format(poPersona.ObtenerFteIngInventario(pnIndice, nUltFte), "#0.00")
        txtactivofijo.Text = Format(poPersona.ObtenerFteIngActivoFijo(pnIndice, nUltFte), "#0.00")

        txtProveedores.Text = Format(poPersona.ObtenerFteIngProveedores(pnIndice, nUltFte), "#0.00")
        txtOtrosPrest.Text = Format(poPersona.ObtenerFteIngOtrosCreditos(pnIndice, nUltFte), "#0.00")
        txtPrestCmact.Text = Format(poPersona.ObtenerFteIngCreditosCmact(pnIndice, nUltFte), "#0.00")

        txtVentas.Text = Format(poPersona.ObtenerFteIngVentas(pnIndice, nUltFte), "#0.00")
        txtrecuperacion.Text = Format(poPersona.ObtenerFteIngRecupCtasxCobrar(pnIndice, nUltFte), "#0.00")
        txtcompras.Text = Format(poPersona.ObtenerFteIngComprasMercad(pnIndice, nUltFte), "#0.00")
        txtOtrosEgresos.Text = Format(poPersona.ObtenerFteIngOtrosEgresos(pnIndice, nUltFte), "#0.00")
        TxtBalIngFam.Text = Format(poPersona.ObtenerFteIngBalIngFam(pnIndice, nUltFte), "#0.00")
        TxtBalEgrFam.Text = Format(poPersona.ObtenerFteIngBalEgrFam(pnIndice, nUltFte), "#0.00")
    End If
End Sub

'ARCV 13-08-2006
Private Sub CargarDatosFuenteIngreso(ByVal prsFte As ADODB.Recordset, _
                                    ByVal pnTipoFteingreso As PersFteIngresoTipo)

    nProcesoEjecutado = 3
    Call HabilitaBalance(False)
    Call HabilitaIngresosEgresos(False)
    
    cMonedaFteCod = Trim(str(prsFte!cPersFIMoneda))

    If pnTipoFteingreso = gPersFteIngresoTipoIndependiente Then
        Call HabilitaBalance(True)
        SSTab1.TabVisible(4) = False
        SSTab1.TabVisible(5) = True
           
        txtDisponible.Text = Format(prsFte!nPersFIActivoDisp, "#0.00")
        txtcuentas.Text = Format(prsFte!nPersFICtasxCobrar, "#0.00")
        txtInventario.Text = Format(prsFte!nPersFIInventarios, "#0.00")
        txtactivofijo.Text = Format(prsFte!nPersFIActivosFijos, "#0.00")
        txtProveedores.Text = Format(prsFte!nPersFIProveedores, "#0.00")
        txtOtrosPrest.Text = Format(prsFte!nPersFICredOtros, "#0.00")
        txtPrestCmact.Text = Format(prsFte!nPersFICredCMACT, "#0.00")
        txtVentas.Text = Format(prsFte!nPersFIVentas, "#0.00")
        txtrecuperacion.Text = Format(prsFte!nPersFIRecupCtasXCobrar, "#0.00")
        txtcompras.Text = Format(prsFte!nPersFICostoVentas, "#0.00")
        txtOtrosEgresos.Text = Format(prsFte!nPersFIEgresosOtros, "#0.00")
        TxtBalIngFam.Text = Format(prsFte!nPersIngFam, "#0.00")
        TxtBalEgrFam.Text = Format(prsFte!nPersEgrFam, "#0.00")
    Else    'Dependiente
        Call HabilitaBalance(False)
        SSTab1.TabVisible(5) = False
        SSTab1.TabVisible(4) = True
        
        'cMonedaFteCod = Trim(Str(prsFte!cPersFIMoneda))
        'Carga Ingresos y Egresos
        DTPFecIni.value = CDate(Format(prsFte!dPersFIInicio, "dd/mm/yyyy"))
        TxtCargo.Text = prsFte!cPersFICargo
        Txtcomentarios.Text = prsFte!cPersFIComentario

        TxtIngCon.Text = Format(prsFte!nPersIngCon, "#0.00")
        txtIngFam.Text = Format(prsFte!nPersIngCli, "#0.00")
        txtOtroIng.Text = Format(prsFte!nPersOtrIng, "#0.00")
        txtEgreFam.Text = Format(prsFte!nPersGastoFam, "#0.00")
    End If
End Sub

'----------------------------
Private Sub HabilitaIngresosEgresos(ByVal pnHabilitar As Boolean)
    txtOtroIng.Enabled = pnHabilitar
    txtIngFam.Enabled = pnHabilitar
    txtEgreFam.Enabled = pnHabilitar
    DTPFecIni.Enabled = pnHabilitar
    TxtCargo.Enabled = pnHabilitar
    Txtcomentarios.Enabled = pnHabilitar
    TxtIngCon.Enabled = pnHabilitar
End Sub

Private Sub HabilitaBalance(ByVal HabBalance As Boolean)
    'Label6.Enabled = HabBalance
    lblActivo.Enabled = HabBalance
    'Label9.Enabled = HabBalance
    lblPasPatrim.Enabled = HabBalance
    'Label7.Enabled = HabBalance
    lblActCirc.Enabled = HabBalance
    Label10.Enabled = HabBalance
    lblPasivo.Enabled = HabBalance
    Label12.Enabled = HabBalance
    txtDisponible.Enabled = HabBalance
    Label19.Enabled = HabBalance
    txtProveedores.Enabled = HabBalance
    Label13.Enabled = HabBalance
    txtcuentas.Enabled = HabBalance
    Label18.Enabled = HabBalance
    txtOtrosPrest.Enabled = HabBalance
    Label14.Enabled = HabBalance
    txtInventario.Enabled = HabBalance
    Label17.Enabled = HabBalance
    txtPrestCmact.Enabled = HabBalance
    'Label8.Enabled = HabBalance
    txtactivofijo.Enabled = HabBalance
    Label11.Enabled = HabBalance
    lblPatrimonio.Enabled = HabBalance
    Label15.Enabled = HabBalance
    txtVentas.Enabled = HabBalance
    Label5.Enabled = HabBalance
    txtcompras.Enabled = HabBalance
    Label20.Enabled = HabBalance
    txtrecuperacion.Enabled = HabBalance
    Label4.Enabled = HabBalance
    txtOtrosEgresos.Enabled = HabBalance
    Frame4.Enabled = HabBalance
    'Frame5.Enabled = HabBalance

    If nTipoFteIngreso = 2 Then
        SSTab1.TabVisible(5) = True
        SSTab1.TabVisible(4) = False
    Else
        SSTab1.TabVisible(5) = False
        SSTab1.TabVisible(4) = True
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------
Private Sub CargaApoderados(ByVal prsApoder As ADODB.Recordset, _
                           ByVal pnIndice As Integer, _
                           ByVal pbMuestraSoloAnalistaActual As Integer)

    CmbApoderado.Clear
    prsApoder.MoveFirst
    Do While Not prsApoder.EOF
        CmbApoderado.AddItem PstaNombre(prsApoder!cPersNombre) & Space(100) & prsApoder!cPersCod
        prsApoder.MoveNext
    Loop
    'R.Close
    'Set R = Nothing
    
    If pbMuestraSoloAnalistaActual = 1 And pnIndice <> -99 Then
        CmbApoderado.Enabled = False
        CmbApoderado.ListIndex = pnIndice
    End If
    
End Sub

Private Sub CargaControles()
    'Cargar los Apoderados
    Dim oCred As COMNCredito.NCOMCredito
    Dim rsDes As ADODB.Recordset
    Dim rsCon As ADODB.Recordset
    Dim rsCon2 As ADODB.Recordset
    Dim rsApoder As ADODB.Recordset
    Dim rsCondCred2 As ADODB.Recordset
    Dim nIndice As Integer
    Dim bMuestraSoloAnalistaActual As Integer
    
    Set oCred = New COMNCredito.NCOMCredito
    '*** PEAC 20091120 - SE AGREGO EL PARAMETRO gsCodAge
    Call oCred.CargaControlesAprobacion(gsCodPersUser, rsDes, rsCon, rsCon2, rsApoder, nIndice, bMuestraSoloAnalistaActual, rsCondCred2, gsCodAge)
    Set oCred = Nothing
    Call Llenar_Combo_con_Recordset(rsDes, CmbDestino)
    Call Llenar_Combo_con_Recordset(rsCon, CmbCondicion)
    Call Llenar_Combo_con_Recordset(rsCon2, CmbCondicionOtra)
    Call Llenar_Combo_con_Recordset(rsCondCred2, CmbCampana)
    Call CargaApoderados(rsApoder, nIndice, bMuestraSoloAnalistaActual)
End Sub

Private Sub LimpiarPantalla()
    fbLimpiaExoSegDes = True 'WIOR 20130710
    'Limpia Datos de Cliente
    bRefinanc = False
    bSustiDeudor = False
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
    bDesembParcialGenerado = False
    nNumGastos = 0
    ReDim MatGracia(0, 0)
    ReDim MatCalend(0, 0)
    ReDim MatDesemb(0, 0)
    ReDim MatGastos(0, 0)
    ReDim MatCalendSegDes(4, 0) 'WIOR 20140528
    SpnPlazo.valor = "0"
    LimpiaControles Me
   ' Call LimpiaFuentesIngreso
    InicializaCombos Me
    LimpiaFlex FERelacCred
    LimpiaFlex FEGarantCred
    
    Me.lblFormEval.Caption = "" 'PEAC 20160822
    
    ActxCta.NroCuenta = ""
    ActxCta.Enabled = True
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    ActxCta.SetFocusProd
    SpnCuotasApr.valor = "0"
    OptTCuota(0).value = True
    OptTCuota(1).Enabled = True
    OptTCuota(2).Enabled = True
    OptTCuota(3).Enabled = True
    OptDesemb(0).value = True
    OptTipoCalend(0).value = True
    OptTPer(0).value = True
    OptTPer(1).Enabled = True
    Frame5.Enabled = False
    cmdGrabar.Enabled = False
    CmdCalend.Enabled = False
    CmdGastos.Enabled = False
    OptTipoCalend(0).value = True
    OptTipoCalend(1).Enabled = True
    ChkCuotaCom.Enabled = True
    ChkCuotaCom.value = 0
    ChkMiViv.Enabled = True
    ChkMiViv.value = 0
    ChkBloqueo.value = 0
    ChkReglamento.value = 0
    ChkExoneracion.value = 0
    If nIndtempo = -99 Then
    Else
        CmbApoderado.ListIndex = nIndtempo
    End If
        
    '**** 03-04-06
    CmbApoderado.ListIndex = -1
    txtBuscarLinea.Text = ""
    '**************
    MatCredVig = ""
    
    'DAOR 20070202 ********************
    chkDesembBcoNac.value = 0
    Set frmCredDesembBcoNac = Nothing
    '**********************************
    txtFechaFija = "__/__/____" 'MAVM 25102010
    'BRGO 20111111 **************************
    If lsTpoProdCod = "517" Then
       LimpiaFlex grdEmpVinculados
       fraEmpVinculados.Visible = False
    End If
    '*** END BRGO **********************
    fnCuotaBalon = 0 'WIOR 20131111
    Set rsExonera = Nothing
    cmdCheckList.Enabled = False 'RECO20150415 ERS010-2015
    fbGrabar = False 'WIOR 20150717
    'APRI20180404 ERS036-2017
    cmdEnvioEstCta.Enabled = False
    CmbDestino.Enabled = True
    CmbCondicion.Enabled = True
    CmbCondicionOtra.Enabled = True
    CmbApoderado.Enabled = True
    CmbCampana.Enabled = True
    cmdVinculados.Enabled = True
    txtTasInt.Enabled = True
    txtBuscarLinea.Enabled = True
    txtBuscarLinea.rs = New ADODB.Recordset
    Dim i As Integer
    For i = 0 To 20
     Clientes(i, 0) = ""
     Clientes(i, 1) = ""
    Next i
    'END APRI
    chkExoneraMYPE.value = 0 'APRI20180821 ERS061-2018
End Sub

Private Function ValidaDatos(Optional ByVal pbValidaLineaCredito As Boolean = False) As Boolean
Dim sCad As String
Dim oCred As COMNCredito.NCOMCredito
Dim sValor As String
'Dim nValor As Double
Dim bEnc As Boolean
Dim bEncPol As Boolean
Dim i As Integer
'WIOR 20120529*******************************************
Dim oCredMM As COMDCredito.DCOMCredito
Dim oCredMMBD As COMDCredito.DCOMCredActBD
Dim rsCredMM As ADODB.Recordset
Dim rsBenfMueb As ADODB.Recordset
'WIOR FIN ***********************************************
'WIOR 20130422 *****************************************************************
Dim oPersonaU As COMDPersona.UCOMPersona
Dim oTipoCam As COMDConstSistema.NCOMTipoCambio
Dim nTC As Double
Dim nTipoConBN As Integer
Dim MatAnexos() As String
Dim sAnexos As String
Dim j As Integer
'WIOR FIN ***********************************************************************
'WIOR 20140820 *****************************************************************
Dim RSRelaciones As ADODB.Recordset
Dim nMinAprobacion As Long
Dim nMaxAprobacion As Long
Dim nPermanencia As Long
Dim nCuotas As Long
Dim nMontoMax As Double
Dim bNecesario As Boolean
Dim nEdadActual As Long
Dim nEdadFuturo As Long
Dim nTiempo As Double
Dim dFuturo As Date
Dim sMensajeValida As String
Dim sMensajePersona As String
Dim nMontoMN As Double
Dim lbCambioTasa As Integer
'WIOR FIN **********************************************************************
Dim lsMsg As String 'EJVG20160527
    
    VerificarFechaSistema Me, True 'EJVG20151020 -> validar fecha de sistema, en caso no apaguen sus PC los usuarios les sacará del sistema
    i = 0
    ValidaDatos = True
    bEnc = False
    bEncPol = False
    'Valida Monto a Aprobar
    If Len(Trim(TxtMontoApr.Text)) = 0 Then
        MsgBox "Ingrese el Monto del Prestamo", vbInformation, "Aviso"
        If gnAgenciaCredEval = 0 Then TxtMontoApr.SetFocus 'JUEZ 20121218
        ValidaDatos = False
        Exit Function
    End If
    If CDbl(TxtMontoApr.Text) = 0 Then
        MsgBox "Monto del Prestamo debe ser Mayor que Cero", vbInformation, "Aviso"
        If TxtMontoApr.Enabled Then
            If gnAgenciaCredEval = 0 Then TxtMontoApr.SetFocus 'JUEZ 20121218
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    ' Valida tasa de Gracia
    If CInt(TxtDiasGracia.Text) > 0 Then
        If CDbl(Me.TxtTasaGracia.Text) = 0 Then
            MsgBox "La Tasa de Interes de Gracia debe ser Mayor Cero", vbInformation, "Aviso"
            ValidaDatos = False
            'TxtDiasGracia.SetFocus 'Comentado Por MAVM 25102010
            'If gnAgenciaCredEval = 0 Then TxtTasaGracia.SetFocus 'JUEZ 20121218 'LUCV2018060101, Comentó según ERS022-2018
            Exit Function
        End If
    End If
    'ALPA 201501150***********************************************
    lbCambioTasa = 0
    If CDbl(txtTasInt.Text) < CDbl(lnTasaInicial) Then
        If CDbl(txtTasInt.Text) <= Val(TxtTasaGracia.Text) Then
            lbCambioTasa = 1
        End If
    End If
    If CInt(TxtDiasGracia.Text) > 0 And TxtTasaGracia.Visible Then
        If CDbl(TxtTasaGracia.Text) < RLinea!nTasaGraciaIni Or CDbl(TxtTasaGracia.Text) > RLinea!nTasaGraciaFin Then
            If lbCambioTasa = 0 Then
                MsgBox "La Tasa de Interes para el Periodo de Gracia, No es Permitida por la Configuración del Producto", vbInformation, "Aviso"
                ValidaDatos = False
                TxtTasaGracia.SetFocus
                Exit Function
            End If
        End If
    End If
    '*************************************************************
    'Valida Fecha de Desembolso
    sCad = ValidaFecha(TxtFecDesemb.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        ValidaDatos = False
        If gnAgenciaCredEval = 0 Then TxtFecDesemb.SetFocus 'JUEZ 20121218
        Exit Function
    End If
    
    'Valida Monto a Aprobar
    If Len(Trim(TxtMontoApr.Text)) = 0 Then
        MsgBox "Ingrese el Monto del Prestamo", vbInformation, "Aviso"
        If gnAgenciaCredEval = 0 Then TxtMontoApr.SetFocus 'JUEZ 20121218
        ValidaDatos = False
        Exit Function
    End If
    If CDbl(TxtMontoApr.Text) = 0 Then
        MsgBox "Monto del Prestamo debe ser Mayor que Cero", vbInformation, "Aviso"
        If TxtMontoApr.Enabled Then
            If gnAgenciaCredEval = 0 Then TxtMontoApr.SetFocus 'JUEZ 20121218
        End If
        ValidaDatos = False
        Exit Function
    End If
    'Valida NroCuotas
    If Not OptTCuota(3).value Then
        If CInt(SpnCuotasApr.valor) = 0 Then
            MsgBox "El Numero de Cuotas debe ser Mayor a Cero", vbInformation, "Aviso"
            If gnAgenciaCredEval = 0 Then SpnCuotasApr.SetFocus 'JUEZ 20121218
            ValidaDatos = False
            Exit Function
        End If
    End If
    'valida el Plazo
    If Not OptTCuota(3).value Then
        If OptTPer(0).value Then
            If CInt(SpnPlazo.valor) = 0 Then
                MsgBox "El Plazo de las Cuotas debe ser Mayor a Cero", vbInformation, "Aviso"
                If gnAgenciaCredEval = 0 Then SpnPlazo.SetFocus 'JUEZ 20121218
                ValidaDatos = False
                Exit Function
            End If
        End If
    End If
    'MADM 20110430
    

   'CTI5 ERS0012020*************************************
   Dim oGarant As New COMNCredito.NCOMGarantia
   Dim oRsGarantiaSuge As New ADODB.Recordset
   '****************************************************
   For i = 1 To FEGarantCred.rows - 1
        'If Trim(FEGarantCred.TextMatrix(i, 10)) = "Desaprobado" Or Trim(FEGarantCred.TextMatrix(i, 10)) = "Pendiente" Then
        If UCase(Trim(FEGarantCred.TextMatrix(i, 10))) = "DESAPROBADO" Or UCase(Trim(FEGarantCred.TextMatrix(i, 10))) = "PENDIENTE" Then 'EJVG20150706
            'CTI5 ERS0012020***********************************************
            If lsTpoCredCod = "853" Or lsTpoCredCod = "854" Then
                Set oGarant = New COMNCredito.NCOMGarantia
                Set oRsGarantiaSuge = oGarant.ObtenerGarantiaBienFuturoCreditoHipo(Trim(FEGarantCred.TextMatrix(i, 9)))
                If Not (oRsGarantiaSuge.EOF And oRsGarantiaSuge.BOF) Then
                    bEnc = False
                Else
                    bEnc = True
                End If
            Else
                bEnc = True
            End If
            'bEnc = True
            '*************************************************************
            Exit For
        End If
    Next i
    'CTI5 ERS0012020*******************************
    RSClose oRsGarantiaSuge
    '**********************************************
    Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
    If objProducto.GetResultadoCondicionCatalogo("N0000073", ActxCta.Prod) Then
    'If Me.ActxCta.Prod <> "517" Then 'BRGO 20111111 Salta verificación del área legal para Créditos Ecotaxi
        If bEnc Then
            MsgBox "Verifique la Garantía en la pestaña Garantias, Comuníquese con el Área de Legal / Sup. Créditos en Agencias", vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
    End If
    'END MADM
    'EJVG20160526 *** Validar cobertura
    lsMsg = ""
    Set oCred = New COMNCredito.NCOMCredito
    lsMsg = oCred.ValidaGarantia(Trim(Me.ActxCta.NroCuenta), gdFecSis, CDbl(TxtMontoApr.Text), fbEsAmpliado)
    If Len(lsMsg) > 0 Then
        MsgBox lsMsg, vbInformation, "Aviso"
        ValidaDatos = False
        SSTab1.Tab = 6
        Exit Function
    End If
    Set oCred = Nothing
    'END EJVG *******
     'MADM 2011050511
   For i = 1 To FEGarantCred.rows - 1
        'If Trim(FEGarantCred.TextMatrix(i, 11)) = "VENCIDO" Then
        If UCase(Trim(FEGarantCred.TextMatrix(i, 11))) = "VENCIDO" Then 'EJVG20150706
            bEncPol = True
            Exit For
        End If
    Next i
    
    If bEncPol Then
        MsgBox "El estado de la Póliza es Vencida, Comuníquese con el Area de Sup. Créditos", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
    End If
    'END MADM
    
    If Not OptTCuota(3).value Then
            If OptTPer(1).value Then
                If CInt(TxtDiaFijo.Text) = 0 Then
                    MsgBox "Ingrese el Dia del mes que Vencen las Cuotas", vbInformation, "Aviso"
                    If gnAgenciaCredEval = 0 Then TxtDiaFijo.SetFocus 'JUEZ 20121218
                    ValidaDatos = False
                    Exit Function
                End If
            End If
     End If
'    End If
    '** End gitu
    'ALPA 20160419********************************************
    If SpnPlazo.valor < 30 And OptTPer(0).value = True Then
        MsgBox "El crédito no cumple con las especificaciones de plazo, no debe ser menor a 30 días", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    '*********************************************************
    If CmbApoderado.ListIndex = -1 Then
        MsgBox "Seleccione el Apoderado del Credito", vbInformation, "Aviso"
        If gnAgenciaCredEval = 0 Then CmbApoderado.SetFocus 'JUEZ 20121218
        ValidaDatos = False
        Exit Function
    End If
    
    If CmbDestino.ListIndex = -1 Then
        MsgBox "Seleccione el Destino del Credito", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    If CmbCondicion.ListIndex = -1 Then
        MsgBox "Seleccione la Condicion del Credito", vbInformation, "Aviso"
        If gnAgenciaCredEval = 0 Then CmbCondicion.SetFocus 'JUEZ 20121218
        ValidaDatos = False
        Exit Function
    End If
    
    ' CMACICA_CSTS - 05/11/2003 -------------------------------------------------------------
    If CmbCondicionOtra.ListIndex = -1 Then
        MsgBox "Seleccione la Condicion 2 del Credito", vbInformation, "Aviso"
        If gnAgenciaCredEval = 0 Then CmbCondicionOtra.SetFocus 'JUEZ 20121218
        ValidaDatos = False
        Exit Function
    End If
    ' ---------------------------------------------------------------------------------------
    
   '**DAOR 20070202**************************
    If chkDesembBcoNac.value = 1 Then
        If frmCredDesembBcoNac.cmbAgencia.ListIndex = -1 Or frmCredDesembBcoNac.txtCodigo.Text = "" Then
            MsgBox "Ingrese los datos para el desembolso en el Banco de la Nación", vbInformation, "Aviso"
            frmCredDesembBcoNac.Show 1
            ValidaDatos = False
            Exit Function
        End If
    End If
    '******************************************
    
    '**DAOR 20080109, Validacion de documento de identidad, solo para desemb. en bco. nac.****
    If chkDesembBcoNac.value = 1 Then
        If fnPersPersoneria = 1 And fnPersTipoDoc <> gPersIdDNI Then
            MsgBox "No es posible aprobar el crédito como desembolso en el banco de la nación debido a que el titular del crédito no posee DNI como documento de identidad", vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
    End If
    '*****************************************************************************************
    
    'MAVM 15112010 ***
    If OptTPer(0).value = True Then
        If CDate(CDate(TxtFecDesemb.Text) + CDate(SpnPlazo.valor) + CDate(TxtDiasGracia.Text)) <> CDate(txtFechaFija.Text) Then
            MsgBox "Falto Presionar Enter en el Campo Fecha de Desembolso", vbInformation, "Aviso"
            If gnAgenciaCredEval = 0 Then TxtFecDesemb.SetFocus 'JUEZ 20121218
            ValidaDatos = False
            Exit Function
        End If
    Else
        If CDate(TxtDiasGracia.Text) <> "0" Then
            If CDate(CDate(TxtFecDesemb.Text) + CDate(30) + CDate(TxtDiasGracia.Text)) <> CDate(txtFechaFija.Text) Then
                MsgBox "Falto Presionar Enter en el Campo Fecha de Desembolso", vbInformation, "Aviso"
                If gnAgenciaCredEval = 0 Then TxtFecDesemb.SetFocus 'JUEZ 20121218
                ValidaDatos = False
                Exit Function
            End If
        End If
    End If
    '***
    ''JOEP ERS066-20161118
    If gnAgenciaCredEval = 1 Then 'Agencia Piloto
            If CmbCampana.Text = "" Then
                MsgBox "El Credito esta  vinculado a una Campaña no Activa, favor de Comunicarse con el Analista del Credito", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
    End If
    'JOEP ERS066-20161118
    'MAVM 20110124 ***
    If CmbCampana.ListIndex = -1 Then
        MsgBox "Seleccione la Campaña del Credito", vbInformation, "Aviso"
        If gnAgenciaCredEval = 0 Then CmbCampana.SetFocus 'JUEZ 20121218
        ValidaDatos = False
        Exit Function
    End If
    'WIOR 20120529*******************************************
    Set oCredMM = New COMDCredito.DCOMCredito
    Set oCredMMBD = New COMDCredito.DCOMCredActBD
    Set rsCredMM = oCredMM.ObtenerMicroseguro(Me.ActxCta.NroCuenta)
    If rsCredMM.RecordCount > 0 Then
        If Trim(rsCredMM!cNumCert) = "" Then
             MsgBox "No ha ingresado el Nº Certificado de Microseguros.", vbInformation, "Aviso"
             ValidaDatos = False
             Exit Function
        End If
        Set rsBenfMueb = oCredMMBD.ObtenerBeneficiariosMicroseguro(Me.ActxCta.NroCuenta)
        If rsBenfMueb.RecordCount <= 0 Then
            If MsgBox("No ha registrado ningun beneficiario en Microseguros, Desea Continuar sin registrar a nadie?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                ValidaDatos = False
                Exit Function
            End If
        End If
    End If
    Set rsCredMM = oCredMM.ObtenerMultiriesgo(Me.ActxCta.NroCuenta)
    If rsCredMM.RecordCount > 0 Then
        If Trim(rsCredMM!cNumCert) = "" Then
             MsgBox "No ha ingresado el Nº Certificado del Seguro Multiriesgo.", vbInformation, "Aviso"
             ValidaDatos = False
             Exit Function
        End If
        Set rsBenfMueb = oCredMMBD.ObtenerMueblesMultiriesgo(Me.ActxCta.NroCuenta)
        If rsBenfMueb.RecordCount <= 0 Then
            MsgBox "Aun no ha registrado ningun Mueble en su Seguro Multiriesgo.", vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
    End If
    Set oCredMM = Nothing
    Set oCredMMBD = Nothing
    Set rsCredMM = Nothing
    Set rsBenfMueb = Nothing
    
    Set oPersonaU = New COMDPersona.UCOMPersona
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    
    nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoMes)
    nTipoConBN = 0
    sAnexos = ""
    
    If oPersona.Personeria = 1 Then
        If oPersona.Nacionalidad <> "04028" Or oPersona.Residencia <> 1 Or oPersona.RPeps = 1 Then
                sAnexos = sAnexos & "05,16,15,"
        ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersona.ObtenerDNI) = "", oPersona.ObtenerNumeroDoc(0), oPersona.ObtenerDNI), oPersona.ObtenerRUC, nTipoConBN, oPersona.NombreCompleto) Then
            If nTipoConBN = 1 Or nTipoConBN = 3 Then
                sAnexos = sAnexos & "05,16,15,"
            End If
        End If
    Else
        If oPersonaU.ValidaEnListaNegativaCondicion(oPersona.ObtenerDNI, oPersona.ObtenerRUC, nTipoConBN, oPersona.NombreCompleto) Then
            If nTipoConBN = 1 Or nTipoConBN = 3 Then
                sAnexos = sAnexos & "05,16,15,"
            End If
        End If
    End If
    
    
    If Mid(Trim(ActxCta.NroCuenta), 9, 1) = "1" Then
        If CDbl(TxtMontoApr.Text) > 200000 Then
            sAnexos = sAnexos & "05,15,"
        End If
    Else
        If (CDbl(TxtMontoApr.Text) * nTC) > 200000 Then
            sAnexos = sAnexos & "05,15,"
        End If
    End If
    
    If Trim(Right(CmbCondicion.Text, 5)) = "1" Then
        sAnexos = sAnexos & "15,"
    End If
    
    If oPersona.SujetoObligado = 1 Then
        sAnexos = sAnexos & "19,"
    End If
    
    If sAnexos <> "" Then
        sAnexos = Mid(sAnexos, 1, Len(sAnexos) - 1)
        MatAnexos = Split(sAnexos, ",")
        
        If UBound(MatAnexos) > 0 Then
            sAnexos = ""
            
            'Quitar Duplicados
            For i = 0 To UBound(MatAnexos)
                For j = 0 To UBound(MatAnexos)
                    If i <> j Then
                        If MatAnexos(i) <> MatAnexos(j) Then
                            If InStr(1, sAnexos, MatAnexos(j)) = 0 Then
                                sAnexos = sAnexos & MatAnexos(j) & ","
                            End If
                        End If
                    End If
                Next j
            Next i
            
            sAnexos = Mid(sAnexos, 1, Len(sAnexos) - 1)
            MatAnexos = Split(sAnexos, ",")
            sAnexos = ""
            
            'Ordenar Lista de Anexos
            For i = 0 To UBound(MatAnexos)
                For j = 0 To UBound(MatAnexos) - 1
                    If CDbl(MatAnexos(j)) > CDbl(MatAnexos(j + 1)) Then
                        sAnexos = MatAnexos(j)
                        MatAnexos(j) = MatAnexos(j + 1)
                        MatAnexos(j + 1) = sAnexos
                    End If
                Next j
            Next i
        End If
        
        sAnexos = ""
        For j = 0 To UBound(MatAnexos)
            sAnexos = sAnexos & "- Anexo Nº " & MatAnexos(j) & Chr(10)
        Next j
        
        If MsgBox("El Expediente de Crédito requiere los Siguientes Anexos: " & Chr(10) & sAnexos & Chr(10) & Space(20) & "Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            ValidaDatos = False
            Exit Function
        End If
    End If
    'ALPA 20140612******
    If (txtMontoMivivienda.Text = 0# Or txtMontoMivivienda.Text = "") And lsTpoProdCod = "801" Then
        MsgBox "No se olvide de asignar el valor de venta del crédito MIVIVIENDA", vbInformation, "Aviso"
        ValidaDatos = False
        SSTab1.Tab = 6
        txtMontoMivivienda.SetFocus
        Exit Function
    End If
    If Round(CDbl(txtMontoMivivienda.Text) - ((CDbl(txtMontoMivivienda.Text) * 1) / 10), 2) < CDbl(TxtMontoApr.Text) And lsTpoProdCod = "801" Then
        If Not (CDbl(txtMontoMivivienda.Text) = CDbl(TxtMontoApr.Text) And lsTpoProdCod = "801") Then
            MsgBox "MIVIVIENDA, no se olvide que el monto de la inicial no debe ser menor al 10% (" & Round((CDbl(TxtMontoApr.Text) * 10) / 9, 2) & ")", vbInformation, "Aviso"
            ValidaDatos = False
            SSTab1.Tab = 6
            txtMontoMivivienda.SetFocus
            Exit Function
        End If
    End If
    '*******************
    Set oPersonaU = Nothing
    Set oTipoCam = Nothing
    'WIOR FIN **********************************************************************
    
    'WIOR 20140820 *****************************************************************
    Set oCredMM = New COMDCredito.DCOMCredito
    Set oCred = New COMNCredito.NCOMCredito
    Set RSRelaciones = oCredMM.RecuperaRelacPers(Trim(ActxCta.NroCuenta))
    Set oCredMM = Nothing
    
    If OptTPer(0).value Then
        nTiempo = (CInt(SpnCuotasApr.valor) - 1) * CInt(SpnPlazo.valor)
        dFuturo = DateAdd("d", nTiempo, CDate(txtFechaFija.Text))
    ElseIf OptTPer(1).value Then
        nTiempo = CInt(SpnCuotasApr.valor) - 1
        dFuturo = DateAdd("m", nTiempo, CDate(txtFechaFija.Text))
    End If
    
    nMontoMN = CDbl(TxtMontoApr.Text) * IIf(Mid(Trim(ActxCta.NroCuenta), 9, 1) = "1", 1, nTC)
    
    If oCred.ConfigEdadProducto(lsTpoProdCod, nMinAprobacion, nMaxAprobacion, nPermanencia) Then
        If Not (RSRelaciones.EOF And RSRelaciones.BOF) Then
            sMensajeValida = ""
            For i = 1 To RSRelaciones.RecordCount
                nEdadActual = 0
                nEdadFuturo = 0
                sMensajePersona = ""
                If Trim(RSRelaciones!nPersPersoneria) = 1 And (Trim(RSRelaciones!nConsValor) = gColRelPersTitular Or Trim(RSRelaciones!nConsValor) = gColRelPersConyugue Or Trim(RSRelaciones!nConsValor) = gColRelPersCodeudor) Then
                    nEdadActual = EdadPersona(CDate(RSRelaciones!dPersNacCreac), gdFecSis)
                    nEdadFuturo = EdadPersona(CDate(RSRelaciones!dPersNacCreac), dFuturo)
                        
                        If nMinAprobacion > nEdadActual And nMinAprobacion > 0 Then
                            If sMensajePersona <> "" Then
                                sMensajePersona = sMensajePersona & vbNewLine
                            End If
                            sMensajePersona = sMensajePersona & "- Supera la edad mínima de " & nMinAprobacion & " años para Aprobación de Crédito."
                        End If
                        
                        If nMaxAprobacion <= nEdadActual Then
                            If sMensajePersona <> "" Then
                                sMensajePersona = sMensajePersona & vbNewLine
                            End If
                            
                            sMensajePersona = sMensajePersona & "- Supera la edad máxima de < " & nMaxAprobacion & " años para Aprobación de Crédito."
                        End If
                        
                        If nPermanencia <= nEdadFuturo And nPermanencia > 0 Then
                            If sMensajePersona <> "" Then
                                sMensajePersona = sMensajePersona & vbNewLine
                            End If
                            sMensajePersona = sMensajePersona & "- Supera la edad máxima de < " & nPermanencia & " años para Permanencia en el Crédito."
                        End If
                        
                        If Trim(sMensajePersona) <> "" Then
                            If Trim(sMensajeValida) <> "" Then
                                sMensajeValida = sMensajeValida & vbNewLine & vbNewLine
                            End If
                            sMensajePersona = "La Persona " & PstaNombre(Trim(RSRelaciones!cPersNombre)) & "(" & Trim(RSRelaciones!cConsDescripcion) & "):" & vbNewLine & sMensajePersona
                        End If
                End If
                
                If Trim(sMensajePersona) <> "" Then
                    sMensajeValida = sMensajeValida & sMensajePersona
                End If
                
                RSRelaciones.MoveNext
            Next i
            
            If Trim(sMensajeValida) <> "" Then
                If frmCredAprobacionAlertas.Inicio(sMensajeValida & vbNewLine & vbNewLine & "Para continuar, favor de ingresar su visto.") = False Then
                    ValidaDatos = False
                    Exit Function
                End If
            End If
        End If
    End If
    
'JOEP20190202 CP
    If Not CP_ValidaRangoFechaApro(lsTpoProdCod, 3000, IIf(OptTPer(0).value, 0, 1), txtFechaFija.Text) Then
            ValidaDatos = False
            txtFechaFija.SetFocus
            Exit Function
    End If
    If Not CP_CondicionesApro(lsTpoProdCod, Mid(ActxCta.NroCuenta, 9, 1), TxtMontoApr.Text, SpnCuotasApr.valor, SpnPlazo.valor, Trim(Right(CmbDestino.Text, 9)), TxtDiasGracia.Text, lsTipoCredito, gModAprobacion) Then
        ValidaDatos = False
        Exit Function
    End If
'JOEP20190202 CP

    If oCred.ConfigPeriodoProducto(lsTpoProdCod, nCuotas, nMontoMax, bNecesario) Then
            sMensajeValida = ""
            
            If CLng(SpnCuotasApr.valor) > nCuotas Then
                sMensajeValida = "- El plazo máximo del credito debe ser " & nCuotas & " meses."
            End If
            
            If nMontoMN > nMontoMax Then
                If sMensajeValida <> "" Then
                    sMensajeValida = sMensajeValida & Chr(10)
                End If
                
                sMensajeValida = sMensajeValida & "- El monto máximo del crédito es de S/. " & Format(nMontoMax, gsFormatoNumeroView) & "."
            End If
            
            If Trim(sMensajeValida) <> "" Then
                MsgBox sMensajeValida, vbInformation, "Aviso"
                If bNecesario Then
                    ValidaDatos = False
                    Exit Function
                End If
            End If
    End If
    
    Set RSRelaciones = Nothing
    Set oCred = Nothing
    'WIOR FIN **********************************************************************
    'EJVG20160526 *** Informe de Riesgos y Comumicar Riesgos
    If Not GenerarDataExposicionRiesgoUnico(Trim(Me.ActxCta.NroCuenta), Trim(LblCodCli.Caption), Trim(LblNomCli.Caption)) Then
        ValidaDatos = False
        Exit Function
    End If
    
    If Not bRefinanc Then
        If Not EmiteInformeRiesgo(eProcesoEmiteInformeRiesgo.Aprobacion, Trim(Me.ActxCta.NroCuenta), lsTpoProdCod, lsTpoCredCod, Trim(LblCodCli.Caption), Trim(LblNomCli.Caption), CDbl(TxtMontoApr.Text), fbEsAmpliado, SpnCuotasApr.valor) Then
            ValidaDatos = False
            Exit Function
        End If
    End If
    'END EJVG *******
End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'FRHU 20140402 ERS026-2014 RQ14131
        If ValidarUltimaActualizacion(ActxCta.NroCuenta) Then
            Exit Sub
        End If
        'FIN FRHU 20140402
        ''FRHU 20140331 ERS042-2014 RQ14180
        'Dim objCred As New COMDCredito.DCOMCredito
        'If objCred.ValidarComunicarRiesgoPend(Trim(ActxCta.NroCuenta)) Then
        '      MsgBox "Para la aprobación de este crédito se requiere la recepción de correo electronico Informativo " & vbNewLine & _
        '             "a la Gerencia de Riesgos, verifiquelo e Intentalo de Nuevo", vbInformation
        '      Exit Sub
        'End If
        'Set objCred = Nothing
        ''FIN FRHU 20140331
        'LUCV20160721 ERS004-2016**********************************************************************
        If Not CumpleCriteriosRatios(ActxCta.NroCuenta) Then
            MsgBox "El crédito no cumple con los criterios de ratios financieros. " & Chr(10) & " - Favor revisar la evaluación del crédito.", vbInformation, "Alerta"
            Exit Sub
        End If
        
        'LUCV FIN *************************************************************************************
        'CTI3 ERS003-2020**********************************************************************
         If Not CumpleCriteriosRatioLimite(ActxCta.NroCuenta) Then
            MsgBox "El crédito no cumple con el porcentaje del limite de cuota. " & Chr(10) & " - Favor revisar la evaluación del crédito.", vbInformation, "Alerta"
            Exit Sub
         End If
        '***MARG ERS007-2017***
         If Not CumpleCriteriosAprobacionCredito(ActxCta.NroCuenta, gsCodUser) Then
            MsgBox "Usted no puede aprobar un crédito que ha solicitado o sugerido", vbInformation, "Alerta"
            Exit Sub
        End If
        '***END MARG***********
        
        'ARLO 20170910
        Dim oDCreditos As COMDCredito.DCOMCreditos
        Set oDCreditos = New COMDCredito.DCOMCreditos
        
        If oDCreditos.VerificaClienteCampania(ActxCta.NroCuenta) Then
        MsgBox "Este Crédito pertenece a la Campaña Automático, por favor solicitar su aprobación por el SICMACM WEB.  ", vbInformation, "Aviso"
        Exit Sub
        End If
        'ARLO 20170910
        
        If CargaDatos(ActxCta.NroCuenta) Then
            ActxCta.Enabled = False
            'If gnAgenciaCredEval = 0 Then Frame5.Enabled = True 'JUEZ 20121218
            Frame5.Enabled = True
            'FRHU 20140212 RQ14005
            If Len(Me.actxCtaAho.NroCuenta) = 18 Then
                Me.btnVer.Enabled = True
            Else
                Me.btnVer.Enabled = False
            End If
            'FIN FRHU20140212
            TxtTasaGracia.Enabled = True 'Add Gitu 06-04-2009
            If gnAgenciaCredEval = 0 Then CmbDestino.SetFocus 'JUEZ 20121218
            cmdGrabar.Enabled = True
            CmdCalend.Enabled = True
            CmdGastos.Enabled = True
            'ALPA 20091007***********************
            cmdVinculados.Enabled = True
            '************************************
            'ALPA 20110525***********************
            Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
            'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            If objProducto.GetResultadoCondicionCatalogo("N0000074", ActxCta.Prod) Then
                Call txtFechaFija_KeyPress(13)
            End If
            '************************************
            cmdCheckList.Enabled = True 'RECO20150415 ERS010-2015
            If gnAgenciaCredEval = 1 Then 'JUEZ 20121218
                'APRI20180404 ERS036-2017
                'FraNiv1.Enabled = False
                CmbDestino.Enabled = False
                CmbCondicion.Enabled = False
                CmbCondicionOtra.Enabled = False
                CmbApoderado.Enabled = False
                CmbCampana.Enabled = False
                cmdVinculados.Enabled = False
                txtTasInt.Enabled = False
                frmCtaAho.Enabled = False
                txtBuscarLinea.Enabled = False
                'END APRI
                Frame12.Enabled = False
                fratipodes.Enabled = False
                Frame11.Enabled = False
                fraMiViv.Enabled = False
                
                If Not fbMIVIVIENDA Then 'WIOR 20160225
                    FraTpoCalend.Enabled = False
                End If
                FraNiv2.Enabled = False
                Frame9.Enabled = False
                OptTPer(0).Enabled = False
                OptTPer(1).Enabled = False
                Call txtFechaFija_KeyPress(13)
            End If
            cmdEnvioEstCta.Enabled = True  'APRI20180404 ERS036-2017
            chkExoneraMYPE.Enabled = True  'APRI20180820 ERS061-2018
            Call bValiExistCheckList 'JOEP20190201 CP
        Else
            ActxCta.Enabled = True
            Frame5.Enabled = False
            cmdGrabar.Enabled = False
            CmdCalend.Enabled = False
            CmdGastos.Enabled = False
            'ALPA 20091007***********************
            cmdVinculados.Enabled = False
            '************************************
            'If gnAgenciaCredEval = 0 And bAprobNivPendiente = False Then
            'If Not (bAprobNivPendiente Or fbSolicitudPoliza) Then 'EJVG20150704
            If Not (bAprobNivPendiente Or fbSolicitudPoliza Or fbSolicitudAutAmp) Then 'JUEZ 20160510
                MsgBox "No se Encontró el Crédito", vbInformation, "Aviso"
            End If
        End If
    End If
End Sub
'*******FRHU20140212 RQ14005
Private Sub btnVer_Click()
    If Len(actxCtaAho.NroCuenta) = 18 Then
        frmCuentaAhorrosDesembolsoTerceros.Inicio (actxCtaAho.NroCuenta)
    Else
        MsgBox "Llena todos los campos Correctamente", vbInformation, "Advertencia"
        actxCtaAho.SetFocusCuenta
    End If
End Sub
'******FRHU20140212 RQ14005

'WIOR 20130418 **************************************
Private Sub ChkExoneracion_Click()

'JOEP-20171007
If ChkExoneracion.value = 1 Then
    fbLimpiaExoSegDes = False
End If

If Not fbLimpiaExoSegDes Then
    Dim nEst As Integer
        If Not fbValidaExoSegDes Then
            frmCredExonSegDesg.Inicio Trim(ActxCta.NroCuenta), nEst
    
            If ChkExoneracion.value = nEst Then
                fbValidaExoSegDes = False
                ChkExoneracion.value = nEst
            Else
                fbValidaExoSegDes = False
                ChkExoneracion.value = 1
            End If
        Else
            fbValidaExoSegDes = False
        End If
End If
'JOEP-20171007
End Sub
'WIOR FIN *******************************************

Private Sub chkExoneraMYPE_Click()
 bCalendGen = False
End Sub

Private Sub ChkMiViv_Click()
    If ChkMiViv.value = 1 Then
        OptTCuota(0).value = True
        OptTCuota(1).Enabled = False
        OptTCuota(2).Enabled = False
        OptTPer(1).value = True
        OptTPer(0).Enabled = False
        ChkCuotaCom.Enabled = False
    Else
        OptTCuota(0).value = True
        OptTCuota(1).Enabled = True
        OptTCuota(2).Enabled = True
        OptTPer(0).value = True
        OptTPer(0).Enabled = True
        ChkCuotaCom.Enabled = True
    End If
End Sub

Private Sub ChkProxMes_Click()
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
End Sub

Private Sub chkTasa_Click()
    If lnExoneracionTasa = 1 Then
        chkTasa.value = 1
    Else
        chkTasa.value = 0
    End If
    If chkTasa.value = 0 Then
        txtInteresTasa.Enabled = False
        txtInteresTasa.Visible = False
        txtInteresTasa.Text = 0#
    Else
        txtInteresTasa.Enabled = True
        txtInteresTasa.Visible = True
        txtInteresTasa.Text = Format(txtTasInt.Text, "#0.0000")
    End If
End Sub

Private Sub ckcPreferencial_Click()
    If lnLogicoBuscarDatos = 1 Then
        If lnCliPreferencial = 1 Then
            ckcPreferencial.value = 1
        Else
            ckcPreferencial.value = 0
        End If
        Call CargarDatosProductoCrediticio
        Call MostrarLineas
    End If
End Sub

Private Sub CmbApoderado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If TxtMontoApr.Enabled Then
           TxtMontoApr.SetFocus
       Else
           If SpnCuotasApr.Enabled = True Then
               SpnCuotasApr.SetFocus
           Else
               SpnPlazo.SetFocus
           End If
       End If
    End If
End Sub

'MAVM 20110613 ***
Private Sub CmbCondicion_Click()
    Call CmbCondicion_Change
End Sub

Private Sub CmbCondicion_Change()
    Call DefineCondicionCredito
End Sub

Private Sub DefineCondicionCredito()
Dim oCred As COMDCredito.DCOMCredito
Dim nValor As Integer

    Set oCred = New COMDCredito.DCOMCredito
    'nValor = oCred.DefineCondicionCredito(oPersona.PersCodigo, , gdFecSis, bRefinanc)
    'nValor = oCred.DefineCondicionCredito(oPersona.PersCodigo, , gdFecSis, bRefinanc, val(SpnCuotasApr.valor)) 'EJVG20130509
    nValor = oCred.DefineCondicionCredito(oPersona.PersCodigo, , gdFecSis, bRefinanc, 0) 'WIOR 20141210
    Set oCred = Nothing
    If nCondCred = 5 Then
        CmbCondicion.ListIndex = IndiceListaCombo(CmbCondicion, 5)
    Else
        CmbCondicion.ListIndex = IndiceListaCombo(CmbCondicion, Trim(str(nValor)))
    End If
End Sub
'***

Private Sub cmbCondicion_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        ' CMACICA_CSTS - 05/11/2003 -----------------------------------------------------------------------
        CmbCondicionOtra.SetFocus
        ' -------------------------------------------------------------------------------------------------
     End If
End Sub

' CMACICA_CSTS - 05/11/2003 -----------------------------------------------------------------------
Private Sub cmbCondicionOtra_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        CmbApoderado.SetFocus
     End If
End Sub

Private Sub cmbDesBcoNac_Click()
     frmCredDesembBcoNac.Show 1
End Sub

Private Sub CmbDestino_Click()
    'By Capi 10042008
    'ALPA 20110301
        'If CmbDestino.ListIndex = 5 Then
        'FRHU20140212-RQ14005 Se Agrego lsTpoProdCod = "706"
        'ALPA 20150701
        'If CmbDestino.ListIndex = 5 Or (CmbDestino.ListIndex = 6 And (lsTpoProdCod = "510" Or lsTpoProdCod = "511" Or lsTpoProdCod = "706")) Then
        If Trim(Right(CmbDestino.Text, 5)) = 10 Or (Trim(Right(CmbDestino.Text, 5)) = 11 And (lsTpoProdCod = "510" Or lsTpoProdCod = "511" Or lsTpoProdCod = "706")) Then
            frmCtaAho.Enabled = True
            CmdCalend.Enabled = False
    'AMDO 20130712  TI-ERS077-2013**
    'ALPA 20150701
        'ElseIf CmbDestino.ListIndex = 9 Then
        ElseIf Trim(Right(CmbDestino.Text, 5)) = 14 Then
            txtDiasVoucher.Visible = True
            lblDiasVoucher.Visible = True
    'END AMDO ***
        Else
            frmCtaAho.Enabled = False
            actxCtaAho.Cuenta = ""
            lblTercero = ""
            CmdCalend.Enabled = True
    'AMDO 20130712  TI-ERS077-2013**
            txtDiasVoucher.Visible = False
            lblDiasVoucher.Visible = False
    'END AMDO ***
        End If
End Sub

' -------------------------------------------------------------------------------------------------
Private Sub cmbDestino_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        
        Me.CmbCondicion.SetFocus
      End If
End Sub

Private Sub cmbFuentes_Change()
Dim oCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim cNumFuente As String
Dim nPersTipFte As Integer
Dim dPersFecEval As Date

If cmbFuentes.Text = "" Then Exit Sub
Set oCred = New COMDCredito.DCOMCredito
cNumFuente = Trim(Mid(cmbFuentes.Text, Len(cmbFuentes.Text) - 150, 50))
nPersTipFte = CInt(Trim(Mid(cmbFuentes.Text, Len(cmbFuentes.Text) - 100, 50)))
dPersFecEval = CDate(Trim(Right(cmbFuentes.Text, 50)))

'*** PEAC 20080704
'Set rs = oCred.RecuperaDatosFuenteingreso(cNumFuente, nPersTipFte, Format(dPersFecEval, "dd/mm/yyyy"))
Set rs = oCred.RecuperaDatosFuenteingreso(cNumFuente, nPersTipFte, dPersFecEval)
Set oCred = Nothing

'*** PEAC 20080704
'    Call CargarDatosFuenteIngreso(rs, nPersTipFte)
If Not (rs.EOF And rs.BOF) Then
    Call CargarDatosFuenteIngreso(rs, nPersTipFte)
End If

Set rs = Nothing
End Sub

Private Sub cmbFuentes_Click()
    Call cmbFuentes_Change
End Sub

Private Sub cmdBuscar_Click()
    cmdCancelar_Click 'EJVG20160428->Limpiar pantalla
    SpnPlazo.valor = "0" 'MAVM 25102010
    txtFechaFija = "__/__/____" 'MAVM 25102010
            
    'ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSug), "Creditos para Aprobacion", , , False, gsCodAge, bLeasing)
    ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSug), "Creditos para Aprobacion", , , False, gsCodAge, bLeasing, , , , gsCodCargo)
    If ActxCta.NroCuenta <> "" Then
    Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
    If objProducto.GetResultadoCondicionCatalogo("N0000075", ActxCta.Prod) Then
        'ALPA 20110525 *LEASING
        'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            Dim oDLeasing As COMDCredito.DCOMleasing
            Dim oRS As ADODB.Recordset
            Dim nNroProveedores As Integer
            Set oDLeasing = New COMDCredito.DCOMleasing
            Set oRS = New ADODB.Recordset
            Set oRS = oDLeasing.ObtenerAprobacionLeasing(ActxCta.NroCuenta, nNroProveedores)
            Set oRS = Nothing
            nNroProveedores = 1
            If nNroProveedores > 0 Then
                Call ActxCta_KeyPress(13)
            Else
                ActxCta.CMAC = gsCodCMAC
                ActxCta.Age = gsCodAge
                ActxCta.SetFocusProd
                ActxCta.Enabled = True
                MsgBox "El Crédito Leasing debe ser aprobado en el SAF", vbInformation, "AVISO!"
            End If
        Else
            Call ActxCta_KeyPress(13)
        End If
        
    Else
        ActxCta.CMAC = gsCodCMAC
        ActxCta.Age = gsCodAge
        ActxCta.SetFocusProd
        ActxCta.Enabled = True
    End If
End Sub

Private Sub cmdCalend_Click()
Dim oNCred As COMNCredito.NCOMCredito
Dim i As Integer
Dim CadTmp As String
Dim MatCalTmp() As String
Dim lscCtaCod As String
Dim bErrorCargaCalendario As Boolean
Dim lsResAproba As String
'ALPA 20110525***************************************************
Dim lsCtaCodLeasing As String

nIntGraInicial = 0 'MAVM 20130209

Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
    If objProducto.GetResultadoCondicionCatalogo("N0000076", ActxCta.Prod) Then
    'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
        lsCtaCodLeasing = ActxCta.GetCuenta
    End If
    '****************************************************************
    'ALPA 20140612******
     If (txtMontoMivivienda.Text = 0# Or txtMontoMivivienda.Text = "") And lsTpoProdCod = "801" Then
         MsgBox "No se olvide de asignar el valor de venta del credito MIVIVIENDA", vbInformation, "Aviso"
         SSTab1.Tab = 6
         txtMontoMivivienda.SetFocus
         Exit Sub
     End If
      If Round(CDbl(txtMontoMivivienda.Text) - ((CDbl(txtMontoMivivienda.Text) * 1) / 10), 2) < CDbl(TxtMontoApr.Text) And lsTpoProdCod = "801" Then
        If Not (CDbl(txtMontoMivivienda.Text) = CDbl(TxtMontoApr.Text) And lsTpoProdCod = "801") Then
            MsgBox "MIVIVIENDA, no se olvide que el monto de la inicial no debe ser menor al 10% (" & Round((CDbl(TxtMontoApr.Text) * 10) / 9, 2) & ")", vbInformation, "Aviso"
            SSTab1.Tab = 6
            txtMontoMivivienda.SetFocus
            Exit Sub
        End If
     End If
     '****************************************************************
    
    'WIOR 20131111 **************************
    Dim lnCuotaBalon As Integer
    lnCuotaBalon = 0
    
    'MADM 20110505
        If (DateDiff("d", CDate(TxtFecDesemb.Text), CDate(txtFechaFija.Text)) = 0) And Len(Trim(lsCtaCodLeasing)) = 0 Then
            MsgBox "La Fecha de Pago es la Misma que la Fecha de Desembolso", vbInformation, "Aviso"
            If gnAgenciaCredEval = 0 Then txtFechaFija.SetFocus 'JUEZ 20121218
            Exit Sub
        End If
    'END MADM
    
    'JOEP 201710 ACTA201
    Dim rsValidaPriFecPago As ADODB.Recordset
    Dim obDCred As COMDCredito.DCOMCredito
    Set obDCred = New COMDCredito.DCOMCredito
    Set rsValidaPriFecPago = obDCred.ValidaPriFecPago(CDate(TxtFecDesemb.Text), CDate(txtFechaFija.Text))
    
    If Not (rsValidaPriFecPago.EOF And rsValidaPriFecPago.BOF) Then
        If rsValidaPriFecPago!cMensaje <> "" Then
            MsgBox rsValidaPriFecPago!cMensaje, vbInformation, "No podrá continuar"
            rsValidaPriFecPago.Close
            Set obDCred = Nothing
            Exit Sub
        End If
    rsValidaPriFecPago.Close
    Set obDCred = Nothing
    End If
    'JOEP 201710 ACTA201

    lscCtaCod = IIf(Len(ActxCta.NroCuenta) = 18, ActxCta.NroCuenta, "")
    If Not ValidaDatos Then
        Exit Sub
    End If
    
    'Valida que se hallan Generado los Diferentes Calendarios y Tipos de Opciones Minimas Requeridas
    'If optTipoGracia(0).value Then vnTipoGracia = gColocTiposGraciaCapitalizada 'LUCV20180601, Comentó según ERS022-2018
    'If optTipoGracia(1).value Then vnTipoGracia = gColocTiposGraciaEnCuotas    'LUCV20180601, Comentó según ERS022-2018
    If optTipoGracia(1).value Then vnTipoGracia = gColocTiposGraciaPrimeraCuota 'LUCV20180601, Agregó según ERS022-2018
    If CInt(TxtDiasGracia.Text) = "0" Then vnTipoGracia = -1
    
    If Trim(lsCtaCodLeasing) = "" Then 'No se usa la validacion para leasing
        If CInt(TxtDiasGracia.Text) > 0 And vnTipoGracia <> gColocTiposGraciaCapitalizada And vnTipoGracia <> gColocTiposGraciaEnCuotas Then
            If gnAgenciaCredEval = 0 Then 'JUEZ 20121218
                If Not bGraciaGen Then
                    MsgBox "Genere los Tipos de Gracia", vbInformation, "Aviso"
                    Exit Sub
                End If
            End If
        End If
    End If
    If OptDesemb(1).value Then
        If Not bDesembParcialGenerado Then
            MsgBox "Defina los Desembolsos Parciales", vbInformation, "Aviso"
            CmdDesembolsos.SetFocus
            Exit Sub
        End If
    End If
    
    If Not OptTCuota(3).value Then
        If Me.OptDesemb(0).value Then
            ReDim MatDesemb(1, 2)
            MatDesemb(0, 0) = TxtFecDesemb.Text
            MatDesemb(0, 1) = Format(TxtMontoApr.Text, "#0.00")
        Else
            
        End If

    'ALPA 20150131***************************************************************
    'If chkTasa.value = 0 Then
    If Trim(lblLineaDesc) = "" Then
        MsgBox "No se ha seleccionado ninguna línea de crédito, ingresar los parametros correctos", vbInformation, "Aviso!"
        Exit Sub
    End If
    If lnTasaInicial = 0 Or lnTasaFinal = 0 Then
        MsgBox "Debe ingresar monto/tasa de Crédito válido para el Producto" & Chr(10) & Chr(10) & "Ver la configuración del Producto", vbInformation, "Aviso!"
        If TxtMontoApr.Enabled = True Then
            TxtMontoApr.SetFocus
        End If
        Exit Sub
    End If
    If txtTasInt.Text = 0 Or txtTasInt.Text = "" Then
        MsgBox "Debe ingresar valor a la tasa de interes compensatorio"
        If txtTasInt.Enabled = True Then
        txtTasInt.SetFocus
        End If
        Exit Sub
    End If
    
    lsResAproba = ObtenerColocacPermisoAprobacion
    If CDbl(txtTasInt.Text) < CDbl(lnTasaInicial) Or CDbl(txtTasInt.Text) > CDbl(lnTasaFinal) Then
        If CDbl(txtTasInt.Text) < CDbl(lnTasaInicial) Then
            
            'Dim lsResAproba As String
            If Left(lsResAproba, 1) = "0" Then
                MsgBox "La tasa de interes debe estar entre [" & lnTasaInicial & "-" & lnTasaFinal & "]"
                Dim objCargos As COMDCredito.DCOMNivelAprobacion
                Dim ClsMov As COMNContabilidad.NCOMContFunciones
                Set objCargos = New COMDCredito.DCOMNivelAprobacion
                Set ClsMov = New COMNContabilidad.NCOMContFunciones
                Dim sMovNroS As String
                sMovNroS = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Call objCargos.GetActualizarColocacPermisoAprobacion(ActxCta.NroCuenta, sMovNroS, "", gsCodPersUser, "", "", txtTasInt.Text, 0, 1, 1)
                MsgBox "La petición de tasa ha sido registrada correctamente, coordinar con el area respectiva para aprobar la petición", vbInformation, "Aviso!"
                Exit Sub
            ElseIf Left(lsResAproba, 1) = "1" Then
                MsgBox "Ya existe el pedido de autorización de tasa" & Chr(10) & Chr(10) & Mid(lsResAproba, 3, Len(lsResAproba) - 2), vbInformation, "Aviso!"
                Exit Sub
            ElseIf Left(lsResAproba, 1) = "2" Then
                MsgBox Mid(lsResAproba, 3, Len(lsResAproba) - 2), vbInformation, "Aviso!"
            ElseIf Left(lsResAproba, 1) = "3" Then
                If txtTasInt.Text <> Val(Trim(Right(lsResAproba, 8))) Then
                     If txtTasInt.Enabled = False Then
                        MsgBox "Debido a la exoneración de tasa se actualizará a " & Val(Trim(Right(lsResAproba, 8))) & "%", vbInformation, "Aviso!"
                        txtTasInt.Text = Val(Trim(Right(lsResAproba, 8)))
                     Else
                        MsgBox Mid(lsResAproba, 3, Len(lsResAproba) - 2), vbInformation, "Aviso!"
                        Exit Sub
                     End If
                End If
            End If
        Else
            If CDbl(txtTasInt.Text) > CDbl(lnTasaFinal) Then
                MsgBox "La tasa está fuera de rango", vbInformation, "Aviso!"
                If txtTasInt.Enabled = True Then
                    txtTasInt.SetFocus
                End If
                Exit Sub
            End If
        End If
    Else
        If Left(lsResAproba, 1) = "1" Then
            If MsgBox("Ya existe el pedido de autorización de tasa " & Chr(10) & Chr(10) & "Desea eliminar el pedido y aprobar con la tasa indicada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                Dim sMovNroA As String
                Set objCargos = New COMDCredito.DCOMNivelAprobacion
                Set ClsMov = New COMNContabilidad.NCOMContFunciones
                sMovNroA = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Call objCargos.GetActualizarColocacPermisoAprobacion(ActxCta.NroCuenta, "", sMovNroA, "", gsCodPersUser, gsCodCargo, 0, 0, 4, 4)
                MsgBox "La petición de tasa se eliminó correctamente ", vbInformation, "Aviso!"
            Else
                Exit Sub
            End If
        End If
    End If
    
    Call ObtenerTasaSeguroDesg(False) 'LUCV20180601, Agregó Según ERS022-2018
    
    'LUCV20180601, Según ERS022-2018
    Dim nMontoPoliza As Double
    Dim nTasaSegInc As Double
    Dim oNGasto As COMNCredito.NCOMGasto
    Set oNGasto = New COMNCredito.NCOMGasto
    nMontoPoliza = 0
    
    nMontoPoliza = oNGasto.RecuperaMontoPoliza(lscCtaCod, _
                                               CInt(SpnCuotasApr.valor), _
                                               gColocConceptoCodGastoPolizaIncendioHipoteca, _
                                               nTasaSegInc, _
                                               IIf(Trim(Right(CmbCondicion.Text, 5)) = "3", True, False))
    
    'Fin LUCV20180601
    
    MatCalend = frmCredCalendPagos.Inicio(CDbl(TxtMontoApr.Text), txtTasInt.Text, CInt(SpnCuotasApr.valor), _
                                    IIf(CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)) > 0, 30, CInt(SpnPlazo.valor)), _
                                    CDate(MatDesemb(0, 0)), DameTipoCuota, IIf(OptTPer(0).value, 1, 2), vnTipoGracia, _
                                    CInt(TxtDiasGracia.Text), TxtTasaGracia.Text, CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                                    ChkProxMes.value, MatGracia, ChkMiViv.value, ChkCuotaCom.value, MatCalend_2, 2, True, _
                                    Me.OptDesemb(1).value, MatDesemb, bQuincenal, lscCtaCod, bErrorCargaCalendario, _
                                    CInt(TxtDiaFijo2.Text), , IIf(TxtDiasGracia.Text <> "0", True, False), _
                                    txtFechaFija.Text, , lsCtaCodLeasing, , nIntGraInicial, lnCuotaBalon, , , CDbl(txtMontoMivivienda.Text), _
                                    lnCSP, fArrMIVIVIENDA, fnTasaSegDes, MatCalendSegDes, ChkExoneracion.value, nMontoPoliza, nTasaSegInc)
                                    'LUCV20180601, Agregró:fnTasaSegDes, MatCalendSegDes, ChkExoneracion.value, nMontoPoliza, nTasaSegInc según ERS022-2018
        'Para evitar la caida del Sistema cuando se encuentre Errores en la Carga del Calendario
        If bErrorCargaCalendario Then Exit Sub
        
        bCalendGen = True
         'WIOR 20120727 *************************************************************
        Dim oCreditoInf As COMDCredito.DCOMCredito
        Dim rsCreditoInf As ADODB.Recordset
        Dim nEstInforme As Integer
        Set oCreditoInf = New COMDCredito.DCOMCredito
        Set rsCreditoInf = oCreditoInf.ObtenerInformeRiesgo(Trim(Me.ActxCta.NroCuenta))
        
        If rsCreditoInf.RecordCount > 0 Then
            nEstInforme = CInt(IIf(IsNull(rsCreditoInf!nEstado), 0, rsCreditoInf!nEstado))
            If nEstInforme = 2 Then
                Call frmCredRiesgosInformeVer.Inicio(Trim(ActxCta.NroCuenta))
            End If
        End If
        'WIOR FIN ******************************************************************

        'Gastos del Calendario
        If Len(Trim(lsCtaCodLeasing)) = 0 Then
            Call CmdGastos_Click
        End If
        
        
        
        For i = 0 To UBound(MatCalend) - 1
            'MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 8)), "#0.00") 'LUCV20180601. Comentó, según ERS0222018
            MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) + CDbl(MatCalend(i, 8)), "#0.00") 'LUCV20180601. Agregó según ERS022-2018
            MatCalend(i, 4) = Format(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 11)), "#0.00")
            'ALPA 20141231*****************************
            If MatCalend(i, 1) = 1 Then
              lnCuotaAprobada = MatCalend(i, 2)
            End If
            '******************************************
        Next i
        Call ObtenerTasaSeguroDesg(True) 'WIOR 20160224
    
        '*****->LUCV20170915, Creó. Según Modificaciones del reglamento (4.2.-Crédito Refinanciado. (b))
        If Not ValidaPeriodoGracia Then
            txtFechaFija.SetFocus
            Exit Sub
        End If
        'Fin LUCV20170915 <-*****
        
        Call frmCredCalendPagos.SoloMuestraMatrices(MatCalend, MatCalend_2, MatGastos, nNumGastos, _
                            CDbl(TxtMontoApr.Text), txtTasInt.Text, CInt(SpnCuotasApr.valor), _
                            CInt(SpnPlazo.valor), CDate(MatDesemb(0, 0)), DameTipoCuota, _
                            IIf(OptTPer(0).value, 1, 2), vnTipoGracia, CInt(TxtDiasGracia.Text), _
                            TxtTasaGracia.Text, CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                            ChkProxMes.value, MatGracia, ChkMiViv.value, ChkCuotaCom.value, MatCalend_2, 2, True, _
                            CInt(IIf(Trim(TxtDiaFijo2.Text) = "", "00", TxtDiaFijo2.Text)), True, _
                            nTasCosEfeAnu, lsCtaCodLeasing) 'DAOR 20070419, nTasCosEfeAnua:Recibir la tasa de costo efectivo anual
                            'nTasCosEfeAnu) 'DAOR 20070419, nTasCosEfeAnua:Recibir la tasa de costo efectivo anual

        bCalendGen = True
        If UBound(MatCalend) <> 0 Then
           nMontoCuota = MatCalend(0, 2)
        End If
        
    Else
        Set oNCred = New COMNCredito.NCOMCredito
        'MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(ActxCta.NroCuenta)
        MatCalend = oNCred.RecuperaMatrizCalendarioInicial(ActxCta.NroCuenta, CDbl(TxtMontoApr.Text))
        Set oNCred = Nothing
        ReDim MatCalTmp(UBound(MatCalend), 6)
        For i = 0 To UBound(MatCalTmp) - 1
            MatCalTmp(i, 0) = MatCalend(i, 1)
            MatCalTmp(i, 1) = MatCalend(i, 0)
            MatCalTmp(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)), "#0.00")
            MatCalTmp(i, 3) = MatCalend(i, 3)
            MatCalTmp(i, 4) = MatCalend(i, 4)
            MatCalTmp(i, 5) = MatCalend(i, 10)
        Next i
                
        If OptDesemb(1).value Then
            bCalendGen = True
        Else
            ReDim MatDesemb(1, 2)
            MatDesemb(0, 0) = TxtFecDesemb.Text
            MatDesemb(0, 1) = Format(TxtMontoApr.Text, "#0.00")
            bCalendGen = True
        End If
        
        ReDim MatCalTmp(UBound(MatCalend), 6)
        For i = 0 To UBound(MatCalend) - 1
            MatCalTmp(i, 0) = MatCalend(i, 1)
            MatCalTmp(i, 1) = MatCalend(i, 0)
            MatCalTmp(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)), "#0.00")
            MatCalTmp(i, 3) = MatCalend(i, 3)
            MatCalTmp(i, 4) = MatCalend(i, 4)
        Next i
        MatCalend = MatCalTmp
        
    End If
    If nAgenciaCredEval = 0 Then
        cmdEvaluacion.Enabled = True
    End If
End Sub

Sub MostrarGastos(ByVal pMatGastos As Variant, ByVal pnNumGastos As Integer)
    MatGastos = pMatGastos
    nNumGastos = pnNumGastos
    bGastosGen = True
    Call frmCredGastos.MuestraGastos(MatGastos, nNumGastos)
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarPantalla
End Sub
'RECO20150421 ERS010-2015 ******************************
Private Sub cmdCheckList_Click()
    'If frmAdmCheckListDocument.Inicio(ActxCta.NroCuenta, lsTpoCredCod, nRegAprobacion) = True Then
    If frmAdmCheckListDocument.Inicio(ActxCta.NroCuenta, lsCategoria, lsTpoProdCod, CCur(TxtMontoApr.Text), lsTipoCredito, nRegAprobacion) = True Then  'JOEP20181229 CP
        bCheckList = True
    Else
        bCheckList = False
    End If
End Sub
'RECO FIN **********************************************

Private Sub CmdCredVig_Click()
    MatCredVig = frmCredVigentes.Inicio(LblCodCli.Caption, LblNomCli.Caption, ActxCta.NroCuenta, MatCredVig)
End Sub

'By Capi 10042008
Private Sub cmdCtaAho_Click()
Dim clsPers As COMDPersona.UCOMPersona 'UPersona
Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
  
    sPers = clsPers.sPersCod
    'By Capi 10042008
    sCliCtaAho = clsPers.sPersNombre
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    
    'By Capi 10042008
    If sPers = gsCodPersUser Then
        MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    MsgBox "Solo se Considera Cuentas Activas y Afectas a ITF", vbInformation, "Aviso"
    '
    Set rsPers = clsCap.GetCuentasPersona(sPers, gCapAhorros, , , , , gsCodAge)
    Set clsCap = Nothing
    'By Capi 10042008
    Dim lbTieneCuentaValida As Boolean
    lbTieneCuentaValida = False
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            'By Capi 10042008
            If rsPers("Exonerada") = "" And sEstado = "ACTIVA" Then
                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
                lbTieneCuentaValida = True
            End If
            rsPers.MoveNext
        Loop
        'By Capi 15042008
        If lbTieneCuentaValida = True Then
            Set clsCuenta = frmCapMantenimientoCtas.inicia
       
            If clsCuenta.sCtaCod <> "" Then
                Me.btnVer.Enabled = True 'FRHU20140212 RQ14005
                actxCtaAho.CMAC = Mid(clsCuenta.sCtaCod, 1, 3) 'FRHU20140212 RQ14005
                actxCtaAho.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                actxCtaAho.Prod = Mid(clsCuenta.sCtaCod, 6, 3) 'FRHU20140212 RQ14005
                actxCtaAho.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                lblTercero.Caption = sCliCtaAho
                actxCtaAho.SetFocusCuenta
                SendKeys "{Enter}"
            End If
            Set clsCuenta = Nothing
        Else
            MsgBox "Persona no posee cuentas validas para depositos.", vbInformation, "Aviso"
        End If
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
Set clsPers = Nothing
actxCtaAho.SetFocusCuenta
End Sub

Private Sub CmdDesembolsos_Click()
Dim nSumaDesPar As Double
Dim i As Integer
Dim MonDesAnt As Double
    
    MonDesAnt = CDbl(TxtMontoApr.Text)
    MatDesemb = frmCredDesembParcial.Inicio(CDate(TxtFecDesemb.Text), MatDesemb)
    If UBound(MatDesemb) > 0 Then
        nSumaDesPar = 0
        For i = 0 To UBound(MatDesemb) - 1
            nSumaDesPar = nSumaDesPar + CDbl(MatDesemb(i, 1))
        Next i
        TxtMontoApr.Text = Format(nSumaDesPar, "#0.00")
        If MonDesAnt <> nSumaDesPar Then
            ReDim MatCalend(0, 0)
        End If
        bDesembParcialGenerado = True
    Else
        bDesembParcialGenerado = False
        ReDim MatCalend(0, 0)
    End If
End Sub
'APRI20180404 ERS036-2017
Private Sub cmdEnvioEstCta_Click()
    frmEnvioEstadoCta.InicioCol ActxCta.NroCuenta, LlenaRecordSet_Cliente, True
    fbRegistraEnvio = frmEnvioEstadoCta.RegistraEnvio
    If fbRegistraEnvio Then
        bCalendGen = False
    End If
End Sub
Private Function LlenaRecordSet_Cliente() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim i As Integer
Set rs = New ADODB.Recordset
    
    With rs
        .Fields.Append "codigo", adVarChar, 13
        .Fields.Append "relacion", adInteger
        .Open
        
       For i = 1 To UBound(Clientes)
            If Clientes(i - 1, 0) <> "" Then
            .AddNew
            .Fields("codigo") = Clientes(i - 1, 0)
            .Fields("relacion") = Clientes(i - 1, 1)
            End If
        Next i
        
    End With
    
    Set LlenaRecordSet_Cliente = rs
End Function
'END APRI

Private Sub cmdEvaluacion_Click()
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
Dim nEstado As Integer
Dim rs As ADODB.Recordset
Dim cPrd As String
Dim cSPrd As String
Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
Dim fnMontoExpEsteCred_NEW As Currency
    
    If Len(ActxCta.NroCuenta) <> "18" Then
        Exit Sub
    End If

    Set rs = oDCOMFormatosEval.RecuperaCredFormEvalProductoEstadoExposicion(ActxCta.NroCuenta)
    fnMontoExpEsteCred_NEW = IIf(IsNull(rs!nMontoExpCredito), 0, rs!nMontoExpCredito)
    nEstado = IIf(IsNull(rs!nPrdEstado), 0, rs!nPrdEstado)
    
    If nEstado = 0 Then
        MsgBox "Nº de Credito no tiene formato registrado", vbInformation, "Aviso"
        Exit Sub
    Else
        If nEstado > 2000 Then
            If rs.RecordCount > 0 Then
                cSPrd = Trim(rs!cTpoProdCod)
                cPrd = Mid(cSPrd, 1, 1) & "00"
            End If
        EvaluarCredito ActxCta.NroCuenta, False, nEstado, cPrd, cSPrd, fnMontoExpEsteCred_NEW, False, True
        Else
            MsgBox "Nº de Crédito no se encuentra en estado Solicitado.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
End Sub

Private Sub CmdGarantia_Click()
    Call frmCredGarantiasCob.Inicio(ActxCta.NroCuenta)
End Sub

Private Sub CmdGastos_Click()
Dim oNGasto As COMNCredito.NCOMGasto
Dim bErrorCargaCalendario As Boolean 'ARCV

'MAVM 20120802 ***
Dim nTramoNoConsMonto As Double
Dim nTramoConsMonto As Double
Dim nTramoNoConsPorcen As Double
Dim nPlazoMiViv As Integer
Dim nPlazoMiVivMax As Integer
'***

    Dim MatCalendIC As Variant
    'WIOR 20131111 **************************
    Dim vArrDatos As Variant 'LUCV20180601, Según ERS022-2018
    On Error GoTo ErrorCmdGastos_Click
    If Not bCalendGen Then
        MsgBox "Genere el Calendario de Pagos para poder Definir los Gastos", vbInformation, "Aviso"
        CmdCalend.SetFocus
        Exit Sub
    End If
    Set oNGasto = New COMNCredito.NCOMGasto
    
    '->***** LUCV20180601, Agregó Según ERS022-2018
    ReDim vArrDatos(2)
    vArrDatos(0) = 0 'Valor del Parámetro pnMontoMivivienda
    vArrDatos(1) = MatCalend
    '<-***** Fin LUCV20180601
        
    MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMontoApr.Text), txtTasInt.Text, CInt(SpnCuotasApr.valor), _
                            IIf(CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)) > 0, 30, CInt(SpnPlazo.valor)), CDate(MatDesemb(0, 0)), DameTipoCuota, _
                            IIf(OptTPer(0).value, 1, 2), vnTipoGracia, CInt(TxtDiasGracia.Text), _
                            TxtTasaGracia.Text, CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                            ChkProxMes.value, MatGracia, 0, ChkCuotaCom.value, MatCalend_2, _
                            MatDesemb, nNumGastos, gdFecSis, _
                            ActxCta.NroCuenta, 1, "DE", IIf(OptTipoGasto(0).value, "F", "V"), _
                            CDbl(MatCalend(0, 2)), CDbl(TxtMontoApr.Text), , , , , , , , True, _
                            2, True, Me.OptDesemb(1).value, MatDesemb, bQuincenal, bErrorCargaCalendario, _
                            CInt(TxtDiaFijo2.Text), True, , _
                             lnCSP, gnITFPorcent, gbITFAplica, ChkExoneracion.value, Mid(ActxCta.NroCuenta, 4, 2), lsTpoProdCod, lsTpoCredCod, , , nIntGraInicial, MatCalendIC, , chkExoneraMYPE.value, ckPolizaBF.value, txtFechaPBF.Text, MatCalendSegDes, vArrDatos) 'Solo Mostrar
                             'WIOR 20131111 AGREGO fnCuotaBalon
                             'WIOR 20140825 AGREGO MatCalendSegDes
                             'ALPA 20141127 lnCSP-gnITFMontoMin
                             'WIOR 20160112 - REEMPLAZO ChkMiViv.value POR 0
                             'WIOR 20160112 - REEMPLAZO IIf(ChkMiViv.value, nTramoNoConsMonto, CDbl(TxtMontoApr.Text)) POR CDbl(TxtMontoApr.Text)
                             'APRI20180821 ERS061-2018 REEMPLAZÓ lnCuotaBalon POR chkExoneraMYPE.Value
                             'LUCV20180601, Agregó vArrDatos por limite de parámetros
                             
    'MAVM 20130209 ***
    If vnTipoGracia = 6 Then
        MatCalend = MatCalendIC
    End If
    '***

    bGastosGen = True
    Set oNGasto = Nothing
    
    Call frmCredGastos.MuestraGastos(MatGastos, nNumGastos, True)
    Exit Sub

ErrorCmdGastos_Click:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'madm 20110303
Public Function ValidaEnListaNegativa() As Boolean
    Dim loVistoElectronico As SICMACT.frmVistoElectronico
    Dim lnCondicion As Integer
    Dim lcCondi As String
    Dim lbResultadoVisto As Boolean
    Set loVistoElectronico = New SICMACT.frmVistoElectronico
    Set objPersona = New COMDPersona.UCOMPersona
    
    lbResultadoVisto = False
    ValidaEnListaNegativa = False

         If objPersona.ValidaEnListaNegativaCondicion(lsTpoIDNroDNI, lsTpoIDNroRUC, lnCondicion, LblNomCli.Caption) Then
            Select Case lnCondicion
               Case 1
                   lcCondi = "NEGATIVO"
               Case 2
                   lcCondi = "FRAUDULENTO"
               Case 3
                   lcCondi = "PEPS"
               Case 5
                   lcCondi = "LISTA OFAC"
               Case 6
                   lcCondi = "LISTA ONU"
            End Select

            If lnCondicion = 1 Or lnCondicion = 3 Or lnCondicion = 5 Or lnCondicion = 6 Then 'madm 20100510
                'MsgBox "Este Cliente esta en la Lista de Negativos como " & lcCondi & ", necesitará un Visto electrónico", vbInformation, "Aviso"'JGPA20200117
                MsgBox "Este Cliente esta en la Lista de Negativos como " & lcCondi, vbInformation, "Aviso" 'JGPA20200117 anexo acta 106-2019
                ValidaEnListaNegativa = False 'JGPA 20191227 Cambió [True] x [False]
                '***JGPA20200117 Comentó
                'lbResultadoVisto = loVistoElectronico.Inicio(3)
                    'If Not lbResultadoVisto Then
                             'ValidaEnListaNegativa = True
                             'Exit Function
                    'Else
                        'ValidaEnListaNegativa = False
                    'End If
                'Set oPersona = Nothing
                '***End JGPA20200117
                Set objPersona = Nothing 'EJVG20121023 X Errores en PEPS
           End If
       End If
End Function
'end madm

Private Sub cmdGrabar_Click()
Dim oNCredito As COMNCredito.NCOMCredito
Dim pnTipoCuota As Integer
Dim sError As String
'Dim sMontosARef As Variant
Dim nTasa As Currency 'ALPA 20150115*********
'Variables para la Grabacion de la Aprobacion
Dim rsDR As ADODB.Recordset
'Dim bAmpliado As Boolean
'Dim bQuincenal As Boolean

'*** PEAC 20080811
Dim lbResultadoVisto As Boolean
Dim sPersVistoCod  As String
Dim sPersVistoCom As String
Dim pnMovMro As Long
Dim loVistoElectronico As SICMACT.frmVistoElectronico
Set loVistoElectronico = New SICMACT.frmVistoElectronico
'WIOR 20140123 *********************
Dim bTrabajadoVinc As Boolean
Dim nTipoVinculado As Integer
bTrabajadoVinc = False
nTipoVinculado = 0
'WIOR FIN **************************
'FRHU 20170915 ERS049-2017
Dim sMovNroM As String
Dim clsMovM As New COMNContabilidad.NCOMContFunciones
sMovNroM = clsMovM.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'FIN FRHU 20170915
'RECO20150421 ERS010-2015 **************************************
Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
If cmdCheckList.Enabled = True And bCheckList = False Then
'If bCheckList = False And ActxCta.Prod <> "703" Then
    MsgBox "Debe registrar el CheckList", vbInformation, "Alerta"
    Exit Sub
End If
'RECO FIN ******************************************************

'NRLO 20180306 Visita Domiciliaria
Dim rsValidaVisita As ADODB.Recordset
Dim obDCredVD As COMDCredito.DCOMCredito
Set obDCredVD = New COMDCredito.DCOMCredito
Set rsValidaVisita = obDCredVD.ValidarVisitaDomiciliaria(ActxCta.NroCuenta)
If Not (rsValidaVisita.EOF And rsValidaVisita.BOF) Then
    If rsValidaVisita!nEstado <> 1 Then
        MsgBox rsValidaVisita!cMensaje, vbInformation, "Aviso"
        If rsValidaVisita!nEstado = 0 Then
            rsValidaVisita.Close
            Set obDCredVD = Nothing
            Exit Sub
        End If
        rsValidaVisita.Close
        Set obDCredVD = Nothing
    End If
Set obDCredVD = Nothing
End If
'NRLO 20180306 Visita Domiciliaria END

'JOEP 201710 Lavado
Dim rsValidaPriFecPago As ADODB.Recordset
Dim obDCredFP As COMDCredito.DCOMCredito
Set obDCredFP = New COMDCredito.DCOMCredito
Set rsValidaPriFecPago = obDCredFP.ValidaPriFecPago(CDate(TxtFecDesemb.Text), CDate(txtFechaFija.Text))
If Not (rsValidaPriFecPago.EOF And rsValidaPriFecPago.BOF) Then
    If rsValidaPriFecPago!cMensaje <> "" Then
        MsgBox rsValidaPriFecPago!cMensaje, vbInformation, "Aviso"
        rsValidaPriFecPago.Close
        Set obDCredFP = Nothing
        Exit Sub
    End If
rsValidaPriFecPago.Close
Set obDCredFP = Nothing
End If
'JOEP 201710 Lavado

    'RIRO20170831 ****************
    If chkDesembBcoNac.value = 1 Or chkDesembBcoNac.value = 2 Then
        If MsgBox("El crédito será desembolsado en el banco de la nacion, por lo tanto una vez enviada la trama, " & vbNewLine & _
               "no podrá ser extornada durante el resto del día" & vbNewLine & _
               "¿desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            
            Exit Sub
            
        End If
    End If
    'END RIRO ********************
    
    'ALPA 20110525***************************************************
    Dim lsCtaCodLeasing As String
    Dim lsResAproba As String 'ALPA20150210
    Set objProducto = New COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
    If objProducto.GetResultadoCondicionCatalogo("N0000078", ActxCta.Prod) Then
    'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
        lsCtaCodLeasing = ActxCta.GetCuenta
    End If
    '****************************************************************
    lbExoneraTasa = 0 'ALPA20150210
    If Trim(lblLineaDesc) = "" Then
        MsgBox "No se ha seleccionado ninguna línea de crédito, ingresar los parametros correctos", vbInformation, "Aviso!"
        Exit Sub
    End If
    
    'If chkTasa.value = 0 Then
    If lnTasaInicial = 0 Or lnTasaFinal = 0 Then
        MsgBox "Debe ingresar monto/tasa de Crédito válido para el Producto"
        If TxtMontoApr.Enabled = True Then
        TxtMontoApr.SetFocus
        End If
        Exit Sub
    End If
    If txtTasInt.Text = 0 Or txtTasInt.Text = "" Then
        MsgBox "Debe ingresar valor a la tasa de interes compensatorio"
        If txtTasInt.Enabled = True Then
            txtTasInt.SetFocus
        End If
        Exit Sub
    End If
    'FRHU 20170914 ERS049-2017
    If gnAgenciaCredEval = 0 Then
        If Not (ValidarTasaMaxima(ActxCta.NroCuenta, CDbl(txtTasInt.Text), Right(Trim(CmbCampana.Text), 3), lsTpoProdCod, CCur(TxtMontoApr.Text))) Then Exit Sub
    End If
    'FIN FRHU 20170914 ERS049-2017
    lsResAproba = ObtenerColocacPermisoAprobacion
If CDbl(txtTasInt.Text) < CDbl(lnTasaInicial) Or CDbl(txtTasInt.Text) > CDbl(lnTasaFinal) Then
    If CDbl(txtTasInt.Text) < CDbl(lnTasaInicial) Then
        
        'Dim lsResAproba As String
        If Left(lsResAproba, 1) = "0" Then
            MsgBox "La tasa de interes debe estar entre [" & lnTasaInicial & "-" & lnTasaFinal & "]"
            Dim objCargos As COMDCredito.DCOMNivelAprobacion
            Dim ClsMov As COMNContabilidad.NCOMContFunciones
            Set objCargos = New COMDCredito.DCOMNivelAprobacion
            Set ClsMov = New COMNContabilidad.NCOMContFunciones
            Dim sMovNroS As String
            sMovNroS = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Call objCargos.GetActualizarColocacPermisoAprobacion(ActxCta.NroCuenta, sMovNroS, "", gsCodPersUser, "", "", txtTasInt.Text, 0, 1, 1)
            MsgBox "La petición de tasa ha sido registrada correctamente, coordinar con el area respectiva para aprobar la petición", vbInformation, "Aviso!"
            Exit Sub
        ElseIf Left(lsResAproba, 1) = "1" Then
            MsgBox "Ya existe el pedido de autorización de tasa" & Chr(10) & Chr(10) & Mid(lsResAproba, 3, Len(lsResAproba) - 2), vbInformation, "Aviso!"
            Exit Sub
        ElseIf Left(lsResAproba, 1) = "2" Then
            MsgBox Mid(lsResAproba, 3, Len(lsResAproba) - 2), vbInformation, "Aviso!"
        ElseIf Left(lsResAproba, 1) = "3" Then
            If txtTasInt.Text <> Val(Trim(Right(lsResAproba, 8))) Then
                 If txtTasInt.Enabled = False Then
                    txtTasInt.Text = Val(Trim(Right(lsResAproba, 8)))
                 Else
                    MsgBox Mid(lsResAproba, 3, Len(lsResAproba) - 2), vbInformation, "Aviso!"
                 End If
                'MsgBox Mid(lsResAproba, 3, Len(lsResAproba) - 2), vbInformation, "Aviso!"
                Exit Sub
            End If
        End If
    Else
        If CDbl(txtTasInt.Text) > CDbl(lnTasaFinal) Then
            MsgBox "La tasa está fuera de rango", vbInformation, "Aviso!"
            If txtTasInt.Enabled = True Then
                txtTasInt.SetFocus
            End If
            Exit Sub
        End If
    End If
Else
    If Left(lsResAproba, 1) = "1" Then
        If MsgBox("Ya existe el pedido de autorización de tasa " & Chr(10) & Chr(10) & "Desea eliminar el pedido y aprobar con la tasa indicada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Dim sMovNroA As String
            'Dim objCargos As COMDCredito.DCOMNivelAprobacion
            'Dim ClsMov As COMNContabilidad.NCOMContFunciones
            Set objCargos = New COMDCredito.DCOMNivelAprobacion
            Set ClsMov = New COMNContabilidad.NCOMContFunciones
            sMovNroA = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Call objCargos.GetActualizarColocacPermisoAprobacion(ActxCta.NroCuenta, "", sMovNroA, "", gsCodPersUser, gsCodCargo, 0, 0, 4, 4)
            MsgBox "La petición de tasa se elimino correctamente ", vbInformation, "Aviso!"
        Else
            Exit Sub
        End If
    End If
End If

    'FRHU 20140422 Observacion - RQ14005
    'El frmCtaAho se habilita o deshabilita solo en el procedimientio cmbDestino_Click()
    If frmCtaAho.Enabled = True Then
        Call actxCtaAho_KeyPress(13)
        If CmdCalend.Enabled = False Then
            Exit Sub
        End If
    End If
    'FIN FRHU 20140422
    If Not bCalendGen Then
        MsgBox "Genere el Calendario de Pagos para poder Definir los Gastos", vbInformation, "Aviso"
        CmdCalend.SetFocus
        Exit Sub
    End If

    If Not ValidaDatos(True) Then
        Exit Sub
    End If
                
    'APRI20171123 ERS028-2017
        Dim oSeguro As COMNCaptaGenerales.NCOMSeguros
        Set oSeguro = New COMNCaptaGenerales.NCOMSeguros
        Dim nTotal As Integer
        nTotal = oSeguro.ObservacionesCredSegDesg(ActxCta.NroCuenta)
        If nTotal > 0 Then
            
         MsgBox "El crédito no puede ser aprobado porque requiere exámenes obligatorios para la afiliación del Seg. Desgravamen." & _
            Chr(13) & Chr(13) & "Por favor comunicarse con el Coordinador de Seguros.", vbInformation, "Aviso"
            
         Exit Sub
        End If
        Set oSeguro = Nothing
    'END APRI
                
    '**ARLO20181211 ERS068-2018
    If bRefinanc Then
        Dim oCredRefin As COMDCredito.DCOMCredito
        Dim rsRefinan As ADODB.Recordset
        Set oCredRefin = New COMDCredito.DCOMCredito
        Set rsRefinan = oCredRefin.obtenerVistoBueno90Dias(Trim(Me.ActxCta.NroCuenta))
        If Not (rsRefinan.EOF And rsRefinan.BOF) Then
            If rsRefinan!sMensaje <> "" Then
                MsgBox rsRefinan!sMensaje, vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        Set rsRefinan = Nothing
        Set oCredRefin = Nothing
    End If
    '**ARLO END
                
    'JOEP ERS066-20161110 Valida si la Campana esta Activa
        Dim oObtenerCampanaAct As COMDCredito.DCOMCredito
        Dim rsCampAct As ADODB.Recordset
        Dim rsAsigCampAct As ADODB.Recordset
        
        Set oObtenerCampanaAct = New COMDCredito.DCOMCredito
        Set rsCampAct = New ADODB.Recordset
        Set rsAsigCampAct = New ADODB.Recordset
        
        
        If gnAgenciaCredEval = 1 Then 'Agencia Piloto
            If CmbCampana.Text = "" Then
                MsgBox "El Credito esta  vinculado a una Campaña no Activa, favor de Comunicarse con el Analista del Credito", vbInformation, "Aviso"
                Exit Sub
            End If
             
        Set rsCampAct = oObtenerCampanaAct.ObtenerCampActivaApro(Right(Trim(CmbCampana.Text), 3)) 'Valida si esta Activa Campana
            If Not (rsCampAct.BOF And rsCampAct.EOF) Then
                If (rsCampAct!bEstado) = False Then
                    MsgBox "El Credito esta  vinculado a una Campaña no Activa, favor de Comunicarse con el Analista del Credito", vbInformation, "Aviso"
                Exit Sub
                End If
            End If
            
        Set rsAsigCampAct = oObtenerCampanaAct.ObtenerAsigCampActiva(Right(Trim(CmbCampana.Text), 3), gsCodAge) 'Valida Asinacion de Campana/Agencia
            If Not (rsAsigCampAct.BOF And rsAsigCampAct.EOF) Then
            
            Else
                    MsgBox "La Campaña no esta Asignada a la Agencia ,favor de Comunicarse con el Analista del Credito", vbInformation, "Aviso"
                Exit Sub
            End If
            
        Else 'Agencia no Piloto
        
        Set rsCampAct = oObtenerCampanaAct.ObtenerCampActivaApro(Right(Trim(CmbCampana.Text), 3)) 'Valida si esta Activa Campana
            If Not (rsCampAct.BOF And rsCampAct.EOF) Then
                If (rsCampAct!bEstado) = False Then
                    MsgBox "El Credito esta  vinculado a una Campaña no Activa, favor de Comunicarse con el Analista del Credito", vbInformation, "Aviso"
                Exit Sub
                End If
            End If
            
        Set rsAsigCampAct = oObtenerCampanaAct.ObtenerAsigCampActiva(Right(Trim(CmbCampana.Text), 3), gsCodAge) 'Valida Asinacion de Campana/Agencia
            If Not (rsAsigCampAct.BOF And rsAsigCampAct.EOF) Then
            
            Else
                    MsgBox "La Campaña no esta Asignada a la Agencia ,favor de Comunicarse con el Analista del Credito", vbInformation, "Aviso"
                Exit Sub
            End If
        
        End If
    'JOEP ERS066-20161110
                
    If Trim(lsCtaCodLeasing) = "" Then 'ALPA 20110526 *No se usa para Leasing
        If Not bRefinanc Then
            If Not bGastosGen Then
                MsgBox "Genere el Plan de Gastos Para Este Credito", vbInformation, "Aviso"
                CmdGastos.SetFocus
                Exit Sub
            End If
        Else
            'Comentado por AVMM -- 25-05-2007
            'ReDim MatGastos(0, 0)
        End If
    End If

    fbGrabar = True 'WIOR 20150717
    'madm 20110303
    If ValidaEnListaNegativa = True Then
        Exit Sub
    End If
    'end madm
    
    '*** PEAC 20080515
    Dim oPers  As COMDPersona.UCOMPersona
    Set oPers = New COMDPersona.UCOMPersona
        If oPers.fgVerificaEmpleado(LblCodCli.Caption) Then
            MsgBox "Este es un Crédito Vinculado...Empleado de la Caja", vbInformation, "Aviso"
            'WIOR 20140123 *********************************
            bTrabajadoVinc = True
            nTipoVinculado = 1
            MsgBox "Se realizará Automáticamente una solicitud de Saldo al Área de Administración de Créditos", vbInformation, "Aviso"
            'WIOR FIN **************************************
        End If
    Set oPers = Nothing
    
    Set oPers = New COMDPersona.UCOMPersona
        If oPers.fgVerificaEmpleadoVincualdo(LblCodCli.Caption) Then
            MsgBox "Este es un Crédito Vinculado...Pariente de Empleado", vbInformation, "Aviso"
            'WIOR 20140123 *********************************
            bTrabajadoVinc = True
            nTipoVinculado = 2
            MsgBox "Se realizará Automáticamente una solicitud de Saldo al Área de Administración de Créditos", vbInformation, "Aviso"
            'WIOR FIN **************************************
        End If
    Set oPers = Nothing
    '*** FIN PEAC 20080515
    
    '**ARLO20171120
    If (Trim(Right(CmbDestino.Text, 5)) = "14") Then
        If (CInt(TxtDiasGracia) > 30) Then
            MsgBox "El periodo de gracia no debe se mayor que 30 dias. ", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    '**************

    '*****************APRI20170630 TI-ERS025-2017*****************
    Dim rsP As ADODB.Recordset
    Set oPers = New COMDPersona.UCOMPersona
    Set rsP = oPers.ObtenerVinculadoRiesgoUnico(LblCodCli.Caption, "", 0)
    
        If Not (rsP.BOF And rsP.EOF) Then
            If rsP.RecordCount = 1 Then
                If rsP!nTotal = 1 Then
                    If MsgBox("El vinculado " & rsP!cPersNombre & " tiene un crédito que se encuentra en " & rsP!cEstado & ". ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                        Exit Sub
                    End If
                 Else
                    If MsgBox("El vinculado " & rsP!cPersNombre & " tiene " & rsP!nTotal & " créditos que se encuentran en mora. ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                        Exit Sub
                    End If
                  End If
            ElseIf rsP.RecordCount > 1 Then
                If MsgBox("El cliente tiene vinculados en persona que se encuentra en mora. ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                    Exit Sub
                   End If
                End If
        End If
        
    Set oPers = Nothing
    '*****************END APRI20170630 ***************************
    '*** PEAC 20100525
    Dim oCredPers  As COMDCredito.DCOMCredito
    Set oCredPers = New COMDCredito.DCOMCredito
        'MAVM 15112010 Observacion de LUCA***
        'If oCredPers.VerificaAprobacionDeAnalistaConSuCredito(ActxCta.NroCuenta, gsCodUser) Then
        If oCredPers.VerificaAprobacionDeAnalistaConSuCredito(ActxCta.NroCuenta, Trim(Right(CmbApoderado.Text, 20))) Then
        '***
            MsgBox "Usuario asignado como analista de un Crédito No puede realizar la Aprobación del mismo crédito.", vbInformation, "Aviso"
            Exit Sub
        End If
    Set oCredPers = Nothing
    '*** FIN PEAC

    'AMDO 20130712 TI-ERS077-2013 ******
        If txtDiasVoucher.Text = "" And Trim(Right(CmbDestino.Text, 5)) = "14" Then
        MsgBox "Es necesario registrar el número de días en que se tendrá el voucher de compra de deuda presentado.", vbInformation, "Aviso"
            'txtDiasVoucher.SetFocus
            EnfocaControl txtDiasVoucher 'EJVG20151112
            Exit Sub
        End If
    'AMDO *******


    '*** PEAC 20080811 ******************************************************
    '*** el codigo de operacio falta definir para aprobar credito por mientras se puso 999999
    lbResultadoVisto = loVistoElectronico.Inicio(1, gColRegistraAprobacion, LblCodCli.Caption)
    If Not lbResultadoVisto Then
        Exit Sub
    End If
    '*** FIN PEAC ************************************************************

    'ALPA 20140724************************************************************
    Set oCredPers = New COMDCredito.DCOMCredito
    Dim oRsVal As ADODB.Recordset
    Set oRsVal = New ADODB.Recordset
    Set oRsVal = oCredPers.RecValidaProcentajeCredito(gdFecSis, ActxCta.Age, CDbl(TxtMontoApr.Text), sLineaTmp)
    If Not (oRsVal.BOF Or oRsVal.EOF) Then
        If oRsVal!nSaldoAdeudado > 0 Then
        If Round((oRsVal!nSaldoCredito / oRsVal!nSaldoAdeudado) * 100, 2) > Round(oRsVal!nPorMax, 2) Then
            MsgBox "No existe saldo para esta aprobación, consultar con el Area de Creditos", vbInformation, "Aviso"
            Exit Sub
        End If
        Else
            MsgBox "No existe saldo para esta aprobación, consultar con el Area de Creditos", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
            MsgBox "No existe saldo para esta aprobación, consultar con el Area de Creditos", vbInformation, "Aviso"
            Exit Sub
    End If
    Set oRsVal = Nothing
    Set oCredPers = Nothing
    '**************************************************************************
    'If Not RecalcularCoberturaGarantias(ActxCta.NroCuenta, bLeasing, lsTpoProdCod, lblSubProducto.Caption, CCur(TxtMontoApr.Text), fvGravamen) Then Exit Sub   'EJVG20150518
    
    Call ObtenerTasaSeguroDesg(True) 'WIOR 20160224
    If MsgBox("Se Va A Grabar los Datos, Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then

        If OptTCuota(3).value Then
            pnTipoCuota = gColocCalendCodCL
        Else
            If OptTPer(0).value Then 'Si es Periodo Fijo
                If OptTCuota(0).value Then
                    If CInt(TxtDiasGracia.Text) > 0 Then
                        pnTipoCuota = gColocCalendCodPFCFPG
                    Else
                        pnTipoCuota = gColocCalendCodPFCF
                    End If
                End If
                If OptTCuota(1).value Then 'Cuota Creciente
                    If CInt(TxtDiasGracia.Text) > 0 Then
                        pnTipoCuota = gColocCalendCodPFCCPG
                    Else
                        pnTipoCuota = gColocCalendCodPFCC
                    End If
                End If
                If OptTCuota(2).value Then 'Cuota Decreciente
                    If CInt(TxtDiasGracia.Text) > 0 Then
                        pnTipoCuota = gColocCalendCodPFCDPG
                    Else
                        pnTipoCuota = gColocCalendCodPFCD
                    End If
                End If
            Else
                If OptTCuota(0).value Then
                    If CInt(TxtDiasGracia.Text) > 0 Then
                        pnTipoCuota = gColocCalendCodFFCFPG
                    Else
                        pnTipoCuota = gColocCalendCodFFCF
                    End If
                End If
                If OptTCuota(1).value Then 'Cuota Creciente
                    If CInt(TxtDiasGracia.Text) > 0 Then
                        pnTipoCuota = gColocCalendCodFFCCPG
                    Else
                        pnTipoCuota = gColocCalendCodFFCC
                    End If
                End If
                If OptTCuota(2).value Then 'Cuota Decreciente
                    If CInt(TxtDiasGracia.Text) > 0 Then
                        pnTipoCuota = gColocCalendCodFFCDPG
                    Else
                        pnTipoCuota = gColocCalendCodFFCD
                    End If
                End If
            End If
        End If
        'Set oNCredito = New COMNCredito.NCOMCredito
        If bRefinanc Then
            'sMontosARef = oNCredito.MontosARefinanciar(ActxCta.NroCuenta, gdFecSis)
            If bCapitalInt Then
                'MAVM 20120529 Correccion Asiento CTB***
                'If MsgBox(IIf(bSustiDeudor, "Montos asumidos por Sustitución de Deudor", "Montos a Refinanciar") & ", Interes : " & sMontosARef(0) & ", Gastos :" & sMontosARef(1) & ", Capital :" & sMontosARef(2) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                If MsgBox(IIf(bSustiDeudor, "Montos asumidos por Sustitución de Deudor", "Montos a Refinanciar") & ", Interes: " & sMontosARef(0) & ", Gastos: " & sMontosARef(1) & ", Capital: " & sMontosARef(2) & " --> TOTAL: " & CCur(sMontosARef(0)) + CCur(sMontosARef(1)) + CCur(sMontosARef(2)) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                '***
                    Exit Sub
                    'Set oNCredito = Nothing
                'MAVM 20120528 Correccion de Asiento CTB
                Else
                    If CCur(TxtMontoApr.Text) <> CCur(CCur(sMontosARef(0)) + CCur(sMontosARef(1)) + CCur(sMontosARef(2))) Then
                        If MsgBox("Por dias transcurridos, Se actualizara el Monto a Refinanciar: " & CDbl(sMontosARef(0)) + CDbl(sMontosARef(1)) + CDbl(sMontosARef(2)) & ", Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                            Exit Sub
                        Else
                            TxtMontoApr.Text = CDbl(sMontosARef(0)) + CDbl(sMontosARef(1)) + CDbl(sMontosARef(2))
                            Set oCredPers = New COMDCredito.DCOMCredito
                            oCredPers.ActualizarMontoColocacEstadoXRefin ActxCta.NroCuenta, TxtMontoApr.Text
                            MsgBox "Genere el Calendario de Pagos para poder Definir los Gastos", vbInformation, "Aviso"
                            bGraciaGen = False
                            bCalendGen = False
                            bGastosGen = False
                            Set oCredPers = Nothing
                            CmdCalend.SetFocus
                            Exit Sub
                        End If
                    End If
                '***
                End If
            Else
                If MsgBox(IIf(bSustiDeudor, "Montos asumidos por Sustitución de Deudor", "Montos a Refinanciar Sin Capitalizar Intereses") & ", Interes : 0.00, Gastos : 0.00" & ", Capital : " & sMontosARef(2) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        
        If Not IsArray(MatCredVig) Then
            ReDim MatCredVig(0)
        End If
        
        Set oNCredito = New COMNCredito.NCOMCredito
        
        'MAVM 20130305 ***
        If vnTipoGracia = 6 Then
            Dim Y As Integer
            Dim MatCalendTemp() As String
            Dim i As Integer
            ReDim MatCalendTemp(UBound(MatCalend) - 1, 13)
            For i = 0 To UBound(MatCalend) - 1
                For Y = 0 To 13
                    MatCalendTemp(i, Y) = MatCalend(i + 1, Y)
                Next Y
            Next i
            Erase MatCalend
            ReDim MatCalend(UBound(MatCalendTemp), 13)
            
            For i = 0 To UBound(MatCalendTemp)
                For Y = 0 To 13
                    MatCalend(i, Y) = MatCalendTemp(i, Y)
                Next Y
            Next i
            Erase MatCalendTemp
        End If
        '***
         nTasa = CDbl(IIf(chkTasa.value = 1, txtInteresTasa.Text, txtTasInt.Text)) 'A
        'MatGastos, nNumgastos  No se tomaran en cuenta
        'Modify Gitu 06-04-2009 se cambio los parametros de tasas de interes para que envie los valores de los Text's respectivos
        'MADM 20091126********************
                'ALPA 20150113**********************************************
        'Set rsExonera = IIf(frmCredSugExonera.feTiposExonera.rows - 1 > 0, frmCredSugExonera.feTiposExonera.GetRsNew(), Nothing) 'LUCV20180601, Comentó según ERS022-2018
        Set rsExonera = Nothing 'LUCV20180601, Agregó según ERS022-2018

        'Dim i As Integer
        Dim bExisteExonera As Boolean
        bExisteExonera = False
        If lbExoneraTasa = 1 Then
            For i = 0 To rsExonera.RecordCount - 1
                If rsExonera!cExoneraCod = "TIP0009" Then
                    rsExonera!Solicitar = 1
                End If
                If rsExonera!Solicitar = 1 Then
                    bExisteExonera = True
                    Exit For
                End If
                rsExonera.MoveNext
            Next i
            If bExisteExonera = False Then
                MsgBox "Debe escoger un tipo de exoneración, de lo contrario desactivar el check de [exoneración Tasa]", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        '************************************************************
        'EJVG20150518 *** Nuevos parametros
        If Not IsArray(fvGravamen) Then
            ReDim fvGravamen(0)
        End If
        
        ReDim vArrDatos(2)
        vArrDatos(0) = MatCalendSegDes
        vArrDatos(1) = fvGravamen
        vArrDatos(2) = lsTpoProdCod
        'END EJVG *******
        'FRHU 20161025 Se modifico el metodo GrabarDatosAprobacion el parametro Right(Trim(CmbCampana.Text), 2) por Right(Trim(CmbCampana.Text), 3)
        Call oNCredito.GrabarDatosAprobacion(ActxCta.NroCuenta, gdFecSis, gColocEstAprob, sLineaTmp, nTasa, CDbl(TxtMontoApr.Text), CInt(SpnCuotasApr.valor), IIf(CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)) > 0, 30, CInt(SpnPlazo.valor)), pnTipoCuota, CInt(TxtDiaFijo.Text), _
                                IIf(ChkProxMes.value = 1, 1, 0), IIf(OptDesemb(0).value, gColocTiposDesembolsoTotal, gColocTiposDesembolsoParcial), IIf(OptTipoCalend(0).value, 0, 1), CInt(TxtDiasGracia.Text), TxtTasaGracia.Text, _
                                vnTipoGracia, nUltNroCalend + 1, MatDesemb, MatCalend, Trim(Right(CmbApoderado.Text, 20)), ChkCuotaCom.value, ChkMiViv.value, nUltNroCalend + 2, MatCalend_2, bRefinanc, _
                                bCapitalInt, gsCodAge, gsCodUser, IIf(OptTipoGasto(0).value, "F", "V"), MatCredVig, sMontosARef, ChkBloqueo.value, rsDR, sError, gsProyectoActual, CInt(TxtDiaFijo2.Text), , IIf(optTipoGracia(1).value, 1, 0), ChkReglamento.value, ChkExoneracion.value, _
                                IIf(chkDesembBcoNac.value = 1, 1, 0), frmCredDesembBcoNac.txtCodigo.Text, Right(frmCredDesembBcoNac.cmbAgencia.Text, 4), nTasCosEfeAnu, actxCtaAho.NroCuenta, pnMovMro, Right(Trim(CmbCampana.Text), 3), Trim(Right(CmbDestino.Text, 5)), Trim(Right(CmbCondicion.Text, 5)), Mid(ActxCta.NroCuenta, 4, 2), lsTpoProdCod, lsTpoCredCod, LblCodCli.Caption, CInt(IIf(txtDiasVoucher.Visible = True, txtDiasVoucher.Text, -1)), chkExoneraMYPE.value, bTrabajadoVinc, nTipoVinculado, CDbl(txtMontoMivivienda.Text), vArrDatos) 'ARCV 29-07-2006(Gracia en Cuotas) , DAOR 20070205 : Parametros para desembolso en banco de la nación, DAOR 20070419 nTasCosEfeAnu
'                               IIf(chkDesembBcoNac.Value = 1, 1, 0), frmCredDesembBcoNac.txtCodigo.Text, Right(frmCredDesembBcoNac.cmbAgencia.Text, 4), nTasCosEfeAnu, actxCtaAho.NroCuenta, pnMovMro, Right(Trim(CmbCampana.Text), 2), Trim(Right(CmbDestino.Text, 5)), Trim(Right(CmbCondicion.Text, 5)), Mid(ActxCta.NroCuenta, 4, 2), lsTpoProdCod, lsTpoCredCod) 'ARCV 29-07-2006(Gracia en Cuotas) , DAOR 20070205 : Parametros para desembolso en banco de la nación, DAOR 20070419 nTasCosEfeAnu
'        Call oNCredito.GrabarDatosAprobacion(ActxCta.NroCuenta, gdFecSis, gColocEstAprob, sLineaTmp, txtTasInt.Text, CDbl(TxtMontoApr.Text), CInt(SpnCuotasApr.Valor), CInt(SpnPlazo.Valor), pnTipoCuota, CInt(TxtDiaFijo.Text), _
'                                IIf(ChkProxMes.value = 1, 1, 0), IIf(OptDesemb(0).value, gColocTiposDesembolsoTotal, gColocTiposDesembolsoParcial), IIf(OptTipoCalend(0).value, 0, 1), CInt(TxtDiasGracia.Text), TxtTasaGracia.Text, _
'                                vnTipoGracia, nUltNroCalend + 1, MatDesemb, MatCalend, Trim(Right(CmbApoderado.Text, 20)), ChkCuotaCom.value, ChkMiViv.value, nUltNroCalend + 2, MatCalend_2, bRefinanc, _
'                                bCapitalInt, gsCodAge, gsCodUser, IIf(OptTipoGasto(0).value, "F", "V"), MatCredVig, sMontosARef, ChkBloqueo.value, rsDR, sError, gsProyectoActual, CInt(TxtDiaFijo2.Text), chkIncremenK.value, IIf(optTipoGracia(1).value, 1, 0), ChkReglamento.value, ChkExoneracion.value, _
'                                IIf(chkDesembBcoNac.value = 1, 1, 0), frmCredDesembBcoNac.txtCodigo.Text, Right(frmCredDesembBcoNac.cmbAgencia.Text, 4), nTasCosEfeAnu, actxCtaAho.NroCuenta, pnMovMro, Right(Trim(CmbCampana.Text), 2), Trim(Right(CmbDestino.Text, 5)), Trim(Right(CmbCondicion.Text, 5))) 'ARCV 29-07-2006(Gracia en Cuotas) , DAOR 20070205 : Parametros para desembolso en banco de la nación, DAOR 20070419 nTasCosEfeAnu
                                'MAVM 20090702 se agrego Trim(Right(CmbDestino.Text, 5)), Trim(Right(CmbCondicion.Text, 5))
                                'By Capi 10042008 se agrego parametro actxctaaho.Nrocuenta para desembolso a terceros
                                'AMDO 20130712 se agrego parametro pnDiasVoucher TI-ERS077-2013
                                'WIOR 20131111 AGREGO fnCuotaBalon
                                'WIOR 20140128 AGREGO bTrabajadoVinc,nTipoVinculado
                                'WIOR 20140825 AGREGO MatCalendSegDes
                                'APRI20180821 ERS061-2018 REEMPLAZO lnCuotaBalon POR chkExoneraMYPE
        '**********************************
        'FRHU 20170915 ERS049-2017
        Dim oLineasM As New COMDCredito.DCOMLineaCredito
        Call oLineasM.InsertarColocHistorialTasaMaxima(ActxCta.NroCuenta, CDbl(lnTasaInicial), CDbl(lnTasaFinal), CDbl(txtTasInt.Text), sMovNroM, gColocEstAprob, gsCodCargo)
        Set oLineasM = Nothing
        'FIN FRHU 20170915
        ''*** PEAC 20090126
        'MAVM 20120529 Correccion Asiento CTB ***
        'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta
        '->***** LUCV20180601, Coment?. Seg?n ERS022-2018
        'If rsExonera.RecordCount > 0 Then
        'If Not rsExonera.EOF Or rsExonera.BOF Then
        '    Call oNCredito.GrabarSugerenciaExoneracion(ActxCta.NroCuenta, lbExoneraTasa, rsExonera)
        'End If
        '<-***** Fin LUCV20180601
        
        'ANGC/RIRO 20200520 ****
        If oNCredito.bAplicaComision Then
            MsgBox "Se aplicó comisión por desembolso bajo el concepto de COFIDE REACTIVA " & _
            "por el monto de " & Format(oNCredito.nComisionCofideReactiva, "##0.00"), vbInformation, "Cofide Reactiva"
        End If
        'END
        
        Set oNCredito = Nothing
        loVistoElectronico.RegistraVistoElectronico (pnMovMro)
        '*** FIN PEAC
        
        If sError <> "" Then
            MsgBox sError, vbInformation, "Aviso"
        Else
        
        'WIOR 20151223 ***
        Set oCredPers = New COMDCredito.DCOMCredito
        Call oCredPers.EliminarDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstAprob)
        
'        If fbMIVIVIENDA Then
'            If IsArray(fArrMIVIVIENDA) Then
'                If Trim(fArrMIVIVIENDA(0)) <> "" Then
'                    Call oCredPers.RegistrarDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstAprob, CDbl(fArrMIVIVIENDA(0)), CDbl(fArrMIVIVIENDA(1)), _
'                    CDbl(fArrMIVIVIENDA(2)), CDbl(fArrMIVIVIENDA(3)), CDbl(fArrMIVIVIENDA(4)), CInt(fArrMIVIVIENDA(5)), CInt(fArrMIVIVIENDA(6)), _
'                    CDbl(fArrMIVIVIENDA(7)), CDbl(fArrMIVIVIENDA(8)), CInt(fArrMIVIVIENDA(10)))
'                End If
'            End If
'        End If
            'CTI5 ERS012020**********************************
        If fbMIVIVIENDA Then
            If IsArray(fArrMIVIVIENDA) Then
               If Trim(fArrMIVIVIENDA(0)) <> "" Then
                     Call oCredPers.RegistrarDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstAprob, CDbl(fArrMIVIVIENDA(0)), CDbl(fArrMIVIVIENDA(1)), _
                     CDbl(fArrMIVIVIENDA(2)), CDbl(fArrMIVIVIENDA(3)), CDbl(fArrMIVIVIENDA(4)), CDbl(fArrMIVIVIENDA(5)), CDbl(fArrMIVIVIENDA(6)), _
                     CDbl(fArrMIVIVIENDA(7)), CDbl(fArrMIVIVIENDA(8)), CInt(fArrMIVIVIENDA(10)), CDbl(fArrMIVIVIENDA(11)))
                     objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Registro de Datos MIVIVIENDA Aprobación", ActxCta.NroCuenta, gCodigoCuenta 'CTI3 ERS001-2020
              End If
            End If
        End If
    '************************************************
        'WIOR FIN ********
        
        'MAVM 20120529 Correccion Asiento CTB ***
        'RECO20161020 ERS060-2016 ********************************************************** 'FRHU 20161110 Descomentado
        Dim oNCOMColocEval As New NCOMColocEval
        Dim lcMovNro As String

        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'JOEP22052017
        objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta 'JOEP22052017

        If Not ValidaExisteRegProceso(ActxCta.NroCuenta, gTpoRegCtrlAprobacion) Then
           'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'Comentado JOEP22052017
           'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta'Comentado JOEP22052017
           Call oNCOMColocEval.insEstadosExpediente(ActxCta.NroCuenta, "Aprobacion de Credito", lcMovNro, "", "", "", 1, 2002, gTpoRegCtrlAprobacion)
           Set oNCOMColocEval = Nothing
        End If
        'RECO FIN **************************************************************************
        
        lnLogicoBuscarDatos = 0 'ALPA20150116
        Call LimpiarPantalla
        End If
    End If
End Sub

Function VerificaTipoCredito() As String
'Devuelve AGRICOLA si  para creditos agricolas
'Devuelve Comercial si es para pequeñas empresas
'Devuelve consumo si es de consumo
 Dim sTipoCredito As String
 Dim sSubTipoCredito As String
 Dim sCuenta As String
 
 sTipoCredito = Mid(ActxCta.NroCuenta, 6, 1) & "00"
 sSubTipoCredito = Mid(ActxCta.NroCuenta, 6, 1) & Mid(ActxCta.NroCuenta, 7, 2)
 If (sTipoCredito = 100 And sSubTipoCredito = "102") Or _
    (sTipoCredito = "200" And sSubTipoCredito = "202") Then
     VerificaTipoCredito = "AGRICOLA"
 ElseIf (sTipoCredito = "300" And sSubTipoCredito = "301") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "302") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "303") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "304") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "305") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "320") Then
          VerificaTipoCredito = "CONSUMO"
 End If
End Function

Private Sub cmdGrabar_LostFocus()
fbGrabar = False 'WIOR 20150717
End Sub

Private Sub cmdLineas_Click()
Dim oLineas As COMDCredito.DCOMLineaCredito
Dim sCtaCod As String

bBuscarLineas = True
sCtaCod = ActxCta.NroCuenta
Set oLineas = New COMDCredito.DCOMLineaCredito
txtBuscarLinea.Text = ""
lblLineaDesc.Caption = ""
If ActxCta.Cuenta = "" Then
    MsgBox "Ingrese Nº de Crédito", vbCritical, "Aviso"
    Exit Sub
End If
'ALPA 20150113*****************************************
'txtBuscarLinea.rs = oLineas.RecuperaLineasProductoArbol(Mid(sCtaCod, 6, 3), Mid(sCtaCod, 9, 1), , Mid(sCtaCod, 4, 2), CInt(SpnPlazo.Valor), CDbl(TxtMontoApr.Text), CInt(SpnCuotasApr.Valor))
'txtBuscarLinea.rs = oLineas.RecuperaLineasProductoArbol(lsTpoCredCod, Mid(sCtaCod, 9, 1), , Mid(sCtaCod, 4, 2), CInt(SpnPlazo.valor), CDbl(TxtMontoApr.Text), CInt(SpnCuotasApr.valor))
txtBuscarLinea.rs = oLineas.RecuperaLineasProductoArbol(lsTpoCredCod, Mid(sCtaCod, 9, 1), , Mid(sCtaCod, 4, 2), CInt(SpnPlazo.valor), CDbl(TxtMontoApr.Text), CInt(SpnCuotasApr.valor), IIf(Trim(txtTasInt.Text) = "", 0, txtTasInt.Text), IIf(Trim(TxtDiasGracia.Text) = "", 0, TxtDiasGracia.Text), gdFecSis)
Set oLineas = Nothing
End Sub


Private Sub cmdsalir_Click()
    Set oPersona = Nothing
    Unload Me
End Sub
'->*****LUCV20180601, Coment? seg?n ERS022-2018
'Private Sub CmdTiposGracia_Click()
'    Dim oCredito As COMNCredito.NCOMCredito
'
'    On Error GoTo ErrorCmdTiposGracia_Click
'    If CInt(TxtDiasGracia.Text) = 0 Then
'        MsgBox "Ingrese los Dias de Gracia", vbInformation, "Aviso"
'        'MAVM 25102010 ***
'        'TxtDiasGracia.SetFocus
'        txtFechaFija.SetFocus
'        '***
'        Exit Sub
'    End If
'    If Not ValidaDatos Then
'        Exit Sub
'    End If
'    Set oCredito = New COMNCredito.NCOMCredito
'    If CInt(TxtDiasGracia.Text) > 0 Then
'        'MatGracia = frmCredGracia.Inicio(CInt(TxtDiasGracia.Text), oCredito.MontoIntPerDias(CDbl(LblTasaGraciaSug.Caption), CInt(TxtDiasGracia.Text), CDbl(TxtMontoApr.Text)), CInt(SpnCuotasApr.Valor), vnTipoGracia, ActxCta.NroCuenta)
'        MatGracia = frmCredGracia.Inicio(CInt(TxtDiasGracia.Text), oCredito.MontoIntPerDias(nTasaGrac, CInt(TxtDiasGracia.Text), CDbl(TxtMontoApr.Text)), CInt(SpnCuotasApr.valor), vnTipoGracia, ActxCta.NroCuenta)
'        If UBound(MatGracia) > 0 Then
'            bGraciaGen = True
'        Else
'            bGraciaGen = False
'        End If
'    End If
'    Set oCredito = Nothing
'    Exit Sub
'
'ErrorCmdTiposGracia_Click:
'        MsgBox Err.Description, vbCritical, "Aviso"
'End Sub
'<-***** Fin LUCV20180601
Private Sub cmdVerEntidades_Click()
    Dim oCredPersRela As UCredRelac_Cli
    Set oCredPersRela = New UCredRelac_Cli
    Call oCredPersRela.CargaRelacPersCred(ActxCta.NroCuenta)
    oCredPersRela.IniciarMatriz
    Dim sDocumento As String
    sDocumento = oCredPersRela.ObtenerDocumento
    If Len(Trim(ActxCta.NroCuenta)) = 18 Then
        Set oRsVerEntidades = New ADODB.Recordset
        Set oRsVerEntidades = frmCredVerEntidades.VerEntidades(ActxCta.NroCuenta, LblCodCli.Caption, sDocumento, 1)
        nLogicoVerEntidades = 1
    End If
End Sub
'ALPA 20090929******************************
Private Sub cmdVinculados_Click()
    frmGruposEconomicos.Show 1
End Sub
'*******************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
    bDesembParcialGenerado = False
    vnTipoGracia = 0
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    'By Capi 10042008
    actxCtaAho.CMAC = gsCodCMAC
    actxCtaAho.Age = gsCodAge
    actxCtaAho.Prod = gCapAhorros
    Call CargaControles
    Frame5.Enabled = False
    ReDim MatGracia(0, 0)
    ReDim MatCalend(0, 0)
    ReDim MatDesemb(0, 0)
    ReDim MatGastos(0, 0)
    ReDim MatCalendSegDes(4, 0) 'WIOR 20140528
    nNumGastos = 0
    ChkBloqueo.value = 0
    ChkReglamento.value = 0
    ChkExoneracion.value = 0
    MatCredVig = ""
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredRegistraAprobacionCredito
    'ALPA 20100605 B2***************************************
        Label123.Visible = False
        lblInstitucionFinanciera.Visible = False
    '*******************************************************
    'ALPA 20110325 *****************************************
    SSTab1.Tab = 3
    '*******************************************************
    fbLimpiaExoSegDes = False 'WIOR 20130710
    fnCuotaBalon = 0 'WIOR 20131111
    fbGrabar = False 'WIOR 20150717
    
    '*** PEAC 20160822 - oculta estos tabs ya que no se usan
    Me.SSTab1.TabVisible(4) = False 'INGRESOS Y EGRESOS
    Me.SSTab1.TabVisible(5) = False 'BALANCE
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Unload(frmCredVigentes)
    MatCredVig = ""
    Set objPista = Nothing
End Sub



Private Sub LblMontoSol_Change()
 If Mid(ActxCta.NroCuenta, 9, 1) = "1" Then
        LblMontoSol.ForeColor = vbBlue
    Else
        LblMontoSol.ForeColor = &H289556
    End If
End Sub

Private Sub LblMontoSug_Click()
    If Mid(ActxCta.NroCuenta, 9, 1) = "1" Then
        LblMontoSug.ForeColor = vbBlue
    Else
        LblMontoSug.ForeColor = &H289556
    End If
End Sub

Private Sub Optdesemb_Click(Index As Integer)
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
    If Index = 1 Then
        TxtMontoApr.Text = "0.00"
        TxtMontoApr.Enabled = False
        CmdDesembolsos.Enabled = True
        ReDim MatDesemb(0, 0)
        bDesembParcialGenerado = False
        Me.SpnCuotasApr.valor = 1
        Me.SpnCuotasApr.Enabled = False
    Else
        TxtMontoApr.Text = "0.00"
        TxtMontoApr.Enabled = True
        CmdDesembolsos.Enabled = False
        bDesembParcialGenerado = False
        Me.SpnCuotasApr.valor = 1
        Me.SpnCuotasApr.Enabled = True
    End If
End Sub

Private Sub opttcuota_Click(Index As Integer)
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
    If Index = 3 Then
        HabilitaCuotaLibre True
    Else
        HabilitaCuotaLibre False
    End If
End Sub

Private Sub optTipoGracia_Click(Index As Integer)
If Index = 0 Then
    'chkIncremenK.Visible = True
Else
    'chkIncremenK.Visible = False
End If
CmdTiposGracia.Enabled = False
End Sub

Private Sub opttper_Click(Index As Integer)
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
    If Index = 1 Then
    
        'Comentado Por MAVM 25102010 ***
        'TxtDiaFijo.Enabled = True
        'TxtDiaFijo.Text = "00"
        'ChkProxMes.Enabled = True
        'If TxtDiaFijo.Enabled And TxtDiaFijo.Visible Then
        '    TxtDiaFijo.SetFocus
        'End If
        '***
        
        'Add by Gitu para ingresar fecha fija para el calculo de dias de gracia
        'Descomentar cuando ten seguros del cambio
        'txtFechaFija.Enabled = True
        'txtFechaFija.Text = "__/__/____"
        'MAVM 20100823 *** Se descomento por efectos en la reprogramacion
        'SpnPlazo.valor = "0"
        SpnPlazo.Enabled = False
        SpnPlazo.valor = "0"
        
        'MAVM 25102010 ***
        ChkProxMes.value = 0
        TxtDiaFijo2.Text = "00"
        optTipoGracia(1).Enabled = False
        optTipoGracia(1).value = False
        
        txtFechaFija.Enabled = True
        TxtDiaFijo.Enabled = False
        TxtDiaFijo2.Enabled = False
        TxtDiasGracia.Enabled = False
        TxtDiasGracia.Text = "0"
        TxtTasaGracia.Text = "0.00"
        optTipoGracia(0).value = False
        optTipoGracia(1).value = False
        CmdTiposGracia.Enabled = True
        txtFechaFija.Text = gdFecSis
        Me.TxtFecDesemb = Format(gdFecSis, "dd/mm/yyyy") 'MAVM 20121113
        GenerarFechaPago
        '***
        
    Else
        'MAVM 25102010 ***
        SpnPlazo.Enabled = True
        ChkProxMes.value = 0
        TxtDiaFijo2.Text = "00"
        '***
        
        TxtDiaFijo.Text = "00"
        TxtDiaFijo.Enabled = False
        ChkProxMes.Enabled = False
        'txtFechaFija.Text = "__/__/____"
        'MAVM 20100823 *** Se descomento por efectos en la reprogramacion
        'spnPlazo.Enabled = False
        '***
        
        'MAVM 25112010 ***
        txtFechaFija.Text = "__/__/____"
        TxtFecDesemb.Text = gdFecSis
        GenerarFechaPago
        TxtDiasGracia.Enabled = False
        TxtDiasGracia.Text = "0"
        TxtTasaGracia.Text = "0.00"
        optTipoGracia(0).value = False
        optTipoGracia(1).value = False
        CmdTiposGracia.Enabled = True
        If gnAgenciaCredEval = 0 Then
            If Not fbMIVIVIENDA Then 'WIOR 20160107
                SpnPlazo.SetFocus 'JUEZ 20121218
            End If 'WIOR 20160107
        End If
        '***
    End If
End Sub

Private Sub SpnCuotasApr_Change()
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
    bDesembParcialGenerado = False
    bBuscarLineas = False
    DefineCondicionCredito
    'ValidaCuotaBalon 'WIOR 20131115 'COMENTADO POR APRI20180820 ERS061-2018
    Call MostrarLineas 'ALPA 20150115
End Sub

Private Sub SpnCuotasApr_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        'MAVM 20101201 ***
        'SpnPlazo.SetFocus
        txtFechaFija.SetFocus
        '***
     End If
End Sub

Private Sub SpnCuotasApr_LostFocus()
    If Trim(SpnCuotasApr.valor) = "" Then
        SpnCuotasApr.valor = "0"
    Else
        SpnCuotasApr.valor = Format(SpnCuotasApr.valor, "#0")
    End If
End Sub

Private Sub SpnPlazo_Change()
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
    bBuscarLineas = False
    GenerarFechaPago 'MAVM 30092010
    
    fraGracia.Enabled = False
    TxtDiasGracia.Text = "0"
    TxtTasaGracia.Text = "0.00"
    TxtTasaGracia.Enabled = False
    optTipoGracia(0).value = False
    optTipoGracia(1).value = False
    Call MostrarLineas 'ALPA 20150115
End Sub

Private Sub spnPlazo_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        TxtFecDesemb.SetFocus
     End If
End Sub

Private Sub txtBuscarLinea_EmiteDatos()
Dim sCodigo As String
'Dim oLineas As COMDCredito.DCOMLineaCredito
Dim oCred As COMNCredito.NCOMCredito
Dim RLinea As ADODB.Recordset
Dim bExisteLineaCred As Boolean
Dim sCtaCodOrigen As String ' DAOR 20070409

sCodigo = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))

If sCodigo <> "" Then
    sLineaTmp = sCodigo
'   txtBuscarLinea.Text = sCodigo
    If txtBuscarLinea.psDescripcion <> "" Then lblLineaDesc = txtBuscarLinea.psDescripcion Else lblLineaDesc = ""
        'VERIFICAR
       'Carga Datos de la Linea de Credito seleccionada
        Set oCred = New COMNCredito.NCOMCredito
        'Comentado por DAOR 20070409***********************************************
        'Call oCred.CargarDatosLinea(sLineaTmp, sCodigo, RLinea, bExisteLineaCred)
        '**************************************************************************
        
        'Call oCred.CargarDatosLinea(sLineaTmp, sCodigo, RLinea, bExisteLineaCred)
        
        'MADM 20110419
'        '**DAOR 20070409***********************************************************
'        If Not bRefinanc Then
'             Call oCred.CargarDatosLinea(sLineaTmp, sCodigo, RLinea, bExisteLineaCred)
'        Else
'             sCtaCodOrigen = Right$(txtBuscarLinea.psDescripcion, 18)
'             Call oCred.CargarDatosLinea(sLineaTmp, sCodigo, RLinea, bExisteLineaCred, sCtaCodOrigen)
'        End If
        '**************************************************************************
        Set oCred = Nothing
        'Set oLineas = New COMDCredito.DCOMLineaCredito
        'Set RLinea = oLineas.RecuperaLineadeCredito(sCodigo)
        'Set oLineas = Nothing
        bExisteLineaCred = True
        'If RLinea.RecordCount > 0 Then
        If bExisteLineaCred Then
        '    Call CargaDatosLinea
            'CargarTasaLinea RLinea 'sLineaTmp
        Else
            MsgBox "No existen Líneas de Crédito con el Plazo seleccionado", vbInformation, "Aviso"
            txtBuscarLinea.Text = ""
            lblLineaDesc = ""
        End If
    
    'txtBuscarLinea.Text = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
Else
    lblLineaDesc = ""
End If
End Sub

'Sub CargarTasaLinea(ByVal pcLineaCred As String)
Sub CargarTasaLinea(ByVal prsLineaCred As ADODB.Recordset)
'ALPA 20091202*************************
'    Dim objLineaTasa As COMDCredito.DCOMLineaCredito
'    Dim rs As ADODB.Recordset
    
'    Set objLineaTasa = New COMDCredito.DCOMLineaCredito
'    Set rs = objLineaTasa.ObtenerTasaLinea(pcLineaCred)
'    Set objLineaTasa = Nothing
'    Do Until prsLineaCred.EOF
'        If CInt(prsLineaCred!nColocLinCredTasaTpo) = "1" Then
'            nTasaComp = CDbl(prsLineaCred!nTasaIni)
'            txtTasInt.Text = nTasaComp
'        ElseIf CInt(prsLineaCred!nColocLinCredTasaTpo) = "3" Then
'            nTasaMor = CDbl(prsLineaCred!nTasaIni)
'        ElseIf CInt(prsLineaCred!nColocLinCredTasaTpo) = "5" Then
'            nTasaGrac = CDbl(prsLineaCred!nTasaIni)
'            TxtTasaGracia.Text = nTasaGrac
'        End If
'        prsLineaCred.MoveNext
'    Loop
'****************************************
End Sub
'******COMENTADO POR APRI20180820 ERS061-2018******
'WIOR 20131116 ************************************************
'Private Sub txtCuotaBalon_Change()
'ValidaCuotaBalon
'End Sub
'WIOR FIN *****************************************************
'*****END APRI

Private Sub TxtDiaFijo_Change()
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
End Sub

Private Sub TxtDiaFijo_GotFocus()
    fEnfoque TxtDiaFijo
End Sub

Private Sub TxtDiaFijo_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
     '->***** LUCV20180601, Comentó según ERS022-2018
'     If KeyAscii = 13 Then
'        ChkProxMes.SetFocus
'     End If
'<-*****Fin LUCV20180601
End Sub

Private Sub TxtDiaFijo_LostFocus()
    If Trim(TxtDiaFijo.Text) = "" Then
        TxtDiaFijo.Text = "00"
    Else
        TxtDiaFijo.Text = Right("00" & TxtDiaFijo.Text, 2)
    End If
End Sub

Private Sub TxtDiaFijo2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub TxtDiasGracia_Change()
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
End Sub

'->***** LUCV20180601, Comentó según ERS022-2018
'Private Sub TxtDiasGracia_GotFocus()
'    fEnfoque TxtDiasGracia
'End Sub
'
'Private Sub TxtDiasGracia_KeyPress(KeyAscii As Integer)
'     KeyAscii = NumerosEnteros(KeyAscii)
'     If KeyAscii = 13 Then
'        If CmdTiposGracia.Enabled Then
'            CmdTiposGracia.SetFocus
'        End If
'     End If
'End Sub
'
'Private Sub TxtDiasGracia_LostFocus()
'    If Trim(TxtDiasGracia.Text) = "" Then
'        TxtDiasGracia.Text = "0"
'    Else
'        TxtDiasGracia.Text = Format(TxtDiasGracia.Text, "#0")
'    End If
'End Sub
'<-***** Fin LUCV20180601

''AMDO20130712 TI-ERS077-2013 ******
Private Sub txtDiasVoucher_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtDiasVoucher_LostFocus()
    If Len(txtDiasVoucher.Text) = 2 Then
        If CInt(txtDiasVoucher.Text) < 0 Or CInt(txtDiasVoucher.Text) > 30 Then
            MsgBox "Los diás en que se tendrá el voucher de compra de deuda presentado no puede ser mayor a 30."
            txtDiasVoucher.Text = ""
            txtDiasVoucher.SetFocus
        End If
    End If
End Sub
''END AMDO

Private Sub TxtFecDesemb_Change()
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
End Sub

Private Sub TxtFecDesemb_GotFocus()
    fEnfoque TxtFecDesemb
End Sub

Private Sub TxtFecDesemb_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        'MAVM 15112010 ***
        'TxtDiasGracia.SetFocus
        GenerarFechaPago
        txtFechaFija.SetFocus
        If OptTPer(0).value = True Then
            SpnPlazo.Enabled = True
            ChkProxMes.value = 0
            TxtDiaFijo2.Text = "00"
            TxtDiaFijo.Text = "00"
            TxtDiaFijo.Enabled = False
            ChkProxMes.Enabled = False
            
            TxtDiasGracia.Enabled = False
            TxtDiasGracia.Text = "0"
            TxtTasaGracia.Text = "0.00"
            optTipoGracia(0).value = False
            optTipoGracia(1).value = False
            CmdTiposGracia.Enabled = True
        Else
            SpnPlazo.Enabled = False
            SpnPlazo.valor = "0"
            
            ChkProxMes.value = 0
            TxtDiaFijo2.Text = "00"
            optTipoGracia(1).Enabled = False
            optTipoGracia(1).value = False
            
            txtFechaFija.Enabled = True
            TxtDiaFijo.Enabled = False
            TxtDiaFijo2.Enabled = False
            TxtDiasGracia.Enabled = False
            TxtDiasGracia.Text = "0"
            TxtTasaGracia.Text = "0.00"
            optTipoGracia(0).value = False
            optTipoGracia(1).value = False
            CmdTiposGracia.Enabled = True
            txtFechaFija.Text = gdFecSis
            GenerarFechaPago
        End If
        '***
     End If
End Sub

Private Sub TxtFecDesemb_LostFocus()
'Descomentar cuando esten seguros del cambio GITU
'    If Not (txtFechaFija = "__/__/____" Or TxtFecDesemb.Text = "__/__/____") Then
'        If CDate(txtFechaFija.Text) < CDate(TxtFecDesemb.Text) Then
'            MsgBox "Fecha tiene que ser menor a la fecha fija", vbInformation, "Atencion!"
'            TxtFecDesemb.SetFocus
'            Exit Sub
'        Else
'            If OptTPer(0).Index = 1 Then
'                nPlazo = SpnPlazo.Valor
'            Else
'                nPlazo = 30
'            End If
'
'            If (CDate(txtFechaFija.Text) - CDate(TxtFecDesemb.Text)) - nPlazo < 0 Then
'                TxtDiasGracia.Text = 0
'            Else
'                TxtDiasGracia.Text = (CDate(txtFechaFija.Text) - CDate(TxtFecDesemb.Text)) - nPlazo
'            End If
'            TxtDiaFijo.Text = Day(CDate(txtFechaFija.Text))
'        End If
'    End If
End Sub

Private Sub txtFechaFija_Change()
    'Descomentar cuando esten seguros del cambio GITU
'    bGraciaGen = False
'    bCalendGen = False
'    bGastosGen = False
End Sub

'MAVM 25102010 ***
Private Sub txtFechaFija_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        CmdTiposGracia.Enabled = True
        optTipoGracia(0).value = 0
        TxtTasaGracia.Text = Format(txtTasInt.Text, "#0.0000") 'LUCV20180601, según ERS022-2018
        If Not Trim(ValidaFecha(txtFechaFija.Text)) = "" Then
            MsgBox Trim(ValidaFecha(txtFechaFija.Text)), vbInformation, "Aviso"
            Exit Sub
        End If
        If CDate(txtFechaFija.Text) < CDate(TxtFecDesemb.Text) Then
            MsgBox "Fecha tiene que ser mayor a la fecha del dia", vbInformation, "Atencion!"
            txtFechaFija.SetFocus
            txtFechaFija.Text = CDate(gdFecSis + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
            Exit Sub
        End If
'->***** LUCV20170915, Agregó y Comentó Según Modificaciones del reglamento (4.2.-Crédito Refinanciado. (b)) ->*****
      If Not ValidaPeriodoGracia Then
            txtFechaFija.SetFocus
            Exit Sub
       End If
'        If OptTPer(0).value = True Then
'            If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") Then
'                MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
'                If SpnPlazo.Enabled Then SpnPlazo.SetFocus
'                Exit Sub
'            End If
'
'            If CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor) > CDate(txtFechaFija.Text) Then
'                MsgBox "La Fecha de Pago No puede ser Menor que el Plazo", vbInformation, "Aviso"
'                txtFechaFija.Text = CDate(CDate(TxtFecDesemb.Text) + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
'                txtFechaFija.SetFocus
'                Exit Sub
'            End If
'
'            If txtFechaFija.Text > CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor) Then
'                fraGracia.Enabled = True
'                TxtTasaGracia.Enabled = True
'                'chkIncremenK.value = 0
'
'                TxtDiasGracia.Text = CInt(CDate(txtFechaFija.Text) - CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor))
'            Else
'                TxtDiasGracia.Text = "0.00"
'
'                fraGracia.Enabled = False
'                TxtDiasGracia.Text = "0"
'                TxtTasaGracia.Text = "0.00"
'                TxtTasaGracia.Enabled = False
'                optTipoGracia(0).value = False
'                optTipoGracia(1).value = False
'
'            End If
'            'ALPA 20160419********************************************
'            If SpnPlazo.valor < 30 Then
'                MsgBox "El crédito no cumple con las especificaciones de plazo, no debe ser menor a 30 días", vbInformation, "Aviso"
'                Exit Sub
'            End If
'            '*********************************************************
'        Else
'            If Month(CDate(TxtFecDesemb.Text)) = Month(txtFechaFija.Text) And Year(CDate(TxtFecDesemb.Text)) = Year(txtFechaFija.Text) Then
'                ChkProxMes.value = 0
'            Else
'                ChkProxMes.value = 1
'            End If
'
'            If CDate(CDate(TxtFecDesemb.Text) + 30) < txtFechaFija.Text Then
'                fraGracia.Enabled = True
'                TxtTasaGracia.Enabled = True
'                TxtDiasGracia.Text = CDate(txtFechaFija.Text) - CDate(CDate(TxtFecDesemb.Text) + 30)
'                optTipoGracia(0).Enabled = True 'MAVM 20130209
'            Else
'                fraGracia.Enabled = False
'                TxtDiasGracia.Text = "0"
'                TxtTasaGracia.Text = "0.00"
'                TxtTasaGracia.Enabled = False
'                optTipoGracia(0).value = False
'                optTipoGracia(1).value = False
'            End If
'
'            TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaFija.Text)), 2)
'        End If
'Fin LUCV20170925 <-*****
    End If
End Sub
'*****->LUCV20170925, Creó. Según Modificaciones del reglamento (4.2.-Crédito Refinanciado. (b))
Public Function ValidaPeriodoGracia()
    Dim nPeriodoGraciaRefinanciado As Integer
    nPeriodoGraciaRefinanciado = 0
    ValidaPeriodoGracia = True
    
    If OptTPer(0).value = True Then
            If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") Then
                MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
                If SpnPlazo.Enabled Then SpnPlazo.SetFocus
                ValidaPeriodoGracia = False
            End If
            
            If CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor) > CDate(txtFechaFija.Text) Then
                MsgBox "La Fecha de Pago No puede ser Menor que el Plazo", vbInformation, "Aviso"
                txtFechaFija.Text = CDate(CDate(TxtFecDesemb.Text) + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
                txtFechaFija.SetFocus
                ValidaPeriodoGracia = False
            End If
        
            If txtFechaFija.Text > CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor) Then
                fraGracia.Enabled = True
                TxtTasaGracia.Enabled = True
                TxtDiasGracia.Text = CInt(CDate(txtFechaFija.Text) - CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor))
                nPeriodoGraciaRefinanciado = TxtDiasGracia.Text
            Else
                TxtDiasGracia.Text = "0.00"
                fraGracia.Enabled = False
                TxtDiasGracia.Text = "0"
                TxtTasaGracia.Text = "0.00"
                TxtTasaGracia.Enabled = False
                optTipoGracia(0).value = False
                optTipoGracia(1).value = False
            End If
            
            If SpnPlazo.valor < 30 Then
                MsgBox "El crédito no cumple con las especificaciones de plazo, no debe ser menor a 30 días", vbInformation, "Aviso"
                ValidaPeriodoGracia = False
        End If

        'JAOR 20200726 COMENTÓ
        'If bRefinanc And nPeriodoGraciaRefinanciado > 30 Then
        ' MsgBox "Periodo de gracia no debe ser mayor a 30 días", vbInformation, "Aviso"
        '   txtFechaFija.SetFocus
        '   txtFechaFija.Text = CDate(gdFecSis + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
        '    TxtDiasGracia.Text = "0"
        '   TxtTasaGracia.Text = "0.00"
        '   ValidaPeriodoGracia = False
        'End If
    Else
        ValidaPeriodoGracia = True
        If Month(CDate(TxtFecDesemb.Text)) = Month(txtFechaFija.Text) And Year(CDate(TxtFecDesemb.Text)) = Year(txtFechaFija.Text) Then
            ChkProxMes.value = 0
        Else
            ChkProxMes.value = 1
        End If

        If CDate(CDate(TxtFecDesemb.Text) + 30) < txtFechaFija.Text Then
            fraGracia.Enabled = True
            TxtTasaGracia.Enabled = True
            TxtDiasGracia.Text = CDate(txtFechaFija.Text) - CDate(CDate(TxtFecDesemb.Text) + 30)
            optTipoGracia(0).Enabled = True
            nPeriodoGraciaRefinanciado = TxtDiasGracia.Text
        Else
            fraGracia.Enabled = False
            TxtDiasGracia.Text = "0"
            TxtTasaGracia.Text = "0.00"
            TxtTasaGracia.Enabled = False
            optTipoGracia(0).value = False
            optTipoGracia(1).value = False
        End If
        TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaFija.Text)), 2)

        'JAOR 20200726 COMENTÓ
        'If bRefinanc And nPeriodoGraciaRefinanciado > 30 Then
        'MsgBox("Periodo de gracia no debe ser mayor a 30 días", vbInformation, "Aviso")
        'txtFechaFija.Text = CDate(gdFecSis + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
        'txtFechaFija.SetFocus()
        'TxtDiasGracia.Text = "0"
        'TxtTasaGracia.Text = "0.00"
        'ValidaPeriodoGracia = False
                'End If
    End If
End Function
'<-***** Fin LUCV20170915<-*****

Private Sub txtFechaFija_LostFocus()
'Add Gitu 20-08-2008
'Descomentar cuando esten seguros del cambio GITU
'    If Not (txtFechaFija = "__/__/____" Or TxtFecDesemb.Text = "__/__/____") Then
'        If CDate(txtFechaFija.Text) < CDate(TxtFecDesemb.Text) Then
'            MsgBox "Fecha tiene que ser mayor a la fecha del desembolso", vbInformation, "Atencion!"
'            txtFechaFija.SetFocus
'            Exit Sub
'        Else
'            If OptTPer(0).value = True Then
'                nPlazo = Val(SpnPlazo.Valor)
'            Else
'                nPlazo = 30
'            End If
'
'            If (CDate(txtFechaFija.Text) - CDate(TxtFecDesemb.Text)) - nPlazo < 0 Then
'                TxtDiasGracia.Text = 0
'            Else
'                TxtDiasGracia.Text = (CDate(txtFechaFija.Text) - CDate(TxtFecDesemb.Text)) - nPlazo
'            End If
'            TxtDiaFijo.Text = Day(CDate(txtFechaFija.Text))
'        End If
'    End If
End Sub

Private Sub TxtMontoApr_Change()
    bGraciaGen = False
    bCalendGen = False
    bGastosGen = False
    bDesembParcialGenerado = False
    If Mid(ActxCta.NroCuenta, 9, 1) = "1" Then
        TxtMontoApr.ForeColor = vbBlue
    Else
        TxtMontoApr.ForeColor = &H289556
    End If
    bBuscarLineas = False
    'WIOR 20131115 *************************************************************
    If Len(Trim(ActxCta.NroCuenta)) = 18 Then
        Dim oNCredito As COMNCredito.NCOMCredito
        Dim nMontoApr As Double
        Set oNCredito = New COMNCredito.NCOMCredito
        
        If TxtMontoApr.Text = "" Or TxtMontoApr.Text = "." Then
            nMontoApr = 0
        Else
            If IsNumeric(TxtMontoApr.Text) Then
                nMontoApr = CDbl(TxtMontoApr.Text)
            Else
                nMontoApr = 0
            End If
        End If
        '******COMENTADO POR APRI20180820 ERS061-2018******
'        If oNCredito.AplicaCuotaBalon(gsCodAge, lsTpoProdCod, nMontoApr, CInt(Mid(ActxCta.NroCuenta, 9, 1))) Then
'            If Not (chkCuotaBalon.Visible And chkCuotaBalon.Enabled) Then
'                chkCuotaBalon.Visible = True
'                chkCuotaBalon.value = 0
'                txtCuotaBalon.Visible = True
'                txtCuotaBalon.Text = "0"
'            End If
'        Else
'            txtCuotaBalon.Text = "0"
'            chkCuotaBalon.Visible = False
'            txtCuotaBalon.Visible = False
'        End If
        '*****END APRI
        Set oNCredito = Nothing
    End If
    'WIOR FIN ******************************************************************

End Sub

Private Sub txtMontoApr_GotFocus()
    fEnfoque TxtMontoApr
End Sub

Private Sub txtMontoApr_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(TxtMontoApr, KeyAscii)
     If KeyAscii = 13 Then
        Call CargarDatosProductoCrediticio 'ALPA 20150115
        Call MostrarLineas 'ALPA 20150115
        
        If gnAgenciaCredEval = 0 Then 'WIOR 20160225
            If SpnCuotasApr.Visible And SpnCuotasApr.Enabled Then
                SpnCuotasApr.SetFocus
            End If
        End If
     End If
End Sub

Private Sub txtMontoApr_LostFocus()
If Not fbGrabar Then 'WIOR 20150717
    If Trim(TxtMontoApr.Text) = "" Then
        TxtMontoApr.Text = "0.00"
    Else
        TxtMontoApr.Text = Format(TxtMontoApr.Text, "#,#0.00")
    End If
    Call CargarDatosProductoCrediticio 'ALPA 20150115
    Call MostrarLineas 'ALPA 20150115
End If 'WIOR 20150717
End Sub

' CMACICA_CSTS - 11/11/2003 - --------------------------------------------------------------------------------------------------------
Private Sub txtactivofijo_Change()
   lblActivo.Caption = Format(CDbl(IIf(Trim(lblActCirc.Caption) = "", "0", lblActCirc.Caption)) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
   lblPatrimonio.Caption = Format(CDbl(Val(lblActivo.Caption)) - CDbl(Val(lblPasivo.Caption)), "#0.00")
   lblPasPatrim.Caption = Format(CDbl(Val(lblPatrimonio.Caption)) + CDbl(Val(lblPasivo.Caption)), "#0.00")
End Sub

Private Sub txtactivofijo_LostFocus()
    txtactivofijo.Text = Format(IIf(Trim(txtactivofijo.Text) = "", 0, txtactivofijo.Text), "#0.00")
End Sub

Private Sub TxtBalEgrFam_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
End Sub

Private Sub TxtBalEgrFam_LostFocus()
    If Len(Trim(TxtBalEgrFam.Text)) = 0 Then
        TxtBalEgrFam.Text = "0.00"
    End If
    TxtBalEgrFam.Text = Format(TxtBalEgrFam.Text, "#0.00")
End Sub

Private Sub TxtBalIngFam_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
End Sub

Private Sub TxtBalIngFam_LostFocus()
    If Len(Trim(TxtBalIngFam.Text)) = 0 Then
        TxtBalIngFam.Text = "0.00"
    End If
    TxtBalIngFam.Text = Format(TxtBalIngFam.Text, "#0.00")
End Sub

Private Sub txtcompras_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
End Sub

Private Sub txtcompras_LostFocus()
    txtcompras.Text = Format(IIf(Trim(txtcompras.Text) = "", "0.00", txtcompras.Text), "#0.00")
End Sub

Private Sub txtcuentas_Change()
    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) _
            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) _
            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
            
    lblActivo.Caption = Format(CDbl(lblActCirc.Caption) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(Val(lblPasivo.Caption)) + CDbl(lblActivo.Caption), "#0.00")
    lblPatrimonio.Caption = Format(CDbl(Val(lblActivo.Caption)) - CDbl(Val(lblPasivo.Caption)), "#0.00")
    
End Sub

Private Sub txtcuentas_LostFocus()
    txtcuentas.Text = Format(IIf(Trim(txtcuentas.Text) = "", 0, txtcuentas.Text), "#0.00")
End Sub

Private Sub txtDisponible_Change()
    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) _
            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) _
            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
            
    lblActivo.Caption = Format(CDbl(Val(lblActCirc.Caption)) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(Val(lblPasivo.Caption)) + CDbl(Val(lblActivo.Caption)), "#0.00")
    lblPatrimonio.Caption = Format(CDbl(Val(lblActivo.Caption)) - CDbl(Val(lblPasivo.Caption)), "#0.00")
    
End Sub

Private Sub txtDisponible_LostFocus()
    txtDisponible.Text = Format(IIf(Trim(txtDisponible.Text) = "", 0, txtDisponible.Text), "#0.00")
End Sub

Private Sub txtEgreFam_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
    
    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00"))
End Sub

Private Sub txtEgreFam_LostFocus()
    If Len(Trim(txtEgreFam.Text)) > 0 Then
        txtEgreFam.Text = Format(txtEgreFam.Text, "#0.00")
    Else
        txtEgreFam.Text = "0.00"
    End If
End Sub

Private Sub TxtIngCon_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
    
    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
    
End Sub

Private Sub TxtIngCon_LostFocus()
    If Len(Trim(TxtIngCon.Text)) = 0 Then
        TxtIngCon.Text = "0.00"
    End If
    TxtIngCon.Text = Format(TxtIngCon.Text, "#0.00")
End Sub

Private Sub txtIngFam_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
    
    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _

End Sub

Private Sub txtIngFam_LostFocus()
    If Len(Trim(txtIngFam.Text)) > 0 Then
        txtIngFam.Text = Format(txtIngFam.Text, "#0.00")
    Else
        txtIngFam.Text = "0.00"
    End If
End Sub

Private Sub txtInventario_Change()
    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) _
            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) _
            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
            
    lblActivo.Caption = Format(CDbl(lblActCirc.Caption) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(Val(lblPasivo.Caption)) + CDbl(Val(lblActivo.Caption)), "#0.00")
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(Val(lblPasivo.Caption)), "#0.00")
        
End Sub

Private Sub txtInventario_LostFocus()
    txtInventario.Text = Format(IIf(Trim(txtInventario.Text) = "", 0, txtInventario.Text), "#0.00")
End Sub

Private Sub txtOtroIng_Change()
    
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
    
    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00"))
            
End Sub

Private Sub txtOtroIng_LostFocus()
    If Len(Trim(txtOtroIng.Text)) > 0 Then
        txtOtroIng.Text = Format(txtOtroIng.Text, "#0.00")
    Else
        txtOtroIng.Text = "0.00"
    End If
End Sub

Private Sub txtOtrosEgresos_Change()
    LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
End Sub

Private Sub txtOtrosEgresos_LostFocus()
    txtOtrosEgresos.Text = Format(IIf(Trim(txtOtrosEgresos.Text) = "", "0.00", txtOtrosEgresos.Text), "#0.00")
End Sub

Private Sub txtOtrosPrest_Change()
    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) _
            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) _
            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
            
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")
    
End Sub

Private Sub txtOtrosPrest_LostFocus()
    If Len(Trim(txtOtrosPrest.Text)) > 0 Then
        txtOtrosPrest.Text = Format(txtOtrosPrest.Text, "#0.00")
    Else
        txtOtrosPrest.Text = "0.00"
    End If
End Sub

Private Sub txtPrestCmact_Change()
    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) _
            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) _
            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
            
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")

End Sub

Private Sub txtPrestCmact_LostFocus()
    If Len(Trim(txtPrestCmact.Text)) > 0 Then
        txtPrestCmact.Text = Format(txtPrestCmact.Text, "#0.00")
    Else
        txtPrestCmact.Text = "0.00"
    End If
End Sub

Private Sub txtProveedores_Change()
    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) _
            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) _
            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
    
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")
End Sub

Private Sub txtProveedores_LostFocus()
    If Len(Trim(txtProveedores.Text)) > 0 Then
        txtProveedores.Text = Format(txtProveedores.Text, "#0.00")
    Else
        txtProveedores.Text = "0.00"
    End If
End Sub

Private Sub txtrecuperacion_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

End Sub

Private Sub txtrecuperacion_LostFocus()
    txtrecuperacion.Text = Format(IIf(Trim(txtrecuperacion.Text) = "", "0.00", txtrecuperacion.Text), "#0.00")
End Sub

'MAVM 25102010 ***
Private Sub TxtTasaGracia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CDbl(txtTasInt.Text) < CDbl(lnTasaInicial) Then
            If CDbl(txtTasInt.Text) <= Val(TxtTasaGracia.Text) Then
                nTasaGrac = Val(TxtTasaGracia.Text)
            End If
        End If
    End If
End Sub
'***

Private Sub txtVentas_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
End Sub

Private Sub txtVentas_LostFocus()
    txtVentas.Text = Format(IIf(Trim(txtVentas.Text) = "", "0.00", txtVentas.Text), "#0.00")
End Sub

'------------------------------------------------------------------------------------------------------------------------------------
Sub CargarLineaForAprobacion(ByVal prsLineas As ADODB.Recordset)

'Dim oLineas As COMDCredito.DCOMLineaCredito
    
    'Carga Lineas de Credito en arbol ---------------------------------------------------------------------------
        
    '    ObtenerLineaCred
                
        txtBuscarLinea.sTitulo = "Lineas de Crédito"
        txtBuscarLinea.psRaiz = "Lineas de Crédito"

        'Dim RLineaProducto As ADODB.Recordset
        'Set oLineas = New COMDCredito.DCOMLineaCredito
        'Set RLineaProducto = oLineas.RecuperaLineasProductoArbol(Mid(ActxCta.NroCuenta, 6, 3), Mid(ActxCta.NroCuenta, 9, 1))
        'Set oLineas = Nothing
        
        txtBuscarLinea.rs = prsLineas 'RLineaProducto
        'Set RLineaProducto = Nothing
        
        txtBuscarLinea.Text = ""
        
        If sLineaTmp <> "" Then
           txtBuscarLinea.Text = IIf(Mid(sLineaTmp, 6, 1) = "1", "CP1-" + sLineaTmp, "LP2-" + sLineaTmp)
           txtBuscarLinea_EmiteDatos
        '--------------------------------------------------------------------------------------------
'        Else
'           Txtinteres.Text = "0.00"
'           LblInteres.Caption = "0.00"
'           TxtTasaGracia.Text = "0.00"
'           LblTasaGracia.Caption = "0.00"
        End If
End Sub

'Sub ObtenerLineaCred()
'    Dim oLinea As COMDCredito.DCOMLineaCredito
'
'    Set oLinea = New COMDCredito.DCOMLineaCredito
'    sLineaTmp = oLinea.ObtenerLinea(ActxCta.NroCuenta)
'    Set oLinea = Nothing
'End Sub

'By Capi 10042008
Private Sub actxCtaAho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCalend.Enabled = False
        If ValidaCuenta(Me.actxCtaAho.NroCuenta) = False Then Exit Sub
        If Mid(Me.ActxCta.NroCuenta, 9, 1) <> Mid(Me.actxCtaAho.NroCuenta, 9, 1) Then
            MsgBox "Moneda de Cuentas de Credito y Ahorro no coinciden", vbInformation, "Aviso"
            Exit Sub
        End If
        Dim sCta As String
        sCta = actxCtaAho.NroCuenta
        ObtieneDatosCuenta sCta
    End If
 End Sub

'By Capi 10042008
Private Function ValidaCuenta(ByVal psCtaCod As String) As Boolean
    If Len(Trim(psCtaCod)) <> 18 Then
        MsgBox "Ingrese un Numero de Cuenta correcta", vbInformation, "Aviso"
        ValidaCuenta = False
    Else
        ValidaCuenta = True
    End If
End Function

'By Capi 10042008
Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim ssql As String, sMoneda As String, sPersona As String
Dim nRow As Long

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing

If Not (rsCta.EOF And rsCta.BOF) Then
    CmdCalend.Enabled = True
    If TxtMontoApr.Enabled Then
        If gnAgenciaCredEval = 0 Then TxtMontoApr.SetFocus 'JUEZ 20121218
    Else
        If SpnCuotasApr.Enabled = True Then
            If gnAgenciaCredEval = 0 Then SpnCuotasApr.SetFocus 'JUEZ 20121218
        Else
            If gnAgenciaCredEval = 0 Then SpnPlazo.SetFocus 'JUEZ 20121218
        End If
    End If
Else
    MsgBox "Cuenta no existe", vbInformation, "Aviso"
    lblTercero.Caption = ""
    actxCtaAho.SetFocusCuenta
End If
End Sub

'MAVM 28102010 ***
Private Sub GenerarFechaPago()
    If OptTPer(0).value = True Then
        txtFechaFija.Text = CDate(CDate(TxtFecDesemb.Text) + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
    End If
    
    If OptTPer(1).value = True Then
        txtFechaFija.Text = CDate(CDate(TxtFecDesemb.Text))
        TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaFija.Text)), 2)
                                    
        If Month(CDate(TxtFecDesemb.Text)) = Month(CDate(txtFechaFija.Text)) And Year(CDate(TxtFecDesemb.Text)) = Year(txtFechaFija.Text) Then
            ChkProxMes.value = 0
        Else
            ChkProxMes.value = 1
        End If
    End If
End Sub

'FRHU 20140402 ERS026-2014 RQ14131
Private Function ValidarUltimaActualizacion(ByVal psCtaCod As String) As Boolean
    Dim oCred As New COMDCredito.DCOMCredito
    Dim oRS As ADODB.Recordset
    Dim sNomPers As String
    Set oRS = oCred.ObtenerUltimaActualizacionXCredito(psCtaCod)
    sNomPers = ""
    If Not oRS.EOF And Not oRS.BOF Then
        Do While Not oRS.EOF
            sNomPers = sNomPers & oRS!cPersNombre & vbNewLine
            oRS.MoveNext
        Loop
        MsgBox "Nombre de los clientes que requieren actualizarse: " & vbNewLine & _
               sNomPers & vbNewLine & _
               "No se puede aprobar el credito debido a que estos clientes tienen información desactualizada (mas de 6 meses de antiguedad)," & vbNewLine & _
               "debe comunicarse con el analista para actualizar la información del cliente.", vbInformation
        
        ValidarUltimaActualizacion = True
    Else
        ValidarUltimaActualizacion = False
    End If
End Function
'FIN FRHU 20140402
'ALPA20140619************************************
Private Sub txtMontoMivivienda_Change()
    Dim nMontoMV As Double
    If Len(Trim(ActxCta.NroCuenta)) = 18 Then
        
        If txtMontoMivivienda.Text = "" Or txtMontoMivivienda.Text = "." Then
            nMontoMV = 0
        Else
            If IsNumeric(txtMontoMivivienda.Text) Then
                nMontoMV = CDbl(txtMontoMivivienda.Text)
            Else
                nMontoMV = 0
            End If
        End If
    End If
End Sub

Private Sub txtMontoMivivienda_GotFocus()
    fEnfoque txtMontoMivivienda
End Sub

Private Sub txtMontoMivivienda_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(txtMontoMivivienda, KeyAscii, 10)
End Sub

Private Sub txtMontoMivivienda_LostFocus()
    If Trim(txtMontoMivivienda.Text) = "" Then
        txtMontoMivivienda.Text = "0.00"
    Else
        txtMontoMivivienda.Text = Format(txtMontoMivivienda.Text, "#0.00")
    End If
End Sub
'ALPA 20150115*************************************************************
Private Sub cmbCampana_Click()
    Call CargarDatosProductoCrediticio
End Sub



Private Sub MostrarLineas()
    txtBuscarLinea.Text = ""
    lblLineaDesc.Caption = ""
    If Val(TxtMontoApr.Text) > 0 Then
    Dim oLineas As COMDCredito.DCOMLineaCredito
    Dim lrsLineas As ADODB.Recordset
    Set lrsLineas = New ADODB.Recordset
    
    Set oLineas = New COMDCredito.DCOMLineaCredito
    Set lrsLineas = oLineas.RecuperaLineasProductoArbol(lsTpoCredCod, Mid(ActxCta.NroCuenta, 9, 1), , Mid(ActxCta.NroCuenta, 4, 2), IIf(Trim(SpnPlazo.valor) = "", 0, SpnPlazo.valor), CDbl(TxtMontoApr.Text), IIf(Trim(SpnCuotasApr.valor) = "", 0, SpnCuotasApr.valor), IIf(Trim(txtTasInt.Text) = "", 0, txtTasInt.Text), IIf(Trim(TxtDiasGracia.Text) = "", 0, TxtDiasGracia.Text), gdFecSis, LblCodCli.Caption, Trim(Right(IIf(CmbCampana.Text = "", 0, CmbCampana.Text), 10))) 'arlo20200429 add cmbCampana
    Set oLineas = Nothing
    txtBuscarLinea.rs = lrsLineas
    End If
End Sub
Private Sub txtTasInt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call MostrarLineas
    Call CargaDatosLinea
End If
End Sub
Private Sub txtTasInt_LostFocus()
    If Trim(txtTasInt.Text) = "" Then
        txtTasInt.Text = "0.0000"
    Else
        'ALPA 20141030************************************
'        If CDbl(txtTasInt.Text) <= lnTasaInicial Then
'            txtTasInt.Text = lnTasaInicial
'        End If
        '*************************************************
        txtTasInt.Text = Format(txtTasInt.Text, "#0.0000")
        txtInteresTasa.Text = Format(txtTasInt.Text, "#0.0000")
        Call MostrarLineas
        Call CargaDatosLinea
    End If
End Sub
Private Sub txtInteresTasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtInteresTasa.Text = Format(txtInteresTasa.Text, "#0.0000")
        KeyAscii = NumerosDecimales(txtInteresTasa, KeyAscii, , 4)
        txtTasInt.Text = txtInteresTasa.Text
        Call MostrarLineas
    End If
End Sub

Private Sub txtInteresTasa_LostFocus()
    If Trim(txtInteresTasa.Text) = "" Then
        txtInteresTasa.Text = "0.0000"
    Else
        If CDbl(txtInteresTasa.Text) >= lnTasaInicial Then
            txtInteresTasa.Text = lnTasaInicial
        End If
        txtInteresTasa.Text = Format(txtInteresTasa.Text, "#0.0000")
        txtTasInt.Text = txtInteresTasa.Text
        Call MostrarLineas
    End If
End Sub
'**************************************************************************

Private Function ObtenerColocacPermisoAprobacion() As String
        Dim objCargos As COMDCredito.DCOMNivelAprobacion
        Dim oRS As ADODB.Recordset
        Set objCargos = New COMDCredito.DCOMNivelAprobacion
        Set oRS = New ADODB.Recordset
        Set oRS = objCargos.GetCargarColocacPermisoAprobacionConsulta(ActxCta.NroCuenta, Format(gdFecSis, "YYYYMMDD"))
        If Not (oRS.BOF Or oRS.EOF) Then
            If oRS!nEstado = 1 Then
                ObtenerColocacPermisoAprobacion = "1-Coordinar con el area respectiva para aprobar la petición"
            Else
                If txtTasInt.Text = oRS!nTasaApr Then
                    ObtenerColocacPermisoAprobacion = "2-La petición ha sido aprobada, tasa aprobada es:     " & oRS!nTasaApr
                Else
                    ObtenerColocacPermisoAprobacion = "3-La petición ha sido aprobada con otra tasa:         " & oRS!nTasaApr
                End If
            End If
            
        Else
            ObtenerColocacPermisoAprobacion = "0-No existe pedido de autorización de tasa"
        End If
End Function

'WIOR 20151223 ***
Private Sub cmdMIVIVIENDA_Click()
If fbMIVIVIENDA Then
    Call frmCredMiViviendaDatos.Inicio(Trim(ActxCta.NroCuenta), fArrMIVIVIENDA, lsTpoCredCod, Right(CmbDestino.Text, 3), False)
'    If IsArray(fArrMIVIVIENDA) Then
'        TxtMontoApr.Text = CDbl(fArrMIVIVIENDA(3))
'    Else
'        TxtMontoApr.Text = "0.00"
'    End If
'    Call TxtMontoApr_Change
'    Call txtMontoApr_KeyPress(13)
'    Call txtMontoApr_LostFocus
'    TxtMontoApr.Enabled = False
'    FraTpoCalend.Enabled = True
'    ChkMiViv.Enabled = False
'    ChkMiViv.value = 1
    cmdMIVIVIENDA.Enabled = True
End If
End Sub
'WIOR FIN ********

'WIOR 20160224 ***
Private Sub ObtenerTasaSeguroDesg(Optional ByVal pbMostrar As Boolean = False)
Dim oNCredito As COMNCredito.NCOMCredito
Set oNCredito = New COMNCredito.NCOMCredito
fnTasaSegDes = oNCredito.ObtenerTasaSeguroDesg(ActxCta.NroCuenta, gdFecSis, fnCantAfiliadosSegDes)

If pbMostrar Then
    If fnCantAfiliadosSegDes > 0 Then
        MsgBox "Tasa del Seguro Desgravamen: " & Format(fnTasaSegDes, "#0.000") & "%" & Chr(10) & fnCantAfiliadosSegDes & " Afiliado(s)", vbInformation, "Tasas"
    End If
End If

Set oNCredito = Nothing
End Sub
'WIOR FIN ********
'JOEP20190201 CP
Private Sub bValiExistCheckList()
Dim rsActivar As ADODB.Recordset
Dim pnActivar As Integer
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito
Set rsActivar = oCredito.getActivarCheckList(lsTpoProdCod)
If Not (rsActivar.BOF And rsActivar.EOF) Then
    If rsActivar!Activar = 1 Then
        cmdCheckList.Enabled = True
    Else
        cmdCheckList.Enabled = False
    End If
End If
Set oCredito = Nothing
RSClose rsActivar
End Sub

Private Function CP_CondicionesApro(ByVal cCodProd As String, ByVal nMonedad As String, ByVal pnMonto As String, _
ByVal pnCuota As String, ByVal pnPlazo As String, ByVal pnDestino As String, ByVal pnPeriodo As String, ByVal pnTpCredito As String, ByVal pnModulo As Integer) As Boolean
    Dim oDCred As COMDCredito.DCOMCredito
    Dim rsDefaut As ADODB.Recordset
    Set oDCred = New COMDCredito.DCOMCredito
    CP_CondicionesApro = True
    
    Set rsDefaut = oDCred.CatalogoCondicionSeg(Trim(ActxCta.NroCuenta), cCodProd, nMonedad, pnMonto, pnCuota, pnPlazo, pnDestino, pnPeriodo, pnTpCredito, pnModulo, Trim(Right(IIf(CmbCampana.Text = "", 0, CmbCampana.Text), 10))) 'arlo20200429 add cmbCampana
        
    If Not (rsDefaut.BOF And rsDefaut.EOF) Then
        If rsDefaut!Mensaje <> "" Then
            MsgBox rsDefaut!Mensaje & rsDefaut!ValorMax, vbInformation, "Aviso"
            CP_CondicionesApro = False
        End If
    End If
End Function
Private Function CP_ValidaRangoFechaApro(ByVal nCodProd As String, ByVal nParCod As Integer, ByVal nTpPago As Integer, ByVal dFechaPago As String) As Boolean
Dim obj As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
On Error GoTo ErrorCP_ValidaRangoFechaApro
CP_ValidaRangoFechaApro = True
    Set obj = New COMDCredito.DCOMCredito
    Set rs = obj.getCatalogoRangoFecha(nCodProd, nParCod, nTpPago, dFechaPago)
        
    If Not (rs.BOF And rs.EOF) Then
        If rs!Mensaje <> "" Then
            MsgBox rs!Mensaje, vbInformation, "AVISO"
            CP_ValidaRangoFechaApro = False
        End If
    End If
Set obj = Nothing
RSClose rs
    Exit Function
ErrorCP_ValidaRangoFechaApro:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
'JOEP20190201 CP
