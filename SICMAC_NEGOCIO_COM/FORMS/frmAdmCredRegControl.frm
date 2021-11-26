VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdmCredRegControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Control de Créditos"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   ForeColor       =   &H00FF0000&
   Icon            =   "frmAdmCredRegControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "Campaña"
      ForeColor       =   &H00FF0000&
      Height          =   680
      Left            =   2880
      TabIndex        =   61
      Top             =   0
      Width           =   3015
      Begin VB.Label lblCampania 
         Alignment       =   2  'Center
         Caption         =   "lblcampaña"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   285
         Width           =   2775
      End
   End
   Begin SICMACT.ActXCodCta_New ActXCodCta 
      Height          =   735
      Left            =   120
      TabIndex        =   56
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Texto           =   "Crédito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Frame frAdmCred 
      Height          =   4200
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Width           =   10695
      Begin VB.Frame Frame7 
         Height          =   660
         Left            =   120
         TabIndex        =   58
         Top             =   3400
         Width           =   2655
         Begin VB.CheckBox chkTpoRevision 
            Caption         =   "Tipo de Revisión"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox cboTipoRev 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdChkList 
         Caption         =   "CheckList"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   57
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   300
         Left            =   7320
         TabIndex        =   55
         Top             =   3460
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   8400
         TabIndex        =   54
         Top             =   3460
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   300
         Left            =   9480
         TabIndex        =   53
         Top             =   3460
         Width           =   1095
      End
      Begin TabDlg.SSTab tabRegistro 
         Height          =   3165
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5583
         _Version        =   393216
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Observaciones"
         TabPicture(0)   =   "frmAdmCredRegControl.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Exoneraciones"
         TabPicture(1)   =   "frmAdmCredRegControl.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frameExoneraciones"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Autorizaciones"
         TabPicture(2)   =   "frmAdmCredRegControl.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FrameAuto"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame6 
            Height          =   2655
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   10215
            Begin VB.CommandButton cmdLevantarObs 
               Caption         =   "Levantar"
               Height          =   300
               Left            =   8640
               TabIndex        =   52
               Top             =   600
               Width           =   1380
            End
            Begin VB.TextBox txtObs 
               Height          =   285
               Left            =   1560
               MaxLength       =   100
               TabIndex        =   49
               Top             =   240
               Width           =   6915
            End
            Begin VB.CommandButton cmdAgrObs 
               Caption         =   "Agrega"
               Height          =   300
               Left            =   8640
               TabIndex        =   48
               Top             =   240
               Width           =   1380
            End
            Begin VB.CommandButton cmdQuiObs 
               Caption         =   "Quitar"
               Height          =   300
               Left            =   8640
               TabIndex        =   47
               Top             =   540
               Visible         =   0   'False
               Width           =   1380
            End
            Begin SICMACT.FlexEdit feObservaciones 
               Height          =   1725
               Left            =   240
               TabIndex        =   50
               Top             =   600
               Width           =   8250
               _ExtentX        =   14552
               _ExtentY        =   3043
               Cols0           =   5
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Observacion-cCtaCod-OK-A"
               EncabezadosAnchos=   "400-6700-0-400-500"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-3-4"
               ListaControles  =   "0-0-0-4-3"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "L-L-C-C-L"
               FormatosEdit    =   "0-0-0-0-0"
               TextArray0      =   "#"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               lbBuscaDuplicadoText=   -1  'True
               ColWidth0       =   405
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Observaciones :"
               Height          =   195
               Left            =   240
               TabIndex        =   51
               ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
               Top             =   240
               Width           =   1155
            End
         End
         Begin VB.Frame frameObservaciones 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   40
            Top             =   360
            Width           =   11175
            Begin VB.TextBox txtObservaciones 
               Height          =   285
               Left            =   1560
               MaxLength       =   100
               TabIndex        =   43
               Top             =   240
               Width           =   6670
            End
            Begin VB.CommandButton cmdAgregaObs 
               Caption         =   "Agrega"
               Height          =   375
               Left            =   9960
               TabIndex        =   42
               Top             =   480
               Width           =   900
            End
            Begin VB.CommandButton cmdBorraObs 
               Caption         =   "Borrar"
               Height          =   375
               Left            =   9960
               TabIndex        =   41
               Top             =   1440
               Width           =   900
            End
            Begin SICMACT.FlexEdit FlexObs 
               Height          =   1845
               Left            =   240
               TabIndex        =   44
               Top             =   720
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   3254
               Cols0           =   3
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Observacion-cCtaCod"
               EncabezadosAnchos=   "400-7000-0"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X"
               ListaControles  =   "0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "L-L-C"
               FormatosEdit    =   "0-0-0"
               TextArray0      =   "#"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   405
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Observaciones :"
               Height          =   195
               Left            =   240
               TabIndex        =   45
               ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
               Top             =   240
               Width           =   1155
            End
         End
         Begin VB.Frame frameExoneraciones 
            Height          =   2535
            Left            =   -74880
            TabIndex        =   32
            ToolTipText     =   "CargarNivelAutorizacion"
            Top             =   480
            Width           =   10215
            Begin VB.CommandButton cmdSelectApoExo 
               Caption         =   "..."
               Height          =   290
               Left            =   9840
               TabIndex        =   63
               Top             =   250
               Width           =   255
            End
            Begin VB.ComboBox cboExoneraciones 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   240
               Width           =   3370
            End
            Begin VB.ComboBox cboNivelExonera 
               Height          =   315
               Left            =   6480
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   240
               Width           =   3370
            End
            Begin VB.CommandButton cmdAgregarExo 
               Caption         =   "Agregar"
               Height          =   300
               Left            =   120
               TabIndex        =   34
               Top             =   2040
               Width           =   1800
            End
            Begin VB.CommandButton cmdQuitaExo 
               Caption         =   "Quitar"
               Height          =   300
               Left            =   1920
               TabIndex        =   33
               Top             =   2040
               Width           =   1800
            End
            Begin SICMACT.FlexEdit FEExoneraciones 
               Height          =   1335
               Left            =   120
               TabIndex        =   35
               Top             =   600
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   2355
               Cols0           =   6
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Exoneración-Nivel de Exoneración-Apoderado 1-Apoderado 2-Apoderado 3"
               EncabezadosAnchos=   "300-2000-2000-2000-2000-2000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X-X-X"
               ListaControles  =   "0-0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L-L-L-L"
               FormatosEdit    =   "0-0-0-0-0-0"
               TextArray0      =   "#"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin VB.Label Label12 
               Caption         =   "Exoneración:"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   285
               Width           =   975
            End
            Begin VB.Label Label16 
               Caption         =   "Nivel de Exoneración:"
               Height          =   255
               Left            =   4800
               TabIndex        =   38
               Top             =   285
               Width           =   1695
            End
         End
         Begin VB.Frame FrameAuto 
            Height          =   2535
            Left            =   -74880
            TabIndex        =   24
            Top             =   480
            Width           =   10215
            Begin VB.CommandButton cmdSelecApoAut 
               Caption         =   "..."
               Height          =   290
               Left            =   9800
               TabIndex        =   64
               Top             =   250
               Width           =   255
            End
            Begin VB.ComboBox cboAutorizacion 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   240
               Width           =   3370
            End
            Begin VB.ComboBox cboNivAutorizacion 
               Height          =   315
               Left            =   6440
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   240
               Width           =   3370
            End
            Begin VB.CommandButton cmdAgregarAutorizacion 
               Caption         =   "Agregar"
               Height          =   300
               Left            =   120
               TabIndex        =   26
               Top             =   2040
               Width           =   1800
            End
            Begin VB.CommandButton cmdQuitarAutorizacion 
               Caption         =   "Quitar"
               Height          =   300
               Left            =   1920
               TabIndex        =   25
               Top             =   2040
               Width           =   1800
            End
            Begin SICMACT.FlexEdit FEAutorizaciones 
               Height          =   1335
               Left            =   120
               TabIndex        =   29
               Top             =   600
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   2355
               Cols0           =   6
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Autorización-Nivel de Autorización-Apoderado 1-Apoderado 2-Apoderado 3"
               EncabezadosAnchos=   "300-2000-2000-2000-2000-2000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-3-X-X"
               ListaControles  =   "0-0-0-3-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L-L-L-L"
               FormatosEdit    =   "0-0-0-0-0-0"
               TextArray0      =   "#"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin VB.Label Label25 
               Caption         =   "Autorización:"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   280
               Width           =   975
            End
            Begin VB.Label Label24 
               Caption         =   "Nivel de Autorización:"
               Height          =   255
               Left            =   4800
               TabIndex        =   30
               Top             =   285
               Width           =   1695
            End
         End
      End
   End
   Begin VB.Frame frameConsulta 
      Caption         =   " Módulos de consulta "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   8280
      TabIndex        =   15
      Top             =   240
      Width           =   2535
      Begin VB.CommandButton cmdFinalidad 
         Caption         =   "Finalidad"
         Height          =   300
         Left            =   120
         TabIndex        =   65
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton cmdInfRiesg 
         Caption         =   "Informe de Riesgos"
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   2410
         Width           =   2295
      End
      Begin VB.CommandButton cmdPoliza 
         Caption         =   "Póliza"
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   1960
         Width           =   2295
      End
      Begin VB.CommandButton cmdGarant 
         Caption         =   "Garantía"
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   1520
         Width           =   2295
      End
      Begin VB.CommandButton cmdConsultPers 
         Caption         =   "Consulta de Persona"
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdPosClient 
         Caption         =   "Posición de Cliente"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   650
         Width           =   2295
      End
      Begin VB.CommandButton cmdDupDoc 
         Caption         =   "Duplicados"
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Datos de la Garantía "
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   10695
      Begin SICMACT.FlexEdit feGarantias 
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   2355
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Garantía-Gravament-Comercial-Realización-Disponible-Titular-Nº Doc-Nº Garantía-Legal-Póliza-cNumPoliza"
         EncabezadosAnchos=   "300-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos de Aprobación "
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   8055
      Begin VB.Label lblFechaPago 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "1º Venc.:"
         Height          =   255
         Left            =   6120
         TabIndex        =   11
         Top             =   375
         Width           =   735
      End
      Begin VB.Label lblNCuotas 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Nº Cuotas:"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "%"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   380
         Width           =   255
      End
      Begin VB.Label lblTem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "TEM:"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   380
         Width           =   375
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblMoneda 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   380
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Créditos "
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8055
      Begin SICMACT.FlexEdit feListaRelac 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2355
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nombre de Cliente-Relación-cPersCod"
         EncabezadosAnchos=   "300-5000-2200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmAdmCredRegControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmAdmCredRegControl
'** Descripción : Formulario que permite realizar el control del crédito
'** Creación    : RECO, 20150421 - ERS010-2015
'**********************************************************************************************

Option Explicit
Dim nCargoNivApr As String
Dim nCanFirmasAut As Integer
Dim nCanFirmasExo As Integer
Dim sPersCargoExo As String
Dim sPersCargoAuto As String
Dim sTpoCredCod As String
Dim vDatos() As Variant
Dim nContador As Integer
Dim nTpoOpe As Integer
Dim sCadSalida As String
'JOEP20190131 CP
Dim lsCategoria As String
Dim lsTipoCredito As String
Dim lsTipoProdCod As String
'JOEP20190131 CP

Public Function Inicio(ByVal psTitulo As String, ByVal psCtaCod As String, ByVal pnTpoOpe As Integer) As String
    nContador = 0
    Call cmdCancelar_Click
    ActXCodCta.NroCuenta = psCtaCod
    Me.Caption = Me.Caption & " " & psTitulo
    feObservaciones.EncabezadosAnchos = IIf(pnTpoOpe = 1, "400-6700-0-400-0", "400-6700-0-400-500")
    If pnTpoOpe = 1 Then Call ActXCodCta_KeyPress(vbKeyReturn)
    If pnTpoOpe = 2 Then Call HabilitaControles(False)
    nTpoOpe = pnTpoOpe
    Screen.MousePointer = 0

    
    '*****************APRI20170630 TI-ERS025-2017*****************
    Dim rsp As ADODB.Recordset
    Dim oPers As COMDPersona.UCOMPersona
    Set oPers = New COMDPersona.UCOMPersona
    Set rsp = oPers.ObtenerVinculadoRiesgoUnico("", psCtaCod, 1)
    
        If Not (rsp.BOF And rsp.EOF) Then
            If rsp.RecordCount = 1 Then
               If rsp!nTotal = 1 Then
                     MsgBox "El vinculado " & Trim(rsp!cPersNombre) & " tiene un crédito que se encuentra en " & rsp!cEstado & ".", vbOKOnly, "Aviso"
                Else
                     MsgBox "El vinculado " & Trim(rsp!cPersNombre) & " tiene " & rsp!nTotal & " créditos que se encuentran en mora.", vbOKOnly, "Aviso"
                End If
            ElseIf rsp.RecordCount > 1 Then
             MsgBox "El cliente tiene vinculados en persona que se encuentra en mora.", vbOKOnly, "Aviso"
            End If
        End If
        
    Set oPers = Nothing
    '*****************END APRI20170630 ***************************
    
    Me.Show 1
    nContador = 0
    Inicio = sCadSalida
End Function

Private Sub CargarExoneraciones()
    Dim loExonera As New COMDCredito.DCOMNivelAprobacion
    Dim lrDatos As New ADODB.Recordset
    Set lrDatos = loExonera.RecuperaTpoExoneraciones
    If Not (lrDatos.BOF And lrDatos.EOF) Then
        Dim nIndex As Integer
        For nIndex = 1 To lrDatos.RecordCount
            cboExoneraciones.AddItem Trim(lrDatos!cExoneraDesc) & Space(100) & Trim((lrDatos!cExoneraCod))
            lrDatos.MoveNext
        Next
    End If
    cboExoneraciones.ListIndex = 0
    Set lrDatos = Nothing
End Sub

Private Sub CargarAutorizaciones()
    Dim loExonera As New COMDCredito.DCOMNivelAprobacion
    Dim lrDatos As New ADODB.Recordset
    Set lrDatos = loExonera.RecuperaTpoAutorizaciones
    If Not (lrDatos.BOF And lrDatos.EOF) Then
        Dim nIndex As Integer
        For nIndex = 1 To lrDatos.RecordCount
            cboAutorizacion.AddItem Trim(lrDatos!cExoneraDesc) & Space(100) & Trim((lrDatos!cExoneraCod))
            lrDatos.MoveNext
        Next
    End If
    cboAutorizacion.ListIndex = 0
    Set lrDatos = Nothing
End Sub

Private Sub CargarNivelExoneracion()
    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
    Dim lrDatos As New ADODB.Recordset
    Dim lnCanFirmas As Integer
    
    Set lrDatos = oDNiv.RecuperaNivelesExoneracion(Trim(Right(cboExoneraciones.Text, 8)))
    
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        Dim nIndex As Integer
        cboNivelExonera.Clear
        nCanFirmasExo = lrDatos!nNumCantFirmas
        For nIndex = 1 To lrDatos.RecordCount
            cboNivelExonera.AddItem Trim(lrDatos!cNivAprDesc) & Space(100) & Trim((lrDatos!cNivAprCod))
            lrDatos.MoveNext
        Next
        cboNivelExonera.ListIndex = 0
    End If
End Sub

Private Sub CargarNivelAutorizacion()
    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
    Dim lrDatos As New ADODB.Recordset
    Dim lnCanFirmas As Integer
    
    Set lrDatos = oDNiv.RecuperaNivelesExoneracion(Trim(Right(cboAutorizacion.Text, 8)))
    
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        Dim nIndex As Integer
        cboNivAutorizacion.Clear
        nCanFirmasAut = lrDatos!nNumCantFirmas
        For nIndex = 1 To lrDatos.RecordCount
            cboNivAutorizacion.AddItem Trim(lrDatos!cNivAprDesc) & Space(100) & Trim((lrDatos!cNivAprCod))
            lrDatos.MoveNext
        Next
        cboNivAutorizacion.ListIndex = 0
    End If
    Set lrDatos = Nothing
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call CargaDatos(ActXCodCta.NroCuenta)
    End If
End Sub

Private Sub cboExoneraciones_Click()
    Call CargarNivelExoneracion
End Sub

Private Sub cboAutorizacion_Click()
    Call CargarNivelAutorizacion
End Sub

Private Sub cboNivelExonera_Click()
    nContador = nContador + 1
End Sub

Private Sub cmdAgregarAutorizacion_Click()
    Dim nIndex As Integer, nCantUser As Integer
    Dim msj As String
    
    msj = ValidaExisteElemento(FEAutorizaciones, Me.cboAutorizacion.Text, 2)
    If msj <> "" Then MsgBox msj, vbInformation, "Alerta": Exit Sub
    FEAutorizaciones.AdicionaFila
    FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 1) = Me.cboAutorizacion.Text
    FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 2) = Me.cboNivAutorizacion.Text

    nCantUser = UBound(vDatos, 2)
    If nCantUser > 0 Then
        If nCantUser > 1 Then
            If nCantUser > 2 Then
                FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 3) = vDatos(2, 0) & Space(75) & vDatos(1, 0)
                FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 4) = vDatos(2, 1) & Space(75) & vDatos(1, 1)
                FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 5) = vDatos(2, 2) & Space(75) & vDatos(1, 2)
            Else
                FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 3) = vDatos(2, 0) & Space(75) & vDatos(1, 0)
                FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 4) = vDatos(2, 1) & Space(75) & vDatos(1, 1)
            End If
        Else
            FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 3) = vDatos(2, 0) & Space(75) & vDatos(1, 0)
        End If
    End If
    cmdAgregarAutorizacion.Enabled = False
End Sub

Private Sub cmdAgregarExo_Click()
    Dim nIndex As Integer, nCantUser As Integer
    Dim msj As String
    
    msj = ValidaExisteElemento(FEExoneraciones, Me.cboExoneraciones.Text, 1)
    If msj <> "" Then MsgBox msj, vbInformation, "Alerta": Exit Sub
    FEExoneraciones.AdicionaFila
    FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 1) = Me.cboExoneraciones.Text
    FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 2) = Me.cboNivelExonera.Text

    nCantUser = UBound(vDatos, 2)
    If nCantUser > 0 Then
        If nCantUser > 1 Then
            If nCantUser > 2 Then
                FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 3) = vDatos(2, 0) & Space(75) & vDatos(1, 0)
                FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 4) = vDatos(2, 1) & Space(75) & vDatos(1, 1)
                FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 5) = vDatos(2, 2) & Space(75) & vDatos(1, 2)
            Else
                FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 3) = vDatos(2, 0) & Space(75) & vDatos(1, 0)
                FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 4) = vDatos(2, 1) & Space(75) & vDatos(1, 1)
            End If
        Else
            FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 3) = vDatos(2, 0) & Space(75) & vDatos(1, 0)
        End If
    End If
    cmdAgregarExo.Enabled = False
End Sub

Private Function CargarCargoNivExoAuto(ByVal cboControlNivelExoAut As ComboBox, ByVal cboControlExoAut As ComboBox, ByVal pnTpoOpe As Integer) As ADODB.Recordset
    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
    Dim lrCargo As New ADODB.Recordset
    Dim lrDatos As New ADODB.Recordset
    Dim lrDRHCargo As New COMDPersona.UCOMAcceso
    Dim lsCadCargo As String
    Dim lnFirmas As Integer

    Set lrCargo = oDNiv.RecuperaNivelesAprValores(Trim(Right(cboControlNivelExoAut.Text, 8)))
    Set lrDatos = oDNiv.RecuperaCargosNivelesExoneracion(Trim(Right(cboControlExoAut.Text, 8)), Trim(Right(cboControlNivelExoAut.Text, 8)))

    If Not (lrCargo.EOF And lrCargo.BOF) Then
        Dim i As Integer
        For i = 1 To lrCargo.RecordCount
            lsCadCargo = lsCadCargo & lrCargo!cValorCod & ","
            lrCargo.MoveNext
        Next
    End If
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        If pnTpoOpe = 1 Then
            nCanFirmasExo = lrDatos!nNumCantFirmas
        ElseIf pnTpoOpe = 2 Then
            nCanFirmasAut = lrDatos!nNumCantFirmas
        End If
    End If
    Dim nIndex As Integer
    Set CargarCargoNivExoAuto = lrDRHCargo.RecuperaPersCargo(lsCadCargo)
    Set lrCargo = Nothing
End Function

Private Sub cmdAgrObs_Click()
    Dim i As Integer
    If Len(txtObs.Text) = 0 Then Exit Sub
    If feObservaciones.Rows - 1 > 0 Then
        For i = 1 To feObservaciones.Rows - 1
            If Trim(feObservaciones.TextMatrix(i, 1)) = Trim(txtObs.Text) Then
                MsgBox "Esta Observación ya fue registrada.", vbOKOnly + vbInformation, "Atención"
                Exit Sub
            End If
        Next i
    End If
    
    With feObservaciones
        .AdicionaFila
        .TextMatrix(feObservaciones.row, 1) = Trim(txtObs.Text)
        .TextMatrix(feObservaciones.row, 2) = Trim(ActXCodCta.NroCuenta)
        If nTpoOpe = 1 Then
            .TextMatrix(feObservaciones.row, 4) = "2"
        End If
    End With
    
    txtObs.Text = ""
    txtObs.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    If nTpoOpe <> 1 Then
        Call LimpiaFormulario
    Else
        Unload Me
    End If
    cmdChkList.Enabled = False
    frameConsulta.Enabled = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    tabRegistro.Enabled = False
    cmdChkList.Enabled = False
    'JOEP20190131 CP
    lsCategoria = ""
lsTipoCredito = ""
'JOEP20190131 CP
End Sub

Private Sub cmdChkList_Click()

    'frmAdmCheckListDocument.Inicio ActXCodCta.NroCuenta, sTpoCredCod, IIf(nTpoOpe = 1, nRegCntrlPre, nRegCntrlPost)'Comento JOEP20190131 CP
    frmAdmCheckListDocument.Inicio ActXCodCta.NroCuenta, lsCategoria, lsTipoProdCod, lblMonto, lsTipoCredito, IIf(nTpoOpe = 1, nRegCntrlPre, nRegCntrlPost)   'JOEP20190131 CP
    
End Sub
Private Sub cmdConsultPers_Click()
    frmPersona.Inicio feListaRelac.TextMatrix(feListaRelac.row, 3), PersonaConsulta
End Sub
Private Sub cmdDupDoc_Click()
    frmCredDupDoc.Iniciaformulario ActXCodCta.NroCuenta
End Sub

Private Sub cmdFinalidad_Click()
Call frmCFHistorial.CargaCFHistorial(ActXCodCta.NroCuenta)
End Sub

Private Sub cmdGarant_Click()
    'EJVG20150725 ***
    'frmPersGarantias.Inicio ConsultaGarant, feGarantias.TextMatrix(feGarantias.row, 8)
    If feGarantias.TextMatrix(1, 0) <> "" Then
        frmGarantia.Consultar Trim(feGarantias.TextMatrix(feGarantias.row, 8))
    Else
        MsgBox "No existen Garantías seleccionadas", vbInformation, "Aviso"
    End If
    'END EJVG *******
End Sub

Private Sub cmdGrabar_Click()
    Dim oCred As COMDCredito.DCOMCreditos
    Dim lsMovNro As String, lsCtaCod As String
    
'    If chkTpoRevision.value = 1 Then
'        If ValidaObsPendientes = False Then
'            MsgBox "Existe una observación pendiente.", vbInformation, "Alerta"
'            Exit Sub
'        End If
'    End If
    If MsgBox("¿Desea Registrar el Control de este Crédito?.", vbInformation + vbYesNo, "Atención") = vbYes Then
        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oCred = New COMDCredito.DCOMCreditos

        Call oCred.ControlAdmCred(ActXCodCta.NroCuenta, Format(gdFecSis, "yyyymmdd"), lsMovNro, LlenarRsObserv(feObservaciones), LlenarRsExonera(FEExoneraciones), LlenarRsAuto(FEAutorizaciones), IIf(nTpoOpe = 2, True, False), Me.chkTpoRevision.value, Right(cboTipoRev.Text, 1))
        Set oCred = Nothing
       
        MsgBox "Los datos se guardaron satisfactoriamente.", vbOKOnly + vbInformation, "Atención"
        lsCtaCod = ActXCodCta.NroCuenta
        Call LimpiaFormulario
        ActXCodCta.NroCuenta = lsCtaCod
        Call ActXCodCta_KeyPress(vbKeyReturn)
        'Unload Me
    End If
End Sub

Private Sub cmdInfRiesg_Click()
    frmCredRiesgosInformeMod.Inicio ActXCodCta.NroCuenta
End Sub

Private Sub cmdLevantarObs_Click()
    feObservaciones.TextMatrix(feObservaciones.row, 3) = 1
End Sub

Private Sub cmdPoliza_Click()
    'EJVG20150725 ***
    If feGarantias.TextMatrix(1, 0) <> "" Then
        Call frmCredPolizaConsulta.Inicio(feGarantias.TextMatrix(feGarantias.row, 11))
    Else
        MsgBox "No existen Garantías seleccionadas", vbInformation, "Aviso"
    End If
    'END EJVG *******
End Sub
Private Sub cmdPosClient_Click()
    frmPosicionCli.iniciarFormulario feListaRelac.TextMatrix(feListaRelac.row, 3)
End Sub

Private Sub cmdQuiObs_Click()
    If MsgBox("¿Esta seguro que desea quitar la observación?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
    Call feObservaciones.EliminaFila(feObservaciones.row)
End Sub

Private Sub cmdQuitaExo_Click()
    FEExoneraciones.EliminaFila FEExoneraciones.row
End Sub

Private Sub cmdQuitarAutorizacion_Click()
    FEAutorizaciones.EliminaFila FEAutorizaciones.row
End Sub

Private Sub cmdsalir_Click()
    sCadSalida = GeneraCadSalida
    Call LimpiaFormulario
    Unload Me
'JOEP20190131 CP
    lsCategoria = ""
lsTipoCredito = ""
'JOEP20190131 CP
End Sub

Private Sub cmdSelecApoAut_Click()
    Dim frm As New frmAdmCredListApoderados
    ReDim vDatos(2, 0)
    vDatos = frm.Inicio(Me.cboNivAutorizacion.Text, nCanFirmasAut, CargarCargoNivExoAuto(Me.cboNivAutorizacion, Me.cboAutorizacion, 2))
    Set frm = Nothing
    cmdAgregarAutorizacion.Enabled = IIf(UBound(vDatos, 2) > 0, True, False)
End Sub

Private Sub cmdSelectApoExo_Click()
    Dim frm As New frmAdmCredListApoderados
    ReDim vDatos(2, 0)
    vDatos = frm.Inicio(Me.cboNivelExonera.Text, nCanFirmasExo, CargarCargoNivExoAuto(Me.cboNivelExonera, Me.cboExoneraciones, 1))
    Set frm = Nothing
    cmdAgregarExo.Enabled = IIf(UBound(vDatos, 2) > 0, True, False)
End Sub

Private Sub Form_Load()

    Dim obj As New COMDConstantes.DCOMConstantes
      'JOEP20190131 CP
    lsCategoria = ""
lsTipoCredito = ""
'JOEP20190131 CP

    Call CargarExoneraciones
    Call CargarAutorizaciones
    Call CargarNivelExoneracion
    Call CargarNivelAutorizacion
    Call CargaCombo(cboTipoRev, obj.RecuperaConstantes(10065))
    Call feObservaciones.CargaCombo(obj.RecuperaConstantes(4046))
    cboTipoRev.ListIndex = 0
    Call LimpiaFormulario
  
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim rsCred As New ADODB.Recordset
    Dim rsComun As New ADODB.Recordset
    Dim prsRelac As New ADODB.Recordset
    Dim bCargado As Boolean, bRegistrado As Boolean, bPostDesem As Boolean
    Dim sEstado As String, sCampania As String
    Dim nPrestamo As Double
    Dim nTpoRevision As Integer
    
    
    bRegistrado = oCred.BuscaRegistroControlAdmCred(psCtaCod, nTpoRevision)
    bCargado = oCred.CargaDatosControlCreditosAdmCred(psCtaCod, rsCred, rsComun, sEstado, , sCampania)
    bPostDesem = oCred.ObtieneCredCFPostDesembolso(psCtaCod)
    
    nPrestamo = IIf(IsNull(rsComun!nMontoCol), rsComun!nMontoSol, rsComun!nMontoCol)
    If nPrestamo = 0 Then
        nPrestamo = rsComun!nMontoSol
    End If
    lblMoneda.Caption = IIf(IsNull(rsComun!cmoneda), "", rsComun!cmoneda)
    lblMonto.Caption = Format(nPrestamo, "#,###,###.00")
    lblFechaPago.Caption = IIf(IsNull(rsComun!dFechaPago), "", rsComun!dFechaPago)
    lblTem.Caption = rsComun!nTasaInteres
    'Modificado PASI20150610
    'lblPlazo.Caption = IIf(IsNull(rsComun!cPlazoCE), "", rsComun!nPlazo)
    lblNCuotas.Caption = rsComun!nCuotas
    'end PASI20150610
    lblCampania.Caption = sCampania
    sTpoCredCod = rsComun!cTpoCredCod
    
    'JOEP20190131 CP
    lsCategoria = rsComun!cCategoria
    lsTipoCredito = rsComun!cTipoCredito
    lsTipoProdCod = rsComun!cTpoProdCod
    Call bValiExistCheckList
    'JOEP20190131 CP
    
    If nTpoRevision > 0 Then
        chkTpoRevision.value = 1
        cboTipoRev.ListIndex = nTpoRevision - 1
        chkTpoRevision.Enabled = False
        cboTipoRev.Enabled = False
        If nTpoOpe = 1 Then
            cmdGrabar.Enabled = False
        End If
    End If
    'cmdChkList.Enabled = True
    frameConsulta.Enabled = True
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    tabRegistro.Enabled = True
    'cmdChkList.Enabled = True
    Call CargarFlexGarantia(psCtaCod)
    Call CargaPersonasRelacCred(psCtaCod, prsRelac)
    lblCampania.Visible = True
    Set oCred = Nothing
    Call CargarDatosObs(psCtaCod)
    Call CargarDatosExoneraciones
    Call CargarDatosAutorizaciones
    
End Function

Private Sub CargarFlexGarantia(ByVal psCtaCod As String)
    Dim RGar As New ADODB.Recordset
    Dim oDGarantia As New COMDCredito.DCOMGarantia
    Dim rsGarantReal As New ADODB.Recordset
    Dim oNCredito As New COMNCredito.NCOMCredito
    
    Set RGar = oNCredito.obtenerGarantxCredito(psCtaCod)
    Do While Not RGar.EOF
        feGarantias.AdicionaFila
        feGarantias.RowHeight(RGar.Bookmark) = 280
        feGarantias.TextMatrix(RGar.Bookmark, 1) = RGar!cTpoGarDescripcion
        feGarantias.TextMatrix(RGar.Bookmark, 2) = Format(RGar!nGravado, "#,#0.00")
        feGarantias.TextMatrix(RGar.Bookmark, 3) = Format(RGar!nTasacion, "#,#0.00")
        feGarantias.TextMatrix(RGar.Bookmark, 4) = Format(RGar!nRealizacion, "#,#0.00")
        feGarantias.TextMatrix(RGar.Bookmark, 5) = Format(RGar!nPorGravar, "#,#0.00")
        feGarantias.TextMatrix(RGar.Bookmark, 6) = Trim(RGar!cPersNombre)
        feGarantias.TextMatrix(RGar.Bookmark, 7) = Trim(RGar!cNroDoc)
        feGarantias.TextMatrix(RGar.Bookmark, 8) = Trim(RGar!cNumGarant)
        
        Set rsGarantReal = oDGarantia.RecuperaGarantiaRealMaxAprobacion(psCtaCod, RGar!cNumGarant)
        Set oDGarantia = Nothing
        feGarantias.TextMatrix(RGar.Bookmark, 11) = Trim(RGar!cNumPoliza)
        If Not (rsGarantReal.EOF Or rsGarantReal.BOF) Then
            If (rsGarantReal!cNumGarant <> "") And ((DateDiff("d", rsGarantReal!dCertifGravamen, gdFecSis) > 365) Or (DateDiff("d", rsGarantReal!dTasacion, gdFecSis) > 730)) Then
                Select Case rsGarantReal!nVerificaLegal
                    Case 1: feGarantias.TextMatrix(RGar.Bookmark, 9) = "Pendiente"
                    Case 2: feGarantias.TextMatrix(RGar.Bookmark, 9) = "Aprobado"
                    Case 3: feGarantias.TextMatrix(RGar.Bookmark, 9) = "Desaprobado"
                    Case 4: feGarantias.TextMatrix(RGar.Bookmark, 9) = "Pendiente por Regularizar"
                    Case 0: feGarantias.TextMatrix(RGar.Bookmark, 9) = "Pendiente"
                End Select
                feGarantias.TextMatrix(RGar.Bookmark, 10) = Trim(IIf(Len(Trim(rsGarantReal!cNumPoliza)) > 1, IIf(Len(Trim(rsGarantReal!nEstadoPolizaNew)) > 1, rsGarantReal!nEstadoPolizaNew, "No registrado"), "No tiene"))
            Else
                feGarantias.TextMatrix(RGar.Bookmark, 9) = "Conforme"
                feGarantias.TextMatrix(RGar.Bookmark, 10) = Trim(IIf(Len(Trim(rsGarantReal!cNumPoliza)) > 1, IIf(Len(Trim(rsGarantReal!nEstadoPolizaNew)) > 1, rsGarantReal!nEstadoPolizaNew, "No registrado"), "No tiene"))
            End If
        Else
            feGarantias.TextMatrix(RGar.Bookmark, 9) = "No aplica"
            feGarantias.TextMatrix(RGar.Bookmark, 10) = "No tiene"
        End If
        RGar.MoveNext
    Loop
End Sub

Private Sub CargaPersonasRelacCred(ByVal psCtaCod As String, ByVal prsRelac As ADODB.Recordset)
Dim s As ListItem
Dim oRelPersCred As UCredRelac_Cli
    Set oRelPersCred = New UCredRelac_Cli
    
    On Error GoTo ErrorCargaPersonasRelacCred
    Call oRelPersCred.CargaRelacPersCred(psCtaCod)
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        feListaRelac.AdicionaFila
        feListaRelac.TextMatrix(feListaRelac.Rows - 1, 1) = oRelPersCred.ObtenerNombre
        feListaRelac.TextMatrix(feListaRelac.Rows - 1, 2) = oRelPersCred.ObtenerRelac
        feListaRelac.TextMatrix(feListaRelac.Rows - 1, 3) = oRelPersCred.ObtenerCodigo
        oRelPersCred.siguiente
    Loop
    Exit Sub
ErrorCargaPersonasRelacCred:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub CargarDatosObs(ByVal psCtaCod As String)
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim rsCred As New ADODB.Recordset
    
    If nTpoOpe = 1 Then
        Set rsCred = oCred.BuscaObsAdmCred(psCtaCod)
    Else
        Set rsCred = oCred.DevuelveObsPostDesembolso(psCtaCod)
    End If
    Set oCred = Nothing
    If Not (rsCred.EOF And rsCred.BOF) Then
        feObservaciones.Clear
        FormateaFlex feObservaciones
        Do While Not rsCred.EOF
            feObservaciones.AdicionaFila
            feObservaciones.TextMatrix(feObservaciones.row, 1) = Trim(rsCred!CDescripcion)
            feObservaciones.TextMatrix(feObservaciones.row, 2) = Trim(rsCred!cCtaCod)
            feObservaciones.TextMatrix(feObservaciones.row, 3) = Trim(rsCred!nRegulariza)
            If nTpoOpe <> 1 Then
                feObservaciones.TextMatrix(feObservaciones.row, 4) = rsCred!cAnalistaObs
            End If
            rsCred.MoveNext
        Loop
        rsCred.Close
        Set rsCred = Nothing
    End If
End Sub

Private Sub CargarDatosExoneraciones()
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim RS As New ADODB.Recordset
    Dim x As Integer
    
    Set RS = oCred.DevulveExoneracionesCredControlAdm(ActXCodCta.NroCuenta)
    FEExoneraciones.Clear
    FormateaFlex FEExoneraciones
    
    If Not (RS.EOF And RS.BOF) Then
        For x = 1 To RS.RecordCount
            FEExoneraciones.AdicionaFila
            FEExoneraciones.TextMatrix(x, 1) = RS!cExoneraCod
            FEExoneraciones.TextMatrix(x, 2) = RS!cNivAprCod
            FEExoneraciones.TextMatrix(x, 3) = RS!Apoderado1
            FEExoneraciones.TextMatrix(x, 4) = RS!Apoderado2
            FEExoneraciones.TextMatrix(x, 5) = RS!Apoderado3
            RS.MoveNext
        Next
    End If
End Sub
Public Sub CargarDatosAutorizaciones()
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim RS As New ADODB.Recordset
    Dim x As Integer
    
    Set RS = oCred.DevulveAutorizacionesCredControlAdm(ActXCodCta.NroCuenta)
    FEAutorizaciones.Clear
    FormateaFlex FEAutorizaciones
    
    If Not (RS.EOF And RS.BOF) Then
        For x = 1 To RS.RecordCount
            FEAutorizaciones.AdicionaFila
            FEAutorizaciones.TextMatrix(x, 1) = RS!cAutorizaCod
            FEAutorizaciones.TextMatrix(x, 2) = RS!cNivAprCod
            FEAutorizaciones.TextMatrix(x, 3) = RS!Apoderado1
            FEAutorizaciones.TextMatrix(x, 4) = RS!Apoderado2
            FEAutorizaciones.TextMatrix(x, 5) = RS!Apoderado3
            RS.MoveNext
        Next
    End If
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdAgrObs_Click
    End If
End Sub

Private Function LlenarRsExonera(ByVal feControl As FlexEdit) As ADODB.Recordset
    Dim rsExoneraciones As New ADODB.Recordset
    Dim nIndex As Integer
    Set rsExoneraciones = New ADODB.Recordset
    
    If feControl.Rows >= 2 Then
        If feControl.TextMatrix(nIndex, 1) = "" Then
                Exit Function
        End If
        rsExoneraciones.CursorType = adOpenStatic
        rsExoneraciones.Fields.Append "cCtaCod", adVarChar, 18, adFldIsNullable
        rsExoneraciones.Fields.Append "nCodExonera", adVarChar, 7, adFldIsNullable
        rsExoneraciones.Fields.Append "cCodQuienExo", adVarChar, 7, adFldIsNullable
        rsExoneraciones.Fields.Append "Apoder1", adVarChar, 4, adFldIsNullable
        rsExoneraciones.Fields.Append "Apoder2", adVarChar, 4, adFldIsNullable
        rsExoneraciones.Fields.Append "Apoder3", adVarChar, 4, adFldIsNullable
        rsExoneraciones.Fields.Append "nDesExoneraCAR", adInteger, 4, adFldIsNullable
        rsExoneraciones.Fields.Append "cDesExoneraOtros", adVarChar, 100, adFldIsNullable
        rsExoneraciones.Open
        
        For nIndex = 1 To feControl.Rows - 1
            rsExoneraciones.AddNew
            rsExoneraciones.Fields("cCtaCod") = ActXCodCta.NroCuenta
            rsExoneraciones.Fields("nCodExonera") = Right(feControl.TextMatrix(nIndex, 1), 7)
            rsExoneraciones.Fields("cCodQuienExo") = Right(feControl.TextMatrix(nIndex, 2), 7)
            rsExoneraciones.Fields("Apoder1") = Right(feControl.TextMatrix(nIndex, 3), 4)
            rsExoneraciones.Fields("Apoder2") = Right(feControl.TextMatrix(nIndex, 4), 4)
            rsExoneraciones.Fields("Apoder3") = Right(feControl.TextMatrix(nIndex, 5), 4)
            rsExoneraciones.Fields("nDesExoneraCAR") = 0
            rsExoneraciones.Fields("cDesExoneraOtros") = ""
            rsExoneraciones.Update
            rsExoneraciones.MoveFirst
        Next
    End If
    Set LlenarRsExonera = rsExoneraciones
End Function
Private Function LlenarRsAuto(ByVal feControl As FlexEdit) As ADODB.Recordset
    Dim rsAutorizaciones As New ADODB.Recordset
    Dim nIndex As Integer
    Set rsAutorizaciones = New ADODB.Recordset
    
    If feControl.Rows >= 2 Then
        If feControl.TextMatrix(nIndex, 1) = "" Then
                Exit Function
        End If
        rsAutorizaciones.CursorType = adOpenStatic
        rsAutorizaciones.Fields.Append "cCtaCod", adVarChar, 18, adFldIsNullable
        rsAutorizaciones.Fields.Append "nCodAutoriza", adVarChar, 7, adFldIsNullable
        rsAutorizaciones.Fields.Append "cCodQuienAuto", adVarChar, 7, adFldIsNullable
        rsAutorizaciones.Fields.Append "Apoder1", adVarChar, 4, adFldIsNullable
        rsAutorizaciones.Fields.Append "Apoder2", adVarChar, 4, adFldIsNullable
        rsAutorizaciones.Fields.Append "Apoder3", adVarChar, 4, adFldIsNullable
        rsAutorizaciones.Fields.Append "AutoOtro", adVarChar, 100, adFldIsNullable
        rsAutorizaciones.Open
        For nIndex = 1 To feControl.Rows - 1
            rsAutorizaciones.AddNew
            rsAutorizaciones.Fields("cCtaCod") = ActXCodCta.NroCuenta
            rsAutorizaciones.Fields("nCodAutoriza") = Right(feControl.TextMatrix(nIndex, 1), 7)
            rsAutorizaciones.Fields("cCodQuienAuto") = Right(feControl.TextMatrix(nIndex, 2), 7)
            rsAutorizaciones.Fields("Apoder1") = Right(feControl.TextMatrix(nIndex, 3), 4)
            rsAutorizaciones.Fields("Apoder2") = IIf(Right(feControl.TextMatrix(nIndex, 4), 4) = "", Space(4), Right(feControl.TextMatrix(nIndex, 4), 4))
            rsAutorizaciones.Fields("Apoder3") = IIf(Right(feControl.TextMatrix(nIndex, 5), 4) = "", Space(4), Right(feControl.TextMatrix(nIndex, 5), 4))
            rsAutorizaciones.Fields("AutoOtro") = ""
            rsAutorizaciones.Update
            rsAutorizaciones.MoveFirst
        Next
    End If
    Set LlenarRsAuto = rsAutorizaciones
End Function
Private Function LlenarRsObserv(ByVal feControl As FlexEdit) As ADODB.Recordset
    Dim rsObs As New ADODB.Recordset
    Dim nIndex As Integer
    
    If feControl.Rows >= 2 Then
        If feControl.TextMatrix(nIndex, 1) = "" Then
            Exit Function
        End If

        rsObs.CursorType = adOpenStatic
        rsObs.Fields.Append "cCtaCod", adVarChar, 18, adFldIsNullable
        rsObs.Fields.Append "observacion", adVarChar, 100, adFldIsNullable
        rsObs.Fields.Append "nRegulariza", adInteger, 4, adFldIsNullable
        rsObs.Fields.Append "nAnalistaObs", adInteger, 4, adFldIsNullable
        rsObs.Open
        
        For nIndex = 1 To feControl.Rows - 1
            rsObs.AddNew
            rsObs.Fields("cCtaCod") = feControl.TextMatrix(nIndex, 2)
            rsObs.Fields("observacion") = feControl.TextMatrix(nIndex, 1)
            rsObs.Fields("nRegulariza") = IIf(feControl.TextMatrix(nIndex, 3) = "", 0, 1)
            rsObs.Fields("nAnalistaObs") = IIf(feControl.TextMatrix(nIndex, 4) = "", 2, Right(feControl.TextMatrix(nIndex, 4), 1))
            rsObs.Update
            rsObs.MoveFirst
        Next
    End If
    Set LlenarRsObserv = rsObs
End Function

Private Sub HabilitaControles(ByVal pbValor As Boolean)
    With tabRegistro
        '.TabVisible(1) = pbValor
        '.TabVisible(2) = pbValor
    End With
    lblCampania.Visible = pbValor
    chkTpoRevision.Enabled = False
    cboTipoRev.Enabled = False
End Sub

Public Function ValidaExisteElemento(ByVal poFlex As FlexEdit, ByVal psExoAut As String, ByVal pnTpo As Integer) As String
    Dim i As Integer
    If poFlex.Rows > 1 And poFlex.TextMatrix(1, 1) <> "" Then
        For i = 1 To poFlex.Rows - 1
            If RTrim(Left(poFlex.TextMatrix(i, 1), Len(poFlex.TextMatrix(i, 1)) - 7)) = RTrim(Left(psExoAut, Len(psExoAut) - 7)) Then
                If pnTpo = 1 Then
                    ValidaExisteElemento = "Ya existe esta exoneración para este crédito"
                Else
                    ValidaExisteElemento = "Ya existe esta autorización para este crédito"
                End If
                Exit Function
            End If
        Next
    End If
End Function

Private Function ValidaObsPendientes() As Boolean
    Dim i As Integer
    ValidaObsPendientes = True
    For i = 1 To feObservaciones.Rows - 1
        If feObservaciones.TextMatrix(i, 1) <> "" And feObservaciones.TextMatrix(i, 3) = "" Then
            ValidaObsPendientes = False
            Exit Function
        End If
    Next
End Function

Private Sub LimpiaFormulario()
    ActXCodCta.NroCuenta = ""
    lblMoneda.Caption = ""
    lblMonto.Caption = ""
    lblTem.Caption = ""
    lblNCuotas.Caption = ""
    lblFechaPago.Caption = ""
    txtObs.Text = ""
    feListaRelac.Clear
    feGarantias.Clear
    feObservaciones.Clear
    FEExoneraciones.Clear
    FEAutorizaciones.Clear
    FormateaFlex feListaRelac
    FormateaFlex feGarantias
    FormateaFlex feObservaciones
    FormateaFlex FEExoneraciones
    FormateaFlex FEAutorizaciones
    cboExoneraciones.ListIndex = 0
    cboAutorizacion.ListIndex = 0
    cboTipoRev.ListIndex = 0
    nCargoNivApr = 0
    nCanFirmasAut = 0
    nCanFirmasExo = 0
    nContador = 0
    sPersCargoExo = ""
    sPersCargoAuto = ""
    sTpoCredCod = ""
    cmdAgregarAutorizacion.Enabled = False
    cmdAgregarExo.Enabled = False
End Sub

Public Function GeneraCadSalida() As String
    Dim i As Integer
    For i = 1 To feObservaciones.Rows - 1
        If feObservaciones.TextMatrix(i, 3) = "." Then
            GeneraCadSalida = "1,"
        End If
    Next
    If chkTpoRevision.value = 0 Then
        GeneraCadSalida = GeneraCadSalida & "2,"
    End If
End Function

'JOEP20190201 CP
Private Sub bValiExistCheckList()
Dim rsActivar As ADODB.Recordset
Dim pnActivar As Integer
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito
Set rsActivar = oCredito.getActivarCheckList(lsTipoProdCod)
If Not (rsActivar.BOF And rsActivar.EOF) Then
    If rsActivar!Activar = 1 Then
        cmdChkList.Enabled = True
    Else
        cmdChkList.Enabled = False
    End If
End If
Set oCredito = Nothing
RSClose rsActivar
End Sub
'JOEP20190201 CP
