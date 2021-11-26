VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPagoServRecaudo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicios - Servicios de Recaudo"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8445
   Icon            =   "frmPagoServRecaudo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Selección del Convenio"
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
      Height          =   8220
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   375
         Left            =   5550
         TabIndex        =   40
         Top             =   7740
         Width           =   1335
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
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
         Left            =   6960
         TabIndex        =   41
         Top             =   7740
         Width           =   1335
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
         Height          =   375
         Left            =   150
         TabIndex        =   42
         Top             =   7740
         Width           =   1335
      End
      Begin TabDlg.SSTab stContenedorValidacion 
         Height          =   5880
         Left            =   105
         TabIndex        =   48
         Top             =   1800
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10372
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Validación Completa"
         TabPicture(0)   =   "frmPagoServRecaudo.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label7"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label8"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame3"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame4"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "grdConceptoPagarVC"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtSubTotalVC"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtComisionVC"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtTotalVC"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Validación Incompleta"
         TabPicture(1)   =   "frmPagoServRecaudo.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label13"
         Tab(1).Control(1)=   "Label14"
         Tab(1).Control(2)=   "Label15"
         Tab(1).Control(3)=   "grdConceptoPagarVI"
         Tab(1).Control(4)=   "Frame5"
         Tab(1).Control(5)=   "Frame6"
         Tab(1).Control(6)=   "txtSubTotalVI"
         Tab(1).Control(7)=   "txtComisionVI"
         Tab(1).Control(8)=   "txtTotalVI"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Sin Validación"
         TabPicture(2)   =   "frmPagoServRecaudo.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdQuitarSV"
         Tab(2).Control(1)=   "cmdAgregarSV"
         Tab(2).Control(2)=   "txtTotalSV"
         Tab(2).Control(3)=   "txtComisionSV"
         Tab(2).Control(4)=   "txtSubTotalSV"
         Tab(2).Control(5)=   "Frame8"
         Tab(2).Control(6)=   "grdConceptoPagarSV"
         Tab(2).Control(7)=   "Label21"
         Tab(2).Control(8)=   "Label20"
         Tab(2).Control(9)=   "Label19"
         Tab(2).ControlCount=   10
         TabCaption(3)   =   "Validación por Importe"
         TabPicture(3)   =   "frmPagoServRecaudo.frx":035E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label29"
         Tab(3).Control(1)=   "Label30"
         Tab(3).Control(2)=   "Label31"
         Tab(3).Control(3)=   "grdConceptoPagarVX"
         Tab(3).Control(4)=   "txtSubTotalVX"
         Tab(3).Control(5)=   "txtComisionVX"
         Tab(3).Control(6)=   "txtTotalVX"
         Tab(3).Control(7)=   "cmdAgregarVX"
         Tab(3).Control(8)=   "cmdQuitarVX"
         Tab(3).Control(9)=   "Frame7"
         Tab(3).ControlCount=   10
         Begin VB.Frame Frame7 
            Caption         =   "Operación"
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
            Height          =   1425
            Left            =   -74895
            TabIndex        =   78
            Top             =   360
            Width           =   7935
            Begin VB.TextBox txtDOIVX 
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
               Height          =   315
               Left            =   3810
               TabIndex        =   32
               Top             =   960
               Width           =   1380
            End
            Begin VB.TextBox txtNombreClienteVX 
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
               Height          =   315
               Left            =   1065
               TabIndex        =   30
               Top             =   600
               Width           =   6390
            End
            Begin VB.ComboBox cboTipoDOIVX 
               Height          =   315
               ItemData        =   "frmPagoServRecaudo.frx":037A
               Left            =   1065
               List            =   "frmPagoServRecaudo.frx":0384
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtOtroCodigoVX 
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
               Height          =   315
               Left            =   5970
               MaxLength       =   10
               TabIndex        =   33
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label Label33 
               Caption         =   "Nombre: "
               Height          =   255
               Left            =   150
               TabIndex        =   84
               Top             =   660
               Width           =   735
            End
            Begin VB.Label Label28 
               Caption         =   "Tipo DOI: "
               Height          =   255
               Left            =   150
               TabIndex        =   83
               Top             =   1020
               Width           =   945
            End
            Begin VB.Label Label27 
               Caption         =   "Nª DOI: "
               Height          =   255
               Left            =   2520
               TabIndex        =   82
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label26 
               Caption         =   "Código: "
               Height          =   255
               Left            =   5280
               TabIndex        =   81
               Top             =   1020
               Width           =   615
            End
            Begin VB.Label lblDescripcionCodigoVP 
               Alignment       =   2  'Center
               Caption         =   "Descripción de la Operación"
               Height          =   255
               Left            =   1065
               TabIndex        =   80
               Top             =   300
               Width           =   6390
            End
            Begin VB.Label Label18 
               Caption         =   "Nª DOI:"
               Height          =   255
               Left            =   3030
               TabIndex        =   79
               Top             =   990
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdQuitarVX 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   -73740
            TabIndex        =   36
            Top             =   4770
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarVX 
            Caption         =   "&Agregar"
            Height          =   375
            Left            =   -74910
            TabIndex        =   35
            Top             =   4770
            Width           =   1095
         End
         Begin VB.TextBox txtTotalVX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Left            =   -68550
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   5400
            Width           =   1575
         End
         Begin VB.TextBox txtComisionVX 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   -68550
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "0.00"
            Top             =   5085
            Width           =   1575
         End
         Begin VB.TextBox txtSubTotalVX 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   -68550
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "0.00"
            Top             =   4770
            Width           =   1575
         End
         Begin VB.CommandButton cmdQuitarSV 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   -73680
            TabIndex        =   26
            Top             =   4755
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarSV 
            Caption         =   "&Agregar"
            Height          =   375
            Left            =   -74880
            TabIndex        =   25
            Top             =   4755
            Width           =   1095
         End
         Begin VB.TextBox txtTotalSV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "0.00"
            Top             =   5400
            Width           =   1575
         End
         Begin VB.TextBox txtComisionSV 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   5085
            Width           =   1575
         End
         Begin VB.TextBox txtSubTotalSV 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "0.00"
            Top             =   4770
            Width           =   1575
         End
         Begin VB.Frame Frame8 
            Caption         =   "Operación"
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
            Height          =   1425
            Left            =   -74880
            TabIndex        =   65
            Top             =   360
            Width           =   7935
            Begin VB.TextBox txtDOISV 
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
               Height          =   315
               Left            =   3690
               TabIndex        =   22
               ToolTipText     =   "Ingresar DOI"
               Top             =   960
               Width           =   1380
            End
            Begin VB.TextBox txtOtroCodigoSV 
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
               Height          =   315
               Left            =   5970
               MaxLength       =   10
               TabIndex        =   23
               ToolTipText     =   "Ingresar Codigo"
               Top             =   960
               Width           =   1485
            End
            Begin VB.ComboBox cboTipoDOISV 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "frmPagoServRecaudo.frx":05E5
               Left            =   1065
               List            =   "frmPagoServRecaudo.frx":05EF
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtNombreClienteSV 
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
               Height          =   315
               Left            =   1065
               TabIndex        =   20
               Top             =   600
               Width           =   6390
            End
            Begin VB.Label Label32 
               Caption         =   "Nª DOI:"
               Height          =   255
               Left            =   3030
               TabIndex        =   77
               Top             =   1020
               Width           =   735
            End
            Begin VB.Label lblDescripcionCodigoSV 
               Alignment       =   2  'Center
               Caption         =   "Descripción de la Operación"
               Height          =   255
               Left            =   1050
               TabIndex        =   76
               Top             =   300
               Width           =   6405
            End
            Begin VB.Label Label24 
               Caption         =   "Código: "
               Height          =   255
               Left            =   5190
               TabIndex        =   72
               Top             =   1020
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "Nª DOI: "
               Height          =   255
               Left            =   2520
               TabIndex        =   71
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label22 
               Caption         =   "Tipo DOI: "
               Height          =   255
               Left            =   150
               TabIndex        =   70
               Top             =   1020
               Width           =   750
            End
            Begin VB.Label Label17 
               Caption         =   "Nombre: "
               Height          =   255
               Left            =   180
               TabIndex        =   66
               Top             =   630
               Width           =   735
            End
         End
         Begin VB.TextBox txtTotalVI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   5415
            Width           =   1575
         End
         Begin VB.TextBox txtComisionVI 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   5100
            Width           =   1575
         End
         Begin VB.TextBox txtSubTotalVI 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   4785
            Width           =   1575
         End
         Begin VB.Frame Frame6 
            Caption         =   "Datos del Cliente"
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
            Left            =   -74880
            TabIndex        =   58
            Top             =   1200
            Width           =   7935
            Begin VB.TextBox txtNombreClienteVI 
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
               Height          =   315
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   270
               Width           =   4575
            End
            Begin VB.TextBox txtDOIVI 
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
               Height          =   315
               Left            =   6090
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   270
               Width           =   1695
            End
            Begin VB.Label Label12 
               Caption         =   "Nombre: "
               Height          =   255
               Left            =   150
               TabIndex        =   60
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label11 
               Caption         =   "DOI: "
               Height          =   255
               Left            =   5640
               TabIndex        =   59
               Top             =   330
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Busqueda"
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
            Left            =   -74880
            TabIndex        =   57
            Top             =   360
            Width           =   7935
            Begin VB.TextBox txtCodigoIDVI 
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
               Height          =   315
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   13
               Top             =   240
               Width           =   1575
            End
            Begin VB.CommandButton cmdBuscarPersonaVI 
               Appearance      =   0  'Flat
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
               Left            =   2535
               TabIndex        =   14
               Top             =   240
               Width           =   400
            End
            Begin VB.Label Label34 
               Caption         =   "Código: "
               Height          =   255
               Left            =   150
               TabIndex        =   87
               Top             =   285
               Width           =   735
            End
            Begin VB.Label lblDescripcionCodigoVI 
               Caption         =   "DESCRIPCION PARA EL USUARIO"
               Height          =   285
               Left            =   3210
               TabIndex        =   86
               Top             =   300
               Width           =   4575
            End
         End
         Begin VB.TextBox txtTotalVC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   5415
            Width           =   1575
         End
         Begin VB.TextBox txtComisionVC 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   5100
            Width           =   1575
         End
         Begin VB.TextBox txtSubTotalVC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
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
            Height          =   315
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   4785
            Width           =   1575
         End
         Begin SICMACT.FlexEdit grdConceptoPagarVC 
            Height          =   2625
            Left            =   120
            TabIndex        =   9
            Top             =   2040
            Width           =   7935
            _extentx        =   13996
            _extenty        =   2355
            cols0           =   10
            fixedcols       =   0
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "ID-Servicio-Concepto-Fec Vec-Moneda-Importe-Mora-Pagar-IDCopy-a"
            encabezadosanchos=   "500-1900-1900-0-1200-1300-0-1000-0-0"
            font            =   "frmPagoServRecaudo.frx":074E
            font            =   "frmPagoServRecaudo.frx":077A
            font            =   "frmPagoServRecaudo.frx":07A6
            font            =   "frmPagoServRecaudo.frx":07D2
            fontfixed       =   "frmPagoServRecaudo.frx":07FE
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-X-X-X-X-7-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-4-0-0"
            encabezadosalineacion=   "C-L-L-C-L-R-R-C-C-C"
            formatosedit    =   "0-0-0-0-2-0-0-0-0-0"
            textarray0      =   "ID"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Frame Frame4 
            Caption         =   "Datos del Cliente"
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
            TabIndex        =   50
            Top             =   1200
            Width           =   7935
            Begin VB.TextBox txtDOIVC 
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
               Height          =   315
               Left            =   6120
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   300
               Width           =   1695
            End
            Begin VB.TextBox txtNombreClienteVC 
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
               Height          =   315
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   300
               Width           =   4575
            End
            Begin VB.Label Label5 
               Caption         =   "DOI: "
               Height          =   255
               Left            =   5640
               TabIndex        =   52
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label4 
               Caption         =   "Nombre: "
               Height          =   255
               Left            =   180
               TabIndex        =   51
               Top             =   330
               Width           =   735
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Busqueda"
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
            TabIndex        =   49
            Top             =   360
            Width           =   7935
            Begin VB.CommandButton cmdBuscarPersonaVC 
               Appearance      =   0  'Flat
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
               Left            =   2535
               TabIndex        =   6
               Top             =   270
               Width           =   400
            End
            Begin VB.TextBox txtCodigoIDVC 
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
               Height          =   315
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   5
               Top             =   270
               Width           =   1575
            End
            Begin VB.Label Label9 
               Caption         =   "Código: "
               Height          =   255
               Left            =   150
               TabIndex        =   85
               Top             =   300
               Width           =   630
            End
            Begin VB.Label lblDescripcionCodigoVC 
               Caption         =   "DESCRIPCION PARA EL USUARIO"
               Height          =   225
               Left            =   3240
               TabIndex        =   56
               Top             =   330
               Width           =   4590
            End
         End
         Begin SICMACT.FlexEdit grdConceptoPagarVI 
            Height          =   2625
            Left            =   -74880
            TabIndex        =   61
            Top             =   2040
            Width           =   7935
            _extentx        =   13996
            _extenty        =   2963
            cols0           =   7
            fixedcols       =   0
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "ID-Servicio-Concepto-Moneda-Pagar-Importe-a"
            encabezadosanchos=   "500-1900-1900-1200-1000-1300-0"
            font            =   "frmPagoServRecaudo.frx":082C
            font            =   "frmPagoServRecaudo.frx":0858
            font            =   "frmPagoServRecaudo.frx":0884
            font            =   "frmPagoServRecaudo.frx":08B0
            fontfixed       =   "frmPagoServRecaudo.frx":08DC
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-X-4-5-X"
            listacontroles  =   "0-0-0-0-4-0-0"
            encabezadosalineacion=   "C-C-C-L-C-R-C"
            formatosedit    =   "0-0-0-0-0-2-2"
            textarray0      =   "ID"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit grdConceptoPagarSV 
            Height          =   2760
            Left            =   -74880
            TabIndex        =   24
            Top             =   1890
            Width           =   7935
            _extentx        =   13996
            _extenty        =   2355
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Concepto-Moneda-Importe-id"
            encabezadosanchos=   "500-5000-1200-1100-0"
            font            =   "frmPagoServRecaudo.frx":090A
            font            =   "frmPagoServRecaudo.frx":0936
            font            =   "frmPagoServRecaudo.frx":0962
            font            =   "frmPagoServRecaudo.frx":098E
            fontfixed       =   "frmPagoServRecaudo.frx":09BA
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-1-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-L-C-R-C"
            formatosedit    =   "0-0-0-2-2"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit grdConceptoPagarVX 
            Height          =   2700
            Left            =   -74910
            TabIndex        =   34
            Top             =   1950
            Width           =   7935
            _extentx        =   13996
            _extenty        =   2355
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Concepto-Moneda-Importe-id"
            encabezadosanchos=   "500-5000-1200-1100-0"
            font            =   "frmPagoServRecaudo.frx":09E8
            font            =   "frmPagoServRecaudo.frx":0A14
            font            =   "frmPagoServRecaudo.frx":0A40
            font            =   "frmPagoServRecaudo.frx":0A6C
            fontfixed       =   "frmPagoServRecaudo.frx":0A98
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-1-X-X-X"
            listacontroles  =   "0-3-0-0-0"
            encabezadosalineacion=   "C-L-C-R-C"
            formatosedit    =   "0-0-0-2-2"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Label Label31 
            Caption         =   "Total: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   75
            Top             =   5460
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "Comisión: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   74
            Top             =   5160
            Width           =   765
         End
         Begin VB.Label Label29 
            Caption         =   "Subtotal: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   73
            Top             =   4830
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Total: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   69
            Top             =   5400
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "Comisión: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   68
            Top             =   5130
            Width           =   840
         End
         Begin VB.Label Label19 
            Caption         =   "Subtotal: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   67
            Top             =   4830
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Total: "
            Height          =   225
            Left            =   -69330
            TabIndex        =   64
            Top             =   5460
            Width           =   555
         End
         Begin VB.Label Label14 
            Caption         =   "Comisión: "
            Height          =   255
            Left            =   -69330
            TabIndex        =   63
            Top             =   5160
            Width           =   765
         End
         Begin VB.Label Label13 
            Caption         =   "Subtotal: "
            Height          =   255
            Left            =   -69330
            TabIndex        =   62
            Top             =   4830
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Total: "
            Height          =   255
            Left            =   5640
            TabIndex        =   55
            Top             =   5460
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Comisión: "
            Height          =   255
            Left            =   5640
            TabIndex        =   54
            Top             =   5160
            Width           =   765
         End
         Begin VB.Label Label6 
            Caption         =   "Subtotal: "
            Height          =   255
            Left            =   5640
            TabIndex        =   53
            Top             =   4830
            Width           =   705
         End
      End
      Begin VB.Frame pnlBusquedaConvenio 
         Appearance      =   0  'Flat
         Caption         =   "Búsqueda de convenio"
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
         Height          =   1395
         Left            =   120
         TabIndex        =   44
         Top             =   300
         Width           =   8175
         Begin VB.CommandButton cmdBuscarConvenio 
            Appearance      =   0  'Flat
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
            Height          =   314
            Left            =   3255
            TabIndex        =   1
            Top             =   240
            Width           =   400
         End
         Begin VB.TextBox txtCodigoBusConvenio 
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
            Height          =   315
            Left            =   1080
            MaxLength       =   18
            TabIndex        =   0
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtNombreConvenio 
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
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   565
            Width           =   5505
         End
         Begin VB.TextBox txtCodigoEmpresa 
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
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   900
            Width           =   1455
         End
         Begin VB.TextBox txtNombreEmpresa 
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
            Height          =   315
            Left            =   2550
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   900
            Width           =   4035
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            Caption         =   "Código: "
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "Convenio: "
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   46
            Top             =   630
            Width           =   840
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "Empresa: "
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   45
            Top             =   960
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "frmPagoServRecaudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************************************************************************
'* NOMBRE         : "frmPagoServRecaudo"
'* DESCRIPCION    : Formulario creado para el pago de servicios de convenios segun proyecto: "Mejora del Sistema y Automatizacion de Ahorros y Servicios"
'* CREACION       : RIRO, 20121213 10:00 AM
'*********************************************************************************************************************************************************

Option Explicit

Private nConvenioSeleccionado As Integer
Private rsUsuarioRecaudo As ADODB.Recordset
Private nmoneda As Integer ' 1= Soles, 2= Extranjera
Private nComisionEmpresa As Double
Private nComisionCliente As Double
Private strCuenta As String
Private clsprevio As New previo.clsprevio 'RIRO20140906
Dim Importes() As String
Dim bFocusGrid As Boolean

'Para validar el DOI
Private Sub cboTipoDOISV_Click()
    If cboTipoDOISV.ListIndex = 0 Then
        txtDOISV.MaxLength = 8
        txtDOISV.Text = Empty
        txtDOISV.SetFocus
    ElseIf cboTipoDOISV.ListIndex = 1 Then
        txtDOISV.MaxLength = 11
        txtDOISV.Text = Empty
        If txtDOISV.Enabled Then txtDOISV.SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 And bFocusGrid Then
        KeyCode = 10
    End If
End Sub
Private Sub cboTipoDOISV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDOISV.SetFocus
    End If
End Sub

'Para validar el DOI
Private Sub cboTipoDOIVX_Click()

    If cboTipoDOIVX.ListIndex = 0 Then
    
        txtDOIVX.MaxLength = 8
        txtDOIVX.Text = Empty
        txtDOIVX.SetFocus
    
    ElseIf cboTipoDOIVX.ListIndex = 1 Then
    
        txtDOIVX.MaxLength = 11
        txtDOIVX.Text = Empty
        txtDOIVX.SetFocus
    
    End If

End Sub

Private Sub cboTipoDOIVX_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtDOIVX.SetFocus
    End If

End Sub

Private Sub cmdAgregarSV_Click()

    grdConceptoPagarSV.AdicionaFila
    grdConceptoPagarSV.SetFocus
    grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.Rows - 1, 2) = IIf(nmoneda = 1, "SOLES", "DOLARES")
    grdConceptoPagarSV.row = grdConceptoPagarSV.Rows - 1
    grdConceptoPagarSV.col = 1
    SendKeys "{F2}"
    
End Sub

Private Sub cmdAgregarSV_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdQuitarSV.SetFocus
    End If

End Sub

Private Sub cmdAgregarVX_Click()

    grdConceptoPagarVX.AdicionaFila
    grdConceptoPagarVX.SetFocus
    grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.Rows - 1, 2) = _
                                                                IIf(nmoneda = 1, "SOLES", "DOLARES")
    SendKeys "{ENTER}"
    
End Sub

Private Sub cmdBuscarPersonaVC_Click()
    
    Dim oBuscaUsuario As frmBuscarUsuarioConvenio
    Set oBuscaUsuario = New frmBuscarUsuarioConvenio
    Set rsUsuarioRecaudo = New Recordset
    Set rsUsuarioRecaudo = oBuscaUsuario.Inicio(Trim(txtCodigoBusConvenio.Text))
    
    If Not rsUsuarioRecaudo Is Nothing Then
        If Not (rsUsuarioRecaudo.EOF And rsUsuarioRecaudo.BOF) Then
            txtNombreClienteVC.Text = rsUsuarioRecaudo!cNomCliente
            txtDOIVC.Text = rsUsuarioRecaudo!cDOI
            txtCodigoIDVC.Text = rsUsuarioRecaudo!cCodCliente
        Else
            Exit Sub
        End If
    Dim nRow  As Integer
    LimpiaFlex grdConceptoPagarVC
    grdConceptoPagarVC_OnCellCheck 1, 5
    Do While Not rsUsuarioRecaudo.EOF
        
        'Se valida el estado del cliente enviado por la empresa
        If rsUsuarioRecaudo!nEstado = Registrado Or rsUsuarioRecaudo!nEstado = Pagando Then
            
            'ID-Servicio-Concepto-Fec Vec-Moneda-Importe-Mora-Pagar-a
            
            grdConceptoPagarVC.AdicionaFila
            nRow = grdConceptoPagarVC.Rows - 1
            
            'Id
            grdConceptoPagarVC.TextMatrix(nRow, 0) = rsUsuarioRecaudo!CID
            'Servicio
            grdConceptoPagarVC.TextMatrix(nRow, 1) = IIf(Len(Trim(rsUsuarioRecaudo!cServicio)) = 0, Space(200) & ".", Trim(rsUsuarioRecaudo!cServicio))
            'Concepto
            grdConceptoPagarVC.TextMatrix(nRow, 2) = rsUsuarioRecaudo!cConcepto
            'Fec Vec
            grdConceptoPagarVC.TextMatrix(nRow, 3) = rsUsuarioRecaudo!dFechaVencimiento
            'Moneda
            grdConceptoPagarVC.TextMatrix(nRow, 4) = rsUsuarioRecaudo!cmoneda
            'Importe
            grdConceptoPagarVC.TextMatrix(nRow, 5) = rsUsuarioRecaudo!nImporte
            grdConceptoPagarVC.TextMatrix(nRow, 5) = Format$(grdConceptoPagarVC.TextMatrix(nRow, 5), "#0.00")
            'Mora
            grdConceptoPagarVC.TextMatrix(nRow, 6) = rsUsuarioRecaudo!nMora
            grdConceptoPagarVC.TextMatrix(nRow, 6) = Format$(grdConceptoPagarVC.TextMatrix(nRow, 6), "#0.00")
            'copy ID
            grdConceptoPagarVC.TextMatrix(nRow, 8) = rsUsuarioRecaudo!CID 'se copia valor para efectuar validaciones
            
        End If
        rsUsuarioRecaudo.MoveNext
    Loop
    rsUsuarioRecaudo.MoveFirst
Else
    MsgBox "Usted no selecciono ningun usuario", vbExclamation, "Aviso"
    limpiaDetalle
End If

End Sub

Private Sub cmdBuscarPersonaVI_Click()
        
    Set rsUsuarioRecaudo = New Recordset
    Set rsUsuarioRecaudo = frmBuscarUsuarioConvenio.Inicio(Trim(txtCodigoBusConvenio.Text))
    
    If Not rsUsuarioRecaudo Is Nothing Then
    
        If Not (rsUsuarioRecaudo.EOF And rsUsuarioRecaudo.BOF) Then

            txtNombreClienteVI.Text = rsUsuarioRecaudo!cNomCliente
            txtDOIVI.Text = rsUsuarioRecaudo!cDOI
            txtCodigoIDVI.Text = rsUsuarioRecaudo!cCodCliente

        End If
        
        Dim nRow  As Integer
        LimpiaFlex grdConceptoPagarVI
        calculoSubTotalComisioVI
        Do While Not rsUsuarioRecaudo.EOF
        
            grdConceptoPagarVI.AdicionaFila
            
            nRow = grdConceptoPagarVI.Rows - 1
            
            grdConceptoPagarVI.TextMatrix(nRow, 0) = rsUsuarioRecaudo!CID
            grdConceptoPagarVI.TextMatrix(nRow, 1) = rsUsuarioRecaudo!cServicio
            grdConceptoPagarVI.TextMatrix(nRow, 2) = rsUsuarioRecaudo!cConcepto
            grdConceptoPagarVI.TextMatrix(nRow, 3) = rsUsuarioRecaudo!cmoneda
            grdConceptoPagarVI.TextMatrix(nRow, 5) = Format(rsUsuarioRecaudo!nDeudaActual, "#,##0.00")
            'grdConceptoPagarVI.TextMatrix(nRow, 5) = Format$(grdConceptoPagarVI.TextMatrix(nRow, 5), "#,##0.00")
            
            ReDim Preserve Importes(nRow)
            Importes(nRow) = CDbl(rsUsuarioRecaudo!nDeudaActual)
            
            rsUsuarioRecaudo.MoveNext
            
        Loop
        rsUsuarioRecaudo.MoveFirst
         
        
    Else
        MsgBox "Usted no selecciono ninguna Empresa", vbExclamation, "Aviso"
         limpiaDetalle
         
    End If

End Sub

Private Sub cmdBuscarConvenio_Click()

    Dim rsRecaudo As Recordset
    Set rsRecaudo = New Recordset
    Set rsRecaudo = frmBuscarConvenio.Inicio

    ' Limpia Detalle de Convevio
    limpiaDetalle
    Set rsUsuarioRecaudo = Nothing
    nmoneda = 1
    nComisionEmpresa = 0#
    strCuenta = ""
    
    If Not rsRecaudo Is Nothing Then
        If Not (rsRecaudo.EOF And rsRecaudo.BOF) Then
            
            Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
            Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento
            Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
            Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento
            
            'Valida el estado de la cuenta, solo permite el pago a cuentas activas
            If Not objValidar.ValidaEstadoCuenta(rsRecaudo!cCtaCod, False) Then
                MsgBox "La cuenta del convenio seleccionado NO Tiene un estado valido para la operacion", vbExclamation + vbDefaultButton1, "Aviso"
                limpiaDetalle
                limipiarCabecera
                Set ClsServicioRecaudo = Nothing
                Set objValidar = Nothing
                Exit Sub
                
            End If

            'Verifica si en el convenio seleccionado, hay clientes a pagar
            If Mid(rsRecaudo!cCodConvenio, 14, 2) <> "SV" Then
                If Not ClsServicioRecaudo.getCantidadRegistrosConvenio(rsRecaudo!cCodConvenio) Then
                    If Mid(rsRecaudo!cCodConvenio, 14, 2) = "VP" Then
                        MsgBox "El convenio seleccionado no posee importes a pagar", vbExclamation, "Aviso"
                    Else
                        MsgBox "El convenio seleccionado no posee clientes a pagar", vbExclamation, "Aviso"
                    End If
                    limpiaDetalle
                    limipiarCabecera
                    Set ClsServicioRecaudo = Nothing
                    Set objValidar = Nothing
                    Exit Sub
                End If
            End If
            
            txtCodigoBusConvenio.Text = rsRecaudo!cCodConvenio
            txtNombreConvenio.Text = rsRecaudo!cNombreConvenio
            txtCodigoEmpresa.Text = rsRecaudo!cPersCod
            txtNombreEmpresa.Text = rsRecaudo!cPersNombre
            nmoneda = IIf(Mid(rsRecaudo!cCtaCod, 9, 1) = 1, 1, 2)
            strCuenta = rsRecaudo!cCtaCod
            
            lblDescripcionCodigoVC.Caption = rsRecaudo!cdescripcion
            lblDescripcionCodigoVI.Caption = rsRecaudo!cdescripcion
            lblDescripcionCodigoSV.Caption = rsRecaudo!cdescripcion
            lblDescripcionCodigoVP.Caption = rsRecaudo!cdescripcion
            
            vistaPestana (txtCodigoBusConvenio.Text)
            
            If Mid(txtCodigoBusConvenio.Text, 14, 2) = "VP" Then
                grdConceptoPagarVX.CargaCombo ClsServicioRecaudo.getListaConceptosCobrarXConvenioPV(Trim(txtCodigoBusConvenio.Text))
            End If
            
            txtCodigoBusConvenio.Locked = True
            cmdGuardar.Enabled = True
            cmdCancelar.Enabled = True
            Set ClsServicioRecaudo = Nothing
            Set objValidar = Nothing
            
        End If
    Else
         MsgBox "No selecciono ningun convenio", vbExclamation, "Aviso"
         limipiarCabecera
    End If

End Sub

'Limpia la parte superior del formulario
Private Sub limipiarCabecera()

    txtCodigoBusConvenio.Text = ""
     txtNombreConvenio.Text = ""
     txtCodigoEmpresa.Text = ""
     txtNombreEmpresa.Text = ""
     strCuenta = ""
     stContenedorValidacion.Enabled = False
     txtCodigoBusConvenio.SetFocus
     txtCodigoBusConvenio.Locked = False
     cmdGuardar.Enabled = False
     cmdCancelar.Enabled = False
     
End Sub
     
'Limpia la parte central del formulario
Private Sub limpiaDetalle()
    
    Set rsUsuarioRecaudo = Nothing
    nmoneda = 1
    nComisionEmpresa = 0#
    
' ***** Validacion Completa *****
    'grdConceptoPagarVC.Clear
    grdConceptoPagarVC.Rows = 2
    LimpiaFlex grdConceptoPagarVC
    txtCodigoIDVC.Text = ""
    txtNombreClienteVC.Text = ""
    txtDOIVC.Text = ""
    txtSubTotalVC.Text = "0.00"
    txtComisionVC.Text = "0.00"
    txtTotalVC.Text = "0.00"
    
' ***** Validacion Incompleta *****
    grdConceptoPagarVI.Clear
    grdConceptoPagarVI.Rows = 2
    grdConceptoPagarVI.FormaCabecera
    txtCodigoIDVI.Text = ""
    txtNombreClienteVI.Text = ""
    txtDOIVI.Text = ""
    txtSubTotalVI.Text = "0.00"
    txtComisionVI.Text = "0.00"
    txtTotalVI.Text = "0.00"

' ***** Sin Validacion *****
    grdConceptoPagarSV.Clear
    grdConceptoPagarSV.Rows = 2
    grdConceptoPagarSV.FormaCabecera
    txtOtroCodigoSV.Text = ""
    txtNombreClienteSV.Text = ""
    txtDOISV.Text = ""
    cboTipoDOISV.ListIndex = 1
    txtSubTotalSV.Text = "0.00"
    txtComisionSV.Text = "0.00"
    txtTotalSV.Text = "0.00"
    
' ***** Validacion por Importe *****
    grdConceptoPagarVX.Clear
    grdConceptoPagarVX.Rows = 2
    grdConceptoPagarVX.FormaCabecera
    txtOtroCodigoVX.Text = ""
    txtNombreClienteVX.Text = ""
    txtDOIVX.Text = ""
    txtSubTotalVX.Text = "0.00"
    txtComisionVX.Text = "0.00"
    txtTotalVX.Text = "0.00"
    
    ReDim Preserve Importes(0)
    
End Sub

Private Sub cmdCancelar_Click()

    If MsgBox("¿Está seguro de cancelar la operacion?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
        limpiaDetalle
        limipiarCabecera
        
    End If
    
End Sub

Private Sub cmdGuardar_Click()
    
    Dim nI As Double
    Dim nCount As Double, nTmpValid As Integer
    Dim sMensaje As String
    Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento

    On Error GoTo error

    If txtCodigoBusConvenio.Text = "" Or txtNombreConvenio.Text = "" Or _
        txtCodigoEmpresa.Text = "" Or txtNombreEmpresa.Text = "" Then
        
        MsgBox "No se Selecciono Ningun Convenio", vbExclamation, "Aviso"
        Exit Sub
                
    End If
    
    ' ===== Valida el estado de la cuenta
    If Not objValidar.ValidaEstadoCuenta(strCuenta, False) Then
        MsgBox "Cuenta NO Tiene un estado valido para la Operacion, consulte con el Asistente de Agencia.", vbExclamation, "aviso"
        Exit Sub
    End If
        
    If nConvenioSeleccionado = 0 Then
            MsgBox "No se selecciono ningun convenio", vbExclamation, "Aviso"
            Exit Sub
    ' Validando Convenio:  Sin Validacion
    ElseIf nConvenioSeleccionado = Convenio_SV Then
        If cboTipoDOISV.ListIndex = 0 Then
            If Len(Trim(txtDOISV.Text)) < 8 Then
                MsgBox "El tipo de documento seleccionado admite 8 caracteres", vbExclamation, "Aviso"
                Exit Sub
            End If
        ElseIf cboTipoDOISV.ListIndex = 1 Then
            If Len(Trim(txtDOISV.Text)) < 11 Then
                MsgBox "El tipo de documento seleccionado admite 11 caracteres", vbExclamation, "Aviso"
                Exit Sub
            End If
        End If
        sMensaje = Trim(validaSVVP)
        If Not sMensaje = "" Then
            MsgBox sMensaje, vbExclamation, "Aviso"
            Exit Sub
        End If
        'Validadndo Pago de Minims y Maximos
        
        
    ' Validando Convenio:  Validacion por Importes
    ElseIf nConvenioSeleccionado = Convenio_VP Then
            If cboTipoDOIVX.ListIndex = 0 Then
                If Len(Trim(txtDOIVX.Text)) < 8 Then
                    MsgBox "El tipo de documento seleccionado admite 8 caracteres", vbExclamation, "Aviso"
                    Exit Sub
                End If
            ElseIf cboTipoDOIVX.ListIndex = 1 Then
                If Len(Trim(txtDOIVX.Text)) < 11 Then
                    MsgBox "El tipo de documento seleccionado admite 11 caracteres", vbExclamation, "Aviso"
                    Exit Sub
                End If
            End If
            sMensaje = Trim(validaSVVP)
            If Not sMensaje = "" Then
                MsgBox sMensaje, vbExclamation, "Aviso"
                Exit Sub
            End If
    ' Validando Convenio:  Validacion Completa
    ElseIf nConvenioSeleccionado = Convenio_VC Or nConvenioSeleccionado = Convenio_VCM Then
            For nI = 1 To grdConceptoPagarVC.Rows - 1
                'If grdConceptoPagarVC.TextMatrix(nI, 5) = "." Then
                If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
                    nCount = nCount + 1
                    If grdConceptoPagarVC.TextMatrix(nI, 0) <> grdConceptoPagarVC.TextMatrix(nI, 8) Then
                        nTmpValid = nTmpValid + 1
                    End If
                End If
            Next
            If nCount = 0 Then
                MsgBox "No se seleccionaron montos a pagar", vbExclamation, "Aviso"
                Exit Sub
            End If
            nCount = 0
            If nTmpValid > 0 Then
                MsgBox "Los valores del campo ID cambiaron, vuelva a cargar los valores", vbExclamation, "Aviso"
                Exit Sub
            End If
            nTmpValid = 0
            If rsUsuarioRecaudo Is Nothing Then
                MsgBox "No se seleccionó un usuario del convenio", vbExclamation, "Aviso"
                Exit Sub
            End If

            For nI = 1 To grdConceptoPagarVC.Rows - 1
                'If val(grdConceptoPagarVC.TextMatrix(nI, 4)) <= 0 Then
                If val(grdConceptoPagarVC.TextMatrix(nI, 5)) <= 0 Then
                    nCount = nCount + 1
                End If
            Next
            If nCount > 0 Then
                MsgBox "No puede efectuar pagos que sean iguales a 0.00 ", vbExclamation, "Aviso"
                Exit Sub
            End If
    ' Validando Convenio:  Validacion Incompleta
    ElseIf nConvenioSeleccionado = Convenio_VI Then
            For nI = 1 To grdConceptoPagarVI.Rows - 1
                If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
                    nCount = nCount + 1
                End If
            Next
            If rsUsuarioRecaudo Is Nothing Then
                MsgBox "No se seleccionó un usuario del convenio", vbExclamation, "Aviso"
                Exit Sub
            End If
            If nCount = 0 Then
                MsgBox "No se seleccionaron montos a pagar", vbExclamation, "Aviso"
                Exit Sub
            End If
            nCount = 0
            
            ' Aplicando Validacion de Importe Minimos y Maximos
            Dim rsImpoteMinMax As ADODB.Recordset
            Dim nMinimo, nMaximo As Double
                   
            For nI = 1 To grdConceptoPagarVI.Rows - 1
                If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
                    If val(grdConceptoPagarVI.TextMatrix(nI, 5)) <= 0 Then
                        nCount = nCount + 1
                    End If
                    
                    Set rsImpoteMinMax = ObtenerImporteMinMax(Trim(txtCodigoBusConvenio.Text), grdConceptoPagarVI.TextMatrix(nI, 0))
                    If Not rsImpoteMinMax Is Nothing Then
                        If Not rsImpoteMinMax.EOF And Not rsImpoteMinMax.BOF Then
                            nMinimo = rsImpoteMinMax!nPagoMin
                            nMaximo = rsImpoteMinMax!nPagoMax
                        Else
                            nMinimo = 0
                            nMaximo = 9999
                        End If
                    Else
                            nMinimo = 0
                            nMaximo = 9999
                    End If
                    
                    If CDbl(grdConceptoPagarVI.TextMatrix(nI, 5)) < nMinimo Or _
                       CDbl(grdConceptoPagarVI.TextMatrix(nI, 5)) > nMaximo Then
                       
                       MsgBox " Verificar importes. Límites por registro: Monto mínimo " & nMinimo & " - Monto máximo " & nMaximo _
                       , vbExclamation, "Aviso"
                                         
                       Set rsImpoteMinMax = Nothing
                       Exit Sub
                        
                    End If
                    
                End If
            Next
            If nCount > 0 Then
                MsgBox "No puede efectuar pagos que sean iguales a 0.00 ", vbExclamation, "Aviso"
                Exit Sub
            End If
    End If
    ' ===== Fin Validacion
    
    'Procediendo a grabar
    If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        
        Dim sImpresion As String
        Dim sMovNro As String, sMovNro2 As String
        Dim oCont As COMNContabilidad.NCOMContFunciones
        Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Sleep 1000
        sMovNro2 = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        If nConvenioSeleccionado = Convenio_VC Or nConvenioSeleccionado = Convenio_VCM Then
            registrarVC sMovNro, sMovNro2
        ElseIf nConvenioSeleccionado = Convenio_VI Then
            registrarVI sMovNro, sMovNro2
        ElseIf nConvenioSeleccionado = Convenio_SV Then
            registrarSV sMovNro, sMovNro2
        ElseIf nConvenioSeleccionado = Convenio_VP Then
            registrarVP sMovNro, sMovNro2
        End If
    End If
    
    Exit Sub
    
error:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Function validaSVVP() As String

    Dim nI As Integer
    Dim nCont As Integer
        
    If nConvenioSeleccionado = Convenio_SV Then
        nCont = 0
        If Trim(txtNombreClienteSV.Text) = "" Then
            validaSVVP = "Verifcar datos de cliente"
            Exit Function
        ElseIf Trim(txtDOISV.Text) = "" Then
            validaSVVP = "Verifcar datos de cliente"
            Exit Function
'        ElseIf Trim(txtOtroCodigoSV.Text) = "" Then
'            validaSVVP = "Verifcar el codigo ingresado"
'            Exit Function
        ElseIf cboTipoDOISV.ListIndex = -1 Then
            validaSVVP = "Verifcar el tipo de DOI seleccionado"
            Exit Function
        End If
                
        ' Aplicando Validacion de Importe Minimos y Maximos
        Dim rsImpoteMinMax As ADODB.Recordset
        Dim nMinimo, nMaximo As Double
        Set rsImpoteMinMax = ObtenerImporteMinMax(Trim(txtCodigoBusConvenio.Text))
        
        If Not rsImpoteMinMax Is Nothing Then
            If Not rsImpoteMinMax.EOF And Not rsImpoteMinMax.BOF Then
                nMinimo = rsImpoteMinMax!nPagoMin
                nMaximo = rsImpoteMinMax!nPagoMax
            End If
        End If
        
        If CDbl(txtSubTotalSV.Text) < nMinimo Or _
           CDbl(txtSubTotalSV.Text) > nMaximo Then

            validaSVVP = " Verificar importes. Límites por operacion: Monto mínimo " & nMinimo & " - Monto máximo " & nMaximo
            Set rsImpoteMinMax = Nothing
            Exit Function

        End If
        
        For nI = 1 To grdConceptoPagarSV.Rows - 1
            nCont = nCont + 1
            If Trim(grdConceptoPagarSV.TextMatrix(nI, 1)) = "" Or _
            val(grdConceptoPagarSV.TextMatrix(nI, 3)) <= 0 Then
                
                validaSVVP = "Verificar conceptos y montos a pagar"
                Exit Function
                
            End If
                       
        Next
        
        If nCont > 0 Then
            validaSVVP = ""
        Else
            validaSVVP = "No se registraron montos y conceptos a pagar"
        End If
        Exit Function
    ElseIf nConvenioSeleccionado = Convenio_VP Then
    
         nCont = 0
         If Trim(txtNombreClienteVX.Text) = "" Then
                validaSVVP = "Verifcar datos de cliente"
                Exit Function
        ElseIf Trim(txtDOIVX.Text) = "" Then
                validaSVVP = "Verifcar datos de cliente"
                Exit Function
'        ElseIf Trim(txtOtroCodigoVX.Text) = "" Then
'                validaSVVP = "Verifcar el codigo ingresado"
'                Exit Function
        ElseIf cboTipoDOIVX.ListIndex = -1 Then
                validaSVVP = "Verifcar el tipo de DOI seleccionado"
                Exit Function
        End If
                
        For nI = 1 To grdConceptoPagarVX.Rows - 1
        
            nCont = nCont + 1
            If Trim(grdConceptoPagarVX.TextMatrix(nI, 1)) = "" Or _
                val(grdConceptoPagarVX.TextMatrix(nI, 3)) <= 0 Then
                
                validaSVVP = "Verificar conceptos y montos a pagar"
                Exit Function

            End If
            
        Next
        
        If nCont > 0 Then
            validaSVVP = ""
        Else
            validaSVVP = "No se registraron montos y conceptos a pagar"
        End If
        Exit Function
        
    End If
    
End Function

Private Sub registrarVP(ByVal sMovNro As String, sMovNro2 As String)
    
    Dim cCodigoConvenio As String
    Dim cCodigoCliente As String
    Dim nTipoDOI As Integer
    Dim cDOI As String
    Dim cNombreCliente As String
    Dim dFechaCobro As Date
    Dim sBoleta As String
    Dim nI As Integer
    Dim nTamanio As Integer
    Dim arrDetalleCobro() As String 'Contiene detalle de conceptos de pago
    Dim cCadenaData As String
    Dim nDeudaActual As Double
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim nITF As Double          'Para validacion de ITF
    Dim nRedondeoITF As Double  'Para validacion de ITF
 
    nTamanio = 0

    ' Validando si Aplica ITF
    
    nITF = Format(fgITFCalculaImpuesto(txtSubTotalVX.Text), "#,##0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))

    If nRedondeoITF > 0 Then
    
       nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
       
    End If

    cCodigoConvenio = Trim(txtCodigoBusConvenio.Text)
    cCodigoCliente = Trim(txtOtroCodigoVX.Text)
    nTipoDOI = Trim(Right(Trim(cboTipoDOIVX.Text), 2))
    cDOI = Trim(txtDOIVX.Text)
    cNombreCliente = Trim(txtNombreClienteVX.Text)
    dFechaCobro = gdFecSis
    
    ' Se llena un arreglo que contendra los conceptos y montos
    ' de pago correspondientes a la operacion
    For nI = 1 To grdConceptoPagarVX.Rows - 1
        
            nTamanio = nTamanio + 1
            
            ReDim Preserve arrDetalleCobro(nTamanio)
            
            'id de la Trama
            cCadenaData = grdConceptoPagarVX.TextMatrix(nI, 4) & "|"
            'Servicio
            cCadenaData = cCadenaData & "" & "|"
            'Concepto
            cCadenaData = cCadenaData & Mid(grdConceptoPagarVX.TextMatrix(nI, 1), _
                                            InStr(1, grdConceptoPagarVX.TextMatrix(nI, 1), "-") + 2, _
                                            Len(grdConceptoPagarVX.TextMatrix(nI, 1))) & "|" ' se hizo una ultima modificacion
            
            'Importe
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVX.TextMatrix(nI, 3)) & "|"
            nDeudaActual = 0#
            'Deuda Actual
            cCadenaData = cCadenaData & CDbl(nDeudaActual) & "|"
            'Monto Cobro
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVX.TextMatrix(nI, 3)) & "|"
            'Estado
            cCadenaData = cCadenaData & Pagado & "|"
            'Mora
            cCadenaData = cCadenaData & "0.00|"
            'Fecha Vencimiento
            'cCadenaData = cCadenaData & Format(CDate(gdFecSis), "yyyyMMdd") & "|"
            cCadenaData = cCadenaData & "|"
            
            arrDetalleCobro(nTamanio - 1) = cCadenaData

    Next
    
    Dim ClsServicioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set ClsServicioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If ClsServicioRecaudo.registrarCobroServicioConvenio(arrDetalleCobro, cCodigoConvenio, _
                                                         cCodigoCliente, nTipoDOI, cDOI, cNombreCliente, _
                                                         dFechaCobro, sMovNro, sMovNro2, nComisionEmpresa, _
                                                         Trim(txtComisionVX.Text), nmoneda, strCuenta, nITF) Then
        
        MsgBox "La operacion se realizo correctamente", vbInformation, "Aviso"
        
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        sBoleta = ClsServicioRecaudo.ImprimeVaucherRecaudo(ClsServicioRecaudo.getDatosVaucherRecaudo(clsCap.GetnMovNro(sMovNro)), gbImpTMU, gsCodUser)
        
        limpiaDetalle
        'limipiarCabecera
        txtNombreClienteVX.SetFocus
        
        Do
          clsprevio.PrintSpool sLpt, sBoleta
          
        Loop While MsgBox("Desea Reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbYes
        
        Set clsCap = Nothing
        
    Else
        MsgBox "No se pudo realizar en cobro del recaudo", vbExclamation, "Aviso"
    End If
    
End Sub

Private Sub registrarSV(ByVal sMovNro As String, sMovNro2 As String)

    Dim cCodigoConvenio As String
    Dim cCodigoCliente As String
    Dim nTipoDOI As Integer
    Dim cDOI As String
    Dim cNombreCliente As String
    Dim dFechaCobro As Date
    Dim sBoleta As String
    Dim nI As Integer
    Dim nTamanio As Integer
    Dim arrDetalleCobro() As String
    Dim cCadenaData As String
    Dim nDeudaActual As Double
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim nITF As Double          ' Para Validar ITF
    Dim nRedondeoITF As Double  ' Para Validar ITF
    
    On Error GoTo ErrGraba
    
    ' Validando si Aplica ITF
    nITF = Format(fgITFCalculaImpuesto(txtSubTotalSV.Text), "#,##0.00")
    
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))

    If nRedondeoITF > 0 Then
    
       nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
       
    End If
            
    cCodigoConvenio = Trim(txtCodigoBusConvenio.Text)
    cCodigoCliente = Trim(txtOtroCodigoSV.Text)
    nTipoDOI = Trim(Right(Trim(cboTipoDOISV.Text), 2))
    cDOI = Trim(txtDOISV.Text)
    cNombreCliente = Trim(txtNombreClienteSV.Text)
    dFechaCobro = gdFecSis
     
    nTamanio = 0
    
    ' Se llena un arreglo que contendra los conceptos y montos
    ' de pago correspondientes a la operacion
    For nI = 1 To grdConceptoPagarSV.Rows - 1
        
        nTamanio = nTamanio + 1
        
        ReDim Preserve arrDetalleCobro(nTamanio)
        'id de la Trama
        cCadenaData = "" & "|"
        'Servicio
        cCadenaData = cCadenaData & "" & "|"
        'Concepto
        cCadenaData = cCadenaData & grdConceptoPagarSV.TextMatrix(nI, 1) & "|"
        'Importe
        cCadenaData = cCadenaData & CDbl(grdConceptoPagarSV.TextMatrix(nI, 3)) & "|"
        'Deuda Actual
        cCadenaData = cCadenaData & "0.00|"
        'Monto Cobro
        cCadenaData = cCadenaData & CDbl(grdConceptoPagarSV.TextMatrix(nI, 3)) & "|"
        'Estado
        cCadenaData = cCadenaData & Pagado & "|"
        'Mora
        cCadenaData = cCadenaData & "0.00|"
        'Fecha Vencimiento
        cCadenaData = cCadenaData & "|"
        
        arrDetalleCobro(nTamanio - 1) = cCadenaData
        
    Next
    
    Dim ClsServicioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set ClsServicioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If ClsServicioRecaudo.registrarCobroServicioConvenio(arrDetalleCobro, cCodigoConvenio, _
                            cCodigoCliente, nTipoDOI, cDOI, cNombreCliente, dFechaCobro, sMovNro, sMovNro2, nComisionEmpresa, Trim(txtComisionSV.Text), nmoneda, strCuenta, nITF) Then
        
        MsgBox "La operacion se realizo correctamente", vbInformation, "Aviso"
        
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        sBoleta = ClsServicioRecaudo.ImprimeVaucherRecaudo(ClsServicioRecaudo.getDatosVaucherRecaudo(clsCap.GetnMovNro(sMovNro)), gbImpTMU, gsCodUser)
        
        limpiaDetalle
        'limipiarCabecera
        txtNombreClienteSV.SetFocus
        
        Do
            clsprevio.PrintSpool sLpt, sBoleta
            
        Loop While MsgBox("Desea Reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbYes
        
        Set clsCap = Nothing
        
    Else
    
        MsgBox "No se pudo realizar en cobro del recaudo", vbExclamation, "Aviso"
        
    End If
    
    Exit Sub
    
ErrGraba:
    Set clsCap = Nothing
    MsgBox Err.Description, vbCritical, "Aviso"
     
End Sub

Private Sub registrarVI(ByVal sMovNro As String, sMovNro2 As String)
    
    Dim cCodigoConvenio As String
    Dim cCodigoCliente As String
    Dim nTipoDOI As Integer
    Dim cDOI As String
    Dim cNombreCliente As String
    Dim dFechaCobro As Date
    Dim sBoleta As String
    Dim nI As Integer
    Dim nTamanio As Integer
    Dim arrDetalleCobro() As String
    Dim cCadenaData As String
    Dim nDeudaActual As Double
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim nITF As Double          ' Validar ITF
    Dim nRedondeoITF As Double  ' Validar ITF
    
    nTamanio = 0
           
    ' Validando si Aplica ITF
    nITF = Format(fgITFCalculaImpuesto(txtSubTotalVI.Text), "#,##0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))
    If nRedondeoITF > 0 Then
       nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
    End If
                
    cCodigoConvenio = Trim(txtCodigoBusConvenio.Text)
    cCodigoCliente = Trim(txtCodigoIDVI.Text)
    nTipoDOI = rsUsuarioRecaudo!nTipoDOI
    cDOI = Trim(txtDOIVI.Text)
    cNombreCliente = Trim(txtNombreClienteVI.Text)
    dFechaCobro = gdFecSis
    
    ' Se llena un arreglo que contendra los conceptos y montos
    ' de pago correspondientes a la operacion
    For nI = 1 To grdConceptoPagarVI.Rows - 1
        If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
            nTamanio = nTamanio + 1
            ReDim Preserve arrDetalleCobro(nTamanio)
            'Id Trama
            cCadenaData = grdConceptoPagarVI.TextMatrix(nI, 0) & "|"
            'Servicio
            cCadenaData = cCadenaData & grdConceptoPagarVI.TextMatrix(nI, 1) & "|"
            'Concepto
            cCadenaData = cCadenaData & grdConceptoPagarVI.TextMatrix(nI, 2) & "|"
            'Importe
            cCadenaData = cCadenaData & CDbl(rsUsuarioRecaudo!nDeudaActual) & "|"
            nDeudaActual = CDbl(rsUsuarioRecaudo!nDeudaActual) - CDbl(grdConceptoPagarVI.TextMatrix(nI, 5))
            'Deuda Actual
            cCadenaData = cCadenaData & nDeudaActual & "|"
            'Monto Cobro
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVI.TextMatrix(nI, 5)) & "|"
            'Estado
            If nDeudaActual <= 0 Then
                cCadenaData = cCadenaData & Pagado & "|"
            Else
                cCadenaData = cCadenaData & Pagando & "|"
            End If
            'Mora
            cCadenaData = cCadenaData & "0.00|"
            'Fecha Vencimiento
            cCadenaData = cCadenaData & "|"
            
            arrDetalleCobro(nTamanio - 1) = cCadenaData
        
        End If
        rsUsuarioRecaudo.MoveNext
        
    Next
    rsUsuarioRecaudo.MoveFirst
    If cCadenaData = "" Then
    
        MsgBox "No se seleccionó ningún concepto de pago", vbExclamation, "Aviso"
        Exit Sub
        
    End If
    
    Dim ClsServicioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set ClsServicioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If ClsServicioRecaudo.registrarCobroServicioConvenio(arrDetalleCobro, cCodigoConvenio, cCodigoCliente, _
                                                        nTipoDOI, cDOI, cNombreCliente, dFechaCobro, _
                                                        sMovNro, sMovNro2, nComisionEmpresa, Trim(txtComisionVI.Text), nmoneda, strCuenta, nITF) Then
    
        MsgBox "La operacion se realizo correctamente", vbInformation, "Aviso"
                
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        sBoleta = ClsServicioRecaudo.ImprimeVaucherRecaudo(ClsServicioRecaudo.getDatosVaucherRecaudo(clsCap.GetnMovNro(sMovNro)), gbImpTMU, gsCodUser)
        
        limpiaDetalle
        'limipiarCabecera
        cmdBuscarPersonaVI.SetFocus
        Do
            clsprevio.PrintSpool sLpt, sBoleta
        Loop While MsgBox("Desea Reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbYes
        Set clsCap = Nothing
    Else
        MsgBox "No se pudo realizar en cobro del recaudo", vbExclamation, "Aviso"
    End If
    
End Sub
Private Sub registrarVC(ByVal sMovNro As String, sMovNro2 As String)
    
    Dim cCodigoConvenio As String
    Dim cCodigoCliente As String
    Dim nTipoDOI As Integer
    Dim cDOI As String
    Dim cNombreCliente As String
    Dim dFechaCobro As Date
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim sBoleta As String
    Dim nITF As Double          ' Validar ITF
    Dim nRedondeoITF As Double  ' Validar ITF
    
    'Validando si Aplica ITF
            
    nITF = Format(fgITFCalculaImpuesto(txtSubTotalVC.Text), "#,##0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))
    If nRedondeoITF > 0 Then
       nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
    End If
    
    'Datos Unicos -----------------------------------------
    cCodigoConvenio = Trim(txtCodigoBusConvenio.Text)
    cCodigoCliente = Trim(txtCodigoIDVC.Text)
    nTipoDOI = rsUsuarioRecaudo!nTipoDOI
    cDOI = Trim(txtDOIVC.Text)
    cNombreCliente = Trim(txtNombreClienteVC.Text)
    dFechaCobro = gdFecSis
    ' Datos Unicos -----------------------------------------
    
    Dim nI As Integer
    Dim nTamanio As Integer
    Dim arrDetalleCobro() As String
    Dim cCadenaData As String
    
    nTamanio = 0
        
    For nI = 1 To grdConceptoPagarVC.Rows - 1
        'If grdConceptoPagarVC.TextMatrix(nI, 5) = "." Then
        If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
            nTamanio = nTamanio + 1
            ReDim Preserve arrDetalleCobro(nTamanio)
            'ID de la Trama
            cCadenaData = Trim(grdConceptoPagarVC.TextMatrix(nI, 0)) & "|"
            'Servicio
            cCadenaData = cCadenaData & Replace(Trim(grdConceptoPagarVC.TextMatrix(nI, 1)), ".", "") & "|"
            'Concepto
            cCadenaData = cCadenaData & Trim(grdConceptoPagarVC.TextMatrix(nI, 2)) & "|"
            'Importe
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVC.TextMatrix(nI, 5)) & "|"
            'Deuda Actual
            cCadenaData = cCadenaData & "0.00" & "|"
            'Monto Cobro
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVC.TextMatrix(nI, 5)) & "|"
            'Estado
            cCadenaData = cCadenaData & Pagado & "|"
            'Mora
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVC.TextMatrix(nI, 6)) & "|"
            'Fecha Vencimiento
            If nConvenioSeleccionado = Convenio_VCM Then
                cCadenaData = cCadenaData & Format(CDate(grdConceptoPagarVC.TextMatrix(nI, 3)), "yyyyMMdd") & "|"
            Else
                cCadenaData = cCadenaData & "|"
            End If
            arrDetalleCobro(nTamanio - 1) = cCadenaData
        End If
    Next
    
    If cCadenaData = "" Then
        MsgBox "no se seleccionó ningún concepto de pago", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim ClsServicioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set ClsServicioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If ClsServicioRecaudo.registrarCobroServicioConvenio(arrDetalleCobro, _
                        cCodigoConvenio, cCodigoCliente, nTipoDOI, cDOI, cNombreCliente, _
                        dFechaCobro, sMovNro, sMovNro2, nComisionEmpresa, Trim(txtComisionVC.Text), _
                        nmoneda, strCuenta, nITF) Then
                                                        
        MsgBox "La operacion se realizo correctamente", vbInformation, "Aviso"
        
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        sBoleta = ClsServicioRecaudo.ImprimeVaucherRecaudo(ClsServicioRecaudo.getDatosVaucherRecaudo(clsCap.GetnMovNro(sMovNro)), gbImpTMU, gsCodUser)
        
        limpiaDetalle
        'limipiarCabecera
        cmdBuscarPersonaVC.SetFocus
        
        Do
        clsprevio.PrintSpool sLpt, sBoleta
        Loop While MsgBox("Desea Reimprimir el voucher?", vbInformation + vbYesNo, "Aviso") = vbYes
        Set clsCap = Nothing
    Else
        MsgBox "No se pudo realizar en cobro del recaudo", vbExclamation, "Aviso"
    End If
           
End Sub

Private Sub cmdQuitarSV_Click()
    grdConceptoPagarSV.EliminaFila grdConceptoPagarSV.row
    grdConceptoPagarSV_OnCellChange 0, 0
End Sub

Private Sub cmdQuitarSV_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    cmdGuardar.SetFocus
    End If

End Sub

Private Sub cmdQuitarVX_Click()
      grdConceptoPagarVX.EliminaFila grdConceptoPagarVX.row
      grdConceptoPagarVX_OnChangeCombo
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub vistaPestana(ByVal strConvenio As String)
    Dim cTipoConvenio As String
    cTipoConvenio = Mid(Trim(strConvenio), 14, 2)
    stContenedorValidacion.Enabled = True
    stContenedorValidacion.TabVisible(0) = False
    stContenedorValidacion.TabVisible(1) = False
    stContenedorValidacion.TabVisible(2) = False
    stContenedorValidacion.TabVisible(3) = False
        
    If cTipoConvenio = "VC" Then
        stContenedorValidacion.TabVisible(0) = True
        stContenedorValidacion.Tab = 0
                
        
        If Mid(Trim(strConvenio), 16, 1) <> "M" Then
            'CONVENIO VC
            grdConceptoPagarVC.ColWidth(1) = 1905
            grdConceptoPagarVC.ColWidth(2) = 1905
            grdConceptoPagarVC.ColWidth(3) = 0
            grdConceptoPagarVC.ColWidth(4) = 1200
            grdConceptoPagarVC.ColWidth(5) = 1305
            grdConceptoPagarVC.ColWidth(6) = 0
            grdConceptoPagarVC.ColWidth(7) = 1005
            stContenedorValidacion.TabCaption(0) = "Validación Completa"
            nConvenioSeleccionado = Convenio_VC
        Else
            'CONVENIO MYPE
            grdConceptoPagarVC.ColWidth(1) = 1000
            grdConceptoPagarVC.ColWidth(2) = 1100
            grdConceptoPagarVC.ColWidth(3) = 1200
            grdConceptoPagarVC.ColWidth(4) = 850
            grdConceptoPagarVC.ColWidth(5) = 1100
            grdConceptoPagarVC.ColWidth(6) = 1100
            grdConceptoPagarVC.ColWidth(7) = 650
            stContenedorValidacion.TabCaption(0) = "Valid. Comp. MYPE"
            nConvenioSeleccionado = Convenio_VCM
        End If
                        
        cmdBuscarPersonaVC.SetFocus
        
    ElseIf cTipoConvenio = "VI" Then
        stContenedorValidacion.TabVisible(1) = True
        stContenedorValidacion.Tab = 1
        nConvenioSeleccionado = Convenio_VI
        cmdBuscarPersonaVI.SetFocus
        
    ElseIf cTipoConvenio = "SV" Then
        stContenedorValidacion.TabVisible(2) = True
        stContenedorValidacion.Tab = 2
        nConvenioSeleccionado = Convenio_SV
        cboTipoDOISV.ListIndex = 0
        txtNombreClienteSV.SetFocus
        
    ElseIf cTipoConvenio = "VP" Then
        stContenedorValidacion.TabVisible(3) = True
        stContenedorValidacion.Tab = 3
        nConvenioSeleccionado = Convenio_VP
        cboTipoDOIVX.ListIndex = 0
        txtNombreClienteVX.SetFocus
        
    End If
    
End Sub

Private Sub Form_Load()

    stContenedorValidacion.TabVisible(0) = True
    stContenedorValidacion.TabVisible(1) = False
    stContenedorValidacion.TabVisible(2) = False
    stContenedorValidacion.TabVisible(3) = False
    lblDescripcionCodigoVC.Caption = ""
    stContenedorValidacion.Enabled = False
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarSV_GotFocus()
    bFocusGrid = True
End Sub

Private Sub grdConceptoPagarSV_LostFocus()
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarSV_OnCellChange(pnRow As Long, pnCol As Long)
    grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 1) = UCase(grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 1))
    Dim nSubTotal As Double
    Dim nI As Integer
    If grdConceptoPagarSV.col = 3 Then
        If Not IsNumeric(grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3)) Then
            MsgBox "Valor de la celda no es numerico", vbExclamation, "Aviso"
            grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3) = "0.00"
        Else
            If grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3) < 0 Then
             MsgBox "Valor de la celda es menor que 0.00", vbExclamation, "Aviso"
             grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3) = "0.00"
            End If
        End If
    End If
    For nI = 1 To grdConceptoPagarSV.Rows - 1
        If grdConceptoPagarSV.TextMatrix(nI, 3) <> "" Then
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarSV.TextMatrix(nI, 3)), 0, grdConceptoPagarSV.TextMatrix(nI, 3)))
        End If
    Next
    
    txtSubTotalSV.Text = Format(nSubTotal, "#,##0.00")
    'calculoComision grdConceptoPagarSV, txtComisionSV, -1, txtSubTotalSV
    
    calculoComisionNew 'RIRO20150320
    
    ' Suma y muesta el Total
    Dim strValor As Double
    strValor = Math.Round((CDbl(txtSubTotalSV.Text) + CDbl(txtComisionSV.Text)), 2)
    txtTotalSV.Text = Format(strValor, "#,##0.00")
    
    If grdConceptoPagarSV.col = 1 And _
       Trim(grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3)) = "" Then
        
        grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3) = "0.00"
        
    End If
     
End Sub

Private Sub grdConceptoPagarSV_OnRowChange(pnRow As Long, pnCol As Long)

    If pnRow = grdConceptoPagarSV.Rows - 1 And pnCol = 3 Then
    
          DoEvents
          cmdAgregarSV.SetFocus
          
    End If

End Sub

Private Sub grdConceptoPagarSV_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

    Dim sColumnas() As String
    sColumnas = Split(grdConceptoPagarSV.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "No es posible editar este campo", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub grdConceptoPagarVC_GotFocus()
    bFocusGrid = True
End Sub

Private Sub grdConceptoPagarVC_LostFocus()
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarVC_KeyPress(KeyAscii As Integer)

    If grdConceptoPagarVC.col = 4 Then
        KeyAscii = 13
        Exit Sub
    End If

End Sub
Private Sub grdConceptoPagarVC_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    
    Dim nSubTotal As Double
    Dim nI As Integer
    
    Dim nColImp As Double ' columna importe
    Dim nColMor As Double ' columna mora
    
    nColImp = 5
    nColMor = 6
    
    For nI = 1 To grdConceptoPagarVC.Rows - 1
        If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
            'Sumando Importe
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarVC.TextMatrix(nI, nColImp)), 0, grdConceptoPagarVC.TextMatrix(nI, nColImp)))
            'Sumando Mora
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarVC.TextMatrix(nI, nColMor)), 0, grdConceptoPagarVC.TextMatrix(nI, nColMor)))
        End If
    Next
    txtSubTotalVC.Text = Format(nSubTotal, "#,##00.00")
    'calculoComision grdConceptoPagarVC, txtComisionVC, 7, txtSubTotalVC
    calculoComisionNew 'RIRO20150320
    
    ' Suma y muesta el Total
    Dim strValor As Double
    strValor = Math.Round((CDbl(txtSubTotalVC.Text) + CDbl(txtComisionVC.Text)), 2)
    txtTotalVC.Text = Format(strValor, "#,##0.00")
    
End Sub

'RIRO20150326 ****************
Private Sub calculoComisionNew()
        
    Dim nI As Integer
    Dim i As Integer
    Dim nColImp As Double 'columna importe
    Dim nColMor As Double 'columna mora

    Dim vLista() As Double '
    ReDim vLista(3, 0) 'columnas de arreglo vLista:1: Autoincrement, 2: importe, 3: mora
        
    Dim oServicio As New COMDCaptaServicios.DCOMServicioRecaudo
        
    If nConvenioSeleccionado = Convenio_VC Or nConvenioSeleccionado = Convenio_VCM Then
        nColImp = 5
        nColMor = 6
        For nI = 1 To grdConceptoPagarVC.Rows - 1
            If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
                i = i + 1
                ReDim Preserve vLista(3, 0 To i)
                vLista(1, i) = i
                vLista(2, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarVC.TextMatrix(nI, nColImp)), 0, grdConceptoPagarVC.TextMatrix(nI, nColImp)))
                vLista(3, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarVC.TextMatrix(nI, nColMor)), 0, grdConceptoPagarVC.TextMatrix(nI, nColMor)))
            End If
        Next
        If UBound(vLista, 2) = 0 Then
            nComisionCliente = 0
            nComisionEmpresa = 0
            txtComisionVC.Text = "0.00"
            Exit Sub
        End If
        oServicio.CalculaComisionRecaudo Trim(txtCodigoBusConvenio.Text), Trim(txtCodigoIDVC.Text), vLista, nComisionCliente, nComisionEmpresa
        txtComisionVC.Text = Format(nComisionCliente, "#,##0.00")
        
    ElseIf nConvenioSeleccionado = Convenio_VI Then
        nColImp = 5
        nColMor = 0 ' no hay mora
        For nI = 1 To grdConceptoPagarVI.Rows - 1
            If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
                i = i + 1
                ReDim Preserve vLista(3, 0 To i)
                vLista(1, i) = i
                vLista(2, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarVI.TextMatrix(nI, nColImp)), 0, grdConceptoPagarVI.TextMatrix(nI, nColImp)))
                vLista(3, i) = 0
            End If
        Next
        If UBound(vLista, 2) = 0 Then
            nComisionCliente = 0
            nComisionEmpresa = 0
            txtComisionVI.Text = "0.00"
            Exit Sub
        End If
        oServicio.CalculaComisionRecaudo Trim(txtCodigoBusConvenio.Text), Trim(txtCodigoIDVI.Text), vLista, nComisionCliente, nComisionEmpresa
        txtComisionVI.Text = Format(nComisionCliente, "#,##0.00")
        
    ElseIf nConvenioSeleccionado = Convenio_SV Then
        nColImp = 3
        nColMor = 0 ' no hay mora
        For nI = 1 To grdConceptoPagarSV.Rows - 1
            If Trim(grdConceptoPagarSV.TextMatrix(nI, 3)) <> "" Then
                i = i + 1
                ReDim Preserve vLista(3, 0 To i)
                vLista(1, i) = i
                vLista(2, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarSV.TextMatrix(nI, nColImp)), 0, grdConceptoPagarSV.TextMatrix(nI, nColImp)))
                vLista(3, i) = 0
            End If
        Next
        If UBound(vLista, 2) = 0 Then
            nComisionCliente = 0
            nComisionEmpresa = 0
            txtComisionSV.Text = "0.00"
            Exit Sub
        End If
        oServicio.CalculaComisionRecaudo Trim(txtCodigoBusConvenio.Text), Trim(txtOtroCodigoSV.Text), vLista, nComisionCliente, nComisionEmpresa   ' revisar
        txtComisionSV.Text = Format(nComisionCliente, "#,##0.00")
        
    ElseIf nConvenioSeleccionado = Convenio_VP Then
        nColImp = 3
        nColMor = 0 ' no hay mora
        Dim nSubTotal As Double
        For nI = 1 To grdConceptoPagarVX.Rows - 1
            If grdConceptoPagarVX.TextMatrix(nI, 3) <> "" Then
                ReDim Preserve vLista(3, 0 To i)
                vLista(1, i) = i
                vLista(2, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarVX.TextMatrix(nI, nColImp)), 0, grdConceptoPagarVX.TextMatrix(nI, nColImp)))
                vLista(3, i) = 0
            End If
        Next
        If UBound(vLista, 2) = 0 Then
            nComisionCliente = 0
            nComisionEmpresa = 0
            txtComisionVX.Text = "0.00"
            Exit Sub
        End If
        oServicio.CalculaComisionRecaudo Trim(txtCodigoBusConvenio.Text), Trim(txtOtroCodigoVX.Text), vLista, nComisionCliente, nComisionEmpresa
        txtComisionVX.Text = Format(nComisionCliente, "#,##0.00")
        
    End If
End Sub
'END RIRO *************************

Private Sub calculoComision(ByVal flxTemp As FlexEdit, _
                            ByVal txtComision As TextBox, _
                            Optional ByVal nIndiceCheck As Integer, _
                            Optional ByVal txtSubTotal As TextBox)
    Dim nI As Integer
    Dim nCount As Integer
    Dim nMonto As Double
    Dim nnMonto As Double
    Dim nnMontoEmp As Double
    
    Dim nMontoTemporal As Double
                    
    '================================================================================================
    If nIndiceCheck = -1 Then
        Dim rsSV_VC As Recordset
        Set rsSV_VC = New Recordset
                
        Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
        Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
        Set rsSV_VC = ClsServicioRecaudo.getBUscarConvenioXCodigo(Trim(txtCodigoBusConvenio.Text))
        Set rsUsuarioRecaudo = rsSV_VC
    End If
    '================================================================================================
    If Not rsUsuarioRecaudo Is Nothing Or nConvenioSeleccionado = Convenio_SV Then
        If rsUsuarioRecaudo!nTipoCobro = TipoCobro_porConcepto Then ' Valida Tipo Cobro X Concepto
            If rsUsuarioRecaudo!nTipoCalculo = TipoCalculo_fijo Then 'Valida Tipo Calculo -> Calculo Fijo
                
                ' Caso: Por Concepto->Fijo==============================================================={
                    
                    ' Valida Distribucion es Fija
                    If rsUsuarioRecaudo!nDistribucion = Distribucion_fijo Then

                       For nI = 1 To flxTemp.Rows - 1
                           If nIndiceCheck = -1 Then nIndiceCheck = 0
                           If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                                nCount = nCount + 1
                           End If
                       Next
                       
                       txtComision.Text = Format(nCount * rsUsuarioRecaudo!nDistCliente, "#,##0.00")
                       nComisionEmpresa = Format(nCount * rsUsuarioRecaudo!nDistEmpresa, "#,##0.00")
                       txtComisionVC.Text = Format$(nCount * rsUsuarioRecaudo!nDistCliente, "#,##0.00")
                       
                    ' Distribucio Porcentaje
                    Else
                       For nI = 1 To flxTemp.Rows - 1
                            If nIndiceCheck = -1 Then nIndiceCheck = 0
                            If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                                nCount = nCount + 1
                            End If
                       Next
                      txtComision.Text = Format(nCount * ((rsUsuarioRecaudo!nDistCliente) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00" _
                                                   )
                      nComisionEmpresa = Format(nCount * ((rsUsuarioRecaudo!nDistEmpresa) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00" _
                                                   )
                       txtComisionVC.Text = Format(nCount * ((rsUsuarioRecaudo!nDistCliente) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00")
                                                 
                    End If
                 ' Caso: Por Concepto->Fijo===============================================================}

            Else ' Valida Tipo Calculo -> Porcentual
                 ' Caso: Por Concepto->Porcentual==============================================================={
                    If rsUsuarioRecaudo!nDistribucion = Distribucion_porcentaje Then ' Valida Distribucion es Porcentual
                       For nI = 1 To flxTemp.Rows - 1
                       If nIndiceCheck = -1 Then nIndiceCheck = 0
                            If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                            nMontoTemporal = (((rsUsuarioRecaudo!nComision) / 100) * _
                                            IIf(Not IsNumeric(flxTemp.TextMatrix(nI, IIf(nIndiceCheck = 0, 3, IIf(nIndiceCheck = 7, 5, 5)))), 0, _
                                                              flxTemp.TextMatrix(nI, IIf(nIndiceCheck = 0, 3, IIf(nIndiceCheck = 7, 5, 5)))))
                                                              
                            'Solo si es convenio Validacion Completa MYPE
                            If Mid(Trim(txtCodigoBusConvenio.Text), 14, 3) = "VCM" Then
                                nMontoTemporal = nMontoTemporal + (((rsUsuarioRecaudo!nComision) / 100) * _
                                                                    IIf(Not IsNumeric(flxTemp.TextMatrix(nI, 6)), 0, _
                                                                                      flxTemp.TextMatrix(nI, 6)))
                            End If
                            'Fin de calculo para Convenios MYPE
                                                              
                            If nMontoTemporal < rsUsuarioRecaudo!nMinimo Then
                                nMontoTemporal = rsUsuarioRecaudo!nMinimo
                            ElseIf nMontoTemporal > rsUsuarioRecaudo!nMaximo Then
                                nMontoTemporal = rsUsuarioRecaudo!nMaximo
                            End If
                            nnMontoEmp = nnMontoEmp + nMontoTemporal * ((rsUsuarioRecaudo!nDistEmpresa) / 100)
                            nnMonto = nnMonto + nMontoTemporal * ((rsUsuarioRecaudo!nDistCliente) / 100)
                            nCount = nCount + 1 '
                           End If
                       Next
                       If nCount = 0 Then
                        txtComision.Text = "0.00"
                        txtTotalVC.Text = "0.00"
                        Exit Sub
                       End If
                       txtComision.Text = Format(nnMonto, "#,##0.00")
                       nComisionEmpresa = nnMontoEmp
                    End If
                ' Caso: Por Concepto->Porcentual===============================================================}
            End If
            
        Else 'Valida Tipo Cobro X Operacion
            If rsUsuarioRecaudo!nTipoCalculo = TipoCalculo_fijo Then 'Valida Tipo Calculo -> Calculo Fijo
                ' Caso: Por Concepto->Fijo==============================================================={
                    If rsUsuarioRecaudo!nDistribucion = Distribucion_fijo Then ' Valida Distribucion es Fija
                       For nI = 1 To flxTemp.Rows - 1
                       If nIndiceCheck = -1 Then nIndiceCheck = 0
                           If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                               nCount = nCount + 1
                           End If
                       Next
                       If nCount = 0 Then
                        txtComision.Text = "0.00"
                        txtTotalVC.Text = "0.00"
                        Exit Sub
                       End If
                       txtComision.Text = Format(rsUsuarioRecaudo!nDistCliente, "#,##0.00")
                       nComisionEmpresa = Format(rsUsuarioRecaudo!nDistEmpresa, "#,##0.00")
                    Else ' Distribucio Porcentaje
                       For nI = 1 To flxTemp.Rows - 1
                       If nIndiceCheck = -1 Then nIndiceCheck = 0
                          If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                              nCount = nCount + 1
                          End If
                       Next
                       If nCount = 0 Then
                        txtComision.Text = "0.00"
                        txtTotalVC.Text = "0.00"
                        Exit Sub
                       End If
                       txtComision.Text = Format(((rsUsuarioRecaudo!nDistCliente) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00" _
                                                   )
                        nComisionEmpresa = Format(((rsUsuarioRecaudo!nDistEmpresa) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00" _
                                                   )
                    End If
                ' Caso: Por Concepto->Fijo===============================================================}
            
            Else ' Valida Tipo Calculo -> Porcentual
               ' Caso: Por Concepto->Porcentual==============================================================={
                    If rsUsuarioRecaudo!nDistribucion = Distribucion_porcentaje Then ' Valida Distribucion es Porcentual
                                           
                       For nI = 1 To flxTemp.Rows - 1
                       If nIndiceCheck = -1 Then nIndiceCheck = 0
                           If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                               nCount = nCount + 1
                           End If
                       Next
                       If nCount = 0 Then
                        txtComision.Text = "0.00"
                        txtTotalVC.Text = "0.00"
                        Exit Sub
                       End If
                       'Dim nMonto As Double
                       
                       nMontoTemporal = (CDbl(txtSubTotal.Text) * ((rsUsuarioRecaudo!nComision) / 100))
                       
                       If nMontoTemporal < rsUsuarioRecaudo!nMinimo Then
                           nMontoTemporal = rsUsuarioRecaudo!nMinimo
                       ElseIf nMontoTemporal > rsUsuarioRecaudo!nMaximo Then
                           nMontoTemporal = rsUsuarioRecaudo!nMaximo
                       End If
                       
                       nMonto = nMontoTemporal * ((rsUsuarioRecaudo!nDistCliente) / 100)
                       
                       nnMontoEmp = nMontoTemporal * ((rsUsuarioRecaudo!nDistEmpresa) / 100)
                       
                       txtComision.Text = Format(nMonto, "#,##0.00")
                       nComisionEmpresa = nnMontoEmp
                                                 
                    End If
                    
                ' Caso: Por Concepto->Porcentual===============================================================}
                
            End If
                        
        End If
        
    End If
    
End Sub

Private Sub grdConceptoPagarVC_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

Dim sColumnas() As String

sColumnas = Split(grdConceptoPagarVC.ColumnasAEditar, "-")

If sColumnas(pnCol) = "X" Then
Cancel = False
MsgBox "No es posible editar este campo", vbInformation, "Aviso"
End If

End Sub

Private Sub grdConceptoPagarVI_GotFocus()
    bFocusGrid = True
End Sub

Private Sub grdConceptoPagarVI_LostFocus()
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarVI_OnCellChange(pnRow As Long, pnCol As Long)
    
    If grdConceptoPagarVI.col = 5 Then
    
        If Not IsNumeric(grdConceptoPagarVI.TextMatrix(grdConceptoPagarVI.row, 5)) Then
        
            MsgBox "Valor de la celda no es numerico", vbExclamation, "Aviso"
            grdConceptoPagarVI.TextMatrix(grdConceptoPagarVI.row, 5) = "0.00"
            
        Else
            
            If grdConceptoPagarVI.TextMatrix(grdConceptoPagarVI.row, 5) < 0 Then
            
             MsgBox "Valor de la celda es menor que 0.00", vbExclamation, "Aviso"
             grdConceptoPagarVI.TextMatrix(grdConceptoPagarVI.row, 5) = "0.00"
            
            End If
                        
        End If
    
    End If
    
    calculoSubTotalComisioVI
    
End Sub
Private Sub calculoSubTotalComisioVI()
    Dim nSubTotal As Double
    Dim nI As Integer
    
    For nI = 1 To grdConceptoPagarVI.Rows - 1
        If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarVI.TextMatrix(nI, 5)), 0, grdConceptoPagarVI.TextMatrix(nI, 5)))
        End If
    Next
    
    txtSubTotalVI.Text = Format$(nSubTotal, "#,##0.00")
    'calculoComision grdConceptoPagarVI, txtComisionVI, 4, txtSubTotalVI
    calculoComisionNew 'RIRO20150320
    
    ' Suma y muesta el Total
    Dim strValor As Double
    strValor = Math.Round((CDbl(txtSubTotalVI.Text) + CDbl(txtComisionVI.Text)), 2)
    txtTotalVI.Text = Format(strValor, "#,##0.00")
End Sub
 Private Sub grdConceptoPagarVI_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    calculoSubTotalComisioVI
End Sub

Private Sub grdConceptoPagarVI_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdConceptoPagarVI.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "No es posible editar este campo", vbInformation, "Aviso"
    End If
End Sub

Private Sub grdConceptoPagarVX_GotFocus()
    bFocusGrid = True
End Sub

Private Sub grdConceptoPagarVX_LostFocus()
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarVX_OnChangeCombo()

    Dim cCodigoImporte As String
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    Dim nI As Integer
    
    Dim rs As Recordset
    Set rs = New Recordset
        
    On Error GoTo error
        
    cCodigoImporte = Trim((Right(Trim(grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 1)), 10)))
    
    Set rs = ClsServicioRecaudo.getMontoConceptoServicio(cCodigoImporte, Trim(txtCodigoBusConvenio))
    
    If Not rs.EOF Then
        grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 3) = rs!nDeudaActual
        grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 3) = _
        Format$(grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 3), "#,##0.00")
        grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 4) = cCodigoImporte
    Else
        grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 3) = 0#
    End If

    Dim nSubTotal As Double
    For nI = 1 To grdConceptoPagarVX.Rows - 1
        If grdConceptoPagarVX.TextMatrix(nI, 3) <> "" Then
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarVX.TextMatrix(nI, 3)), 0, grdConceptoPagarVX.TextMatrix(nI, 3)))
        End If
    Next
    txtSubTotalVX.Text = Format(nSubTotal, "#,##0.00")
    'calculoComision grdConceptoPagarVX, txtComisionVX, -1, txtSubTotalVX
    calculoComisionNew 'RIRO20150320
    
    ' Suma y muesta el Total
    Dim strValor As Double
    strValor = Math.Round((CDbl(txtSubTotalVX.Text) + CDbl(txtComisionVX.Text)), 2)
    txtTotalVX.Text = Format(strValor, "#,##0.00")
      
    Exit Sub
    
error:
    
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub grdConceptoPagarVX_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

    Dim sColumnas() As String
    sColumnas = Split(grdConceptoPagarVX.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "No es posible editar este campo", vbInformation, "Aviso"
    End If

End Sub

Private Sub txtCodigoBusConvenio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim strCodigo As String
        strCodigo = txtCodigoBusConvenio.Text
        Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
        Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
        Dim rsRecaudo As Recordset
        Set rsRecaudo = New Recordset
        Set rsRecaudo = ClsServicioRecaudo.getBuscaConvenioXCodigo(strCodigo)
        ' Limpia Detalle de Convevio
        limpiaDetalle
        Set rsUsuarioRecaudo = Nothing
        nmoneda = 1
        nComisionEmpresa = 0#
        If Not rsRecaudo Is Nothing Then
            If Not (rsRecaudo.EOF And rsRecaudo.BOF) Then
                Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento
                Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento
                'Valida el estado de la cuenta, solo permite el pago a cuentas activas
                If Not objValidar.ValidaEstadoCuenta(rsRecaudo!cCtaCod, False) Then
                    MsgBox "La cuenta del convenio seleccionado NO Tiene un estado valido para la operacion", vbExclamation + vbDefaultButton1, "Aviso"
                    limpiaDetalle
                    limipiarCabecera
                    Set ClsServicioRecaudo = Nothing
                    Set objValidar = Nothing
                    Exit Sub
                End If
                'valida que el convenio seleccionado cuente con registros de pago.
                If Mid(rsRecaudo!cCodConvenio, 14, 2) <> "SV" Then
                    'If ClsServicioRecaudo.getBuscarUsuarioRecaudo(, , , rsRecaudo!cCodConvenio).RecordCount = 0 Then
                    If Not ClsServicioRecaudo.getCantidadRegistrosConvenio(rsRecaudo!cCodConvenio) Then
                        If Mid(rsRecaudo!cCodConvenio, 14, 2) = "VP" Then
                            MsgBox "El convenio seleccionado no posee importes a pagar", vbExclamation, "Aviso"
                        Else
                            MsgBox "El convenio seleccionado no posee clientes a pagar", vbExclamation, "Aviso"
                        End If
                        limpiaDetalle
                        limipiarCabecera
                        Exit Sub
                    End If
                End If
                txtCodigoBusConvenio.Text = rsRecaudo!cCodConvenio
                txtNombreConvenio.Text = rsRecaudo!cNombreConvenio
                txtCodigoEmpresa.Text = rsRecaudo!cPersCod
                txtNombreEmpresa.Text = rsRecaudo!cPersNombre
                lblDescripcionCodigoVC.Caption = rsRecaudo!cdescripcion
                lblDescripcionCodigoVI.Caption = rsRecaudo!cdescripcion
                lblDescripcionCodigoSV.Caption = rsRecaudo!cdescripcion
                lblDescripcionCodigoVP.Caption = rsRecaudo!cdescripcion
                nmoneda = IIf(Mid(rsRecaudo!cCtaCod, 9, 1) = 1, 1, 2)
                strCuenta = rsRecaudo!cCtaCod
                vistaPestana (txtCodigoBusConvenio.Text)
                If Mid(txtCodigoBusConvenio.Text, 14, 2) = "VP" Then
                    grdConceptoPagarVX.CargaCombo ClsServicioRecaudo.getListaConceptosCobrarXConvenioPV(Trim(txtCodigoBusConvenio.Text))
                End If
                txtCodigoBusConvenio.Locked = True
                cmdGuardar.Enabled = True
                cmdCancelar.Enabled = True
            End If
        Else
             MsgBox "Usted no selecciono ninguna Empresa", vbExclamation, "Aviso"
    '             limipiarCabeceras
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
Private Sub txtDOISV_Change()

    If txtDOISV.Text <> "" Then
        
        If Not IsNumeric(txtDOISV.Text) Then
            txtDOISV.Text = ""
            MsgBox "Solo debe ingresar datos numéricos", vbInformation, "Aviso"
            Exit Sub
        End If
        
    End If
    
End Sub

Private Sub txtDOISV_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 13 Then
        txtOtroCodigoSV.SetFocus
    End If

End Sub

Private Sub txtDOIVX_Change()

    If txtDOIVX.Text <> "" Then
        
        If Not IsNumeric(txtDOIVX.Text) Then
            txtDOIVX.Text = ""
            MsgBox "Solo debe ingresar datos numéricos", vbInformation, "Aviso"
            Exit Sub
        End If
        
    End If

End Sub

Private Sub txtDOIVX_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtOtroCodigoVX.SetFocus
    End If

End Sub

Private Sub txtNombreClienteSV_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cboTipoDOISV.SetFocus
        
    Else
        KeyAscii = Letras(KeyAscii)
    
    End If

End Sub

Private Sub txtNombreClienteVX_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cboTipoDOIVX.SetFocus
        
    Else
        KeyAscii = Letras(KeyAscii)
            
    End If

End Sub

Private Sub txtOtroCodigoSV_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdAgregarSV.SetFocus
    End If

End Sub

Private Sub txtOtroCodigoVX_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then
        cmdAgregarVX.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
    
End Sub

Private Function ObtenerImporteMinMax(ByVal pCodConvenio As String, Optional pId As String = "") As ADODB.Recordset
   
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Dim rsImporteMinMax As ADODB.Recordset
    
    On Error GoTo errImporteMinMax
    
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    Set rsImporteMinMax = ClsServicioRecaudo.getImporteMinMax(pCodConvenio, pId)
    
    Set ObtenerImporteMinMax = rsImporteMinMax
    Set rsImporteMinMax = Nothing
   
    Exit Function
    
errImporteMinMax:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Function

