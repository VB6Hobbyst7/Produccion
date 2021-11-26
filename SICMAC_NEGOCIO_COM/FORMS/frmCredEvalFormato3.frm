VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredEvalFormato3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos - Evaluación - Formato 3"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10890
   Icon            =   "frmCredEvalFormato3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOperaciones 
      Height          =   615
      Left            =   120
      TabIndex        =   104
      Top             =   9120
      Width           =   10695
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9360
         TabIndex        =   38
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8040
         TabIndex        =   37
         Top             =   180
         Width           =   1170
      End
   End
   Begin TabDlg.SSTab SSTOperaciones 
      Height          =   7575
      Left            =   120
      TabIndex        =   52
      Top             =   1560
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13361
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Activos y Pasivos"
      TabPicture(0)   =   "frmCredEvalFormato3.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Ingresos y Egresos"
      TabPicture(1)   =   "frmCredEvalFormato3.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Propuesta y Comentarios"
      TabPicture(2)   =   "frmCredEvalFormato3.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame13"
      Tab(2).Control(1)=   "Frame12"
      Tab(2).Control(2)=   "Frame11"
      Tab(2).Control(3)=   "Frame10"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame13 
         Caption         =   " Propuesta "
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Width           =   10455
         Begin VB.CommandButton cmdCalcular 
            Caption         =   "Calcular"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            TabIndex        =   31
            Top             =   360
            Width           =   1170
         End
         Begin SICMACT.EditMoney txtCalcMonto 
            Height          =   300
            Left            =   840
            TabIndex        =   28
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin SICMACT.EditMoney txtCalcTEM 
            Height          =   300
            Left            =   5040
            TabIndex        =   30
            Top             =   360
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin Spinner.uSpinner spnCalcCuotas 
            Height          =   315
            Left            =   3000
            TabIndex        =   29
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            Max             =   999
            Min             =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
         End
         Begin VB.Label Label15 
            Caption         =   "Capacidad Pago Total : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   80
            Top             =   1245
            Width           =   1815
         End
         Begin VB.Label lblCapPagoTotal 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8640
            TabIndex        =   45
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label Label11 
            Caption         =   "Capacidad Pago Empr. : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   79
            Top             =   1240
            Width           =   1935
         End
         Begin VB.Label lblCapPagoEmp 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   44
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label Label26 
            Caption         =   "Endeud. Total : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   1240
            Width           =   1335
         End
         Begin VB.Label Label24 
            Caption         =   "Cuota estimada: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   850
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Monto máximo del crédito : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   76
            Top             =   850
            Width           =   2055
         End
         Begin VB.Label lblEndeudTotal 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   43
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label lblCuotaEstima 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   41
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label lblMontoMax 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   42
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label22 
            Caption         =   "Nº Cuotas :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   75
            Top             =   395
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "TEM (%) : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   74
            Top             =   395
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Monto :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   395
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " Verificación "
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
         Height          =   1095
         Left            =   -74880
         TabIndex        =   46
         Top             =   6120
         Width           =   10455
         Begin VB.TextBox txtVerif 
            Height          =   705
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   10215
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " Referencias "
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
         Height          =   2655
         Left            =   -74880
         TabIndex        =   71
         Top             =   3360
         Width           =   10455
         Begin VB.CommandButton cmdAgregarRef 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Top             =   2160
            Width           =   1050
         End
         Begin VB.CommandButton cmdQuitarRef 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   35
            Top             =   2160
            Width           =   1050
         End
         Begin SICMACT.FlexEdit feReferencia 
            Height          =   1815
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   10200
            _ExtentX        =   17992
            _ExtentY        =   3201
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Nº-Nombre-DNI-Teléfono-Referido-DNI REF.-Aux"
            EncabezadosAnchos=   "300-2730-1500-1200-2730-1500-0"
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
            ColumnasAEditar =   "X-1-2-3-4-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "Nº"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " Comentario "
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
         Height          =   1095
         Left            =   -74880
         TabIndex        =   70
         Top             =   2160
         Width           =   10455
         Begin VB.TextBox txtComent 
            Height          =   705
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   10215
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   " Ratios empresariales y personales "
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
         Height          =   1620
         Left            =   120
         TabIndex        =   61
         Top             =   5760
         Width           =   10455
         Begin SICMACT.FlexEdit feRatios 
            Height          =   1275
            Left            =   0
            TabIndex        =   27
            Top             =   240
            Width           =   10320
            _ExtentX        =   18203
            _ExtentY        =   2249
            Rows            =   4
            Cols0           =   8
            FixedCols       =   2
            HighLight       =   1
            EncabezadosNombres=   "-Ratios-End. Corto Plazo-End. Largo Plazo-Liquidez Corriente-Capital de Trabajo-Rotación de Invent.-Tipo"
            EncabezadosAnchos=   "0-2050-1500-1550-1650-1650-1800-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " Estado de Ganancia y Pérdidas Empresa "
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
         Height          =   1020
         Left            =   120
         TabIndex        =   60
         Top             =   4680
         Width           =   10455
         Begin SICMACT.FlexEdit feEstadoGP 
            Height          =   675
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   10200
            _ExtentX        =   17992
            _ExtentY        =   1191
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Ingresos-Egresos-Margen Bruto-Gastos del Negocio-Ingreso Neto"
            EncabezadosAnchos=   "0-1950-1950-1950-1950-1950"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " Ingresos del Negocio "
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
         Height          =   2175
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   5175
         Begin SICMACT.EditMoney txtTCMIngNeg 
            Height          =   300
            Left            =   3840
            TabIndex        =   13
            Top             =   1740
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin SICMACT.FlexEdit feIngNeg 
            Height          =   1455
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   2566
            Rows            =   12
            Cols0           =   7
            FixedCols       =   2
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Venta-Tipo-Prod1-Prod2-Prod3-Resultado"
            EncabezadosAnchos=   "0-1200-0-1000-1000-1000-1000"
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
            ColumnasAEditar =   "X-X-X-3-4-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C-C-C-C"
            FormatosEdit    =   "0-0-0-4-4-4-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label Label9 
            Caption         =   "Tasa de Crecimiento Mensual (TCM) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1785
            Width           =   2775
         End
         Begin VB.Label Label8 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4680
            TabIndex        =   68
            Top             =   1785
            Width           =   255
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Otros Ingresos "
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
         Height          =   1935
         Left            =   120
         TabIndex        =   58
         Top             =   2640
         Width           =   5175
         Begin VB.CommandButton cmdAgregarOtrosIng 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   1500
            Width           =   1050
         End
         Begin VB.CommandButton cmdQuitarOtrosIng 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   20
            Top             =   1500
            Width           =   1050
         End
         Begin SICMACT.FlexEdit feOtrosIng 
            Height          =   1215
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   2143
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto Ingreso-Monto-Aux"
            EncabezadosAnchos=   "0-3000-1500-0"
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
            ColumnasAEditar =   "X-1-2-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-C-C-C"
            FormatosEdit    =   "0-0-4-4"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin SICMACT.EditMoney txtTCMOtrosIng 
            Height          =   300
            Left            =   3960
            TabIndex        =   21
            Top             =   1500
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label7 
            Caption         =   "Inflación :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   67
            Top             =   1545
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   66
            Top             =   1545
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Gastos Familiares "
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
         Height          =   1935
         Left            =   5400
         TabIndex        =   57
         Top             =   2640
         Width           =   5175
         Begin VB.CommandButton cmdAgregarGastoFam 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   1500
            Width           =   1050
         End
         Begin VB.CommandButton cmdQuitarGastoFam 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   24
            Top             =   1500
            Width           =   1050
         End
         Begin SICMACT.FlexEdit feGastoFam 
            Height          =   1215
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   2143
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto Gasto-Monto-Aux"
            EncabezadosAnchos=   "0-3000-1500-0"
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
            ColumnasAEditar =   "X-1-2-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-C-C-C"
            FormatosEdit    =   "0-0-4-4"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin SICMACT.EditMoney txtTCMGastoFam 
            Height          =   300
            Left            =   3960
            TabIndex        =   25
            Top             =   1500
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label4 
            Caption         =   "TCM :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   65
            Top             =   1545
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   64
            Top             =   1545
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Gastos del Negocio "
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
         Height          =   2175
         Left            =   5400
         TabIndex        =   56
         Top             =   360
         Width           =   5175
         Begin VB.CommandButton cmdAgregarGastoNeg 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   1740
            Width           =   1050
         End
         Begin VB.CommandButton cmdQuitarGastoNeg 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   16
            Top             =   1740
            Width           =   1050
         End
         Begin SICMACT.FlexEdit feGastoNeg 
            Height          =   1455
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   2566
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto Gasto-Tipo-Monto-Aux"
            EncabezadosAnchos=   "0-2500-800-1200-0"
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
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-0-3-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin SICMACT.EditMoney txtTCMGastoNeg 
            Height          =   300
            Left            =   3960
            TabIndex        =   17
            Top             =   1740
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label2 
            Caption         =   "TCM :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   63
            Top             =   1785
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   62
            Top             =   1785
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Declaración PDT "
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
         Height          =   1575
         Left            =   -74880
         TabIndex        =   55
         Top             =   5880
         Width           =   10425
         Begin SICMACT.FlexEdit fePDT 
            Height          =   1215
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   10200
            _ExtentX        =   17992
            _ExtentY        =   2143
            Rows            =   3
            Cols0           =   7
            FixedCols       =   2
            HighLight       =   1
            EncabezadosNombres=   "-Mes / Detalle-Tipo----Promedio"
            EncabezadosAnchos=   "400-2000-0-1900-1900-1900-1800"
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
            ColumnasAEditar =   "X-X-X-3-4-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C-C-C-C"
            FormatosEdit    =   "0-0-2-4-4-4-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            CellBackColor   =   -2147483633
         End
         Begin SICMACT.EditMoney txtCompraMes3 
            Height          =   315
            Left            =   2130
            TabIndex        =   96
            Top             =   545
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtVentaMes3 
            Height          =   315
            Left            =   2130
            TabIndex        =   97
            Top             =   840
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtCompraMes2 
            Height          =   315
            Left            =   4030
            TabIndex        =   98
            Top             =   545
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtVentaMes2 
            Height          =   315
            Left            =   4030
            TabIndex        =   99
            Top             =   840
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtCompraMes1 
            Height          =   315
            Left            =   5940
            TabIndex        =   100
            Top             =   545
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtVentaMes1 
            Height          =   315
            Left            =   5940
            TabIndex        =   101
            Top             =   840
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label txtPromMes3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   7840
            TabIndex        =   103
            Top             =   840
            Width           =   1920
         End
         Begin VB.Label lblCompraProm 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   7840
            TabIndex        =   102
            Top             =   540
            Width           =   1920
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Ventas"
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
            Left            =   120
            TabIndex        =   95
            Top             =   845
            Width           =   2025
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Compras"
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
            Left            =   120
            TabIndex        =   94
            Top             =   545
            Width           =   2025
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Prom"
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
            Left            =   7840
            TabIndex        =   93
            Top             =   240
            Width           =   1920
         End
         Begin VB.Label lblMes1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   5940
            TabIndex        =   92
            Top             =   240
            Width           =   1920
         End
         Begin VB.Label lblMes2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   4030
            TabIndex        =   91
            Top             =   240
            Width           =   1920
         End
         Begin VB.Label lblMes3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   2130
            TabIndex        =   90
            Top             =   240
            Width           =   1920
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mes / Detalle"
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
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   2025
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Pasivos "
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
         Height          =   5535
         Left            =   -69480
         TabIndex        =   54
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton cmdAgregarPasNoCorriente 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   4120
            Width           =   1050
         End
         Begin VB.CommandButton cmdQuitarPasNoCorriente 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   10
            Top             =   4120
            Width           =   1050
         End
         Begin VB.CommandButton cmdAgregarPasCorriente 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   1960
            Width           =   1050
         End
         Begin VB.CommandButton cmdQuitarPasCorriente 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   1960
            Width           =   1050
         End
         Begin SICMACT.FlexEdit fePasCorriente 
            Height          =   1695
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   2990
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Pasivo Corriente-P.E.-P.P.-TOTAL"
            EncabezadosAnchos=   "0-1800-800-800-950"
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
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-C"
            FormatosEdit    =   "0-0-2-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin SICMACT.FlexEdit fePasNoCorriente 
            Height          =   1695
            Left            =   120
            TabIndex        =   8
            Top             =   2400
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   2990
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Pasivo No Corriente-P.E.-P.P.-TOTAL"
            EncabezadosAnchos=   "0-1800-800-800-950"
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
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-C"
            FormatosEdit    =   "0-0-2-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOT. PASIVO Y PAT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            TabIndex        =   88
            Top             =   5040
            Width           =   1575
         End
         Begin VB.Label lbTotPasPatrimonioTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3840
            TabIndex        =   87
            Top             =   5040
            Width           =   1125
         End
         Begin VB.Label lbTotPasPatrimonioPP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2610
            TabIndex        =   86
            Top             =   5040
            Width           =   1245
         End
         Begin VB.Label lbTotPasPatrimonioPE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1560
            TabIndex        =   85
            Top             =   5040
            Width           =   1065
         End
         Begin VB.Label lblPatrimonioTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3840
            TabIndex        =   84
            Top             =   4560
            Width           =   1125
         End
         Begin VB.Label lblPatrimonioPP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2610
            TabIndex        =   83
            Top             =   4560
            Width           =   1245
         End
         Begin VB.Label lblPatrimonioPE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1560
            TabIndex        =   82
            Top             =   4560
            Width           =   1065
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PATRIMONIO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            TabIndex        =   81
            Top             =   4560
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Activos "
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
         Height          =   5535
         Left            =   -74880
         TabIndex        =   53
         Top             =   360
         Width           =   5295
         Begin SICMACT.FlexEdit feActivos 
            Height          =   5175
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   5040
            _ExtentX        =   8890
            _ExtentY        =   9128
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Concepto-Tipo-P.E.-P.P.-TOTAL"
            EncabezadosAnchos=   "0-2100-0-1000-1000-1200"
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
            ColumnasAEditar =   "X-X-X-3-4-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R-R-C"
            FormatosEdit    =   "0-0-0-2-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1335
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   2355
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Información del Negocio"
      TabPicture(0)   =   "frmCredEvalFormato3.frx":035E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "spnCuotas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ActXCodCta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCuotaPagar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtMontoSol"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtGiroNeg"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboMontoSol"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.ComboBox cboMontoSol 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCredEvalFormato3.frx":037A
         Left            =   8400
         List            =   "frmCredEvalFormato3.frx":0384
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtGiroNeg 
         Height          =   300
         Left            =   5640
         TabIndex        =   0
         Top             =   420
         Width           =   4875
      End
      Begin SICMACT.EditMoney txtMontoSol 
         Height          =   300
         Left            =   9240
         TabIndex        =   3
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin SICMACT.EditMoney txtCuotaPagar 
         Height          =   300
         Left            =   2520
         TabIndex        =   1
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Crédito"
      End
      Begin Spinner.uSpinner spnCuotas 
         Height          =   315
         Left            =   5640
         TabIndex        =   2
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Max             =   999
         Min             =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
      End
      Begin VB.Label Label5 
         Caption         =   "Probable cuota a pagar (mes) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   885
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Nº Cuotas :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   50
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Monto solicitado :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   49
         Top             =   885
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Giro del Negocio :"
         Height          =   255
         Left            =   4320
         TabIndex        =   48
         Top             =   450
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCredEvalFormato3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredEvalFormato3
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 3
'**               creado segun RFC090-2012
'** Creación : WIOR, 20120903 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim fnTipoCliente As Integer
Dim sCtaCod As String
Dim gsOpeCod As String
Dim fnTipoRegMant As Integer
Dim fnTipoPermiso As Integer
Dim fbPermiteGrabar As Boolean
Dim fbBloqueaTodo As Boolean
Dim lnIndMaximaCapPago As Double
Dim lnIndCuotaUNM As Double
Dim lnIndCuotaExcFam As Double
Dim lnCondLocal As Integer
Dim rsCredEval As ADODB.Recordset
Dim rsInd As ADODB.Recordset
Dim rsDatGastoNeg As ADODB.Recordset
Dim rsDatGastoFam As ADODB.Recordset
Dim rsDatOtrosIng As ADODB.Recordset
Dim rsDatRef As ADODB.Recordset
Dim fsCliente As String
Dim fsUserAnalista  As String

Dim rsDatActivos As ADODB.Recordset
Dim rsDatPasivos As ADODB.Recordset
Dim rsDatPasivosNo As ADODB.Recordset
Dim rsDatPDT As ADODB.Recordset
Dim rsDatPDTDet As ADODB.Recordset
Dim rsDatPatrimonio As ADODB.Recordset
Dim rsDatPasPat As ADODB.Recordset

Dim rsDatEstadoGP As ADODB.Recordset
Dim rsDatRatios As ADODB.Recordset
Dim rsDatIngNeg As ADODB.Recordset
Dim nTasaIngNeg As Double
Dim nTasaGastoNeg As Double
Dim nTasaGastoFam As Double
Dim nTasaOtrosIng As Double

Dim fnPasivoPE As Double
Dim fnPasivoPP As Double
Dim fnPasivoTOTAL As Double

Dim fnActivoPE As Double
Dim fnActivoPP As Double
Dim fnActivoTOTAL As Double


Dim cSPrd As String, cPrd As String
Dim DCredito As COMDCredito.DCOMCredito
Dim objPista As COMManejador.Pista
Dim nFormato, nPersoneria As Integer
Dim fnMontoIni As Double
Dim lnMin As Double, lnMax As Double
Dim lnMinDol As Double, lnMaxDol As Double
Dim nTC As Double
Dim i, j As Integer
Dim sMes1 As String, sMes2 As String, sMes3 As String
Dim nMes1 As Integer, nMes2 As Integer, nMes3 As Integer
Dim nAnio1 As Integer, nAnio2 As Integer, nAnio3 As Integer
Dim nMontoPDT, nMontoAct, nMontoPas, nMontoPasN As Double

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGiroNeg.SetFocus
End If
End Sub

Private Sub cmdAgregarGastoFam_Click()
If feGastoFam.Rows - 1 < 25 Then
    feGastoFam.lbEditarFlex = True
    feGastoFam.AdicionaFila
    feGastoFam.SetFocus
    SendKeys "{Enter}"
Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdAgregarGastoNeg_Click()
If feGastoNeg.Rows - 1 < 25 Then
    feGastoNeg.lbEditarFlex = True
    feGastoNeg.AdicionaFila
    feGastoNeg.SetFocus
    SendKeys "{Enter}"
Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdAgregarOtrosIng_Click()
If feOtrosIng.Rows - 1 < 25 Then
    feOtrosIng.lbEditarFlex = True
    feOtrosIng.AdicionaFila
    feOtrosIng.SetFocus
    SendKeys "{Enter}"
Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdAgregarPasCorriente_Click()
If fePasCorriente.Rows - 1 < 25 Then
    fePasCorriente.lbEditarFlex = True
    fePasCorriente.AdicionaFila
    fePasCorriente.SetFocus
    SendKeys "{Enter}"
Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdAgregarPasNoCorriente_Click()
If fePasNoCorriente.Rows - 1 < 25 Then
    fePasNoCorriente.lbEditarFlex = True
    fePasNoCorriente.AdicionaFila
    fePasNoCorriente.SetFocus
    SendKeys "{Enter}"
Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdAgregarRef_Click()
If feReferencia.Rows - 1 < 25 Then
    feReferencia.lbEditarFlex = True
    feReferencia.AdicionaFila
    feReferencia.SetFocus
    SendKeys "{Enter}"
Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdCalcular_Click()

On Error GoTo ErrorCalcular
    Dim pnTEM As Double, pnMonto As Double
    Dim MontoMax As Double, CuotaEstima As Double
    Dim CuotaUNM As Double, CuotaExcedeFam As Double
    Dim pnCuotas As Integer
    Dim pnFormula As Double
    
    Dim pnEndedTotal As Double
    Dim pnCapEmp As Double
    Dim pnCapTotal As Double
    Dim pnR As Double
    Dim pnIngresos As Double
    Dim nEndedTotal As Double
    Dim nPatrimonio As Double
    Dim oCredito As COMDCredito.DCOMCredito
    Dim rsCredito As ADODB.Recordset
    Dim nArriba As Double
    Dim Abajo As Double
    
    Set oCredito = New COMDCredito.DCOMCredito
    Set rsCredito = oCredito.RecuperaDatosIndicadCredEval(fnTipoCliente, 3, 2)
    If rsCredito.RecordCount > 0 Then
        For i = 0 To rsCredito.RecordCount - 1
            If fnTipoCliente = 1 Then
                If Trim(rsCredito!cIndicadorID) = "IND015" Then
                    lnIndMaximaCapPago = CDbl(Trim(rsCredito!cIndicadorPorc)) / 100
                End If
            ElseIf fnTipoCliente = 2 Then
                If Trim(rsCredito!cIndicadorID) = "IND016" Then
                    lnIndMaximaCapPago = CDbl(Trim(rsCredito!cIndicadorPorc)) / 100
                End If
            End If
            rsCredito.MoveNext
        Next i
    End If
    
    
    pnIngresos = CDbl(IIf(Trim(feEstadoGP.TextMatrix(1, 5)) = "", 0, feEstadoGP.TextMatrix(1, 5)))
    pnR = (pnIngresos + SumarCampo(feOtrosIng, 2)) '- SumarCampo(feGastoFam, 2)


    nPatrimonio = CDbl(IIf(Trim(lblPatrimonioTotal.Caption) = "", 0, lblPatrimonioTotal.Caption))
    
    
    pnMonto = CDbl(txtCalcMonto.Text)
    pnTEM = CDbl(txtCalcTEM.Text) / 100
    pnCuotas = CInt(spnCalcCuotas.valor)
    
    nArriba = ((pnTEM * ((1 + pnTEM) ^ pnCuotas)))
    Abajo = (((1 + pnTEM) ^ pnCuotas) - 1)
    
    pnFormula = IIf(Abajo = 0, 0, (nArriba / IIf(Abajo = 0, 1, Abajo)))
    
    MontoMax = IIf(pnFormula = 0, 0, (pnR * lnIndMaximaCapPago) / IIf(pnFormula = 0, 1, pnFormula))
    CuotaEstima = pnMonto * pnFormula
    
    lblMontoMax.Caption = Format(MontoMax, "#,##0.00")
    lblCuotaEstima.Caption = Format(CuotaEstima, "#,##0.00")
    
    nEndedTotal = SumarCampo(fePasCorriente, 4) + SumarCampo(fePasNoCorriente, 4) + pnMonto
    
    lblEndeudTotal.Caption = Format(IIf(nPatrimonio = 0, 0, (nEndedTotal / IIf(nPatrimonio = 0, 1, nPatrimonio))) * 100, "#,##0.00")
    lblCapPagoEmp.Caption = Format(IIf(pnIngresos = 0, 0, (CuotaEstima / IIf(pnIngresos = 0, 1, pnIngresos))) * 100, "#,##0.00")
    lblCapPagoTotal.Caption = Format(IIf(pnR = 0, 0, (CuotaEstima / IIf(pnR = 0, 1, pnR))), "#,##0.00")
    
    'txtComent.SetFocus
    Exit Sub
ErrorCalcular:
    MsgBox err.Description + ": Verifique que todos los datos esten ingresados", vbCritical, "Error"

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdGrabar_Click()
If validaDatos Then
    Dim oCred As COMNCredito.NCOMCredito
    Dim GrabarDatos As Boolean
    Dim rsGastoNeg As ADODB.Recordset
    Dim rsGastoFam As ADODB.Recordset
    Dim rsOtrosIng As ADODB.Recordset
    Dim rsActivos As ADODB.Recordset
    Dim rsPasivos As ADODB.Recordset
    Dim rsPasivosNo As ADODB.Recordset
    Dim rsPDT As ADODB.Recordset
    Dim rsEstadoGP As ADODB.Recordset
    Dim rsRatios As ADODB.Recordset
    Dim rsRef As ADODB.Recordset
    Dim rsIngNeg As ADODB.Recordset
    
    Set rsGastoNeg = IIf(feGastoNeg.Rows - 1 > 0, feGastoNeg.GetRsNew(), Nothing)
    Set rsGastoFam = IIf(feGastoFam.Rows - 1 > 0, feGastoFam.GetRsNew(), Nothing)
    Set rsOtrosIng = IIf(feOtrosIng.Rows - 1 > 0, feOtrosIng.GetRsNew(), Nothing)
    Set rsActivos = IIf(feActivos.Rows - 1 > 0, feActivos.GetRsNew(), Nothing)
    Set rsPasivos = IIf(fePasCorriente.Rows - 1 > 0, fePasCorriente.GetRsNew(), Nothing)
    Set rsPasivosNo = IIf(fePasNoCorriente.Rows - 1 > 0, fePasNoCorriente.GetRsNew(), Nothing)
    fePDT.TextMatrix(0, 3) = "Mes1"
    fePDT.TextMatrix(0, 4) = "Mes2"
    fePDT.TextMatrix(0, 5) = "Mes3"
    Set rsPDT = IIf(fePDT.Rows - 1 > 0, fePDT.GetRsNew(), Nothing)
    fePDT.TextMatrix(0, 3) = sMes3
    fePDT.TextMatrix(0, 4) = sMes2
    fePDT.TextMatrix(0, 5) = sMes1
    Set rsEstadoGP = IIf(feEstadoGP.Rows - 1 > 0, feEstadoGP.GetRsNew(), Nothing)
    Set rsRatios = IIf(feRatios.Rows - 1 > 0, feRatios.GetRsNew(), Nothing)
    Set rsRef = IIf(feReferencia.Rows - 1 > 0, feReferencia.GetRsNew(), Nothing)
    Set rsIngNeg = IIf(feIngNeg.Rows - 1 > 0, feIngNeg.GetRsNew(), Nothing)
    
    If MsgBox("Esta Seguro de Guardar los Datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oCred = New COMNCredito.NCOMCredito
        If fnTipoPermiso = 3 Then
    
        GrabarDatos = oCred.GrabarRSFormato3y4(sCtaCod, fnTipoRegMant, Trim(txtGiroNeg.Text), CDbl(txtCuotaPagar.Text), CDbl(spnCuotas.valor), _
                                                CInt(Trim(Right(cboMontoSol.Text, 2))), CDbl(txtMontoSol.Text), CDbl(txtCalcMonto.value), _
                                                CDbl(txtCalcTEM.value), CDbl(spnCalcCuotas.valor), CDbl(lblMontoMax.Caption), _
                                                CDbl(lblCuotaEstima.Caption), CDbl(lblEndeudTotal.Caption), CDbl(lblCapPagoEmp.Caption), _
                                                CDbl(lblCapPagoTotal.Caption), 3, Trim(txtComent.Text), rsGastoNeg, rsGastoFam, rsOtrosIng, rsRef, rsActivos, _
                                                rsPasivos, rsPasivosNo, rsPDT, rsEstadoGP, rsRatios, rsIngNeg, _
                                                nMes1, nMes2, nMes3, nAnio1, nAnio2, nAnio3, CDbl(txtTCMIngNeg.Text), _
                                                CDbl(txtTCMGastoNeg.Text), CDbl(txtTCMGastoFam.Text), CDbl(txtTCMOtrosIng.Text), _
                                                CDbl(IIf(Trim(lblPatrimonioPE.Caption) = "", 0, lblPatrimonioPE.Caption)), CDbl(IIf(Trim(lblPatrimonioPP.Caption) = "", 0, lblPatrimonioPP.Caption)), _
                                                CDbl(IIf(Trim(lblPatrimonioTotal.Caption) = "", 0, lblPatrimonioTotal.Caption)), CDbl(IIf(Trim(lbTotPasPatrimonioPE.Caption) = "", 0, lbTotPasPatrimonioPE.Caption)), _
                                                CDbl(IIf(Trim(lbTotPasPatrimonioPP.Caption) = "", 0, lbTotPasPatrimonioPP.Caption)), CDbl(IIf(Trim(lbTotPasPatrimonioTotal.Caption) = "", 0, lbTotPasPatrimonioTotal.Caption)))
        Else
            GrabarDatos = oCred.GrabarCredEvaluacionVerif(sCtaCod, Trim(txtVerif.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
        
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 3", sCtaCod, gCodigoCuenta
        
        If GrabarDatos Then
            If fnTipoRegMant = 1 Then
                MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            Else
                MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
            End If
            Call GeneraExcelFormato
            Unload Me
        Else
            MsgBox "Hubo errores al grabar la información", vbError, "Error"
        End If
    End If

End If
End Sub

Private Sub cmdQuitarGastoFam_Click()
If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    feGastoFam.EliminaFila (feGastoFam.Row)
    Call CalcularGanaciasPerdidas
End If
End Sub

Private Sub cmdQuitarGastoNeg_Click()
If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    feGastoNeg.EliminaFila (feGastoNeg.Row)
    Call CalcularGanaciasPerdidas
End If
End Sub

Private Sub cmdQuitarOtrosIng_Click()
If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    feOtrosIng.EliminaFila (feOtrosIng.Row)
    Call CalcularGanaciasPerdidas
End If
End Sub

Private Sub cmdQuitarPasCorriente_Click()
If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    fePasCorriente.EliminaFila (fePasCorriente.Row)
    Call CalcularActivoPatrimonio
End If
End Sub

Private Sub cmdQuitarPasNoCorriente_Click()
If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    fePasNoCorriente.EliminaFila (fePasNoCorriente.Row)
    Call CalcularActivoPatrimonio
End If
End Sub

Private Sub cmdQuitarRef_Click()
If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    feReferencia.EliminaFila (feReferencia.Row)
End If
End Sub

Private Sub feActivos_OnCellChange(pnRow As Long, pnCol As Long)
nMontoAct = 0
If pnRow = 1 Or pnRow = 5 Or pnRow = 10 Or pnRow = 16 Then
    MsgBox "No se puede Editar este Regsitro", vbInformation, "Aviso"
    feActivos.TextMatrix(pnRow, pnCol) = ""
End If

'FILA MODIFICADA
nMontoAct = 0
For i = 3 To 4
    nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(pnRow, i)) = "", 0, Trim(feActivos.TextMatrix(pnRow, i))))
Next i
feActivos.TextMatrix(pnRow, 5) = Format(nMontoAct, "#0.00")


'Inventario
For i = 3 To 4
    nMontoAct = 0
    For j = 6 To 9
        nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(j, i)) = "", 0, Trim(feActivos.TextMatrix(j, i))))
    Next j
    feActivos.TextMatrix(5, i) = Format(nMontoAct, "#0.00")
Next i
nMontoAct = 0
nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(5, 3)) = "", 0, Trim(feActivos.TextMatrix(5, 3))))
nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(5, 4)) = "", 0, Trim(feActivos.TextMatrix(5, 4))))
feActivos.TextMatrix(5, 5) = Format(nMontoAct, "#0.00")



'Activo Corriente
For i = 3 To 4
    nMontoAct = 0
    For j = 2 To 5
        nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(j, i)) = "", 0, Trim(feActivos.TextMatrix(j, i))))
    Next j
    feActivos.TextMatrix(1, i) = Format(nMontoAct, "#0.00")
Next i
nMontoAct = 0
nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(1, 3)) = "", 0, Trim(feActivos.TextMatrix(1, 3))))
nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(1, 4)) = "", 0, Trim(feActivos.TextMatrix(1, 4))))
feActivos.TextMatrix(1, 5) = Format(nMontoAct, "#0.00")
'Activo Fijo
For i = 3 To 4
    nMontoAct = 0
    For j = 11 To 15
        nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(j, i)) = "", 0, Trim(feActivos.TextMatrix(j, i))))
    Next j
    feActivos.TextMatrix(10, i) = Format(nMontoAct, "#0.00")
Next i
nMontoAct = 0
nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(10, 3)) = "", 0, Trim(feActivos.TextMatrix(10, 3))))
nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(10, 4)) = "", 0, Trim(feActivos.TextMatrix(10, 4))))
feActivos.TextMatrix(10, 5) = Format(nMontoAct, "#0.00")

'Activos Totales
nMontoAct = 0
For i = 3 To 4
    nMontoAct = 0
    nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(1, i)) = "", 0, Trim(feActivos.TextMatrix(1, i))))
    nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(10, i)) = "", 0, Trim(feActivos.TextMatrix(10, i))))
    feActivos.TextMatrix(16, i) = Format(nMontoAct, "#0.00")
Next i
nMontoAct = 0
nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(1, 5)) = "", 0, Trim(feActivos.TextMatrix(1, 5))))
nMontoAct = nMontoAct + CDbl(IIf(Trim(feActivos.TextMatrix(10, 5)) = "", 0, Trim(feActivos.TextMatrix(10, 5))))
feActivos.TextMatrix(16, 5) = Format(nMontoAct, "#0.00")


fnActivoPE = CDbl(IIf(Trim(feActivos.TextMatrix(16, 3)) = "", 0, Trim(feActivos.TextMatrix(16, 3))))
fnActivoPP = CDbl(IIf(Trim(feActivos.TextMatrix(16, 4)) = "", 0, Trim(feActivos.TextMatrix(16, 4))))
fnActivoTOTAL = CDbl(IIf(Trim(feActivos.TextMatrix(16, 5)) = "", 0, Trim(feActivos.TextMatrix(16, 5))))
Call CalcularActivoPatrimonio
End Sub



Private Sub feGastoFam_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 1 Then
    feGastoFam.TextMatrix(pnRow, pnCol) = UCase(feGastoFam.TextMatrix(pnRow, pnCol))
End If

Call CalcularGanaciasPerdidas
End Sub

Private Sub feGastoNeg_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 1 Then
    feGastoNeg.TextMatrix(pnRow, pnCol) = UCase(feGastoNeg.TextMatrix(pnRow, pnCol))
End If

Call CalcularGanaciasPerdidas
End Sub

Private Sub feIngNeg_OnCellChange(pnRow As Long, pnCol As Long)


If pnRow = 8 Or pnRow = 11 Then
    MsgBox "No se puede Editar este Regsitro", vbInformation, "Aviso"
    feIngNeg.TextMatrix(pnRow, pnCol) = ""
End If

nMontoPas = 0
If pnRow = 9 Then
     For i = 3 To 5
        nMontoPas = nMontoPas + CDbl(IIf(Trim(feIngNeg.TextMatrix(pnRow, i)) = "", 0, Trim(feIngNeg.TextMatrix(pnRow, i))))
    Next i
    feIngNeg.TextMatrix(pnRow, 6) = Format(nMontoPas / 3, "#0.00")
ElseIf pnRow = 10 Then
    For i = 3 To 5
        nMontoPas = nMontoPas + CDbl(IIf(Trim(feIngNeg.TextMatrix(pnRow, i)) = "", 0, Trim(feIngNeg.TextMatrix(pnRow, i))))
    Next i
    feIngNeg.TextMatrix(pnRow, 6) = Format(nMontoPas, "#0.00")
Else
    For i = 3 To 5
        nMontoPas = nMontoPas + CDbl(IIf(Trim(feIngNeg.TextMatrix(pnRow, i)) = "", 0, Trim(feIngNeg.TextMatrix(pnRow, i))))
    Next i
    feIngNeg.TextMatrix(pnRow, 6) = Format(nMontoPas * 4, "#0.00")
End If

For i = 3 To 6
    nMontoPas = 0
    For j = 1 To 7
        nMontoPas = nMontoPas + CDbl(IIf(Trim(feIngNeg.TextMatrix(j, i)) = "", 0, Trim(feIngNeg.TextMatrix(j, i))))
    Next j
    feIngNeg.TextMatrix(j, i) = Format(nMontoPas, "#0.00")
Next i

For i = 3 To 6
    nMontoPas = 0
    nMontoPas = nMontoPas + CDbl(IIf(Trim(feIngNeg.TextMatrix(9, i)) = "", 0, Trim(feIngNeg.TextMatrix(9, i))))
    nMontoPas = (nMontoPas / 100) * CDbl(IIf(Trim(feIngNeg.TextMatrix(10, i)) = "", 0, Trim(feIngNeg.TextMatrix(10, i))))
    feIngNeg.TextMatrix(11, i) = Format(nMontoPas, "#0.00")
Next i

nMontoPas = 0
For i = 3 To 5
    nMontoPas = nMontoPas + CDbl(IIf(Trim(feIngNeg.TextMatrix(11, i)) = "", 0, Trim(feIngNeg.TextMatrix(11, i))))
Next i
feIngNeg.TextMatrix(11, 6) = Format(nMontoPas, "#0.00")


Call CalcularGanaciasPerdidas
Call CalcularRatios
End Sub

Private Sub feOtrosIng_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 1 Then
    feOtrosIng.TextMatrix(pnRow, pnCol) = UCase(feOtrosIng.TextMatrix(pnRow, pnCol))
End If
Call CalcularGanaciasPerdidas
End Sub

Private Sub fePasCorriente_OnCellChange(pnRow As Long, pnCol As Long)
If Trim(fePasCorriente.TextMatrix(pnRow, pnCol)) = "." Then
    fePasCorriente.TextMatrix(pnRow, pnCol) = ""
End If
If pnCol = 1 Then
    fePasCorriente.TextMatrix(pnRow, pnCol) = UCase(fePasCorriente.TextMatrix(pnRow, pnCol))
End If

nMontoPas = 0
For i = 2 To 3
    nMontoPas = nMontoPas + CDbl(IIf(Trim(fePasCorriente.TextMatrix(pnRow, i)) = "", 0, Trim(fePasCorriente.TextMatrix(pnRow, i))))
Next i
fePasCorriente.TextMatrix(pnRow, 4) = Format(nMontoPas, "#0.00")


fnPasivoPE = SumarCampo(fePasCorriente, 2) + SumarCampo(fePasNoCorriente, 2)
fnPasivoPP = SumarCampo(fePasCorriente, 3) + SumarCampo(fePasNoCorriente, 3)
fnPasivoTOTAL = SumarCampo(fePasCorriente, 4) + SumarCampo(fePasNoCorriente, 4)
Call CalcularActivoPatrimonio
End Sub


Private Sub fePasNoCorriente_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 1 Then
    fePasNoCorriente.TextMatrix(pnRow, pnCol) = UCase(fePasNoCorriente.TextMatrix(pnRow, pnCol))
End If
nMontoPasN = 0
For i = 2 To 3
    nMontoPasN = nMontoPasN + CDbl(IIf(Trim(fePasNoCorriente.TextMatrix(pnRow, i)) = "", 0, Trim(fePasNoCorriente.TextMatrix(pnRow, i))))
Next i
fePasNoCorriente.TextMatrix(pnRow, 4) = Format(nMontoPasN, "#0.00")


fnPasivoPE = SumarCampo(fePasCorriente, 2) + SumarCampo(fePasNoCorriente, 2)
fnPasivoPP = SumarCampo(fePasCorriente, 3) + SumarCampo(fePasNoCorriente, 3)
fnPasivoTOTAL = SumarCampo(fePasCorriente, 4) + SumarCampo(fePasNoCorriente, 4)
Call CalcularActivoPatrimonio
End Sub

Private Sub fePDT_OnCellChange(pnRow As Long, pnCol As Long)

nMontoPDT = 0
For i = 3 To 5
    nMontoPDT = nMontoPDT + CDbl(IIf(Trim(fePDT.TextMatrix(pnRow, i)) = "", 0, Trim(fePDT.TextMatrix(pnRow, i))))
Next i
fePDT.TextMatrix(pnRow, 6) = Format(nMontoPDT / 3, "#0.00")

End Sub


Private Sub feReferencia_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 1 Or pnCol = 4 Then
    feReferencia.TextMatrix(pnRow, pnCol) = UCase(feReferencia.TextMatrix(pnRow, pnCol))
End If
End Sub


Public Sub Inicio(ByVal psCtaCod As String, ByVal psTipoRegMant As Integer)
Call CargaControlesInicio
Dim oCred As COMNCredito.NCOMCredito
    Dim rsDCredito As ADODB.Recordset
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    
    Set oCred = New COMNCredito.NCOMCredito
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)
    
    sCtaCod = psCtaCod
    fnTipoRegMant = psTipoRegMant
    
    
    ActXCodCta.NroCuenta = sCtaCod
    fnTipoPermiso = oCred.ObtieneTipoPermisoCredEval(gsCodCargo)
   
    Set DCredito = New COMDCredito.DCOMCredito
    Set rsDCredito = DCredito.RecuperaSolicitudDatoBasicos(sCtaCod)
    fnMontoIni = Trim(rsDCredito!nMonto)
    nPersoneria = CInt(rsDCredito!nPersPersoneria)
    fsCliente = Trim(rsDCredito!cPersNombre)
    fsUserAnalista = Trim(rsDCredito!UserAnalista)
    fnTipoCliente = CInt(rsDCredito!nColocCondicion)
    fnTipoCliente = IIf(fnTipoCliente > 1, 2, 1)
    
    If nPersoneria = 1 Then
        feActivos.ColumnasAEditar = "X-X-X-3-4-X"
        fePasCorriente.ColumnasAEditar = "X-1-2-3-X"
        fePasNoCorriente.ColumnasAEditar = "X-1-2-3-X"
    Else
        feActivos.ColumnasAEditar = "X-X-X-3-X-X"
        fePasCorriente.ColumnasAEditar = "X-1-2-X-X"
        fePasNoCorriente.ColumnasAEditar = "X-1-2-X-X"
    End If
    
    cSPrd = Trim(rsDCredito!cTpoProdCod)
    cPrd = Mid(cSPrd, 1, 1) & "00"
    fbPermiteGrabar = False
    fbBloqueaTodo = False
    
    Set rsDCredEval = DCredito.RecuperaColocacCredEval(sCtaCod)
    If fnTipoPermiso = 2 Then
        If rsDCredEval.RecordCount = 0 Then
            MsgBox "El analista no ha registrado la Evaluacion respectiva", vbExclamation, "Aviso"
            fbPermiteGrabar = False
        Else
            fbPermiteGrabar = True
        End If
    End If
    Set rsDCredito = Nothing
    Set rsDCredEval = Nothing
    
    Set rsDColCred = DCredito.RecuperaColocacCred(sCtaCod)
    If rsDColCred!nVerifCredEval = 1 Then
        MsgBox "Ud. no puede editar la evaluación, ya se realizó la verificacion del credito", vbExclamation, "Aviso"
        fbBloqueaTodo = True
    End If
    
    nFormato = DCredito.AsignarFormato(cPrd, cSPrd, fnMontoIni, lnMin, lnMax)
    lnMinDol = lnMin / nTC
    lnMaxDol = lnMax / nTC
    
    Set DCredito = Nothing
    Set oTipoCam = Nothing
    If CargaDatos Then
        If CargaControles(fnTipoPermiso, fbPermiteGrabar, fbBloqueaTodo) Then
            If fnTipoRegMant = 1 Then
                If Not rsCredEval.EOF Then
                    Call Mantenimiento
                    fnTipoRegMant = 2
                Else
                    Call Registro
                    fnTipoRegMant = 1
                End If
            Else
                If rsCredEval.EOF Then
                    Call Registro
                    fnTipoRegMant = 1
                Else
                    Call Mantenimiento
                    fnTipoRegMant = 2
                End If
            End If
        Else
            Unload Me
            Exit Sub
        End If
    Else
        If CargaControles(1, False) Then
        End If
    End If
    
    Me.Show 1
End Sub

Private Function CargaDatos() As Boolean

On Error GoTo ErrorCargaDatos

    Dim oCred As COMNCredito.NCOMCredito
    Dim i As Integer
    Set oCred = New COMNCredito.NCOMCredito
       
    CargaDatos = oCred.CargaDatosCredEvaluacion3y4(sCtaCod, 1, rsCredEval, rsInd, rsDatGastoNeg, _
                                                rsDatGastoFam, rsDatOtrosIng, rsDatRef, rsDatActivos, _
                                                rsDatPasivos, rsDatPasivosNo, rsDatPDT, rsDatPDTDet, _
                                                rsDatPatrimonio, rsDatPasPat, rsDatEstadoGP, rsDatRatios, _
                                                rsDatIngNeg, nTasaIngNeg, nTasaGastoNeg, nTasaGastoFam, _
                                                nTasaOtrosIng)
    

    
    If CargaDatos Then
        For i = 1 To rsInd.RecordCount
            If rsInd!cIndicadorID = "IND001" Or rsInd!cIndicadorID = "IND002" Then lnIndMaximaCapPago = rsInd!cIndicadorPorc / 100
            rsInd.MoveNext
        Next
    End If
    Exit Function
ErrorCargaDatos:
    CargaDatos = False
    MsgBox err.Description + ": Error al carga datos", vbCritical, "Error"
End Function

Private Function CargaControles(ByVal TipoPermiso As Integer, ByVal pPermiteGrabar As Boolean, Optional ByVal pBloqueaTodo As Boolean = False) As Boolean
    If TipoPermiso = 1 Then
        Call HabilitaControles(False, False, False)
        CargaControles = True
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False, True, pPermiteGrabar)
        CargaControles = True
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True, False, True)
        CargaControles = True
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        CargaControles = False
    End If
    If pBloqueaTodo Then
        Call HabilitaControles(False, False, False)
        CargaControles = True
    End If
     If nPersoneria = 1 Then
        feGastoFam.Enabled = True
        cmdAgregarGastoFam.Enabled = True
        cmdQuitarGastoFam.Enabled = True
        txtTCMGastoFam.Enabled = True
    Else
        feGastoFam.Enabled = False
        cmdAgregarGastoFam.Enabled = False
        cmdQuitarGastoFam.Enabled = False
        txtTCMGastoFam.Enabled = False
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean, ByVal pbHabilitaB As Boolean, ByVal pbHabilitaGuardar As Boolean)
    txtGiroNeg.Enabled = pbHabilitaA
    'Flex Edit
    feActivos.lbEditarFlex = pbHabilitaA
    fePasCorriente.lbEditarFlex = pbHabilitaA
    fePasNoCorriente.lbEditarFlex = pbHabilitaA
    fePDT.lbEditarFlex = pbHabilitaA
    feIngNeg.lbEditarFlex = pbHabilitaA
    fePDT.lbEditarFlex = pbHabilitaA
    feGastoNeg.lbEditarFlex = pbHabilitaA
    feGastoFam.lbEditarFlex = pbHabilitaA
    feOtrosIng.lbEditarFlex = pbHabilitaA
    feEstadoGP.lbEditarFlex = pbHabilitaA
    feRatios.lbEditarFlex = pbHabilitaA
    feReferencia.Enabled = pbHabilitaA
    
    'Botones
    cmdAgregarPasCorriente.Enabled = pbHabilitaA
    cmdQuitarPasCorriente.Enabled = pbHabilitaA
    cmdAgregarPasNoCorriente.Enabled = pbHabilitaA
    cmdQuitarPasNoCorriente.Enabled = pbHabilitaA
    cmdAgregarGastoNeg.Enabled = pbHabilitaA
    cmdQuitarGastoNeg.Enabled = pbHabilitaA
    cmdAgregarGastoFam.Enabled = pbHabilitaA
    cmdQuitarGastoFam.Enabled = pbHabilitaA
    cmdAgregarOtrosIng.Enabled = pbHabilitaA
    cmdQuitarOtrosIng.Enabled = pbHabilitaA
    cmdAgregarRef.Enabled = pbHabilitaA
    cmdQuitarRef.Enabled = pbHabilitaA
    cmdCalcular.Enabled = pbHabilitaA
    cmdGrabar.Enabled = pbHabilitaGuardar
     
    txtCuotaPagar.Enabled = pbHabilitaA
    spnCuotas.Enabled = pbHabilitaA
    txtMontoSol.Enabled = pbHabilitaA
    
    txtCalcMonto.Enabled = pbHabilitaA
    txtCalcTEM.Enabled = pbHabilitaA
    spnCalcCuotas.Enabled = pbHabilitaA
    
    txtComent.Enabled = pbHabilitaA
    txtVerif.Enabled = pbHabilitaB
    
    If Mid(sCtaCod, 9, 1) = "2" Then
        Me.txtMontoSol.BackColor = RGB(200, 255, 200)
        Me.txtCuotaPagar.BackColor = RGB(200, 255, 200)
        
        txtCalcMonto.BackColor = RGB(200, 255, 200)
        lblMontoMax.BackColor = RGB(200, 255, 200)
        lblCuotaEstima.BackColor = RGB(200, 255, 200)
        
        Set DCredito = Nothing
    Else
        Me.txtMontoSol.BackColor = &HFFFFFF
        Me.txtCuotaPagar.BackColor = &HFFFFFF
        txtCalcMonto.BackColor = &HFFFFFF
    
        lblMontoMax.BackColor = &HFFFFFF
        lblCuotaEstima.BackColor = &HFFFFFF
    End If
End Function
Private Function Registro()
    gsOpeCod = gCredRegistrarEvaluacionCred
    txtMontoSol.Text = Format(fnMontoIni, "#,##0.00")
    cboMontoSol.ListIndex = IndiceListaCombo(cboMontoSol, Mid(sCtaCod, 9, 1))
    txtCalcMonto.Text = Format(fnMontoIni, "#,##0.00")
    lblCapPagoTotal.Caption = Format(0, "#,##0.00")
    lblMontoMax.Caption = Format(0, "#,##0.00")
    lblCuotaEstima.Caption = Format(0, "#,##0.00")
    lblEndeudTotal.Caption = Format(0, "#,##0.00")
    lblCapPagoEmp.Caption = Format(0, "#,##0.00")
End Function

Private Function Mantenimiento()
    Dim lnFila As Integer
    If fnTipoPermiso = 3 Then
        gsOpeCod = gCredMantenimientoEvaluacionCred
    Else
        gsOpeCod = gCredVerificacionEvaluacionCred
    End If

    txtGiroNeg.Text = rsCredEval!cGiroNeg
    txtCuotaPagar.Text = Format(rsCredEval!cCuotaPagar, "#,##0.00")
    spnCuotas.valor = rsCredEval!nCuotas
    cboMontoSol.ListIndex = IndiceListaCombo(cboMontoSol, rsCredEval!nmoneda)
    txtMontoSol.Text = Format(rsCredEval!nMontoSol, "#,##0.00")
    txtCalcMonto.Text = Format(rsCredEval!nMontoCalc, "#,##0.00")
    txtCalcTEM.Text = Format(rsCredEval!nTEMCalc, "#,##0.00")
    spnCalcCuotas.valor = Format(rsCredEval!nCuotasCalc, "#,##0.00")
    lblMontoMax.Caption = Format(rsCredEval!nMontoMax, "#,##0.00")
    lblCuotaEstima.Caption = Format(rsCredEval!nCuotaEstima, "#,##0.00")
    txtComent.Text = rsCredEval!cComent
    lblEndeudTotal.Caption = Format(rsCredEval!nEndeudTotal, "#,##0.00")
    lblCapPagoEmp.Caption = Format(rsCredEval!nCapPagoEmp, "#,##0.00")
    lblCapPagoTotal.Caption = Format(rsCredEval!nCapPagoTotal, "#,##0.00")
    
    'Gastos Negocio
    Call LimpiaFlex(feGastoNeg)
    Do While Not rsDatGastoNeg.EOF
        feGastoNeg.AdicionaFila
        lnFila = feGastoNeg.Row
        feGastoNeg.TextMatrix(lnFila, 1) = rsDatGastoNeg!cConcepto
        feGastoNeg.TextMatrix(lnFila, 2) = rsDatGastoNeg!cDesc & Space(25) & rsDatGastoNeg!nTpoGasto
        feGastoNeg.TextMatrix(lnFila, 3) = Format(rsDatGastoNeg!nMonto, "#,##0.00")
        
        rsDatGastoNeg.MoveNext
    Loop
    rsDatGastoNeg.Close
    Set rsDatGastoNeg = Nothing
    
    'Gasto Familiares
    Call LimpiaFlex(feGastoFam)
    Do While Not rsDatGastoFam.EOF
        feGastoFam.AdicionaFila
        lnFila = feGastoFam.Row
        feGastoFam.TextMatrix(lnFila, 1) = rsDatGastoFam!cConcepto
        feGastoFam.TextMatrix(lnFila, 2) = Format(rsDatGastoFam!nMonto, "#,##0.00")
        rsDatGastoFam.MoveNext
    Loop
    rsDatGastoFam.Close
    Set rsDatGastoFam = Nothing
    
    'Otros Ingresos
    Call LimpiaFlex(feOtrosIng)
    Do While Not rsDatOtrosIng.EOF
        feOtrosIng.AdicionaFila
        lnFila = feOtrosIng.Row
        feOtrosIng.TextMatrix(lnFila, 1) = rsDatOtrosIng!cConcepto
        feOtrosIng.TextMatrix(lnFila, 2) = Format(rsDatOtrosIng!nMonto, "#,##0.00")
        rsDatOtrosIng.MoveNext
    Loop
    rsDatOtrosIng.Close
    Set rsDatOtrosIng = Nothing
    
    'Referencia
    Call LimpiaFlex(feReferencia)
    Do While Not rsDatRef.EOF
        feReferencia.AdicionaFila
        lnFila = feReferencia.Row
        feReferencia.TextMatrix(lnFila, 1) = rsDatRef!cNombre
        feReferencia.TextMatrix(lnFila, 2) = rsDatRef!cDNI
        feReferencia.TextMatrix(lnFila, 3) = rsDatRef!cTelef
        feReferencia.TextMatrix(lnFila, 4) = rsDatRef!cReferido
        feReferencia.TextMatrix(lnFila, 5) = rsDatRef!cDNIRef
        rsDatRef.MoveNext
    Loop
    rsDatRef.Close
    Set rsDatRef = Nothing
    
    'Activos
    lnFila = 1
    Do While Not rsDatActivos.EOF
        feActivos.TextMatrix(lnFila, 3) = Format(rsDatActivos!nPE, "#,##0.00")
        feActivos.TextMatrix(lnFila, 4) = Format(rsDatActivos!nPP, "#,##0.00")
        feActivos.TextMatrix(lnFila, 5) = Format(rsDatActivos!nTotal, "#,##0.00")
        rsDatActivos.MoveNext
        lnFila = lnFila + 1
    Loop
    rsDatActivos.Close
    Set rsDatActivos = Nothing
    
    'Pasivo Corriente
    Call LimpiaFlex(fePasCorriente)
    Do While Not rsDatPasivos.EOF
        fePasCorriente.AdicionaFila
        lnFila = fePasCorriente.Row
        fePasCorriente.TextMatrix(lnFila, 1) = rsDatPasivos!cDesc
        fePasCorriente.TextMatrix(lnFila, 2) = Format(rsDatPasivos!nPE, "#,##0.00")
        fePasCorriente.TextMatrix(lnFila, 3) = Format(rsDatPasivos!nPP, "#,##0.00")
        fePasCorriente.TextMatrix(lnFila, 4) = Format(rsDatPasivos!nTotal, "#,##0.00")
        rsDatPasivos.MoveNext
    Loop
    rsDatPasivos.Close
    Set rsDatPasivos = Nothing
    
    
    Call LimpiaFlex(fePasNoCorriente)
    Do While Not rsDatPasivosNo.EOF
        fePasNoCorriente.AdicionaFila
        lnFila = fePasNoCorriente.Row
        fePasNoCorriente.TextMatrix(lnFila, 1) = rsDatPasivosNo!cDesc
        fePasNoCorriente.TextMatrix(lnFila, 2) = Format(rsDatPasivosNo!nPE, "#,##0.00")
        fePasNoCorriente.TextMatrix(lnFila, 3) = Format(rsDatPasivosNo!nPP, "#,##0.00")
        fePasNoCorriente.TextMatrix(lnFila, 4) = Format(rsDatPasivosNo!nTotal, "#,##0.00")
        rsDatPasivosNo.MoveNext
    Loop
    rsDatPasivosNo.Close
    Set rsDatPasivosNo = Nothing
    
    'Call LimpiaFlex(fePDT)
    lnFila = 1
    Do While Not rsDatPDTDet.EOF
        'fePDT.AdicionaFila
        fePDT.TextMatrix(lnFila, 3) = Format(rsDatPDTDet!nMontoMes1, "#,##0.00")
        fePDT.TextMatrix(lnFila, 4) = Format(rsDatPDTDet!nMontoMes2, "#,##0.00")
        fePDT.TextMatrix(lnFila, 5) = Format(rsDatPDTDet!nMontoMes3, "#,##0.00")
        fePDT.TextMatrix(lnFila, 6) = Format(rsDatPDTDet!nPromedio, "#,##0.00")
        rsDatPDTDet.MoveNext
        lnFila = lnFila + 1
    Loop
    rsDatPDTDet.Close
    Set rsDatPDTDet = Nothing

    fePDT.TextMatrix(0, 3) = DevolverMesDatos(CInt(rsDatPDT!nMes1))
    fePDT.TextMatrix(0, 4) = DevolverMesDatos(CInt(rsDatPDT!nMes2))
    fePDT.TextMatrix(0, 5) = DevolverMesDatos(CInt(rsDatPDT!nMes3))
    
    'Call LimpiaFlex(fePasNoCorriente)
    lnFila = 1
    Do While Not rsDatIngNeg.EOF
        'fePasNoCorriente.AdicionaFila
        'lnFila = feIngNeg.Row
        feIngNeg.TextMatrix(lnFila, 3) = Format(rsDatIngNeg!nProd1, "#,##0.00")
        feIngNeg.TextMatrix(lnFila, 4) = Format(rsDatIngNeg!nProd2, "#,##0.00")
        feIngNeg.TextMatrix(lnFila, 5) = Format(rsDatIngNeg!nProd3, "#,##0.00")
        feIngNeg.TextMatrix(lnFila, 6) = Format(rsDatIngNeg!nResultado, "#,##0.00")
        rsDatIngNeg.MoveNext
        lnFila = lnFila + 1
    Loop
    rsDatIngNeg.Close
    Set rsDatIngNeg = Nothing

    lblPatrimonioPE.Caption = Format(rsDatPatrimonio!nPE, "#,##0.00")
    lblPatrimonioPP.Caption = Format(rsDatPatrimonio!nPP, "#,##0.00")
    lblPatrimonioTotal.Caption = Format(rsDatPatrimonio!nTotal, "#,##0.00")
    
    lbTotPasPatrimonioPE.Caption = Format(rsDatPasPat!nPE, "#,##0.00")
    lbTotPasPatrimonioPP.Caption = Format(rsDatPasPat!nPP, "#,##0.00")
    lbTotPasPatrimonioTotal.Caption = Format(rsDatPasPat!nTotal, "#,##0.00")

    txtTCMIngNeg.Text = Format(nTasaIngNeg, "#,##0.00")
    txtTCMGastoNeg.Text = Format(nTasaGastoNeg, "#,##0.00")
    txtTCMOtrosIng.Text = Format(nTasaOtrosIng, "#,##0.00")
    txtTCMGastoFam.Text = Format(nTasaGastoFam, "#,##0.00")
    
    txtVerif.Text = rsCredEval!cVerif
    lnFila = 1
    Do While Not rsDatEstadoGP.EOF
        'fePasNoCorriente.AdicionaFila
        'lnFila = feEstadoGP.Row
        feEstadoGP.TextMatrix(lnFila, 1) = Format(rsDatEstadoGP!nIngresos, "#,##0.00")
        feEstadoGP.TextMatrix(lnFila, 2) = Format(rsDatEstadoGP!nEgresos, "#,##0.00")
        feEstadoGP.TextMatrix(lnFila, 3) = Format(rsDatEstadoGP!nMargen, "#,##0.00")
        feEstadoGP.TextMatrix(lnFila, 4) = Format(rsDatEstadoGP!nGastos, "#,##0.00")
        feEstadoGP.TextMatrix(lnFila, 5) = Format(rsDatEstadoGP!nIngresoNeto, "#,##0.00")
        rsDatEstadoGP.MoveNext
        lnFila = lnFila + 1
    Loop
    rsDatEstadoGP.Close
    Set rsDatEstadoGP = Nothing
    
    lnFila = 1
    Do While Not rsDatRatios.EOF
        'fePasNoCorriente.AdicionaFila
        'lnFila = feRatios.Row
        feRatios.TextMatrix(lnFila, 2) = Format(rsDatRatios!nEndCortoPlazo, "#,##0.00")
        feRatios.TextMatrix(lnFila, 3) = Format(rsDatRatios!nEndLargoPlazo, "#,##0.00")
        feRatios.TextMatrix(lnFila, 4) = Format(rsDatRatios!nLiqCorriente, "#,##0.00")
        feRatios.TextMatrix(lnFila, 5) = Format(rsDatRatios!nCapTrabajo, "#,##0.00")
        feRatios.TextMatrix(lnFila, 6) = Format(rsDatRatios!nRotInventario, "#,##0.00")
        rsDatRatios.MoveNext
        lnFila = lnFila + 1
    Loop
    rsDatRatios.Close
    Set rsDatRatios = Nothing
    
    
End Function


Private Sub CargarFlexEdit()
'ACTIVOS
feActivos.Clear
feActivos.FormaCabecera
feActivos.Rows = 2
    
For i = 1 To 16
    feActivos.AdicionaFila
    feActivos.TextMatrix(i, 1) = Choose(i, " ACTIVO CORRIENTE", "   Caja", "    Banco", "   Ctas. x Cobrar", "   Inventario", "      Mercaderia", "      Productos Terminados", "      Productos en Proceso", "      Materia Prima (insumos)", " ACTIVO FIJO", "   Muebles y Equipos de Ofic.", "   Vehículos", "   Maquinarias y Equipos", "   Edificios y Terrenos", "   Herramientas y Otros", "   ACTIVO TOTAL")
    feActivos.TextMatrix(i, 2) = Choose(i, "100", "110", "120", "130", "140", "141", "142", "143", "144", "200", "210", "220", "230", "240", "250", "300")

    Select Case i
        Case 1, 10, 16:   Call feActivos.BackColorRow(RGB(239, 235, 222), True)
        Case 5:   Call feActivos.BackColorRow(RGB(239, 235, 222))
    End Select
Next i

'DECLARACION PDT
sMes1 = DevolverMes(1, nAnio1, nMes3)
sMes2 = DevolverMes(2, nAnio2, nMes2)
sMes3 = DevolverMes(3, nAnio3, nMes1)
    
fePDT.Clear
fePDT.FormaCabecera
fePDT.Rows = 2
    
fePDT.TextMatrix(0, 3) = sMes3
fePDT.TextMatrix(0, 4) = sMes2
fePDT.TextMatrix(0, 5) = sMes1

fePDT.TextMatrix(0, 1) = "Mes/Detalle" & Space(8)
For i = 1 To 2
    fePDT.AdicionaFila
    fePDT.TextMatrix(i, 1) = Choose(i, "Compras" & Space(8), "Ventas" & Space(8))
    fePDT.TextMatrix(i, 2) = Choose(i, "1", "2")
Next i


'Ingresos del Negocio
feIngNeg.Clear
feIngNeg.FormaCabecera
feIngNeg.Rows = 2
For i = 1 To 11
    feIngNeg.AdicionaFila
    feIngNeg.TextMatrix(i, 1) = Choose(i, "Lunes" & Space(8), "Martes" & Space(6), "Miercoles" & Space(1), _
                                "Jueves" & Space(7), "Viernes" & Space(6), "Sábado" & Space(6), "Domingo" & Space(4), "Total" & Space(11), _
                                "% Costo" & Space(5), "% Particip." & Space(1), "% Real" & Space(8))
    feIngNeg.TextMatrix(i, 2) = Choose(i, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11")

    If i = 8 Or i = 11 Then
        Call feIngNeg.BackColorRow(RGB(239, 235, 222), True)
    End If
Next i

'RATIOS
feRatios.Clear
feRatios.FormaCabecera
feRatios.Rows = 2
For i = 1 To 3
    feRatios.AdicionaFila
    feRatios.TextMatrix(i, 1) = Choose(i, "Patrimonio Empresarial", "Patrimonio Personal", "Total")
    feRatios.TextMatrix(i, 2) = Format(0, "#,##0.00")
    feRatios.TextMatrix(i, 3) = Format(0, "#,##0.00")
    feRatios.TextMatrix(i, 4) = Format(0, "#,##0.00")
    feRatios.TextMatrix(i, 5) = Format(0, "#,##0.00")
    feRatios.TextMatrix(i, 6) = Format(0, "#,##0.00")
    feRatios.TextMatrix(i, 7) = Choose(i, "1", "2", "3")
Next i
'Estados GANACIAS Y PERDIDAS
feEstadoGP.AdicionaFila
End Sub
Private Function DevolverMes(ByVal pnMes As Integer, ByRef pnAnio As Integer, ByRef pnMesN As Integer) As String
Dim nIndMes As Integer
nIndMes = CInt(Mid(gdFecSis, 4, 2)) - pnMes
pnAnio = CInt(Mid(gdFecSis, 7, 4))
If nIndMes < 1 Then
    nIndMes = nIndMes + 12
    pnAnio = pnAnio - 1
End If

pnMesN = nIndMes

DevolverMes = Choose(nIndMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

Private Function DevolverMesDatos(ByVal pnMes As Integer) As String
DevolverMesDatos = Choose(pnMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function






Private Sub spnCalcCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCalcTEM.SetFocus
End If
End Sub

Private Sub spnCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMontoSol.SetFocus
End If
End Sub

Private Sub txtCalcMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    spnCalcCuotas.SetFocus
End If
End Sub

Private Sub txtCalcTEM_Change()
    If Trim(txtCalcTEM.Text) <> "." Then
        If CDbl(txtCalcTEM.Text) > 100 Then
            txtCalcTEM.Text = Replace(Mid(txtCalcTEM.Text, 1, Len(txtCalcTEM.Text) - 1), ",", "")
        End If
    End If
End Sub



Private Sub txtCalcTEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCalcular.SetFocus
End If
End Sub

'Private Sub txtComent_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cmdAgregarRef.SetFocus
'End If
'End Sub

Private Sub txtComent_LostFocus()
txtComent.Text = UCase(txtComent.Text)
End Sub


Private Sub txtCuotaPagar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    spnCuotas.SetFocus
End If
End Sub

Private Sub txtGiroNeg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCuotaPagar.SetFocus
End If
End Sub

Private Sub txtGiroNeg_LostFocus()
txtGiroNeg.Text = UCase(txtGiroNeg.Text)
End Sub



Private Sub txtMontoSol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    feActivos.SetFocus
    SSTOperaciones.Tab = 0
End If
End Sub

Private Sub txtTCMGastoFam_Change()
    If Trim(txtTCMGastoFam.Text) <> "." Then
        If CDbl(txtTCMGastoFam.Text) > 100 Then
            txtTCMGastoFam.Text = Replace(Mid(txtTCMGastoFam.Text, 1, Len(txtTCMGastoFam.Text) - 1), ",", "")
        End If
    End If
End Sub

Private Sub txtTCMGastoNeg_Change()
    If Trim(txtTCMGastoNeg.Text) <> "." Then
        If CDbl(txtTCMGastoNeg.Text) > 100 Then
            txtTCMGastoNeg.Text = Replace(Mid(txtTCMGastoNeg.Text, 1, Len(txtTCMGastoNeg.Text) - 1), ",", "")
        End If
    End If
End Sub

Private Sub txtTCMGastoNeg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAgregarOtrosIng.SetFocus
End If
End Sub

Private Sub txtTCMIngNeg_Change()
    If Trim(txtTCMIngNeg.Text) <> "." Then
        If CDbl(txtTCMIngNeg.Text) > 100 Then
            txtTCMIngNeg.Text = Replace(Mid(txtTCMIngNeg.Text, 1, Len(txtTCMIngNeg.Text) - 1), ",", "")
        End If
    End If
End Sub

Private Sub txtTCMIngNeg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAgregarGastoNeg.SetFocus
End If
End Sub

Private Sub txtTCMOtrosIng_Change()
    If Trim(txtTCMOtrosIng.Text) <> "." Then
        If CDbl(txtTCMOtrosIng.Text) > 100 Then
            txtTCMOtrosIng.Text = Replace(Mid(txtTCMOtrosIng.Text, 1, Len(txtTCMOtrosIng.Text) - 1), ",", "")
        End If
    End If
End Sub
Private Sub CargarConstantes()
Dim oCons As COMDConstantes.DCOMConstantes
Dim rsConstante As ADODB.Recordset

'Carga Tipo de Gasto del Negocio
Set oCons = New COMDConstantes.DCOMConstantes
Set rsConstante = oCons.RecuperaConstantes(7050)
feGastoNeg.CargaCombo rsConstante

End Sub
Private Sub VerPestana(ByVal i As Integer)
If i = 0 Then
    Me.SSTOperaciones.TabVisible(i) = False
    Me.SSTOperaciones.TabVisible(i + 1) = False
    Me.SSTOperaciones.TabVisible(i + 2) = False
    Me.SSTOperaciones.TabVisible(i) = True
    Me.SSTOperaciones.TabVisible(i + 1) = True
    Me.SSTOperaciones.TabVisible(i + 2) = True
ElseIf i = 1 Then
    Me.SSTOperaciones.TabVisible(i) = False
    Me.SSTOperaciones.TabVisible(i - 1) = False
    Me.SSTOperaciones.TabVisible(i + 1) = False
    Me.SSTOperaciones.TabVisible(i) = True
    Me.SSTOperaciones.TabVisible(i - 1) = True
    Me.SSTOperaciones.TabVisible(i + 1) = True
ElseIf i = 2 Then
    Me.SSTOperaciones.TabVisible(i) = False
    Me.SSTOperaciones.TabVisible(i - 2) = False
    Me.SSTOperaciones.TabVisible(i - 1) = False
    Me.SSTOperaciones.TabVisible(i) = True
    Me.SSTOperaciones.TabVisible(i - 2) = True
    Me.SSTOperaciones.TabVisible(i - 1) = True
End If
End Sub

Public Function validaDatos() As Boolean
If fnTipoPermiso = 3 Then
   
    If Trim(txtGiroNeg.Text) = "" Then
        MsgBox "Falta ingresar el Giro del Negocio", vbInformation, "Aviso"
        txtGiroNeg.SetFocus
        validaDatos = False
        Exit Function
    End If
    
    If txtCuotaPagar.value = 0 Then
        MsgBox "Falta ingresar la Probable cuota a pagar", vbInformation, "Aviso"
        txtCuotaPagar.SetFocus
        validaDatos = False
        Exit Function
    End If
    If spnCuotas.valor = 0 Then
        MsgBox "Falta ingresar el Nro de cuotas", vbInformation, "Aviso"
        spnCuotas.SetFocus
        validaDatos = False
        Exit Function
    End If
 
    If cboMontoSol.ListIndex = -1 Then
        MsgBox "Falta seleccionar la moneda", vbInformation, "Aviso"
        cboMontoSol.SetFocus
        validaDatos = False
        Exit Function
    End If
    
    If txtMontoSol.value = 0 Then
        MsgBox "Falta ingresar el Monto solicitado", vbInformation, "Aviso"
        txtMontoSol.SetFocus
        validaDatos = False
        Exit Function
    End If
 
  
    If ValidaGridGastoNeg(feGastoNeg) = False Then
        MsgBox "Faltan datos en la lista de Gastos del Negocio", vbInformation, "Aviso"
        SSTOperaciones.Tab = 1
        validaDatos = False
        Exit Function
    End If
    
    If ValidaGrillas(feGastoFam) = False Then
        MsgBox "Faltan datos en la lista de Gastos Familiares", vbInformation, "Aviso"
        SSTOperaciones.Tab = 1
        validaDatos = False
        Exit Function
    End If
    
    If ValidaGrillas(feOtrosIng) = False Then
        MsgBox "Faltan datos en la lista de Otros Ingresos", vbInformation, "Aviso"
        SSTOperaciones.Tab = 1
        validaDatos = False
        Exit Function
    End If
        
    If nPersoneria = 1 Then
        For i = 1 To 16
            If i = 1 Or i = 10 Or i = 16 Then
                For j = 3 To 5
                    If Trim(feActivos.TextMatrix(i, j)) = "" Then
                        MsgBox "Faltan datos en la lista de Activos", vbInformation, "Aviso"
                        SSTOperaciones.Tab = 0
                        validaDatos = False
                        Exit Function
                    End If
                Next j
            End If
        Next i
    Else
        For i = 1 To 16
            If i = 1 Or i = 10 Or i = 16 Then
                For j = 3 To 5
                    If i <> 4 Then
                        If Trim(feActivos.TextMatrix(i, j)) = "" Then
                            MsgBox "Faltan datos en la lista de Activos.", vbInformation, "Aviso"
                            SSTOperaciones.Tab = 0
                            validaDatos = False
                            Exit Function
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    
 
    If ValidaGridPasivos(fePasCorriente) = False Then
        MsgBox "Faltan datos en la lista de Pasivos Corriente", vbInformation, "Aviso"
        SSTOperaciones.Tab = 0
        validaDatos = False
        Exit Function
    End If
    
    If ValidaGridPasivos(fePasNoCorriente) = False Then
        MsgBox "Faltan datos en la lista de Pasivos No Corriente", vbInformation, "Aviso"
        SSTOperaciones.Tab = 0
        validaDatos = False
        Exit Function
    End If
    
    For i = 1 To 2
        For j = 3 To 6
            If Trim(fePDT.TextMatrix(i, j)) = "" Then
                MsgBox "Faltan datos en la Declaración PDT.", vbInformation, "Aviso"
                SSTOperaciones.Tab = 0
                validaDatos = False
                Exit Function
            End If
        Next j
    Next i
    
    For i = 1 To 11
        If i = 8 Or i = 11 Then
            For j = 3 To 6
                If j = 3 Or j = 6 Then
                    If Trim(feIngNeg.TextMatrix(i, j)) = "" Then
                        MsgBox "Faltan datos en la lista de Ingresos del Negocio", vbInformation, "Aviso"
                        SSTOperaciones.Tab = 1
                        feIngNeg.SetFocus
                        validaDatos = False
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i
    
    Dim nMontonFila As Double
    Dim nMontonTotal As Double
    Dim DescFila As String
    Dim DesColumna As String
    nMontonFila = 0
    nMontonTotal = 0
    
    For i = 9 To 10
    DescFila = Choose(i - 8, "% Costo", "% Particip.")
    nMontonTotal = 0
        For j = 3 To 5
        DesColumna = Choose(j - 2, "Prod1", "Prod2", "Prod3")
            nMontonFila = 0
            nMontonFila = CDbl(IIf(Trim(feIngNeg.TextMatrix(i, j)) = "", 0, Trim(feIngNeg.TextMatrix(i, j))))
            nMontonTotal = nMontonTotal + nMontonFila
            If nMontonFila > 100 Then
                MsgBox "Solo se permiten datos del 1 al 100 en el " & DescFila & " del " & DesColumna & ".", vbInformation, "Aviso"
                SSTOperaciones.Tab = 1
                feIngNeg.SetFocus
                validaDatos = False
                Exit Function
            End If
        Next j
        
        If nMontonTotal <> 100 And i = 10 Then
            MsgBox "La Sumatoria del Porcetanje de Participación debe ser 100 ", vbInformation, "Aviso"
            SSTOperaciones.Tab = 1
            feIngNeg.SetFocus
            validaDatos = False
            Exit Function
        End If
    Next i
    
    If Trim(lblMontoMax.Caption) = "0.00" Then
        MsgBox "Faltan datos para el calculo del Monto maximo del credito", vbInformation, "Aviso"
        SSTOperaciones.Tab = 2
        validaDatos = False
        Exit Function
    End If
    
    If Trim(lblCuotaEstima.Caption) = "0.00" Then
        MsgBox "Faltan datos para el calculo de la cuota estimada", vbInformation, "Aviso"
        SSTOperaciones.Tab = 2
        validaDatos = False
        Exit Function
    End If
    
    If Round(CDbl(lblMontoMax.Caption), 2) < Round(CDbl(txtCalcMonto.Text), 2) Then
        MsgBox "El Monto Máximo del Credito es menor al ingresado en el calculo", vbInformation, "Aviso"
        txtCalcMonto.SetFocus
        SSTOperaciones.Tab = 2
        validaDatos = False
        Exit Function
    End If
    If Round(CDbl(lblCuotaEstima.Caption), 2) > Round(CDbl(txtCuotaPagar.Text), 2) Then
        MsgBox "La Couta Estimada a Pagar es mayor a la Probable Cuota por Pagar", vbInformation, "Aviso"
        txtCuotaPagar.SetFocus
        SSTOperaciones.Tab = 2
        validaDatos = False
        Exit Function
    End If
    
    If Trim(txtComent.Text) = "" Then
        MsgBox "Faltan ingresar el comentario", vbInformation, "Aviso"
        txtComent.SetFocus
        SSTOperaciones.Tab = 2
        validaDatos = False
        Exit Function
    End If
    
    If ValidaDatosReferencia = False Then
        validaDatos = False
        Exit Function
    End If
    
    
ElseIf fnTipoPermiso = 2 Then
        If Trim(txtVerif.Text) = "" Then
            MsgBox "Favor de Ingresar la Validación Respectiva.", vbInformation, "Aviso"
            txtVerif.SetFocus
            SSTOperaciones.Tab = 2
            validaDatos = False
            Exit Function
        End If
End If
    

validaDatos = True
End Function


Private Sub CalcularActivoPatrimonio()
lblPatrimonioPE.Caption = Format(fnActivoPE - fnPasivoPE, "#0.00")
lblPatrimonioPP.Caption = Format(fnActivoPP - fnPasivoPP, "#0.00")
lblPatrimonioTotal.Caption = Format(fnActivoTOTAL - fnPasivoTOTAL, "#0.00")

lbTotPasPatrimonioPE.Caption = Format(fnPasivoPE + CDbl(IIf(Trim(lblPatrimonioPE.Caption) = "", 0, lblPatrimonioPE.Caption)), "#0.00")
lbTotPasPatrimonioPP.Caption = Format(fnPasivoPP + CDbl(IIf(Trim(lblPatrimonioPP.Caption) = "", 0, lblPatrimonioPP.Caption)), "#0.00")
lbTotPasPatrimonioTotal.Caption = Format(fnPasivoTOTAL + CDbl(IIf(Trim(lblPatrimonioTotal.Caption) = "", 0, lblPatrimonioTotal.Caption)), "#0.00")

Call CalcularRatios
Call cmdCalcular_Click
End Sub

Private Sub CalcularGanaciasPerdidas()
feEstadoGP.lbEditarFlex = True
Dim sMargen As String
Dim sIngNeto As String
Dim nEgresos As Double
Dim nGasto As Double
nGasto = SumarCampo(feGastoNeg, 3) + SumarCampo(feGastoFam, 2)
nEgresos = CDbl(IIf(Trim(feIngNeg.TextMatrix(8, 6)) = "", 0, feIngNeg.TextMatrix(8, 6))) * ((CDbl(IIf(Trim(feIngNeg.TextMatrix(11, 6)) = "", 0, feIngNeg.TextMatrix(11, 6)))) / 100)
feEstadoGP.TextMatrix(1, 1) = feIngNeg.TextMatrix(8, 6)
feEstadoGP.TextMatrix(1, 2) = Format(nEgresos, "#0.00")
sMargen = Format(CDbl(IIf(Trim(feIngNeg.TextMatrix(8, 6)) = "", 0, feIngNeg.TextMatrix(8, 6))) - nEgresos, "#0.00")
sIngNeto = Format(CDbl(IIf(sMargen = "", 0, sMargen)) - nGasto, "#0.00")
feEstadoGP.TextMatrix(1, 3) = sMargen
feEstadoGP.TextMatrix(1, 4) = Format(nGasto, "#0.00")
feEstadoGP.TextMatrix(1, 5) = sIngNeto
feEstadoGP.lbEditarFlex = False
Call cmdCalcular_Click
End Sub

Private Sub CalcularRatios()
feRatios.TextMatrix(1, 2) = Format((SumarCampo(fePasCorriente, 2) / CDbl(IIf(Trim(lblPatrimonioPE.Caption) = 0, 1, lblPatrimonioPE.Caption))) * 100, "#0.00")
feRatios.TextMatrix(2, 2) = Format((SumarCampo(fePasCorriente, 3) / CDbl(IIf(Trim(lblPatrimonioPP.Caption) = 0, 1, lblPatrimonioPP.Caption))) * 100, "#0.00")
feRatios.TextMatrix(3, 2) = Format((SumarCampo(fePasCorriente, 4) / CDbl(IIf(Trim(lblPatrimonioTotal.Caption) = 0, 1, lblPatrimonioTotal.Caption))) * 100, "#0.00")

feRatios.TextMatrix(1, 3) = Format((SumarCampo(fePasNoCorriente, 2) / CDbl(IIf(Trim(lblPatrimonioPE.Caption) = 0, 1, lblPatrimonioPE.Caption))) * 100, "#0.00")
feRatios.TextMatrix(2, 3) = Format((SumarCampo(fePasNoCorriente, 3) / CDbl(IIf(Trim(lblPatrimonioPP.Caption) = 0, 1, lblPatrimonioPP.Caption))) * 100, "#0.00")
feRatios.TextMatrix(3, 3) = Format((SumarCampo(fePasNoCorriente, 4) / CDbl(IIf(Trim(lblPatrimonioTotal.Caption) = 0, 1, lblPatrimonioTotal.Caption))) * 100, "#0.00")


feRatios.TextMatrix(1, 4) = Format(CDbl(IIf(Trim(feActivos.TextMatrix(1, 3)) = "", 0, feActivos.TextMatrix(1, 3))) / IIf(SumarCampo(fePasCorriente, 2) = 0, 1, SumarCampo(fePasCorriente, 2)), "#0.00")
feRatios.TextMatrix(2, 4) = Format(CDbl(IIf(Trim(feActivos.TextMatrix(1, 4)) = "", 0, feActivos.TextMatrix(1, 4))) / IIf(SumarCampo(fePasCorriente, 3) = 0, 1, SumarCampo(fePasCorriente, 3)), "#0.00")
feRatios.TextMatrix(3, 4) = Format(CDbl(IIf(Trim(feActivos.TextMatrix(1, 5)) = "", 0, feActivos.TextMatrix(1, 5))) / IIf(SumarCampo(fePasCorriente, 4) = 0, 1, SumarCampo(fePasCorriente, 4)), "#0.00")


feRatios.TextMatrix(1, 5) = Format(CDbl(IIf(Trim(feActivos.TextMatrix(1, 3)) = "", 0, feActivos.TextMatrix(1, 3))) - IIf(SumarCampo(fePasCorriente, 2) = 0, 0, SumarCampo(fePasCorriente, 2)), "#0.00")
feRatios.TextMatrix(2, 5) = Format(CDbl(IIf(Trim(feActivos.TextMatrix(1, 4)) = "", 0, feActivos.TextMatrix(1, 4))) - IIf(SumarCampo(fePasCorriente, 3) = 0, 0, SumarCampo(fePasCorriente, 3)), "#0.00")
feRatios.TextMatrix(3, 5) = Format(CDbl(IIf(Trim(feActivos.TextMatrix(1, 5)) = "", 0, feActivos.TextMatrix(1, 5))) - IIf(SumarCampo(fePasCorriente, 4) = 0, 0, SumarCampo(fePasCorriente, 4)), "#0.00")

feRatios.TextMatrix(1, 6) = Format((CDbl(IIf(Trim(feEstadoGP.TextMatrix(1, 2)) = "", 0, feEstadoGP.TextMatrix(1, 2))) / CDbl(IIf(Trim(feActivos.TextMatrix(5, 3)) = "0.00" Or Trim(feActivos.TextMatrix(5, 3)) = "", 1, feActivos.TextMatrix(5, 3)))) * 100, "#0.00")
feRatios.TextMatrix(2, 6) = Format((CDbl(IIf(Trim(feEstadoGP.TextMatrix(1, 2)) = "", 0, feEstadoGP.TextMatrix(1, 2))) / CDbl(IIf(Trim(feActivos.TextMatrix(5, 4)) = "0.00" Or Trim(feActivos.TextMatrix(5, 3)) = "", 1, feActivos.TextMatrix(5, 4)))) * 100, "#0.00")
feRatios.TextMatrix(3, 6) = Format((CDbl(IIf(Trim(feEstadoGP.TextMatrix(1, 2)) = "", 0, feEstadoGP.TextMatrix(1, 2))) / CDbl(IIf(Trim(feActivos.TextMatrix(5, 5)) = "0.00" Or Trim(feActivos.TextMatrix(5, 3)) = "", 1, feActivos.TextMatrix(5, 5)))) * 100, "#0.00")

End Sub

Private Sub CargaControlesInicio()
cboMontoSol.ListIndex = IndiceListaCombo(cboMontoSol, Mid(sCtaCod, 9, 1))
Call CargarFlexEdit
Call CargarConstantes
Call VerPestana(0)
End Sub

Private Sub txtTCMOtrosIng_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAgregarGastoFam.SetFocus
End If
End Sub

Private Sub txtVerif_LostFocus()
txtVerif.Text = UCase(txtVerif.Text)
End Sub

Public Function ValidaGrillas(ByVal Flex As FlexEdit) As Boolean
    Dim i As Integer
    ValidaGrillas = False
    For i = 1 To Flex.Rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 2)) = "" Then
                ValidaGrillas = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrillas = True
End Function
Public Function ValidaGridPasivos(ByVal Flex As FlexEdit) As Boolean
    Dim i As Integer
    
    For i = 1 To Flex.Rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            If nPersoneria = 1 Then
                If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 2)) = "" Or Trim(Flex.TextMatrix(i, 3)) = "" Or Trim(Flex.TextMatrix(i, 4)) = "" Then
                    ValidaGridPasivos = False
                    Exit Function
                End If
            Else
                If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 2)) = "" Or Trim(Flex.TextMatrix(i, 4)) = "" Then
                    ValidaGridPasivos = False
                    Exit Function
                End If
            End If
        End If
    Next i
    ValidaGridPasivos = True
End Function





Public Function ValidaGridGastoNeg(ByVal Flex As FlexEdit) As Boolean
    Dim i As Integer
    
    For i = 1 To Flex.Rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 2)) = "" Or Trim(Flex.TextMatrix(i, 3)) = "" Then
                ValidaGridGastoNeg = False
                Exit Function
            End If
        End If
    Next i
    ValidaGridGastoNeg = True
End Function
Public Function ValidaDatosReferencia() As Boolean
    Dim i As Integer, j As Integer
    ValidaDatosReferencia = False
    
    If feReferencia.Rows - 1 < 2 Then
        MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
        cmdAgregarRef.SetFocus
        ValidaDatosReferencia = False
        Exit Function
    End If
    
    'Verfica Tipo de Valores del DNI
    For i = 1 To feReferencia.Rows - 1
        If Trim(feReferencia.TextMatrix(i, 1)) <> "" Then
            For j = 1 To Len(Trim(feReferencia.TextMatrix(i, 2)))
                If (Mid(feReferencia.TextMatrix(i, 2), j, 1) < "0" Or Mid(feReferencia.TextMatrix(i, 2), j, 1) > "9") Then
                   MsgBox "Uno de los Digitos del primer DNI no es un Numero", vbInformation, "Aviso"
                   feReferencia.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next j
        End If
    Next i
    
    'Verfica Longitud del DNI
    For i = 1 To feReferencia.Rows - 1
        If Trim(feReferencia.TextMatrix(i, 1)) <> "" Then
            If Len(Trim(feReferencia.TextMatrix(i, 2))) <> gnNroDigitosDNI Then
                MsgBox "Primer DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                feReferencia.SetFocus
                ValidaDatosReferencia = False
                Exit Function
            End If
        End If
    Next i
    
    'Verfica Tipo de Valores del Telefono
    For i = 1 To feReferencia.Rows - 1
        If Trim(feReferencia.TextMatrix(i, 1)) <> "" Then
            For j = 1 To Len(Trim(feReferencia.TextMatrix(i, 3)))
                If (Mid(feReferencia.TextMatrix(i, 3), j, 1) < "0" Or Mid(feReferencia.TextMatrix(i, 3), j, 1) > "9") Then
                   MsgBox "Uno de los Digitos del teléfono no es un Numero", vbInformation, "Aviso"
                   feReferencia.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next j
        End If
    Next i
    
    'Verfica Tipo de Valores del DNI 2
    For i = 1 To feReferencia.Rows - 1
        If Trim(feReferencia.TextMatrix(i, 1)) <> "" Then
            For j = 1 To Len(Trim(feReferencia.TextMatrix(i, 5)))
                If (Mid(feReferencia.TextMatrix(i, 5), j, 1) < "0" Or Mid(feReferencia.TextMatrix(i, 5), j, 1) > "9") Then
                   MsgBox "Uno de los Digitos del segundo DNI no es un Numero", vbInformation, "Aviso"
                   feReferencia.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next j
        End If
    Next i
    
    'Verfica Longitud del DNI 2
    For i = 1 To feReferencia.Rows - 1
        If Trim(feReferencia.TextMatrix(i, 1)) <> "" Then
            If Len(Trim(feReferencia.TextMatrix(i, 5))) <> gnNroDigitosDNI Then
                MsgBox "Segundo DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                feReferencia.SetFocus
                ValidaDatosReferencia = False
                Exit Function
            End If
        End If
    Next i
    
    'Verfica ambos DNI que no sean iguales
    For i = 1 To feReferencia.Rows - 1
        If Trim(feReferencia.TextMatrix(i, 1)) <> "" Then
            If Trim(feReferencia.TextMatrix(i, 2)) = Trim(feReferencia.TextMatrix(i, 5)) Then
                MsgBox "Los DNI de la fila " & i & " son iguales", vbInformation, "Aviso"
                feReferencia.SetFocus
                ValidaDatosReferencia = False
                Exit Function
            End If
        End If
    Next i
    ValidaDatosReferencia = True
End Function


Private Sub GeneraExcelFormato()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim FilaPertenece As Integer
    Dim pnCondLocalCol As Integer
    Dim nTEA As Double
    Dim nCuotaUNM As Double, nCuotaExdFam As Double
    Dim i As Integer
    Dim IniTablas As Integer, IniTablaOtroIng As Integer, FinTablas As Integer
    Dim CeldaVacia1 As Integer, CeldaVacia2 As Integer
    Dim CeldaVacia3 As Integer, CeldaVacia4 As Integer
    Dim Celda As String
    
    On Error GoTo ErrorGeneraExcelFormato
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    lsNomHoja = "FORMATO3-1"
    lsFile = "CredEvalFormato3"
    
    lsArchivo = "\spooler\" & "Evaluacion_" & sCtaCod & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    'HOJA FORMATO 3-1
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    fsCliente = PstaNombre(fsCliente, True)
    nTEA = ((1 + (CDbl(txtCalcTEM.Text) / 100)) ^ 12) - 1

    
    xlHoja1.Cells(2, 2) = "FORMATO 3. EVALUACIÓN DE CRÉDITOS HASTA " & IIf(Mid(sCtaCod, 9, 1) = "1", Format(lnMax, "#,##0.00"), Format(lnMaxDol, "#,##0.00"))
    xlHoja1.Cells(4, 3) = fsCliente
    xlHoja1.Cells(4, 16) = fsUserAnalista
    xlHoja1.Cells(8, 3) = sCtaCod
    xlHoja1.Cells(8, 11) = txtGiroNeg.Text
    xlHoja1.Cells(10, 5) = txtCuotaPagar.Text
    xlHoja1.Cells(10, 11) = spnCuotas.valor
    xlHoja1.Cells(10, 16) = txtMontoSol.Text

    'ACTIVOS
    IniTablas = 15
    For i = 1 To feActivos.Rows - 1
        xlHoja1.Cells(IniTablas + i, 6) = feActivos.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 7) = feActivos.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 8) = feActivos.TextMatrix(i, 5)
        If (IniTablas + i = 24) Or (IniTablas + i) = 31 Then
            IniTablas = IniTablas + 1
        End If
    Next i


    'PASIVOS CORRIENTE Y NO CORRIENTE
    IniTablas = 37
    For i = 1 To fePasCorriente.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = fePasCorriente.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 6) = fePasCorriente.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 7) = fePasCorriente.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 8) = fePasCorriente.TextMatrix(i, 4)
    Next i
    CeldaVacia1 = IniTablas + i
    FinTablas = 62
    
    xlHoja1.Cells(37, 6) = Format(SumarCampo(fePasCorriente, 2), "#,##0.00")
    xlHoja1.Cells(37, 7) = Format(SumarCampo(fePasCorriente, 3), "#,##0.00")
    xlHoja1.Cells(37, 8) = Format(SumarCampo(fePasCorriente, 4), "#,##0.00")
       
    IniTablas = 37
    For i = 1 To fePasNoCorriente.Rows - 1
        xlHoja1.Cells(IniTablas + i, 10) = fePasNoCorriente.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 14) = fePasNoCorriente.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 15) = fePasNoCorriente.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 16) = fePasNoCorriente.TextMatrix(i, 4)
    Next i
    CeldaVacia2 = IniTablas + i
    FinTablas = 62
    
    xlHoja1.Cells(37, 14) = Format(SumarCampo(fePasNoCorriente, 2), "#,##0.00")
    xlHoja1.Cells(37, 15) = Format(SumarCampo(fePasNoCorriente, 3), "#,##0.00")
    xlHoja1.Cells(37, 16) = Format(SumarCampo(fePasNoCorriente, 4), "#,##0.00")
    
    If IIf(CeldaVacia1 > CeldaVacia2, CeldaVacia1, CeldaVacia2) < FinTablas Then
        For i = IIf(CeldaVacia1 > CeldaVacia2, CeldaVacia1, CeldaVacia2) To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    
    
    xlHoja1.Cells(64, 6) = Format(SumarCampo(fePasCorriente, 2) + SumarCampo(fePasNoCorriente, 2), "#,##0.00")
    xlHoja1.Cells(64, 7) = Format(SumarCampo(fePasCorriente, 3) + SumarCampo(fePasNoCorriente, 3), "#,##0.00")
    xlHoja1.Cells(64, 8) = Format(SumarCampo(fePasCorriente, 4) + SumarCampo(fePasNoCorriente, 4), "#,##0.00")
    
    xlHoja1.Cells(66, 6) = Format(lblPatrimonioPE.Caption, "#,##0.00")
    xlHoja1.Cells(66, 7) = Format(lblPatrimonioPP.Caption, "#,##0.00")
    xlHoja1.Cells(66, 8) = Format(lblPatrimonioTotal.Caption, "#,##0.00")
    
    xlHoja1.Cells(68, 6) = Format(lbTotPasPatrimonioPE.Caption, "#,##0.00")
    xlHoja1.Cells(68, 7) = Format(lbTotPasPatrimonioPP.Caption, "#,##0.00")
    xlHoja1.Cells(68, 8) = Format(lbTotPasPatrimonioTotal.Caption, "#,##0.00")
    
    
    'Ingreso del negocio y Gasto Negocio
    IniTablas = 73
    For i = 1 To feIngNeg.Rows - 1
        For j = 3 To 6
            xlHoja1.Cells(IniTablas + i, 2 + j) = feIngNeg.TextMatrix(i, j)
        Next j
    Next i
    
    
    
    IniTablas = 73
    For i = 1 To feGastoNeg.Rows - 1
        xlHoja1.Cells(IniTablas + i, 10) = feGastoNeg.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 14) = Trim(Left(feGastoNeg.TextMatrix(i, 2), 4))
        xlHoja1.Cells(IniTablas + i, 15) = feGastoNeg.TextMatrix(i, 3)
    Next i
    CeldaVacia1 = IniTablas + i
    FinTablas = 98
    xlHoja1.Cells(99, 15) = SumarCampo(feGastoNeg, 3)
    
    If IIf(CeldaVacia1 > 84, CeldaVacia1, 84) < FinTablas Then
        For i = IIf(CeldaVacia1 > 84, CeldaVacia1, 84) To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    'otros ingresos
    IniTablas = 102
    For i = 1 To feOtrosIng.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = feOtrosIng.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 7) = feOtrosIng.TextMatrix(i, 2)
    Next i
    CeldaVacia1 = IniTablas + i
    FinTablas = 127
    xlHoja1.Cells(128, 7) = SumarCampo(feOtrosIng, 2)
    
    'gasto familiares
    IniTablas = 102
    For i = 1 To feGastoFam.Rows - 1
        xlHoja1.Cells(IniTablas + i, 10) = feGastoFam.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 15) = feGastoFam.TextMatrix(i, 2)
    Next i
    CeldaVacia2 = IniTablas + i
    FinTablas = 127
    xlHoja1.Cells(128, 15) = SumarCampo(feGastoFam, 2)
    If IIf(CeldaVacia1 > CeldaVacia2, CeldaVacia1, CeldaVacia2) < FinTablas Then
        For i = IIf(CeldaVacia1 > CeldaVacia2, CeldaVacia1, CeldaVacia2) To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    'PDT
    xlHoja1.Cells(131, 5) = sMes3
    xlHoja1.Cells(131, 6) = sMes2
    xlHoja1.Cells(131, 7) = sMes1
    IniTablas = 131
    For i = 1 To fePDT.Rows - 1
        For j = 3 To 6
            xlHoja1.Cells(IniTablas + i, 2 + j) = fePDT.TextMatrix(i, j)
        Next j
    Next i
    CeldaVacia2 = IniTablas + i
    
    xlHoja1.Cells(134, 5) = CDbl(IIf(Trim(xlHoja1.Cells(132, 5)) = "", 0, xlHoja1.Cells(133, 5))) + CDbl(IIf(Trim(xlHoja1.Cells(133, 5)) = "", 0, xlHoja1.Cells(132, 5)))
    xlHoja1.Cells(134, 6) = CDbl(IIf(Trim(xlHoja1.Cells(132, 6)) = "", 0, xlHoja1.Cells(133, 6))) + CDbl(IIf(Trim(xlHoja1.Cells(133, 6)) = "", 0, xlHoja1.Cells(132, 6)))
    xlHoja1.Cells(134, 7) = CDbl(IIf(Trim(xlHoja1.Cells(132, 7)) = "", 0, xlHoja1.Cells(133, 7))) + CDbl(IIf(Trim(xlHoja1.Cells(133, 7)) = "", 0, xlHoja1.Cells(132, 7)))
    xlHoja1.Cells(134, 8) = CDbl(IIf(Trim(xlHoja1.Cells(132, 8)) = "", 0, xlHoja1.Cells(133, 8))) + CDbl(IIf(Trim(xlHoja1.Cells(133, 8)) = "", 0, xlHoja1.Cells(132, 8)))
    
    xlHoja1.Cells(138, 7) = feEstadoGP.TextMatrix(1, 1)
    xlHoja1.Cells(139, 7) = feEstadoGP.TextMatrix(1, 2)
    xlHoja1.Cells(140, 7) = feEstadoGP.TextMatrix(1, 3)
    xlHoja1.Cells(141, 7) = feEstadoGP.TextMatrix(1, 4)
    xlHoja1.Cells(142, 7) = feEstadoGP.TextMatrix(1, 5)
    
    'RATIOS
    IniTablas = 137
    For i = 1 To feRatios.Rows - 1
        For j = 2 To 6
            xlHoja1.Cells(IniTablas + (j - 1), i + 13) = feRatios.TextMatrix(i, j)
        Next j
    Next i

    'HOJA FORMATO 3-2
    lsNomHoja = "FORMATO3-2"
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    xlHoja1.Cells(6, 2) = gdFecSis
    xlHoja1.Cells(6, 4) = IIf(Mid(sCtaCod, 9, 1) = "1", "SOLES", "DOLARES")
    xlHoja1.Cells(6, 5) = txtCalcMonto.Text
    xlHoja1.Cells(6, 10) = spnCalcCuotas.valor
    xlHoja1.Cells(6, 12) = CStr(CDbl(txtCalcTEM.Text)) ' / 100)
    xlHoja1.Cells(6, 14) = nTEA
    xlHoja1.Cells(6, 16) = Format(lblCuotaEstima.Caption, "#,##0.00")
    

    xlHoja1.Cells(9, 4) = Format(lblEndeudTotal.Caption, "#,##0.00")
    xlHoja1.Cells(9, 10) = Format(lblCapPagoEmp.Caption, "#,##0.00")
    xlHoja1.Cells(9, 17) = Format(lblCapPagoTotal.Caption, "#,##0.00")
    
    xlHoja1.Cells(12, 2) = Trim(txtComent.Text)
    xlHoja1.Cells(49, 2) = Trim(txtVerif.Text)
       
    IniTablas = 16
    For i = 1 To feReferencia.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = i
        xlHoja1.Cells(IniTablas + i, 3) = feReferencia.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 7) = feReferencia.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 9) = feReferencia.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 11) = feReferencia.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 16) = feReferencia.TextMatrix(i, 5)
    Next i
    CeldaVacia4 = IniTablas + i
    FinTablas = 41
    
    If CeldaVacia4 < FinTablas Then
        For i = CeldaVacia4 To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    
    
    'REGRESAR A LA PRIMERA HOJA FORMATO 3-1
    lsNomHoja = "FORMATO3-1"
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
   
    
    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
    
    Exit Sub
ErrorGeneraExcelFormato:
    MsgBox err.Description, vbInformation, "Error!!"
End Sub

