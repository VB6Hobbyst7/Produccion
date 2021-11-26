VERSION 5.00
Begin VB.Form frmConfigHojaTrabajoFE 
   Caption         =   "Configuración de Flujo de Efectivo"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigHojaTrabajoFE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Ajustes"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   13095
      Begin VB.CommandButton cmdBajarAjuste 
         Caption         =   "Bajar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   44
         Top             =   1400
         Width           =   975
      End
      Begin VB.CommandButton cmdSubirAjuste 
         Caption         =   "Subir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   43
         Top             =   1000
         Width           =   975
      End
      Begin VB.CommandButton cmdQuitarAjuste 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   42
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarAjuste 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   41
         Top             =   2200
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregarAjuste 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   40
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardarAjuste 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   38
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         HideSelection   =   0   'False
         Left            =   470
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "frmConfigHojaTrabajoFE.frx":030A
         Top             =   250
         Width           =   3030
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "frmConfigHojaTrabajoFE.frx":0326
         Top             =   250
         Width           =   360
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3890
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   35
         Text            =   "Ajuste y Clasifi."
         Top             =   250
         Width           =   1940
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5810
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   34
         Text            =   "Act. Operación"
         Top             =   250
         Width           =   1940
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   7720
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   33
         Text            =   "Act. Inversión"
         Top             =   250
         Width           =   1940
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   9650
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   32
         Text            =   "Act. Financiamiento"
         Top             =   250
         Width           =   1940
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         HideSelection   =   0   'False
         Left            =   3480
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "frmConfigHojaTrabajoFE.frx":032E
         Top             =   250
         Width           =   420
      End
      Begin Sicmact.FlexEdit fgAjuste 
         Height          =   1935
         Left            =   120
         TabIndex        =   39
         Top             =   550
         Width           =   11770
         _ExtentX        =   20770
         _ExtentY        =   3413
         Cols0           =   13
         ScrollBars      =   2
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Id-Descripcion-Nivel-Debe-Haber-Debe-Haber-Debe-Haber-Debe-Haber-Orden"
         EncabezadosAnchos=   "350-0-3010-400-955-955-955-955-955-955-955-955-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-8-9-10-11-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-3-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pasivo"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   13095
      Begin VB.CommandButton cmdBajarPasivo 
         Caption         =   "Bajar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   29
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdSubirPasivo 
         Caption         =   "Subir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   28
         Top             =   1100
         Width           =   975
      End
      Begin VB.CommandButton cmdQuitarPasivo 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   27
         Top             =   700
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarPasivo 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   26
         Top             =   2300
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregarPa 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   25
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardarPasivo 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   24
         Top             =   1900
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         HideSelection   =   0   'False
         Left            =   470
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "frmConfigHojaTrabajoFE.frx":0336
         Top             =   420
         Width           =   1510
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "frmConfigHojaTrabajoFE.frx":0352
         Top             =   420
         Width           =   360
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1970
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         Text            =   "Formula Cuentas"
         Top             =   420
         Width           =   1940
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3890
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   19
         Text            =   "Ajuste y Clasifi."
         Top             =   420
         Width           =   1940
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5810
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         Text            =   "Act. Operación"
         Top             =   420
         Width           =   1940
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   7720
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Text            =   "Act. Inversión"
         Top             =   420
         Width           =   1940
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   9650
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Text            =   "Act. Financiamiento"
         Top             =   420
         Width           =   1940
      End
      Begin Sicmact.FlexEdit fgPasivo 
         Height          =   1935
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   11770
         _ExtentX        =   20770
         _ExtentY        =   3413
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Id-Descripcion-<2013->2013-Debe-Haber-Debe-Haber-Debe-Haber-Debe-Haber-Orden"
         EncabezadosAnchos=   "350-0-1500-955-955-955-955-955-955-955-955-955-955-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-8-9-10-11-12-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Activo"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      Begin VB.CommandButton cmdAgregarActivo 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   13
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardarActivo 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdQuitarActivo 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdSubirActivo 
         Caption         =   "Subir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   10
         Top             =   1000
         Width           =   975
      End
      Begin VB.CommandButton cmdBajarActivo 
         Caption         =   "Bajar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   9
         Top             =   1400
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarActivo 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   8
         Top             =   2200
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         HideSelection   =   0   'False
         Left            =   470
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "frmConfigHojaTrabajoFE.frx":035A
         Top             =   300
         Width           =   1510
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmConfigHojaTrabajoFE.frx":0376
         Top             =   300
         Width           =   360
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1970
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Text            =   "Formula Cuentas"
         Top             =   300
         Width           =   1940
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3890
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Text            =   "Ajuste y Clasifi."
         Top             =   300
         Width           =   1940
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5800
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Text            =   "Act. Operación"
         Top             =   300
         Width           =   1940
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   7730
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Text            =   "Act. Inversión"
         Top             =   300
         Width           =   1940
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   9650
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Text            =   "Act. Financiamiento"
         Top             =   300
         Width           =   1940
      End
      Begin Sicmact.FlexEdit fgActivo 
         Height          =   1935
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   11770
         _ExtentX        =   20770
         _ExtentY        =   3413
         Cols0           =   14
         ScrollBars      =   2
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Id-Descripcion-<2013->2013-Debe-Haber-Debe-Haber-Debe-Haber-Debe-Haber-Orden"
         EncabezadosAnchos=   "350-0-1500-955-955-955-955-955-955-955-955-955-955-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-8-9-10-11-12-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmConfigHojaTrabajoFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lcDescripcion As String
Dim lnNivel As Integer
Dim lcTipo As String
Dim lnTipo As Integer
Dim lnTipoVal As Integer
Dim lcValor As String
Dim lbPeriodo As String
Dim lnId As Integer
Dim lnOrden As Integer
Dim lbEstado As Boolean
Dim i As Integer
Dim guardar As Integer


Private Sub Form_Load()
    CentraForm Me
    'Hoja de Trabajo
    IniciarControlesPasivo
    IniciarControlesActivo
    IniciarControlesAjuste
End Sub
Public Sub Inicio(ByVal psOpeCod As String, psOpeDesc As String)
    Show 1
End Sub
'******************************* Hojas de Trabajo [Flujo de Efectivo] *************************
'******************************* Activo *************************
Private Sub cmdAgregarActivo_Click()
    If Not (fgActivo.Rows - 1 = 1 And Len(Trim(fgActivo.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
        If validarRegistroActivo = False Then Exit Sub
    End If
    fgActivo.AdicionaFila
    fgActivo.SetFocus
    SendKeys "{Enter}"
End Sub
Private Sub cmdQuitarActivo_Click()
Dim row, col As Integer
Dim oRep As New NRepFormula
Dim lnId As Integer
row = fgActivo.row
lnId = Me.fgActivo.TextMatrix(row, 1)
Call oRep.ModificarEstadoHojaTrabajo(lnId)
Call IniciarControlesActivo
End Sub
Private Sub cmdSubirActivo_Click()
Dim lnId_1 As Integer, lnId_2 As Integer
Dim lcDescripcion_1 As String, lcDescripcion_2 As String
Dim lcForMen2013_1 As String, lcForMen2013_2 As String
Dim lcForMay2013_1 As String, lcForMay2013_2 As String
Dim lcAjusDebe_1 As String, lcAjusDebe_2 As String
Dim lcAjusHaber_1 As String, lcAjusHaber_2 As String
Dim lcOperDebe_1 As String, lcOperDebe_2 As String
Dim lcOperHaber_1 As String, lcOperHaber_2 As String
Dim lcInveDebe_1 As String, lcInveDebe_2 As String
Dim lcInveHaber_1 As String, lcInveHaber_2 As String
Dim lcFinaDebe_1 As String, lcFinaDebe_2 As String
Dim lcFinaHaber_1 As String, lcFinaHaber_2 As String

 If fgActivo.row > 1 Then
        lnId_1 = fgActivo.TextMatrix(fgActivo.row - 1, 1)
        lcDescripcion_1 = fgActivo.TextMatrix(fgActivo.row - 1, 2)
        lcForMen2013_1 = fgActivo.TextMatrix(fgActivo.row - 1, 3)
        lcForMay2013_1 = fgActivo.TextMatrix(fgActivo.row - 1, 4)
        lcAjusDebe_1 = fgActivo.TextMatrix(fgActivo.row - 1, 5)
        lcAjusHaber_1 = fgActivo.TextMatrix(fgActivo.row - 1, 6)
        lcOperDebe_1 = fgActivo.TextMatrix(fgActivo.row - 1, 7)
        lcOperHaber_1 = fgActivo.TextMatrix(fgActivo.row - 1, 8)
        lcInveDebe_1 = fgActivo.TextMatrix(fgActivo.row - 1, 9)
        lcInveHaber_1 = fgActivo.TextMatrix(fgActivo.row - 1, 10)
        lcFinaDebe_1 = fgActivo.TextMatrix(fgActivo.row - 1, 11)
        lcFinaHaber_1 = fgActivo.TextMatrix(fgActivo.row - 1, 12)
        
        lnId_2 = fgActivo.TextMatrix(fgActivo.row, 1)
        lcDescripcion_2 = fgActivo.TextMatrix(fgActivo.row, 2)
        lcForMen2013_2 = fgActivo.TextMatrix(fgActivo.row, 3)
        lcForMay2013_2 = fgActivo.TextMatrix(fgActivo.row, 4)
        lcAjusDebe_2 = fgActivo.TextMatrix(fgActivo.row, 5)
        lcAjusHaber_2 = fgActivo.TextMatrix(fgActivo.row, 6)
        lcOperDebe_2 = fgActivo.TextMatrix(fgActivo.row, 7)
        lcOperHaber_2 = fgActivo.TextMatrix(fgActivo.row, 8)
        lcInveDebe_2 = fgActivo.TextMatrix(fgActivo.row, 9)
        lcInveHaber_2 = fgActivo.TextMatrix(fgActivo.row, 10)
        lcFinaDebe_2 = fgActivo.TextMatrix(fgActivo.row, 11)
        lcFinaHaber_2 = fgActivo.TextMatrix(fgActivo.row, 12)
        
        fgActivo.TextMatrix(fgActivo.row - 1, 1) = lnId_2
        fgActivo.TextMatrix(fgActivo.row - 1, 2) = lcDescripcion_2
        fgActivo.TextMatrix(fgActivo.row - 1, 3) = lcForMen2013_2
        fgActivo.TextMatrix(fgActivo.row - 1, 4) = lcForMay2013_2
        fgActivo.TextMatrix(fgActivo.row - 1, 5) = lcAjusDebe_2
        fgActivo.TextMatrix(fgActivo.row - 1, 6) = lcAjusHaber_2
        fgActivo.TextMatrix(fgActivo.row - 1, 7) = lcOperDebe_2
        fgActivo.TextMatrix(fgActivo.row - 1, 8) = lcOperHaber_2
        fgActivo.TextMatrix(fgActivo.row - 1, 9) = lcInveDebe_2
        fgActivo.TextMatrix(fgActivo.row - 1, 10) = lcInveHaber_2
        fgActivo.TextMatrix(fgActivo.row - 1, 11) = lcFinaDebe_2
        fgActivo.TextMatrix(fgActivo.row - 1, 12) = lcFinaHaber_2
        
        fgActivo.TextMatrix(fgActivo.row, 1) = lnId_1
        fgActivo.TextMatrix(fgActivo.row, 2) = lcDescripcion_1
        fgActivo.TextMatrix(fgActivo.row, 3) = lcForMen2013_1
        fgActivo.TextMatrix(fgActivo.row, 4) = lcForMay2013_1
        fgActivo.TextMatrix(fgActivo.row, 5) = lcAjusDebe_1
        fgActivo.TextMatrix(fgActivo.row, 6) = lcAjusHaber_1
        fgActivo.TextMatrix(fgActivo.row, 7) = lcOperDebe_1
        fgActivo.TextMatrix(fgActivo.row, 8) = lcOperHaber_1
        fgActivo.TextMatrix(fgActivo.row, 9) = lcInveDebe_1
        fgActivo.TextMatrix(fgActivo.row, 10) = lcInveHaber_1
        fgActivo.TextMatrix(fgActivo.row, 11) = lcFinaDebe_1
        fgActivo.TextMatrix(fgActivo.row, 12) = lcFinaHaber_1
               
        fgActivo.row = fgActivo.row - 1
        fgActivo.SetFocus
    End If
End Sub
Private Sub cmdBajarActivo_Click()
Dim lnId_1 As Integer, lnId_2 As Integer
Dim lcDescripcion_1 As String, lcDescripcion_2 As String
Dim lcForMen2013_1 As String, lcForMen2013_2 As String
Dim lcForMay2013_1 As String, lcForMay2013_2 As String
Dim lcAjusDebe_1 As String, lcAjusDebe_2 As String
Dim lcAjusHaber_1 As String, lcAjusHaber_2 As String
Dim lcOperDebe_1 As String, lcOperDebe_2 As String
Dim lcOperHaber_1 As String, lcOperHaber_2 As String
Dim lcInveDebe_1 As String, lcInveDebe_2 As String
Dim lcInveHaber_1 As String, lcInveHaber_2 As String
Dim lcFinaDebe_1 As String, lcFinaDebe_2 As String
Dim lcFinaHaber_1 As String, lcFinaHaber_2 As String

 If fgActivo.row < fgActivo.Rows - 1 Then
        lnId_1 = fgActivo.TextMatrix(fgActivo.row + 1, 1)
        lcDescripcion_1 = fgActivo.TextMatrix(fgActivo.row + 1, 2)
        lcForMen2013_1 = fgActivo.TextMatrix(fgActivo.row + 1, 3)
        lcForMay2013_1 = fgActivo.TextMatrix(fgActivo.row + 1, 4)
        lcAjusDebe_1 = fgActivo.TextMatrix(fgActivo.row + 1, 5)
        lcAjusHaber_1 = fgActivo.TextMatrix(fgActivo.row + 1, 6)
        lcOperDebe_1 = fgActivo.TextMatrix(fgActivo.row + 1, 7)
        lcOperHaber_1 = fgActivo.TextMatrix(fgActivo.row + 1, 8)
        lcInveDebe_1 = fgActivo.TextMatrix(fgActivo.row + 1, 9)
        lcInveHaber_1 = fgActivo.TextMatrix(fgActivo.row + 1, 10)
        lcFinaDebe_1 = fgActivo.TextMatrix(fgActivo.row + 1, 11)
        lcFinaHaber_1 = fgActivo.TextMatrix(fgActivo.row + 1, 12)
        
        lnId_2 = fgActivo.TextMatrix(fgActivo.row, 1)
        lcDescripcion_2 = fgActivo.TextMatrix(fgActivo.row, 2)
        lcForMen2013_2 = fgActivo.TextMatrix(fgActivo.row, 3)
        lcForMay2013_2 = fgActivo.TextMatrix(fgActivo.row, 4)
        lcAjusDebe_2 = fgActivo.TextMatrix(fgActivo.row, 5)
        lcAjusHaber_2 = fgActivo.TextMatrix(fgActivo.row, 6)
        lcOperDebe_2 = fgActivo.TextMatrix(fgActivo.row, 7)
        lcOperHaber_2 = fgActivo.TextMatrix(fgActivo.row, 8)
        lcInveDebe_2 = fgActivo.TextMatrix(fgActivo.row, 9)
        lcInveHaber_2 = fgActivo.TextMatrix(fgActivo.row, 10)
        lcFinaDebe_2 = fgActivo.TextMatrix(fgActivo.row, 11)
        lcFinaHaber_2 = fgActivo.TextMatrix(fgActivo.row, 12)
        
        fgActivo.TextMatrix(fgActivo.row + 1, 1) = lnId_2
        fgActivo.TextMatrix(fgActivo.row + 1, 2) = lcDescripcion_2
        fgActivo.TextMatrix(fgActivo.row + 1, 3) = lcForMen2013_2
        fgActivo.TextMatrix(fgActivo.row + 1, 4) = lcForMay2013_2
        fgActivo.TextMatrix(fgActivo.row + 1, 5) = lcAjusDebe_2
        fgActivo.TextMatrix(fgActivo.row + 1, 6) = lcAjusHaber_2
        fgActivo.TextMatrix(fgActivo.row + 1, 7) = lcOperDebe_2
        fgActivo.TextMatrix(fgActivo.row + 1, 8) = lcOperHaber_2
        fgActivo.TextMatrix(fgActivo.row + 1, 9) = lcInveDebe_2
        fgActivo.TextMatrix(fgActivo.row + 1, 10) = lcInveHaber_2
        fgActivo.TextMatrix(fgActivo.row + 1, 11) = lcFinaDebe_2
        fgActivo.TextMatrix(fgActivo.row + 1, 12) = lcFinaHaber_2
        
        fgActivo.TextMatrix(fgActivo.row, 1) = lnId_1
        fgActivo.TextMatrix(fgActivo.row, 2) = lcDescripcion_1
        fgActivo.TextMatrix(fgActivo.row, 3) = lcForMen2013_1
        fgActivo.TextMatrix(fgActivo.row, 4) = lcForMay2013_1
        fgActivo.TextMatrix(fgActivo.row, 5) = lcAjusDebe_1
        fgActivo.TextMatrix(fgActivo.row, 6) = lcAjusHaber_1
        fgActivo.TextMatrix(fgActivo.row, 7) = lcOperDebe_1
        fgActivo.TextMatrix(fgActivo.row, 8) = lcOperHaber_1
        fgActivo.TextMatrix(fgActivo.row, 9) = lcInveDebe_1
        fgActivo.TextMatrix(fgActivo.row, 10) = lcInveHaber_1
        fgActivo.TextMatrix(fgActivo.row, 11) = lcFinaDebe_1
        fgActivo.TextMatrix(fgActivo.row, 12) = lcFinaHaber_1
               
        fgActivo.row = fgActivo.row + 1
        fgActivo.SetFocus
    End If
End Sub
Private Sub cmdGuardarActivo_Click()
    Dim oRep As New NRepFormula
    Dim lsMovNro As String
    Dim lbExito As Boolean
    Dim MatFlujo As Variant
    Dim i As Long
    
    If validarGrabarActivo = False Then Exit Sub
    If validarRegistroActivo = False Then Exit Sub
    
    If MsgBox("¿Esta seguro de guardar la configuración de flujo de Efectivo?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    ReDim MatFlujo(1 To 17, 0)
    For i = 1 To fgActivo.Rows - 1
        ReDim Preserve MatFlujo(1 To 17, 1 To i)
        MatFlujo(1, i) = IIf(Trim(fgActivo.TextMatrix(i, 1)) = "", 0, Trim(fgActivo.TextMatrix(i, 1)))
        MatFlujo(2, i) = 1 'Tipo - 1 es Activo
        MatFlujo(3, i) = 0 'Nivel -
        MatFlujo(4, i) = Trim(fgActivo.TextMatrix(i, 2)) 'Descripcion
        MatFlujo(5, i) = Trim(fgActivo.TextMatrix(i, 3))  'cForMen2013
        MatFlujo(6, i) = Trim(fgActivo.TextMatrix(i, 4))  'cForMay2013
        MatFlujo(7, i) = Trim(fgActivo.TextMatrix(i, 5)) 'cAjusDebe
        MatFlujo(8, i) = Trim(fgActivo.TextMatrix(i, 6)) 'cAjusHaber
        MatFlujo(9, i) = Trim(fgActivo.TextMatrix(i, 7)) 'cOperDebe
        MatFlujo(10, i) = Trim(fgActivo.TextMatrix(i, 8))  'cOperHaber
        MatFlujo(11, i) = Trim(fgActivo.TextMatrix(i, 9))  'cInveDebe
        MatFlujo(12, i) = Trim(fgActivo.TextMatrix(i, 10)) 'cInveHaber
        MatFlujo(13, i) = Trim(fgActivo.TextMatrix(i, 11))  'cFinaDebe
        MatFlujo(14, i) = Trim(fgActivo.TextMatrix(i, 12)) 'cFinaHaber
        MatFlujo(15, i) = True 'bEstado
        MatFlujo(16, i) = Trim(fgActivo.TextMatrix(i, 0)) 'nOrden
        MatFlujo(17, i) = False 'bMov
    Next
    
    lbExito = oRep.RegistrarHojaTrabajoFE(MatFlujo)
      
    If lbExito Then
        MsgBox "Se ha grabado satisfactoriamente los cambios en los flujos de efectivos", vbInformation, "Aviso"
        'cmdCancelar_Click
    Else
        MsgBox "No se ha podido grabar los cambios realizados, vuelva a intentarlo, si persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If

    Call IniciarControlesActivo
    
    Set oRep = Nothing
    Set MatFlujo = Nothing
End Sub
Private Sub cmdCancelarActivo_Click()
IniciarControlesActivo
End Sub
Private Function validarRegistroActivo() As Boolean
    validarRegistroActivo = True
    Dim i As Long, j As Long
    For i = 1 To fgActivo.Rows - 1 'valida fila x fila
        For j = 1 To fgActivo.Cols - 2 '2 porque el id esta oculto
            If j = 2 Then 'solo la descripcion
                If Trim(fgActivo.TextMatrix(i, j)) = "" Then
                    validarRegistroActivo = False
                    MsgBox "Ud. debe de ingresar el dato '" & UCase(fgActivo.TextMatrix(0, j)) & "'", vbInformation, "Aviso"
                    fgActivo.row = i
                    fgActivo.col = j
                    fgActivo.SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
End Function
Private Function validarGrabarActivo() As Boolean
    Dim i As Long
    validarGrabarActivo = True
    If fgActivo.Rows - 1 = 1 And fgActivo.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de registrar los flujos de efectivos", vbCritical, "Aviso"
        validarGrabarActivo = False
        fgActivo.SetFocus
        Exit Function
    End If
End Function
Private Sub IniciarControlesActivo()
    Call ListarHojaTrabajoActivo
    Me.cmdSubirActivo.Enabled = False
    Me.cmdBajarActivo.Enabled = False
End Sub
Private Sub ListarHojaTrabajoActivo()
    Dim oRep As New NRepFormula
    Dim nValor As Integer
    Dim rsHoja As New ADODB.Recordset
    Dim i As Long
    nValor = 1 ' 1 = Activo
    Set rsHoja = oRep.RecuperaHojaTrabajoFE(nValor)
    Call LimpiaFlex(fgActivo)
    If Not RSVacio(rsHoja) Then
        For i = 1 To rsHoja.RecordCount
            Me.fgActivo.AdicionaFila
            Me.fgActivo.TextMatrix(fgActivo.row, 1) = rsHoja!nId
            Me.fgActivo.TextMatrix(fgActivo.row, 2) = rsHoja!cDescripcion
            'Me.fgActivo.TextMatrix(fgActivo.row, 3) = rsHoja!nNivel
            'Me.fgActivo.TextMatrix(fgActivo.row, 4) = rsHoja!nTipo
            Me.fgActivo.TextMatrix(fgActivo.row, 3) = rsHoja!cForMen2013
            Me.fgActivo.TextMatrix(fgActivo.row, 4) = rsHoja!cForMay2013
            Me.fgActivo.TextMatrix(fgActivo.row, 5) = rsHoja!cAjusDebe
            Me.fgActivo.TextMatrix(fgActivo.row, 6) = rsHoja!cAjusHaber
            Me.fgActivo.TextMatrix(fgActivo.row, 7) = rsHoja!cOperDebe
            Me.fgActivo.TextMatrix(fgActivo.row, 8) = rsHoja!cOperHaber
            Me.fgActivo.TextMatrix(fgActivo.row, 9) = rsHoja!cInveDebe
            Me.fgActivo.TextMatrix(fgActivo.row, 10) = rsHoja!cInveHaber
            Me.fgActivo.TextMatrix(fgActivo.row, 11) = rsHoja!cFinaDebe
            Me.fgActivo.TextMatrix(fgActivo.row, 12) = rsHoja!cFinaHaber
            Me.fgActivo.TextMatrix(fgActivo.row, 13) = rsHoja!nOrden
            rsHoja.MoveNext
        Next
        Me.cmdQuitarActivo.Enabled = True
        If rsHoja.RecordCount >= 3 Then
            Me.cmdSubirActivo.Enabled = True
            Me.cmdBajarActivo.Enabled = True
        End If
    Else
    Me.cmdQuitarActivo.Enabled = False
    Me.cmdSubirActivo.Enabled = False
    Me.cmdBajarActivo.Enabled = False
    End If
    Set oRep = Nothing
    Set rsHoja = Nothing
End Sub
Private Sub fgActivo_Click()
If fgActivo.row > 0 Then
        If fgActivo.TextMatrix(fgActivo.row, 0) <> "" Then
            Me.cmdSubirActivo.Enabled = True
            Me.cmdBajarActivo.Enabled = True
        End If
End If
End Sub
'*******************************************************************************************
'******************************* Pasivo *************************
Private Sub cmdQuitarPasivo_Click()
Dim row, col As Integer
Dim oRep As New NRepFormula
Dim lnId As Integer
row = fgPasivo.row
col = fgPasivo.col
lnId = Me.fgPasivo.TextMatrix(row, 1)
Call oRep.ModificarEstadoHojaTrabajo(lnId)
Call IniciarControlesPasivo
End Sub
Private Sub cmdSubirPasivo_Click()
Dim lnId_1 As Integer, lnId_2 As Integer
Dim lcDescripcion_1 As String, lcDescripcion_2 As String
Dim lcForMen2013_1 As String, lcForMen2013_2 As String
Dim lcForMay2013_1 As String, lcForMay2013_2 As String
Dim lcAjusDebe_1 As String, lcAjusDebe_2 As String
Dim lcAjusHaber_1 As String, lcAjusHaber_2 As String
Dim lcOperDebe_1 As String, lcOperDebe_2 As String
Dim lcOperHaber_1 As String, lcOperHaber_2 As String
Dim lcInveDebe_1 As String, lcInveDebe_2 As String
Dim lcInveHaber_1 As String, lcInveHaber_2 As String
Dim lcFinaDebe_1 As String, lcFinaDebe_2 As String
Dim lcFinaHaber_1 As String, lcFinaHaber_2 As String

 If fgPasivo.row > 1 Then
        lnId_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 1)
        lcDescripcion_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 2)
        lcForMen2013_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 3)
        lcForMay2013_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 4)
        lcAjusDebe_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 5)
        lcAjusHaber_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 6)
        lcOperDebe_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 7)
        lcOperHaber_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 8)
        lcInveDebe_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 9)
        lcInveHaber_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 10)
        lcFinaDebe_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 11)
        lcFinaHaber_1 = fgPasivo.TextMatrix(fgPasivo.row - 1, 12)
        
        lnId_2 = fgPasivo.TextMatrix(fgPasivo.row, 1)
        lcDescripcion_2 = fgPasivo.TextMatrix(fgPasivo.row, 2)
        lcForMen2013_2 = fgPasivo.TextMatrix(fgPasivo.row, 3)
        lcForMay2013_2 = fgPasivo.TextMatrix(fgPasivo.row, 4)
        lcAjusDebe_2 = fgPasivo.TextMatrix(fgPasivo.row, 5)
        lcAjusHaber_2 = fgPasivo.TextMatrix(fgPasivo.row, 6)
        lcOperDebe_2 = fgPasivo.TextMatrix(fgPasivo.row, 7)
        lcOperHaber_2 = fgPasivo.TextMatrix(fgPasivo.row, 8)
        lcInveDebe_2 = fgPasivo.TextMatrix(fgPasivo.row, 9)
        lcInveHaber_2 = fgPasivo.TextMatrix(fgPasivo.row, 10)
        lcFinaDebe_2 = fgPasivo.TextMatrix(fgPasivo.row, 11)
        lcFinaHaber_2 = fgPasivo.TextMatrix(fgPasivo.row, 12)
        
        fgPasivo.TextMatrix(fgPasivo.row - 1, 1) = lnId_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 2) = lcDescripcion_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 3) = lcForMen2013_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 4) = lcForMay2013_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 5) = lcAjusDebe_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 6) = lcAjusHaber_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 7) = lcOperDebe_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 8) = lcOperHaber_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 9) = lcInveDebe_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 10) = lcInveHaber_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 11) = lcFinaDebe_2
        fgPasivo.TextMatrix(fgPasivo.row - 1, 12) = lcFinaHaber_2
        
        fgPasivo.TextMatrix(fgPasivo.row, 1) = lnId_1
        fgPasivo.TextMatrix(fgPasivo.row, 2) = lcDescripcion_1
        fgPasivo.TextMatrix(fgPasivo.row, 3) = lcForMen2013_1
        fgPasivo.TextMatrix(fgPasivo.row, 4) = lcForMay2013_1
        fgPasivo.TextMatrix(fgPasivo.row, 5) = lcAjusDebe_1
        fgPasivo.TextMatrix(fgPasivo.row, 6) = lcAjusHaber_1
        fgPasivo.TextMatrix(fgPasivo.row, 7) = lcOperDebe_1
        fgPasivo.TextMatrix(fgPasivo.row, 8) = lcOperHaber_1
        fgPasivo.TextMatrix(fgPasivo.row, 9) = lcInveDebe_1
        fgPasivo.TextMatrix(fgPasivo.row, 10) = lcInveHaber_1
        fgPasivo.TextMatrix(fgPasivo.row, 11) = lcFinaDebe_1
        fgPasivo.TextMatrix(fgPasivo.row, 12) = lcFinaHaber_1
               
        fgPasivo.row = fgPasivo.row - 1
        fgPasivo.SetFocus
    End If
End Sub
Private Sub cmdBajarPasivo_Click()
Dim lnId_1 As Integer, lnId_2 As Integer
Dim lcDescripcion_1 As String, lcDescripcion_2 As String
Dim lcForMen2013_1 As String, lcForMen2013_2 As String
Dim lcForMay2013_1 As String, lcForMay2013_2 As String
Dim lcAjusDebe_1 As String, lcAjusDebe_2 As String
Dim lcAjusHaber_1 As String, lcAjusHaber_2 As String
Dim lcOperDebe_1 As String, lcOperDebe_2 As String
Dim lcOperHaber_1 As String, lcOperHaber_2 As String
Dim lcInveDebe_1 As String, lcInveDebe_2 As String
Dim lcInveHaber_1 As String, lcInveHaber_2 As String
Dim lcFinaDebe_1 As String, lcFinaDebe_2 As String
Dim lcFinaHaber_1 As String, lcFinaHaber_2 As String

 If fgPasivo.row < fgPasivo.Rows - 1 Then
        lnId_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 1)
        lcDescripcion_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 2)
        lcForMen2013_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 3)
        lcForMay2013_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 4)
        lcAjusDebe_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 5)
        lcAjusHaber_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 6)
        lcOperDebe_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 7)
        lcOperHaber_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 8)
        lcInveDebe_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 9)
        lcInveHaber_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 10)
        lcFinaDebe_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 11)
        lcFinaHaber_1 = fgPasivo.TextMatrix(fgPasivo.row + 1, 12)
        
        lnId_2 = fgPasivo.TextMatrix(fgPasivo.row, 1)
        lcDescripcion_2 = fgPasivo.TextMatrix(fgPasivo.row, 2)
        lcForMen2013_2 = fgPasivo.TextMatrix(fgPasivo.row, 3)
        lcForMay2013_2 = fgPasivo.TextMatrix(fgPasivo.row, 4)
        lcAjusDebe_2 = fgPasivo.TextMatrix(fgPasivo.row, 5)
        lcAjusHaber_2 = fgPasivo.TextMatrix(fgPasivo.row, 6)
        lcOperDebe_2 = fgPasivo.TextMatrix(fgPasivo.row, 7)
        lcOperHaber_2 = fgPasivo.TextMatrix(fgPasivo.row, 8)
        lcInveDebe_2 = fgPasivo.TextMatrix(fgPasivo.row, 9)
        lcInveHaber_2 = fgPasivo.TextMatrix(fgPasivo.row, 10)
        lcFinaDebe_2 = fgPasivo.TextMatrix(fgPasivo.row, 11)
        lcFinaHaber_2 = fgPasivo.TextMatrix(fgPasivo.row, 12)
        
        fgPasivo.TextMatrix(fgPasivo.row + 1, 1) = lnId_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 2) = lcDescripcion_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 3) = lcForMen2013_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 4) = lcForMay2013_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 5) = lcAjusDebe_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 6) = lcAjusHaber_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 7) = lcOperDebe_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 8) = lcOperHaber_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 9) = lcInveDebe_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 10) = lcInveHaber_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 11) = lcFinaDebe_2
        fgPasivo.TextMatrix(fgPasivo.row + 1, 12) = lcFinaHaber_2
        
        fgPasivo.TextMatrix(fgPasivo.row, 1) = lnId_1
        fgPasivo.TextMatrix(fgPasivo.row, 2) = lcDescripcion_1
        fgPasivo.TextMatrix(fgPasivo.row, 3) = lcForMen2013_1
        fgPasivo.TextMatrix(fgPasivo.row, 4) = lcForMay2013_1
        fgPasivo.TextMatrix(fgPasivo.row, 5) = lcAjusDebe_1
        fgPasivo.TextMatrix(fgPasivo.row, 6) = lcAjusHaber_1
        fgPasivo.TextMatrix(fgPasivo.row, 7) = lcOperDebe_1
        fgPasivo.TextMatrix(fgPasivo.row, 8) = lcOperHaber_1
        fgPasivo.TextMatrix(fgPasivo.row, 9) = lcInveDebe_1
        fgPasivo.TextMatrix(fgPasivo.row, 10) = lcInveHaber_1
        fgPasivo.TextMatrix(fgPasivo.row, 11) = lcFinaDebe_1
        fgPasivo.TextMatrix(fgPasivo.row, 12) = lcFinaHaber_1
               
        fgPasivo.row = fgPasivo.row + 1
        fgPasivo.SetFocus
    End If
End Sub
Private Sub cmdAgregarPa_Click()
    If Not (fgPasivo.Rows - 1 = 1 And Len(Trim(fgPasivo.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
        If validarRegistroPasivo = False Then Exit Sub
    End If
    fgPasivo.AdicionaFila
    fgPasivo.SetFocus
    SendKeys "{Enter}"
End Sub
Private Sub cmdGuardarPasivo_Click()
    Dim oRep As New NRepFormula
    Dim lsMovNro As String
    Dim lbExito As Boolean
    Dim MatFlujo As Variant
    Dim i As Long
    
    If validarGrabarPasivo = False Then Exit Sub
    If validarRegistroPasivo = False Then Exit Sub
    
    If MsgBox("¿Esta seguro de guardar la configuración de flujo de Efectivo?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    ReDim MatFlujo(1 To 17, 0)
    For i = 1 To fgPasivo.Rows - 1
        ReDim Preserve MatFlujo(1 To 17, 1 To i)
        MatFlujo(1, i) = IIf(Trim(fgPasivo.TextMatrix(i, 1)) = "", 0, Trim(fgPasivo.TextMatrix(i, 1)))
        MatFlujo(2, i) = 2 'Tipo - Pasivo
        MatFlujo(3, i) = 0 'Nivel -
        MatFlujo(4, i) = Trim(fgPasivo.TextMatrix(i, 2)) 'Descripcion
        MatFlujo(5, i) = Trim(fgPasivo.TextMatrix(i, 3))  'cForMen2013
        MatFlujo(6, i) = Trim(fgPasivo.TextMatrix(i, 4))  'cForMay2013
        MatFlujo(7, i) = Trim(fgPasivo.TextMatrix(i, 5)) 'cAjusDebe
        MatFlujo(8, i) = Trim(fgPasivo.TextMatrix(i, 6)) 'cAjusHaber
        MatFlujo(9, i) = Trim(fgPasivo.TextMatrix(i, 7)) 'cOperDebe
        MatFlujo(10, i) = Trim(fgPasivo.TextMatrix(i, 8))  'cOperHaber
        MatFlujo(11, i) = Trim(fgPasivo.TextMatrix(i, 9))  'cInveDebe
        MatFlujo(12, i) = Trim(fgPasivo.TextMatrix(i, 10)) 'cInveHaber
        MatFlujo(13, i) = Trim(fgPasivo.TextMatrix(i, 11))  'cFinaDebe
        MatFlujo(14, i) = Trim(fgPasivo.TextMatrix(i, 12)) 'cFinaHaber
        MatFlujo(15, i) = True 'bEstado
        MatFlujo(16, i) = Trim(fgPasivo.TextMatrix(i, 0)) 'nOrden
        MatFlujo(17, i) = False 'bMov
    Next
    
    lbExito = oRep.RegistrarHojaTrabajoFE(MatFlujo)
      
    If lbExito Then
        MsgBox "Se ha grabado satisfactoriamente los cambios en los flujos de efectivos", vbInformation, "Aviso"
    Else
        MsgBox "No se ha podido grabar los cambios realizados, vuelva a intentarlo, si persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If

    Call IniciarControlesPasivo
    
    Set oRep = Nothing
    Set MatFlujo = Nothing
End Sub
Private Sub cmdCancelarPasivo_Click()
IniciarControlesPasivo
End Sub
Private Function validarRegistroPasivo() As Boolean
    validarRegistroPasivo = True
    Dim i As Long, j As Long
    For i = 1 To fgPasivo.Rows - 1 'valida fila x fila
        For j = 1 To fgPasivo.Cols - 2 '2 porque el id esta oculto
            If j = 2 Then 'solo la descripcion
                If Trim(fgPasivo.TextMatrix(i, j)) = "" Then
                    validarRegistroPasivo = False
                    MsgBox "Ud. debe de ingresar el dato '" & UCase(fgPasivo.TextMatrix(0, j)) & "'", vbInformation, "Aviso"
                    fgPasivo.row = i
                    fgPasivo.col = j
                    fgPasivo.SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
End Function
Private Function validarGrabarPasivo() As Boolean
    Dim i As Long
    validarGrabarPasivo = True
    If fgPasivo.Rows - 1 = 1 And fgPasivo.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de registrar los flujos de efectivos", vbCritical, "Aviso"
        validarGrabarPasivo = False
        fgPasivo.SetFocus
        Exit Function
    End If
End Function
Private Sub IniciarControlesPasivo()
    Call ListarHojaTrabajoPasivo
    Me.cmdSubirPasivo.Enabled = False
    Me.cmdBajarPasivo.Enabled = False
End Sub
Private Sub ListarHojaTrabajoPasivo()
    Dim oRep As New NRepFormula
    Dim nValor As Integer
    Dim rsHoja As New ADODB.Recordset
    Dim i As Long
    nValor = 2
    Set rsHoja = oRep.RecuperaHojaTrabajoFE(nValor)
    Call LimpiaFlex(fgPasivo)
    If Not RSVacio(rsHoja) Then
        For i = 1 To rsHoja.RecordCount
            Me.fgPasivo.AdicionaFila
            Me.fgPasivo.TextMatrix(fgPasivo.row, 1) = rsHoja!nId
            Me.fgPasivo.TextMatrix(fgPasivo.row, 2) = rsHoja!cDescripcion
            'Me.fgPasivo.TextMatrix(fgPasivo.row, 3) = rsHoja!nNivel
            'Me.fgPasivo.TextMatrix(fgPasivo.row, 4) = rsHoja!nTipo
            Me.fgPasivo.TextMatrix(fgPasivo.row, 3) = rsHoja!cForMen2013
            Me.fgPasivo.TextMatrix(fgPasivo.row, 4) = rsHoja!cForMay2013
            Me.fgPasivo.TextMatrix(fgPasivo.row, 5) = rsHoja!cAjusDebe
            Me.fgPasivo.TextMatrix(fgPasivo.row, 6) = rsHoja!cAjusHaber
            Me.fgPasivo.TextMatrix(fgPasivo.row, 7) = rsHoja!cOperDebe
            Me.fgPasivo.TextMatrix(fgPasivo.row, 8) = rsHoja!cOperHaber
            Me.fgPasivo.TextMatrix(fgPasivo.row, 9) = rsHoja!cInveDebe
            Me.fgPasivo.TextMatrix(fgPasivo.row, 10) = rsHoja!cInveHaber
            Me.fgPasivo.TextMatrix(fgPasivo.row, 11) = rsHoja!cFinaDebe
            Me.fgPasivo.TextMatrix(fgPasivo.row, 12) = rsHoja!cFinaHaber
            Me.fgPasivo.TextMatrix(fgPasivo.row, 13) = rsHoja!nOrden
            rsHoja.MoveNext
        Next
        Me.cmdQuitarPasivo.Enabled = True
        If rsHoja.RecordCount >= 3 Then
            Me.cmdSubirPasivo.Enabled = True
            Me.cmdBajarPasivo.Enabled = True
        End If
    Else
    Me.cmdQuitarPasivo.Enabled = False
    Me.cmdSubirPasivo.Enabled = False
    Me.cmdBajarPasivo.Enabled = False
    End If
    Set oRep = Nothing
    Set rsHoja = Nothing
End Sub
Private Sub fgPasivo_Click()
If fgPasivo.row > 0 Then
        If fgPasivo.TextMatrix(fgPasivo.row, 0) <> "" Then
            Me.cmdSubirPasivo.Enabled = True
            Me.cmdBajarPasivo.Enabled = True
        End If
End If
End Sub
'******************************* FIN PASIVO **********************
'*****************************************************************
'******************************* Ajustes *************************

Private Sub cmdAgregarAjuste_Click()
If Not (fgAjuste.Rows - 1 = 1 And Len(Trim(fgAjuste.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
        If validarRegistroPasivo = False Then Exit Sub
    End If
    fgAjuste.AdicionaFila
    fgAjuste.SetFocus
    SendKeys "{Enter}"
End Sub
Private Sub cmdQuitarAjuste_Click()
Dim row, col As Integer
Dim oRep As New NRepFormula
Dim lnId As Integer
row = fgAjuste.row
col = fgAjuste.col
lnId = Me.fgAjuste.TextMatrix(row, 1)
Call oRep.ModificarEstadoHojaTrabajo(lnId)
Call IniciarControlesAjuste
End Sub

'FRHU-Para subir y bajar con un check
'Private Sub cmdSubirAjuste_Click()
'Dim lnId_1 As Integer, lnId_2 As Integer
'Dim lcDescripcion_1 As String, lcDescripcion_2 As String
'Dim lnNivel_1 As Integer, lnNivel_2 As Integer
'Dim lcAjusDebe_1 As String, lcAjusDebe_2 As String
'Dim lcAjusHaber_1 As String, lcAjusHaber_2 As String
'Dim lcOperDebe_1 As String, lcOperDebe_2 As String
'Dim lcOperHaber_1 As String, lcOperHaber_2 As String
'Dim lcInveDebe_1 As String, lcInveDebe_2 As String
'Dim lcInveHaber_1 As String, lcInveHaber_2 As String
'Dim lcFinaDebe_1 As String, lcFinaDebe_2 As String
'Dim lcFinaHaber_1 As String, lcFinaHaber_2 As String
'Dim lcMov_1 As String, lcMov_2 As String
'
' If fgAjuste.row > 1 Then
'        lnId_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 1)
'        lcDescripcion_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 2)
'        lnNivel_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 3)
'        lcAjusDebe_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 4)
'        lcAjusHaber_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 5)
'        lcOperDebe_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 6)
'        lcOperHaber_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 7)
'        lcInveDebe_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 8)
'        lcInveHaber_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 9)
'        lcFinaDebe_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 10)
'        lcFinaHaber_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 11)
'        'lcMov_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 13)
'        If fgAjuste.TextMatrix(fgAjuste.row - 1, 13) = "." Then
'        lcMov_1 = "1"
'        End If
'
'        lnId_2 = fgAjuste.TextMatrix(fgAjuste.row, 1)
'        lcDescripcion_2 = fgAjuste.TextMatrix(fgAjuste.row, 2)
'        lnNivel_2 = fgAjuste.TextMatrix(fgAjuste.row, 3)
'        lcAjusDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 4)
'        lcAjusHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 5)
'        lcOperDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 6)
'        lcOperHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 7)
'        lcInveDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 8)
'        lcInveHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 9)
'        lcFinaDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 10)
'        lcFinaHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 11)
'        'lcMov_2 = fgAjuste.TextMatrix(fgAjuste.row, 13)
'        If fgAjuste.TextMatrix(fgAjuste.row, 13) = "." Then
'        lcMov_2 = "1"
'        End If
'
'        If lcMov_1 = "" And lcMov_2 = "" Then
'        Else
'            If lcMov_1 = "1" And lcMov_2 = "1" Then
'            Else
'                fgAjuste.row = fgAjuste.row - 1
'                fgAjuste.SetFocus
'                If lcMov_2 = "" Then
'                fgAjuste.SeleccionaChekTecla
'                fgAjuste.row = fgAjuste.row + 1
'                fgAjuste.SetFocus
'                If lcMov_1 = "1" Then
'                fgAjuste.SeleccionaChekTecla
'                End If
'                End If
'
'                If lcMov_2 = "1" Then
'                fgAjuste.TextMatrix(fgAjuste.row, 13) = "1"
'                If lcMov_1 = "" Then
'                fgAjuste.row = fgAjuste.row + 1
'                fgAjuste.SetFocus
'                fgAjuste.SeleccionaChekTecla
'                End If
'                End If
'            End If
'        End If
'
'
'
'
'
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 1) = lnId_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 2) = lcDescripcion_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 3) = lnNivel_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 4) = lcAjusDebe_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 5) = lcAjusHaber_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 6) = lcOperDebe_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 7) = lcOperHaber_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 8) = lcInveDebe_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 9) = lcInveHaber_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 10) = lcFinaDebe_2
'        fgAjuste.TextMatrix(fgAjuste.row - 1, 11) = lcFinaHaber_2
'
'
'
'        'fgAjuste.TextMatrix(fgAjuste.row - 1, 13) = ""
'
'        fgAjuste.TextMatrix(fgAjuste.row, 1) = lnId_1
'        fgAjuste.TextMatrix(fgAjuste.row, 2) = lcDescripcion_1
'        fgAjuste.TextMatrix(fgAjuste.row, 3) = lnNivel_1
'        fgAjuste.TextMatrix(fgAjuste.row, 4) = lcAjusDebe_1
'        fgAjuste.TextMatrix(fgAjuste.row, 5) = lcAjusHaber_1
'        fgAjuste.TextMatrix(fgAjuste.row, 6) = lcOperDebe_1
'        fgAjuste.TextMatrix(fgAjuste.row, 7) = lcOperHaber_1
'        fgAjuste.TextMatrix(fgAjuste.row, 8) = lcInveDebe_1
'        fgAjuste.TextMatrix(fgAjuste.row, 9) = lcInveHaber_1
'        fgAjuste.TextMatrix(fgAjuste.row, 10) = lcFinaDebe_1
'        fgAjuste.TextMatrix(fgAjuste.row, 11) = lcFinaHaber_1
'        'fgAjuste.SeleccionaChekTecla
'        'fgAjuste.TextMatrix(fgAjuste.row, 13) = "1"
'
'        fgAjuste.row = fgAjuste.row - 1
'        fgAjuste.SetFocus
'    End If
'
'End Sub
'FRHU-FIN Para subir y bajar con un check

Private Sub cmdSubirAjuste_Click()
Dim lnId_1 As Integer, lnId_2 As Integer
Dim lcDescripcion_1 As String, lcDescripcion_2 As String
Dim lnNivel_1 As Integer, lnNivel_2 As Integer
Dim lcAjusDebe_1 As String, lcAjusDebe_2 As String
Dim lcAjusHaber_1 As String, lcAjusHaber_2 As String
Dim lcOperDebe_1 As String, lcOperDebe_2 As String
Dim lcOperHaber_1 As String, lcOperHaber_2 As String
Dim lcInveDebe_1 As String, lcInveDebe_2 As String
Dim lcInveHaber_1 As String, lcInveHaber_2 As String
Dim lcFinaDebe_1 As String, lcFinaDebe_2 As String
Dim lcFinaHaber_1 As String, lcFinaHaber_2 As String

 If fgAjuste.row > 1 Then
        lnId_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 1)
        lcDescripcion_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 2)
        lnNivel_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 3)
        lcAjusDebe_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 4)
        lcAjusHaber_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 5)
        lcOperDebe_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 6)
        lcOperHaber_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 7)
        lcInveDebe_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 8)
        lcInveHaber_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 9)
        lcFinaDebe_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 10)
        lcFinaHaber_1 = fgAjuste.TextMatrix(fgAjuste.row - 1, 11)
        
        lnId_2 = fgAjuste.TextMatrix(fgAjuste.row, 1)
        lcDescripcion_2 = fgAjuste.TextMatrix(fgAjuste.row, 2)
        lnNivel_2 = fgAjuste.TextMatrix(fgAjuste.row, 3)
        lcAjusDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 4)
        lcAjusHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 5)
        lcOperDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 6)
        lcOperHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 7)
        lcInveDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 8)
        lcInveHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 9)
        lcFinaDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 10)
        lcFinaHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 11)
            
        fgAjuste.TextMatrix(fgAjuste.row - 1, 1) = lnId_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 2) = lcDescripcion_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 3) = lnNivel_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 4) = lcAjusDebe_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 5) = lcAjusHaber_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 6) = lcOperDebe_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 7) = lcOperHaber_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 8) = lcInveDebe_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 9) = lcInveHaber_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 10) = lcFinaDebe_2
        fgAjuste.TextMatrix(fgAjuste.row - 1, 11) = lcFinaHaber_2
        
        fgAjuste.TextMatrix(fgAjuste.row, 1) = lnId_1
        fgAjuste.TextMatrix(fgAjuste.row, 2) = lcDescripcion_1
        fgAjuste.TextMatrix(fgAjuste.row, 3) = lnNivel_1
        fgAjuste.TextMatrix(fgAjuste.row, 4) = lcAjusDebe_1
        fgAjuste.TextMatrix(fgAjuste.row, 5) = lcAjusHaber_1
        fgAjuste.TextMatrix(fgAjuste.row, 6) = lcOperDebe_1
        fgAjuste.TextMatrix(fgAjuste.row, 7) = lcOperHaber_1
        fgAjuste.TextMatrix(fgAjuste.row, 8) = lcInveDebe_1
        fgAjuste.TextMatrix(fgAjuste.row, 9) = lcInveHaber_1
        fgAjuste.TextMatrix(fgAjuste.row, 10) = lcFinaDebe_1
        fgAjuste.TextMatrix(fgAjuste.row, 11) = lcFinaHaber_1
            
        fgAjuste.row = fgAjuste.row - 1
        fgAjuste.SetFocus
    End If

End Sub
Private Sub cmdBajarAjuste_Click()
Dim lnId_1 As Integer, lnId_2 As Integer
Dim lcDescripcion_1 As String, lcDescripcion_2 As String
Dim lnNivel_1 As Integer, lnNivel_2 As Integer
Dim lcAjusDebe_1 As String, lcAjusDebe_2 As String
Dim lcAjusHaber_1 As String, lcAjusHaber_2 As String
Dim lcOperDebe_1 As String, lcOperDebe_2 As String
Dim lcOperHaber_1 As String, lcOperHaber_2 As String
Dim lcInveDebe_1 As String, lcInveDebe_2 As String
Dim lcInveHaber_1 As String, lcInveHaber_2 As String
Dim lcFinaDebe_1 As String, lcFinaDebe_2 As String
Dim lcFinaHaber_1 As String, lcFinaHaber_2 As String

 If fgAjuste.row < fgAjuste.Rows - 1 Then
        lnId_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 1)
        lcDescripcion_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 2)
        lnNivel_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 3)
        lcAjusDebe_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 4)
        lcAjusHaber_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 5)
        lcOperDebe_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 6)
        lcOperHaber_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 7)
        lcInveDebe_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 8)
        lcInveHaber_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 9)
        lcFinaDebe_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 10)
        lcFinaHaber_1 = fgAjuste.TextMatrix(fgAjuste.row + 1, 11)
        
        lnId_2 = fgAjuste.TextMatrix(fgAjuste.row, 1)
        lcDescripcion_2 = fgAjuste.TextMatrix(fgAjuste.row, 2)
        lnNivel_2 = fgAjuste.TextMatrix(fgAjuste.row, 3)
        lcAjusDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 4)
        lcAjusHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 5)
        lcOperDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 6)
        lcOperHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 7)
        lcInveDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 8)
        lcInveHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 9)
        lcFinaDebe_2 = fgAjuste.TextMatrix(fgAjuste.row, 10)
        lcFinaHaber_2 = fgAjuste.TextMatrix(fgAjuste.row, 11)
        
        fgAjuste.TextMatrix(fgAjuste.row + 1, 1) = lnId_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 2) = lcDescripcion_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 3) = lnNivel_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 4) = lcAjusDebe_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 5) = lcAjusHaber_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 6) = lcOperDebe_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 7) = lcOperHaber_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 8) = lcInveDebe_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 9) = lcInveHaber_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 10) = lcFinaDebe_2
        fgAjuste.TextMatrix(fgAjuste.row + 1, 11) = lcFinaHaber_2
        
        fgAjuste.TextMatrix(fgAjuste.row, 1) = lnId_1
        fgAjuste.TextMatrix(fgAjuste.row, 2) = lcDescripcion_1
        fgAjuste.TextMatrix(fgAjuste.row, 3) = lnNivel_1
        fgAjuste.TextMatrix(fgAjuste.row, 4) = lcAjusDebe_1
        fgAjuste.TextMatrix(fgAjuste.row, 5) = lcAjusHaber_1
        fgAjuste.TextMatrix(fgAjuste.row, 6) = lcOperDebe_1
        fgAjuste.TextMatrix(fgAjuste.row, 7) = lcOperHaber_1
        fgAjuste.TextMatrix(fgAjuste.row, 8) = lcInveDebe_1
        fgAjuste.TextMatrix(fgAjuste.row, 9) = lcInveHaber_1
        fgAjuste.TextMatrix(fgAjuste.row, 10) = lcFinaDebe_1
        fgAjuste.TextMatrix(fgAjuste.row, 11) = lcFinaHaber_1
               
        fgAjuste.row = fgAjuste.row + 1
        fgAjuste.SetFocus
    End If
End Sub
Private Sub cmdGuardarAjuste_Click()
 Dim oRep As New NRepFormula
    Dim lsMovNro As String
    Dim lbExito As Boolean
    Dim MatFlujo As Variant
    Dim i As Long
    
    If validarGrabarAjuste = False Then Exit Sub
    If validarRegistroAjuste = False Then Exit Sub

    If MsgBox("¿Esta seguro de guardar la configuración de flujo de Efectivo?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    ReDim MatFlujo(1 To 17, 0)
    For i = 1 To fgAjuste.Rows - 1
        ReDim Preserve MatFlujo(1 To 17, 1 To i)
        MatFlujo(1, i) = IIf(Trim(fgAjuste.TextMatrix(i, 1)) = "", 0, Trim(fgAjuste.TextMatrix(i, 1)))
        MatFlujo(2, i) = 3 'Tipo - Ajuste
        MatFlujo(3, i) = Trim(fgAjuste.TextMatrix(i, 3)) 'Nivel
        MatFlujo(4, i) = Trim(fgAjuste.TextMatrix(i, 2)) 'Descripcion
        MatFlujo(5, i) = 0  'cForMen2013
        MatFlujo(6, i) = 0  'cForMay2013
        MatFlujo(7, i) = Trim(fgAjuste.TextMatrix(i, 4)) 'cAjusDebe
        MatFlujo(8, i) = Trim(fgAjuste.TextMatrix(i, 5)) 'cAjusHaber
        MatFlujo(9, i) = Trim(fgAjuste.TextMatrix(i, 6)) 'cOperDebe
        MatFlujo(10, i) = Trim(fgAjuste.TextMatrix(i, 7))  'cOperHaber
        MatFlujo(11, i) = Trim(fgAjuste.TextMatrix(i, 8))  'cInveDebe
        MatFlujo(12, i) = Trim(fgAjuste.TextMatrix(i, 9)) 'cInveHaber
        MatFlujo(13, i) = Trim(fgAjuste.TextMatrix(i, 10))  'cFinaDebe
        MatFlujo(14, i) = Trim(fgAjuste.TextMatrix(i, 11)) 'cFinaHaber
        MatFlujo(15, i) = True 'bEstado
        MatFlujo(16, i) = Trim(fgAjuste.TextMatrix(i, 0)) 'nOrden
        'If fgAjuste.TextMatrix(i, 13) = "." Then
        'MatFlujo(17, i) = True 'bMov
        'Else
        MatFlujo(17, i) = False 'bMov
        'End If
        
    Next
    
    lbExito = oRep.RegistrarHojaTrabajoFE(MatFlujo)
      
    If lbExito Then
        MsgBox "Se ha grabado satisfactoriamente los cambios en los flujos de efectivos", vbInformation, "Aviso"
    Else
        MsgBox "No se ha podido grabar los cambios realizados, vuelva a intentarlo, si persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If

    Call IniciarControlesAjuste
    
    Set oRep = Nothing
    Set MatFlujo = Nothing
End Sub
Private Sub cmdCancelarAjuste_Click()
Call IniciarControlesAjuste
End Sub
Private Function validarRegistroAjuste() As Boolean
    validarRegistroAjuste = True
    Dim i As Long, j As Long
    For i = 1 To fgAjuste.Rows - 1 'valida fila x fila
        For j = 1 To fgAjuste.Cols - 2 '2 porque el id esta oculto
            If j = 2 Or j = 3 Then 'solo la descripcion
                If Trim(fgAjuste.TextMatrix(i, j)) = "" Then
                    validarRegistroAjuste = False
                    MsgBox "Ud. debe de ingresar el dato '" & UCase(fgAjuste.TextMatrix(0, j)) & "'", vbInformation, "Aviso"
                    fgAjuste.row = i
                    fgAjuste.col = j
                    fgAjuste.SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
End Function
Private Function validarGrabarAjuste() As Boolean
    Dim i As Long
    validarGrabarAjuste = True
    If fgAjuste.Rows - 1 = 1 And fgAjuste.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de registrar los flujos de efectivos", vbCritical, "Aviso"
        validarGrabarAjuste = False
        fgAjuste.SetFocus
        Exit Function
    End If
End Function
Private Sub IniciarControlesAjuste()
    Call ListarHojaTrabajoAjuste
    Me.cmdSubirAjuste.Enabled = False
    Me.cmdBajarAjuste.Enabled = False
    Me.cmdQuitarAjuste.Enabled = False
End Sub
Private Sub ListarHojaTrabajoAjuste()
    Dim oRep As New NRepFormula
    Dim nValor As Integer
    Dim rsHoja As New ADODB.Recordset
    Dim i As Long
    nValor = 3
    Set rsHoja = oRep.RecuperaHojaTrabajoFE(nValor)
    Call LimpiaFlex(fgAjuste)
    If Not RSVacio(rsHoja) Then
        For i = 1 To rsHoja.RecordCount
            Me.fgAjuste.AdicionaFila
            Me.fgAjuste.TextMatrix(fgAjuste.row, 1) = rsHoja!nId
            Me.fgAjuste.TextMatrix(fgAjuste.row, 2) = rsHoja!cDescripcion
            Me.fgAjuste.TextMatrix(fgAjuste.row, 3) = rsHoja!nNivel
            'Me.fgAjuste.TextMatrix(fgAjuste.row, 4) = rsHoja!nTipo
            'Me.fgAjuste.TextMatrix(fgAjuste.row, 5) = rsHoja!cForMen2013
            'Me.fgAjuste.TextMatrix(fgAjuste.row, 6) = rsHoja!cForMay2013
            Me.fgAjuste.TextMatrix(fgAjuste.row, 4) = rsHoja!cAjusDebe
            Me.fgAjuste.TextMatrix(fgAjuste.row, 5) = rsHoja!cAjusHaber
            Me.fgAjuste.TextMatrix(fgAjuste.row, 6) = rsHoja!cOperDebe
            Me.fgAjuste.TextMatrix(fgAjuste.row, 7) = rsHoja!cOperHaber
            Me.fgAjuste.TextMatrix(fgAjuste.row, 8) = rsHoja!cInveDebe
            Me.fgAjuste.TextMatrix(fgAjuste.row, 9) = rsHoja!cInveHaber
            Me.fgAjuste.TextMatrix(fgAjuste.row, 10) = rsHoja!cFinaDebe
            Me.fgAjuste.TextMatrix(fgAjuste.row, 11) = rsHoja!cFinaHaber
            Me.fgAjuste.TextMatrix(fgAjuste.row, 12) = rsHoja!nOrden
            'If rsHoja!bMov = True Then
            '    Me.fgAjuste.TextMatrix(fgAjuste.row, 13) = "1"
            'End If
            rsHoja.MoveNext
        Next
        Me.cmdQuitarAjuste.Enabled = True
        If rsHoja.RecordCount >= 3 Then
            Me.cmdSubirAjuste.Enabled = True
            Me.cmdBajarAjuste.Enabled = True
        End If
    Else
    Me.cmdQuitarAjuste.Enabled = False
    Me.cmdSubirAjuste.Enabled = False
    Me.cmdBajarAjuste.Enabled = False
    End If
    Set oRep = Nothing
    Set rsHoja = Nothing
End Sub
Private Sub fgAjuste_Click()
If fgAjuste.row > 0 Then
        If fgAjuste.TextMatrix(fgAjuste.row, 0) <> "" Then
            Me.cmdSubirAjuste.Enabled = True
            Me.cmdBajarAjuste.Enabled = True
            Me.cmdQuitarAjuste.Enabled = True
        End If
End If
End Sub

'******************** FIN AJUSTES ***********************************

