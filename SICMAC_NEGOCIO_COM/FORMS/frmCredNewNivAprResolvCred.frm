VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredNewNivAprResolvCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resolver Crédito"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   Icon            =   "frmCredNewNivAprResolvCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSolModif 
      Caption         =   "Sol. Modif."
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
      Left            =   2520
      TabIndex        =   29
      Top             =   8760
      Width           =   1170
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
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
      Left            =   6120
      TabIndex        =   28
      Top             =   8760
      Width           =   1170
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "Aprobar"
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
      Left            =   120
      TabIndex        =   27
      Top             =   8760
      Width           =   1170
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
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
      Left            =   1320
      TabIndex        =   26
      Top             =   8760
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos del Crédito"
      TabPicture(0)   =   "frmCredNewNivAprResolvCred.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblMoneda"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMontoSol"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCuotas"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPlazo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblCondicion"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblGracia"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label19"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblCredCmac"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblCredAge"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblCredSubProd"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblCredNum"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label24"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label25"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblExoneracion"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblAgencia"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblProducto"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblTasa"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblTitular"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblNivApr"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "feExonera"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "feNivApr"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtComent"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdVerDetalle"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdHojaAprob"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdEval"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      Begin VB.CommandButton cmdEval 
         Caption         =   "Hoja &Evaluación"
         Height          =   340
         Left            =   5715
         TabIndex        =   38
         Top             =   7920
         Width           =   1365
      End
      Begin VB.CommandButton cmdHojaAprob 
         Caption         =   "Hoja &Aprobación"
         Height          =   340
         Left            =   5715
         TabIndex        =   37
         Top             =   7560
         Width           =   1365
      End
      Begin VB.CommandButton cmdVerDetalle 
         Caption         =   "&Ver Detalle"
         Height          =   340
         Left            =   5715
         TabIndex        =   36
         Top             =   7200
         Width           =   1365
      End
      Begin VB.TextBox txtComent 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   7200
         Width           =   5535
      End
      Begin SICMACT.FlexEdit feNivApr 
         Height          =   2055
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3625
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "-cNivAprCod-Nivel-Estado-Riesgo-Usuario-Comentario"
         EncabezadosAnchos=   "300-0-2500-700-1000-800-1300"
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
         ColumnasAEditar =   "X-X-2-3-X-X-X"
         ListaControles  =   "0-0-0-4-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit feExonera 
         Height          =   1335
         Left            =   120
         TabIndex        =   32
         Top             =   5280
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2355
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "-Exoneración-Nivel-Estado-Aux"
         EncabezadosAnchos=   "300-3000-2500-800-0"
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
         ColumnasAEditar =   "X-X-X-3-X"
         ListaControles  =   "0-0-0-4-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblNivApr 
         Alignment       =   2  'Center
         Caption         =   "Nivel a Aprobar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   6765
         Width           =   6975
      End
      Begin VB.Label lblTitular 
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   915
         Width           =   5930
      End
      Begin VB.Label lblTasa 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label lblProducto 
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   2040
         Width           =   2925
      End
      Begin VB.Label lblAgencia 
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   960
         TabIndex        =   14
         Top             =   1680
         Width           =   2925
      End
      Begin VB.Label lblExoneracion 
         Caption         =   "Su aprobación incluye exoneración"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   5040
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label25 
         Caption         =   "Exoneraciones Solicitadas: "
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
         TabIndex        =   33
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Niveles de Aprobación: "
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
         TabIndex        =   31
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblCredNum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   2220
         TabIndex        =   25
         Top             =   530
         Width           =   1200
      End
      Begin VB.Label lblCredSubProd 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   1770
         TabIndex        =   24
         Top             =   530
         Width           =   435
      End
      Begin VB.Label lblCredAge 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   1410
         TabIndex        =   23
         Top             =   525
         Width           =   345
      End
      Begin VB.Label lblCredCmac 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   960
         TabIndex        =   22
         Top             =   530
         Width           =   435
      End
      Begin VB.Label Label19 
         Caption         =   "Días Gracia:"
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
         TabIndex        =   21
         Top             =   2085
         Width           =   975
      End
      Begin VB.Label lblGracia 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   5760
         TabIndex        =   20
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label17 
         Caption         =   "Condición:"
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
         TabIndex        =   19
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label lblCondicion 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   5760
         TabIndex        =   18
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label15 
         Caption         =   "Producto: "
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
         TabIndex        =   17
         Top             =   2085
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Agencia: "
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
         TabIndex        =   15
         Top             =   1725
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Plazo (Dias):"
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
         TabIndex        =   13
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label lblPlazo 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   5760
         TabIndex        =   12
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label9 
         Caption         =   "Cuotas:"
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
         Left            =   2520
         TabIndex        =   11
         Top             =   1365
         Width           =   615
      End
      Begin VB.Label lblCuotas 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   3240
         TabIndex        =   10
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label7 
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
         Left            =   1800
         TabIndex        =   9
         Top             =   1335
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Tasa:"
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
         TabIndex        =   8
         Top             =   1340
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Titular:"
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
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Crédito:"
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
         TabIndex        =   4
         Top             =   585
         Width           =   735
      End
      Begin VB.Label lblMontoSol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   5580
         TabIndex        =   3
         Top             =   530
         Width           =   1305
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   5160
         TabIndex        =   2
         Top             =   530
         Width           =   390
      End
      Begin VB.Label Label28 
         Caption         =   "Monto Solicitado:"
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
         Left            =   3840
         TabIndex        =   1
         Top             =   585
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCredNewNivAprResolvCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredNewNivAprResolvCred
'** Descripción : Formulario para Resolver los Creditos para su Aprobación creado segun RFC110-2012
'** Creación : JUEZ, 20121207 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim oNNiv As COMNCredito.NCOMNivelAprobacion
Dim oDNiv As COMDCredito.DCOMNivelAprobacion
Dim oDCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim fsCtaCod As String
Dim fsNivAprCod As String
Dim fsDNI, fsRUC As String
Dim fsLineaCred As String
Dim fbPreferencial As Boolean 'JUEZ 20160218

Public Sub ResolverCredito(ByVal psCtaCod As String, ByVal psNivAprCod As String)
    fsCtaCod = psCtaCod
    fsNivAprCod = psNivAprCod
    If CargaDatos Then
        Me.Show 1
    Else
        MsgBox "No se Encontro el Credito", vbExclamation, "Aviso"
        Unload Me
    End If
End Sub

Private Function CargaDatos() As Boolean
    CargaDatos = False
    Dim oAge As COMDConstantes.DCOMAgencias
    'Dim oNCred As COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
    Dim prsCredEval As ADODB.Recordset
    Dim lnFila As Integer
    Set oDCred = New COMDCredito.DCOMCredito
    
    CargarNivelAprUsuarioCred 'JUEZ 20160302
    
    Set prsCredEval = oDCred.RecuperaColocacCredEval(fsCtaCod)
    Set rs = oDCred.RecuperaDatosAprobacion(fsCtaCod, IIf(gnAgenciaCredEval And Not prsCredEval.EOF, 1, 0))
    Set oDCred = Nothing
    If Not rs.EOF Then
        fsDNI = rs!DNI
        fsRUC = rs!Ruc
        fsLineaCred = rs!cLineaCredCod
        lblCredCmac.Caption = "109"
        lblCredAge.Caption = Mid(fsCtaCod, 4, 2)
        lblCredSubProd.Caption = Mid(fsCtaCod, 6, 3)
        lblCredNum.Caption = Right(fsCtaCod, 10)
        lblMoneda.Caption = IIf(Mid(fsCtaCod, 9, 1) = "1", "S/", "$")
        lblMontoSol.Caption = Format(rs!nMonto, "#,##0.00")
        lblTitular.Caption = " " & rs!cPersNombre
        lblTasa.Caption = Format(rs!nTasaInteres, "#,##0.0000")
        lblCuotas = rs!nCuotas
        lblPlazo.Caption = rs!nPlazo
        Set oAge = New COMDConstantes.DCOMAgencias
        lblAgencia.Caption = " " & oAge.NombreAgencia(rs!cAgeCodAct)
        Set oAge = Nothing
        lblCondicion.Caption = rs!cCondicionCred
        lblProducto.Caption = " " & UCase(rs!cSTipoProdDescrip)
        lblGracia.Caption = rs!nPeriodoGracia
        
        'JUEZ 20160218 *******************************
        'Set oNCred = New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
            'fbPreferencial = oNCred.ValidarClientePreferencial(rs!cPersCod) 'COMENTADO POR ARLO 20170722
            fbPreferencial = False 'ARLO 20170722
        'Set oNCred = Nothing                    'COMENTADO POR ARLO 20170722
        'END JUEZ ************************************
        
        Set rs = Nothing
        
        CargarNivelesAprobadosResultado 'JUEZ 20160302
        
        Set oDNiv = New COMDCredito.DCOMNivelAprobacion
        Set rs = oDNiv.RecuperaDatosCredExoneracion(fsCtaCod)
        Set oDNiv = Nothing
        Call LimpiaFlex(feExonera)
        If Not rs.EOF Then
            Do While Not rs.EOF
                If fsNivAprCod = rs!cNivAprCod Then
                    lblExoneracion.Visible = True
                    feExonera.AdicionaFila
                    lnFila = feExonera.row
                    feExonera.TextMatrix(lnFila, 1) = rs!cExoneraDesc
                    feExonera.TextMatrix(lnFila, 2) = rs!cNivAprDesc
                    feExonera.TextMatrix(lnFila, 3) = IIf(rs!nEstado = 1, 1, "")
                End If
                rs.MoveNext
            Loop
            feExonera.TopRow = 1
        Else
            lblExoneracion.Visible = False
        End If
        Set rs = Nothing
        
        CargarNivelApr fsNivAprCod 'JUEZ 20160302
        
        CargaDatos = True
    Else
        CargaDatos = False
    End If
End Function

'JUEZ 20160302 ******************************************************
Private Sub CargarNivelesAprobadosResultado()
Dim lnFila As Integer
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaDatosCredResultado(fsCtaCod)
    Set oDNiv = Nothing
    Call LimpiaFlex(feNivApr)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feNivApr.AdicionaFila
            lnFila = feNivApr.row
            feNivApr.TextMatrix(lnFila, 1) = rs!cNivAprCod
            feNivApr.TextMatrix(lnFila, 2) = rs!cNivAprDesc
            feNivApr.TextMatrix(lnFila, 3) = IIf(rs!nEstado = 1, 1, "")
            feNivApr.TextMatrix(lnFila, 4) = rs!cTipoRiesgo
            feNivApr.TextMatrix(lnFila, 5) = rs!cUserApr
            feNivApr.TextMatrix(lnFila, 6) = "Ver"
            rs.MoveNext
        Loop
        feNivApr.TopRow = feNivApr.row
    End If
    Set rs = Nothing
End Sub

Private Sub CargarNivelApr(ByVal psNivAprCod As String)
Dim sNivAprDesc As String
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
        sNivAprDesc = oDNiv.RecuperaNivApr(fsNivAprCod)!cNivAprDesc
    Set oDNiv = Nothing
    
    lblNivApr.Caption = "Nivel a Aprobar: " & UCase(sNivAprDesc)
End Sub

Private Sub CargarNivelAprUsuarioCred(Optional ByVal psNivAprCod As String = "")
Dim sNivAprCod As String
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
        sNivAprCod = oDNiv.RecuperaNivelAprUsuarioCredito(gsCodUser, fsCtaCod)
    Set oDNiv = Nothing
    
    fsNivAprCod = IIf(sNivAprCod <> "", sNivAprCod, IIf(psNivAprCod <> "", psNivAprCod, fsNivAprCod))
End Sub
'END JUEZ ***********************************************************

Private Sub cmdRechazar_Click()
    If ValidarTexto Then
        Call frmCredRechazo.RechazarPorNiveles(fsCtaCod)
        Dim oNCred As COMNCredito.NCOMCredito
    Dim prsCredEval As ADODB.Recordset
    Dim lnFila As Integer
    Set oDCred = New COMDCredito.DCOMCredito
    
    Set prsCredEval = oDCred.RecuperaColocacCredEval(fsCtaCod)
        
        Set rs = oDCred.RecuperaDatosAprobacion(fsCtaCod, IIf(gnAgenciaCredEval And Not prsCredEval.EOF, 1, 0))
        If rs.EOF Then
            Set oNNiv = New COMNCredito.NCOMNivelAprobacion
            Call oNNiv.dDesactivaAprobacionResultado(fsCtaCod)
            Set oNNiv = Nothing
            Unload Me
        End If
    End If
End Sub

Private Sub cmdSolModif_Click()
    If ValidarTexto Then
        If MsgBox("Esta seguro de solicitar la modificación de la sugerencia ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set oNNiv = New COMNCredito.NCOMNivelAprobacion
        Call oNNiv.dSolicitaModifAprobacionResultado(fsCtaCod, Trim(txtComent.Text))
        Set oNNiv = Nothing
        MsgBox "Se ha solicitado la modificación de la Sugerencia del crédito", vbInformation, "Aviso"
        Unload Me
    End If
End Sub

Private Sub txtComent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdAprobar.SetFocus
    End If
End Sub

Private Function ValidarTexto() As Boolean
    ValidarTexto = False
    If Trim(txtComent.Text) = "" Then
        MsgBox "Debe escribir un comentario", vbInformation, "Aviso"
        txtComent.SetFocus
        ValidarTexto = False
        Exit Function
    End If
    ValidarTexto = True
End Function

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdVerDetalle_Click()
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaNivApr(fsNivAprCod)
    Set oDNiv = Nothing
    Call frmCredSugerencia_NEW.InicioVerDetalleSug(fsCtaCod, IIf(rs!nCorrigeSug = 1, True, False))
    CargaDatos
End Sub

Private Sub feNivApr_Click()
    If feNivApr.TextMatrix(feNivApr.row, 0) <> "" Then
        If feNivApr.Col = 6 Then
            Set oDNiv = New COMDCredito.DCOMNivelAprobacion
                Set rs = oDNiv.RecuperaDatosCredResultado(fsCtaCod, feNivApr.TextMatrix(feNivApr.row, 1))
            Set oDNiv = Nothing
            frmCredListaDatos.InicioTextBox "Comentarios", rs!cComent
        End If
    End If
End Sub

Private Sub cmdHojaAprob_Click()
    Call frmCredDupDoc.ImprimeAprobacionCreditos(fsCtaCod, fsDNI, fsRUC)
End Sub

Private Sub cmdEval_Click()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim nEstado As Integer
    Dim rs As ADODB.Recordset
    Dim cPrd As String
    Dim cSPrd As String
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim fnMontoExpEsteCred_NEW As Currency
    Set rs = oDCOMFormatosEval.RecuperaCredFormEvalProductoEstadoExposicion(fsCtaCod)
    
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
        EvaluarCredito fsCtaCod, False, 2001, cPrd, cSPrd, fnMontoExpEsteCred_NEW, False, True
        
        Else
            MsgBox "Nº de Crédito no se encuentra en estado Solicitado.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If



'*****-> Comentado por LUCV20160811, Según ERS004-2016
'    Dim oTipoCam As COMDConstSistema.NCOMTipoCambio
'    Dim nFormato As Integer
'    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
'    Set oDCred = New COMDCredito.DCOMCredito
'    Set rs = oDCred.RecuperaSolicitudDatoBasicos(fsCtaCod)
'        If rs.RecordCount > 0 Then
'            nFormato = oDCred.AsignarFormato(Mid(Trim(rs!cTpoProdCod), 1, 1) & "00", Trim(rs!cTpoProdCod), CDbl(lblMontoSol.Caption) * IIf(Mid(fsCtaCod, 9, 1) = "1", 1, CDbl(oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia))))
'            Select Case nFormato
'                Case 0: MsgBox "Crédito no fue evaluado por los formatos.", vbInformation, "Aviso"
'                Case 1: Call frmCredEvalFormato1.Inicio(fsCtaCod, 2)
'                Case 2: Call frmCredEvalFormato2.Inicio(fsCtaCod, 2)
'                Case 3: Call frmCredEvalFormato3.Inicio(fsCtaCod, 2)
'            End Select
'        Else
'            MsgBox "Crédito no fue evaluado por los formatos", vbInformation, "Aviso"
'        End If
'<-***** Fin comentario
End Sub

Private Sub cmdAprobar_Click()
    Dim bAprobado As Boolean 'JUEZ 20160218
    Dim bNivAprNext As Boolean 'JUEZ 20160302
    If ValidarTexto Then
        'JUEZ 20160218 **************************************************************************
        'Call frmCredNewNivAprAprobar.AprobarCredito(fsCtaCod, fsNivAprCod, CInt(Trim(lblCuotas.Caption)), CDbl(Trim(Me.lblTasa.Caption)), CDbl(Trim(Me.lblMontoSol.Caption)), fsLineaCred, Trim(txtComent.Text))
        bAprobado = frmCredNewNivAprAprobar.AprobarCredito(fsCtaCod, fsNivAprCod, CInt(Trim(lblCuotas.Caption)), CDbl(Trim(Me.lblTasa.Caption)), CDbl(Trim(Me.lblMontoSol.Caption)), fsLineaCred, Trim(txtComent.Text))
        'Set oDNiv = New COMDCredito.DCOMNivelAprobacion
        'Set rs = oDNiv.RecuperaDatosCredResultado(fsCtaCod, fsNivAprCod)
        'If Not rs.EOF Then
        If bAprobado Then
            'Unload Me
            'JUEZ 20160302 *************************************************************
            Dim rsNivelUsu As ADODB.Recordset, rsNivelPend As ADODB.Recordset, rsResult As ADODB.Recordset
            
            Set oDNiv = New COMDCredito.DCOMNivelAprobacion
            Set rsNivelPend = oDNiv.RecuperaHistorialCredAprobados(fsCtaCod, IIf(fbPreferencial = True, 2, 1))
            Set rsNivelUsu = oDNiv.RecuperaNivelesUsuario(gsCodUser)
            
            Do While Not rsNivelPend.EOF
                If CInt(rsNivelPend!nFirmasSolic) = CInt(rsNivelPend!nFirmasAprob) Or rsNivelPend!cUserApr = gsCodUser Then
                    rsNivelPend.MoveNext
                Else
                    rsNivelUsu.MoveFirst
                    Do While Not rsNivelUsu.EOF
                        If rsNivelPend!cNivAprCod = rsNivelUsu!cNivAprCod Then
                            Set rsResult = oDNiv.RecuperaDatosCredResultado(fsCtaCod, rsNivelPend!cNivAprCod, gsCodUser)
                            If rsResult.EOF And rsResult.BOF Then
                                txtComent.Text = ""
                                CargarNivelesAprobadosResultado
                                CargarNivelAprUsuarioCred rsNivelPend!cNivAprCod
                                CargarNivelApr rsNivelPend!cNivAprCod
                                Exit Sub
                            End If
                        End If
                        rsNivelUsu.MoveNext
                    Loop
                    rsNivelPend.MoveNext
                End If
            Loop
            'END JUEZ ******************************************************************
            Unload Me
        End If
        'END JUEZ *******************************************************************************
    End If
End Sub

