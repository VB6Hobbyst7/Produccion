VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMntCtasContabCambios 
   Caption         =   "Cuentas Contables: Control de Cambios"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10920
   Icon            =   "frmMntCtasContabCambios.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   9360
      TabIndex        =   7
      Top             =   5310
      Width           =   1275
   End
   Begin VB.Frame FraRangoFechas 
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   270
      TabIndex        =   4
      Top             =   30
      Width           =   9030
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7470
         TabIndex        =   2
         Top             =   180
         Width           =   1410
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   300
         Left            =   2565
         TabIndex        =   0
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHasta 
         Height          =   300
         Left            =   4485
         TabIndex        =   1
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3900
         TabIndex        =   6
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1920
         TabIndex        =   5
         Top             =   210
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   4335
      Left            =   90
      TabIndex        =   3
      Top             =   870
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cCtaContCod"
         Caption         =   "Cuenta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cCtaContDesc"
         Caption         =   "Descripción"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cFecha"
         Caption         =   "Fecha - Hora"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cMaquina"
         Caption         =   "Maquina"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2340.284
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   3254.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2160
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMntCtasContabCambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdProcesar_Click()
Dim oCon As New DConecta
Dim rs   As New ADODB.Recordset
oCon.AbreConexion
Set rs = oCon.CargaRecordSet("SELECT cCtaContCod, cCtaContDesc, LEFT(cUltimaActualizacion,8)+'-'+SubString(cUltimaActualizacion,9,6) cFecha, SubString(cUltimaActualizacion,15,50) cMaquina FROM CtaContHisto WHERE LEFT(cUltimaActualizacion,8) BETWEEN '" & Format(txtDesde, gsFormatoMovFecha) & "' and '" & Format(txtHasta, gsFormatoMovFecha) & "' ", adLockReadOnly)
Set dg.DataSource = rs
dg.SetFocus
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub Form_Load()
CentraForm Me
txtDesde = gdFecSis
txtHasta = gdFecSis
End Sub

Private Sub txtDesde_GotFocus()
fEnfoque txtDesde
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtHasta.SetFocus
End Sub

Private Sub txtHasta_GotFocus()
fEnfoque txtHasta
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdProcesar.SetFocus
End Sub

