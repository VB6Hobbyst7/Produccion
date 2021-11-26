VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPITConsultaMovimientosAut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones InterCajas - Consulta de Movimientos"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8400
      TabIndex        =   15
      Top             =   6720
      Width           =   1230
   End
   Begin VB.Frame fraFechas 
      Caption         =   " Rango de fechas de operaciones InterCajas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   9615
      Begin VB.CommandButton cmdMovimientos 
         Caption         =   "Buscar Movimientos"
         Height          =   375
         Left            =   7320
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskFechaDe 
         Height          =   300
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaHasta 
         Height          =   300
         Left            =   6000
         TabIndex        =   12
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta: "
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "De: "
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   " Cliente "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton cmdCliente 
         Caption         =   "Busqueda de Cliente"
         Height          =   375
         Left            =   7320
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblNroRUC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblNroDNI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. RUC: "
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nro. DNI: "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraMovimientos 
      Caption         =   " Movimientos "
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
      Height          =   4575
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   9735
      Begin MSDataGridLib.DataGrid dtgMovCliente 
         Height          =   4095
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "cFecha"
            Caption         =   "Fecha"
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
            DataField       =   "cDescripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "cCuenta"
            Caption         =   "Nro. Cuenta"
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
            DataField       =   "cMoneda"
            Caption         =   "Moneda"
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
         BeginProperty Column04 
            DataField       =   "nMontoTran"
            Caption         =   "Monto"
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
         BeginProperty Column05 
            DataField       =   "nMovNro"
            Caption         =   "Nro. Mov"
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
            Size            =   182
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPITConsultaMovimientosAut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sPersCod As String

Private Sub cmdCliente_Click()
Dim loBusqPers As frmPITBuscaPersona

    sPersCod = ""
    Set loBusqPers = New frmPITBuscaPersona
    sPersCod = loBusqPers.Inicio()
    If (sPersCod = "") Then
        Call MsgBox("No se realizó la busqueda o no se encontró al cliente", vbInformation)
    Else
        lblPersCod.Caption = loBusqPers.PersCod
        LblNombre.Caption = loBusqPers.PersNombre
        lblNroDNI.Caption = loBusqPers.PersNroDNI
        lblNroRUC.Caption = loBusqPers.PersNroRUC
    End If
End Sub

Private Sub cmdMovimientos_Click()
    If Not IsDate(Me.mskFechaDe.Text) Then
        MsgBox "Fecha no valida.", vbInformation, "Aviso"
        mskFechaDe.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.mskFechaHasta.Text) Then
        MsgBox "Fecha no valida.", vbInformation, "Aviso"
        mskFechaHasta.SetFocus
        Exit Sub
    End If
    
    If sPersCod = "" Then
        MsgBox "Cliente no seleccionado", vbInformation, "Aviso"
        cmdCliente.SetFocus
        Exit Sub
    End If
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub




