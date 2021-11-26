VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredMiViviendaAlertasAdm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración  de Alertas para Créditos ""MIVIVIENDA"""
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   Icon            =   "frmCredMiViviendaAlertasAdm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Créditos Pendientes"
      TabPicture(0)   =   "frmCredMiViviendaAlertasAdm.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feCreditos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDarPase"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdActualizar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   240
         TabIndex        =   5
         Top             =   4440
         Width           =   9135
         Begin VB.Label lblmonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7320
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblcredito 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5280
            TabIndex        =   6
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label lblpago 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            TabIndex        =   8
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblcliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label1 
            Caption         =   "Datos del Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Monto a Pagar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   7
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
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
         Left            =   6480
         TabIndex        =   4
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   7920
         TabIndex        =   2
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdDarPase 
         Caption         =   "Dar Pase"
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
         Left            =   240
         TabIndex        =   1
         Top             =   5520
         Width           =   1335
      End
      Begin SICMACT.FlexEdit feCreditos 
         Height          =   3975
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7011
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Crédito-Fecha Alerta-Cliente-Monto Cancelar (Calend.)-Id-tcredito-tpago"
         EncabezadosAnchos=   "500-1800-2000-2500-1900-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-R-L-C-C"
         FormatosEdit    =   "0-0-5-0-2-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCredMiViviendaAlertasAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fDCredito As COMDCredito.DCOMCredito

Private Sub LlenarGrid()
Dim rs As ADODB.Recordset
Dim i As Integer
Set fDCredito = New COMDCredito.DCOMCredito

Set rs = fDCredito.ObtenerCredAlertaMIVIVIENDA()

LimpiaFlex feCreditos

If Not (rs.EOF And rs.BOF) Then
    For i = 1 To rs.RecordCount
        feCreditos.AdicionaFila
        feCreditos.TextMatrix(i, 1) = rs!cCtaCod
        feCreditos.TextMatrix(i, 2) = Format(rs!Fecha, "dd/mm/yyyy")
        feCreditos.TextMatrix(i, 3) = rs!cPersNombre
        feCreditos.TextMatrix(i, 4) = Format(rs!nMonto, "###," & String(15, "#") & "#0.00")
        feCreditos.TextMatrix(i, 5) = rs!nId
        feCreditos.TextMatrix(i, 6) = rs!cConsDescripcion
        feCreditos.TextMatrix(i, 7) = rs!sTpoPago
        rs.MoveNext
    Next i
End If

Set fDCredito = Nothing
End Sub

Private Sub cmdActualizar_Click()
        'CTI3 : ERS085-2018
        lblcliente.Caption = ""
        lblmonto.Caption = ""
        lblpago.Caption = ""
        lblcredito.Caption = ""
        '---------------------
LlenarGrid
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdDarPase_Click()
If Trim(feCreditos.TextMatrix(feCreditos.row, 0)) <> "" Then
    Dim Id As Long
    Id = CDbl(feCreditos.TextMatrix(feCreditos.row, 5))
    If MsgBox("Estas seguro de Dar Pase al Crédito ''" & feCreditos.TextMatrix(feCreditos.row, 1) & "''?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set fDCredito = New COMDCredito.DCOMCredito
        Call fDCredito.ActualizaMiViviendaAlertasPago(Id, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), 1)
        'CTI3 : ERS085-2018
        lblcliente.Caption = ""
        lblmonto.Caption = ""
        lblpago.Caption = ""
        lblcredito.Caption = ""
        '---------------------
        Call LlenarGrid
    End If
Else
    MsgBox "Favor de seleccionar el credito a Dar Pase correctamente", vbInformation, "Aviso"
End If
End Sub

Private Sub feCreditos_Click()
lblcliente.Caption = UCase(feCreditos.TextMatrix(feCreditos.row, 3))
lblmonto.Caption = Format(feCreditos.TextMatrix(feCreditos.row, 4), "###," & String(15, "#") & "#0.00")
lblpago.Caption = UCase(feCreditos.TextMatrix(feCreditos.row, 7))
lblcredito.Caption = UCase(feCreditos.TextMatrix(feCreditos.row, 6))
End Sub

Private Sub Form_Load()
LlenarGrid
End Sub

