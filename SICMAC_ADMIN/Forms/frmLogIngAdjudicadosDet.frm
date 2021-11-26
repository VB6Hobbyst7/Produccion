VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogIngAdjudicados 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   1320
   ClientTop       =   2430
   ClientWidth     =   9630
   Icon            =   "frmLogIngAdjudicados.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Cliente "
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
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton cmdPersona 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   345
      End
      Begin VB.TextBox txtPersCod 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   330
         Width           =   1515
      End
      Begin VB.TextBox txtPersona 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   330
         Width           =   5475
      End
      Begin VB.TextBox txtDNI 
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
         Left            =   7920
         TabIndex        =   1
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "D.N.I."
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
         Left            =   7320
         TabIndex        =   2
         Top             =   405
         Width           =   525
      End
   End
   Begin VB.Frame fraVis 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   9375
      Begin VB.Frame Frame5 
         Caption         =   "Datos del Crédito "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   3255
         Begin VB.TextBox Text6 
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
            Left            =   1500
            TabIndex        =   39
            Top             =   1380
            Width           =   1515
         End
         Begin VB.TextBox Text7 
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
            Left            =   1500
            TabIndex        =   37
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox Text5 
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
            Left            =   1500
            TabIndex        =   35
            Top             =   660
            Width           =   1515
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            TabIndex        =   41
            Top             =   300
            Width           =   1515
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cta de Crédito"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Deuda Total"
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
            Left            =   240
            TabIndex        =   40
            Top             =   1455
            Width           =   1065
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Intereses / Otros"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   1095
            Width           =   1185
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Capital"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   735
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   3360
         TabIndex        =   30
         Top             =   0
         Width           =   6015
         Begin VB.CommandButton Command2 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1260
            TabIndex        =   33
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   1095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDoc 
            Height          =   1095
            Left            =   120
            TabIndex        =   31
            Top             =   300
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   1931
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483633
            ScrollBars      =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8100
         TabIndex        =   27
         Top             =   3540
         Width           =   1275
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1260
         TabIndex        =   26
         Top             =   3540
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   3540
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSBien 
         Height          =   1575
         Left            =   0
         TabIndex        =   24
         Top             =   1920
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame fraReg 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   6900
         TabIndex        =   23
         Top             =   3540
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8220
         TabIndex        =   22
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos de la Adjudicación "
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
         Height          =   2235
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   9375
         Begin VB.TextBox txtDireccion 
            Height          =   315
            Left            =   1200
            TabIndex        =   28
            Top             =   1380
            Width           =   7935
         End
         Begin VB.ComboBox cboTipoBien 
            Height          =   315
            ItemData        =   "frmLogIngAdjudicados.frx":08CA
            Left            =   1200
            List            =   "frmLogIngAdjudicados.frx":08D4
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   360
            Width           =   2115
         End
         Begin VB.ComboBox cboClaseBien 
            Height          =   315
            ItemData        =   "frmLogIngAdjudicados.frx":08EA
            Left            =   8280
            List            =   "frmLogIngAdjudicados.frx":08F4
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboOrigen 
            Height          =   315
            ItemData        =   "frmLogIngAdjudicados.frx":0900
            Left            =   5160
            List            =   "frmLogIngAdjudicados.frx":090A
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   360
            Width           =   1995
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   615
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   720
            Width           =   7935
         End
         Begin VB.TextBox txtUbiGeoDesc 
            Height          =   315
            Left            =   2460
            TabIndex        =   12
            Top             =   1740
            Width           =   6675
         End
         Begin VB.CommandButton cmdUbiGeo 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2040
            TabIndex        =   10
            Top             =   1770
            Width           =   405
         End
         Begin VB.TextBox txtUbiGeoCod 
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   1740
            Width           =   1275
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Bien"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   420
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   780
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Localización"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   1800
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7260
            TabIndex        =   18
            Top             =   420
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Origen"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4560
            TabIndex        =   17
            Top             =   420
            Width           =   585
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   1635
         Left            =   0
         TabIndex        =   8
         Top             =   2280
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   4
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
End
Attribute VB_Name = "frmLogIngAdjudicados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAgregar_Click()
fraVis.Visible = False
fraReg.Visible = True
cboTipoBien.ListIndex = 0
cboOrigen.ListIndex = 0
cboClaseBien.ListIndex = 0
txtDescripcion.Text = ""
txtDireccion.Text = ""
txtUbiGeoCod.Text = ""
txtUbiGeoDesc.Text = ""
End Sub

Private Sub cmdCancelar_Click()
fraVis.Visible = True
fraReg.Visible = False
End Sub

Private Sub cmdGrabar_Click()
fraVis.Visible = True
fraReg.Visible = False
End Sub

Private Sub cmdPersona_Click()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio

If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   txtPersona.Text = X.sPersNombre
   txtPersCod = X.sPersCod
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdUbiGeo_Click()
frmLogProSelSeleUbiGeo.Show 1
If Len(Trim(frmLogProSelSeleUbiGeo.vpCodUbigeo)) = 0 Then
   txtUbiGeoCod.Text = frmLogProSelSeleUbiGeo.vpCodUbigeo
   txtUbiGeoDesc.Text = frmLogProSelSeleUbiGeo.vpUbigeoDesc
End If
End Sub

Private Sub Form_Load()
CentraForm Me
FormaFlex
End Sub


Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 4
MSFlex.ColWidth(0) = 0
MSFlex.RowHeight(0) = 340
MSFlex.RowHeight(1) = 300
MSFlex.RowHeight(2) = 300
MSFlex.RowHeight(3) = 300

MSFlex.TextMatrix(0, 1) = "Descripción": MSFlex.ColWidth(1) = 2000
MSFlex.TextMatrix(0, 2) = "Fecha":       MSFlex.ColWidth(2) = 1000
MSFlex.TextMatrix(0, 3) = "Moneda":      MSFlex.ColWidth(3) = 600
MSFlex.TextMatrix(0, 4) = "Monto":       MSFlex.ColWidth(4) = 1000
MSFlex.TextMatrix(1, 1) = "Valor de Adjudicación"
MSFlex.TextMatrix(2, 1) = "Valor de Realización"
MSFlex.TextMatrix(3, 1) = "Valor Ultima Subasta"
MSFlex.Col = 1:  MSFlex.row = 1:  MSFlex.CellFontBold = True ': MSFlex.CellBackColor = ""
MSFlex.Col = 1:  MSFlex.row = 2:  MSFlex.CellFontBold = True ': MSFlex.CellBackColor = ""
MSFlex.Col = 1:  MSFlex.row = 3:  MSFlex.CellFontBold = True ': MSFlex.CellBackColor = ""
End Sub

Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub
