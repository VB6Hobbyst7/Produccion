VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl ActXColPDesCon 
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   ScaleHeight     =   3555
   ScaleWidth      =   7560
   ToolboxBitmap   =   "ActXColPDesCon.ctx":0000
   Begin VB.Frame fraContenedor 
      Height          =   1005
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7515
      Begin MSComctlLib.ListView lstClientes 
         Height          =   735
         Left            =   60
         TabIndex        =   34
         Top             =   180
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   1296
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Direccion"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Telefono"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cod Zona"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cod Ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TidoCi"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Doc Ident."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "TidoTr"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Nro Doc Trib"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.TextBox txtTEM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         TabIndex        =   15
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblEtiqueta 
         Alignment       =   2  'Center
         Caption         =   "Tasa Efectiva Mensual"
         Height          =   405
         Index           =   1
         Left            =   6360
         TabIndex        =   16
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Kilataje"
      Height          =   1350
      Index           =   5
      Left            =   6120
      TabIndex        =   0
      Top             =   1020
      Width           =   1335
      Begin VB.TextBox txt14k 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   555
         TabIndex        =   32
         Top             =   180
         Width           =   720
      End
      Begin VB.TextBox txt16k 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   555
         TabIndex        =   31
         Top             =   445
         Width           =   720
      End
      Begin VB.TextBox txt18k 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   555
         TabIndex        =   30
         Top             =   740
         Width           =   720
      End
      Begin VB.TextBox txt21k 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   555
         TabIndex        =   29
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "21 K"
         Height          =   210
         Index           =   21
         Left            =   150
         TabIndex        =   4
         Top             =   1065
         Width           =   465
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "18 K"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   3
         Top             =   795
         Width           =   495
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "16 K"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   2
         Top             =   510
         Width           =   420
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "14 K"
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   495
      End
   End
   Begin VB.Frame fraContenedor 
      Height          =   1350
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   1020
      Width           =   6135
      Begin VB.Label lblCodEstado 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3960
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   4560
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Estado "
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblFechaVencimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fec.Vencim."
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblSaldoCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblFechaPrestamo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblMontoPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fec. Prestamo"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblValorTasacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPiezas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblOroNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblOroBruto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Oro Bruto(gr)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Oro Neto(gr)"
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Piezas "
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Saldo Cap."
         Height          =   255
         Index           =   13
         Left            =   2040
         TabIndex        =   8
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tasación"
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Prestamo"
         Height          =   255
         Index           =   12
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Descripcion Lote"
      Height          =   1110
      Index           =   3
      Left            =   0
      TabIndex        =   12
      Top             =   2520
      Width           =   7455
      Begin ComctlLib.Slider Slider1 
         Height          =   30
         Left            =   1320
         TabIndex        =   36
         Top             =   840
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
         _Version        =   327682
      End
      Begin MSComctlLib.ListView lstJoyasDet 
         Height          =   855
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Piezas"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Oro Bruto"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Oro Neto"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Val. Tasac"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Descripcion"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.TextBox txtDescLote 
         Enabled         =   0   'False
         Height          =   795
         Left            =   90
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   210
         Width           =   7245
      End
   End
End
Attribute VB_Name = "ActXColPDesCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event KeyPressDesLot(KeyAscii As Integer)

'ListaClientes - lstClientes
Public Property Get listaClientes() As ListView
    Set listaClientes = lstClientes
    
End Property
Public Property Let listaClientes(ByVal NewlstClientes As ListView)
    Set lstClientes = NewlstClientes
    PropertyChanged "ListaClientes"
End Property

'RECO20140410 ERS044-2014************************************
'TipoCuenta - txttipcta.tex
'Public Property Get TipoCuenta() As String
'    TipoCuenta = txtTipCta.Text
'End Property
'Public Property Let TipoCuenta(ByVal NewTipoCta As String)
'    txtTipCta.Text = NewTipoCta
'    PropertyChanged "TipoCta"
'End Property
'RECO FIN****************************************************

'OroBruto - lblOroBruto.text
Public Property Get OroBruto() As Double
    OroBruto = val(lblOroBruto.Caption)
End Property
Public Property Let OroBruto(ByVal NewOroBruto As Double)
    lblOroBruto.Caption = Format(NewOroBruto, "#0.00")
    PropertyChanged "OroBruto"
End Property

'OroNeto - lblOroNeto.caption
Public Property Get OroNeto() As Double
    OroNeto = val(lblOroNeto.Caption)
End Property
Public Property Let OroNeto(ByVal NewOroNeto As Double)
    lblOroNeto.Caption = Format(NewOroNeto, "#0.00")
    PropertyChanged "OroNeto"
End Property

'Piezas - lblpiezas.caption
Public Property Get Piezas() As Integer
    Piezas = val(lblPiezas.Caption)
End Property
Public Property Let Piezas(ByVal NewPiezas As Integer)
    lblPiezas.Caption = Format(NewPiezas, "#0")
    PropertyChanged "Piezas"
End Property

'SaldoCapital - lblSaldoCapital.caption
Public Property Get SaldoCapital() As Currency
    SaldoCapital = val(lblSaldoCapital.Caption)
End Property
Public Property Let SaldoCapital(ByVal NewSaldoCapital As Currency)
    lblSaldoCapital.Caption = Format(NewSaldoCapital, "#0.00")
    PropertyChanged "SaldoCapital"
End Property

'ValTasa - lblValorTasacion.caption
Public Property Get ValTasa() As Currency
    ValTasa = val(lblValorTasacion.Caption)
End Property
Public Property Let ValTasa(ByVal NewValTasa As Currency)
    lblValorTasacion.Caption = Format(NewValTasa, "#0.00")
    PropertyChanged "ValTasa"
End Property

'MonPres - lblMontoPrestamo.caption
Public Property Get prestamo() As Currency
    prestamo = val(lblMontoPrestamo.Caption)
End Property
Public Property Let prestamo(ByVal Newprestamo As Currency)
    lblMontoPrestamo.Caption = Format(Newprestamo, "#0.00")
    PropertyChanged "Prestamo"
End Property

'Fecha Prestamo - lblFechaPrestamo.Caption
Public Property Get FechaPrestamo() As String
    FechaPrestamo = Format(lblFechaPrestamo.Caption, "dd/mm/yyyy")
End Property
Public Property Let FechaPrestamo(ByVal NewFechaPrestamo As String)
    lblFechaPrestamo.Caption = Format(NewFechaPrestamo, "dd/mm/yyyy")
    PropertyChanged "FechaPrestamo"
End Property

'Fecha Vencimiento - lblFechaVencimiento.Caption
Public Property Get FechaVencimiento() As String
    FechaVencimiento = Format(lblFechaVencimiento.Caption, "dd/mm/yyyy")
End Property
Public Property Let FechaVencimiento(ByVal NewFechaVencimiento As String)
    lblFechaVencimiento.Caption = Format(NewFechaVencimiento, "dd/mm/yyyy")
    PropertyChanged "FechaVencimiento"
End Property

'Codig Estado Credito  - lblCodEstado.Caption
Public Property Get CodEstadoCred() As Integer
    CodEstadoCred = val(lblcodestado.Caption)
End Property
Public Property Let CodEstadoCred(ByVal NewCodEstadoCred As Integer)
    lblcodestado.Caption = Trim(str(NewCodEstadoCred))
    PropertyChanged "CodEstadoCred"
End Property

'Estado Credito  - lblEstado.Caption
Public Property Get EstadoCred() As String
    EstadoCred = lblEstado.Caption
End Property
Public Property Let EstadoCred(ByVal NewEstadoCred As String)
    lblEstado.Caption = fgEstadoCredPigDesc(NewEstadoCred)
    PropertyChanged "EstadoCred"
End Property

'Oro14 - txt14k.text
Public Property Get Oro14() As Double
    Oro14 = val(txt14k.Text)
End Property
Public Property Let Oro14(ByVal NewOro14 As Double)
    txt14k.Text = Format(NewOro14, "#0.00")
    PropertyChanged "Oro14"
End Property

'Oro16 - txt16k.text
Public Property Get Oro16() As Double
    Oro16 = val(txt16k.Text)
End Property
Public Property Let Oro16(ByVal NewOro16 As Double)
    txt16k.Text = Format(NewOro16, "#0.00")
    PropertyChanged "Oro16"
End Property

'Oro18 - txt18k.text
Public Property Get Oro18() As Double
    Oro18 = val(txt18k.Text)
End Property
Public Property Let Oro18(ByVal NewOro18 As Double)
    txt18k.Text = Format(NewOro18, "#0.00")
    PropertyChanged "Oro18"
End Property

'Oro21 - txt21k.text
Public Property Get Oro21() As Double
    Oro21 = val(txt21k.Text)
End Property
Public Property Let Oro21(ByVal NewOro21 As Double)
    txt21k.Text = Format(NewOro21, "#0.00")
    PropertyChanged "Oro21"
End Property

'DescLote - txtDescLote.text
Public Property Get DescLote() As String
    DescLote = txtDescLote.Text
End Property
Public Property Let DescLote(ByVal NewDescLote As String)
    txtDescLote.Text = NewDescLote
    PropertyChanged "DescLote"
End Property
Public Property Get EnabledDescLot() As Boolean
    EnabledDescLot = txtDescLote.Enabled
End Property
Public Property Let EnabledDescLot(ByVal NewEnabledDescLot As Boolean)
    txtDescLote.Enabled = NewEnabledDescLot
    PropertyChanged "EnabledDescLot"
End Property
Public Sub SetFocusDesLot()
    txtDescLote.SetFocus
End Sub

Private Sub txtDescLote_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPressDesLot(KeyAscii)
End Sub

'Limpia Controles
Public Sub Limpiar()
    lstClientes.ListItems.Clear
    lblOroBruto.Caption = Format(0, "0.00")
    lblOroNeto.Caption = Format(0, "0.00")
    lblPiezas.Caption = 0
    lblValorTasacion.Caption = Format(0, "0.00")
    lblMontoPrestamo.Caption = Format(0, "0.00")
    lblSaldoCapital.Caption = Format(0, "0.00")
    lblFechaPrestamo.Caption = "  /  /  "
    lblFechaVencimiento.Caption = "  /  /  "
    lblEstado.Caption = ""
    lblcodestado.Caption = ""
    txt14k.Text = Format(0, "0.00")
    txt16k.Text = Format(0, "0.00")
    txt18k.Text = Format(0, "0.00")
    txt21k.Text = Format(0, "0.00")
    txtDescLote.Text = ""
    lstJoyasDet.ListItems.Clear
    txtTEM.Text = "" ' RECO20140410 ERS044
End Sub

'ListaClientes - lstClientes
Public Property Get listaJoyasDet() As ListView
    Set listaJoyasDet = lstJoyasDet
    
End Property
Public Property Let listaJoyasDet(ByVal NewlstJoyasDet As ListView)
    Dim lstJ As ListView
    Set lstJ = NewlstJoyasDet
    PropertyChanged "ListaJoyasDet"
End Property

'RECO20140410 ERS044-2014************************************
'Tasa efectiva Mensual - txtTEM
Public Property Get TasaEfectivaMensual() As String
    TasaEfectivaMensual = txtTEM.Text
End Property
Public Property Let TasaEfectivaMensual(ByVal NewTasaEfectivaMensual As String)
    txtTEM.Text = NewTasaEfectivaMensual
    PropertyChanged "TasaEfectivaMensual"
End Property
'RECO FIN****************************************************
