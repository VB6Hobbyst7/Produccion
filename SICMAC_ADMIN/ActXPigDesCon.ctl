VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ActXPigDesCon 
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ScaleHeight     =   2010
   ScaleWidth      =   7500
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   15
      TabIndex        =   22
      Top             =   990
      Width           =   7440
      Begin VB.Label lblFechaPrestamo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6345
         TabIndex        =   34
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lblFechaVencimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3495
         TabIndex        =   33
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Plazo"
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   32
         Top             =   225
         Width           =   615
      End
      Begin VB.Label lblPiezas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1035
         TabIndex        =   31
         Top             =   180
         Width           =   855
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fec. Prestamo"
         Height          =   165
         Index           =   3
         Left            =   5010
         TabIndex        =   30
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fec.Vencim."
         Height          =   165
         Index           =   4
         Left            =   2280
         TabIndex        =   29
         Top             =   255
         Width           =   915
      End
      Begin VB.Label lblneto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6360
         TabIndex        =   28
         Top             =   555
         Width           =   975
      End
      Begin VB.Label lblcomtas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3495
         TabIndex        =   27
         Top             =   555
         Width           =   975
      End
      Begin VB.Label lblPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1035
         TabIndex        =   26
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Prestamo "
         Height          =   180
         Left            =   105
         TabIndex        =   25
         Top             =   585
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Com. Tasacion"
         Height          =   195
         Left            =   2265
         TabIndex        =   24
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Neto a Recibir "
         Height          =   255
         Left            =   5010
         TabIndex        =   23
         Top             =   600
         Width           =   1275
      End
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Descripcion Lote"
      Height          =   1110
      Index           =   3
      Left            =   420
      TabIndex        =   13
      Top             =   5460
      Width           =   7455
      Begin VB.TextBox txtDescLote 
         Enabled         =   0   'False
         Height          =   795
         Left            =   180
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   690
         Width           =   7245
      End
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Kilataje"
      Height          =   1785
      Index           =   5
      Left            =   6390
      TabIndex        =   4
      Top             =   5400
      Width           =   1500
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
         TabIndex        =   8
         Top             =   1035
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
         TabIndex        =   7
         Top             =   740
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
         TabIndex        =   6
         Top             =   445
         Width           =   720
      End
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
         TabIndex        =   5
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "14 K"
         Height          =   210
         Index           =   14
         Left            =   255
         TabIndex        =   12
         Top             =   555
         Width           =   495
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "16 K"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   11
         Top             =   510
         Width           =   420
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "18 K"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   10
         Top             =   795
         Width           =   495
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "21 K"
         Height          =   210
         Index           =   21
         Left            =   150
         TabIndex        =   9
         Top             =   1065
         Width           =   465
      End
   End
   Begin VB.Frame fraContenedor 
      Height          =   1005
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   7440
      Begin MSComctlLib.ListView lstClientes 
         Height          =   735
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   1296
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
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
      Begin VB.TextBox txtTipCta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6015
         TabIndex        =   2
         Top             =   525
         Width           =   780
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tipo de Contrato "
         Height          =   405
         Index           =   1
         Left            =   6285
         TabIndex        =   3
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Label lblestado 
      Caption         =   "Label1"
      Height          =   465
      Left            =   5775
      TabIndex        =   21
      Top             =   7230
      Width           =   1140
   End
   Begin VB.Label lblcodestado 
      Caption         =   "Label1"
      Height          =   555
      Left            =   900
      TabIndex        =   20
      Top             =   6900
      Width           =   1680
   End
   Begin VB.Label lblMontoPrestamo 
      Caption         =   "Label1"
      Height          =   1185
      Left            =   8115
      TabIndex        =   19
      Top             =   5475
      Width           =   435
   End
   Begin VB.Label lblvalortasacion 
      Caption         =   "Label1"
      Height          =   450
      Left            =   4575
      TabIndex        =   18
      Top             =   5100
      Width           =   1560
   End
   Begin VB.Label lblsaldocapital 
      Caption         =   "Label1"
      Height          =   420
      Left            =   2505
      TabIndex        =   17
      Top             =   4935
      Width           =   960
   End
   Begin VB.Label lbloroneto 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3165
      TabIndex        =   16
      Top             =   6945
      Width           =   1815
   End
   Begin VB.Label lblorobruto 
      Caption         =   "Label1"
      Height          =   405
      Left            =   1905
      TabIndex        =   15
      Top             =   6330
      Width           =   2205
   End
End
Attribute VB_Name = "ActXPigDesCon"
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

'TipoCuenta - txttipcta.tex
Public Property Get TipoCuenta() As String
    TipoCuenta = txtTipCta.Text
End Property
Public Property Let TipoCuenta(ByVal NewTipoCta As String)
    txtTipCta.Text = NewTipoCta
    PropertyChanged "TipoCta"
End Property
'OroBruto - lblOroBruto.text
Public Property Get OroBruto() As Double
    OroBruto = Val(lblorobruto.Caption)
End Property
Public Property Let OroBruto(ByVal NewOroBruto As Double)
    lblorobruto.Caption = Format(NewOroBruto, "#0.00")
    PropertyChanged "OroBruto"
End Property
'OroNeto - lblOroNeto.caption
Public Property Get OroNeto() As Double
    OroNeto = Val(lbloroneto.Caption)
End Property
Public Property Let OroNeto(ByVal NewOroNeto As Double)
    lbloroneto.Caption = Format(NewOroNeto, "#0.00")
    PropertyChanged "OroNeto"
End Property
'Piezas - lblpiezas.caption
Public Property Get Piezas() As Integer
    Piezas = Val(lblPiezas.Caption)
End Property
Public Property Let Piezas(ByVal NewPiezas As Integer)
    lblPiezas.Caption = Format(NewPiezas, "#0")
    PropertyChanged "Piezas"
End Property

'SaldoCapital - lblSaldoCapital.caption
Public Property Get SaldoCapital() As Currency
    SaldoCapital = Val(lblsaldocapital.Caption)
End Property
Public Property Let SaldoCapital(ByVal NewSaldoCapital As Currency)
    lblsaldocapital.Caption = Format(NewSaldoCapital, "#0.00")
    PropertyChanged "SaldoCapital"
End Property

'ValTasa - lblValorTasacion.caption
Public Property Get ValTasa() As Currency
    ValTasa = Val(lblvalortasacion.Caption)
End Property
Public Property Let ValTasa(ByVal NewValTasa As Currency)
    lblvalortasacion.Caption = Format(NewValTasa, "#0.00")
    PropertyChanged "ValTasa"
End Property

'MonPres - lblMontoPrestamo.caption
Public Property Get prestamo() As Currency
    prestamo = Val(lblMontoPrestamo.Caption)
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
    CodEstadoCred = Val(lblcodestado.Caption)
End Property
Public Property Let CodEstadoCred(ByVal NewCodEstadoCred As Integer)
    lblcodestado.Caption = Trim(Str(NewCodEstadoCred))
    PropertyChanged "CodEstadoCred"
End Property

'Estado Credito  - lblEstado.Caption
Public Property Get EstadoCred() As String
    EstadoCred = lblestado.Caption
End Property
Public Property Let EstadoCred(ByVal NewEstadoCred As String)
    lblestado.Caption = fgEstadoCredPigDesc(NewEstadoCred)
    PropertyChanged "EstadoCred"
End Property

'Oro14 - txt14k.text
Public Property Get Oro14() As Double
    Oro14 = Val(txt14k.Text)
End Property
Public Property Let Oro14(ByVal NewOro14 As Double)
    txt14k.Text = Format(NewOro14, "#0.00")
    PropertyChanged "Oro14"
End Property

'Oro16 - txt16k.text
Public Property Get Oro16() As Double
    Oro16 = Val(txt16k.Text)
End Property
Public Property Let Oro16(ByVal NewOro16 As Double)
    txt16k.Text = Format(NewOro16, "#0.00")
    PropertyChanged "Oro16"
End Property

'Oro18 - txt18k.text
Public Property Get Oro18() As Double
    Oro18 = Val(txt18k.Text)
End Property
Public Property Let Oro18(ByVal NewOro18 As Double)
    txt18k.Text = Format(NewOro18, "#0.00")
    PropertyChanged "Oro18"
End Property
'CMCPL - PRESTAMO
'Prestamo - lblprestamo
Public Property Get prestamo1() As Double
    prestamo1 = Val(lblPrestamo.Caption)
End Property
Public Property Let prestamo1(ByVal Newprestamo As Double)
    lblPrestamo.Caption = Format(Newprestamo, "#0.00")
    PropertyChanged "prestamo1"
End Property
'CMCPL - COMISION
'Prestamo - lblcontas
Public Property Get comision1() As Double
    comision1 = Val(lblcomtas.Caption)
End Property
Public Property Let comision1(ByVal Newcomision As Double)
    lblcomtas.Caption = Format(Newcomision, "#0.00")
    PropertyChanged "comision1"
End Property
'CMCPL -NETO
'Prestamo - lblneto
Public Property Get neto1() As Double
    neto1 = Val(lblneto.Caption)
End Property
Public Property Let neto1(ByVal Newneto As Double)
    lblneto.Caption = Format(Newneto, "#0.00")
    PropertyChanged "neto1"
End Property
'Oro21 - txt21k.text
Public Property Get Oro21() As Double
    Oro21 = Val(txt21k.Text)
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
    lblorobruto.Caption = Format(0, "0.00")
    lbloroneto.Caption = Format(0, "0.00")
    lblPiezas.Caption = 0
    '***********************************
    lblPrestamo.Caption = Format(0, "0.00")
    lblcomtas.Caption = Format(0, "0.00")
    lblneto.Caption = Format(0, "0.00")
    '***********************************
    lblsaldocapital.Caption = Format(0, "0.00")
    lblFechaPrestamo.Caption = "  /  /  "
    lblFechaVencimiento.Caption = "  /  /  "
    lblestado.Caption = ""
    lblcodestado.Caption = ""
End Sub

