VERSION 5.00
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
      Begin VB.TextBox txtTipCta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   15
         Top             =   660
         Width           =   780
      End
      Begin VB.PictureBox lstCliente 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   60
         ScaleHeight     =   780
         ScaleWidth      =   6435
         TabIndex        =   16
         Top             =   120
         Width           =   6495
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tipo de Contrato "
         Height          =   405
         Index           =   1
         Left            =   6600
         TabIndex        =   17
         Top             =   120
         Width           =   705
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Estado "
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fec.Vencim."
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fec. Prestamo"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
      Left            =   30
      TabIndex        =   12
      Top             =   2400
      Width           =   7455
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
Public Property Get ListaClientes() As ListView
    Set ListaClientes = lstCliente
    
End Property
Public Property Let ListaClientes(ByVal NewlstClientes As ListView)
    Set Me.lstClientes = NewlstClientes
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
    OroBruto = Val(lblOroBruto.Caption)
End Property
Public Property Let OroBruto(ByVal NewOroBruto As Double)
    lblOroBruto.Caption = Format(NewOroBruto, "#0.00")
    PropertyChanged "OroBruto"
End Property

'OroNeto - lblOroNeto.caption
Public Property Get OroNeto() As Double
    OroNeto = Val(lblOroNeto.Caption)
End Property
Public Property Let OroNeto(ByVal NewOroNeto As Double)
    lblOroNeto.Caption = Format(NewOroNeto, "#0.00")
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
    SaldoCapital = Val(lblSaldoCapital.Caption)
End Property
Public Property Let SaldoCapital(ByVal NewSaldoCapital As Currency)
    lblSaldoCapital.Caption = Format(NewSaldoCapital, "#0.00")
    PropertyChanged "SaldoCapital"
End Property

'ValTasa - lblValorTasacion.caption
Public Property Get ValTasa() As Currency
    ValTasa = Val(lblValorTasacion.Caption)
End Property
Public Property Let ValTasa(ByVal NewValTasa As Currency)
    lblValorTasacion.Caption = Format(NewValTasa, "#0.00")
    PropertyChanged "ValTasa"
End Property

'MonPres - lblMontoPrestamo.caption
Public Property Get Prestamo() As Currency
    Prestamo = Val(lblMontoPrestamo.Caption)
End Property
Public Property Let Prestamo(ByVal NewPrestamo As Currency)
    lblMontoPrestamo.Caption = Format(NewPrestamo, "#0.00")
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
    CodEstadoCred = Val(lblCodEstado.Caption)
End Property
Public Property Let CodEstadoCred(ByVal NewCodEstadoCred As Integer)
    lblCodEstado.Caption = Trim(Str(NewCodEstadoCred))
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
    lstCliente.ListItems.Clear
    lblOroBruto.Caption = Format(0, "0.00")
    lblOroNeto.Caption = Format(0, "0.00")
    lblPiezas.Caption = 0
    lblValorTasacion.Caption = Format(0, "0.00")
    lblMontoPrestamo.Caption = Format(0, "0.00")
    lblSaldoCapital.Caption = Format(0, "0.00")
    lblFechaPrestamo.Caption = "  /  /  "
    lblFechaVencimiento.Caption = "  /  /  "
    lblEstado.Caption = ""
    lblCodEstado.Caption = ""
    txt14k.Text = Format(0, "0.00")
    txt16k.Text = Format(0, "0.00")
    txt18k.Text = Format(0, "0.00")
    txt21k.Text = Format(0, "0.00")
    txtDescLote.Text = ""
End Sub

