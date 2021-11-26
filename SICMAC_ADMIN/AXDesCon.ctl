VERSION 5.00
Begin VB.UserControl AXDesCon 
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   ScaleHeight     =   2535
   ScaleWidth      =   7425
   ToolboxBitmap   =   "AXDesCon.ctx":0000
   Begin VB.Frame fraContenedor 
      Caption         =   "Kilataje"
      Height          =   1410
      Index           =   5
      Left            =   5910
      TabIndex        =   0
      Top             =   -15
      Width           =   1470
      Begin VB.TextBox txt21k 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   4
         Top             =   1035
         Width           =   750
      End
      Begin VB.TextBox txt18k 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   3
         Top             =   750
         Width           =   750
      End
      Begin VB.TextBox txt16k 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   2
         Top             =   465
         Width           =   750
      End
      Begin VB.TextBox txt14k 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   1
         Top             =   180
         Width           =   750
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "21 K"
         Height          =   210
         Index           =   21
         Left            =   150
         TabIndex        =   8
         Top             =   1065
         Width           =   465
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "18 K"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   7
         Top             =   795
         Width           =   495
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "16 K"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   6
         Top             =   510
         Width           =   420
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "14 K"
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   495
      End
   End
   Begin VB.Frame fraContenedor 
      Height          =   1410
      Index           =   2
      Left            =   30
      TabIndex        =   9
      Top             =   -15
      Width           =   5790
      Begin VB.TextBox txtOroBruto 
         Alignment       =   1  'Right Justify
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
         Left            =   1575
         TabIndex        =   15
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txtOroNeto 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   4395
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   255
         Width           =   1110
      End
      Begin VB.TextBox txtPiezas 
         Alignment       =   1  'Right Justify
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
         Left            =   1575
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtValorTasacion 
         Alignment       =   1  'Right Justify
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
         Left            =   1575
         TabIndex        =   12
         Top             =   945
         Width           =   1095
      End
      Begin VB.TextBox txtMontoPrestamo 
         Alignment       =   1  'Right Justify
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
         Left            =   4395
         TabIndex        =   11
         Top             =   945
         Width           =   1110
      End
      Begin VB.TextBox txtPlazo 
         Alignment       =   1  'Right Justify
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
         Left            =   4395
         TabIndex        =   10
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Oro Bruto : (gr)"
         Height          =   255
         Index           =   0
         Left            =   255
         TabIndex        =   21
         Top             =   285
         Width           =   1185
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Oro Neto : (gr)"
         Height          =   210
         Index           =   11
         Left            =   3015
         TabIndex        =   20
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Piezas :"
         Height          =   210
         Index           =   2
         Left            =   255
         TabIndex        =   19
         Top             =   615
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Plazo : (dias)"
         Height          =   255
         Index           =   13
         Left            =   3015
         TabIndex        =   18
         Top             =   615
         Width           =   1155
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Valor Tasación :"
         Height          =   255
         Index           =   9
         Left            =   255
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Monto Prestamo :"
         Height          =   255
         Index           =   12
         Left            =   3015
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Descripcion Lote"
      Height          =   1110
      Index           =   3
      Left            =   30
      TabIndex        =   22
      Top             =   1380
      Width           =   7350
      Begin VB.TextBox txtDescLote 
         Enabled         =   0   'False
         Height          =   795
         Left            =   150
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   210
         Width           =   7020
      End
   End
End
Attribute VB_Name = "AXDesCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event KeyPressDesLot(KeyAscii As Integer)

'OroBruto - txtOroBruto.text
Public Property Get OroBruto() As Double
    OroBruto = Val(txtOroBruto.Text)
End Property
Public Property Let OroBruto(ByVal NewOroBruto As Double)
    txtOroBruto.Text = Format(NewOroBruto, "#0.00")
    PropertyChanged "OroBruto"
End Property

'OroNeto - txtOroNeto.text
Public Property Get OroNeto() As Double
    OroNeto = Val(txtOroNeto.Text)
End Property
Public Property Let OroNeto(ByVal NewOroNeto As Double)
    txtOroNeto.Text = Format(NewOroNeto, "#0.00")
    PropertyChanged "OroNeto"
End Property

'Piezas - txtPiezas.text
Public Property Get Piezas() As Integer
    Piezas = Val(txtPiezas.Text)
End Property
Public Property Let Piezas(ByVal NewPiezas As Integer)
    txtPiezas.Text = Format(NewPiezas, "#0.00")
    PropertyChanged "Piezas"
End Property

'Plazo - txtPlazo.text
Public Property Get Plazo() As Integer
    Plazo = Val(txtPlazo.Text)
End Property
Public Property Let Plazo(ByVal NewPlazo As Integer)
    txtPlazo.Text = Format(NewPlazo, "#0.00")
    PropertyChanged "Plazo"
End Property

'ValTasa - txtValorTasacion.text
Public Property Get ValTasa() As Double
    ValTasa = Val(txtvalorTasacion.Text)
End Property
Public Property Let ValTasa(ByVal NewValTasa As Double)
    txtvalorTasacion.Text = Format(NewValTasa, "#0.00")
    PropertyChanged "ValTasa"
End Property

'MonPres - txtMontoPrestamo.text
Public Property Get Prestamo() As Double
    Prestamo = Val(txtMontoPrestamo.Text)
End Property
Public Property Let Prestamo(ByVal NewPrestamo As Double)
    txtMontoPrestamo.Text = Format(NewPrestamo, "#0.00")
    PropertyChanged "Prestamo"
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
Public Sub Clear()
txtOroBruto.Text = Format(0, "0.00")
txtOroNeto.Text = Format(0, "0.00")
txtPiezas.Text = 0
txtPlazo = 0
txtvalorTasacion.Text = Format(0, "0.00")
txtMontoPrestamo.Text = Format(0, "0.00")
txt14k.Text = Format(0, "0.00")
txt16k.Text = Format(0, "0.00")
txt18k.Text = Format(0, "0.00")
txt21k.Text = Format(0, "0.00")
txtDescLote.Text = ""
End Sub

