VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ActXPigComision 
   BackColor       =   &H8000000B&
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   ScaleHeight     =   2085
   ScaleWidth      =   7470
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   30
      TabIndex        =   4
      Top             =   990
      Width           =   7410
      Begin VB.TextBox txtDuplicadp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3870
         TabIndex        =   16
         Top             =   615
         Width           =   600
      End
      Begin VB.TextBox txtmtotas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1110
         TabIndex        =   14
         Top             =   615
         Width           =   1095
      End
      Begin VB.TextBox txtsalpag 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   6240
         TabIndex        =   9
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Duplicados"
         Height          =   225
         Left            =   2505
         TabIndex        =   15
         Top             =   705
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Mto.Tasacion"
         Height          =   255
         Left            =   75
         TabIndex        =   13
         Top             =   705
         Width           =   1005
      End
      Begin VB.Label txtctocus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   6225
         TabIndex        =   12
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label txtdiastra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   3870
         TabIndex        =   11
         Top             =   255
         Width           =   600
      End
      Begin VB.Label txtfecpag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1110
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label label4 
         Caption         =   "Disponible"
         ForeColor       =   &H80000001&
         Height          =   225
         Left            =   4860
         TabIndex        =   8
         Top             =   675
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Costo Custodia"
         Height          =   180
         Left            =   4845
         TabIndex        =   7
         Top             =   345
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Dias Trancurrido"
         Height          =   180
         Left            =   2505
         TabIndex        =   6
         Top             =   345
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pago"
         Height          =   195
         Left            =   75
         TabIndex        =   5
         Top             =   330
         Width           =   870
      End
   End
   Begin VB.Frame fraContenedor 
      Height          =   1005
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7440
      Begin VB.TextBox txtTipCta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6015
         TabIndex        =   2
         Top             =   1065
         Width           =   780
      End
      Begin MSComctlLib.ListView lstClientes 
         Height          =   735
         Left            =   60
         TabIndex        =   1
         Top             =   135
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
End
Attribute VB_Name = "ActXPigComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ListaClientes - lstClientes
Public Property Get listaClientes() As ListView
    Set listaClientes = lstClientes
End Property
Public Property Let listaClientes(ByVal NewlstClientes As ListView)
    Set lstClientes = NewlstClientes
    PropertyChanged "ListaClientes"
End Property
Public Property Get Oro14() As Double
    Oro14 = Val(txt14k.Text)
End Property
Public Property Let Oro14(ByVal NewOro14 As Double)
    txt14k.Text = Format(NewOro14, "#0.00")
    PropertyChanged "Oro14"
End Property
'Fecha de Pago - txtfecpag
Public Property Get Fechapago() As String
    Fechapago = txtfecpag.Caption
    End Property
Public Property Let Fechapago(ByVal Newfecha As String)
    txtfecpag.Caption = Format(Newfecha, "Short Date")
    PropertyChanged "FechaPago"
End Property
'Dias de Atraso - txtdiastra
Public Property Get DiasAtraso() As String
    DiasAtraso = txtdiastra.Caption
End Property
Public Property Let DiasAtraso(ByVal NewDia As String)
    txtdiastra.Caption = NewDia
    PropertyChanged "DiasAtraso"
End Property
'Dias de Costo de Cuota  - txtctocus
Public Property Get CuotaCosto() As String
    CuotaCosto = txtctocus.Caption
End Property
Public Property Let CuotaCosto(ByVal NewCosto As String)
    txtctocus.Caption = NewCosto
    PropertyChanged "CuotaCosto"
End Property
Public Sub Limpiar()
    lstClientes.ListItems.Clear
    '***********************************
    txtfecpag.Caption = "  /  /  "
    txtdiastra.Caption = Format(0, "0")
     txtctocus.Caption = Format(0, "0,00")
 End Sub
