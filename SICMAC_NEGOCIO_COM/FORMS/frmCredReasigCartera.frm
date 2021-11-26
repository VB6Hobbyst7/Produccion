VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredReasigCartera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reasignacion de Cartera en Lote"
   ClientHeight    =   7710
   ClientLeft      =   765
   ClientTop       =   2505
   ClientWidth     =   10605
   ClipControls    =   0   'False
   Icon            =   "frmCredReasigCartera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTransferencia 
      Caption         =   "Total Transferido"
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
      Height          =   1815
      Left            =   5760
      TabIndex        =   42
      Top             =   5280
      Width           =   4695
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Nº Créditos:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Mora 8-30:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Mora >30:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Mora 8-15:"
         Height          =   195
         Left            =   2400
         TabIndex        =   52
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Mora >15:"
         Height          =   195
         Left            =   2400
         TabIndex        =   51
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "(Sólo minorista)"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "(Sólo no minorista)"
         Height          =   195
         Left            =   2400
         TabIndex        =   49
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label lblNCreditosT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   48
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblSaldoCapT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   47
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label lblMora830T 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   46
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblMoraM30T 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   45
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblMora815T 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   3360
         TabIndex        =   44
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblMoraM15T 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   3360
         TabIndex        =   43
         Top             =   1140
         Width           =   1155
      End
   End
   Begin VB.Frame fraOrigen 
      Caption         =   "Total Origen"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   27
      Top             =   5280
      Width           =   4695
      Begin VB.Label lblMoraM15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   3360
         TabIndex        =   41
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblMora815 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   3360
         TabIndex        =   40
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblMoraM30 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblMora830 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblSaldoCap 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label lblNCreditos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "(Sólo no minorista)"
         Height          =   195
         Left            =   2400
         TabIndex        =   35
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "(Sólo minorista)"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Mora >15:"
         Height          =   195
         Left            =   2400
         TabIndex        =   33
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Mora 8-15:"
         Height          =   195
         Left            =   2400
         TabIndex        =   32
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Mora >30:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Mora 8-30:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Créditos:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   15
      TabIndex        =   7
      Top             =   60
      Width           =   10410
      Begin VB.CheckBox CheckEx 
         Caption         =   "Ex-Usuarios"
         Height          =   255
         Left            =   3360
         TabIndex        =   57
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtMotivo 
         Height          =   285
         Left            =   5720
         TabIndex        =   25
         Top             =   1680
         Width           =   4620
      End
      Begin VB.ComboBox CboProducto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1650
         Width           =   2505
      End
      Begin VB.CheckBox ChkProducto 
         Caption         =   "Producto"
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
         Left            =   150
         TabIndex        =   23
         Top             =   1650
         Width           =   1155
      End
      Begin VB.CommandButton CmdUbicacion 
         Caption         =   "Ubicacion Geografica"
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
         Height          =   435
         Left            =   2520
         TabIndex        =   21
         Top             =   1200
         Width           =   1905
      End
      Begin VB.CheckBox ChkUbi 
         Caption         =   "Ubicación Geografica"
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
         Left            =   150
         TabIndex        =   20
         Top             =   1230
         Width           =   2235
      End
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   810
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   210
         Width           =   2475
      End
      Begin VB.ComboBox CmbAnalista 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   810
         Width           =   4425
      End
      Begin VB.ComboBox CmbAnalistaN 
         Height          =   315
         Left            =   5700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1260
         Width           =   4710
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Motivo:"
         Height          =   195
         Left            =   5160
         TabIndex        =   26
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ESTA OPCION ES IRREVERSIBLE LA OPERACION DE ASIGNACION DE CARTERA...           POR FAVOR CON MUCHO CUIDADO"
         ForeColor       =   &H00000080&
         Height          =   675
         Left            =   5610
         TabIndex        =   22
         Top             =   240
         Width           =   4665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Analista Asignado : "
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Analista Reasignado : "
         Height          =   195
         Left            =   5700
         TabIndex        =   10
         Top             =   1035
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2790
      Left            =   30
      TabIndex        =   2
      Top             =   2160
      Width           =   10500
      Begin MSComctlLib.ListView LstCredAna 
         Height          =   2565
         Left            =   105
         TabIndex        =   12
         Top             =   150
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4524
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Credito"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SalCap"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DiasAtraso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Mora830"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "MoraM30"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Mora815"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "MoraM15"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   390
         Left            =   4680
         TabIndex        =   6
         Top             =   495
         Width           =   1035
      End
      Begin VB.CommandButton CmdAgregarT 
         Caption         =   "&Todos >>"
         Height          =   390
         Left            =   4680
         TabIndex        =   5
         Top             =   960
         Width           =   1035
      End
      Begin VB.CommandButton CmdQuitar 
         Caption         =   "&Quitar"
         Height          =   390
         Left            =   4680
         TabIndex        =   4
         Top             =   1635
         Width           =   1035
      End
      Begin VB.CommandButton CmdQuitarT 
         Caption         =   "To&dos <<"
         Height          =   390
         Left            =   4680
         TabIndex        =   3
         Top             =   2145
         Width           =   1035
      End
      Begin MSComctlLib.ListView LstCredAnaN 
         Height          =   2565
         Left            =   5850
         TabIndex        =   13
         Top             =   150
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4524
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Credito"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SalCap"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DiasAtraso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Mora830"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "MoraM30"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Mora815"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "MoraM15"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   9360
      TabIndex        =   1
      Top             =   7200
      Width           =   1080
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   405
      Left            =   8160
      TabIndex        =   0
      Top             =   7200
      Width           =   1080
   End
   Begin VB.Label LblTotalAnaN 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6855
      TabIndex        =   17
      Top             =   5010
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6180
      TabIndex        =   16
      Top             =   4995
      Width           =   570
   End
   Begin VB.Label LblTotalAna 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   810
      TabIndex        =   15
      Top             =   5010
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   14
      Top             =   4995
      Width           =   570
   End
End
Attribute VB_Name = "frmCredReasigCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sUbicacionGeo As String

Private Sub GrabarNuevaCartera()
Dim i As Integer
Dim oCredito As COMDCredito.DCOMCredito
Dim oSeguridad As COMManejador.Pista
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim MatCuentas As Variant
Dim lsMovNro As String

    
    On Error GoTo ErrorGrabarNuevaCartera
'   Set oCredito = New COMDCredito.DCOMCredito
    ReDim MatCuentas(LstCredAnaN.ListItems.count, 8) 'WIOR 20131118
    
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    For i = 1 To LstCredAnaN.ListItems.count
    '    Call oCredito.ActualizaCreditoAnalista(LstCredAnaN.ListItems(i).Text, Right(CmbAnalistaN.Text, 13))
        MatCuentas(i, 1) = LstCredAnaN.ListItems(i).Text 'WIOR 20131118
        'WIOR 20131118 *****************************************
        MatCuentas(i, 2) = LstCredAnaN.ListItems(i).SubItems(2)
        MatCuentas(i, 3) = LstCredAnaN.ListItems(i).SubItems(3)
        MatCuentas(i, 4) = LstCredAnaN.ListItems(i).SubItems(4)
        MatCuentas(i, 5) = LstCredAnaN.ListItems(i).SubItems(5)
        MatCuentas(i, 6) = LstCredAnaN.ListItems(i).SubItems(6)
        MatCuentas(i, 7) = LstCredAnaN.ListItems(i).SubItems(7)
        MatCuentas(i, 8) = LstCredAnaN.ListItems(i).SubItems(8)
        'WIOR FIN **********************************************
    Next i
    Set oCredito = New COMDCredito.DCOMCredito
    'Call oCredito.ReasignaCarteraAnalista(MatCuentas, Right(CmbAnalistaN.Text, 13), Right(CmbAnalista.Text, 13), CDate(gdFecSis & " " & Time)) 'WIOR 20131118 CAMBIO gdFecSis POR CDate(gdFecSis & " " & Time)
    Call oCredito.ReasignaCarteraAnalista(MatCuentas, Right(CmbAnalistaN.Text, 13), Right(CmbAnalista.Text, 13), CDate(gdFecSis & " " & Time), Trim(Right(cboAgencia.Text, 3))) 'JUEZ 20160425
    Set oCredito = Nothing
    
    'BRGO 20110628 ***********************
    Set oSeguridad = New COMManejador.Pista
    oSeguridad.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, "Reasignación de cartera"
    Exit Sub
    'End BRGO

ErrorGrabarNuevaCartera:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub
'*** PEAC 20130510
Private Sub RegistroReasignacion(ByRef pcMovNro As String)
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    Set loContFunct = New COMNContabilidad.NCOMContFunciones

    Dim MatCuentasReasigna As Variant

    ReDim MatCuentasReasigna(LstCredAnaN.ListItems.count)
    Dim i As Integer
    Dim lsMovNro As String
        
    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    For i = 1 To LstCredAnaN.ListItems.count
        MatCuentasReasigna(i) = LstCredAnaN.ListItems(i).Text
    Next i

    Call RegistraCuentasReasignacion(lsMovNro, MatCuentasReasigna)
    Call RegistraReasignacionCartera(lsMovNro, Right(CmbAnalista.Text, 13), Right(CmbAnalistaN.Text, 13), 0, 0, "")
    pcMovNro = lsMovNro
    'cuarto parametro 0=vigente, 1=anulado
    'quinto parametro 1=aceptada, 2=rechazada, 0=aun no resuelto
    
End Sub
'*** PEAC 20130510
Private Function VerificaReasignacion(ByVal pcMovNro As String) As String
        
    Dim oCredito As COMDCredito.DCOMCredito
    Dim R As ADODB.Recordset

    Set oCredito = New COMDCredito.DCOMCredito
    Set R = oCredito.ConsultaReasignacionCartera(pcMovNro)
    Set oCredito = Nothing
 
    If Not R.EOF Then
        VerificaReasignacion = R!cEstadoTransfe 'R!nEstadoTransfe 'APRI20170530 ERS026-2017
    End If
    R.Close
    Set R = Nothing
    
End Function
'*** PEAC 20130510
Private Sub CancelaRegistroReasignacion(ByVal pcMovNro As String)

    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
    Dim lsMovNro As String
    Dim i As Integer
    Dim oCredito As COMDCredito.DCOMCredito

    On Error GoTo ErrorCancelaRegistroReasignacion

    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    Set oCredito = New COMDCredito.DCOMCredito
    Call oCredito.MantReasignacionCartera(pcMovNro, 1, 9, lsMovNro)
    'parametro nmovflag 1=anulado, 0=activo
    'parametro estadoasigna 1=aceptado 2=rechazado 0=aun no resuelto 9=mantiene el estado(no visible en la tabla)
    Set oCredito = Nothing
    
    Exit Sub

ErrorCancelaRegistroReasignacion:
        MsgBox Err.Description, vbCritical, "Aviso"

End Sub
'*** PEAC 20130510
Private Sub RegistraCuentasReasignacion(ByVal pcMovNro As String, ByVal pcMatrizCtas As Variant)

Dim i As Integer
Dim oCredito As COMDCredito.DCOMCredito

    On Error GoTo ErrorRegistraCuentasReasignacion

    Set oCredito = New COMDCredito.DCOMCredito
        
    For i = 1 To UBound(pcMatrizCtas)
        Call oCredito.RegDetalleCuentasReasignacion(pcMovNro, pcMatrizCtas(i))
    Next i

    Set oCredito = Nothing
    
    Exit Sub

ErrorRegistraCuentasReasignacion:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub
'*** PEAC 20130514
Private Sub RegistraReasignacionCartera(ByVal psMovNro As String, ByVal pcAnaAsignado As String, ByVal pcAnaReasignado As String, ByVal pnMovFlag As Integer, ByVal pnEstadoAsigna As Integer, ByVal psMovNroResuelto As String)
Dim i As Integer
Dim oCredito As COMDCredito.DCOMCredito

    On Error GoTo ErrorRegistraReasignacionCartera

    Set oCredito = New COMDCredito.DCOMCredito
    Call oCredito.RegReasignacionCartera(psMovNro, pcAnaAsignado, pcAnaReasignado, pnMovFlag, pnEstadoAsigna, psMovNroResuelto)
    Set oCredito = Nothing
    
    Exit Sub

ErrorRegistraReasignacionCartera:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

'Private Sub CargaAnalistas()
'Dim R As ADODB.Recordset
'Dim sSQL As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim oGen As COMDConstSistema.DCOMGeneral
'Dim sAnalistas As String
'
'    On Error GoTo ERRORCargaAnalistas
'
'    Set oGen = New COMDConstSistema.DCOMGeneral
'    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
'    Set oGen = Nothing
'
'    sSQL = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
'    sSQL = sSQL & " AND nRHEstado = 201 "
'    sSQL = sSQL & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
'    sSQL = sSQL & " where  RC.cRHCargoCod in (" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
'    sSQL = sSQL & " order by P.cPersNombre "
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set R = oConecta.CargaRecordSet(sSQL)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    CmbAnalista.Clear
'    CmbAnalistaN.Clear
'    Do While Not R.EOF
'        CmbAnalista.AddItem PstaNombre(R!cPersNombre) & Space(100) & R!cPersCod
'        CmbAnalistaN.AddItem PstaNombre(R!cPersNombre) & Space(100) & R!cPersCod
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    If CmbAnalista.ListCount > 0 Then
'        CmbAnalista.ListIndex = 0
'        CmbAnalistaN.ListIndex = 0
'    End If
'    Exit Sub
'ERRORCargaAnalistas:
'    MsgBox Err.Description, vbCritical, "Aviso"
'End Sub

Private Sub CargaListaCarteraAnalista()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim L As ListItem
Dim sTipoProducto As String
'WIOR 20131118 *********************
Dim oTC As COMDConstSistema.NCOMTipoCambio
Dim lnTC As Double
Dim lnTCUsa As Double
Dim lnSalCap As Double
Dim lnMora830 As Double
Dim lnMoraM30 As Double
Dim lnMora815 As Double
Dim lnMoraM15 As Double

Set oTC = New COMDConstSistema.NCOMTipoCambio
lnTC = oTC.EmiteTipoCambio(gdFecSis, TCFijoMes)

lnSalCap = 0
lnMora830 = 0
lnMoraM30 = 0
lnMora815 = 0
lnMoraM15 = 0
'WIOR FIN **************************

    On Error GoTo ErrorCargaListaCarteraAnalista
     If ChkProducto.value = 1 And CboProducto.ListIndex <> -1 Then
        sTipoProducto = CStr(CboProducto.ItemData(CboProducto.ListIndex))
     Else
        sTipoProducto = ""
     End If
     
    
    Set oCredito = New COMDCredito.DCOMCredito
    Set R = oCredito.CarteraAnalista(Right(CmbAnalista.Text, 13), Right(cboAgencia.List(cboAgencia.ListIndex), 2), Trim(Right(sUbicacionGeo, 18)), sTipoProducto)
    Set oCredito = Nothing
    'LstCredAna.ListItems.Clear'WIOR 20131119 COMENTÓ
    LimpiaDatosTransferidos 'WIOR 20131119
    Do While Not R.EOF
        Set L = LstCredAna.ListItems.Add(, , R!cCtaCod)
        L.SubItems(1) = R!cPersNombre
        'WIOR 20131118 ****************************
        L.SubItems(2) = R!SalCap
        L.SubItems(3) = R!cTpoCredCod
        L.SubItems(4) = R!nDiasAtraso
        L.SubItems(5) = R!Mora830
        L.SubItems(6) = R!MoraM30
        L.SubItems(7) = R!Mora815
        L.SubItems(8) = R!MoraM15
        
        lnTCUsa = CDbl(IIf(Mid(Trim(R!cCtaCod), 9, 1) = "1", 1, lnTC))
        lnSalCap = lnSalCap + CDbl(R!SalCap) * lnTCUsa
        lnMora830 = lnMora830 + CDbl(R!Mora830) * lnTCUsa
        lnMoraM30 = lnMoraM30 + CDbl(R!MoraM30) * lnTCUsa
        lnMora815 = lnMora815 + CDbl(R!Mora815) * lnTCUsa
        lnMoraM15 = lnMoraM15 + CDbl(R!MoraM15) * lnTCUsa
        'WIOR FIN *********************************
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    LblTotalAna.Caption = LstCredAna.ListItems.count
    LblTotalAnaN.Caption = LstCredAnaN.ListItems.count
    'WIOR 20131118 ***********************************
    lblNCreditos.Caption = LstCredAna.ListItems.count
    lblSaldoCap.Caption = Format(Round(lnSalCap, 2), "###," & String(15, "#") & "#0.00")
    lblMora830.Caption = Format(Round(lnMora830, 2), "###," & String(15, "#") & "#0.00")
    lblMoraM30.Caption = Format(Round(lnMoraM30, 2), "###," & String(15, "#") & "#0.00")
    lblMora815.Caption = Format(Round(lnMora815, 2), "###," & String(15, "#") & "#0.00")
    lblMoraM15.Caption = Format(Round(lnMoraM15, 2), "###," & String(15, "#") & "#0.00")
    'WIOR FIN ****************************************
    Exit Sub

ErrorCargaListaCarteraAnalista:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub ChkProducto_Click()
    If ChkProducto.value = 1 Then
        CboProducto.Enabled = True
    Else
        CboProducto.Enabled = False
      '  sCodTipoProducto = ""
    End If
End Sub

Private Sub ChkUbi_Click()
    If ChkUbi.value = 1 Then
        CmdUbicacion.Enabled = True
    Else
        CmdUbicacion.Enabled = False
        sUbicacionGeo = ""
    End If
End Sub

Private Sub CmbAnalista_Click()
    If cboAgencia.ListCount > 0 Then
        If cboAgencia.ListIndex = -1 Then
            MsgBox "Debe seleccionar una agencia", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    Call CargaListaCarteraAnalista
    LstCredAnaN.ListItems.Clear
    Screen.MousePointer = 0
End Sub

Private Sub CmbAnalistaN_Click()
    LstCredAnaN.ListItems.Clear 'WIOR 20131119 COMENTÒ
    'LimpiaDatosTransferidos 'WIOR 20131119
End Sub

Private Sub cmdAgregar_Click()
Dim L As ListItem

    On Error GoTo ErrorCmdAgregar_Click
    
    Set L = LstCredAnaN.ListItems.Add(, , LstCredAna.SelectedItem.Text)
    L.SubItems(1) = LstCredAna.SelectedItem.SubItems(1)
    'WIOR 20131118 ************************************
    L.SubItems(2) = LstCredAna.SelectedItem.SubItems(2)
    L.SubItems(3) = LstCredAna.SelectedItem.SubItems(3)
    L.SubItems(4) = LstCredAna.SelectedItem.SubItems(4)
    L.SubItems(5) = LstCredAna.SelectedItem.SubItems(5)
    L.SubItems(6) = LstCredAna.SelectedItem.SubItems(6)
    L.SubItems(7) = LstCredAna.SelectedItem.SubItems(7)
    L.SubItems(8) = LstCredAna.SelectedItem.SubItems(8)
    'WIOR FIN *****************************************
    Call LstCredAna.ListItems.Remove(LstCredAna.SelectedItem.Index)
    LblTotalAna.Caption = LstCredAna.ListItems.count
    LblTotalAnaN.Caption = LstCredAnaN.ListItems.count
    CalcularTransferencia 'WIOR 20131118
    Exit Sub

ErrorCmdAgregar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CmdAgregarT_Click()
Dim i As Integer
Dim L As ListItem
    On Error GoTo ErrorCmdAgregarT_Click
    Screen.MousePointer = 11
    For i = 1 To LstCredAna.ListItems.count
        Set L = LstCredAnaN.ListItems.Add(, , LstCredAna.ListItems(i).Text)
        L.SubItems(1) = LstCredAna.ListItems(i).SubItems(1)
        'WIOR 20131118 ************************************
        L.SubItems(2) = LstCredAna.ListItems(i).SubItems(2)
        L.SubItems(3) = LstCredAna.ListItems(i).SubItems(3)
        L.SubItems(4) = LstCredAna.ListItems(i).SubItems(4)
        L.SubItems(5) = LstCredAna.ListItems(i).SubItems(5)
        L.SubItems(6) = LstCredAna.ListItems(i).SubItems(6)
        L.SubItems(7) = LstCredAna.ListItems(i).SubItems(7)
        L.SubItems(8) = LstCredAna.ListItems(i).SubItems(8)
        'WIOR FIN *****************************************
    Next i
    
    Do While LstCredAna.ListItems.count > 0
        Call LstCredAna.ListItems.Remove(1)
    Loop
    LblTotalAna.Caption = LstCredAna.ListItems.count
    LblTotalAnaN.Caption = LstCredAnaN.ListItems.count
    Screen.MousePointer = 0
    CalcularTransferencia 'WIOR 20131118
    Exit Sub

ErrorCmdAgregarT_Click:
    Screen.MousePointer = 0
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdGrabar_Click()

    '*** PEAC 20130510
    If Len(Trim(Me.txtMotivo.Text)) = 0 Then
        MsgBox "Porfavor ingrese el motivo de la reasignación.", vbExclamation + vbOKOnly, "Atención"
        Me.txtMotivo.SetFocus
        Exit Sub
    End If
    '*** FIN PEAC

    If LstCredAnaN.ListItems.count > 0 Then
        If MsgBox("Se va a Actualizar la Cartera del Analista : " & Trim(Left(CmbAnalistaN.Text, 40)) & ", Desea Continuar ? ", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Dim bCondi As Boolean
            Dim lcMovNro As String
            '*** PEAC 20130510
            
            Call RegistroReasignacion(lcMovNro)
            bCondi = False
            Dim cEstadoTransferencia As String 'APRI20170530 ERS026-2017
            While bCondi = False
                If MsgBox("Para realizar la transferencia se requiere en VoBo del jefe de Negocios Territoriales, una vez que éste haya resuelto la transferencia en el Sistema pulse el botón ''Aceptar'', en caso de que no desee continuar con el proceso pulse ''Cancelar'', Desea Continuar ? ", vbInformation + vbOKCancel, "Aviso") = 1 Then
                    cEstadoTransferencia = VerificaReasignacion(lcMovNro) 'APRI20170530 ERS026-2017
                    If CInt(Right(cEstadoTransferencia, 1)) = 1 Then 'aceptada
                        If Left(cEstadoTransferencia, 6) = "002029" Then
                            MsgBox "La Transferencia ha sido [Aceptada] por el Jefe de Operaciones.", vbInformation, "Aviso" 'APRI20170530 ERS026-2017
                        Else
                            MsgBox "La Transferencia ha sido [Aceptada] por el Jefe de Negocios Territoriales.", vbInformation, "Aviso"
                        End If
                        Call GrabarNuevaCartera
                        'LstCredAnaN.ListItems.Clear 'WIOR 20131119 COMENTO
                        LimpiaDatosTransferidos 'WIOR 20131119 AGREGÓ
                        MsgBox "La reasignación de Cartera se realizó Satisfactoriamente.", vbInformation, "Aviso"
                        bCondi = True
                    ElseIf CInt(Right(cEstadoTransferencia, 1)) = 2 Then 'rechazada
                        If Left(cEstadoTransferencia, 6) = "002029" Then
                            MsgBox "La Transferencia ha sido [Rechazada] por el Jefe de Operaciones.", vbInformation, "Aviso" 'APRI20170530 ERS026-2017
                        Else
                            MsgBox "La Transferencia ha sido [Rechazada] por el Jefe de Negocios Territoriales.", vbInformation, "Aviso"
                        End If
                        LstCredAnaN.ListItems.Clear
                        bCondi = True
                    Else
                        'MsgBox "La Transferencia aún no fue resuelta por el Jefe de Negocios Territoriales.", vbInformation, "Aviso"
                        MsgBox "La Transferencia aún no fue resuelta.", vbInformation, "Aviso" 'APRI20170530 ERS026-2017
                    End If
                Else
                    Call CancelaRegistroReasignacion(lcMovNro)
                    bCondi = True
                    
                    MsgBox "La reasignacion de cartera fue cancelada.", vbInformation, "Aviso"
                    
                End If
            Wend

'            Call GrabarNuevaCartera
'            LstCredAnaN.ListItems.Clear
            '*** FIN PEAC
        End If
    Else
        MsgBox "No Existen Creditos a Reasignar", vbInformation, "Aviso"
    End If
    LblTotalAna.Caption = LstCredAna.ListItems.count
    LblTotalAnaN.Caption = LstCredAnaN.ListItems.count
End Sub

Private Sub cmdQuitar_Click()
Dim L As ListItem

On Error GoTo ErrorCmdQuitar_Click
    Set L = LstCredAna.ListItems.Add(, , LstCredAnaN.SelectedItem.Text)
    L.SubItems(1) = LstCredAnaN.SelectedItem.SubItems(1)
    'WIOR 20131118 ************************************
    L.SubItems(2) = LstCredAnaN.SelectedItem.SubItems(2)
    L.SubItems(3) = LstCredAnaN.SelectedItem.SubItems(3)
    L.SubItems(4) = LstCredAnaN.SelectedItem.SubItems(4)
    L.SubItems(5) = LstCredAnaN.SelectedItem.SubItems(5)
    L.SubItems(6) = LstCredAnaN.SelectedItem.SubItems(6)
    L.SubItems(7) = LstCredAnaN.SelectedItem.SubItems(7)
    L.SubItems(8) = LstCredAnaN.SelectedItem.SubItems(8)
    'WIOR FIN *****************************************
    
    Call LstCredAnaN.ListItems.Remove(LstCredAnaN.SelectedItem.Index)
    LblTotalAna.Caption = LstCredAna.ListItems.count
    LblTotalAnaN.Caption = LstCredAnaN.ListItems.count
    CalcularTransferencia 'WIOR 20131118
    Exit Sub

ErrorCmdQuitar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdQuitarT_Click()
Dim L As ListItem
Dim i As Integer
    On Error GoTo ErrorCmdQuitarT_Click
    Screen.MousePointer = 11
    For i = 1 To LstCredAnaN.ListItems.count
        Set L = LstCredAna.ListItems.Add(, , LstCredAnaN.ListItems(i).Text)
        L.SubItems(1) = LstCredAnaN.ListItems(i).SubItems(1)
        'WIOR 20131118 ************************************
        L.SubItems(2) = LstCredAnaN.ListItems(i).SubItems(2)
        L.SubItems(3) = LstCredAnaN.ListItems(i).SubItems(3)
        L.SubItems(4) = LstCredAnaN.ListItems(i).SubItems(4)
        L.SubItems(5) = LstCredAnaN.ListItems(i).SubItems(5)
        L.SubItems(6) = LstCredAnaN.ListItems(i).SubItems(6)
        L.SubItems(7) = LstCredAnaN.ListItems(i).SubItems(7)
        L.SubItems(8) = LstCredAnaN.ListItems(i).SubItems(8)
        'WIOR FIN *****************************************
    Next i
    Do While LstCredAnaN.ListItems.count > 0
        Call LstCredAnaN.ListItems.Remove(1)
    Loop
    LblTotalAna.Caption = LstCredAna.ListItems.count
    LblTotalAnaN.Caption = LstCredAnaN.ListItems.count
    Screen.MousePointer = 0
    CalcularTransferencia 'WIOR 20131118
    Exit Sub

ErrorCmdQuitarT_Click:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CmdUbicacion_Click()
 sUbicacionGeo = frmUbicacionGeo.Inicio
End Sub

Private Sub Command1_Click()
    frmCredConfirmaReasignaCartera.Show 1
End Sub

Private Sub Form_Load()
'    Call CargarAgencias
'    Call CargaAnalistas
'    Call CargarProductos
    Call CargarControles
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    gsOpeCod = "190290"
End Sub

Private Sub CargarControles(Optional ByVal exU As Integer = -1) 'PTI1 20170630 se agrego parametro exU
Dim rsAn As ADODB.Recordset
Dim rsProd As ADODB.Recordset
Dim rsAge As ADODB.Recordset
Dim i As Integer
Dim oCred As COMDCredito.DCOMCredito
Dim rsAnEx As ADODB.Recordset 'PTI1 20170630
Dim ex As Integer 'PTI1 20170630
ex = exU 'PTI1 20170630

On Error GoTo ERRORCargarControles
        
    Set oCred = New COMDCredito.DCOMCredito
    Call oCred.CargarControlesReasignacion(rsAn, rsAge, rsProd, rsAnEx) 'PTI1 20170630 SE AGREGO rsAnEx
    Set oCred = Nothing
    
    CmbAnalista.Clear
    CmbAnalistaN.Clear
    
    If ex = 1 Then 'PTI1 20170630 INICIO **************************************************
        Do While Not rsAnEx.EOF
        LstCredAna.ListItems.Clear
            CmbAnalista.AddItem PstaNombre(rsAnEx!cPersNombre) & Space(100) & rsAnEx!cPersCod
            rsAnEx.MoveNext
        Loop
         Do While Not rsAn.EOF
            CmbAnalistaN.AddItem PstaNombre(rsAn!cPersNombre) & Space(100) & rsAn!cPersCod
            rsAn.MoveNext
        Loop
        
        If CmbAnalista.ListCount > 0 Then
            CmbAnalista.ListIndex = 0
            CmbAnalistaN.ListIndex = 0
        End If
        
        cboAgencia.Clear
        Do Until rsAge.EOF
            cboAgencia.AddItem rsAge!cAgeDescripcion & Space(100) & rsAge!cAgeCod
            rsAge.MoveNext
        Loop
        
        For i = 0 To cboAgencia.ListCount - 1
            If Right(cboAgencia.List(i), 2) = gsCodAge Then
                cboAgencia.ListIndex = i
                Exit For
            End If
        Next i
        
        Do Until rsProd.EOF
            CboProducto.AddItem rsProd!cConsDescripcion
            CboProducto.ItemData(CboProducto.NewIndex) = rsProd!nConsValor
            rsProd.MoveNext
        Loop
        
        CboProducto.ListIndex = 0
    Else 'PTI1 20170630 FIN **************************************************
        Do While Not rsAn.EOF
            CmbAnalista.AddItem PstaNombre(rsAn!cPersNombre) & Space(100) & rsAn!cPersCod
            CmbAnalistaN.AddItem PstaNombre(rsAn!cPersNombre) & Space(100) & rsAn!cPersCod
            rsAn.MoveNext
        Loop
        
        If CmbAnalista.ListCount > 0 Then
            CmbAnalista.ListIndex = 0
            CmbAnalistaN.ListIndex = 0
        End If
        
        cboAgencia.Clear
        Do Until rsAge.EOF
            cboAgencia.AddItem rsAge!cAgeDescripcion & Space(100) & rsAge!cAgeCod
            rsAge.MoveNext
        Loop
        
        For i = 0 To cboAgencia.ListCount - 1
            If Right(cboAgencia.List(i), 2) = gsCodAge Then
                cboAgencia.ListIndex = i
                Exit For
            End If
        Next i
        
        Do Until rsProd.EOF
            CboProducto.AddItem rsProd!cConsDescripcion
            CboProducto.ItemData(CboProducto.NewIndex) = rsProd!nConsValor
            rsProd.MoveNext
        Loop
        
        CboProducto.ListIndex = 0
    End If
    Exit Sub
    
ERRORCargarControles:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub


Private Sub LstCredAna_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstCredAna.SortKey = ColumnHeader.SubItemIndex
    LstCredAna.SortOrder = lvwAscending
    LstCredAna.Sorted = True
End Sub

Private Sub LstCredAna_DblClick()
    If LstCredAna.ListItems.count > 0 Then
        cmdAgregar_Click
    End If
End Sub

Private Sub LstCredAnaN_DblClick()
    If LstCredAnaN.ListItems.count > 0 Then
        cmdQuitar_Click
    End If
End Sub

'Sub CargarAgencias()
'    Dim sSQL As String
'    Dim oConec As COMConecta.DCOMConecta
'    Dim rs As ADODB.Recordset
'    Dim i As Integer
'
'    sSQL = "select cagecod,cAgeDescripcion from agencias"
'
'    Set oConec = New COMConecta.DCOMConecta
'    oConec.AbreConexion
'    Set rs = oConec.CargaRecordSet(sSQL)
'    oConec.CierraConexion
'    Set oConec = Nothing
'    cboAgencia.Clear
'    Do Until rs.EOF
'        cboAgencia.AddItem rs!cAgeDescripcion & Space(100) & rs!cAgeCod
'     rs.MoveNext
'    Loop
'    Set rs = Nothing
'
'    For i = 0 To cboAgencia.ListCount - 1
'        If Right(cboAgencia.List(i), 2) = gsCodAge Then
'            cboAgencia.ListIndex = i
'            Exit For
'        End If
'    Next i
'
'End Sub

'Public Sub CargarProductos()
'    Dim rs As ADODB.Recordset
'    Dim odCredito As COMDCredito.DCOMCredito
'
'    Set odCredito = New COMDCredito.DCOMCredito
'    Set rs = odCredito.CargarProductos
'    Set odCredito = Nothing
'
'    Do Until rs.EOF
'        CboProducto.AddItem rs!cConsDescripcion
'        CboProducto.ItemData(CboProducto.NewIndex) = rs!nConsValor
'        rs.MoveNext
'    Loop
'    Set rs = Nothing
'    CboProducto.ListIndex = 0
'End Sub

'WIOR 20131118 *****************************
Private Sub CalcularTransferencia()
Dim oTC As COMDConstSistema.NCOMTipoCambio
Dim lnTC As Double
Dim lnTCUsa As Double
Dim i As Integer
Dim lnSalCap As Double
Dim lnMora830 As Double
Dim lnMoraM30 As Double
Dim lnMora815 As Double
Dim lnMoraM15 As Double

Set oTC = New COMDConstSistema.NCOMTipoCambio
lnTC = oTC.EmiteTipoCambio(gdFecSis, TCFijoMes)

lnSalCap = 0
lnMora830 = 0
lnMoraM30 = 0
lnMora815 = 0
lnMoraM15 = 0

If LstCredAnaN.ListItems.count > 0 Then
    For i = 1 To LstCredAnaN.ListItems.count
        lnTCUsa = CDbl(IIf(Mid(Trim(LstCredAnaN.ListItems(i).Text), 9, 1) = "1", 1, lnTC))
        lnSalCap = lnSalCap + CDbl(LstCredAnaN.ListItems(i).SubItems(2)) * lnTCUsa
        lnMora830 = lnMora830 + CDbl(LstCredAnaN.ListItems(i).SubItems(5)) * lnTCUsa
        lnMoraM30 = lnMoraM30 + CDbl(LstCredAnaN.ListItems(i).SubItems(6)) * lnTCUsa
        lnMora815 = lnMora815 + CDbl(LstCredAnaN.ListItems(i).SubItems(7)) * lnTCUsa
        lnMoraM15 = lnMoraM15 + CDbl(LstCredAnaN.ListItems(i).SubItems(8)) * lnTCUsa
    Next i
End If

lblNCreditosT.Caption = LstCredAnaN.ListItems.count
lblSaldoCapT.Caption = Format(Round(lnSalCap, 2), "###," & String(15, "#") & "#0.00")
lblMora830T.Caption = Format(Round(lnMora830, 2), "###," & String(15, "#") & "#0.00")
lblMoraM30T.Caption = Format(Round(lnMoraM30, 2), "###," & String(15, "#") & "#0.00")
lblMora815T.Caption = Format(Round(lnMora815, 2), "###," & String(15, "#") & "#0.00")
lblMoraM15T.Caption = Format(Round(lnMoraM15, 2), "###," & String(15, "#") & "#0.00")

Set oTC = Nothing
End Sub

Private Sub LimpiaDatosTransferidos()
LstCredAna.ListItems.Clear
LstCredAnaN.ListItems.Clear

txtMotivo.Text = ""
lblNCreditos.Caption = "0"
lblSaldoCap.Caption = "0.00"
lblMora830.Caption = "0.00"
lblMoraM30.Caption = "0.00"
lblMora815.Caption = "0.00"
lblMoraM15.Caption = "0.00"

lblNCreditosT.Caption = "0"
lblSaldoCapT.Caption = "0.00"
lblMora830T.Caption = "0.00"
lblMoraM30T.Caption = "0.00"
lblMora815T.Caption = "0.00"
lblMoraM15T.Caption = "0.00"
End Sub
'WIOR FIN **********************************
'PTI1 20170706 INICIO
Private Sub cboAgencia_Click()
Dim oAnali As COMDCredito.DCOMCredito
Set oAnali = New COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
If CheckEx.value = 1 Then
Set rs = oAnali.CargaAnalistasEx(cboAgencia.Text)
Else
Set rs = oAnali.CargaAnalistas(cboAgencia.Text)
End If
CmbAnalista.Clear
Do While Not rs.EOF
    LstCredAna.ListItems.Clear
        CmbAnalista.AddItem PstaNombre(rs!cPersNombre) & Space(100) & rs!cPersCod
        rs.MoveNext
    Loop
If CmbAnalista.ListCount > 0 Then
        CmbAnalista.ListIndex = 0
        CmbAnalistaN.ListIndex = 0
End If
End Sub
'END PTI1 20170706
Private Sub CheckEx_Click() 'PTI1 20170630
LstCredAna.ListItems.Clear
Call CargarControles(CheckEx.value)
End Sub
'END PTI1 20170630
