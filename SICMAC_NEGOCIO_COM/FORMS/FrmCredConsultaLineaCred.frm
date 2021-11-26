VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCredConsultaLineaCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Linea de Credito"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   Icon            =   "FrmCredConsultaLineaCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9765
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   60
      TabIndex        =   8
      Top             =   1200
      Width           =   11055
      Begin MSComctlLib.ListView lvwTarifario 
         Height          =   3100
         Left            =   60
         TabIndex        =   10
         Top             =   200
         Width           =   10900
         _ExtentX        =   19235
         _ExtentY        =   5477
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo Linea"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Plazo Min. Dias"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Plazo Max Dias"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Monto Min"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Monto Max"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Tasa Compesatoria"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Tasa Moratoria"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Tasa de Gracia"
            Object.Width           =   2469
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.CommandButton CmbFiltrar 
         Caption         =   "Filtrar"
         Height          =   375
         Left            =   9960
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox CboMoneda 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox CboProducto 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   7080
         TabIndex        =   5
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto: "
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   280
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agencia: "
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   280
         Width           =   675
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Tarifario por Agencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   840
      Width           =   11055
   End
End
Attribute VB_Name = "FrmCredConsultaLineaCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub CargaControles()
    Dim oLinea As COMDCredito.DCOMLineaCredito
    Dim lrsFondos As ADODB.Recordset
    Dim lrsProductos As ADODB.Recordset
    Dim lrsAgencias As ADODB.Recordset

    Set oLinea = New COMDCredito.DCOMLineaCredito
    Call oLinea.Cargar_Datos_Objetos_LineaCredito(lrsFondos, lrsProductos, lrsAgencias)

    CboMoneda.Clear
    CboMoneda.AddItem "SOLES" & Space(95) & "1" & Space(100) & "1"
    CboMoneda.AddItem "DOLARES" & Space(93) & "2" & Space(100) & "2"
        
    CboProducto.Clear
    Do While Not lrsProductos.EOF
        CboProducto.AddItem Trim(lrsProductos!cConsDescripcion) & Space(100 - Len(Trim(lrsProductos!cConsDescripcion))) & lrsProductos!nConsValor
        lrsProductos.MoveNext
    Loop
    
    CboAgencia.Clear
    With lrsAgencias
        Do While Not .EOF
            CboAgencia.AddItem Trim(!cAgeDescripcion) & Space(25) & !cAgeCod
            .MoveNext
        Loop
    End With
    
    CboMoneda.ListIndex = 0
    CboProducto.ListIndex = 0
    CboAgencia.ListIndex = 0
End Sub

Private Sub CboAgencia_Click()
    lvwTarifario.ListItems.Clear
End Sub

Private Sub CboMoneda_Click()
    lvwTarifario.ListItems.Clear
End Sub

Private Sub CboProducto_Click()
    lvwTarifario.ListItems.Clear
End Sub

Private Sub CmbFiltrar_Click()
    BuscarLineaAgencia
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargaControles
End Sub

Public Sub BuscarLineaAgencia()
     Dim oDCred As COMDCredito.DCOMLineaCredito
     Dim RS As New ADODB.Recordset
     Dim sCodAge As String
     Dim sCodProducto As String
     Dim sMoneda As String
     Dim lstTari As ListItem
     
     sCodAge = Right(CboAgencia.Text, 2)
     sCodProducto = Right(CboProducto, 3)
     sMoneda = Right(CboMoneda, 1)
     
     lvwTarifario.ListItems.Clear
     Set oDCred = New COMDCredito.DCOMLineaCredito
        Set RS = oDCred.RecuperaLineasCreditoAgencia(sCodAge, sCodProducto, sMoneda)
     Set oDCred = Nothing
     If Not (RS.EOF And RS.BOF) Then
        Do Until RS.EOF
            Set lstTari = lvwTarifario.ListItems.Add(, , RS(0))
            lstTari.SubItems(1) = RS(1)
            lstTari.SubItems(2) = Format(RS(3), "0.00")
            lstTari.SubItems(3) = Format(RS(2), "0.00")
            lstTari.SubItems(4) = Format(RS(5), "0.00")
            lstTari.SubItems(5) = Format(RS(4), "0.00")
            lstTari.SubItems(6) = Format(RS(6), "0.00")
            lstTari.SubItems(7) = Format(RS(7), "0.00")
            lstTari.SubItems(8) = Format(RS(8), "0.00")
            RS.MoveNext
        Loop
     End If
     
     
End Sub
