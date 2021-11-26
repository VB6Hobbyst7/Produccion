VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajeroCierreAgencias 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmCajeroCierreAgencias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   5580
      Width           =   1185
   End
   Begin VB.CommandButton cmdResumenIE 
      Caption         =   "&Resumen de Ingresos y Egresos"
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   5580
      Width           =   2625
   End
   Begin VB.Frame fraAgencias 
      Caption         =   "Agencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5370
      Left            =   90
      TabIndex        =   6
      Top             =   135
      Width           =   6135
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton optSel 
         Caption         =   "&Ninguna"
         Height          =   240
         Index           =   1
         Left            =   1260
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optSel 
         Caption         =   "&Todas"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   1230
      End
      Begin MSComctlLib.ListView lvwAgencia 
         Height          =   4470
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   7885
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agencia"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cierre?"
            Object.Width           =   1764
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCajeroCierreAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()
    Dim oCajero As COMNCajaGeneral.NCOMCajero
    Dim L As MSComctlLib.ListItem
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    For Each L In lvwAgencia.ListItems
        If L.Checked Then
            If oCajero.YaRealizoCierreAgencia(L.Text, gdFecSis) Then
                L.SubItems(2) = "SI"
            Else
                L.SubItems(2) = "NO"
            End If
        Else
            L.SubItems(2) = ""
        End If
    Next
    Set oCajero = Nothing
End Sub

Private Sub cmdResumenIE_Click()
    frmCajeroIngEgre.Inicia False, True, True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim rs As ADODB.Recordset
    Dim L As MSComctlLib.ListItem
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    Set rs = oGen.GetAgencias()
    Set oGen = Nothing
    
    Do While Not rs.EOF
        Set L = lvwAgencia.ListItems.Add(, , rs("cAgeCod"))
        L.SubItems(1) = rs("cAgeDescripcion")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Me.Caption = "Verifica Cierre de Agencias"
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub optSel_Click(Index As Integer)
    Dim L As MSComctlLib.ListItem
    Dim bSeleccion As Boolean
    Select Case Index
        Case 0
            bSeleccion = True
        Case 1
            bSeleccion = False
    End Select
    For Each L In lvwAgencia.ListItems
        L.Checked = bSeleccion
    Next
End Sub
