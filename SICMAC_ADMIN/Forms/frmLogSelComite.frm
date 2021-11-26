VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogSelComite 
   Caption         =   "Mantenimiento de Comite Predeterminado"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   Icon            =   "frmLogSelComite.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   9870
   Begin VB.ComboBox cmbtipoComite 
      Height          =   315
      ItemData        =   "frmLogSelComite.frx":030A
      Left            =   120
      List            =   "frmLogSelComite.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Editar"
      Height          =   390
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   3
      Left            =   3000
      TabIndex        =   5
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   390
      Left            =   8520
      TabIndex        =   0
      Top             =   4080
      Width           =   1305
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Comites Predeterminados"
      TabPicture(0)   =   "frmLogSelComite.frx":030E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FlexCoimite"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdMant(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdMant(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdMant 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   1305
      End
      Begin VB.CommandButton cmdMant 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   390
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   3000
         Width           =   1305
      End
      Begin Sicmact.FlexEdit FlexCoimite 
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4471
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Codigo-Nombres-Cargo-UltimaActualizacion"
         EncabezadosAnchos=   "550-1500-3400-1300-2600"
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
         ColumnasAEditar =   "X-1-X-3-X"
         ListaControles  =   "0-1-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbPuntero       =   -1  'True
         ColWidth0       =   555
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmLogSelComite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim clsDGnral As DLogGeneral
Dim clsDGAdqui As DLogAdquisi
Dim ClsNAdqui As NActualizaProcesoSelecLog
Dim oCons As DConstantes
Dim saccion As String



Private Sub cmbtipoComite_Click()
mostrar_Comite_Tipo Right(cmbtipoComite.Text, 1)
End Sub

Private Sub cmdMant_Click(Index As Integer)
Select Case Index
Case 0
        FlexCoimite.AdicionaFila
        FlexCoimite.SetFocus
Case 1
        nBSRow = FlexCoimite.Row
        If MsgBox("¿ Estás seguro de eliminar " & FlexCoimite.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            FlexCoimite.EliminaFila nBSRow
            clsDGAdqui.EliminaSeleccionTipoComite Right(cmbtipoComite.Text, 1)
        End If
End Select
End Sub

Private Sub cmdReq_Click(Index As Integer)
Dim sactualiza As String
sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
Select Case Index
Case 1 'Editar
        saccion = "E"
        cmdReq(1).Enabled = False  'Editar
        cmdReq(2).Enabled = True  'Cancelar
        cmdReq(3).Enabled = True  'Grabar
        cmdMant(1).Enabled = True  'Eliminar
        cmdMant(0).Enabled = True  'Nuevo
        FlexCoimite.Enabled = True
        cmbtipoComite.Enabled = False
        
Case 2 'Cancelar
        saccion = "C"
        cmdReq(1).Enabled = True  'Editar
        cmdReq(2).Enabled = False  'Cancelar
        cmdReq(3).Enabled = False 'Grabar
        cmdMant(1).Enabled = False  'Eliminar
        cmdMant(0).Enabled = False 'Nuevo
        FlexCoimite.Enabled = False
        cmbtipoComite.Enabled = True
        
Case 3 'Grabar
        If cmbtipoComite.Text = "" Then
           MsgBox "Seleccione el tipo de Comite", vbInformation, "Seleccione un Numero de Proceso"
           txtSeleccionA.SetFocus
           Exit Sub
        End If
        If FlexCoimite.Rows <= 2 And FlexCoimite.TextMatrix(1, 1) = "" Then
            MsgBox "debe ingresar los integrantes del Comite   " & Left(cmbtipoComite.Text, 30), vbInformation, "Ingrese Los Integrantes del Comite"
            FlexCoimite.SetFocus
            Exit Sub
        End If
        For i = 0 To FlexCoimite.Rows - 1
            If FlexCoimite.TextMatrix(i, 0) = "" Then
                MsgBox "Falta Ingresar el Integrante del comite  del Item  Nº " & FlexCoimite.TextMatrix(i, 0), vbInformation, "Ingrese el Integrante del Comite"
                FlexCoimite.SetFocus
                Exit Sub
            End If
            
            If FlexCoimite.TextMatrix(i, 3) = "" Then
                MsgBox "Falta Ingresar el cargo del integrante del Comite " & FlexCoimite.TextMatrix(i, 0), vbInformation, "Ingrese el Integrante del Comite"
                FlexCoimite.SetFocus
                Exit Sub
            End If
            
            
            
        Next
        Select Case saccion
            Case "E"
                    If FlexCoimite.Rows = 2 And FlexCoimite.TextMatrix(1, 1) = "" Then
                                'Elimina
                                clsDGAdqui.EliminaSeleccionTipoComite Right(cmbtipoComite.Text, 1)
                                Exit Sub
                    End If
                                ClsNAdqui.AgregaSeleccionTipoComite Right(cmbtipoComite.Text, 1), FlexCoimite.GetRsNew, sactualiza
         End Select
        cmdReq(1).Enabled = True  'Editar
        cmdReq(2).Enabled = False  'Cancelar
        cmdReq(3).Enabled = False 'Grabar
        cmdMant(1).Enabled = False  'Eliminar
        cmdMant(0).Enabled = False 'Nuevo
        FlexCoimite.Enabled = False
        saccion = "G"
        mostrar_Comite_Tipo Right(cmbtipoComite.Text, 1)
        cmbtipoComite.Enabled = True
        
End Select
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 4999
Me.Width = 9990
Dim sAno As String
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDGAdqui = New DLogAdquisi
Set ClsNAdqui = New NActualizaProcesoSelecLog
Set rs = clsDGAdqui.CargaSelTipoComite()
Call CargaCombo(rs, cmbtipoComite)
cmbtipoComite.ListIndex = 1

Set rs = clsDGAdqui.CargaSelcargos
Me.FlexCoimite.CargaCombo rs

FlexCoimite.BackColorBkg = -2147483643
End Sub

Sub mostrar_Comite_Tipo(nLogTipoComite As Long)
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = clsDGAdqui.CargaLogSelComitePre(nLogTipoComite)
If rs.EOF = True Then
    FlexCoimite.Rows = 2
    FlexCoimite.Clear
    FlexCoimite.FormaCabecera
    Else
    Set FlexCoimite.Recordset = rs
End If
End Sub

