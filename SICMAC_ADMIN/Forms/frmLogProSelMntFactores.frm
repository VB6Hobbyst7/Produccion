VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelMntFactores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Factores"
   ClientHeight    =   3165
   ClientLeft      =   1365
   ClientTop       =   3525
   ClientWidth     =   7905
   Icon            =   "frmLogProSelMntFactores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7905
   Begin VB.Frame fraVis 
      Caption         =   "Factores de Evaluacion del Procesos de Selección "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6300
         TabIndex        =   4
         Top             =   2520
         Width           =   1155
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   2520
         Width           =   1155
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   2520
         Width           =   1155
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSEta 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         WordWrap        =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Registro de Etapas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmLogProSelMntFactores.frx":08CA
         Left            =   1920
         List            =   "frmLogProSelMntFactores.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox txtUnidades 
         Height          =   315
         Left            =   1920
         TabIndex        =   15
         Top             =   1320
         Width           =   4995
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Width           =   4995
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   4380
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5700
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Factor"
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Unidades"
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   900
         TabIndex        =   12
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1200
         TabIndex        =   11
         Top             =   660
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
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
      Left            =   180
      TabIndex        =   13
      Top             =   840
      Width           =   1020
   End
End
Attribute VB_Name = "frmLogProSelMntFactores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String, nFuncion As Integer

Private Sub cmdEliminar_Click()
    On Error GoTo cmdEliminarErr
    Dim oCon As DConecta, sSQL As String
    Set oCon = New DConecta
    
    If Len(MSEta.TextMatrix(MSEta.row, 1)) = 0 Then Exit Sub
    
    If MsgBox("Seguro que Desea Eliminar...", vbQuestion + vbYesNo) = vbYes Then
        If oCon.AbreConexion Then
            sSQL = "update LogProSelFactor set nFactorEstado=0 where nFactorNro=" & MSEta.TextMatrix(MSEta.row, 1)
            oCon.Ejecutar sSQL
            MsgBox "Factor Eliminado Correctamente...", vbInformation
            ListaEtapas
            oCon.CierraConexion
        End If
    End If
    Exit Sub
cmdEliminarErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta
Dim nConsValor As Integer

If MsgBox("¿ Está seguro de agregar la Etapa indicada ?" + Space(10), vbQuestion + vbYesNo, "Confirme operación") = vbYes Then
   nConsValor = CInt(txtCodigo)
   Select Case nFuncion
       Case 1
            sSQL = "INSERT INTO LogProSelFactor (cFactorDescripcion,cUnidades,nTipo) " & _
                   " VALUES ('" & txtDescripcion.Text & "','" & txtUnidades.Text & "'," & cboTipo.ItemData(cboTipo.ListIndex) & ")"
       Case 2
            
            sSQL = "UPDATE LogProSelFactor SET cFactorDescripcion = '" & txtDescripcion.Text & "', cUnidades= '" & txtUnidades.Text & "', nTipo=" & cboTipo.ItemData(cboTipo.ListIndex) & " WHERE nFactorNro = " & txtCodigo.Text
   End Select
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
   End If
   fraReg.Visible = False
   fraVis.Visible = True
   ListaEtapas
End If
End Sub

Private Sub Form_Load()
CentraForm Me
ListaEtapas
cboTipo.ListIndex = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub ListaEtapas()
Dim Rs As New ADODB.Recordset, oConn As New DConecta, i As Integer
If Not oConn.AbreConexion Then
   Exit Sub
End If
LimpiaFlexEta

sSQL = "SELECT nFactorNro, cFactorDescripcion, cUnidades, nTipo FROM LogProSelFactor WHERE nFactorEstado=1"

Set Rs = oConn.CargaRecordSet(sSQL)
If Not Rs.EOF Then
   Do While Not Rs.EOF
      i = i + 1
      InsRow MSEta, i
      MSEta.TextMatrix(i, 1) = Rs!nFactorNro
      MSEta.TextMatrix(i, 2) = Rs!cFactorDescripcion
      MSEta.TextMatrix(i, 3) = Rs!cUnidades
      MSEta.TextMatrix(i, 4) = IIf(Rs!nTipo, "Economica", "Tecnica")
      Rs.MoveNext
   Loop
End If
End Sub

Sub LimpiaFlexEta()
MSEta.Clear
MSEta.Rows = 2
MSEta.RowHeight(0) = 320
MSEta.RowHeight(1) = 8
MSEta.ColWidth(0) = 0
MSEta.ColWidth(1) = 260:  MSEta.ColAlignment(1) = 4
MSEta.ColWidth(2) = 3800
MSEta.ColWidth(3) = 2000
MSEta.ColWidth(4) = 1200
End Sub

Private Sub cmdCancelar_Click()
fraReg.Visible = False
fraVis.Visible = True
End Sub

Private Sub cmdAgregar_Click()
Dim oConn As New DConecta, Rs As New ADODB.Recordset

nFuncion = 1
txtDescripcion.Text = ""
fraVis.Visible = False
fraReg.Visible = True

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet("Select nMaxNro=coalesce(Max(nFactorNro),0) from LogProSelFactor")
   If Not Rs.EOF Then
      txtCodigo.Text = Rs!nMaxNro + 1
   Else
      txtCodigo.Text = 1
   End If
End If
End Sub

Private Sub cmdModificar_Click()
nFuncion = 2

If Len(MSEta.TextMatrix(MSEta.row, 1)) = 0 Then Exit Sub

txtCodigo.Text = MSEta.TextMatrix(MSEta.row, 1)
txtDescripcion.Text = MSEta.TextMatrix(MSEta.row, 2)
txtUnidades.Text = MSEta.TextMatrix(MSEta.row, 3)
cboTipo.Text = MSEta.TextMatrix(MSEta.row, 4)
fraVis.Visible = False
fraReg.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelMntFactores = Nothing
End Sub
