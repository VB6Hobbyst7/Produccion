VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelEtapas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Etapas del Proceso de Seleccion"
   ClientHeight    =   3165
   ClientLeft      =   1800
   ClientTop       =   3315
   ClientWidth     =   6720
   Icon            =   "frmLogProSelEtapas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6720
   Begin VB.Frame fraVis 
      Caption         =   "Etapas en Procesos de Selección "
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
      Width           =   6495
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Nueva"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   1155
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   2520
         Width           =   1155
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   2520
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   5220
         TabIndex        =   1
         Top             =   2520
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSEta 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
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
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4980
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   3660
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   4995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1380
         Width           =   840
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
      Left            =   300
      TabIndex        =   11
      Top             =   960
      Width           =   1020
   End
End
Attribute VB_Name = "frmLogProSelEtapas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String, nFuncion As Integer

Private Sub cmdEliminar_Click()
On Error GoTo cmdEliminarErr
    Dim oCon As DConecta, sSQL As String, nConsValor As Integer
    
    If Len(MSEta.TextMatrix(MSEta.row, 1)) = 0 Then Exit Sub
    
    Set oCon = New DConecta
    nConsValor = Val(MSEta.TextMatrix(MSEta.row, 1))
    If MsgBox("Seguro que Desea Eliminar...", vbQuestion + vbYesNo) = vbYes Then
        If oCon.AbreConexion Then
            'sSQL = "delete Constante where nConsCod=" & gcEtapasProcesoSel & " and nConsValor = " & CInt(txtCodigo) & " "
            sSQL = " UPDATE LogEtapa SET nEstado = 0 WHERE nEtapaCod = " & nConsValor
            oCon.Ejecutar sSQL
            MsgBox "Eliminado Correctamente...", vbInformation
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
            sSQL = " declare @tmp int "
            sSQL = sSQL & "set @tmp = (select isnull(count(nEtapaCod),0) from LogEtapa where cDescripcion= '" & txtDescripcion.Text & "') "
            sSQL = sSQL & "if @tmp = 0 "
            sSQL = sSQL & "INSERT INTO LogEtapa(nEtapaCod,cDescripcion) " & _
                        " VALUES (" & nConsValor & ",'" & txtDescripcion.Text & "') "
            sSQL = sSQL & " else "
            sSQL = sSQL & " UPDATE LogEtapa SET nEstado = 1 WHERE nEtapaCod = @tmp"
       Case 2
            sSQL = "UPDATE LogEtapa SET cDescripcion = '" & txtDescripcion.Text & "' WHERE nEtapaCod = " & nConsValor
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

sSQL = "SELECT nEtapaCod, cDescripcion FROM LogEtapa WHERE nEstado = 1"

Set Rs = oConn.CargaRecordSet(sSQL)
If Not Rs.EOF Then
   Do While Not Rs.EOF
      i = i + 1
      InsRow MSEta, i
      MSEta.TextMatrix(i, 1) = Rs!nEtapaCod
      MSEta.TextMatrix(i, 2) = Rs!cDescripcion
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
MSEta.ColWidth(2) = 5800
MSEta.ColWidth(3) = 0
MSEta.ColWidth(4) = 0
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
   Set Rs = oConn.CargaRecordSet("Select nMaxNro=isnull(Max(nEtapaCod),0) from LogEtapa ")   'where nEtapaCod = " & gcEtapasProcesoSel)
   If Not Rs.EOF Then
      txtCodigo.Text = Rs!nMaxNro + 1
   Else
      txtCodigo.Text = 1
   End If
End If
End Sub

Private Sub cmdModificar_Click()
If Len(MSEta.TextMatrix(MSEta.row, 1)) = 0 Then Exit Sub
nFuncion = 2
txtCodigo.Text = MSEta.TextMatrix(MSEta.row, 1)
txtDescripcion.Text = MSEta.TextMatrix(MSEta.row, 2)
fraVis.Visible = False
fraReg.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmLogProSelEtapas = Nothing
End Sub
