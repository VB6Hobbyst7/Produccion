VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelCnsPostores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Postores"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmLogProSelCnsPostores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "S&alir"
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton CmndSeleccionar 
      Caption         =   "&Seleccionar"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10455
      Begin VB.TextBox txtPostor 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   280
         Width           =   8500
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSItemPostores 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4048
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   -2147483647
         ForeColorSel    =   -2147483624
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de l Postor"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmLogProSelCnsPostores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gnProSelNro As Integer, cCadena As String
Public gcPersCod As String, gcPersNombre As String

Public Sub Inicio(pnProSelNro As Integer, pcCadena As String)
gnProSelNro = pnProSelNro
cCadena = pcCadena
If cCadena <> "" Then
    cCadena = " and x.cPersCod not in('" & cCadena & "')"
End If
Me.Show 1
End Sub

Private Sub CargarPostores(pnPSN As Integer, pNombre As String, pcCadena As String)
    On Error GoTo CargarPostoresErr
    Dim oConn As New DConecta, Rs As ADODB.Recordset, sSQL As String, i As Integer
    sSQL = "select x.nMovNroVentaBase, x.dFecha, p.cPersCod, p.cPersNombre  from LogProSelPostor x " & _
           "    inner join Persona p on x.cPersCod = p.cPersCod " & _
           "    where p.cPersNombre like '" & pNombre & "%'" & " and x.nProSelNro=" & pnPSN & pcCadena & _
           "    order by x.nMovNroVentaBase"
    oConn.AbreConexion
    FormaFlexItemPostor
    Set Rs = oConn.CargaRecordSet(sSQL)
    i = 1
    Do While Not Rs.EOF
        InsRow MSItemPostores, i
        MSItemPostores.TextMatrix(i, 0) = Rs!nMovNroVentaBase
        MSItemPostores.TextMatrix(i, 1) = Format(Rs!dFecha, "dd/mm/yyyy")
        MSItemPostores.TextMatrix(i, 2) = Rs!cPersCod
        MSItemPostores.TextMatrix(i, 3) = Rs!cPersNombre
        i = i + 1
        Rs.MoveNext
    Loop
    oConn.CierraConexion
    Exit Sub
CargarPostoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Sub FormaFlexItemPostor()
With MSItemPostores
    .Clear
    .Rows = 2
    .Cols = 4
    .RowHeight(0) = 320
    .RowHeight(1) = 8
    .ColWidth(0) = 1000:     .ColAlignment(1) = 4:   .TextMatrix(0, 0) = " Item"
    .ColWidth(1) = 1500:     .ColAlignment(1) = 4:   .TextMatrix(0, 1) = " Fecha"
    .ColWidth(2) = 1500:     .ColAlignment(2) = 4:   .TextMatrix(0, 2) = " Código"
    .ColWidth(3) = 6000:    .TextMatrix(0, 3) = " Nombre"
End With
End Sub

Private Sub cmdSalir_Click()
    gcPersCod = ""
    Unload Me
End Sub

Private Sub CmndSeleccionar_Click()
    gcPersCod = MSItemPostores.TextMatrix(MSItemPostores.Row, 2)
    gcPersNombre = MSItemPostores.TextMatrix(MSItemPostores.Row, 3)
    Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
FormaFlexItemPostor
gcPersCod = "": gcPersNombre = ""
End Sub

Private Sub MSItemPostores_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then _
        CmndSeleccionar_Click
End Sub

Private Sub txtPostor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CargarPostores gnProSelNro, txtPostor, cCadena
        MSItemPostores.SetFocus
    End If
End Sub
