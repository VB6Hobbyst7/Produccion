VERSION 5.00
Begin VB.Form frmCapConvenioMant 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   Icon            =   "frmCapConvenioMant.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6780
      TabIndex        =   7
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7860
      TabIndex        =   6
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton cmdPlanPagos 
      Caption         =   "&Plan Pagos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1140
      TabIndex        =   5
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton cmdCuentas 
      Caption         =   "&Cuentas"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   3780
      Width           =   975
   End
   Begin VB.Frame fraConvenio 
      Caption         =   "Personas Convenio"
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
      Height          =   3615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8955
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   6720
         TabIndex        =   3
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   7785
         TabIndex        =   2
         Top             =   3120
         Width           =   975
      End
      Begin SICMACT.FlexEdit grdConvenio 
         Height          =   2820
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   8655
         _extentx        =   15266
         _extenty        =   4974
         cols0           =   6
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Codigo-Nombre-Tipo-nConvCod-Flag"
         encabezadosanchos=   "300-1500-3500-2800-0-0"
         font            =   "frmCapConvenioMant.frx":030A
         font            =   "frmCapConvenioMant.frx":0332
         font            =   "frmCapConvenioMant.frx":035A
         font            =   "frmCapConvenioMant.frx":0382
         font            =   "frmCapConvenioMant.frx":03AA
         fontfixed       =   "frmCapConvenioMant.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         tipobusqueda    =   3
         columnasaeditar =   "X-1-X-3-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-1-0-3-0-0"
         encabezadosalineacion=   "C-L-L-L-C-C"
         formatosedit    =   "0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         colwidth0       =   300
         rowheight0      =   300
         tipobuspersona  =   1
         forecolorfixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapConvenioMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HabilitaControles()
If grdConvenio.Rows >= 2 Then
    cmdEliminar.Enabled = True
    cmdCuentas.Enabled = True
    cmdCuentas.Enabled = True
    cmdPlanPagos.Enabled = True
Else
    If grdConvenio.TextMatrix(1, 1) = "" Then
        cmdEliminar.Enabled = False
        cmdCuentas.Enabled = False
        cmdCuentas.Enabled = False
        cmdPlanPagos.Enabled = False
    End If
End If
cmdAgregar.Enabled = True
cmdGrabar.Enabled = True
End Sub

Private Sub CmdAgregar_Click()
    grdConvenio.AdicionaFila
    grdConvenio.TextMatrix(grdConvenio.Rows - 1, 5) = "N"
    grdConvenio.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdCuentas_Click()
Dim nFila As Long
    nFila = grdConvenio.Row
    frmCapServConvCuentas.Inicia grdConvenio.TextMatrix(nFila, 1), grdConvenio.TextMatrix(nFila, 2), 101
End Sub

Private Sub cmdEliminar_Click()
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios

    If MsgBox("¿¿Está seguro de eliminar a la persona de la relación??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
        clsServ.EliminaServConvenios grdConvenio.TextMatrix(grdConvenio.Row, 1)
        
        grdConvenio.EliminaFila grdConvenio.Row
        
        HabilitaControles
    End If

Set clsServ = Nothing

End Sub

Private Sub cmdGrabar_Click()
Dim rs As New ADODB.Recordset
If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
        Set rs = grdConvenio.GetRsNew()
        clsServ.ActualizaServConvenios rs
        grdConvenio.TextMatrix(grdConvenio.Rows - 1, 5) = ""
    Set clsServ = Nothing
    HabilitaControles
End If
End Sub

Private Sub CmdPlanPagos_Click()
Dim nFila As Long
    nFila = grdConvenio.Row
    frmCapServConvPlanPag.Inicia grdConvenio.TextMatrix(nFila, 1), grdConvenio.TextMatrix(nFila, 2)
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsConv As Recordset
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
'Set rsConv = clsServ.GetServConvenios(gCapConvUNT)

Set rsConv = clsServ.GetServConvenios2()
Set clsServ = Nothing
Me.Icon = LoadPicture(App.path & gsRutaIcono)
If Not (rsConv.EOF And rsConv.BOF) Then
    Set grdConvenio.Recordset = rsConv
End If
Set rsConv = Nothing
Set clsGen = New COMDConstSistema.DCOMGeneral
    grdConvenio.CargaCombo clsGen.GetConstante(gCaptacConvenios)
Set clsGen = Nothing
End Sub

Private Sub grdConvenio_Click()
    If grdConvenio.TextMatrix(grdConvenio.Row, 4) = "101" And grdConvenio.TextMatrix(grdConvenio.Row, 5) <> "N" And grdConvenio.TextMatrix(grdConvenio.Row, 5) <> "S" Then
        cmdPlanPagos.Enabled = False
        cmdCuentas.Enabled = True
    ElseIf grdConvenio.TextMatrix(grdConvenio.Row, 4) = "102" And grdConvenio.TextMatrix(grdConvenio.Row, 5) <> "N" And grdConvenio.TextMatrix(grdConvenio.Row, 5) <> "S" Then
        cmdPlanPagos.Enabled = True
        cmdCuentas.Enabled = True
    ElseIf grdConvenio.TextMatrix(grdConvenio.Row, 5) = "N" And grdConvenio.TextMatrix(grdConvenio.Row, 5) = "S" Then
    
        cmdPlanPagos.Enabled = False
        cmdCuentas.Enabled = False
    End If
End Sub

Private Sub grdConvenio_GotFocus()
    If grdConvenio.TextMatrix(grdConvenio.Row, 4) = "101" And grdConvenio.TextMatrix(grdConvenio.Row, 5) <> "N" And grdConvenio.TextMatrix(grdConvenio.Row, 5) <> "S" Then
        cmdPlanPagos.Enabled = True
        cmdCuentas.Enabled = True
    Else
        cmdPlanPagos.Enabled = False
        cmdCuentas.Enabled = False
    End If
End Sub

Private Sub grdConvenio_OnCellChange(pnRow As Long, pnCol As Long)

If grdConvenio.TextMatrix(pnRow, 4) <> "101" And grdConvenio.TextMatrix(pnRow, 5) <> "N" And grdConvenio.TextMatrix(grdConvenio.Row, 5) <> "S" Then
    cmdPlanPagos.Enabled = True
    cmdCuentas.Enabled = True
Else

    cmdPlanPagos.Enabled = False
    cmdCuentas.Enabled = False
End If
End Sub

Private Sub grdConvenio_OnChangeCombo()
'grdConvenio.TextMatrix(grdConvenio.Row, 5) = "S"

'If Right(grdConvenio.TextMatrix(grdConvenio.Row, 3), 3) <> "101" Then
'    cmdPlanPagos.Enabled = True
'Else
'    cmdPlanPagos.Enabled = False
'End If
End Sub

Private Sub grdConvenio_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
If psDataCod <> "" Then
'    If grdConvenio.PersPersoneria = gPersonaNat Then
'        MsgBox "Los convenios sólo pueden ser personas jurídicas.", vbInformation, "Aviso"
'        grdConvenio.EliminaFila pnRow
'    End If
End If
End Sub

Private Sub grdConvenio_OnRowChange(pnRow As Long, pnCol As Long)
If grdConvenio.TextMatrix(pnRow, 5) = "N" Then
    cmdEliminar.Enabled = True
Else
    cmdEliminar.Enabled = False
End If
End Sub
