VERSION 5.00
Begin VB.Form frmMantOpCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Pago de Caja General"
   ClientHeight    =   6060
   ClientLeft      =   765
   ClientTop       =   1950
   ClientWidth     =   10020
   Icon            =   "frmMantOpCaja.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fracontrol 
      Height          =   2925
      Left            =   8325
      TabIndex        =   14
      Top             =   60
      Width           =   1590
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   2430
         Width           =   1350
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   105
         TabIndex        =   7
         Top             =   660
         Width           =   1350
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   105
         TabIndex        =   6
         Top             =   225
         Width           =   1350
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   105
         TabIndex        =   8
         Top             =   225
         Width           =   1350
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   105
         TabIndex        =   9
         Top             =   660
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8490
      TabIndex        =   11
      Top             =   5445
      Width           =   1350
   End
   Begin VB.Frame FraRangoDet 
      Caption         =   "Detalle de Ordenes de Pagos"
      Height          =   2775
      Left            =   1290
      TabIndex        =   13
      Top             =   3030
      Width           =   6960
      Begin VB.CommandButton cmdCancela 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5415
         TabIndex        =   4
         Top             =   600
         Width           =   1260
      End
      Begin VB.TextBox txtDescRangoDet 
         Height          =   690
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1905
         Width           =   4935
      End
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "&Adicionar"
         Height          =   375
         Left            =   5415
         TabIndex        =   3
         Top             =   225
         Width           =   1260
      End
      Begin VB.CommandButton cmdeliminaDet 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   5430
         TabIndex        =   5
         Top             =   1380
         Width           =   1260
      End
      Begin Sicmact.FlexEdit fgRangoDet 
         Height          =   1605
         Left            =   165
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2831
         Cols0           =   6
         HighLight       =   2
         AllowUserResizing=   1
         EncabezadosNombres=   "#-CodRango-Rango Inicial-Rango Final-Estado-cDescripcion"
         EncabezadosAnchos=   "350-0-1500-1500-1200-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-4-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-3-0"
         EncabezadosAlineacion=   "C-C-R-R-L-L"
         FormatosEdit    =   "0-0-3-3-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   285
      End
   End
   Begin VB.Frame fraRangoPadre 
      Caption         =   "Rangos de Ordenes de Pagos"
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
      Height          =   2940
      Left            =   105
      TabIndex        =   12
      Top             =   45
      Width           =   8130
      Begin Sicmact.FlexEdit fgOpRangos 
         Height          =   2460
         Left            =   150
         TabIndex        =   0
         Top             =   255
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   4339
         Cols0           =   7
         HighLight       =   2
         AllowUserResizing=   1
         EncabezadosNombres=   "#-Cod Rango-Rango Inicial-Rango Final-Fecha Ing-Estado-Moneda"
         EncabezadosAnchos=   "350-1200-1200-1200-1000-1200-1200"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-X-5-6"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-3-3"
         EncabezadosAlineacion=   "C-C-R-R-C-L-L"
         FormatosEdit    =   "0-0-3-3-3-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   285
      End
   End
End
Attribute VB_Name = "frmMantOpCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbNuevo As Boolean
Dim lnItem As Long
Dim lnItemDet As Long
Dim NCajaGen As nCajaGeneral
Dim oContFunc As NContFunciones
Dim oGen As DGeneral

Private Sub cmdAdicionar_Click()
Dim lsRangoCod As String
Dim lsMovUltAct As String
Dim lnRangoIni As Long
Dim lnRangoFin As Long
Dim lsEstado As String


If fgOpRangos.TextMatrix(1, 0) = 0 Then
    MsgBox "No encuenta Rangos de OP principales ", vbInformation, "Aviso"
    Exit Sub
End If

If cmdAdicionar.Caption = "&Adicionar" Then
    fgRangoDet.CargaCombo NCajaGen.GetEstadosOp(True, Anuladas, Extraviadas)
    lnItem = fgOpRangos.Row
    lsRangoCod = fgOpRangos.TextMatrix(lnItem, 1)
    If Right(fgOpRangos.TextMatrix(lnItem, 5), 1) <> CGEstadosOp.PorEmitir Then
        MsgBox "No existen Op por emitir dentro del rango seleccionado", vbInformation, "Aviso"
        fgRangoDet.SetFocus
        Exit Sub
    End If
    
    
    fgRangoDet.AdicionaFila
    lnItemDet = fgRangoDet.Row
    fgRangoDet.TextMatrix(lnItemDet, 1) = lsRangoCod
    fgRangoDet.SoloFila = True
    fgRangoDet.lbEditarFlex = True
    cmdAdicionar.Caption = "&Grabar"
    cmdCancela.Enabled = True
    fracontrol.Enabled = False
    fraRangoPadre.Enabled = False
    fgRangoDet.Col = 2
    fgRangoDet.SetFocus
    SendKeys "{Enter}"
Else
    If Len(Trim(txtDescRangoDet)) = 0 Then
        MsgBox "Glosa o Concepto no ingresado", vbInformation, "Aviso"
        txtDescRangoDet.SetFocus
        Exit Sub
    End If
    lsRangoCod = fgRangoDet.TextMatrix(lnItemDet, 1)
    lnRangoIni = fgRangoDet.TextMatrix(lnItemDet, 2)
    lnRangoFin = fgRangoDet.TextMatrix(lnItemDet, 3)
    lsEstado = Right(fgRangoDet.TextMatrix(lnItemDet, 4), 1)
    If MsgBox("Desea Guardar la Información", vbYesNo + vbQuestion) = vbYes Then
        lsMovUltAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        NCajaGen.GrabaOpDetCajaGeneral lsRangoCod, txtDescRangoDet, lnRangoIni, lnRangoFin, Val(lsEstado), lsMovUltAct, False
        cmdCancela_Click
    End If
End If
End Sub
Private Sub cmdCancela_Click()
CargaRangosOPDetalle fgOpRangos.TextMatrix(lnItem, 1)
fracontrol.Enabled = False
fraRangoPadre.Enabled = False
cmdAdicionar.Caption = "&Adicionar"
cmdCancela.Enabled = False
fgRangoDet.lbEditarFlex = False
fgRangoDet.SoloFila = False
Me.fracontrol.Enabled = True
Me.fraRangoPadre.Enabled = True
End Sub
Private Sub cmdCancelar_Click()
fgOpRangos.lbEditarFlex = False
fgOpRangos.SoloFila = False
CargaRangosOP
HabilitaControles False
lbNuevo = False
End Sub
Private Sub cmdEditar_Click()
If fgRangoDet.TextMatrix(1, 0) <> "" Then
    MsgBox "Posee Ordenes de Pago en Detalle. El Proceso no Podrá Continuar", vbInformation, "Aviso"
    Exit Sub
End If
lnItem = fgOpRangos.Row
If Right(fgOpRangos.TextMatrix(lnItem, 5), 1) <> CGEstadosOp.PorEmitir Then
    MsgBox "Sólo se puede Editar Rangos de Ordenes que se encuentren en Estado de Emitidas", vbInformation, "Aviso"
    Exit Sub
End If
HabilitaControles True
fgOpRangos.lbEditarFlex = True
fgOpRangos.SoloFila = True
fgOpRangos.ColumnasAEditar = "X-X-2-3-X-X-X"
fgOpRangos.CargaCombo NCajaGen.GetEstadosOp()
lbNuevo = False
fgOpRangos.SetFocus
End Sub
Private Function ValidaInterfaz() As Boolean
ValidaInterfaz = True
If Val(fgOpRangos.TextMatrix(lnItem, 2)) = 0 Then
    MsgBox "Rango Inicial no Válido", vbInformation, "Aviso"
    ValidaInterfaz = False
    fgOpRangos.SetFocus
    Exit Function
End If
If Val(fgOpRangos.TextMatrix(lnItem, 3)) = 0 Then
    MsgBox "Rango Final no Válido", vbInformation, "Aviso"
    ValidaInterfaz = False
    fgOpRangos.SetFocus
    Exit Function
End If
If Len(Trim(fgOpRangos.TextMatrix(lnItem, 5))) = 0 Then
    MsgBox "Estado no válido", vbInformation, "Aviso"
    ValidaInterfaz = False
    fgOpRangos.SetFocus
    Exit Function
End If

End Function
Private Sub cmdeliminaDet_Click()
Dim lsRangoCod As String
Dim lnRangoIni As Long
Dim lnRangoFin As Long

If Me.fgRangoDet.TextMatrix(1, 0) <> "" Then
    If Right(fgOpRangos.TextMatrix(fgOpRangos.Row, 5), 1) <> CGEstadosOp.PorEmitir Then
        MsgBox "No existen Op por emitir dentro del rango seleccionado", vbInformation, "Aviso"
        fgRangoDet.SetFocus
        Exit Sub
    End If

    lnItemDet = fgRangoDet.Row
    lsRangoCod = fgRangoDet.TextMatrix(lnItemDet, 1)
    lnRangoIni = fgRangoDet.TextMatrix(lnItemDet, 2)
    lnRangoFin = fgRangoDet.TextMatrix(lnItemDet, 3)
    If MsgBox("Desea Eliminar el Detalle del Rango N° " & lsRangoCod & " que se encuentra entre " & lnRangoIni & " y " & lnRangoFin, vbYesNo + vbQuestion, "Aviso") = vbYes Then
        NCajaGen.EliminaOpDetCaja lsRangoCod, lnRangoIni, lnRangoFin
        CargaRangosOPDetalle lsRangoCod
    End If
End If
End Sub
Private Sub cmdEliminar_Click()
Dim lsRangoCod As String
Dim lnRangoIni As Long
Dim lnRangoFin As Long

If fgRangoDet.TextMatrix(1, 0) <> "" Then
    MsgBox "Posee Ordenes de Pago en Detalle. El Proceso no Podrá Continuar", vbInformation, "Aviso"
    Exit Sub
End If

lnItem = fgOpRangos.Row
lnRangoIni = Me.fgOpRangos.TextMatrix(lnItem, 2)
lnRangoFin = Me.fgOpRangos.TextMatrix(lnItem, 3)
lsRangoCod = Me.fgOpRangos.TextMatrix(lnItem, 1)
If fgOpRangos.TextMatrix(1, 0) <> "" Then
    If MsgBox("Desea Eliminar el Rango N°" & lsRangoCod & " que se encuentra entre " & lnRangoIni & " y " & lnRangoFin, vbYesNo + vbQuestion, "Aviso") = vbYes Then
        NCajaGen.EliminaOpCajaGeneral lsRangoCod, False
        CargaRangosOP
    End If
End If
End Sub

Private Sub cmdGrabar_Click()
Dim lsMovUltAct As String
Dim lsRangoCod As String
Dim lnRangoIni As Long
Dim lnRangoFin As Long
Dim ldFechaRango As Date
Dim lsEstado As String
Dim lsMoneda As String
If ValidaInterfaz = False Then Exit Sub
lsRangoCod = fgOpRangos.TextMatrix(lnItem, 1)
lnRangoIni = fgOpRangos.TextMatrix(lnItem, 2)
lnRangoFin = fgOpRangos.TextMatrix(lnItem, 3)
ldFechaRango = CDate(fgOpRangos.TextMatrix(lnItem, 4))
lsEstado = Right(fgOpRangos.TextMatrix(lnItem, 5), 1)
lsMoneda = Right(fgOpRangos.TextMatrix(lnItem, 6), 1)

If ValidaMoneda(lsMoneda) = False Then
    fgOpRangos.SetFocus
    Exit Sub
End If
If MsgBox("Desea Grabar la Informacón", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    lsMovUltAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    If lbNuevo Then
        NCajaGen.GrabaRangoOpCajaGeneral lsRangoCod, ldFechaRango, lnRangoIni, lnRangoFin, lsEstado, lsMovUltAct, lsMoneda, False
    Else
        NCajaGen.UpdateOpCajaGeneral lsRangoCod, lnRangoIni, lnRangoFin, lsMovUltAct, False
    End If
    cmdCancelar_Click
End If
End Sub

Private Sub cmdNuevo_Click()
fgOpRangos.CargaCombo NCajaGen.GetEstadosOp(True, PorEmitir, PorEmitir)
lbNuevo = True
HabilitaControles True
fgOpRangos.lbEditarFlex = True
fgOpRangos.ColumnasAEditar = "X-X-2-3-X-5-6"
fgOpRangos.AdicionaFila
lnItem = fgOpRangos.Row
fgOpRangos.TextMatrix(lnItem, 1) = NCajaGen.GetNroCodOp
fgOpRangos.TextMatrix(lnItem, 4) = gdFecSis
fgOpRangos.SoloFila = True
fgRangoDet.Clear
fgRangoDet.FormaCabecera
fgRangoDet.Rows = 2
fgOpRangos.Col = 2
fgOpRangos.SetFocus
SendKeys "{Enter}"
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub



Private Sub fgOpRangos_Click()
If fgOpRangos.TextMatrix(1, 0) <> "" Then
'    Select Case Me.fgOpRangos.Col
'        Case 5
'            If lbNuevo Then
'                fgOpRangos.CargaCombo NCajaGen.GetEstadosOp(True, PorEmitir, PorEmitir)
'            Else
'                fgOpRangos.CargaCombo NCajaGen.GetEstadosOp()
'            End If
'        Case 6
'            fgOpRangos.CargaCombo oContFunc.GetConstantes(gMoneda)
'    End Select
    CargaRangosOPDetalle fgOpRangos.TextMatrix(fgOpRangos.Row, 1)
End If
End Sub
Private Sub fgOpRangos_OnRowChange(pnRow As Long, pnCol As Long)
If fgOpRangos.TextMatrix(1, 0) <> "" Then
    CargaRangosOPDetalle fgOpRangos.TextMatrix(fgOpRangos.Row, 1)
End If
End Sub
Private Sub fgOpRangos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
 If Val(fgOpRangos.TextMatrix(lnItem, 3)) > 0 And Val(fgOpRangos.TextMatrix(lnItem, 2)) > 0 Then
    Select Case fgOpRangos.Col
        Case 2
            If CCur(fgOpRangos.TextMatrix(lnItem, 2)) > CCur(fgOpRangos.TextMatrix(lnItem, 3)) Then
                MsgBox "Rango Inicial no puede ser mayor que rango Final", vbInformation, "Aviso"
                Cancel = False
            End If
        Case 3
            If CCur(fgOpRangos.TextMatrix(lnItem, 3)) < CCur(fgOpRangos.TextMatrix(lnItem, 2)) Then
                MsgBox "Rango Final no puede ser menor que Rango Inicial", vbInformation, "Aviso"
                Cancel = False
            End If
    End Select
End If
End Sub
Private Sub fgOpRangos_RowColChange()
If cmdGrabar.Visible Then
    Set oGen = New DGeneral
    Select Case fgOpRangos.Col
        Case 5
            If lbNuevo Then
                fgOpRangos.CargaCombo NCajaGen.GetEstadosOp(True, PorEmitir, PorEmitir)
            Else
                fgOpRangos.CargaCombo NCajaGen.GetEstadosOp()
            End If
        Case 6
            fgOpRangos.CargaCombo oGen.GetConstante(gMoneda)
    End Select
End If
End Sub

Private Sub fgRangoDet_Click()
If fgRangoDet.TextMatrix(1, 0) <> "" Then
    txtDescRangoDet = fgRangoDet.TextMatrix(fgRangoDet.Row, 5)
End If
End Sub

Private Sub fgRangoDet_OnRowChange(pnRow As Long, pnCol As Long)
If fgRangoDet.TextMatrix(1, 0) <> "" Then
    txtDescRangoDet = fgRangoDet.TextMatrix(fgRangoDet.Row, 5)
End If
End Sub
Private Sub fgRangoDet_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnRangoEmiIni As Long
Dim lnRangoEmiFin As Long
 If Val(fgRangoDet.TextMatrix(lnItemDet, fgRangoDet.Col)) > 0 Then
    lnRangoEmiIni = fgOpRangos.TextMatrix(lnItem, 2)
    lnRangoEmiFin = fgOpRangos.TextMatrix(lnItem, 3)
    If Not (CCur(fgRangoDet.TextMatrix(lnItemDet, fgRangoDet.Col)) >= lnRangoEmiIni And CCur(fgRangoDet.TextMatrix(lnItemDet, fgRangoDet.Col)) <= lnRangoEmiFin) Then
        MsgBox "El Valor Ingresado no se encuentra entre los rangos establecidos", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    Select Case fgRangoDet.Col
        Case 2
            If VerificaValor(CCur(fgRangoDet.TextMatrix(lnItemDet, 2))) Then
                Cancel = False
                Exit Sub
            End If
            If Val(fgRangoDet.TextMatrix(lnItemDet, 3)) = 0 Then Exit Sub
            If CCur(fgRangoDet.TextMatrix(lnItemDet, 2)) > CCur(fgRangoDet.TextMatrix(lnItemDet, 3)) Then
                MsgBox "Rango Inicial no puede ser mayor que rango Final", vbInformation, "Aviso"
                Cancel = False
            End If
        Case 3
            If VerificaValor(CCur(fgRangoDet.TextMatrix(lnItemDet, 3))) = False Then Exit Sub
            If Val(fgRangoDet.TextMatrix(lnItemDet, 2)) = 0 Then Exit Sub
            If CCur(fgRangoDet.TextMatrix(lnItemDet, 3)) < CCur(fgRangoDet.TextMatrix(lnItemDet, 2)) Then
                MsgBox "Rango Final no puede ser menor que Rango Inicial", vbInformation, "Aviso"
                Cancel = False
            End If
    End Select
    
End If
End Sub

Private Sub Form_Load()
Set NCajaGen = New nCajaGeneral
Set oContFunc = New NContFunciones
Set oGen = New DGeneral
CentraForm Me

fgOpRangos.CargaCombo NCajaGen.GetEstadosOp()
fgRangoDet.CargaCombo NCajaGen.GetEstadosOp()
CentraForm Me
lnItem = 0
CargaRangosOP
HabilitaControles False
End Sub
Sub HabilitaControles(lbHab As Boolean)
cmdNuevo.Visible = Not lbHab
cmdEditar.Visible = Not lbHab
cmdcancelar.Visible = lbHab
cmdGrabar.Visible = lbHab
cmdeliminar.Enabled = Not lbHab
FraRangoDet.Enabled = Not lbHab
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set NCajaGen = Nothing
End Sub
Private Sub CargaRangosOP()
Dim rs As ADODB.Recordset
Dim lnFila As Long
Set rs = New ADODB.Recordset
fgOpRangos.Clear
fgOpRangos.FormaCabecera
fgOpRangos.Rows = 2
Set rs = NCajaGen.GetOpCajaGeneral
Do While Not rs.EOF
    Me.fgOpRangos.AdicionaFila
    lnFila = Me.fgOpRangos.Row
    fgOpRangos.TextMatrix(lnFila, 1) = rs!cRangoCod
    fgOpRangos.TextMatrix(lnFila, 2) = Format(rs!nRangoEmiIni, "#,#0")
    fgOpRangos.TextMatrix(lnFila, 3) = Format(rs!nRangoEmiFin, "#,#0")
    fgOpRangos.TextMatrix(lnFila, 4) = Format(rs!dFechaIng, "dd/mm/yyyy")
    fgOpRangos.TextMatrix(lnFila, 5) = rs!cConsDescripcion & Space(50) & rs!cEstadoRango
    fgOpRangos.TextMatrix(lnFila, 6) = rs!cConsDescMon & Space(50) & rs!cmoneda
    rs.MoveNext
Loop
rs.Close: Set rs = Nothing
End Sub
Private Sub CargaRangosOPDetalle(psRangoCod As String)
Dim rs As ADODB.Recordset
Dim lnFila As Long
Set rs = New ADODB.Recordset

fgRangoDet.Clear
fgRangoDet.FormaCabecera
fgRangoDet.Rows = 2
Set rs = NCajaGen.GetOpDetCajaGeneral(psRangoCod)
Do While Not rs.EOF
    fgRangoDet.AdicionaFila
    lnFila = fgRangoDet.Row
    fgRangoDet.TextMatrix(lnFila, 1) = rs!cRangoCod
    fgRangoDet.TextMatrix(lnFila, 2) = Format(rs!nRangoIni, "#,#0")
    fgRangoDet.TextMatrix(lnFila, 3) = Format(rs!nRangoFin, "#,#0")
    fgRangoDet.TextMatrix(lnFila, 4) = rs!cConsDescripcion & Space(50) & rs!cEstadoDet
    fgRangoDet.TextMatrix(lnFila, 5) = rs!cDescDet
    rs.MoveNext
Loop
rs.Close: Set rs = Nothing
End Sub
Private Function VerificaValor(pnValor As Long) As Boolean
Dim I As Long
Dim lnValorIni As Long
Dim lnValorFin As Long
VerificaValor = False
For I = 1 To fgRangoDet.Rows - 2
    lnValorIni = CCur(IIf(fgRangoDet.TextMatrix(I, 2) = "", "0", fgRangoDet.TextMatrix(I, 2)))
    lnValorFin = CCur(IIf(fgRangoDet.TextMatrix(I, 3) = "", "0", fgRangoDet.TextMatrix(I, 3)))
    If pnValor >= lnValorIni And pnValor <= lnValorFin Then
        MsgBox "Valor Ingresado se encuentra dentro de rangos ya Ingresados", vbInformation, "Aviso"
        VerificaValor = True
        Exit Function
    End If
Next
End Function
Private Function ValidaMoneda(psMoneda As String) As Boolean
Dim I As Integer
ValidaMoneda = True
For I = 1 To Me.fgOpRangos.Rows - 2
    If Right(fgOpRangos.TextMatrix(I, 5), 1) = CGEstadosOp.PorEmitir And Right(fgOpRangos.TextMatrix(I, 6), 1) = psMoneda Then
        MsgBox "Aun Existen Ordenes de Pagos Por emitir de la Moneda Seleccionada", vbInformation, "Aviso"
        ValidaMoneda = False
        Exit Function
    Exit Function
    End If
Next I
End Function
