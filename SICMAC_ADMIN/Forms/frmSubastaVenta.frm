VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogSubastaVenta 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmSubastaVenta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3450
      TabIndex        =   27
      Text            =   "002-00000001"
      Top             =   0
      Width           =   2985
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9525
      TabIndex        =   26
      Top             =   5835
      Width           =   1125
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   285
      Left            =   9375
      TabIndex        =   15
      Top             =   360
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraPer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   0
      TabIndex        =   2
      Top             =   645
      Width           =   10650
      Begin Sicmact.TxtBuscar txtPersona 
         Height          =   300
         Left            =   1140
         TabIndex        =   3
         Top             =   217
         Width           =   1860
         _extentx        =   3281
         _extenty        =   529
         appearance      =   0
         appearance      =   0
         font            =   "frmSubastaVenta.frx":030A
         appearance      =   0
         tipobusqueda    =   3
         stitulo         =   ""
      End
      Begin VB.Label lblRucG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4770
         TabIndex        =   9
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblRuc 
         Caption         =   "Ruc :"
         Height          =   195
         Left            =   4335
         TabIndex        =   13
         Top             =   885
         Width           =   510
      End
      Begin VB.Label lblDNI 
         Caption         =   "DNI :"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   885
         Width           =   945
      End
      Begin VB.Label lblDir 
         Caption         =   "Dirección  :"
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   570
         Width           =   945
      End
      Begin VB.Label lblFonoG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9150
         TabIndex        =   10
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label lblDNIG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1140
         TabIndex        =   8
         Top             =   847
         Width           =   1395
      End
      Begin VB.Label lblDirG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1140
         TabIndex        =   7
         Top             =   540
         Width           =   9405
      End
      Begin VB.Label lblPerG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3060
         TabIndex        =   5
         Top             =   225
         Width           =   7485
      End
      Begin VB.Label lblPer 
         Caption         =   "Pesona :"
         Height          =   195
         Left            =   105
         TabIndex        =   4
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lblFono 
         Caption         =   "Telefono :"
         Height          =   195
         Left            =   8355
         TabIndex        =   14
         Top             =   885
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7110
      TabIndex        =   1
      Top             =   5835
      Width           =   1125
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   8325
      TabIndex        =   0
      Top             =   5835
      Width           =   1125
   End
   Begin VB.Frame fraVenta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   0
      TabIndex        =   17
      Top             =   1740
      Width           =   10650
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   345
         Left            =   75
         TabIndex        =   21
         Top             =   3315
         Width           =   1125
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   1260
         TabIndex        =   20
         Top             =   3315
         Width           =   1125
      End
      Begin Sicmact.FlexEdit FlexSerie 
         Height          =   630
         Left            =   4635
         TabIndex        =   18
         Top             =   3300
         Visible         =   0   'False
         Width           =   630
         _extentx        =   1111
         _extenty        =   1111
         highlight       =   1
         rowsizingmode   =   1
         encabezadosnombres=   "#-Serie"
         encabezadosanchos=   "300-1700"
         font            =   "frmSubastaVenta.frx":0336
         font            =   "frmSubastaVenta.frx":035E
         font            =   "frmSubastaVenta.frx":0386
         font            =   "frmSubastaVenta.frx":03AE
         font            =   "frmSubastaVenta.frx":03D6
         fontfixed       =   "frmSubastaVenta.frx":03FE
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         tipobusqueda    =   0
         columnasaeditar =   "X-1"
         textstylefixed  =   3
         listacontroles  =   "0-1"
         encabezadosalineacion=   "C-L"
         formatosedit    =   "0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         appearance      =   0
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit FlexDetalle 
         Height          =   3045
         Left            =   75
         TabIndex        =   19
         Top             =   210
         Width           =   10470
         _extentx        =   18468
         _extenty        =   5371
         cols0           =   9
         highlight       =   1
         rowsizingmode   =   1
         encabezadosnombres=   "#-Lote-Codigo-Producto-P. Unit-Cant-Total-Stock-Almacen"
         encabezadosanchos=   "300-1200-1200-3500-900-900-1000-900-0"
         font            =   "frmSubastaVenta.frx":0424
         font            =   "frmSubastaVenta.frx":044C
         font            =   "frmSubastaVenta.frx":0474
         font            =   "frmSubastaVenta.frx":049C
         font            =   "frmSubastaVenta.frx":04C4
         fontfixed       =   "frmSubastaVenta.frx":04EC
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         columnasaeditar =   "X-1-2-X-X-5-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-0-1-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-R-R-R-R-C"
         formatosedit    =   "0-0-0-0-2-2-2-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         appearance      =   0
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label lblTotalG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   300
         Left            =   8925
         TabIndex        =   23
         Top             =   3300
         Width           =   1275
      End
      Begin VB.Label lblTotalMG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   300
         Left            =   8925
         TabIndex        =   22
         Top             =   3630
         Width           =   1275
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sub Total             :"
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
         Left            =   7200
         TabIndex        =   24
         Top             =   3300
         Width           =   3000
      End
      Begin VB.Label lblTotM 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total + %Martillero :"
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
         Left            =   7200
         TabIndex        =   25
         Top             =   3630
         Width           =   3000
      End
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha :"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8565
      TabIndex        =   16
      Top             =   390
      Width           =   660
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Boleta de Venta : "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   -15
      TabIndex        =   6
      Top             =   0
      Width           =   3450
   End
End
Attribute VB_Name = "frmLogSubastaVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String
Dim lsOpeCod As String
Dim lsDocTpo As Long
Dim lbBoleta As Boolean
Dim lnPorMartillero As Currency
Dim lnPorIGV As Currency
Dim lnMovNroSubActual As Long
Dim lsSubActual As String


Public Sub Ini(psOpeCod As String, psCaption As String, pbBoleta As Boolean)
    lsCaption = psCaption
    lsOpeCod = psOpeCod
    lbBoleta = pbBoleta
    
    Me.Show 1
End Sub

Private Function Valida() As Boolean
    Dim i As Integer
    
    If Me.txtPersona.Text = "" Then
        MsgBox "Debe Ingresar una persona.", vbInformation, "Aviso"
        txtPersona.SetFocus
        Valida = False
        Exit Function
    End If
    
    For i = 1 To Me.FlexDetalle.Rows - 1
        If Me.FlexDetalle.TextMatrix(i, 1) = "" Then
            MsgBox "Debe Ingresar un codigo de lote.", vbInformation, "Aviso"
            FlexDetalle.Col = 1
            FlexDetalle.Row = i
            txtPersona.SetFocus
            Valida = False
            Exit Function
        ElseIf Me.FlexDetalle.TextMatrix(i, 2) = "" Then
            MsgBox "Debe Ingresar un codigo de Bien a subastar.", vbInformation, "Aviso"
            FlexDetalle.Col = 1
            FlexDetalle.Row = i
            txtPersona.SetFocus
            Valida = False
            Exit Function
        ElseIf Me.FlexDetalle.TextMatrix(i, 2) = "" Then
            MsgBox "Debe Ingresar un codigo de Bien a subastar.", vbInformation, "Aviso"
            FlexDetalle.Col = 1
            FlexDetalle.Row = i
            txtPersona.SetFocus
            Valida = False
            Exit Function
        End If
    Next i
    
    Valida = True
End Function

Private Sub cmdGrabar_Click()
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oGen As DLogGeneral
    Set oGen = New DLogGeneral
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim i As Integer
    Dim lnContador As Integer
    Dim lsCta As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    Dim lnMontoProvAnual As Currency
    Dim lnMontoProvAnterior As Currency
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim odoc As NContFunciones
    Set odoc = New NContFunciones
    If Not Valida Then Exit Sub
    
    If MsgBox("Desea Grabar los cambios ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Me.txtNumDoc.Text = odoc.GeneraDocNro(lsDocTpo, gMonedaExtranjera, oGen.GetOpeSerie(lsOpeCod))
    
'    Insert MovVenta(nMovNro, nMovItem, nMovNroSub, nMovNroIng)
'Values()

    
    oMov.BeginTrans
        lsMovNro = oMov.GeneraMovNro(gdFecSis, Left(gsCodAge, 2), gsCodUser)
        oMov.InsertaMov lsMovNro, lsOpeCod, "Venta Subasta " & lsSubActual, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        oMov.InsertaMovDoc lnMovNro, 1, Me.txtNumDoc.Text, Format(gdFecSis, gsFormatoFecha)
        
        For i = 1 To Me.FlexDetalle.Rows - 1
            lsCta = oOpe.EmiteOpeCta(lsOpeCod, "D", , Right(gsCodAge, 2), ObjCMACAgencias)
            oMov.InsertaMovBS lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 8), Me.FlexDetalle.TextMatrix(i, 2)
            oMov.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 5)
            oMov.InsertaMovCta lnMovNro, i, lsCta, Me.FlexDetalle.TextMatrix(i, 6)
            oMov.InsertaMovVenta lnMovNro, i, lnMovNroSubActual, Me.FlexDetalle.TextMatrix(i, 1)
        Next i
        
        oMov.InsertaMovCta lnMovNro, i, lsCta, Format((CCur(Me.lblTotalMG.Caption) - CCur(Me.lblTotalG.Caption)), "#.00")
        
        lnContador = i + 1
        
        'Ctas Pendientes para provicion - Reversion
        For i = 1 To Me.FlexDetalle.Rows - 1
            Set rs = oSubasta.GetProvSubasta(Me.FlexDetalle.TextMatrix(i, 1), Me.FlexDetalle.TextMatrix(i, 2))
            
            lnMontoProvAnual = 0
            
            While Not rs.EOF
                lnMontoProvAnual = lnMontoProvAnual + rs!Monto
                rs.MoveNext
            Wend
            
            'lsCta = oMov.GetOpeCtaCtaOtro(gnAlmaIngXAdjudicacion, "", GetCtaCntBS(Me.FlexDetalle.TextMatrix(i, 2), gnAlmaIngXAdjudicacion), False)
            oMov.InsertaMovCta lnMovNro, lnContador, lsCta, lnMontoProvAnual * CCur(Me.FlexDetalle.TextMatrix(i, 5))
            lnContador = lnContador + 1
            rs.Close
        Next i
    
        'Ctas Pendientes para provicion - Reversion ----------   11111
        For i = 1 To Me.FlexDetalle.Rows - 1
            Set rs = oSubasta.GetProvSubasta(Me.FlexDetalle.TextMatrix(i, 1), Me.FlexDetalle.TextMatrix(i, 2))
            
            lnMontoProvAnual = 0
            lnMontoProvAnterior = 0
            
            While Not rs.EOF
                If CCur(rs!Anio) = Year(gdFecSis) Then
                    lnMontoProvAnual = lnMontoProvAnual + rs!Monto
                Else
                    lnMontoProvAnterior = lnMontoProvAnterior + rs!Monto
                End If
                rs.MoveNext
            Wend
            
            lsCta = oMov.GetOpeCtaCtaOtro(gnAlmaIngXAdjudicacion, "", GetCtaCntBS(Me.FlexDetalle.TextMatrix(i, 2), gnAlmaIngXAdjudicacion), True)
            oMov.InsertaMovCta lnMovNro, lnContador, lsCta, lnMontoProvAnual * CCur(Me.FlexDetalle.TextMatrix(i, 5)) * -1
            lnContador = lnContador + 1
            
            If lnMontoProvAnterior <> 0 Then
                lsCta = oOpe.EmiteOpeCta(lsOpeCod, "H", 4)
                oMov.InsertaMovCta lnMovNro, lnContador, lsCta, lnMontoProvAnual * CCur(Me.FlexDetalle.TextMatrix(i, 5)) * -1
                lnContador = lnContador + 1
            End If
        Next i
        
        lsCta = oOpe.EmiteOpeCta(lsOpeCod, "H", 1)
        oMov.InsertaMovCta lnMovNro, lnContador, lsCta, Format(CCur(Me.lblTotalG.Caption) * -1 * (1 - lnPorIGV), "#.00")
        lsCta = oOpe.EmiteOpeCta(lsOpeCod, "H", 2)
        oMov.InsertaMovCta lnMovNro, lnContador + 1, lsCta, Format(CCur(Me.lblTotalG.Caption) * -1 * lnPorIGV, "#.00")
        lsCta = oOpe.EmiteOpeCta(lsOpeCod, "H", 3)
        oMov.InsertaMovCta lnMovNro, lnContador + 2, lsCta, Format((CCur(Me.lblTotalMG.Caption) - CCur(Me.lblTotalG.Caption)) * -1, "#.00")
    oMov.CommitTrans
    
    Me.txtNumDoc.Text = odoc.GeneraDocNro(lsDocTpo, gMonedaExtranjera, oGen.GetOpeSerie(lsOpeCod))
  '5101
    
    If MsgBox("¿ Desea Ingresar otro Documento ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Me.txtPersona.Text = ""
        Me.lblPerG.Caption = ""
        Me.lblDirG.Caption = ""
        Me.lblDNIG.Caption = ""
        Me.lblFonoG.Caption = ""
        
        Me.FlexDetalle.Clear
        Me.FlexDetalle.FormaCabecera
        Me.FlexDetalle.Rows = 2
        
        Me.lblTotalG.Caption = "0.00"
        Me.lblTotalMG.Caption = "0.00"
        
        Exit Sub
    End If
     
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdAgregar_Click()
    Me.FlexDetalle.AdicionaFila
End Sub

Private Sub CmdCancelar_Click()
    If MsgBox("Desea Salir sin grabar? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea eliminar el Item ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Me.FlexDetalle.EliminaFila FlexDetalle.Row
End Sub

Private Sub FlexDetalle_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim i As Integer
    Dim lnMonto As Currency
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    
    If FlexDetalle.TextMatrix(pnRow, 1) <> "" And pnCol = 1 Then
        If Not oSubasta.ValidaLote(lnMovNroSubActual, FlexDetalle.TextMatrix(pnRow, 1)) Then
            MsgBox "Debe ingresar un código de lote valido.", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
        FlexDetalle.TextMatrix(pnRow, 2) = ""
        FlexDetalle.TextMatrix(pnRow, 3) = ""
        FlexDetalle.TextMatrix(pnRow, 4) = ""
        FlexDetalle.TextMatrix(pnRow, 5) = ""
        FlexDetalle.TextMatrix(pnRow, 6) = ""
        FlexDetalle.TextMatrix(pnRow, 7) = ""
    End If
    
    If FlexDetalle.TextMatrix(pnRow, 5) <> "" And FlexDetalle.TextMatrix(pnRow, 2) <> "" Then
        If CCur(FlexDetalle.TextMatrix(pnRow, 5)) > CCur(FlexDetalle.TextMatrix(pnRow, 7)) Then
            Cancel = False
            Exit Sub
        End If
        FlexDetalle.TextMatrix(pnRow, 6) = Format(FlexDetalle.TextMatrix(pnRow, 4) * FlexDetalle.TextMatrix(pnRow, 5), "#,##0.0")
    End If
    
    lnMonto = 0
    For i = 1 To Me.FlexDetalle.Rows - 1
        If IsNumeric(FlexDetalle.TextMatrix(i, 6)) Then
            lnMonto = lnMonto + CCur(FlexDetalle.TextMatrix(i, 6))
        End If
    Next i

    Me.lblTotalG.Caption = Format(lnMonto, "#,##0.00")
    Me.lblTotalMG.Caption = Format(lnMonto * (1 + lnPorMartillero), "#,##0.00")
End Sub

Private Sub FlexDetalle_RowColChange()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lnSubasta As Long
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    
    
    If Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1) <> "" Then
        Set rs = oSubasta.GetBSLote(lnMovNroSubActual, Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1))
        Me.FlexDetalle.rsTextBuscar = rs
    End If
    
    If Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1) <> "" And Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 2) <> "" Then
        Set rs = oSubasta.GetSubastaSockPrecio(lnMovNroSubActual, Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1), Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 2))
        Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 4) = Format(rs.Fields(1), "#,##0.00")
        Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 7) = Format(rs.Fields(0), "#,##0.00")
        Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 8) = Format(rs.Fields(2), "#,##0.00")
    End If
    
    Set oSubasta = Nothing
    Set oALmacen = Nothing
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    Dim oGen As DLogGeneral
    Set oGen = New DLogGeneral
    
    Dim ofun As NContFunciones
    Set ofun = New NContFunciones
    
    
    lnPorMartillero = oGen.CargaParametro(5000, 1003) / 100
    lnPorIGV = oGen.CargaParametro(5000, 1004) / 100
    
    
    lnMovNroSubActual = oSubasta.GetUltimaSubasta(lsSubActual)
    
    
    lsDocTpo = oGen.GetOpeDocTpo(lsOpeCod)
    
    Me.txtNumDoc.Text = ofun.GeneraDocNro(lsDocTpo, gMonedaExtranjera, oGen.GetOpeSerie(lsOpeCod))
    
    If lbBoleta Then
        Me.lblTit.Caption = "Boleta de Venta :"
    Else
        Me.lblTit.Caption = "Factura de Venta :"
    End If
    
    Caption = lsCaption
    
    Me.FlexDetalle.rsTextBuscar = oALmacen.CargaBSSubasta
    
    Me.mskFecha = Format(gdFecSis, gsFormatoFechaView)
End Sub

Private Sub txtPersona_EmiteDatos()
    Dim oPersona As UPersona
    Set oPersona = New UPersona
    
    Me.lblPerG.Caption = txtPersona.psDescripcion
    
    If txtPersona.Text <> "" Then
        oPersona.ObtieneClientexCodigo Me.txtPersona.Text
        
        Me.lblDirG.Caption = oPersona.sPersDireccDomicilio
        Me.lblDNIG.Caption = oPersona.sPersIdnroDNI
        Me.lblRucG.Caption = oPersona.sPersIdnroRUC
        Me.lblFonoG.Caption = oPersona.sPersTelefono
    End If
    
End Sub
