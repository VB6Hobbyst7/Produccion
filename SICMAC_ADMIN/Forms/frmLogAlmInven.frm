VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmLogAlmInven 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Almacén : Inventario"
   ClientHeight    =   5670
   ClientLeft      =   1245
   ClientTop       =   2235
   ClientWidth     =   10740
   Icon            =   "frmLogAlmInven.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdValida 
      Caption         =   "&Validar"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   5220
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar Barrita 
      Height          =   195
      Left            =   1740
      TabIndex        =   13
      Top             =   4980
      Visible         =   0   'False
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Exportar a Excel"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   5220
      Width           =   1455
   End
   Begin VB.CommandButton cmdCodBarra 
      Caption         =   "&Código de Barras"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5220
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   3675
      Left            =   120
      TabIndex        =   9
      Top             =   1260
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6482
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      HighLight       =   2
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
      _Band(0).Cols   =   8
   End
   Begin VB.Frame fraOpcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   10515
      Begin VB.CheckBox chkStockMin 
         Appearance      =   0  'Flat
         Caption         =   "Stock Minimo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   270
         Width           =   1440
      End
      Begin VB.CheckBox chkStockMayor 
         Appearance      =   0  'Flat
         Caption         =   "Stock Mayor a 0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   540
         Width           =   1560
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   7020
         TabIndex        =   14
         Top             =   420
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboAlmFue 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   3060
      End
      Begin VB.ComboBox cboTpoAlm 
         Height          =   315
         ItemData        =   "frmLogAlmInven.frx":08CA
         Left            =   1020
         List            =   "frmLogAlmInven.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   645
         Width           =   3075
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   8760
         TabIndex        =   4
         Top             =   420
         Width           =   1425
      End
      Begin VB.CheckBox chkOpc 
         Appearance      =   0  'Flat
         Caption         =   "Saldos en Cero"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   810
         Width           =   1440
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   6360
         TabIndex        =   15
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Lista de"
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
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   705
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdImprime 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   8100
      TabIndex        =   1
      Top             =   5220
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9420
      TabIndex        =   0
      Top             =   5220
      Width           =   1245
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "Exportando a Excel"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   4980
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "frmLogAlmInven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

'TORE 20190810**************************
Dim appExcel As Excel.Application
Dim wbExcel As Excel.Workbook
Dim wsHojas As Excel.Worksheet
'***************************************

Private Sub cboAlmFue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTpoAlm.SetFocus
    End If
End Sub

Private Sub cboTpoAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFecha.SetFocus
    End If
End Sub

Private Sub cmdCodBarra_Click()
Dim I As Integer
I = MSFlex.row
'frmLogAlmCodBarra.Codigo MSFlex.TextMatrix(i, 2), MSFlex.TextMatrix(i, 3)
End Sub


Private Sub cmdExcel_Click()
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Set fs = New Scripting.FileSystemObject
Set appExcel = New Excel.Application
Dim clsDAlm As New DLogAlmacen 'TORE


'TORE
Dim lsArchivo As String
Dim lsFile As String
Dim lsNombreHoja As String
Dim lbExisteHoja As Boolean
Dim hHoraCrea  As String
Dim lnValorIni As Integer
Dim nValor As Integer
Dim lsTipDoc As String
On Error GoTo ErrorGenerar


hHoraCrea = Format(Time, "yyyyMMdd") & CStr(Hour(Time) & Minute(Time) & Second(Time))
lsNombreHoja = "Inventario"
lsFile = "FormatoInventario"

If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xlsx") Then
    Set wbExcel = appExcel.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xlsx")
    lsTipDoc = ".xlsx"
ElseIf fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
    Set wbExcel = appExcel.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    lsTipDoc = ".xls"
Else
    MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
    Exit Sub
End If

lsArchivo = "\Spooler\" & "InventarioAgencia" & hHoraCrea & lsTipDoc

'END TORE

For Each wsHojas In wbExcel.Worksheets
    If wsHojas.Name = lsNombreHoja Then
        wsHojas.Activate
        lbExisteHoja = True
        Exit For
    End If
Next

If lbExisteHoja = False Then
    Set wsHojas = wbExcel.Worksheets
    wsHojas.Name = lsNombreHoja
End If

lblMsg.Caption = "Exportando a EXCEL"
lblMsg.Visible = True
Barrita.Visible = True
DoEvents

 If chkStockMin.value = 1 Then
    nValor = 1
 End If
 If chkStockMayor.value = 1 Then
    nValor = 2
 End If
 If chkStockMin.value = 1 And chkStockMayor.value = 1 Then
    nValor = 3
 End If
 
lnValorIni = 2
Set rs = clsDAlm.CargaAlmacenBS(Val(Right(cboAlmFue.Text, 2)), Val(Right(Me.cboTpoAlm.Text, 5)), CDate(Me.mskFecha.Text), IIf(chkOpc.value = 1, True, False), , nValor)
Barrita.Max = rs.RecordCount

wsHojas.Cells(1, 1) = "INVENTARIO " & UCase(Left(Me.cboAlmFue.Text, 100))
Do While Not rs.EOF
    lnValorIni = lnValorIni + 1
    
    wsHojas.Cells(lnValorIni, 1) = rs!cBSCod
    wsHojas.Cells(lnValorIni, 2) = rs!cBSDescripcion
    wsHojas.Cells(lnValorIni, 3) = rs!nAlmBSStock
    wsHojas.Cells(lnValorIni, 4) = rs!cConsUnidad
    wsHojas.Cells(lnValorIni, 5) = rs!nAlmBSPrePromedio
    wsHojas.Cells(lnValorIni, 6) = rs!TotalSaldo
    wsHojas.Cells(lnValorIni, 7) = rs!cCtaContCod
    wsHojas.Cells(lnValorIni, 8) = rs!nStockMinimo
    wsHojas.Range("A" & Trim(Str(lnValorIni)) & ":" & "H" & Trim(Str(lnValorIni))).Borders.LineStyle = 1
    
    Barrita.value = IIf(lnValorIni <= Barrita.Max, lnValorIni, 0)
    rs.MoveNext
Loop

Barrita.Visible = False
lblMsg.Visible = False

wsHojas.SaveAs App.path & lsArchivo
appExcel.Visible = True
appExcel.Windows(1).Visible = True

Set appExcel = Nothing
Set wbExcel = Nothing
Set wsHojas = Nothing

        'ARLO 20160126 ***
        gsopecod = LogPistaInventarioAlmacen
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Genero Reporte de Inventario del Almacen a la Fecha : " & mskFecha.Text
        Set objPista = Nothing
        '**************
ErrorGenerar:

End Sub


Private Sub cmdImprime_Click()
    Dim lsCadena As String
    Dim clsDAlm As DLogAlmacen
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Set clsDAlm = New DLogAlmacen
    Dim I As Integer
    Dim nValor As Integer
    
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Fecha No valida.", vbInformation, "Aviso"
        mskFecha.SetFocus
        Exit Sub
    End If
'    lsCadena = ""
'    lsCadena = "Código" & Space(8) & "Descripción" & Space(14) & "Stock" & Space(6) & "Unidad" & Space(6) & "Precio Prom." & Space(6) & "Total" & Space(6) & "Cta. Cont" & Space(10) & "Stock Minimo" & oImpresora.gPrnSaltoLinea
'    For i = 1 To MSFlex.Rows - 1
'       lsCadena = lsCadena & MSFlex.TextMatrix(i, 2) & Space(5) & MSFlex.TextMatrix(i, 3) & Space(8) & MSFlex.TextMatrix(i, 4) & Space(8) & MSFlex.TextMatrix(i, 5) & Space(8) & MSFlex.TextMatrix(i, 6) & Space(8) & MSFlex.TextMatrix(i, 7) & Space(8) & MSFlex.TextMatrix(i, 8) & Space(8) & MSFlex.TextMatrix(i, 9) & oImpresora.gPrnSaltoLinea
'    Next
    If chkStockMin.value = 1 Then
        nValor = 1
    End If
    If chkStockMayor.value = 1 Then
        nValor = 2
    End If
    If chkStockMin.value = 1 And chkStockMayor.value = 1 Then
        nValor = 3
    End If
    
    
    lsCadena = clsDAlm.ReporteAlmacenBS(Left(Me.cboAlmFue.Text, 30), gsEmpresa, gdFecSis, Val(Right(cboAlmFue.Text, 2)), Val(Right(Me.cboTpoAlm.Text, 5)), CDate(Me.mskFecha.Text), IIf(chkOpc.value = 1, True, False), , nValor)
    
    oPrevio.Show lsCadena, Caption, True, 66
    
        'ARLO 20160126 ***
        gsopecod = LogPistaInventarioAlmacen
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, "", "", "Se Imprimio el Reporte de Inventario del Almacen a la Fecha : " & mskFecha.Text
        Set objPista = Nothing
        '**************
End Sub

Private Sub cmdProcesar_Click()
    Dim rs As ADODB.Recordset
    Dim clsDAlm As DLogAlmacen
    Dim I As Integer, nBSTpo As Integer, cDocNro As String
    Dim nValor As Integer
    
    
    Set rs = New ADODB.Recordset
    Set clsDAlm = New DLogAlmacen
     
    If Trim(cboAlmFue.Text) = "" Then Exit Sub
    
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Fecha No valida.", vbInformation, "Aviso"
        mskFecha.SetFocus
        Exit Sub
    End If
    
    If chkStockMin.value = 1 Then
        nValor = 1
    End If
    If chkStockMayor.value = 1 Then
        nValor = 2
    End If
    If chkStockMin.value = 1 And chkStockMayor.value = 1 Then
        nValor = 3
    End If
    
'    If OptAmbos.value = True Then
'        nValor = 3
'    End If
'
    'fgeAlmFue.Clear
    'fgeAlmFue.FormaCabecera
    'fgeAlmFue.Rows = 2
    
    'Detalle de bienes/servicios del almacen seleccionado
    'Set rs = clsDAlm.CargaAlmacenBS(Val(Right(cboAlmFue.Text, 2)), Val(Right(Me.cboTpoAlm.Text, 5)), IIf(chkOpc.value = 1, True, False))
    
'    If rs.RecordCount > 0 Then
'        Set fgeAlmFue.Recordset = rs
'        fgeAlmFue.lbEditarFlex = True
'    End If
    nBSTpo = Val(Right(Me.cboTpoAlm.Text, 5))
    LimpiaFlex nBSTpo
    DoEvents
    I = 0
   'VALIDA SI YA HAY UN CIERRE DEL MES ANTERIO ANPS
    Set rs = clsDAlm.VerificarCierrecBSCod(CDate(Me.mskFecha.Text), Val(Right(cboAlmFue.Text, 2)), Val(Right(Me.cboTpoAlm.Text, 5)))
If Not rs.EOF Then
If rs!cBSCod = 0 Then
Exit Sub
End If
End If
'fin valida anps
    Set rs = clsDAlm.CargaAlmacenBS(Val(Right(cboAlmFue.Text, 2)), Val(Right(Me.cboTpoAlm.Text, 5)), CDate(Me.mskFecha.Text), IIf(chkOpc.value = 1, True, False), , nValor)
    If Not rs.EOF Then
       If nBSTpo = 1 Then
          cDocNro = ""
          Do While Not rs.EOF
             i = i + 1
             InsRow MSFlex, i
             If cDocNro <> rs!cDocNro Then
                MSFlex.TextMatrix(i, 1) = rs!cDocNro
                cDocNro = rs!cDocNro
             End If
             MSFlex.TextMatrix(i, 2) = rs!cBSCod
             MSFlex.TextMatrix(i, 3) = rs!cBSDescripcion
             MSFlex.TextMatrix(i, 4) = rs!nAlmBSStock
             MSFlex.TextMatrix(i, 5) = rs!cConsUnidad
             MSFlex.TextMatrix(i, 6) = Format(rs!nAlmBSPrePromedio, "#,###,##0.00")
             MSFlex.TextMatrix(i, 7) = Format(rs!TotalSaldo, "#,###,##0.00")
             MSFlex.TextMatrix(i, 8) = rs!cCtaContCod & ""
             MSFlex.TextMatrix(i, 9) = rs!nStockMinimo
             rs.MoveNext
          Loop
       Else
          Do While Not rs.EOF
             i = i + 1
             InsRow MSFlex, i
             MSFlex.TextMatrix(i, 2) = rs!cBSCod
             MSFlex.TextMatrix(i, 3) = rs!cBSDescripcion
             MSFlex.TextMatrix(i, 4) = rs!nAlmBSStock
             MSFlex.TextMatrix(i, 5) = rs!cConsUnidad
             MSFlex.TextMatrix(i, 6) = Format(rs!nAlmBSPrePromedio, "#,###,##0.00")
             MSFlex.TextMatrix(i, 7) = Format(rs!TotalSaldo, "#,###,##0.00")
             MSFlex.TextMatrix(i, 8) = rs!cCtaContCod & ""
             MSFlex.TextMatrix(i, 9) = rs!nStockMinimo
             rs.MoveNext
          Loop
       End If
       cmdImprime.Enabled = True
    End If

    
    Set clsDAlm = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdValida_Click()
    Dim lsCadena As String
    Dim clsDAlm As DLogAlmacen
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Set clsDAlm = New DLogAlmacen
    
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Fecha No valida.", vbInformation, "Aviso"
        mskFecha.SetFocus
        Exit Sub
    End If
    
    lsCadena = clsDAlm.ReporteAlmacenBS(Left(Me.cboAlmFue.Text, 30), gsEmpresa, gdFecSis, Val(Right(cboAlmFue.Text, 2)), Val(Right(Me.cboTpoAlm.Text, 5)), CDate(Me.mskFecha.Text), IIf(chkOpc.value = 1, True, False), True)
    
    oPrevio.Show lsCadena, Caption, True, 66
End Sub

Private Sub Form_Load()
    Dim clsDAlm As DLogAlmacen
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oCon.GetConstante(5010, False)
    CargaCombo rs, cboTpoAlm
    cboTpoAlm.ListIndex = 0
    LimpiaFlex 0
    Call CentraForm(Me)
    

    Set clsDAlm = New DLogAlmacen
    
    Set rs = clsDAlm.CargaAlmacen(ATodos)
    If rs.RecordCount > 0 Then
        CargaCombo rs, cboAlmFue
    End If
    
    Set clsDAlm = Nothing
    Set rs = Nothing
    
    Me.mskFecha.Text = Format(gdFecSis, gsFormatoFechaView)
End Sub

Sub LimpiaFlex(nTipo As Integer)

MSFlex.Clear
MSFlex.Cols = 10
MSFlex.Rows = 2
MSFlex.RowHeight(1) = 8
MSFlex.RowHeight(0) = 320
MSFlex.ColWidth(0) = 0
If nTipo = 1 Then
   MSFlex.ColWidth(1) = 1050
   MSFlex.ColWidth(3) = 2800
Else
   MSFlex.ColWidth(1) = 0
   MSFlex.ColWidth(3) = 3850:
End If
MSFlex.ColAlignment(1) = 4: MSFlex.TextMatrix(0, 1) = "      LOTE"
MSFlex.ColWidth(2) = 850:   MSFlex.TextMatrix(0, 2) = "   Código":  MSFlex.ColAlignment(2) = 4
  MSFlex.TextMatrix(0, 3) = " Descripción"
MSFlex.ColWidth(4) = 800:   MSFlex.TextMatrix(0, 4) = "   Stock": MSFlex.ColAlignment(4) = 4
MSFlex.ColWidth(5) = 1000:  MSFlex.TextMatrix(0, 5) = " Unidad"
MSFlex.ColWidth(6) = 1000:  MSFlex.TextMatrix(0, 6) = " Precio Prom."
MSFlex.ColWidth(7) = 1000:  MSFlex.TextMatrix(0, 7) = " Total"
MSFlex.ColWidth(8) = 1200:  MSFlex.TextMatrix(0, 8) = " Cta. Cont"
MSFlex.ColWidth(9) = 1000:  MSFlex.TextMatrix(0, 9) = " Stock Minimo"

End Sub


Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 50
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub
