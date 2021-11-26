VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAnexo15BPosicionLiquidez 
   Caption         =   "Anexo 15B: Posición Mensual de Liquidez"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   Icon            =   "frmAnexo15BPosicionLiquidez.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRep 
      Height          =   1305
      Left            =   150
      TabIndex        =   15
      Top             =   30
      Width           =   6915
      Begin VB.Frame Frame3 
         Caption         =   "Periodo"
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
         Height          =   795
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Width           =   6525
         Begin VB.TextBox txtAnio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1230
            MaxLength       =   4
            TabIndex        =   18
            Top             =   300
            Width           =   855
         End
         Begin VB.ComboBox CboMes 
            Height          =   315
            ItemData        =   "frmAnexo15BPosicionLiquidez.frx":030A
            Left            =   3750
            List            =   "frmAnexo15BPosicionLiquidez.frx":0332
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   300
            Width           =   2460
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Año :"
            Height          =   195
            Left            =   540
            TabIndex        =   20
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes :"
            Height          =   195
            Left            =   2970
            TabIndex        =   19
            Top             =   360
            Width           =   390
         End
      End
   End
   Begin TabDlg.SSTab sTab 
      Height          =   3585
      Left            =   150
      TabIndex        =   12
      Top             =   1470
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   6324
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Conceptos "
      TabPicture(0)   =   "frmAnexo15BPosicionLiquidez.frx":039A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraConcepto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraConcepto 
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
         Height          =   3165
         Left            =   60
         TabIndex        =   13
         Top             =   330
         Width           =   6795
         Begin VB.CommandButton cmdNuevoConcep 
            Caption         =   "&Nuevo"
            Height          =   315
            Left            =   150
            TabIndex        =   5
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdModificaConcep 
            Caption         =   "&Modificar"
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminaConcep 
            Caption         =   "&Eliminar"
            Height          =   315
            Left            =   2010
            TabIndex        =   7
            Top             =   2790
            Width           =   915
         End
         Begin VB.CommandButton cmdGrabaConcep 
            Caption         =   "&Grabar"
            Height          =   315
            Left            =   4800
            TabIndex        =   8
            Top             =   2790
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.CommandButton cmdCancelaConcep 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   5730
            TabIndex        =   9
            Top             =   2790
            Visible         =   0   'False
            Width           =   915
         End
         Begin MSComctlLib.ListView lvConcep 
            Height          =   2475
            Left            =   150
            TabIndex        =   0
            Top             =   255
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   4366
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clase"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Concep"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descripción"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Fórmula"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Frame fraDatConcep 
            Height          =   525
            Left            =   150
            TabIndex        =   14
            Top             =   2220
            Visible         =   0   'False
            Width           =   6495
            Begin VB.TextBox txtCtaCod 
               Height          =   315
               Left            =   4320
               MaxLength       =   255
               TabIndex        =   4
               Top             =   150
               Width           =   2085
            End
            Begin VB.TextBox txtClase 
               Height          =   315
               Left            =   60
               MaxLength       =   1
               TabIndex        =   1
               Top             =   150
               Width           =   315
            End
            Begin VB.TextBox txtConcep 
               Height          =   315
               Left            =   360
               MaxLength       =   2
               TabIndex        =   2
               Top             =   150
               Width           =   435
            End
            Begin VB.TextBox txtConcepDesc 
               Height          =   315
               Left            =   810
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   3
               Top             =   150
               Width           =   3495
            End
         End
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   345
      Left            =   4500
      TabIndex        =   10
      Top             =   5160
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   5700
      TabIndex        =   11
      Top             =   5160
      Width           =   1155
   End
End
Attribute VB_Name = "frmAnexo15BPosicionLiquidez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim sSql As String
Dim lNuevo As Boolean
Dim nColRango As Integer
Dim nFil    As Integer
Dim nTpoCambioAj As Currency
Dim nTpoCambioF  As Currency

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim oCon As DConecta

Dim nActS() As Currency
Dim nPasS() As Currency
Dim nTotS() As Currency

Dim nActD() As Currency
Dim nPasD() As Currency
Dim nTotD() As Currency

Dim nPromA() As Currency
Dim nPromP() As Currency

Dim nCantAct As Integer
Dim nCantPas As Integer
Dim nDias As Integer

Dim nActualAct As Integer
Dim nActualPas As Integer
Dim bandera As Integer

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub

Private Sub cmdCancelaConcep_Click()
HabilitaConcepto False
lvConcep.SetFocus
End Sub

Private Sub cmdEliminaConcep_Click()
If lvConcep.ListItems.Count = 0 Then
   Exit Sub
End If
If MsgBox("¿ Seguro que desea Eliminar Concepto ?", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
sSql = "DELETE Anx15BConcepto WHERE cOpeCod = '" & gsOpeCod & "' and cCodClase = '" & lvConcep.SelectedItem.Text & "' and cCodConcep = '" & lvConcep.SelectedItem.SubItems(2) & "'"
oCon.Ejecutar sSql
lvConcep.ListItems.Remove lvConcep.SelectedItem.Index

End Sub

Private Sub cmdGenerar_Click()
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lsRuta      As String
Dim lbLibroOpen As Boolean
Dim N           As Integer

On Error GoTo ErrImprime

MousePointer = 11
If Not ValidaDatos Then
   MousePointer = 0
   Exit Sub
End If
'-LIMPIAMOS EL TEMPORAL DE FECHA TMPFECHA
gdFecha = CDate("01/" & Format(CboMes.ListIndex + 1, "00") & "/" & txtAnio)

'sSql = "DELETE TmpFecha WHERE datedIff(m,dFecha, '" & Format(gdFecha, gsFormatoFecha) & "') = 0 "
'oCon.Ejecutar sSql
'
'For N = 1 To Day(DateAdd("m", 1, gdFecha) - 1)
'    sSql = "INSERT tmpFecha VALUES ('" & Format(gdFecha + N - 1, gsFormatoFecha) & "')"
'    oCon.Ejecutar sSql
'Next
gdFecha = DateAdd("m", 1, gdFecha) - 1
Me.Enabled = False

If bandera = 1 Then
    lsRuta = App.path & "\Spooler\"
    lsArchivo = lsRuta & "ANX15BLIQUIDEZ" & "_" & txtAnio & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If lbLibroOpen Then
      Set xlHoja1 = xlLibro.Worksheets(1)
      
      ExcelAddHoja CboMes, xlLibro, xlHoja1, False
       
      Call CabeceraExcel
      nFil = 8
      
      Call ImprimeConceptos(bandera, "1", nFil)
      nFil = nFil + 2
      Call ImprimeConceptos(bandera, "2", nFil)
      
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
      CargaArchivo lsArchivo, lsRuta
        
      MsgBox "Reporte Generado satisfactoriamente", vbInformation, "Aviso!!!"
   End If
Else
      nActualAct = 0
      nActualPas = 0
     
      nFil = 8
      
      Call ImprimeConceptos(bandera, "1", nFil)
      nFil = nFil + 2
      
      nActualAct = 0
      nActualPas = 0
    
      Call ImprimeConceptos(bandera, "2", nFil)
      
      CalculaArreglos gdFecha
      MsgBox "Reporte SUCAVE Generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
End If
Me.Enabled = True
MousePointer = 0
 
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   MousePointer = 0
   Me.Enabled = True
End Sub

Private Sub cmdGrabaConcep_Click()
Dim nPos As Integer
If Not ValidaDatosConcep() Then
   Exit Sub
End If
nPos = 1
If lvConcep.ListItems.Count > 0 Then
    nPos = lvConcep.SelectedItem.Index
End If
If lNuevo Then
   sSql = "INSERT Anx15BConcepto (cOpeCod, cCodClase, cCodConcep, cDescrip, cFormula) " _
        & "VALUES ('" & gsOpeCod & "','" & txtClase & "','" & txtConcep & "','" & txtConcepDesc & "','" & txtCtaCod & "')"
Else
   sSql = "UPDATE Anx15BConcepto SET cDescrip = '" & txtConcepDesc & "', cFormula = '" & txtCtaCod & "' WHERE cOpeCod = '" & gsOpeCod & "' and cCodClase = '" & txtClase & "' and cCodConcep = '" & txtConcep & "' "
End If
oCon.Ejecutar sSql
HabilitaConcepto False
CargaConceptos
lvConcep.ListItems(nPos).Selected = True
lvConcep.SetFocus
End Sub

Private Sub cmdModificaConcep_Click()
lNuevo = False
HabilitaConcepto True
txtConcepDesc.SetFocus
End Sub

Private Sub cmdNuevoConcep_Click()
lNuevo = True
HabilitaConcepto True
txtClase.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
frmReportes.Enabled = False
Set oCon = New DConecta
oCon.AbreConexion
Me.Caption = gsOpeDesc
CargaConceptos
txtAnio = Year(gdFecSis)
CboMes.ListIndex = Month(gdFecSis) - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
frmReportes.Enabled = True
End Sub
 
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
   If KeyAscii = 13 Then
      CboMes.SetFocus
   End If
End Sub

Private Sub txtClase_GotFocus()
fEnfoque txtClase
End Sub

Private Sub txtClase_KeyPress(KeyAscii As Integer)
If InStr("12F", Chr(KeyAscii)) = 0 And Not KeyAscii = 13 And Not KeyAscii = 8 Then
   KeyAscii = 0
End If
If KeyAscii = 13 Then
   txtConcep.SetFocus
End If
End Sub

Private Sub txtConcep_GotFocus()
fEnfoque txtConcep
End Sub

Private Sub txtConcep_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtConcepDesc.SetFocus
End If
End Sub

Private Sub txtConcepDesc_GotFocus()
fEnfoque txtConcepDesc
End Sub

Private Sub txtConcepDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtCtaCod.SetFocus
End If
End Sub

Private Sub txtCtaCod_KeyPress(KeyAscii As Integer)
If InStr("0123456789_[]%", Chr(KeyAscii)) = 0 And Not KeyAscii = 13 And Not KeyAscii = 8 Then
   Exit Sub
End If
If KeyAscii = 13 Then
   cmdGrabaConcep.SetFocus
End If
End Sub

Private Sub HabilitaConcepto(lActiva As Boolean)
cmdGrabaConcep.Visible = lActiva
cmdCancelaConcep.Visible = lActiva
fraDatConcep.Visible = lActiva
cmdNuevoConcep.Visible = Not lActiva
cmdModificaConcep.Visible = Not lActiva
cmdEliminaConcep.Visible = Not lActiva
cmdGenerar.Enabled = Not lActiva
fraRep.Enabled = Not lActiva
If lActiva Then
   lvConcep.Height = 1965
Else
   lvConcep.Height = 2445
End If
If lActiva Then
   txtClase.Enabled = lNuevo
   txtConcep.Enabled = lNuevo
   If lNuevo Then
      txtConcepDesc = ""
   Else
      txtClase = lvConcep.SelectedItem.Text
      txtConcep = lvConcep.SelectedItem.SubItems(1)
      txtConcepDesc = lvConcep.SelectedItem.SubItems(2)
      txtCtaCod = lvConcep.SelectedItem.SubItems(3)
   End If
End If
End Sub

Private Function ValidaDatosConcep() As Boolean
ValidaDatosConcep = True
If txtClase = "" Then
   MsgBox "Falta ingresar Clase de Concepto", vbInformation, "¡Aviso!"
   txtClase.SetFocus
   Exit Function
End If
If txtConcep = "" Then
   MsgBox "Falta ingresar Código de Concepto", vbInformation, "¡Aviso!"
   txtConcep.SetFocus
   Exit Function
End If
ValidaDatosConcep = True
End Function

Private Sub CargaConceptos()
Dim lvItm As ListItem
lvConcep.ListItems.Clear
sSql = "SELECT cCodClase, cCodConcep, cDescrip, ISNULL(cFormula,'') cFormula " _
     & "FROM Anx15BConcepto " _
     & "WHERE cOpeCod = '" & gsOpeCod & "'"
Set rs = oCon.CargaRecordSet(sSql)
Do While Not rs.EOF
   Set lvItm = lvConcep.ListItems.Add(, , rs!cCodClase)
   lvItm.SubItems(1) = rs!cCodConcep
   lvItm.SubItems(2) = rs!cDescrip
   lvItm.SubItems(3) = rs!cFormula
   rs.MoveNext
Loop
Set lvItm = Nothing
RSClose rs
End Sub
Private Function ValidaDatos() As Boolean
ValidaDatos = False
    If Len(Trim(txtAnio.Text)) = 0 Then
        MsgBox "Ingrese Año de Proceso", vbInformation, "¡Aviso!"
        txtAnio.SetFocus
        Exit Function
    End If
    If CInt(txtAnio.Text) < 1950 Then
        MsgBox "Año no valido menor a 1950 ", vbInformation, "¡Aviso!"
        txtAnio.SetFocus
        Exit Function
    End If
ValidaDatos = True
End Function

Private Sub CabeceraExcel()
Dim nCol As Integer
Dim sCol As String
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.Zoom = 55
xlHoja1.Cells(1, 1) = "SUPERINTENDENCIA DE BANCA Y SEGURO"
xlHoja1.Cells(2, 2) = "ANEXO 15B"
xlHoja1.Cells(3, 2) = "POSICION MENSUAL DE LIQUIDEZ"
xlHoja1.Cells(4, 2) = txtAnio & "/" & Trim(CboMes)
xlHoja1.Cells(4, 11) = "FECHA: " & Mid(gdFecha, 1, 2) & "/" & Format(CboMes, "00") & "/" & txtAnio
xlHoja1.Cells(4, 1) = "EMPRESA: " & gsNomCmac

xlHoja1.Range("B2:J2").Merge
xlHoja1.Range("B3:J3").Merge
xlHoja1.Range("B4:J4").Merge
xlHoja1.Range("B2:J4").HorizontalAlignment = xlHAlignCenter

xlHoja1.Range("B2:J2").Font.Size = 14

xlHoja1.Range("A1:A1").ColumnWidth = 50
nColRango = Day(gdFecha) + 2


End Sub

Private Function getTipCambioRep() As Boolean
Dim prs As ADODB.Recordset
Dim oTC As New nTipoCambio
Set prs = New ADODB.Recordset
getTipCambioRep = False
   nTpoCambioF = oTC.EmiteTipoCambio(gdFecha, TCFijoMes)
   sSql = "SELECT nMovOtroImporte FROM MovOtrosItem MO JOIN Mov M ON M.nMovNro = MO.nMovNro WHERE not M.nMovFlag IN ('1','2','3') and M.cMovNro LIKE '" & Format(gdFecha, "yyyymmdd") & "%' and cmovotrovariable = 'TC2'"
   Set prs = oCon.CargaRecordSet(sSql)
   If Not prs.EOF Then
      nTpoCambioAj = prs!nMovOtroImporte
   Else
      MsgBox "Aún no se genera Asiento de Ajuste por Tipo de Cambio. Se usaré el Tipo de Cambio Fijo definido en el Mes", vbInformation, "¡Aviso!"
      nTpoCambioAj = nTpoCambioF
   End If
   RSClose prs
getTipCambioRep = True
Set oTC = Nothing
End Function

Private Sub ImprimeConceptos(pnBandera As Integer, psMoneda As String, ByVal pnFilIni As Integer)
Dim sCodAnt As String
Dim nPosIni As Integer
Dim pnCol   As Integer
Dim sCol    As String
Dim nCol    As Integer
Dim sColFin As String
Dim sTotales As String
Dim N        As Integer
Dim lnDia As Integer
Dim lsSql As String
Dim rsPF As ADODB.Recordset
Set rsPF = New ADODB.Recordset

sSql = "SELECT cCodClase, cCodConcep, cDescrip, ISNULL(cFormula,'') cFormula FROM Anx15BConcepto WHERE cOpeCod = '" & gsOpeCod & "'"
Set rs = oCon.CargaRecordSet(sSql)

If pnBandera = 2 Then
    If nCantAct = 0 And nCantPas = 0 Then
        Do While Not rs.EOF
            If rs!cCodClase = 1 And Trim(rs!cFormula) <> "H" And Left(rs!cFormula, 3) <> "SUM" And Len(Trim(rs!cFormula)) > 0 Then
                nCantAct = nCantAct + 1
            ElseIf rs!cCodClase = 2 And Trim(rs!cFormula) <> "H" And Left(rs!cFormula, 3) <> "SUM" And Len(Trim(rs!cFormula)) > 0 Then
                nCantPas = nCantPas + 1
            End If
            rs.MoveNext
        Loop
        rs.MoveFirst
    End If
End If

sCodAnt = rs!cCodClase

If psMoneda = "2" Then
   If Not getTipCambioRep() Then
      Exit Sub
   End If
Else
    nTpoCambioAj = 1
    nTpoCambioF = 1
End If

If pnBandera = 1 Then
    xlHoja1.Cells(pnFilIni - 1, 1) = psMoneda & ". RATIO DE LIQUIDEZ MONEDA " & IIf(psMoneda = "1", "NACIONAL", "EXTRANJERA")
    For nCol = 2 To Day(gdFecha) + 1
        xlHoja1.Cells(pnFilIni - 1, nCol) = nCol - 1
    Next
    xlHoja1.Cells(pnFilIni - 1, nCol) = "Promedio"
    xlHoja1.Cells(pnFilIni, nCol) = "Mensual"
ElseIf pnBandera = 2 Then
    If nDias = 0 Then
        'Definimos el nro de dias
        nDias = Day(gdFecha)
    
        'Redefinimos los arreglos

        ReDim nActS(nCantAct, nDias) As Currency
        ReDim nPasS(nCantPas, nDias) As Currency
        ReDim nTotS(3, nDias) As Currency
        
        ReDim nActD(nCantAct, nDias) As Currency
        ReDim nPasD(nCantPas, nDias) As Currency
        ReDim nTotD(3, nDias) As Currency
        
        ReDim nPromA(2, nCantAct) As Currency
        ReDim nPromP(2, nCantPas) As Currency
    End If
     
End If
 
If pnBandera = 1 Then
    If nCol <= 26 Then
        sColFin = Chr(64 + nCol)
    Else
        sColFin = "A" + Chr(38 + nCol)
    End If
    xlHoja1.Range("A1:" & sColFin & "4").Font.Bold = True
    xlHoja1.Range("A" & pnFilIni - 1 & ":" & sColFin & pnFilIni).Font.Bold = True
    xlHoja1.Range("B" & pnFilIni - 1 & ":" & sColFin & pnFilIni).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & pnFilIni - 1 & ":" & sColFin & pnFilIni).BorderAround xlContinuous
    xlHoja1.Range("B" & pnFilIni - 1 & ":" & sColFin & pnFilIni).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range("B1:" & sColFin & "1").ColumnWidth = 14
End If

sTotales = ""
Do While Not rs.EOF
   If pnBandera = 1 Then
        If Not rs!cCodClase = sCodAnt Then
           xlHoja1.Range(xlHoja1.Cells(nFil - 1, 1), xlHoja1.Cells(nFil - 1, nColRango)).BorderAround xlContinuous
        End If
        xlHoja1.Cells(nFil, 1) = rs!cDescrip
    
        'Formulas de Totales
        If Left(rs!cFormula, 4) = "SUMA" Then
           nPosIni = Val(Mid(rs!cFormula, 6, Len(rs!cFormula) - 1))
           For pnCol = 2 To Day(gdFecha) + 2
             If pnCol <= 26 Then
                 sCol = Chr(64 + pnCol)
             Else
                 sCol = "A" + Chr(38 + pnCol)
             End If
             xlHoja1.Range(xlHoja1.Cells(nFil, pnCol), xlHoja1.Cells(nFil, pnCol)).Formula = "=SUM(" & sCol & nFil - nPosIni - 1 & ":" & sCol & nFil - 1 & ")"
             xlHoja1.Range(xlHoja1.Cells(nFil, pnCol), xlHoja1.Cells(nFil, pnCol)).Font.Bold = True
           Next
           sTotales = sTotales & "," & nFil
        ElseIf rs!cDescrip = "Obligaciones por Cuentas a Plazo" Then
            lsSql = " "
            lsSql = lsSql & " Select fec.dfecha, (Sum(nSaldCnt) + dbo.getsaldoctaacumulado(fec.dfecha,'2[13]" & Trim(Str(psMoneda)) & "803%'," & Trim(Str(psMoneda)) & ")) as nSaldoAC"
            lsSql = lsSql & " from capsaldosdiarios csd, dbo.fechatmp('" & Format(gdFecha, gsFormatoFecha) & "') fec"
            lsSql = lsSql & " Where csd.dFecha >= fec.dFecha And csd.dFecha < DateAdd(Day, 1, fec.dFecha)"
            lsSql = lsSql & "     And cctacod like '10___233" & Trim(Str(psMoneda)) & "%' and nPlazo <= 360"
            lsSql = lsSql & "     And cCtaCod Not In (Select cCtaCod from productobloqueos where cCtaCod like '_____233" & Trim(Str(psMoneda)) & "%'"
            lsSql = lsSql & "     And nBlqMotivo = 3 and cMovNro < Convert(Varchar(8),DateAdd(Day,1,fec.dfecha),112) And (cMovNroDbl Is Null Or cMovNroDbl > Convert(Varchar(8),DateAdd(Day,1,fec.dfecha),112)))"
            lsSql = lsSql & " Group By fec.dfecha order by fec.dfecha"
            Set rsPF = oCon.CargaRecordSet(lsSql)
           For pnCol = 2 To Day(gdFecha) + 2
             If Not rsPF.EOF Then
                xlHoja1.Range(xlHoja1.Cells(nFil, pnCol), xlHoja1.Cells(nFil, pnCol)).Formula = rsPF.Fields(1)
                rsPF.MoveNext
             End If
           Next
        
        ElseIf rs!cFormula = "H" Then
             xlHoja1.Range("A" & nFil & ":" & sColFin & nFil).BorderAround xlContinuous
        ElseIf Len(Trim(rs!cFormula)) > 0 Then
            
           CalculaSaldosCuenta pnBandera, rs!cCodClase, rs!cFormula, nFil, psMoneda
           
        End If
   
       nFil = nFil + 1
   ElseIf pnBandera = 2 Then
        If Left(rs!cFormula, 4) = "SUMA" Then
        ElseIf rs!cFormula = "H" Then
        ElseIf Len(Trim(rs!cFormula)) > 0 Then
            If rs!cCodClase = 1 Then
              nActualAct = nActualAct + 1
            Else
              nActualPas = nActualPas + 1
            End If
            CalculaSaldosCuenta pnBandera, rs!cCodClase, rs!cFormula, nFil, psMoneda
        End If
        nFil = nFil + 1
   End If
   
   sCodAnt = rs!cCodClase
   rs.MoveNext
Loop
RSClose rs

If pnBandera = 1 Then
    xlHoja1.Range(xlHoja1.Cells(nFil - 1, 1), xlHoja1.Cells(nFil - 1, nColRango)).BorderAround xlContinuous
    xlHoja1.Range(xlHoja1.Cells(pnFilIni - 1, 1), xlHoja1.Cells(nFil - 1, nColRango)).BorderAround xlContinuous
    xlHoja1.Range(xlHoja1.Cells(pnFilIni - 1, 1), xlHoja1.Cells(nFil - 1, nColRango)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(pnFilIni, 1), xlHoja1.Cells(nFil - 1, nColRango)).NumberFormat = "##,###,##0.00"
    xlHoja1.Range(xlHoja1.Cells(1, nColRango), xlHoja1.Cells(1, nColRango)).ColumnWidth = 14
 
    'Resumen Ratio
    For N = pnFilIni + 1 To nFil - 1
        If xlHoja1.Cells(N, 1) <> "" Then
            If nColRango <= 26 Then
                sCol = Chr(64 + nColRango - 1)
            Else
                sCol = "A" + Chr(38 + nColRango - 1)
            End If
            
            
            xlHoja1.Range(xlHoja1.Cells(N, nColRango), xlHoja1.Cells(N, nColRango)).Formula = "=COUNT(B" & N & ":" & sCol & N & ")"
            
            If xlHoja1.Cells(N, nColRango) = 0 Then
                xlHoja1.Cells(N, nColRango) = 0
            Else
                xlHoja1.Range(xlHoja1.Cells(N, nColRango), xlHoja1.Cells(N, nColRango)).Formula = "=AVERAGE(B" & N & ":" & sCol & N & ")"
            End If
            
        End If
    Next
    
    Dim nPos As Integer
    Dim nValor As Integer
    sTotales = Mid(sTotales, 2, Len(sTotales))
    xlHoja1.Cells(nFil, 1) = "Ratio de Liquidez " & IIf(psMoneda = "1", "M.N.", "M.E.") & " (a)/(b)*100"
    Do While Len(sTotales) > 0
        nPos = InStr(sTotales, ",")
        If nPos > 0 Then
            nValor = Val(Mid(sTotales, 1, nPos - 1))
            sTotales = Mid(sTotales, nPos + 1, Len(sTotales))
        Else
            nValor = Val(sTotales)
            sTotales = ""
        End If
        For N = 2 To Day(gdFecha) + 2
            If sTotales <> "" Then
                xlHoja1.Cells(nFil, N) = xlHoja1.Cells(nValor, N)
            Else
                If Val(xlHoja1.Cells(nValor, N)) <> 0 Then
                    xlHoja1.Cells(nFil, N) = Round(Val(xlHoja1.Cells(nFil, N)) * 100 / Val(xlHoja1.Cells(nValor, N)), 2)
                End If
            End If
        Next
    Loop
    xlHoja1.Range(xlHoja1.Cells(nFil, 1), xlHoja1.Cells(nFil, nColRango)).BorderAround xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nFil, 1), xlHoja1.Cells(nFil, nColRango)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nFil, 1), xlHoja1.Cells(nFil, nColRango)).Font.Bold = True
    
    For N = 1 To 3
        nFil = nFil + 1
        Select Case N
            Case 1: xlHoja1.Cells(nFil, 1) = "Activos Líquidos Ajustados por Recursos Préstado en " & IIf(psMoneda = "1", "M.N.", "M.E.") & " (c)"
            Case 2: xlHoja1.Cells(nFil, 1) = "Pasivos de Corto Plazo Ajustados por Recursos Prestados en " & IIf(psMoneda = "1", "M.N.", "M.E.") & " (d)"
            Case 3: xlHoja1.Cells(nFil, 1) = "Ratio de Liquidez Ajustado por Recursos Prestados en " & IIf(psMoneda = "1", "M.N.", "M.E.") & " [(c/d)]*100"
        End Select
        xlHoja1.Range(xlHoja1.Cells(nFil, 1), xlHoja1.Cells(nFil, nColRango)).BorderAround xlContinuous
        xlHoja1.Range(xlHoja1.Cells(nFil, 1), xlHoja1.Cells(nFil, nColRango)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(nFil, 1), xlHoja1.Cells(nFil, nColRango)).Font.Bold = True
    Next
    nFil = nFil + 2
End If
End Sub

Private Sub CalculaSaldosCuenta(ByVal pn_Bandera As Integer, ByVal psActPas As Integer, ByVal psFormula As String, ByVal pnFil As Integer, ByVal psMoneda As String)
Dim sCtaCod As String
Dim nPosSimbolo As Integer
Dim nImporte As Double
Dim sSimbolo As String
Dim sSimboloOpe As String
Dim prs As ADODB.Recordset
Dim prsIni As ADODB.Recordset
Dim sCta As String
Dim j As Integer
Dim N As Integer
nImporte = 0
sSimbolo = "+"
Do While Len(Trim(psFormula)) > 0
    If InStr(psFormula, "-") > 0 Then
        nPosSimbolo = IIf(InStr(psFormula, "+") < InStr(psFormula, "-"), InStr(psFormula, "+"), 0)
        If nPosSimbolo = 0 Then
            nPosSimbolo = InStr(psFormula, "-")
        End If
    Else
       nPosSimbolo = InStr(psFormula, "+")
    End If
    sSimboloOpe = sSimbolo
    If nPosSimbolo > 0 Then
        sCtaCod = Left(psFormula, nPosSimbolo - 1)
        sSimbolo = Mid(psFormula, nPosSimbolo, 1)
        psFormula = Mid(psFormula, nPosSimbolo + 1, Len(psFormula))
    Else
        sCtaCod = Trim(psFormula)
        psFormula = ""
    End If
    Dim ldFecha As Date
    ldFecha = CDate("01/" & Format(Month(gdFecha), "00") & "/" & Year(gdFecha))
    sSql = "SELECT '" & Format(ldFecha - 1, gsFormatoFecha) & "' dCtaSaldoFecha, ISNULL(SUM(nCtaSaldoImporte),0) nSaldo FROM CtaSaldo cs " _
         & "WHERE cCtaContCod LIKE '__" & psMoneda & "%' and cCtaContCod LIKE '" & sCtaCod & "%' and " _
         & "dCtaSaldoFecha = (SELECT max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= '" & Format(ldFecha - 1, gsFormatoFecha) & "') " _
         & "UNION " _
         & "SELECT Fec.dFecha, SUM(cs.nCtaSaldoImporte) nSaldo " _
         & "FROM FechaTmp('" & Format(ldFecha, gsFormatoFecha) & "') Fec, CtaSaldo cs " _
         & "WHERE Fec.dFecha BETWEEN '" & Format(ldFecha, gsFormatoFecha) & "' and '" & Format(gdFecha, gsFormatoFecha) & "' " _
         & "      and cCtaContCod LIKE '__" & psMoneda & "%' and cCtaContCod LIKE '" & sCtaCod & "%' " _
         & "      and cs.dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= Fec.dFecha) " _
         & "GROUP BY Fec.dFecha " _
         & "ORDER BY dCtaSaldoFecha"

    Set prsIni = oCon.CargaRecordSet(sSql)
    Dim nSaldo As Currency
    If Not prsIni.EOF Then
        nSaldo = prsIni!nSaldo
        prsIni.MoveNext
        Do While Not prsIni.EOF
            If ldFecha < prsIni!dCtaSaldofecha Then
                Do While ldFecha < prsIni!dCtaSaldofecha
                    N = Day(ldFecha) + 1
                    nImporte = nSaldo
                    If nImporte <> 0 Then
                        nImporte = Round(nImporte / IIf(psMoneda = "2", IIf(ldFecha = gdFecha, nTpoCambioAj, nTpoCambioF), 1), 2)
                    End If
                    If pn_Bandera = 1 Then
                        If sSimboloOpe = "+" Then
                            xlHoja1.Cells(pnFil, N) = Val(xlHoja1.Cells(pnFil, N)) + nImporte
                        Else
                            xlHoja1.Cells(pnFil, N) = Val(xlHoja1.Cells(pnFil, N)) - nImporte
                        End If
                    Else
                        'revisar
                        If sSimboloOpe = "+" Then
                            If psMoneda = "1" Then
                                If psActPas = 1 Then
                                    nActS(nActualAct, N - 1) = nActS(nActualAct, N - 1) + nImporte
                                Else
                                    nPasS(nActualPas, N - 1) = nPasS(nActualPas, N - 1) + nImporte
                                End If
                            ElseIf psMoneda = "2" Then
                                If psActPas = 1 Then
                                    nActD(nActualAct, N - 1) = nActD(nActualAct, N - 1) + nImporte
                                Else
                                    nPasD(nActualPas, N - 1) = nPasD(nActualPas, N - 1) + nImporte
                                End If
                            End If
                        Else
                            If psMoneda = "1" Then
                                If psActPas = 1 Then
                                    nActS(nActualAct, N - 1) = nActS(nActualAct, N - 1) - nImporte
                                Else
                                    nPasS(nActualPas, N - 1) = nPasS(nActualPas, N - 1) - nImporte
                                End If
                            ElseIf psMoneda = "2" Then
                                If psActPas = 1 Then
                                    nActD(nActualAct, N - 1) = nActD(nActualAct, N - 1) - nImporte
                                Else
                                    nPasD(nActualPas, N - 1) = nPasD(nActualPas, N - 1) - nImporte
                                End If
                            End If
                        End If
                        'revisar
                    End If
                    ldFecha = ldFecha + 1
                Loop
            End If
            N = Day(ldFecha) + 1
            nImporte = prsIni!nSaldo
            If nImporte <> 0 Then
                nImporte = Round(nImporte / IIf(psMoneda = "2", IIf(prsIni!dCtaSaldofecha = gdFecha, nTpoCambioAj, nTpoCambioF), 1), 2)
            End If
            If sSimboloOpe = "+" Then
                If pn_Bandera = 1 Then
                    xlHoja1.Cells(pnFil, N) = Val(xlHoja1.Cells(pnFil, N)) + nImporte
                ElseIf pn_Bandera = 2 Then
                 
                    If psMoneda = "1" Then
                        If psActPas = 1 Then
                            nActS(nActualAct, N - 1) = nActS(nActualAct, N - 1) + nImporte
                        Else
                            nPasS(nActualPas, N - 1) = nPasS(nActualPas, N - 1) + nImporte
                        End If
                    ElseIf psMoneda = "2" Then
                        If psActPas = 1 Then
                            nActD(nActualAct, N - 1) = nActD(nActualAct, N - 1) + nImporte
                        Else
                            nPasD(nActualPas, N - 1) = nPasD(nActualPas, N - 1) + nImporte
                        End If
                    End If
                End If
            
            Else
                If pn_Bandera = 1 Then
                    xlHoja1.Cells(pnFil, N) = Val(xlHoja1.Cells(pnFil, N)) - nImporte
                Else
                    If psMoneda = "1" Then
                        If psActPas = 1 Then
                            nActS(nActualAct, N - 1) = nActS(nActualAct, N - 1) - nImporte
                        Else
                            nPasS(nActualPas, N - 1) = nPasS(nActualPas, N - 1) - nImporte
                        End If
                    ElseIf psMoneda = "2" Then
                        If psActPas = 1 Then
                            nActD(nActualAct, N - 1) = nActD(nActualAct, N - 1) - nImporte
                        Else
                            nPasD(nActualPas, N - 1) = nPasD(nActualPas, N - 1) - nImporte
                        End If
                    End If
                End If
            End If
             
            ldFecha = prsIni!dCtaSaldofecha + 1
            nSaldo = prsIni!nSaldo
            prsIni.MoveNext
        Loop
        Do While ldFecha <= gdFecha
            N = Day(ldFecha) + 1
            nImporte = nSaldo
            If nImporte <> 0 Then
                nImporte = Round(nImporte / IIf(psMoneda = "2", IIf(ldFecha = gdFecha, nTpoCambioAj, nTpoCambioF), 1), 2)
            End If
            If sSimboloOpe = "+" Then
                If pn_Bandera = 1 Then
                    xlHoja1.Cells(pnFil, N) = Val(xlHoja1.Cells(pnFil, N)) + nImporte
                Else
                    If psMoneda = "1" Then
                        If psActPas = 1 Then
                            nActS(nActualAct, N - 1) = nActS(nActualAct, N - 1) + nImporte
                        Else
                            nPasS(nActualPas, N - 1) = nPasS(nActualPas, N - 1) + nImporte
                        End If
                    ElseIf psMoneda = "2" Then
                        If psActPas = 1 Then
                            nActD(nActualAct, N - 1) = nActD(nActualAct, N - 1) + nImporte
                        Else
                            nPasD(nActualPas, N - 1) = nPasD(nActualPas, N - 1) + nImporte
                        End If
                    End If
                End If
                
            Else
                If pn_Bandera = 1 Then
                    xlHoja1.Cells(pnFil, N) = Val(xlHoja1.Cells(pnFil, N)) - nImporte
                Else
                    If psMoneda = "1" Then
                        If psActPas = 1 Then
                            nActS(nActualAct, N - 1) = nActS(nActualAct, N - 1) - nImporte
                        Else
                            nPasS(nActualPas, N - 1) = nPasS(nActualPas, N - 1) - nImporte
                        End If
                    ElseIf psMoneda = "2" Then
                        If psActPas = 1 Then
                            nActD(nActualAct, N - 1) = nActD(nActualAct, N - 1) - nImporte
                        Else
                            nPasD(nActualPas, N - 1) = nPasD(nActualPas, N - 1) - nImporte
                        End If
                    End If
                End If
                
            End If
            ldFecha = ldFecha + 1
        Loop
    End If
Loop

End Sub

Private Sub CalculaArreglos(ByVal pdFecha As String)

Dim psArchivoAGrabar As String
 
Dim I As Integer
Dim j As Integer
Dim sCadTemp As String
Dim sCadSoles As String
Dim sCadDolares As String
Dim sCadPromedio As String
Dim nTempo1 As Currency
Dim nTempo2 As Currency

    'activo soles y dolares
    For j = 1 To nDias
        For I = 1 To nCantAct
            'Activo
            nTotS(1, j) = nTotS(1, j) + nActS(I, j)
            nTotD(1, j) = nTotD(1, j) + nActD(I, j)
        Next
    Next
    
    'pasivo soles y dolares
    For j = 1 To nDias
        For I = 1 To nCantPas
            'Activo
            nTotS(2, j) = nTotS(2, j) + nPasS(I, j)
            nTotD(2, j) = nTotD(2, j) + nPasD(I, j)
        Next
    Next
    
    'aca calculo ratios
    For I = 1 To nDias
        If nTotS(2, I) = 0 Then
            nTotS(3, I) = 0
        Else
            nTotS(3, I) = Format(nTotS(1, I) / nTotS(2, I) * 100, "0.00")
        End If
     
        If nTotD(2, I) = 0 Then
            nTotD(3, I) = 0
        Else
            nTotD(3, I) = Format(nTotD(1, I) / nTotD(2, I) * 100, "0.00")
        End If
    Next
    
    'Promedio de activos en soles y dolares
    For I = 1 To nCantAct
        For j = 1 To nDias
            nPromA(1, I) = nPromA(1, I) + nActS(I, j)
            nPromA(2, I) = nPromA(2, I) + nActD(I, j)
        Next
    Next
    
    'Aca Calculo promedio
    For I = 1 To nCantAct
        nPromA(1, I) = Format(nPromA(1, I) / nDias, "0.00")
        nPromA(2, I) = Format(nPromA(2, I) / nDias, "0.00")
    Next
    
    'Promedio de pasivos en soles y dolares
    For I = 1 To nCantPas
        For j = 1 To nDias
            nPromP(1, I) = nPromP(1, I) + nPasS(I, j)
            nPromP(2, I) = nPromP(2, I) + nPasD(I, j)
        Next
    Next
    
    'Aca Calculo promedio
    For I = 1 To nCantPas
        nPromP(1, I) = Format(nPromP(1, I) / nDias, "0.00")
        nPromP(2, I) = Format(nPromP(2, I) / nDias, "0.00")
    Next
    
    psArchivoAGrabar = App.path & "\SPOOLER\03" & Format(gdFecha, "YYMMdd") & ".115"

    Open psArchivoAGrabar For Output As #1
    Print #1, "01150300" & gsCodCMAC & Format(pdFecha, "YYYYMMDD") & "012" '0& "000000000000000"
    sCadTemp = ""
    For I = 1 To (nCantAct + nCantPas + 9)
        sCadTemp = sCadTemp & LlenaCerosSUCAVE(0)
    Next
    Print #1, " 100" & "  0" & sCadTemp
    
    'SOLES
    
    For I = 1 To nDias
        
        sCadTemp = ""
        sCadSoles = ""
           
        'Trabajo con Activo soles/dolares primero
        For j = 1 To nCantAct
            If j = 3 Then '3 No existe
                sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(0)
            End If
            sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(nActS(j, I))
        Next
        
        '7 Tampoco Existe
        sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(0)
            
        '8 es total de activos Soles/Dolares
        sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(nTotS(1, I))
        
        'Trabajo con Pasivo Soles/Dolares
        For j = 1 To nCantPas
            If j = 6 Then '6(13) No existe
                sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(0)
            End If
            sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(nPasS(j, I))
        Next
        
        '16 es total de pasivo soles/Dolares
        sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(nTotS(2, I))
        
        '17 es ratio Soles/Dolares
        sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(nTotS(3, I))
        
        '18 19 y 20  son ceros
        For j = 1 To 3
            sCadSoles = sCadSoles & "" & LlenaCerosSUCAVE(0)
        Next
               
        Print #1, " " & Trim(Str(100 + I)) & IIf(I < 10, "  ", " ") & Trim(Str(I)) & sCadSoles
    Next
      
    'Aca Calculo promedio soles
    For I = 1 To nCantAct
        sCadTemp = ""
        sCadPromedio = ""
           
        nTempo1 = 0
        nTempo2 = 0
           
        'Trabajo con Activo soles primero
        For j = 1 To nCantAct
            If j = 3 Then '3 No existe
                sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(0)
            Else
                nTempo1 = nTempo1 + nPromA(1, j)
            End If
            sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(nPromA(1, j))
        Next
           
        '7 Tampoco Existe
        sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(0)
            
        '8 es total de activos Soles
          
        sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(nTempo1)
         
        'Trabajo con Pasivo soles primero
        For j = 1 To nCantPas
            If j = 6 Then '6(13) No existe
                sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(0)
            End If
            nTempo2 = nTempo2 + nPromP(1, j)
            
            sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(nPromP(1, j))
        Next
          
        '16 es total de pasivo soles/PROMEDIO
        sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(nTempo2)
        
        '17 es ratio Soles/PROMEDIO
        sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(Format(nTempo1 / nTempo2 * 100, "0.00"))
        
        '18 19 y 20  son ceros
        For j = 1 To 3
            sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(0)
        Next
              
    Next
    
    Print #1, " 200" & "  0" & sCadPromedio
    
    sCadTemp = ""
    For I = 1 To (nCantAct + nCantPas + 9)
        sCadTemp = sCadTemp & LlenaCerosSUCAVE(0)
    Next
    Print #1, " 300" & "  0" & sCadTemp
        
    'DOLARES
    For I = 1 To nDias
        
        sCadTemp = ""
        sCadDolares = ""
           
        'Trabajo con Activo soles/dolares primero
        For j = 1 To nCantAct
            If j = 3 Then '3 No existe
                sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(0)
            End If
            sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(nActD(j, I))
        Next
        
        '7 Tampoco Existe
        sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(0)
            
        '8 es total de activos Soles/Dolares
        sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(nTotD(1, I))
        
        'Trabajo con Pasivo Soles/Dolares
        For j = 1 To nCantPas
            If j = 6 Then '6(13) No existe
                sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(0)
            End If
            sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(nPasD(j, I))
        Next
        
        '16 es total de pasivo soles/Dolares
        sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(nTotD(2, I))
        
        '17 es ratio Soles/Dolares
        sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(nTotD(3, I))
        
        '18 19 y 20  son ceros
        For j = 1 To 3
            sCadDolares = sCadDolares & "" & LlenaCerosSUCAVE(0)
        Next
             
         Print #1, " " & Trim(Str(300 + I)) & IIf(I < 10, "  ", " ") & Trim(Str(I)) & sCadDolares
    Next
    
    'Aca Calculo promedio dolares
    sCadPromedio = ""
    For I = 1 To nCantAct
        sCadTemp = ""
        sCadPromedio = ""
           
        nTempo1 = 0
        nTempo2 = 0
           
        'Trabajo con Activo soles primero
        For j = 1 To nCantAct
            If j = 3 Then '3 No existe
                sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(0)
            Else
                nTempo1 = nTempo1 + nPromA(2, j)
            End If
            sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(nPromA(2, j))
        Next
         
        
        '7 Tampoco Existe
        sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(0)
            
        '8 es total de activos Soles
          
        sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(nTempo1)
         
        'Trabajo con Pasivo soles primero
        For j = 1 To nCantPas
            If j = 6 Then '6(13) No existe
                sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(0)
            End If
            nTempo2 = nTempo2 + nPromP(2, j)
            
            sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(nPromP(2, j))
        Next
         
        
        '16 es total de pasivo soles/PROMEDIO
        sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(nTempo2)
        
        '17 es ratio Soles/PROMEDIO
        sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(Format(nTempo1 / nTempo2 * 100, "0.00"))
        
        '18 19 y 20  son ceros
        For j = 1 To 3
            sCadPromedio = sCadPromedio & "" & LlenaCerosSUCAVE(0)
        Next
        
    Next
    
    Print #1, " 400" & "  0" & sCadPromedio
    
    Close #1
End Sub

Public Sub Inicia(ByVal pBandera As Integer, nmodo As Integer, sforma As Form)
    bandera = pBandera
    Me.Show Val(nmodo), sforma
End Sub
