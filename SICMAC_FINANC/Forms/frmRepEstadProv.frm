VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmRepEstadProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadística de Gastos por Proveedor"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "frmRepEstadProv.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   2790
      TabIndex        =   7
      Top             =   2970
      Width           =   1290
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      Height          =   390
      Left            =   1080
      TabIndex        =   6
      Top             =   2970
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   2700
      Left            =   120
      TabIndex        =   8
      Top             =   90
      Width           =   5145
      Begin VB.TextBox txtNomProv 
         Enabled         =   0   'False
         Height          =   315
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "txtNombre"
         Top             =   2085
         Width           =   4725
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Consultar Un Solo Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   1410
         Width           =   2895
      End
      Begin Spinner.uSpinner uSpinner1 
         Height          =   315
         Left            =   3825
         TabIndex        =   1
         Top             =   435
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         Max             =   2025
         Min             =   100
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Arial"
         FontSize        =   9.75
      End
      Begin VB.TextBox TxtOrigenDesc 
         Enabled         =   0   'False
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   840
         Width           =   4770
      End
      Begin Sicmact.TxtBuscar txtOrigen 
         Height          =   345
         Left            =   1005
         TabIndex        =   0
         Top             =   420
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
      End
      Begin Sicmact.TxtBuscar txtCodProv 
         Height          =   345
         Left            =   810
         TabIndex        =   4
         Top             =   1695
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         TipoBusPers     =   2
      End
      Begin VB.Label Label2 
         Caption         =   "Código"
         Height          =   285
         Left            =   210
         TabIndex        =   11
         Top             =   1755
         Width           =   825
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Left            =   90
         Top             =   1305
         Width           =   4935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3105
         TabIndex        =   10
         Top             =   465
         Width           =   645
      End
      Begin VB.Label Label10 
         Caption         =   "Cuenta Contable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   195
         TabIndex        =   9
         Top             =   285
         Width           =   735
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmRepEstadProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub Genera(ByVal nBandera As Integer)
Dim sSql As String
Dim sCuenta As String
Dim sAnno As String
Dim reg As New ADODB.Recordset
Dim lsArchivo   As String
Dim lsRuta      As String
Dim lbLibroOpen As Boolean
Dim N           As Integer
Dim nFil As Integer
Dim nTot(1 To 13) As Currency
Dim ntempo As Currency
Dim oCon As DConecta
Dim I As Integer
Dim sRan As String
Dim sCad As String
Dim sProv As String
'Dim sDescripTempo As String

If Len(Trim(txtOrigen.Text)) = 0 Then
    MsgBox "Especifique una cuenta de gasto", vbExclamation, "Aviso!!!"
    txtOrigen.SetFocus
    Exit Sub
Else
    If uSpinner1.Valor = 0 Then
        MsgBox "Ingrese un año a buscar", vbExclamation, "Aviso!!!"
        uSpinner1.SetFocus
        Exit Sub
    End If
End If

sCuenta = Mid(txtOrigen.Text, 1, 2) & IIf(Mid(txtOrigen.Text, 3, 1) = "0", "[12]", Mid(txtOrigen.Text, 3, 1)) & Mid(txtOrigen.Text, 4, Len(txtOrigen.Text))
sAnno = uSpinner1.Valor
sProv = txtCodProv.psCodigoPersona

If nBandera = 1 Then
    sSql = " Select mg.cPersCod, ISNULL(p.cPersNombre, '- NO DEFINIDO') cPersNombre, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '01' THEN nMovImporte ELSE 0 END) Ene, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '02' THEN nMovImporte ELSE 0 END) Feb, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '03' THEN nMovImporte ELSE 0 END) Mar, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '04' THEN nMovImporte ELSE 0 END) Abr, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '05' THEN nMovImporte ELSE 0 END) May, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '06' THEN nMovImporte ELSE 0 END) Jun, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '07' THEN nMovImporte ELSE 0 END) Jul, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '08' THEN nMovImporte ELSE 0 END) Ago, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '09' THEN nMovImporte ELSE 0 END) Seti, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '10' THEN nMovImporte ELSE 0 END) Oct, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '11' THEN nMovImporte ELSE 0 END) Nov, " & _
           " SUM(CASE WHEN SubString(m.cMovNro,5,2) = '12' THEN nMovImporte ELSE 0 END) Dic " & _
           " from mov m join movcta mc on mc.nmovnro = m.nmovnro " & _
           " left join movgasto mg on mg.nmovnro = m.nmovnro " & _
           " left join Persona  p  on p.cPersCod = mg.cPersCod " & _
           " where m.cmovnro like '2002%' and m.nmovestado = 10 and m.nmovflag in (0,2,3) " & _
           " and mc.cctacontcod like '" & sCuenta & "%' and not m.copecod like '70185%' " & _
           " GROUP BY mg.cPersCod, ISNULL(p.cPersNombre, '- NO DEFINIDO') " & _
           " order by cPersNombre"
           
           
           
           '" and mc.cctacontcod like '45[12]3010501%' and not m.copecod like '70185%' "
           
ElseIf nBandera = 2 Then
    sSql = " Select left(m.cmovnro,8) fecha, md.ndoctpo, doc.cdocabrev, md.cdocnro, mc.nmovimporte, m.cmovdesc, m.cmovnro " & _
           " from mov m join movcta mc on mc.nmovnro = m.nmovnro " & _
           " left join movdoc md on md.nmovnro = m.nmovnro " & _
           " left join movgasto mg on mg.nmovnro = m.nmovnro " & _
           " left join Persona  p  on p.cPersCod = mg.cPersCod  " & _
           " inner join documento Doc on md.ndoctpo=doc.ndoctpo " & _
           " where m.cmovnro like '2002%' and m.nmovestado = 10 and m.nmovflag in (0,2,3) " & _
           " and mg.cperscod ='" & sProv & "' " & _
           " and mc.cctacontcod like '" & sCuenta & "%' and not m.copecod like '70185%' " & _
           " Order by m.cmovnro"
           
           'and mg.cperscod = '1120700177865'
           'and mc.cctacontcod like '45[12]3010501%' and not m.copecod like '70185%' "
End If

Set oCon = New DConecta

oCon.AbreConexion
Set reg = oCon.CargaRecordSet(sSql)
If reg.BOF Then
    Set reg = Nothing
    oCon.CierraConexion
    MsgBox "No existen datos para generar el reporte", vbExclamation, "Aviso!!!"
    Exit Sub
End If

    lsRuta = App.path & "\Spooler\"
    
    If nBandera = 1 Then
        lsArchivo = lsRuta & "RepEstadListaProv" & "_" & uSpinner1.Valor & ".xls"
    ElseIf nBandera = 2 Then
        lsArchivo = lsRuta & "RepEstadConsProv" & "_" & uSpinner1.Valor & ".xls"
    End If
    
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    
    If lbLibroOpen Then
        
        Set xlHoja1 = xlLibro.Worksheets(1)
        
        ExcelAddHoja uSpinner1.Valor, xlLibro, xlHoja1
     
        CabeceraExcel nBandera
        
        If nBandera = 1 Then
            
            nFil = 7
      
            Do While Not reg.EOF
            
                nFil = nFil + 1
                
                xlHoja1.Cells(nFil, 1) = reg!cPersCod
                xlHoja1.Cells(nFil, 2) = reg!cPersNombre
                
                xlHoja1.Cells(nFil, 3) = reg!Ene
                xlHoja1.Cells(nFil, 4) = reg!Feb
                xlHoja1.Cells(nFil, 5) = reg!Mar
                xlHoja1.Cells(nFil, 6) = reg!Abr
                xlHoja1.Cells(nFil, 7) = reg!May
                xlHoja1.Cells(nFil, 8) = reg!Jun
                xlHoja1.Cells(nFil, 9) = reg!Jul
                xlHoja1.Cells(nFil, 10) = reg!Ago
                xlHoja1.Cells(nFil, 11) = reg!Seti
                xlHoja1.Cells(nFil, 12) = reg!Oct
                xlHoja1.Cells(nFil, 13) = reg!Nov
                xlHoja1.Cells(nFil, 14) = reg!Dic
                 
                ntempo = 0
                ntempo = reg!Ene + reg!Feb + reg!Mar + reg!Abr + reg!May + reg!Jun + reg!Jul + reg!Ago + reg!Seti + reg!Oct + reg!Nov + reg!Dic
                
                xlHoja1.Cells(nFil, 15) = ntempo
                 
                nTot(1) = nTot(1) + reg!Ene
                nTot(2) = nTot(2) + reg!Feb
                nTot(3) = nTot(3) + reg!Mar
                nTot(4) = nTot(4) + reg!Abr
                nTot(5) = nTot(5) + reg!May
                nTot(6) = nTot(6) + reg!Jun
                nTot(7) = nTot(7) + reg!Jul
                nTot(8) = nTot(8) + reg!Ago
                nTot(9) = nTot(9) + reg!Seti
                nTot(10) = nTot(10) + reg!Oct
                nTot(11) = nTot(11) + reg!Nov
                nTot(12) = nTot(12) + reg!Dic
                nTot(13) = nTot(13) + ntempo
                
                reg.MoveNext
            Loop
      
            nFil = nFil + 1
            
            xlHoja1.Cells(nFil, 1) = "Totales"
            xlHoja1.Cells(nFil + 1, 1) = "Totales Acumulados"
            
            ntempo = 0
            
            For I = 3 To 15
                  
                  xlHoja1.Cells(nFil, I) = nTot(I - 2)
                  
                  If I < 15 Then
                      ntempo = ntempo + nTot(I - 2)
                      xlHoja1.Cells(nFil + 1, I) = ntempo
                  End If
            Next
                   
            For I = 1 To 3
            
                If I = 1 Then
                    sRan = "A7:O" & Trim(Str(nFil + 1))
                ElseIf I = 2 Then
                    sRan = "A8:O" & Trim(Str(nFil - 1))
                ElseIf I = 3 Then
                    sRan = "C7:N" & Trim(Str(nFil + 1))
                End If
                
                Range(sRan).Select
                Range(sRan).Activate
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Next
      
            xlHoja1.Range("A" & Trim(Str(nFil)) & ":B" & Trim(Str(nFil))).MergeCells = True
            xlHoja1.Range("A" & Trim(Str(nFil + 1)) & ":B" & Trim(Str(nFil + 1))).MergeCells = True
            
            xlHoja1.Range("A" & Trim(Str(nFil)) & ":B" & Trim(Str(nFil))).Font.Bold = True
            xlHoja1.Range("A" & Trim(Str(nFil + 1)) & ":B" & Trim(Str(nFil + 1))).Font.Bold = True
            
            
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
            xlHoja1.Cells.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            
            xlHoja1.Range("A1:O7").Font.Bold = True
            
            Columns("B:B").ColumnWidth = 40.29
        ElseIf nBandera = 2 Then
            
            nFil = 8
      
            Do While Not reg.EOF
            
                nFil = nFil + 1
                
                xlHoja1.Cells(nFil, 1) = Mid(reg!Fecha, 7, 2) & "/" & Mid(reg!Fecha, 5, 2) & "/" & Mid(reg!Fecha, 1, 4)
                xlHoja1.Cells(nFil, 2) = reg!cDocAbrev
                xlHoja1.Cells(nFil, 3) = reg!cDocNro
                 
                xlHoja1.Cells(nFil, 4) = Format(reg!nMovImporte, "###,###,##0.00")
 
                xlHoja1.Cells(nFil, 5) = Mid(Replace(Replace(reg!cMovDesc, Chr(10), ""), Chr(13), ""), 1, 70)
                 
                reg.MoveNext
            Loop
             
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("A8:E" & Trim(Str(nFil + 1))).Select
            
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
             
            xlHoja1.Cells(nFil + 1, 1) = "Total"
            xlHoja1.Cells(nFil + 1, 3) = "Docs: " & Trim(Str(nFil - 8))
            
            With xlHoja1.Range("A" & Trim(Str(nFil + 1)) & ":B" & Trim(Str(nFil + 1)))
                .MergeCells = True
                .HorizontalAlignment = xlLeft
            End With
            
            xlHoja1.Range("A" & Trim(Str(nFil + 1)) & ":E" & Trim(Str(nFil + 1))).Font.Bold = True
            
            xlHoja1.Range("D" & Trim(Str(nFil + 1)) & ":D" & Trim(Str(nFil + 1))).Formula = "=SUM(D9..D" & Trim(Str(nFil)) & ")"
            
            Columns("D:D").Select
            Selection.NumberFormat = "0.00"
            
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
            
            xlHoja1.Range("A1:E8").Font.Bold = True
        End If
        
        Set reg = Nothing
        oCon.CierraConexion
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        
        CargaArchivo lsArchivo, lsRuta
        
        MsgBox "Reporte Generado satisfactoriamente", vbInformation, "Aviso!!!"
    
    End If
    
End Sub

Private Sub CabeceraExcel(ByVal nBan As Integer)
    
    xlHoja1.Cells(1, 1) = gsNomCmac
    xlHoja1.Cells(3, 1) = "Estadística de Gastos por Proveedor - Periodo " & uSpinner1.Valor
    xlHoja1.Cells(5, 1) = "Cuenta Contable: [" & txtOrigen.Text & "] " & TxtOrigenDesc
    
    If nBan = 1 Then
        xlHoja1.Cells(7, 1) = "Código"
        xlHoja1.Cells(7, 2) = "Proveedor"
        xlHoja1.Cells(7, 3) = "Enero"
        xlHoja1.Cells(7, 4) = "Febrero"
        xlHoja1.Cells(7, 5) = "Marzo"
        xlHoja1.Cells(7, 6) = "Abril"
        xlHoja1.Cells(7, 7) = "Mayo"
        xlHoja1.Cells(7, 8) = "Junio"
        xlHoja1.Cells(7, 9) = "Julio"
        xlHoja1.Cells(7, 10) = "Agosto"
        xlHoja1.Cells(7, 11) = "Setiembre"
        xlHoja1.Cells(7, 12) = "Octubre"
        xlHoja1.Cells(7, 13) = "Noviembre"
        xlHoja1.Cells(7, 14) = "Diciembre"
        xlHoja1.Cells(7, 15) = "Total"
    
        xlHoja1.Range("A3:O3").MergeCells = True
        xlHoja1.Range("A1:B1").MergeCells = True
        xlHoja1.Range("A5:B5").MergeCells = True
    ElseIf nBan = 2 Then
        
        xlHoja1.Cells(6, 1) = "Proveedor: [" & txtCodProv.Text & "] " & txtNomProv.Text
        
        xlHoja1.Cells(8, 1) = "Fecha"
        xlHoja1.Cells(8, 2) = "Doc."
        xlHoja1.Cells(8, 3) = "Número"
        xlHoja1.Cells(8, 4) = "Importe"
        xlHoja1.Cells(8, 5) = "Descripción"

        xlHoja1.Range("A3:E3").MergeCells = True
        xlHoja1.Range("A6:E6").MergeCells = True
        xlHoja1.Range("A1:E1").MergeCells = True
        xlHoja1.Range("A5:E5").MergeCells = True
    End If
    

    
    xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
    
End Sub
  
Private Sub Check1_Click()
    txtCodProv.Enabled = IIf(Check1.value = 1, True, False)
    txtCodProv.Text = ""
    txtNomProv.Text = ""
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCodProv.EnabledText = True Then
            txtCodProv.SetFocus
        Else
            cmdExcel.SetFocus
        End If
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdExcel_Click()
    Genera IIf(Check1.value = 1, 2, 1)
End Sub


Private Sub Form_Load()
    
    Dim clsCta As New DCtaCont
    Dim rsCta As New ADODB.Recordset

    CentraForm Me
 
    Set rsCta = clsCta.CargaCtaCont("cCtaContCod LIKE '4%'", gsCentralCom & "CtaCont", adLockReadOnly)
    Set clsCta = Nothing
    
    uSpinner1.Valor = Year(Date) - 1
    
    txtOrigen.rs = rsCta
    txtOrigen.TipoBusqueda = BuscaGrid
    txtOrigen.lbUltimaInstancia = False
  
    txtCodProv.TipoBusqueda = BuscaPersona
    txtCodProv.TipoBusPers = BusPersDocumentoRuc
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRepEstadProv = Nothing
End Sub

Private Sub txtCodProv_EmiteDatos()
txtNomProv = txtCodProv.psDescripcion
txtCodProv.Tag = txtCodProv.psCodigoPersona
 
If txtNomProv <> "" Then
    cmdExcel.SetFocus
End If
End Sub
 
Private Sub txtOrigen_EmiteDatos()
    If txtOrigen.psDescripcion <> "" Then
      TxtOrigenDesc.Text = txtOrigen.psDescripcion
          uSpinner1.SetFocus
   End If
End Sub

Private Sub uSpinner1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check1.SetFocus
End If
End Sub
