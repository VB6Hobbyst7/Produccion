VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAjusteDeprecia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Ajustes por Depreciacion Historica"
   ClientHeight    =   5205
   ClientLeft      =   3435
   ClientTop       =   2940
   ClientWidth     =   6285
   Icon            =   "frmAjusteDeprecia.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   90
      TabIndex        =   7
      Top             =   1530
      Width           =   6120
      Begin MSComctlLib.ListView LstRep 
         Height          =   2640
         Left            =   120
         TabIndex        =   3
         Top             =   225
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Archivo"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CuentaD"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CuentaH"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   2250
      TabIndex        =   5
      Top             =   4710
      Width           =   1440
   End
   Begin VB.Frame Frame4 
      Height          =   705
      Left            =   90
      TabIndex        =   10
      Top             =   750
      Width           =   6120
      Begin VB.ComboBox cboDec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "frmAjusteDeprecia.frx":030A
         Left            =   1710
         List            =   "frmAjusteDeprecia.frx":0320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2985
      End
      Begin VB.Label Label3 
         Caption         =   "Decimales"
         Height          =   285
         Left            =   210
         TabIndex        =   11
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.CommandButton CmdAsiento 
      Caption         =   "&Asiento"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   4710
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3780
      TabIndex        =   6
      Top             =   4710
      Width           =   1365
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   90
      TabIndex        =   8
      Top             =   15
      Width           =   6120
      Begin VB.TextBox TxtAño 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   4
         TabIndex        =   1
         Top             =   225
         Width           =   915
      End
      Begin VB.ComboBox CboMes 
         Height          =   315
         ItemData        =   "frmAjusteDeprecia.frx":0336
         Left            =   1725
         List            =   "frmAjusteDeprecia.frx":035E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Proceso:"
         Height          =   225
         Left            =   225
         TabIndex        =   9
         Top             =   255
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAjusteDeprecia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dFecha As Date
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nIni As Integer
Dim nCol As Integer
Dim LibroAb As Boolean
Dim nDepHis As Double
Dim nDepAj As Double

Dim oDep As DAjusteDeprecia

Private Function ValidaProcesar() As Boolean
   ValidaProcesar = False
    If Len(Trim(TxtAño.Text)) = 0 Then
        MsgBox "Falta Ingresar el Año", vbInformation, "Aviso"
        Exit Function
    End If
    If Len(Trim(TxtAño.Text)) < 4 Then
        MsgBox "Faltan Digitos a Año Ingresado ", vbInformation, "Aviso"
        Exit Function
    End If
    ValidaProcesar = True
End Function
Private Sub CargaDatos()
Dim sSql As String
Dim R As New ADODB.Recordset
Dim l As ListItem
    LstRep.ListItems.Clear
    Set R = oDep.CargaAjusteDeprecia()
        Do While Not R.EOF
            Set l = LstRep.ListItems.Add(, , Trim(R!cDescrip))
            l.SubItems(1) = Trim(Str(R!nCodigo))
            l.SubItems(2) = Trim(R!cNomArch)
            l.SubItems(3) = Trim(R!cCtaContCodD)
            l.SubItems(4) = Trim(R!cCtaContCodH)
            R.MoveNext
        Loop
    R.Close
End Sub
Private Sub Cabecera(ByVal sTit As String, sFecValHis As String, sFecValAj As String)

    'xlAplicacion.Range("A1:R100").Font.Size = 10
    xlHoja1.Range("A1:H1").Font.Name = "ARIAL"
    xlHoja1.Range("A1:H1").Font.Size = 14
    xlHoja1.Range("A1:H1").MergeCells = True
    xlHoja1.Range("A1:H1").HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A1:H1").VerticalAlignment = xlHAlignCenter
    xlHoja1.Range("A1:H1") = UCase(gsNomCmac)
    
    xlHoja1.Range("A3:H3").Font.Name = "ARIAL"
    xlHoja1.Range("A3:H3").Font.Size = 16
    xlHoja1.Range("A3:H3").MergeCells = True
    xlHoja1.Range("A3:H3").HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A3:H3").VerticalAlignment = xlHAlignCenter
    xlHoja1.Range("A3:H3") = sTit
    
    xlHoja1.Range("A5:H5").Font.Name = "ARIAL"
    xlHoja1.Range("A5:H5").Font.Size = 14
    xlHoja1.Range("A5:H5").MergeCells = True
    xlHoja1.Range("A5:H5").HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A5:H5").VerticalAlignment = xlHAlignCenter
    xlHoja1.Range("A5:H5") = Format(dFecha, "mmmm-yyyy")
        
    xlHoja1.Range("A7:K7").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
'    xlHoja1.Range("A7:K7").FormatCondition = xlHoja1.Range("A7:K7").FormatConditions(1)
    
    xlHoja1.Cells(nIni, nCol + 1) = "ITEM"
    xlHoja1.Cells(nIni, nCol + 2) = "CODIGO"
    xlHoja1.Cells(nIni, nCol + 3) = "DESCRIPCION"
    xlHoja1.Cells(nIni, nCol + 4) = "UBICACION"
    xlHoja1.Cells(nIni, nCol + 5) = "VALOR HISTORICO AL " & sFecValHis
    xlHoja1.Cells(nIni, nCol + 6) = "FECHA DE ADQUISICION"
    xlHoja1.Cells(nIni, nCol + 7) = "FACTOR DE AJUSTE"
    xlHoja1.Cells(nIni, nCol + 8) = "VALOR AJUSTADO AL" & sFecValAj
    xlHoja1.Cells(nIni, nCol + 9) = "VIDA UTIL DEL ACTIVO"
    xlHoja1.Cells(nIni, nCol + 10) = "MESES DE DEPREC."
    xlHoja1.Cells(nIni, nCol + 11) = "DEPREC. HISTORICA"
    xlHoja1.Cells(nIni, nCol + 12) = "DEPREC. AJUSTADA"
    
End Sub

Private Function ValorHistorico(ByVal FecAdq As Date, FecRep As Date, ByVal nValHis As Double)
Dim DifAgnos As Integer
Dim i As Integer
Dim FecTemp As Date
Dim FecAdqTemp As Date
Dim ValHisTemp As Double
Dim bOk As Boolean
On Error GoTo ValorErr
    ValHisTemp = nValHis
    If Year(FecAdq) < Year(FecRep) Then
        DifAgnos = Year(FecRep) - Year(FecAdq)
        FecTemp = CDate("31/12/" & Trim(Str(Year(FecAdq))))
        ValHisTemp = FactorAjuste(FecAdq, FecTemp) * ValHisTemp
        For i = 1 To DifAgnos - 1
            FecAdqTemp = CDate("01/01/" & Trim(Str(Year(FecAdq) + i)))
            FecTemp = CDate("31/12/" & Trim(Str(Year(FecAdq) + i)))
            ValHisTemp = FactorAjuste(FecAdqTemp, FecTemp) * ValHisTemp
        Next i
    Else
        ValHisTemp = nValHis
    End If
    ValorHistorico = ValHisTemp
   Exit Function
ValorErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function
Private Sub procesar()
Dim sSql As String
Dim R As New ADODB.Recordset
Dim i As Integer
Dim lsArchivo As String
Dim dFecValHis As Date
Dim sCad As String
Dim fs As New Scripting.FileSystemObject
Dim j As Integer
Dim bOk As Boolean
Dim nFacI As Double
Dim nValHis As Double
Dim dFechaTemp As Date
Dim bExiste As Boolean


   Screen.MousePointer = 11
   dFecha = DateAdd("m", 1, CDate("01/" & Format(CboMes.ListIndex + 1, "00") & "/" & TxtAño.Text)) - 1
   
   'Graba y Cierra Aplicacion de Exel
    lsArchivo = App.Path & "\Spooler\" & Trim(LstRep.SelectedItem.SubItems(2)) & TxtAño & ".XLS"
    
   'Propiedades de Exel
   If Not ExcelBegin(lsArchivo, xlAplicacion, xlLibro) Then
      Exit Sub
   End If
   
    Set xlHoja1 = xlLibro.Worksheets(1)
    xlAplicacion.Range("A1:R100").Font.Size = 10
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 55
    ExcelAddHoja CboMes.Text, xlLibro, xlHoja1
            
            i = 1
            nDepHis = 0
            nDepAj = 0
            Set R = oDep.CargaActivosDeprecia(LstRep.SelectedItem.SubItems(1), Format(dFecha, gsFormatoFecha))
                If Not R.BOF And Not R.EOF Then
                    bExiste = True
                    Call Cabecera(LstRep.SelectedItem.Text, Format(R!dFecAdq, "dd/mm/yyyy"), Format(R!dFecAdq, "dd/mm/yyyy"))
                Else
                    bExiste = False
                End If
                Do While Not R.EOF
                    xlHoja1.Cells(nIni + i, nCol + 1) = i
                    xlHoja1.Cells(nIni + i, nCol + 2) = R!cBSCod
                    xlHoja1.Cells(nIni + i, nCol + 3) = R!cDescrip
                    xlHoja1.Cells(nIni + i, nCol + 4) = R!cUbicacion
                    nValHis = ValorHistorico(R!dFecAdq, dFecha, R!nValHis)
                    xlHoja1.Cells(nIni + i, nCol + 5) = nValHis
                    xlHoja1.Cells(nIni + i, nCol + 6) = Format(R!dFecAdq, "mm/dd/yyyy")
                    If Year(R!dFecAdq) < Year(dFecha) Then
                        dFechaTemp = CDate("31/12/" & Trim(Str(Year(dFecha) - 1)))
                        nFacI = FactorAjuste(dFechaTemp, dFecha)
                    Else
                        nFacI = FactorAjuste(R!dFecAdq, dFecha)
                    End If
                    If nFacI <> 0 Then
                        If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
                            xlLibro.Close
                        Else
                            xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
                            xlLibro.Close
                        End If
                        'Cierra el libro de trabajo
                        xlAplicacion.Quit
                        'Libera los objetos.
                        Set xlAplicacion = Nothing
                        Set xlLibro = Nothing
                        Set xlHoja1 = Nothing
                        R.Close
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                    xlHoja1.Cells(nIni + i, nCol + 7) = nFacI
                    xlHoja1.Cells(nIni + i, nCol + 8) = CDbl(Format(nValHis * nFacI, "#0.00"))
                    xlHoja1.Cells(nIni + i, nCol + 9) = R!nVidaUtil
                    xlHoja1.Cells(nIni + i, nCol + 10) = CInt((dFecha - R!dFecAdq) / 30)
                    xlHoja1.Cells(nIni + i, nCol + 11) = CDbl(Format((nValHis / R!nVidaUtil) * CInt((dFecha - R!dFecAdq) / 30), "#0.00"))
                    nDepHis = nDepHis + CDbl(Format((nValHis / R!nVidaUtil) * CInt((dFecha - R!dFecAdq) / 30), "#0.00"))
                    xlHoja1.Cells(nIni + i, nCol + 12) = CDbl(Format((CDbl(Format(nValHis * nFacI, "#0.00")) / R!nVidaUtil) * CInt((dFecha - R!dFecAdq) / 30), "#0.00"))
                    nDepAj = nDepAj + CDbl(Format((CDbl(Format(nValHis * nFacI, "#0.00")) / R!nVidaUtil) * CInt((dFecha - R!dFecAdq) / 30), "#0.00"))
                    i = i + 1
                    R.MoveNext
                Loop
            R.Close
            If bExiste Then
                sCad = Chr(65 + (nCol + 1)) & Trim(Str(nIni + i)) & ":" & Chr(65 + (nCol + 2)) & Trim(Str(nIni + i))
                xlHoja1.Range(sCad).MergeCells = True
                xlHoja1.Cells(nIni + i, nCol + 2) = "T O T A L"
                
                sCad = "=Sum(" & Chr(65 + (nCol + 4)) & Trim(Str(nIni + 1)) & ":" & Chr(65 + (nCol + 4)) & Trim(Str(nIni + (i - 1))) & ")"
                xlHoja1.Range(xlHoja1.Cells(nIni + i, nCol + 5), xlHoja1.Cells(nIni + i, nCol + 5)).Formula = sCad
                                
                sCad = "=Sum(" & Chr(65 + (nCol + 7)) & Trim(Str(nIni + 1)) & ":" & Chr(65 + (nCol + 7)) & Trim(Str(nIni + (i - 1))) & ")"
                xlHoja1.Range(xlHoja1.Cells(nIni + i, nCol + 8), xlHoja1.Cells(nIni + i, nCol + 8)).Formula = sCad
                
                sCad = "=Sum(" & Chr(65 + (nCol + 10)) & Trim(Str(nIni + 1)) & ":" & Chr(65 + (nCol + 10)) & Trim(Str(nIni + (i - 1))) & ")"
                xlHoja1.Range(xlHoja1.Cells(nIni + i, nCol + 11), xlHoja1.Cells(nIni + i, nCol + 11)).Formula = sCad
                
                sCad = "=Sum(" & Chr(65 + (nCol + 11)) & Trim(Str(nIni + 1)) & ":" & Chr(65 + (nCol + 11)) & Trim(Str(nIni + (i - 1))) & ")"
                xlHoja1.Range(xlHoja1.Cells(nIni + i, nCol + 12), xlHoja1.Cells(nIni + i, nCol + 12)).Formula = sCad
        End If
   
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    CargaArchivo lsArchivo, App.Path & "\SPOOLER"
    LibroAb = True
    
    Screen.MousePointer = 0
End Sub
Private Sub GeneraAsiento()
Dim sSql As String
Dim nMov As Long
Dim oMov As New DMov
    oMov.BeginTrans
    
    'Genera asiento de Dep. historica
    gsMovNro = oMov.GeneraMovNro(dFecha, gsCodAge, gsCodUser)
    gsGlosa = "Asiento de Depreciación Histórica: " & LstRep.SelectedItem.Text & " - " & CboMes.Text & " " & TxtAño.Text
    
    oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable
    nMov = oMov.GetnMovNro(gsMovNro)
    
    oMov.InsertaMovCont nMov, nDepHis, 0, ""
    oMov.InsertaMovCta nMov, 1, LstRep.SelectedItem.SubItems(3), Format(nDepHis, "#0.00")
    oMov.InsertaMovCta nMov, 2, LstRep.SelectedItem.SubItems(4), Format(-1 * nDepHis, "#0.00")
    
    'Genera asiento de Deprecia. Ajustada
    gsMovNro = oMov.GeneraMovNro(dFecha, gsCodAge, gsCodUser)
    gsGlosa = "Asiento de Depreciación Ajustada: " & LstRep.SelectedItem.Text & " - " & CboMes.Text & " " & TxtAño.Text
    
    oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable
    nMov = oMov.GetnMovNro(gsMovNro)
    
    oMov.InsertaMovCont nMov, nDepHis, 0, ""
    oMov.InsertaMovCta nMov, 1, Left(LstRep.SelectedItem.SubItems(3), 2) & "6" & Mid(LstRep.SelectedItem.SubItems(3), 4, 20), Format(nDepAj, "#0.00")
    oMov.InsertaMovCta nMov, 2, Left(LstRep.SelectedItem.SubItems(4), 2) & "6" & Mid(LstRep.SelectedItem.SubItems(3), 4, 20), Format(-1 * nDepAj, "#0.00")
       
    oMov.CommitTrans
    MsgBox "Asiento se Genero con Exito", vbInformation, "Aviso"
    
    Dim oFun As New NContImprimir
    EnviaPrevio oFun.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "Asiento de Depreciación"), "Asiento de Depreciación", gnLinPage, False
    Set oFun = Nothing
End Sub
Private Sub CboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtAño.SetFocus
    End If
End Sub

Private Sub cmdAsiento_Click()
    If MsgBox("¿Desea Generar Asiento de Depreciación?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Call GeneraAsiento
    End If
End Sub

Private Sub cmdProcesar_Click()
    If ValidaProcesar Then
       Call procesar
       cmdAsiento.Enabled = True
       cmdAsiento.SetFocus
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oDep = New DAjusteDeprecia
    CentraForm Me
    nIni = 7
    nCol = 0
    Call CargaDatos
    TxtAño.Text = Trim(Str(Year(gdFecSis)))
    CboMes.ListIndex = Month(gdFecSis) - 1
    cboDec.ListIndex = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oDep = Nothing
End Sub

Private Sub LstRep_DblClick()
    LibroAb = True
End Sub

Private Sub TxtAño_GotFocus()
    fEnfoque TxtAño
End Sub

Private Sub TxtAño_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        CmdProcesar.SetFocus
    End If
End Sub
