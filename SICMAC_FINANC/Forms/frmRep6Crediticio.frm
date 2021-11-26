VERSION 5.00
Begin VB.Form frmRep6Crediticio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte 06: Reporte Diario de Tasas de Interés"
   ClientHeight    =   690
   ClientLeft      =   1350
   ClientTop       =   2340
   ClientWidth     =   5565
   Icon            =   "frmRep6Crediticio.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmRep6Crediticio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim ldFecha  As Date
Dim oBarra As clsProgressBar
Dim oCon As DConecta
 
Dim sservidorconsolidada As String
     
Public Sub ImprimeAnexo6Crediticio(psOpeCod As String, pdFecha As Date, ByVal sLstage As ListBox)
On Error GoTo GeneraEstadError
Dim oConecta As DConecta
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim cConsol As String

    Dim i As Integer
    Dim lsAgencia As String
    
    lsAgencia = ""
    cConsol = "N"
    For i = 0 To sLstage.ListCount - 1
        If sLstage.Selected(i) Then
            If Right(sLstage.List(i), 5) = "ONSOL" Then
                lsAgencia = lsAgencia & Right(sLstage.List(i), 2) & ","
                cConsol = "S"
            Else
                lsAgencia = lsAgencia & Right(sLstage.List(i), 2) & ","
            End If
        End If
    Next i
    
    lsAgencia = Mid(lsAgencia, 1, Len(lsAgencia) - 1)


   Set oConecta = New DConecta
   oConecta.AbreConexion
   Set rs = oConecta.CargaRecordSet("select nconssisvalor from constsistema where nconssiscod=" & gConstSistServCentralRiesgos)
    If rs.BOF Then
    Else
        sservidorconsolidada = rs!nConsSisValor
    End If
    Set rs = Nothing
   
   oConecta.CierraConexion
   Set oConecta = Nothing
   
   Set oCon = New DConecta
   oCon.AbreConexion 'Remota gsCodAge, True, False, "03"
   
   ldFecha = pdFecha
   lsArchivo = App.path & "\SPOOLER\" & "SBSRep6_" & Format(pdFecha, "mmyyyy") & ".XLS"
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If lbExcel Then
   
      ExcelAddHoja "Rep_6A", xlLibro, xlHoja1
      Genera6A 1, pdFecha, lsAgencia, cConsol 'Coloc Ok
      Genera6A 2, pdFecha, lsAgencia, cConsol 'Coloc OK
'''      ExcelAddHoja "Rep_6A_DM1", xlLibro, xlHoja1
'''      Genera6A_D 1, pdFecha 'Coloc Ok
'''      ExcelAddHoja "Rep_6A_DM2", xlLibro, xlHoja1
'''      Genera6A_D 2, pdFecha 'Coloc Ok

      ExcelAddHoja "Rep_6B", xlLibro, xlHoja1
        Genera6B 1, pdFecha, lsAgencia, cConsol 'Capta OK
        Genera6B 2, pdFecha, lsAgencia, cConsol 'Capta OK
'''      ExcelAddHoja "Rep_6B_DM1", xlLibro, xlHoja1
'''        Genera6B_D 1, pdFecha 'Capta OK
'''      ExcelAddHoja "Rep_6B_DM2", xlLibro, xlHoja1
'''        Genera6B_D 2, pdFecha 'Capta OK

''
''      ExcelAddHoja "Rep_6D", xlLibro, xlHoja1
''      Genera6D 1, pdFecha 'Coloc Ok
''      Genera6D 2, pdFecha 'Coloc Ok
''      ExcelAddHoja "Rep_6D_DM1", xlLibro, xlHoja1
''      Genera6D_D 1, pdFecha 'Coloc Ok
''      ExcelAddHoja "Rep_6D_DM2", xlLibro, xlHoja1
''      Genera6D_D 2, pdFecha 'Coloc Ok

''      ExcelAddHoja "Rep_6E", xlLibro, xlHoja1
''      Genera6E 1, pdFecha 'Capta
''      Genera6E 2, pdFecha 'Capta
''      ExcelAddHoja "Rep_6E_DM1", xlLibro, xlHoja1
''      Genera6E_D 1, pdFecha 'Capta
''      ExcelAddHoja "Rep_6E_DM2", xlLibro, xlHoja1
''      Genera6E_D 2, pdFecha 'Capta

''      'peac 20071127

      ExcelAddHoja "Rep_6D (nuevo)", xlLibro, xlHoja1
      Genera6D1 1, pdFecha, lsAgencia, cConsol 'Coloc oK
      Genera6D1 2, pdFecha, lsAgencia, cConsol 'Coloc Ok
'''       ExcelAddHoja "Rep_6D1_DM1 (nuevo)", xlLibro, xlHoja1
'''      Genera6D1_D 1, pdFecha 'Coloc oK
'''      ExcelAddHoja "Rep_6D1_DM2 (nuevo)", xlLibro, xlHoja1
'''      Genera6D1_D 2, pdFecha 'Coloc Ok
''
      'peac 20071127
      ExcelAddHoja "Rep_6E (nuevo)", xlLibro, xlHoja1
      Genera6E1 1, pdFecha, lsAgencia, cConsol 'Capta
      Genera6E1 2, pdFecha, lsAgencia, cConsol 'Capta
      
'''      ExcelAddHoja "Rep_6E1_DM1 (nuevo)", xlLibro, xlHoja1
'''      Genera6E1_D 1, pdFecha 'Capta
'''      ExcelAddHoja "Rep_6E1_DM2 (nuevo)", xlLibro, xlHoja1
'''      Genera6E1_D 2, pdFecha 'Capta
      
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
      If lsArchivo <> "" Then
         CargaArchivo lsArchivo, App.path & "\SPOOLER\"
      End If
   End If
   oCon.CierraConexion
   Set oCon = Nothing
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub
Public Sub GeneraAnexo6CrediticioSUCAVE(psOpeCod As String, pdFecha As Date)

On Error GoTo GeneraEstadError

Dim psArchivoA_Leer As String
Dim psArchivoAGrabar As String
Dim Hojas(1 To 4) As String
Dim psAnno As String
Dim psMes As String

ldFecha = pdFecha
psArchivoA_Leer = App.path & "\SPOOLER\" & "SBSRep6_" & Format(pdFecha, "mmyyyy") & ".XLS"
 
Hojas(1) = "Rep_6A"
Hojas(2) = "Rep_6B"
Hojas(3) = "Rep_6D"
Hojas(4) = "Rep_6E"
  
psMes = Mid(pdFecha, 4, 2)
psAnno = Mid(pdFecha, 7, 4)

Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject
Dim sCad As String
Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim nFil As Integer
Dim nCol As Integer
 
Dim bRep1 As Boolean
Dim bRep2 As Boolean
Dim bRep3 As Boolean
Dim bRep4 As Boolean
Dim bRep5 As Boolean

Dim matImprimir(43, 7) As Double
 
bExiste = fs.FileExists(psArchivoA_Leer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoA_Leer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
 
    Set xlAplicacion = New Excel.Application

    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoA_Leer)
    
Genera01:
     
    bRep1 = False
     
    'Primer Archivo a Grabar
    psArchivoAGrabar = App.path & "\SPOOLER\01" & Format(pdFecha, "YYMMdd") & ".206"
     
    bEncontrado = False
    For Each xlHoja1 In xlLibro.Worksheets
        If Trim(UCase(xlHoja1.Name)) = Trim(UCase(Hojas(1))) Then
            bEncontrado = True
            xlHoja1.Activate
            Exit For
        End If
    Next

    If bEncontrado = False Then
        'ExcelEnd psArchivoAGrabar, xlAplicacion, xlLibro, xlHoja1, True
        'MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        GoTo Genera02
    End If
  
  
    For i = 1 To 28
        If i >= 1 And i <= 9 Then
            nFil = i + 11
        ElseIf i >= 10 And i <= 14 Then
            nFil = i + 12
        ElseIf i >= 15 And i <= 28 Then
            nFil = i + 13
        End If
     
        For j = 2 To 5
            matImprimir(i, j) = xlHoja1.Cells(nFil, j + 1)
            matImprimir(i, j) = Format(matImprimir(i, j), "0.00")
        Next
    
        k = IIf(i = 1, 50, IIf(i = 2, 80, IIf(i = 3, 100, IIf(i = 4, 350, IIf(i = 5, 400, IIf(i = 6, 600, IIf(i = 7, 700, IIf(i = 8, 900, IIf(i = 9, 1020, 0)))))))))
        
        If k = 0 Then
            k = IIf(i = 10, 1100, IIf(i = 11, 1130, IIf(i = 12, 1180, IIf(i = 13, 1200, IIf(i = 14, 1250, IIf(i = 15, 1300, IIf(i = 16, 1350, IIf(i = 17, 1360, IIf(i = 18, 1375, 0)))))))))
        End If
        
        If k = 0 Then
            k = IIf(i = 19, 1400, IIf(i = 20, 1450, IIf(i = 21, 1475, IIf(i = 22, 1500, IIf(i = 23, 1510, IIf(i = 24, 1515, IIf(i = 25, 1517, IIf(i = 26, 1520, IIf(i = 27, 1550, 1600)))))))))
        End If
            
        matImprimir(i, 1) = k
          
    Next
      
    Open psArchivoAGrabar For Output As #1

    Print #1, "02060100" & gsCodCMAC & Format(ldFecha, "YYYYMMDD") & "012"
    sCad = ""

    For i = 1 To 28
        sCad = ""
        For j = 1 To 5
            If j = 1 Then
                sCad = sCad & Right("    " & Trim(Str(matImprimir(i, j))), 4)
            ElseIf j = 2 Or j = 4 Then
                sCad = sCad & LlenaCerosSUCAVE(matImprimir(i, j), 3)
            Else
                sCad = sCad & LlenaCerosSUCAVE(matImprimir(i, j))
            End If
        Next
        
'        If I = 5 Or I = 7 Then
'            sCad = sCad & "   s"
'        ElseIf I = 8 Then
'            sCad = sCad & " s  "
'        ElseIf I >= 19 And I <= 24 Then
'            sCad = sCad & "ssss"
'        ElseIf I = 25 Then
'            sCad = sCad & "ss  "
'        Else
'            sCad = sCad & "    "
'        End If
        Print #1, sCad
    Next
     
    Close #1
    bRep1 = True
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Genera02:
    bRep2 = False
    
    'Segundo Archivo a Grabar
    psArchivoAGrabar = App.path & "\SPOOLER\02" & Format(pdFecha, "YYMMdd") & ".206"
    
    bEncontrado = False
    For Each xlHoja1 In xlLibro.Worksheets
        If Trim(UCase(xlHoja1.Name)) = Trim(UCase(Hojas(2))) Then
            bEncontrado = True
            xlHoja1.Activate
            Exit For
        End If
    Next

    If bEncontrado = False Then
        'ExcelEnd psArchivoAGrabar, xlAplicacion, xlLibro, xlHoja1, True
        'MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        'Exit Sub
        GoTo Genera03
    End If
  
    For i = 1 To 28
        
        nFil = i + 11
          
        For j = 2 To 5
            matImprimir(i, j) = xlHoja1.Cells(nFil, j + 1)
            matImprimir(i, j) = Format(matImprimir(i, j), "0.00")
        Next
    
        k = IIf(i = 1, 100, IIf(i = 2, 200, IIf(i = 3, 300, IIf(i = 4, 400, IIf(i = 5, 420, IIf(i = 6, 450, IIf(i = 7, 500, IIf(i = 8, 600, IIf(i = 9, 700, 0)))))))))
        
        If k = 0 Then
            k = IIf(i = 10, 800, IIf(i = 11, 900, IIf(i = 12, 1000, IIf(i = 13, 1100, IIf(i = 14, 1150, IIf(i = 15, 1180, IIf(i = 16, 1200, IIf(i = 17, 1300, IIf(i = 18, 1400, 0)))))))))
        End If
        
        If k = 0 Then
            k = IIf(i = 19, 1500, IIf(i = 20, 1600, IIf(i = 21, 1700, IIf(i = 22, 1800, IIf(i = 23, 1900, IIf(i = 24, 2000, 2100))))))
        End If
            
        matImprimir(i, 1) = k
          
    Next
      
    Open psArchivoAGrabar For Output As #1

    Print #1, "02060200" & gsCodCMAC & Format(ldFecha, "YYYYMMDD") & "012"
    sCad = ""

    For i = 1 To 25
        sCad = ""
        For j = 1 To 5
            If j = 1 Then
                sCad = sCad & Right("    " & Trim(Str(matImprimir(i, j))), 4)
            ElseIf j = 2 Or j = 4 Then
                sCad = sCad & LlenaCerosSUCAVE(matImprimir(i, j), 3)
            Else
                sCad = sCad & LlenaCerosSUCAVE(matImprimir(i, j))
            End If
        Next
 
        Print #1, sCad
    Next
     
    Close #1
    
    bRep2 = True
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Genera03:

    bRep3 = False
    
    'Tercer archivo a grabar
    psArchivoAGrabar = App.path & "\SPOOLER\03" & Format(pdFecha, "YYMMdd") & ".206"
    
    Open psArchivoAGrabar For Output As #1

    Print #1, "02060300" & gsCodCMAC & Format(ldFecha, "YYYYMMDD") & "012"
    sCad = ""

    For i = 1 To 20
        sCad = ""
        If i >= 1 And i <= 9 Then
            k = i * 100
        ElseIf i = 10 Then
            k = 950
        Else
            k = (i - 1) * 100
        End If
        sCad = Right("    " & Trim(Str(k)), 4) & "     0     0"
        Print #1, sCad
    Next
     
    Close #1
    
    bRep3 = True

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Genera04:
    
    bRep4 = False
    
    'Cuarto Archivo a Grabar
    psArchivoAGrabar = App.path & "\SPOOLER\04" & Format(pdFecha, "YYMMdd") & ".206"
    
    bEncontrado = False
    For Each xlHoja1 In xlLibro.Worksheets
        If Trim(UCase(xlHoja1.Name)) = Trim(UCase(Hojas(3))) Then
            bEncontrado = True
            xlHoja1.Activate
            Exit For
        End If
    Next

    If bEncontrado = False Then
        'ExcelEnd psArchivoAGrabar, xlAplicacion, xlLibro, xlHoja1, True
        'MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        'Exit Sub
        GoTo Genera05
    End If
  
    For i = 1 To 43
         
        nFil = i + 11
  
        For j = 2 To 7
            matImprimir(i, j) = xlHoja1.Cells(nFil, j + 1)
            matImprimir(i, j) = Format(matImprimir(i, j), "0.00")
        Next
    
        If i <= 25 Then
            k = i * 100
        Else
            k = 0
        End If
        
        If k = 0 Then
            k = IIf(i = 26, 2700, IIf(i = 27, 2900, IIf(i = 28, 3100, IIf(i = 29, 3300, IIf(i = 30, 3400, IIf(i = 31, 3600, IIf(i = 32, 3620, IIf(i = 33, 3650, IIf(i = 34, 3800, 0)))))))))
        End If
        
        If k = 0 Then
            k = IIf(i = 35, 3820, IIf(i = 36, 3850, IIf(i = 37, 4000, IIf(i = 38, 4020, IIf(i = 39, 4050, IIf(i = 40, 4075, IIf(i = 41, 4100, IIf(i = 42, 4300, 4400))))))))
        End If
            
        matImprimir(i, 1) = k
          
    Next
      
    Open psArchivoAGrabar For Output As #1

    Print #1, "02060400" & gsCodCMAC & Format(ldFecha, "YYYYMMDD") & "012"
    sCad = ""

    For i = 1 To 43
        sCad = matImprimir(i, 1)
        sCad = ""
        
        For j = 1 To 7
            If j = 1 Then
                sCad = sCad & Right("    " & Trim(Str(matImprimir(i, j))), 4)
            ElseIf j = 2 Or j = 3 Or j = 5 Or j = 6 Then
                sCad = sCad & LlenaCerosSUCAVE(matImprimir(i, j), 3)
            Else
                sCad = sCad & LlenaCerosSUCAVE(matImprimir(i, j))
            End If
        Next
 
        Print #1, sCad
    Next
     
    Close #1
    
    bRep4 = True
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Genera05:

    bRep5 = False

    'Quinto Archivo a Grabar
    psArchivoAGrabar = App.path & "\SPOOLER\05" & Format(pdFecha, "YYMMdd") & ".206"
    
    bEncontrado = False
    For Each xlHoja1 In xlLibro.Worksheets
        If Trim(UCase(xlHoja1.Name)) = Trim(UCase(Hojas(4))) Then
            bEncontrado = True
            xlHoja1.Activate
            Exit For
        End If
    Next

    If bEncontrado = False Then
        'ExcelEnd psArchivoAGrabar, xlAplicacion, xlLibro, xlHoja1, True
        'MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        'Exit Sub
        GoTo Final
    End If
  
    For i = 1 To 21
        
        nFil = i + 11
        
        For j = 2 To 5
            matImprimir(i, j) = xlHoja1.Cells(nFil, j + 1)
            matImprimir(i, j) = Format(matImprimir(i, j), "0.00")
        Next
    
        k = i * 100
        
        matImprimir(i, 1) = k
          
    Next
      
      
    Open psArchivoAGrabar For Output As #1

    Print #1, "02060500" & gsCodCMAC & Format(ldFecha, "YYYYMMDD") & "012"
    sCad = ""

    For i = 1 To 21
        sCad = ""
        For j = 1 To 5
            If j = 1 Then
                sCad = sCad & Right("    " & Trim(Str(matImprimir(i, j))), 4)
            ElseIf j = 2 Or j = 4 Then
                sCad = sCad & LlenaCerosSUCAVE(matImprimir(i, j), 3)
            Else
                sCad = sCad & LlenaCerosSUCAVE(matImprimir(i, j))
            End If
        Next
 
        Print #1, sCad
    Next
     
    Close #1
 
Final:
 
    sCad = ""
    
    If bRep1 = True Then
        sCad = "Reporte_6A"
    End If
    If bRep2 = True Then
        If Len(Trim(sCad)) = 0 Then
            sCad = "Reporte_6B"
        Else
            sCad = sCad & Chr(13) & "Reporte_6B"
        End If
    End If
    If bRep3 = True Then
        If Len(Trim(sCad)) = 0 Then
            sCad = "Reporte_6C"
        Else
            sCad = sCad & Chr(13) & "Reporte_6C"
        End If
    End If
    If bRep4 = True Then
        If Len(Trim(sCad)) = 0 Then
            sCad = "Reporte_6D"
        Else
            sCad = sCad & Chr(13) & "Reporte_6D"
        End If
    End If
    If bRep2 = True Then
        If Len(Trim(sCad)) = 0 Then
            sCad = "Reporte_6E"
        Else
            sCad = sCad & Chr(13) & "Reporte_6E"
        End If
    End If
 
    ExcelEnd psArchivoA_Leer, xlAplicacion, xlLibro, xlHoja1, False

    If Len(Trim(sCad)) = 0 Then
        MsgBox "No se genero ningun reporte" & Chr(13) & "Probablemente no existen registros en el archivo excel"
    ElseIf Trim(sCad) = "Reporte_6C" Then
        MsgBox "Se genero Unicamente el Reporte_6C" & Chr(13) & "Verifique que existan en el archivo excel las hojas Rep_6A / Rep_6B / Rep_6D / Rep_6E"
    Else
        MsgBox "Reportes SUCAVE " & Chr(13) & "===============" & Chr(13) & sCad & Chr(13) & Chr(13) & "Generados satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
    End If
  
    Exit Sub
    
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Sub Genera6A(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
      Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim sSql As String
    Dim sSqlTemp As String
    Dim rs As New ADODB.Recordset
    Dim cPigno As String
    Dim cVigente As String
    
    'GCOLOCESTRECREFJUD
    'xxx cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "'"
    cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "' , '" & gColocEstRecVigJud & "','2205'"
    cPigno = "'" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "'"
   
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    
If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS SOBRE SALDOS", "REPORTE 6A", 6)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL     (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En Nuevos Soles)" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL PROMEDIO           ( % ) "
    xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En dolares N.A) "
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 12
    xlHoja1.Range("D1").ColumnWidth = 18
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 18
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
Else
    lnFila = 10
End If

    oBarra.Progress 1, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
        '******ALPA****2008/06/03**********************************************************************************
'        sSql = "SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, ISNULL(round(SUM(nTasaInt)/count(*),2),0) nTasaInt_Mes, ((SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol))/SUM(col.nMontoCol)) nTasaInt_Anio, "
'        sSql = sSql & "       ISNULL(SUM(CASE WHEN c.cCtaCod LIKE '_____3%' and not c.cCtaCod LIKE '_____305%' and c.nDiasAtraso >=31 and c.nDiasAtraso <=90 THEN c.nSaldoCap - c.nCapVencido "
'        sSql = sSql & "                  ELSE c.nSaldoCap END ),0) nSaldo "
'        sSql = sSql & "  FROM  " & sservidorconsolidada & "Rango r LEFT JOIN "
'        sSql = sSql & "     (SELECT c.cCtaCod, nTasaInt, c.nDiasAtraso, c.nSaldoCap, nCuotasApr, c.nCapVencido, "
'        sSql = sSql & "             nCuotasApr * CASE WHEN nPlazoApr = 0 THEN 30 ELSE NPLAZOAPR END nDias "
'        sSql = sSql & "      FROM " & sservidorconsolidada & "CREDITOSALDOCONSOL c "
'        sSql = sSql & "           JOIN " & sservidorconsolidada & "CREDITOCONSOLTOTAL ct ON ct.cCtaCod = c.cCtaCod "
'        sSql = sSql & "      WHERE  DATEDIFF(D,DFECHA,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 And ct.cRefinan = 'N' "
'        sSql = sSql & "         AND c.nPrdEstado IN (" & cVigente & ", " & cPigno & ")  and c.nSaldoCap > 0 "
'        sSql = sSql & "         AND SUBSTRING(c.cCtaCod,9,1) = '" & psMoneda & "'  and not SubString(c.cCtaCod,6,3) IN ('320') "
'        sSql = sSql & "         AND ((c.cCtaCod LIKE '_____1%' and c.nDiasAtraso < 16) or (c.cCtaCod LIKE '_____2%' and c.nDiasAtraso < 31) or (c.cCtaCod LIKE '_____[34]%' and not c.cCtaCod LIKE '_____305%' and c.nDiasAtraso <= 90) or (c.cCtaCod LIKE '_____305%' and c.nDiasAtraso < 31) )"
'        sSql = sSql & "     ) C ON C.nDias BETWEEN r.nRangoIni and r.nRangoFin And nSaldoCap BETWEEN r.nMontoIni and r.nMontoFin "
'        sSql = sSql & "        And r.cProdCod = SubString(C.cCtaCod,6,1) "
'        sSql = sSql & "        And ((r.bPrdIn = 1  And Substring(c.cCtaCod,6,3) In (Select cPrdIn From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod))"
'        sSql = sSql & "        Or   (r.bPrdOut = 1 And Substring(c.cCtacod,6,3) Not In (Select cPrdOut From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod)))"
'        sSql = sSql & "left join colocaciones col on (C.cCtaCod=col.cCtaCod)"
'        sSql = sSql & "         WHERE r.cTipoAnx = 'A' And r.nMoneda = " & psMoneda & " GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod "
'        sSql = sSql & "         ORDER BY nRangoCod "
        '***END**ALPA*************************************************************************************************
    'End If
    
        sSql = " SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, "
        'sSql = sSql & " ISNULL(round(SUM(nTasaInt)/count(*),2),0) nTasaInt_Mes, "
        sSql = sSql & " round(ISNULL(SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "  THEN case "
        sSql = sSql & "   when(c.nSaldoCap - c.nCapVencido) = 0 "
        sSql = sSql & "  then  0 "
        sSql = sSql & "   else  nTasaInt END "
        sSql = sSql & "  else nTasaInt end)/ "
        sSql = sSql & "  SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "     THEN case"
        sSql = sSql & "       when(c.nSaldoCap - c.nCapVencido) = 0 "
        sSql = sSql & "                        then  1 " 'NAGL 20190918 Cambió de 0 a 1 Según Correo e INC1909180006
        sSql = sSql & "    else  1 end "
        sSql = sSql & "      else 1 end),0),2) nTasaInt_Mes, "
        'sSql = sSql & " ISNULL((SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol))/SUM(col.nMontoCol),0) nTasaInt_Anio, "
        sSql = sSql & "  ROUND(ISNULL((SUM( "
        sSql = sSql & "      CASE WHEN cSitCtb='2' "
        sSql = sSql & "             THEN CASE "
        sSql = sSql & "                 WHEN(c.nSaldoCap - c.nCapVencido) = 0"
        sSql = sSql & "                     THEN  0"
        sSql = sSql & "                 ELSE  (power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol "
        sSql = sSql & "                 End "
        sSql = sSql & "            ELSE (power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol "
        sSql = sSql & "      End "
        sSql = sSql & "  ))/SUM(CASE WHEN cSitCtb='2' "
        sSql = sSql & "              THEN CASE "
        sSql = sSql & "                        WHEN(c.nSaldoCap - c.nCapVencido) = 0"
        sSql = sSql & "                          THEN 1" 'NAGL 20190918 Cambió de 0 a 1 Según Correo
        sSql = sSql & "              ELSE  col.nMontoCol END"
        sSql = sSql & "     ELSE "
        sSql = sSql & "             Col.nMontoCol "
        sSql = sSql & "     end),0),2) nTasaInt_Anio,"
        'sSql = sSql & " ISNULL(SUM(CASE WHEN C.cTpoCredCod LIKE '7%' AND NOT C.cTpoCredCod = '755' and c.nDiasAtraso >=31 and c.nDiasAtraso <=90 THEN c.nSaldoCap - c.nCapVencido  ELSE c.nSaldoCap END),0) nSaldo "
        sSql = sSql & " ISNULL(SUM(CASE WHEN cSitCtb='2' THEN c.nSaldoCap - c.nCapVencido  ELSE c.nSaldoCap END),0) nSaldo "
        sSql = sSql & " FROM  " & sservidorconsolidada & "Rangox r "
        sSql = sSql & " LEFT JOIN( "
        sSql = sSql & " SELECT cc.cSitCtb,c.cCtaCod, Ct.cTpoCredCod, ct.nTasaInt, c.nDiasAtraso, c.nSaldoCap, ct.nCuotasApr, c.nCapVencido, "
        sSql = sSql & " ct.nCuotasApr * CASE WHEN ct.nPlazoApr = 0 THEN 30 ELSE ct.nPlazoApr END nDias "
        sSql = sSql & " FROM " & sservidorconsolidada & "CreditoSaldoConsol c "
        sSql = sSql & " JOIN " & sservidorconsolidada & "CreditoConsolTotal ct ON ct.cCtaCod = c.cCtaCod "
        sSql = sSql & " JOIN " & sservidorconsolidada & "CreditoConsol cc ON ct.cCtaCod = cc.cCtaCod "
        sSql = sSql & " WHERE  DATEDIFF(D,DFECHA,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 And ct.cRefinan = 'N' "
        sSql = sSql & " AND c.nPrdEstado IN (" & cVigente & ", " & cPigno & ") "
        sSql = sSql & " AND c.nSaldoCap > 0 AND SUBSTRING(c.cCtaCod,9,1) = " & psMoneda
        'sSql = sSql & " AND NOT Ct.cTpoCredCod IN ('756') "
        'sSql = sSql & "     AND ("
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[123]%' AND c.nDiasAtraso < 16) or"
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[45]%' AND c.nDiasAtraso < 31) or"
        'sSql = sSql & "       (CT.cTpoCredCod LIKE '[678]%' AND NOT ct.cTpoCredCod = '755' AND c.nDiasAtraso <= 90) or"
        'sSql = sSql & "       (CT.cTpoCredCod='755' AND c.nDiasAtraso < 31)"
        'sSql = sSql & "     )"
        sSql = sSql & " and cc.cSitCtb in ('1','2')"
        If psConsol = "N" Then
            sSql = sSql & " And CT.cAgeCodAct in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql & " ) C ON C.nDias BETWEEN r.nRangoIni and r.nRangoFin And nSaldoCap"
        sSql = sSql & "   Between R.nMontoIni And R.nMontoFin And R.cProdCod = SubString(c.cTpoCredCod, 1, 1)"
'        sSql = sSql & "   AND ((r.bPrdIn = 1  And C.cTpoCredCod"
'        sSql = sSql & "   IN (Select cPrdIn From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod))"
'        sSql = sSql & "     OR (r.bPrdOut = 1 And C.cTpoCredCod"
'        sSql = sSql & "     NOT In (SELECT cPrdOut From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod)))"
        sSql = sSql & "   LEFT JOIN Colocaciones col on (C.cCtaCod=col.cCtaCod)"
        sSql = sSql & " WHERE r.cTipoAnx = 'A' AND r.nMoneda = " & psMoneda & " " ' AND ISNULL((CASE WHEN cSitCtb='2' THEN c.nSaldoCap - c.nCapVencido  ELSE c.nSaldoCap END),0)<> 0" 'NAGL 20190918 Agregó la última condición
        sSql = sSql & " GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod" 'NAGL 20190918 Traslado INC1909180006
        sSql = sSql & " ORDER BY nRangoCod"
        
    
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    
    Dim Pro1, Pro2 As Double
    Dim SumSaldo As Double
    
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 2)) = rs!nSaldo
            If lnFila = 39 Or lnFila = 40 Then
'                xlHoja1.Range(xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2))).Formula = "=((" & _
'                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 1 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 1, 1 + (Val(psMoneda) * 2))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 2))).Address & ") + (" & _
'                    xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 1 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 2, 1 + (Val(psMoneda) * 2))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 2))).Address & "))/(" & _
'                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 2))).Address & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 2))).Address & ")"
                rs.MoveNext
                Pro1 = rs!nTasaInt_Anio * rs!nSaldo
                SumSaldo = rs!nSaldo
                rs.MoveNext
                Pro2 = rs!nTasaInt_Anio * rs!nSaldo
                SumSaldo = SumSaldo + rs!nSaldo
                xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = (Pro1 + Pro2) / SumSaldo
            Else
                xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = rs!nTasaInt_Anio
            End If
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub

Private Sub Genera6A_D(psMoneda As String, pdFecha As Date)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Long, lnFilaIni As Long
    Dim lsProd As String
    Dim sSql As String
    Dim sSqlTemp As String
    Dim rs As New ADODB.Recordset
    Dim cPigno As String
    Dim cVigente As String
    
    'GCOLOCESTRECREFJUD
    'xxx cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "'"
    cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "' , '" & gColocEstRecVigJud & "','2205'"
    cPigno = "'" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "'"
   
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS-DETALLE", "Cargando Datos", "", vbBlue
    
'If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS SOBRE SALDOS-DETALLE", "REPORTE 6A", 6)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    xlHoja1.Cells(lnFila, 3) = "CREDITO"
    If psMoneda = "1" Then
        xlHoja1.Cells(lnFila, 5) = "MONEDA NACIONAL"
    Else
        xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    End If
    
    If psMoneda = "1" Then
        xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL     (%)"
        xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En Nuevos Soles)"
    Else
        xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL PROMEDIO           ( % ) "
        xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En dolares N.A) "
    End If
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 12
    xlHoja1.Range("D1").ColumnWidth = 18
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 18
    
    'xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila + 1, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
'Else
'    lnFila = 10
'End If

    oBarra.Progress 1, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
   
        '******ALPA****2008/06/03**********************************************************************************
        sSql = "SELECT isnull(c.cCtaCod,0) cCtaCod,r.cProdCod Prod, r.nRangoCod, r.nRangoDes, ISNULL(round(SUM(nTasaInt)/count(*),2),0) nTasaInt_Mes, isnull(((SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol))/SUM(col.nMontoCol)),0) nTasaInt_Anio, "
        sSql = sSql & "       ISNULL(SUM(CASE WHEN c.cCtaCod LIKE '_____3%' and not c.cCtaCod LIKE '_____305%' and c.nDiasAtraso >=31 and c.nDiasAtraso <=90 THEN c.nSaldoCap - c.nCapVencido "
        sSql = sSql & "                  ELSE c.nSaldoCap END ),0) nSaldo "
        sSql = sSql & "  FROM  " & sservidorconsolidada & "Rango r LEFT JOIN "
        sSql = sSql & "     (SELECT c.cCtaCod, nTasaInt, c.nDiasAtraso, c.nSaldoCap, nCuotasApr, c.nCapVencido, "
        sSql = sSql & "             nCuotasApr * CASE WHEN nPlazoApr = 0 THEN 30 ELSE NPLAZOAPR END nDias "
        sSql = sSql & "      FROM " & sservidorconsolidada & "CREDITOSALDOCONSOL c "
        sSql = sSql & "           JOIN " & sservidorconsolidada & "CREDITOCONSOLTOTAL ct ON ct.cCtaCod = c.cCtaCod "
        sSql = sSql & "      WHERE  DATEDIFF(D,DFECHA,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 And ct.cRefinan = 'N' "
        sSql = sSql & "         AND c.nPrdEstado IN (" & cVigente & ", " & cPigno & ")  and c.nSaldoCap > 0 "
        sSql = sSql & "         AND SUBSTRING(c.cCtaCod,9,1) = '" & psMoneda & "'  and not SubString(c.cCtaCod,6,3) IN ('320') "
        sSql = sSql & "         AND ((c.cCtaCod LIKE '_____1%' and c.nDiasAtraso < 16) or (c.cCtaCod LIKE '_____2%' and c.nDiasAtraso < 31) or (c.cCtaCod LIKE '_____[34]%' and not c.cCtaCod LIKE '_____305%' and c.nDiasAtraso <= 90) or (c.cCtaCod LIKE '_____305%' and c.nDiasAtraso < 31) )"
        sSql = sSql & "     ) C ON C.nDias BETWEEN r.nRangoIni and r.nRangoFin And nSaldoCap BETWEEN r.nMontoIni and r.nMontoFin "
        sSql = sSql & "        And r.cProdCod = SubString(C.cCtaCod,6,1) "
        sSql = sSql & "        And ((r.bPrdIn = 1  And Substring(c.cCtaCod,6,3) In (Select cPrdIn From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod))"
        sSql = sSql & "        Or   (r.bPrdOut = 1 And Substring(c.cCtacod,6,3) Not In (Select cPrdOut From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod)))"
        sSql = sSql & "left join colocaciones col on (C.cCtaCod=col.cCtaCod)"
        sSql = sSql & "         WHERE r.cTipoAnx = 'A' And r.nMoneda = " & psMoneda & " GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod,c.cCtaCod "
        'sSql = sSql & "         r.nRangoDes ='%C.4 Créditos Pignoraticios%'"
        sSql = sSql & "         ORDER BY nRangoCod "
        '***END**ALPA*************************************************************************************************
             
    'End If
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        'If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        'End If
        If rs!nSaldo <> 0 Then
            If lnFila = 32 Or lnFila = 35 Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2))).Formula = "=((" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 1 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 1, 1 + (Val(psMoneda) * 2))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 2))).Address & ") + (" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 1 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 2, 1 + (Val(psMoneda) * 2))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 2))).Address & "))/(" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 2))).Address & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 2)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 2))).Address & ")"
            Else
                xlHoja1.Cells(lnFila, 5) = rs!nTasaInt_Anio
            End If
            'ALPA***30/06/2008
            xlHoja1.Cells(lnFila, 3) = rs!cCtaCod
            xlHoja1.Cells(lnFila, 6) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6A: TASAS DE INTERES ACTIVAS SOBRE SALDOS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
Private Function CabeceraReporte(pdFecha As Date, psTitulo As String, psReporte As String, pnCols As Integer) As Integer
Dim lnFila As Integer
    xlHoja1.Range("A1:R100").Font.Size = 8

    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(8, pnCols)).Font.Bold = True
    
    lnFila = 1
    xlHoja1.Cells(lnFila, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    lnFila = lnFila + 3
    xlHoja1.Cells(lnFila, 2) = psTitulo
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 5) = psReporte
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 1) = "EMPRESA : " & gsNomCmac:
    xlHoja1.Cells(lnFila + 1, 1) = "Fecha : AL " & Format(pdFecha, "dd mmmm yyyy")
    
    lnFila = lnFila + 3
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).VerticalAlignment = xlVAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).WrapText = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, pnCols)).Font.Bold = True
    CabeceraReporte = lnFila
End Function

Private Sub Genera6B(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    
If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES PASIVAS SOBRE SALDOS", "REPORTE 6B", 6)
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 12
    xlHoja1.Range("D1").ColumnWidth = 18
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 18
    
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL     (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En Nuevos Soles)" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL PROMEDIO           ( % ) "
    xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En dolares N.A) "
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
Else
    lnFila = 10
End If

    oBarra.Progress 1, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
    If gbBitCentral = True Then
        'sSql = "SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, ISNULL(nTasaInt_Anio,0) nTasaInt_Anio, ISNULL(SUM(nSaldo),0) nSaldo " _
             & "FROM Rango R LEFT JOIN " _
             & "    ( SELECT -10 nDias, ROUND(SUM(nTasaIntCTS)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCntCTS) nSaldo " _
             & "      FROM CTSConsol where nEstCtaCTS not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' " _
             & "      UNION " _
             & "      Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCntAc) nSaldo " _
             & "      From AhorroCConsol " _
             & "      where nEstCtaAC not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' " _
             & "      UNION " _
             & "      SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo " _
             & "      FROM PlazoFijoConsol pf JOIN Rango R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin " _
             & "      WHERE nEstCtaPF not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' and cTipoAnx = 'B' " _
             & "      GROUP BY R1.nRangoIni " _
             & "    ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin " _
             & "WHERE cTipoAnx = 'B' " _
             & "GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio " _
             & "ORDER BY r.nRangoCod"
        '((SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol))/SUM(col.nMontoCol)) nTasaInt_Anio
        '**ALPA***2008/06/03**************************************************************************************************************************************
        '*********************************************************************************************************************************************************
        sSql = "         SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, "
        sSql = sSql + "         ISNULL(nTasaInt_Anio,0) nTasaInt_Anio, "
        sSql = sSql + "         ISNULL(SUM(nSaldo),0) nSaldo"
        sSql = sSql + "  FROM    " & sservidorconsolidada & "Rango R"
        sSql = sSql + "           LEFT JOIN     ( SELECT -10 nDias, ROUND(SUM(nTasaIntCTS)/count(*),2) nTasaNAnual,"
        'sSql = sSql + "                                 ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
        sSql = sSql + "                                 SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldCntCTS)/sum(nSaldCntCTS) nTasaInt_Anio,"
        sSql = sSql + "                                 sum(nSaldCntCTS) nSaldo"
        sSql = sSql + "                            From " & sservidorconsolidada & "CTSConsol "
        'sSql = sSql + "                            inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "CTSConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                            where   nEstCtaCTS not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "'  "
        If psConsol = "N" Then
            sSql = sSql & "                                     And SubString(cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql + "                             Union"
        sSql = sSql + "                             Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                                     ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, "
        sSql = sSql + "                                 SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldCntAc)/sum(nSaldCntAc) nTasaInt_Anio,"
        sSql = sSql + "                                 sum(nSaldCntAc) nSaldo"
        sSql = sSql + "                             From " & sservidorconsolidada & "AhorroCConsol"
        'sSql = sSql + "                             inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "AhorroCConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                             where   nEstCtaAC not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' And nPersoneria in (1,2,3) And bInactiva = 0 "
        If psConsol = "N" Then
            sSql = sSql & "                                     And SubString(cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql + "                             Union"
        sSql = sSql + "                             SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                             ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
        sSql = sSql + "                                 SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio,"
        sSql = sSql + "                                     SUM(nSaldCntPF) nSaldo "
        sSql = sSql + "                             FROM    " & sservidorconsolidada & "PlazoFijoConsol pf "
        sSql = sSql + "                                     JOIN " & sservidorconsolidada & "Rango R1 ON pf.nPlazo "
        sSql = sSql + "                                     Between R1.nRangoIni And R1.nRangoFin "
        'sSql = sSql + "                                     inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "PlazoFijoConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                             WHERE   nEstCtaPF not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' And nPersoneria In (1,2,3) And cTipoAnx = 'B' "
        sSql = sSql + "                                     AND NOT EXISTS (SELECT  PC.CCTACOD "
        sSql = sSql + "                                                     FROM    " & sservidorconsolidada & "ProductoBloqueosConsol PC "
        sSql = sSql + "                                                     WHERE   PC.CCTACOD = PF.CCTACOD AND "
        sSql = sSql + "                                                             cMovNroDbl IS NULL "
        sSql = sSql + "                                                     AND nBlqMotivo = 3)"
        If psConsol = "N" Then
            sSql = sSql & "                                     AND SubString(cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql + "                                     GROUP BY R1.nRangoIni) Dat "
        sSql = sSql + "                                     ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql + "                                     WHERE cTipoAnx = 'B' AND nRangoCod <> 20"
        sSql = sSql + "                              GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
        sSql = sSql + "      Union"
        sSql = sSql + "      SELECT Prod, nRangoCod, nRangoDes, SUM(nTasaInt_Anio) AS nTasaInt_Anio ,"
        sSql = sSql + "             SUM(nSaldo) As nSaldo "
        sSql = sSql + "      From"
        sSql = sSql + "           (SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes,"
        sSql = sSql + "                   ISNULL(nTasaInt_Anio,000000000.00) nTasaInt_Anio,"
        sSql = sSql + "                    ISNULL(SUM(nSaldo),00000000.00000) nSaldo"
        sSql = sSql + "            FROM    " & sservidorconsolidada & "Rango R"
        sSql = sSql + "                    JOIN     (  "
        sSql = sSql + "                                 SELECT  -1 nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                                         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, "
        sSql = sSql + "                                         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio, "
        sSql = sSql + "                                         SUM(nSaldCntPF) nSaldo "
        sSql = sSql + "                                  FROM    " & sservidorconsolidada & "PlazoFijoConsol pf "
        sSql = sSql + "                                  WHERE   nEstCtaPF not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' "
        sSql = sSql + "                                          AND EXISTS (   SELECT  PC.CCTACOD "
        sSql = sSql + "                                                         FROM    " & sservidorconsolidada & "ProductoBloqueosConsol PC "
        sSql = sSql + "                                                         WHERE   PC.CCTACOD = PF.CCTACOD AND "
        sSql = sSql + "                                                                 PF.nEstCtaPF IN (1100,1200) AND cMovNroDbl IS NULL "
        sSql = sSql + "                                                                 AND nBlqMotivo = 3)"
        If psConsol = "N" Then
            sSql = sSql & "                                          And SubString(cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
        End If
        sSql = sSql + "                                             ) Dat "
        sSql = sSql + "                                 ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin"
        sSql = sSql + "                                 WHERE cTipoAnx = 'B' AND nRangoCod = 20 "
        sSql = sSql + "                                 GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio ) AS x "
        sSql = sSql + "                                 GROUP BY Prod, nRangoCod, nRangoDes"
        sSql = sSql + "                                 ORDER BY Prod, nRangoCod, nRangoDes" 'EJVG20131202
        '*********************************************************************************************************************************************************
        
    Else
        sSql = "SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, ISNULL(nTasaInt_Anio,0) nTasaInt_Anio, ISNULL(SUM(nSaldo),0) nSaldo "
        sSql = sSql & "FROM " & sservidorconsolidada & "Rango R LEFT JOIN "
        sSql = sSql & "    ( SELECT -10 nDias, ROUND(SUM(nTasaIntCTS)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCntCTS) nSaldo "
        sSql = sSql & "      FROM " & sservidorconsolidada & "CTSConsol where cEstCtaCTS not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' "
        sSql = sSql & "      UNION "
        sSql = sSql & "      Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldCntAc)/sum(nSaldCntAc) nTasaInt_Anio, sum(nSaldCntAc) nSaldo "
        'sSql = sSql & "      Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCntAc) nSaldo "
        sSql = sSql & "      From " & sservidorconsolidada & "AhorroCConsol "
        sSql = sSql & "      where cEstCtaAC not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' "
        sSql = sSql & "      UNION "
        'sSql = sSql & "      SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & "      SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & "      FROM " & sservidorconsolidada & "PlazoFijoConsol pf JOIN Rango R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin "
        sSql = sSql & "      WHERE cEstCtaPF not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' and cTipoAnx = 'B' "
        sSql = sSql & "      GROUP BY R1.nRangoIni "
        sSql = sSql & "    ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql & "WHERE cTipoAnx = 'B' "
        sSql = sSql & "GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
        sSql = sSql & "ORDER BY r.nRangoCod"
    End If
    '*********************************************************************************************************************************************************
    '**END***2008/06/03**************************************************************************************************************************************
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 2)) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
Private Sub Genera6B_D(psMoneda As String, pdFecha As Date)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Long, lnFilaIni As Long
    Dim lsProd As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    
''If psMoneda = "1" Then
''    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES PASIVAS SOBRE SALDOS", "REPORTE 6B", 6)
''
''    xlHoja1.Range("A1").ColumnWidth = 23
''    xlHoja1.Range("B1").ColumnWidth = 14
''    xlHoja1.Range("C1").ColumnWidth = 12
''    xlHoja1.Range("D1").ColumnWidth = 18
''    xlHoja1.Range("E1").ColumnWidth = 14
''    xlHoja1.Range("F1").ColumnWidth = 18
''
''    xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
''    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
''    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
''
''    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL     (%)"
''    xlHoja1.Cells(lnFila + 1, 4) = "SALDO (En Nuevos Soles)"
''    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL PROMEDIO           ( % ) "
''    xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En dolares N.A) "
''
''    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
''    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
''Else
''    lnFila = 10
''End If
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS SOBRE SALDOS-DETALLE", "REPORTE 6A", 6)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
    xlHoja1.Cells(lnFila, 3) = "CREDITO"
    If psMoneda = "1" Then
        xlHoja1.Cells(lnFila, 5) = "MONEDA NACIONAL"
    Else
        xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    End If
    
    If psMoneda = "1" Then
        xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL     (%)"
        xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En Nuevos Soles)"
    Else
        xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL PROMEDIO           ( % ) "
        xlHoja1.Cells(lnFila + 1, 6) = "SALDO (En dolares N.A) "
    End If
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 12
    xlHoja1.Range("D1").ColumnWidth = 18
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 18

    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila + 1, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True

    oBarra.Progress 1, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
    If gbBitCentral = True Then
    
        sSql = "         SELECT isnull(cCtaCod,0) cCtaCod,r.cProdCod Prod, r.nRangoCod, r.nRangoDes, "
        sSql = sSql + "         ISNULL(nTasaInt_Anio,0) nTasaInt_Anio, "
        sSql = sSql + "         ISNULL(SUM(nSaldo),0) nSaldo"
        sSql = sSql + "  FROM    " & sservidorconsolidada & "Rango R"
        sSql = sSql + "           LEFT JOIN     ( SELECT cCtaCod,-10 nDias, ROUND(SUM(nTasaIntCTS)/count(*),2) nTasaNAnual,"
        'sSql = sSql + "                                 ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
        sSql = sSql + "                                 SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldCntCTS)/case sum(nSaldCntCTS) when 0 then 1 else sum(nSaldCntCTS) end nTasaInt_Anio,"
        sSql = sSql + "                                 sum(nSaldCntCTS) nSaldo"
        sSql = sSql + "                            From " & sservidorconsolidada & "CTSConsol "
        'sSql = sSql + "                            inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "CTSConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                            where   nEstCtaCTS not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' group by cCtaCod "
        sSql = sSql + "                             Union"
        sSql = sSql + "                             Select cCtaCod,-5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                                     ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, "
        sSql = sSql + "                                 SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldCntAc)/case sum(nSaldCntAc) when 0 then 1 else sum(nSaldCntAc) end  nTasaInt_Anio,"
        sSql = sSql + "                                 sum(nSaldCntAc) nSaldo"
        sSql = sSql + "                             From " & sservidorconsolidada & "AhorroCConsol"
        'sSql = sSql + "                             inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "AhorroCConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                             where   nEstCtaAC not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' And nPersoneria in (1,2,3) And bInactiva = 0 group by cCtaCod"
        sSql = sSql + "                             Union"
        sSql = sSql + "                             SELECT cCtaCod,R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                             ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
        sSql = sSql + "                                 SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case sum(nSaldCntPF) when 0 then 1 else sum(nSaldCntPF) end  nTasaInt_Anio,"
        sSql = sSql + "                                     SUM(nSaldCntPF) nSaldo "
        sSql = sSql + "                             FROM    " & sservidorconsolidada & "PlazoFijoConsol pf "
        sSql = sSql + "                                     JOIN " & sservidorconsolidada & "Rango R1 ON pf.nPlazo "
        sSql = sSql + "                                     Between R1.nRangoIni And R1.nRangoFin "
        'sSql = sSql + "                                     inner join captaciones cap on cap.cCtaCod= " & sservidorconsolidada & "PlazoFijoConsol.cCtaCod " ' Aqui Cambio
        sSql = sSql + "                             WHERE   nEstCtaPF not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' And nPersoneria In (1,2,3) And cTipoAnx = 'B' "
        sSql = sSql + "                                     AND NOT EXISTS (SELECT  PC.CCTACOD "
        sSql = sSql + "                                                     FROM    " & sservidorconsolidada & "ProductoBloqueosConsol PC "
        sSql = sSql + "                                                     WHERE   PC.CCTACOD = PF.CCTACOD AND "
        sSql = sSql + "                                                             cMovNroDbl IS NULL "
        sSql = sSql + "                                                     AND nBlqMotivo = 3)"
        sSql = sSql + "                                     GROUP BY cCtaCod,R1.nRangoIni) Dat "
        sSql = sSql + "                                     ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql + "                                     WHERE cTipoAnx = 'B' AND nRangoCod <> 20"
        sSql = sSql + "                              GROUP BY cCtaCod,r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
        'sSql = sSql + "                              ORDER BY r.cProdCod, r.nRangoCod, r.nRangoDes,cCtaCod, nTasaInt_Anio "
        sSql = sSql + "      Union"
        sSql = sSql + "      SELECT isnull(cCtaCod,0) cCtaCod,Prod, nRangoCod, nRangoDes, SUM(nTasaInt_Anio) AS nTasaInt_Anio ,"
        sSql = sSql + "             SUM(nSaldo) As nSaldo "
        sSql = sSql + "      From"
        sSql = sSql + "           (SELECT cCtaCod,r.cProdCod Prod, r.nRangoCod, r.nRangoDes,"
        sSql = sSql + "                   ISNULL(nTasaInt_Anio,000000000.00) nTasaInt_Anio,"
        sSql = sSql + "                    ISNULL(SUM(nSaldo),00000000.00000) nSaldo"
        sSql = sSql + "            FROM    " & sservidorconsolidada & "Rango R"
        sSql = sSql + "                    JOIN     (  "
        sSql = sSql + "                                 SELECT  cCtaCod,-1 nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, "
        'sSql = sSql + "                                         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, "
        sSql = sSql + "                                         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case SUM(nSaldCntPF) when 0 then 1 else SUM(nSaldCntPF)  end nTasaInt_Anio, "
        sSql = sSql + "                                         SUM(nSaldCntPF) nSaldo "
        sSql = sSql + "                                  FROM    " & sservidorconsolidada & "PlazoFijoConsol pf "
        sSql = sSql + "                                  WHERE   nEstCtaPF not in (1300,1400) and Substring(cCtaCod,9,1) = '" & psMoneda & "' "
        sSql = sSql + "                                          AND EXISTS (   SELECT  PC.CCTACOD "
        sSql = sSql + "                                                         FROM    " & sservidorconsolidada & "ProductoBloqueosConsol PC "
        sSql = sSql + "                                                         WHERE   PC.CCTACOD = PF.CCTACOD AND "
        sSql = sSql + "                                                                 PF.nEstCtaPF IN (1100,1200) AND cMovNroDbl IS NULL "
        sSql = sSql + "                                                                 AND nBlqMotivo = 3)"
        sSql = sSql + "                                   group by cCtaCod) Dat "
        sSql = sSql + "                                 ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin"
        sSql = sSql + "                                 WHERE cTipoAnx = 'B' AND nRangoCod = 20 "
        sSql = sSql + "                                 GROUP BY cCtaCod,r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio ) AS x "
        sSql = sSql + "                                 GROUP BY cCtaCod,Prod, nRangoCod, nRangoDes"
        sSql = sSql + "                                 order BY Prod,nRangoCod, nRangoDes,cCtaCod"
        '*********************************************************************************************************************************************************
        
    Else
        sSql = "SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, ISNULL(nTasaInt_Anio,0) nTasaInt_Anio, ISNULL(SUM(nSaldo),0) nSaldo "
        sSql = sSql & "FROM " & sservidorconsolidada & "Rango R LEFT JOIN "
        sSql = sSql & "    ( SELECT -10 nDias, ROUND(SUM(nTasaIntCTS)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCntCTS) nSaldo "
        sSql = sSql & "      FROM " & sservidorconsolidada & "CTSConsol where cEstCtaCTS not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' "
        sSql = sSql & "      UNION "
        sSql = sSql & "      Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldCntAc)/sum(nSaldCntAc) nTasaInt_Anio, sum(nSaldCntAc) nSaldo "
        'sSql = sSql & "      Select -5 nDias, ROUND(SUM(nTasaIntAC)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCntAc) nSaldo "
        sSql = sSql & "      From " & sservidorconsolidada & "AhorroCConsol "
        sSql = sSql & "      where cEstCtaAC not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' "
        sSql = sSql & "      UNION "
        'sSql = sSql & "      SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & "      SELECT R1.nRangoIni nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & "      FROM " & sservidorconsolidada & "PlazoFijoConsol pf JOIN Rango R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin "
        sSql = sSql & "      WHERE cEstCtaPF not in ('C','U') and Substring(cCodCta,6,1) = '" & psMoneda & "' and cTipoAnx = 'B' "
        sSql = sSql & "      GROUP BY R1.nRangoIni "
        sSql = sSql & "    ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql & "WHERE cTipoAnx = 'B' "
        sSql = sSql & "GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
        sSql = sSql & "ORDER BY r.nRangoCod"
    End If
    '*********************************************************************************************************************************************************
    '**END***2008/06/03**************************************************************************************************************************************
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "ANEXO 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        'If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
       ' End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 3) = rs!cCtaCod
            xlHoja1.Cells(lnFila, 5) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 6) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6B: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    'If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    'End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
Private Sub Genera6D(psMoneda As String, pdFecha As Date)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim cJudicial As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    cJudicial = gColocEstRecVigJud & ", " & gColocEstRecVigCast & ", " & gColocEstSolic & ", " & gColocEstSug & ", " & gColocEstRetirado & ", " & gColocEstRech
    
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6D: TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue

If psMoneda = "1" Then
    xlHoja1.PageSetup.Zoom = 80
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Reporte 6D", 8)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 6) = "MONEDA EXTRANJERA"
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL     (%)"
    xlHoja1.Cells(lnFila + 1, 4) = "PROMEDIO DE COMISIONES, PORTES Y OTROS CARGOS DIRECTOS (en nuevos soles)"
    xlHoja1.Cells(lnFila + 1, 5) = "MONTO DESEMBOLSADO (en nuevos soles)  "
    xlHoja1.Cells(lnFila + 1, 6) = "TASA EFECTIVA ANUAL            ( % ) "
    xlHoja1.Cells(lnFila + 1, 7) = "PROMEDIO DE COMISIONES, PORTES Y OTROS CARGOS DIRECTOS (en nuevos soles)"
    xlHoja1.Cells(lnFila + 1, 8) = "MONTO DESEMBOLSADO (en nuevos soles)   "
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 10
    xlHoja1.Range("D1").ColumnWidth = 14
    xlHoja1.Range("E1").ColumnWidth = 17
    xlHoja1.Range("F1").ColumnWidth = 13
    xlHoja1.Range("G1").ColumnWidth = 12
    xlHoja1.Range("H1").ColumnWidth = 12
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 8)).MergeCells = True
Else
    lnFila = 10
End If

    oBarra.Progress 1, "REPORTE 6D: TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
    'If gbBitCentral = True Then
       ' sSql = "SELECT r.nRangoCod, r.nRangoDes, r.cProdCod Prod, " _
             & "       round(SUM(nTasaInt)/count(*),2) nTasaInt_Mes, (power(1+ round(SUM(nTasaInt)/count(*),2)/100,12)-1)*100 nTasaInt_Anio, Sum(nMontoDesemb) nSaldo " _
             & "FROM Rango r LEFT JOIN ( " _
             & "     SELECT c.cCtaCod, nTasaInt, nMontoDesemb, " _
             & "            nCuotasApr *  CASE WHEN nPlazoApr = 0 THEN 30 ELSE NPLAZOAPR END nDias " _
             & "     FROM CreditoConsolTotal c " _
             & "     WHERE datediff(m,dFecVig,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and not SubString(c.cCtaCod,6,3) IN ('320','305','423') " _
             & "       and SubString(cCtaCod,9,1) = '" & psMoneda & "' and NOT nPrdEstado in (" & cJudicial & ") " _
             & "       and cRefinan = 'N' " _
             & "     ) c ON c.nDias BETWEEN r.nRangoIni and r.nRangoFin and r.cProdCod = SubString(c.cCtaCod,6,1) " _
             & "WHERE r.cTipoAnx = 'D' " _
             & "GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod " _
             & "ORDER BY nRangoCod"
    'Else
'        sSql = "SELECT r.nRangoCod, r.nRangoDes, r.cProdCod Prod, " _
'             & "       round(SUM(nTasaInt)/count(*),2) nTasaInt_Mes,Round(SUM(CASE WHEN SubString(c.cCodCta,3,3) = '305' THEN nTasaInt * 100 ELSE (power(1+ nTasaInt/100 ,12)-1)*100 END)/Count(*),2) nTasaInt_Anio, Sum(nMontoDesemb) nSaldo, round(SUM(nGasto)/count(*),2) PromedioGasto " _
'             & "FROM Rango r LEFT JOIN ( " _
'             & "     SELECT c.cCodCta, nTasaInt, nMontoDesemb, " _
'             & "            nCuotasApr *  CASE WHEN nPlazoApr = 0 THEN 30 ELSE NPLAZOAPR END nDias, IsNull((Select Sum(IsNull(KC.nOtrGas,0)) From KardexConsol KC Where KC.cCodOpe like '0101%' And c.cCodCta = KC.cCodCta),0) nGasto   " _
'             & "     FROM CreditoConsolTotal c " _
'             & "     WHERE datediff(m,dFecVig,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and not SubString(c.cCodCta,3,3) IN ('320','423') " _
'             & "       and SubString(cCodCta,6,1) = '" & psMoneda & "' and NOT cEstado in ('V','W','A','B','K','L','1','4','6','7') " _
'             & "       and cRefinan = 'N' " _
'             & "     ) c ON c.nDias BETWEEN r.nRangoIni and r.nRangoFin and r.cProdCod = SubString(c.cCodCta,3,1) " _
'             & "     And nMontoDesemb BETWEEN r.nMontoIni and r.nMontoFin " _
'             & "     And ((r.bPrdIn = 1  And Substring(c.cCodCta,3,3) In (Select cPrdIn From RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod))" _
'             & "     Or   (r.bPrdOut = 1 And Substring(c.cCodCta,3,3) Not In (Select cPrdOut From RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod)))" _
'             & " WHERE r.cTipoAnx = 'D' And r.nMoneda = " & psMoneda & " " _
'             & " GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod " _
'             & " ORDER BY nRangoCod"
    '*********************************************************************************************************************************************
    '**ALPA***2008/06/03**************************************************************************************************************************
    sSql = "SELECT r.nRangoCod, r.nRangoDes, r.cProdCod Prod, "
    sSql = sSql & "       round(SUM(nTasaInt)/count(*),2) nTasaIt_Mes, "
    'sSql = sSql & "Round(SUM((power(1+ nTasaInt/100 ,12)-1)*100)/Count(*),2) nTasaInt_Anio, "
    sSql = sSql & "(SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol)/sum(col.nMontoCol)) nTasaInt_Anio, "
    sSql = sSql & "Sum(nMontoDesemb) nSaldo, round(SUM(nGasto)/count(*),2) PromedioGasto "
    sSql = sSql & "FROM " & sservidorconsolidada & "Rango r LEFT JOIN ( "
    sSql = sSql & "     SELECT c.cCtaCod, nTasaInt, nMontoDesemb, "
    sSql = sSql & "            nCuotasApr *  CASE WHEN nPlazoApr = 0 THEN 30 ELSE NPLAZOAPR END nDias, "
    sSql = sSql & "            0 as  nGasto   "
    sSql = sSql & "     FROM " & sservidorconsolidada & "CreditoConsolTotal c "
    sSql = sSql & "     WHERE datediff(m,dFecVig,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and not SubString(c.cCtaCod,6,3) IN ('320','423') "
    sSql = sSql & "        and SubString(cCtaCod,9,1) = '" & psMoneda & "' and nPrdEstado IN ('2020', '2021', '2022', '2030', '2031', '2032', '2101', '2104', '2106', '2107') "
    sSql = sSql & "       and cRefinan = 'N' "
    sSql = sSql & "     ) c ON c.nDias BETWEEN r.nRangoIni and r.nRangoFin and r.cProdCod = SubString(c.cCtaCod,6,1) "
    sSql = sSql & "     And nMontoDesemb BETWEEN r.nMontoIni and r.nMontoFin "
    sSql = sSql & "     And ((r.bPrdIn = 1  And Substring(c.cCtaCod,6,3) In (Select cPrdIn From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod))"
    sSql = sSql & "     Or   (r.bPrdOut = 1 And Substring(c.cCtaCod,6,3) Not In (Select cPrdOut From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod)))"
    sSql = sSql & "     left join colocaciones col on (C.cCtaCod=col.cCtaCod) "
    sSql = sSql & " WHERE r.cTipoAnx = 'D' And r.nMoneda = " & psMoneda & " "
    sSql = sSql & " GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod "
    sSql = sSql & " ORDER BY nRangoCod"
    'End If
    'END******************************************************************************************************************************************
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6D: TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
            If lnFila = 45 Or lnFila = 48 Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila, (Val(psMoneda) * 3))).Formula = "=((" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, (Val(psMoneda) * 3))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3))).Address & ") + (" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 2, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, (Val(psMoneda) * 3))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3))).Address & "))/(" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3))).Address & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3))).Address & ")"
            Else
                xlHoja1.Cells(lnFila, (Val(psMoneda) * 3)) = rs!nTasaInt_Anio
            End If
            xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 3)) = rs!PromedioGasto
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 3)) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6D: TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 8)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00"
    End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
Private Sub Genera6D_D(psMoneda As String, pdFecha As Date)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Long, lnFilaIni As Long
    Dim lsProd As String
    Dim cJudicial As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    cJudicial = gColocEstRecVigJud & ", " & gColocEstRecVigCast & ", " & gColocEstSolic & ", " & gColocEstSug & ", " & gColocEstRetirado & ", " & gColocEstRech
    
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6D: TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue

'If psMoneda = "1" Then
    xlHoja1.PageSetup.Zoom = 80
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Reporte 6D", 6)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    
    xlHoja1.Cells(lnFila + 1, 3) = "CREDITO"
    xlHoja1.Cells(lnFila + 1, 4) = "TASA EFECTIVA ANUAL     (%)"
    xlHoja1.Cells(lnFila + 1, 5) = "PROMEDIO DE COMISIONES, PORTES Y OTROS CARGOS DIRECTOS (en nuevos soles)"
    xlHoja1.Cells(lnFila + 1, 6) = "MONTO DESEMBOLSADO (en nuevos soles)  "
    
  If psMoneda = "1" Then
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
  Else
    xlHoja1.Cells(lnFila, 6) = "MONEDA EXTRANJERA"
  End If
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 10
    xlHoja1.Range("D1").ColumnWidth = 14
    xlHoja1.Range("E1").ColumnWidth = 17
    xlHoja1.Range("F1").ColumnWidth = 13
    xlHoja1.Range("G1").ColumnWidth = 12
    xlHoja1.Range("H1").ColumnWidth = 12
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    'xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 8)).MergeCells = True
'Else
'    lnFila = 10
'End If

    oBarra.Progress 1, "REPORTE 6D: TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    '*********************************************************************************************************************************************
    '**ALPA***2008/06/03**************************************************************************************************************************
    sSql = "SELECT c.cCtaCod,r.nRangoCod, r.nRangoDes, r.cProdCod Prod, "
    sSql = sSql & "       round(SUM(nTasaInt)/count(*),2) nTasaIt_Mes, "
    'sSql = sSql & "Round(SUM((power(1+ nTasaInt/100 ,12)-1)*100)/Count(*),2) nTasaInt_Anio, "
    sSql = sSql & "(SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol)/sum(col.nMontoCol)) nTasaInt_Anio, "
    sSql = sSql & "Sum(nMontoDesemb) nSaldo, round(SUM(nGasto)/count(*),2) PromedioGasto "
    sSql = sSql & "FROM " & sservidorconsolidada & "Rango r LEFT JOIN ( "
    sSql = sSql & "     SELECT c.cCtaCod, nTasaInt, nMontoDesemb, "
    sSql = sSql & "            nCuotasApr *  CASE WHEN nPlazoApr = 0 THEN 30 ELSE NPLAZOAPR END nDias, "
    sSql = sSql & "            0 as  nGasto   "
    sSql = sSql & "     FROM " & sservidorconsolidada & "CreditoConsolTotal c "
    sSql = sSql & "     WHERE datediff(m,dFecVig,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and not SubString(c.cCtaCod,6,3) IN ('320','423') "
    sSql = sSql & "        and SubString(cCtaCod,9,1) = '" & psMoneda & "' and nPrdEstado IN ('2020', '2021', '2022', '2030', '2031', '2032', '2101', '2104', '2106', '2107') "
    sSql = sSql & "       and cRefinan = 'N' "
    sSql = sSql & "     ) c ON c.nDias BETWEEN r.nRangoIni and r.nRangoFin and r.cProdCod = SubString(c.cCtaCod,6,1) "
    sSql = sSql & "     And nMontoDesemb BETWEEN r.nMontoIni and r.nMontoFin "
    sSql = sSql & "     And ((r.bPrdIn = 1  And Substring(c.cCtaCod,6,3) In (Select cPrdIn From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod))"
    sSql = sSql & "     Or   (r.bPrdOut = 1 And Substring(c.cCtaCod,6,3) Not In (Select cPrdOut From " & sservidorconsolidada & "RangoDet RD Where RD.cTipoAnx = r.cTipoAnx And RD.cProdCod = r.cProdCod And RD.nRangoCod = r.nRangoCod)))"
    sSql = sSql & "     left join colocaciones col on (C.cCtaCod=col.cCtaCod) "
    sSql = sSql & " WHERE r.cTipoAnx = 'D' And r.nMoneda = " & psMoneda & " "
    sSql = sSql & " GROUP BY c.cCtaCod,r.nRangoCod, r.nRangoDes, r.cProdCod "
    sSql = sSql & " ORDER BY nRangoCod"
    'End If
    'END******************************************************************************************************************************************
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6D: TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        'If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                'xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                'xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        'End If
        If rs!nSaldo <> 0 Then
            If lnFila = 45 Or lnFila = 48 Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila, (Val(psMoneda) * 3))).Formula = "=((" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, (Val(psMoneda) * 3))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3))).Address & ") + (" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 2, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, (Val(psMoneda) * 3))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3))).Address & "))/(" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3))).Address & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3))).Address & ")"
            Else
                xlHoja1.Cells(lnFila, 4) = rs!nTasaInt_Anio
            End If
            xlHoja1.Cells(lnFila, 3) = rs!cCtaCod
            xlHoja1.Cells(lnFila, 5) = rs!PromedioGasto
            xlHoja1.Cells(lnFila, 6) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6D: TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    'If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    'End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
Private Sub Genera6E(psMoneda As String, pdFecha As Date)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim oConecta As DConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    Dim lsFechaIni As String
    Dim lsFechaFin As String
    
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6E: TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
   
   lsFechaIni = "01/" & Mid(CStr(pdFecha), 4, 2) & "/" & Mid(CStr(pdFecha), 7, 4)
   lsFechaFin = CStr(pdFecha)

If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Reporte 6E", 6)
    xlHoja1.Range("A1").ColumnWidth = 19
    xlHoja1.Range("B1").ColumnWidth = 18
    xlHoja1.Range("C1").ColumnWidth = 11
    xlHoja1.Range("D1").ColumnWidth = 13
    xlHoja1.Range("E1").ColumnWidth = 11
    xlHoja1.Range("F1").ColumnWidth = 14
    
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
    xlHoja1.Cells(lnFila + 1, 4) = "MONTO RECIBIDO 2/                                                      (en nuevos soles)"
    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
    xlHoja1.Cells(lnFila + 1, 6) = "MONTO RECIBIDO 2/                                                      (en dólares de N.A.)"
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
Else
    lnFila = 10
End If

    oBarra.Progress 1, "ANEXO 6E: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
    If gbBitCentral = True Then
        
        Set oConecta = New DConecta
        oConecta.AbreConexion
        '*************************************************************************************************************************************************
        '**ALPA**2008/06/03*******************************************************************************************************************************
        sSql = "SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio, SUM(nSaldo) nSaldo "
        sSql = sSql & "FROM " & sservidorconsolidada & "Rango R LEFT JOIN "
        sSql = sSql & "   ( SELECT -10 nDias, "
        sSql = sSql & "   SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldoContable)/case sum(nSaldoContable) when 0 then 1 else sum(nSaldoContable) end nTasaInt_Anio, "
        sSql = sSql & "   sum(nSaldoContable) nSaldo "
        sSql = sSql & " from Mov MV Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro left join " & sservidorconsolidada & "CTSConsol ah on ah.cctacod = MVC.cCtaCod "
        sSql = sSql & " where MVC.cOpeCod like '220[1-2]%' and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & " ' and  '" & Format(lsFechaFin, "YYYYMMdd") & "' and nMovflag =0 "
        sSql = sSql & " and substring(MVC.cctacod,6,3) = '" & Producto.gCapCTS & "' and substring(MVC.cctacod,9,1) = '" & psMoneda & "'  "
        sSql = sSql & " Union "
        sSql = sSql & " SELECT -5 nDias, SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldoContable)/case sum(nSaldoContable) when 0 then 1 else sum(nSaldoContable) end nTasaInt_Anio, sum(nSaldoContable) nSaldo "
        sSql = sSql & " FROM  Mov MV Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro left join " & sservidorconsolidada & "ahorrocConsol ah on ah.cctacod = MVC.cctacod "
        sSql = sSql & " where (MVC.cOpeCod like '2001%' or MVC.cOpeCod = '100102') "
        sSql = sSql & " And substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & " ' and  '" & Format(lsFechaFin, "YYYYMMdd") & "' and (nMovflag =0 or nMovflag <> 1 ) "
        sSql = sSql & " and substring(MVC.cctacod,6,3) = '" & Producto.gCapAhorros & "' and substring(MVC.cctacod,9,1) = '" & psMoneda & "' "
        sSql = sSql & " Union "
        sSql = sSql & " SELECT R1.nRangoIni nDias, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case SUM(nSaldCntPF) when 0 then 1 else SUM(nSaldCntPF) end  nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & " FROM Mov MV Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro left join " & sservidorconsolidada & "PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod "
        sSql = sSql & " JOIN  " & sservidorconsolidada & "Rango R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin "
        sSql = sSql & " where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & " ' and  '" & Format(lsFechaFin, "YYYYMMdd") & "' and nMovflag =0 AND nRangoCod <> 20 "
        sSql = sSql & " and substring(MVC.cctacod,6,3) = '" & Producto.gCapPlazoFijo & "' and nEstCtaPF not in (1300,1400) "
        sSql = sSql & " and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E' "
        sSql = sSql & " GROUP BY R1.nRangoIni "
        sSql = sSql & "   ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql & "WHERE cTipoAnx = 'E' and r.nRangoCod <> 20  "
        sSql = sSql & "GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
        sSql = sSql & " Union "
        sSql = sSql & " SELECT Prod, nRangoCod, nRangoDes, SUM(nTasaInt_Anio*nSaldo)/case SUM(nSaldo) when 0 then 1 else SUM(nSaldo) end  AS nTasaInt_Anio ,  SUM(nSaldo) As nSaldo"
        sSql = sSql & " From  (SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes,  ISNULL(nTasaInt_Anio,000000000.00) nTasaInt_Anio, "
        sSql = sSql & " ISNULL(SUM(nSaldo),00000000.00000) nSaldo "
        sSql = sSql & " FROM DBConsolidada..Rango R "
        sSql = sSql & " JOIN (SELECT  -1 nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldoContable)/case SUM(nSaldoContable) when 0 then 1 else SUM(nSaldoContable) end nTasaInt_Anio, "
        sSql = sSql & " SUM(nSaldoContable) nSaldo "
        sSql = sSql & " FROM Mov MV INNER JOIN "
        sSql = sSql & " MovCap MVC ON MV.nMovNro=MVC.nMovNro "
        sSql = sSql & " left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod "
        sSql = sSql & " where MVC.cOpeCod LIKE '210[16]%' "
        sSql = sSql & " and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & " ' and  '" & Format(lsFechaFin, "YYYYMMdd") & "'  and mv.nmovflag=0"
        sSql = sSql & " and substring(MVC.cctacod,6,3) = '233'"
        sSql = sSql & " and substring(MVC.cctacod,9,1) ='" & psMoneda & "'"
        sSql = sSql & " and PF.cCtaCod in (select P.cCtaCod from Producto P "
        sSql = sSql & " join ProductoBloqueos PB on P.cCtaCod=PB.cCtaCod "
        sSql = sSql & " where nBlqMotivo=3 and substring(p.cctacod,6,3)='233' and nPrdEstado in(1100,1200)  "
        sSql = sSql & " )"
        sSql = sSql & " ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql & " WHERE cTipoAnx = 'E' AND nRangoCod = 20 "
        sSql = sSql & " GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio)X "
        sSql = sSql & " GROUP BY Prod, nRangoCod, nRangoDes"
   
    Else
    
    sSql = "SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio, SUM(nSaldo) nSaldo "
    sSql = sSql & "FROM " & sservidorconsolidada & "Rango R LEFT JOIN "
    'sSql = sSql & "   ( SELECT -10 nDias, ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCnt) nSaldo "
    sSql = sSql & "   ( SELECT -10 nDias, SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldCnt)/case sum(nSaldCnt) when 0 then 1 else sum(nSaldCnt) end nTasaInt_Anio, sum(nSaldCnt) nSaldo "
    sSql = sSql & "     from " & sservidorconsolidada & "transahoconsol ta left join " & sservidorconsolidada & "CTSConsol ah on ah.ccodcta = ta.ccodcta "
    sSql = sSql & "     where ccodope like '220[1-2]%' and datediff(mm,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cflag is null "
    sSql = sSql & "           and substring(ta.ccodcta,3,3) = '234' and substring(ta.ccodcta,6,1) = '" & psMoneda & "' "
    sSql = sSql & "     UNION "
'    sSql = sSql & "     SELECT -5 nDias, ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCnt) nSaldo "
    sSql = sSql & "     SELECT -5 nDias, SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldCnt)/case sum(nSaldCnt) when 0 then 1 else sum(nSaldCnt) end nTasaInt_Anio, sum(nSaldCnt) nSaldo "
    sSql = sSql & "     FROM " & sservidorconsolidada & "transahoconsol ta left join " & sservidorconsolidada & "ahorrocConsol ah on ah.ccodcta = ta.ccodcta "
    sSql = sSql & "     where (ccodope like '2001%' or ccodope = '010102') "
    sSql = sSql & "           and datediff(mm,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and (cFlag IS NULL or cFlag <> 'X') "
    sSql = sSql & "           and substring(ta.ccodcta,3,3) = '232' and substring(ta.ccodcta,6,1) = '" & psMoneda & "' "
    sSql = sSql & "     UNION "
'    sSql = sSql & "     SELECT R1.nRangoIni nDias, ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
    sSql = sSql & "     SELECT R1.nRangoIni nDias, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case SUM(nSaldCntPF) when 0 then 1 else SUM(nSaldCntPF) end  nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
    sSql = sSql & "     FROM " & sservidorconsolidada & "transahoconsol ta left join " & sservidorconsolidada & "PlazoFijoConsol pf ON pf.cCodCta = ta.cCodCta "
    sSql = sSql & "        JOIN " & sservidorconsolidada & "Rango R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin "
    sSql = sSql & "     where ccodope like '210[16]%' and datediff(mm,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cflag is null "
    sSql = sSql & "           and substring(ta.ccodcta,3,3) = '233' and cEstCtaPF not in ('C','U') "
    sSql = sSql & "        and Substring(ta.cCodCta,6,1) = '" & psMoneda & "' and cTipoAnx = 'E' "
    sSql = sSql & "     GROUP BY R1.nRangoIni "
    sSql = sSql & "   ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
    sSql = sSql & "WHERE cTipoAnx = 'E' "
    sSql = sSql & "GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
    sSql = sSql & "ORDER BY r.nRangoCod"
    
    End If
    '*************************************************************************************************************************************************
    '***END ALPA******************************************/********************************************************************************************
    lsProd = "": lnFilaIni = lnFila
    
    If gbBitCentral = True Then
        Set rs = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Else
        Set rs = oCon.CargaRecordSet(sSql)
    End If
    oBarra.Progress 2, "ANEXO 6E: TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 2)) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6E: TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
Private Sub Genera6E_D(psMoneda As String, pdFecha As Date)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Long, lnFilaIni As Long
    Dim lsProd As String
    Dim oConecta As DConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    Dim lsFechaIni As String
    Dim lsFechaFin As String
    
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6E: TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
   
   lsFechaIni = "01/" & Mid(CStr(pdFecha), 4, 2) & "/" & Mid(CStr(pdFecha), 7, 4)
   lsFechaFin = CStr(pdFecha)

'If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Reporte 6E", 6)
    xlHoja1.Range("A1").ColumnWidth = 19
    xlHoja1.Range("B1").ColumnWidth = 18
    xlHoja1.Range("C1").ColumnWidth = 11
    xlHoja1.Range("D1").ColumnWidth = 13
    xlHoja1.Range("E1").ColumnWidth = 11
    xlHoja1.Range("F1").ColumnWidth = 14
    
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
    
    If psMoneda = "1" Then
        xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
        xlHoja1.Cells(lnFila + 1, 3) = "CREDITO"
        xlHoja1.Cells(lnFila + 1, 4) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
        xlHoja1.Cells(lnFila + 1, 5) = "MONTO RECIBIDO 2/                                                      (en nuevos soles)"
    Else
        xlHoja1.Cells(lnFila, 3) = "MONEDA EXTRANJERA"
        xlHoja1.Cells(lnFila + 1, 3) = "CREDITO"
        xlHoja1.Cells(lnFila + 1, 4) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
        xlHoja1.Cells(lnFila + 1, 5) = "MONTO RECIBIDO 2/                                                      (en dólares de N.A.)"
    End If
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).MergeCells = True
    'xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
'Else
'    lnFila = 10
'End If

    oBarra.Progress 1, "ANEXO 6E: TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
    If gbBitCentral = True Then
        
        Set oConecta = New DConecta
        oConecta.AbreConexion
        '*************************************************************************************************************************************************
        '**ALPA**2008/06/03*******************************************************************************************************************************
        sSql = "SELECT isnull(cCtaCod,0) cCtaCod,r.cProdCod Prod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio, SUM(nSaldo) nSaldo "
        sSql = sSql & "FROM " & sservidorconsolidada & "Rango R LEFT JOIN "
        sSql = sSql & "   ( SELECT MVC.cCtaCod,-10 nDias, "
        sSql = sSql & "   SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldoContable)/case sum(nSaldoContable) when 0 then 1 else sum(nSaldoContable) end nTasaInt_Anio, "
        sSql = sSql & "   sum(nSaldoContable) nSaldo "
        sSql = sSql & " from Mov MV Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro left join " & sservidorconsolidada & "CTSConsol ah on ah.cctacod = MVC.cCtaCod "
        sSql = sSql & " where MVC.cOpeCod like '220[1-2]%' and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & " ' and  '" & Format(lsFechaFin, "YYYYMMdd") & "' and nMovflag =0 "
        sSql = sSql & " and substring(MVC.cctacod,6,3) = '" & Producto.gCapCTS & "' and substring(MVC.cctacod,9,1) = '" & psMoneda & "' group by MVC.cCtaCod "
        sSql = sSql & " Union "
        'sSql = sSql & " SELECT -5 nDias, ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldoContable) nSaldo "
        sSql = sSql & " SELECT MVC.cCtaCod,-5 nDias, SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldoContable)/case sum(nSaldoContable) when 0 then 1 else sum(nSaldoContable) end nTasaInt_Anio, sum(nSaldoContable) nSaldo "
        sSql = sSql & " FROM  Mov MV Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro left join " & sservidorconsolidada & "ahorrocConsol ah on ah.cctacod = MVC.cctacod "
        sSql = sSql & " where (MVC.cOpeCod like '2001%' or MVC.cOpeCod = '100102') "
        sSql = sSql & " And substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & " ' and  '" & Format(lsFechaFin, "YYYYMMdd") & "' and (nMovflag =0 or nMovflag <> 1 ) "
        sSql = sSql & " and substring(MVC.cctacod,6,3) = '" & Producto.gCapAhorros & "' and substring(MVC.cctacod,9,1) = '" & psMoneda & "' group by MVC.cCtaCod"
        sSql = sSql & " Union "
        'sSql = sSql & " SELECT R1.nRangoIni nDias, ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & " SELECT MVC.cCtaCod,R1.nRangoIni nDias, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case SUM(nSaldCntPF) when 0 then 1 else SUM(nSaldCntPF) end  nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
        sSql = sSql & " FROM Mov MV Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro left join " & sservidorconsolidada & "PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod "
        sSql = sSql & " JOIN  " & sservidorconsolidada & "Rango R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin "
        sSql = sSql & " where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & " ' and  '" & Format(lsFechaFin, "YYYYMMdd") & "' and nMovflag =0 AND nRangoCod <> 20 "
        sSql = sSql & " and substring(MVC.cctacod,6,3) = '" & Producto.gCapPlazoFijo & "' and nEstCtaPF not in (1300,1400) "
        sSql = sSql & " and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E' "
        sSql = sSql & " GROUP BY MVC.cCtaCod,R1.nRangoIni "
        sSql = sSql & "   ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql & "WHERE cTipoAnx = 'E' and r.nRangoCod <> 20  "
        sSql = sSql & "GROUP BY cCtaCod,r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
        sSql = sSql & " Union "
        'sSql = sSql & " SELECT Prod, nRangoCod, nRangoDes, SUM(nTasaInt_Anio) AS nTasaInt_Anio ,  SUM(nSaldo) As nSaldo"
        sSql = sSql & " SELECT isnull(cCtaCod,0) cCtaCod,Prod, nRangoCod, nRangoDes, SUM(nTasaInt_Anio*nSaldo)/case SUM(nSaldo) when 0 then 1 else SUM(nSaldo) end  AS nTasaInt_Anio ,  SUM(nSaldo) As nSaldo"
        sSql = sSql & " From  (SELECT Dat.cCtaCod,r.cProdCod Prod, r.nRangoCod, r.nRangoDes,  ISNULL(nTasaInt_Anio,000000000.00) nTasaInt_Anio, "
        sSql = sSql & " ISNULL(SUM(nSaldo),00000000.00000) nSaldo "
        sSql = sSql & " FROM DBConsolidada..Rango R "
        'sSql = sSql & " JOIN (SELECT  -1 nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, "
        sSql = sSql & " JOIN (SELECT  MVC.cCtaCod,-1 nDias, ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldoContable)/case SUM(nSaldoContable) when 0 then 1 else SUM(nSaldoContable) end nTasaInt_Anio, "
        sSql = sSql & " SUM(nSaldoContable) nSaldo "
        sSql = sSql & " FROM Mov MV INNER JOIN "
        sSql = sSql & " MovCap MVC ON MV.nMovNro=MVC.nMovNro "
        sSql = sSql & " left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod "
        sSql = sSql & " where MVC.cOpeCod LIKE '210[16]%' "
        sSql = sSql & " and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & " ' and  '" & Format(lsFechaFin, "YYYYMMdd") & "'  and mv.nmovflag=0"
        sSql = sSql & " and substring(MVC.cctacod,6,3) = '233'"
        sSql = sSql & " and substring(MVC.cctacod,9,1) ='" & psMoneda & "'"
        sSql = sSql & " and PF.cCtaCod in (select P.cCtaCod from Producto P "
        sSql = sSql & " join ProductoBloqueos PB on P.cCtaCod=PB.cCtaCod "
        sSql = sSql & " where nBlqMotivo=3 and substring(p.cctacod,6,3)='233' and nPrdEstado in(1100,1200)  "
        sSql = sSql & " ) group by mvc.cCtaCod"
        sSql = sSql & " ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
        sSql = sSql & " WHERE cTipoAnx = 'E' AND nRangoCod = 20 "
        sSql = sSql & " GROUP BY cCtaCod,r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio)X "
        sSql = sSql & " GROUP BY cCtaCod,Prod, nRangoCod, nRangoDes"
   
    Else
    
    sSql = "SELECT r.cProdCod Prod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio, SUM(nSaldo) nSaldo "
    sSql = sSql & "FROM " & sservidorconsolidada & "Rango R LEFT JOIN "
    'sSql = sSql & "   ( SELECT -10 nDias, ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCnt) nSaldo "
    sSql = sSql & "   ( SELECT -10 nDias, SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldCnt)/case sum(nSaldCnt) when 0 then 1 else sum(nSaldCnt) end nTasaInt_Anio, sum(nSaldCnt) nSaldo "
    sSql = sSql & "     from " & sservidorconsolidada & "transahoconsol ta left join " & sservidorconsolidada & "CTSConsol ah on ah.ccodcta = ta.ccodcta "
    sSql = sSql & "     where ccodope like '220[1-2]%' and datediff(mm,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cflag is null "
    sSql = sSql & "           and substring(ta.ccodcta,3,3) = '234' and substring(ta.ccodcta,6,1) = '" & psMoneda & "' "
    sSql = sSql & "     UNION "
'    sSql = sSql & "     SELECT -5 nDias, ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, sum(nSaldCnt) nSaldo "
    sSql = sSql & "     SELECT -5 nDias, SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldCnt)/case sum(nSaldCnt) when 0 then 1 else sum(nSaldCnt) end nTasaInt_Anio, sum(nSaldCnt) nSaldo "
    sSql = sSql & "     FROM " & sservidorconsolidada & "transahoconsol ta left join " & sservidorconsolidada & "ahorrocConsol ah on ah.ccodcta = ta.ccodcta "
    sSql = sSql & "     where (ccodope like '2001%' or ccodope = '010102') "
    sSql = sSql & "           and datediff(mm,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and (cFlag IS NULL or cFlag <> 'X') "
    sSql = sSql & "           and substring(ta.ccodcta,3,3) = '232' and substring(ta.ccodcta,6,1) = '" & psMoneda & "' "
    sSql = sSql & "     UNION "
'    sSql = sSql & "     SELECT R1.nRangoIni nDias, ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
    sSql = sSql & "     SELECT R1.nRangoIni nDias, SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case SUM(nSaldCntPF) when 0 then 1 else SUM(nSaldCntPF) end  nTasaInt_Anio, SUM(nSaldCntPF) nSaldo "
    sSql = sSql & "     FROM " & sservidorconsolidada & "transahoconsol ta left join " & sservidorconsolidada & "PlazoFijoConsol pf ON pf.cCodCta = ta.cCodCta "
    sSql = sSql & "        JOIN " & sservidorconsolidada & "Rango R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin "
    sSql = sSql & "     where ccodope like '210[16]%' and datediff(mm,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cflag is null "
    sSql = sSql & "           and substring(ta.ccodcta,3,3) = '233' and cEstCtaPF not in ('C','U') "
    sSql = sSql & "        and Substring(ta.cCodCta,6,1) = '" & psMoneda & "' and cTipoAnx = 'E' "
    sSql = sSql & "     GROUP BY R1.nRangoIni "
    sSql = sSql & "   ) Dat ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin "
    sSql = sSql & "WHERE cTipoAnx = 'E' "
    sSql = sSql & "GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio "
    sSql = sSql & "ORDER BY r.nRangoCod"
    
    End If
    '*************************************************************************************************************************************************
    '***END ALPA******************************************/********************************************************************************************
    lsProd = "": lnFilaIni = lnFila
    
    If gbBitCentral = True Then
        Set rs = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Else
        Set rs = oCon.CargaRecordSet(sSql)
    End If
    oBarra.Progress 2, "ANEXO 6E: TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        'If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        'End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 3) = rs!cCtaCod
            xlHoja1.Cells(lnFila, 4) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 5) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6E: TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
'peac 20071127

Private Sub Genera6D1(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim cJudicial As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    cJudicial = gColocEstRecVigJud & ", " & gColocEstRecVigCast & ", " & gColocEstSolic & ", " & gColocEstSug & ", " & gColocEstRetirado & ", " & gColocEstRech
    
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue

If psMoneda = "1" Then
    xlHoja1.PageSetup.Zoom = 80
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Reporte 6D", 8)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 6) = "MONEDA EXTRANJERA"
    
'    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL     (%)"
'    xlHoja1.Cells(lnFila + 1, 4) = "PROMEDIO DE COMISIONES, PORTES Y OTROS CARGOS DIRECTOS (en nuevos soles)"
'    xlHoja1.Cells(lnFila + 1, 5) = "MONTO DESEMBOLSADO (en nuevos soles)  "
'    xlHoja1.Cells(lnFila + 1, 6) = "TASA EFECTIVA ANUAL            ( % ) "
'    xlHoja1.Cells(lnFila + 1, 7) = "PROMEDIO DE COMISIONES, PORTES Y OTROS CARGOS DIRECTOS (en nuevos soles)"
'    xlHoja1.Cells(lnFila + 1, 8) = "MONTO DESEMBOLSADO (en nuevos soles)   "
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL PROMEDIO   (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "MONTO DESEMBOLSADO (en nuevos soles)  " 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "MONTO DESEMBOLSADO (en " & StrConv(gcPEN_PLURAL, vbLowerCase) & ")  "    'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA DE COSTO EFECTIVO ANUAL PROMEDIO    (%)"
    xlHoja1.Cells(lnFila + 1, 6) = "TASA EFECTIVA ANUAL PROMEDIO   (%)"
    xlHoja1.Cells(lnFila + 1, 7) = "MONTO DESEMBOLSADO (en Dolares americanos)  "
    xlHoja1.Cells(lnFila + 1, 8) = "TASA DE COSTO EFECTIVO ANUAL PROMEDIO    (%)"
       
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 10
    xlHoja1.Range("D1").ColumnWidth = 14
    xlHoja1.Range("E1").ColumnWidth = 17
    xlHoja1.Range("F1").ColumnWidth = 13
    xlHoja1.Range("G1").ColumnWidth = 12
    xlHoja1.Range("H1").ColumnWidth = 12
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 8)).MergeCells = True
Else
    lnFila = 10
End If

    oBarra.Progress 1, "REPORTE 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
'declare @fecha char(8),@moneda char(1)
'set @fecha='20071031'
'set @moneda='1'
'***BRGO**2010/09/09*********************************************************************
sSql = " SELECT  r.nRangoCod,"
sSql = sSql & "    r.nRangoDes,"
sSql = sSql & "     r.cProdCod Prod,"
sSql = sSql & "     round(SUM(C.nTasaInt)/count(*),2) nTasaInt_Mes,"
sSql = sSql & "     round(SUM(C.ntasCosEfeAnu)/count(*),2) nTCEA,"
sSql = sSql & "     (SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol)/sum(col.nMontoCol)) nTasaInt_Anio, "
sSql = sSql & "     Sum(C.nMontoDesemb) nSaldo,"
sSql = sSql & "     round(SUM(C.nGasto)/count(*),2) PromedioGasto"
sSql = sSql & " FROM " & sservidorconsolidada & "Rangox r"
sSql = sSql & " LEFT JOIN (SELECT   W.cCtaCod, W.cTpoCredCod, "
sSql = sSql & "             W.nTasaInt,"
sSql = sSql & "             isnull(cc.ntasCosEfeAnu,0) ntasCosEfeAnu,"
sSql = sSql & "             W.nMontoDesemb,"
sSql = sSql & "             W.nCuotasApr * CASE WHEN W.nPlazoApr = 0 THEN 30 ELSE W.NPLAZOAPR END nDias,"
sSql = sSql & "             0 as  nGasto "
sSql = sSql & "         FROM " & sservidorconsolidada & "CreditoConsolTotal W"
sSql = sSql & "         left JOIN ColocacCred cc on W.cctacod=cc.cctacod"
sSql = sSql & "         WHERE datediff(m,W.dFecVig,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 "
sSql = sSql & "         and SubString(W.cCtaCod,9,1) = " & psMoneda & " and W.nPrdEstado"
sSql = sSql & "         IN (2020, 2021, 2022, 2030, 2031, 2032, 2101, 2104, 2106, 2107, 2022, 2092, 2201, 2205)"
'ALPA 20110706*****************************
If psConsol = "N" Then
    sSql = sSql & "      And W.cAgeCodAct in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
'******************************************
sSql = sSql & "     ) C ON C.nDias BETWEEN r.nRangoIni and r.nRangoFin"
sSql = sSql & "     and r.cProdCod = SubString(C.cTpoCredCod,1,1)"
sSql = sSql & "     And nMontoDesemb BETWEEN r.nMontoIni and r.nMontoFin"
sSql = sSql & "     And ((r.bPrdIn = 1"
sSql = sSql & "         And C.cTpoCredCod In (Select cPrdIn From " & sservidorconsolidada & "RangoDet1 RD"
sSql = sSql & "                         Where rd.cTipoAnx = R.cTipoAnx And rd.cProdCod = R.cProdCod"
sSql = sSql & "                         And RD.nRangoCod = r.nRangoCod))"
sSql = sSql & "         Or (r.bPrdOut = 1 And C.cTpoCredCod "
sSql = sSql & "             Not In (Select cPrdOut From " & sservidorconsolidada & "RangoDet1 RD"
sSql = sSql & "                     Where rd.cTipoAnx = R.cTipoAnx And rd.cProdCod = R.cProdCod"
sSql = sSql & "                     And RD.nRangoCod = r.nRangoCod)))"
sSql = sSql & "left join colocaciones col on (C.cCtaCod=col.cCtaCod) "
sSql = sSql & " WHERE r.cTipoAnx = 'D' And r.nMoneda = " & psMoneda
sSql = sSql & " GROUP BY r.nRangoCod, r.nRangoDes, r.cProdCod"
sSql = sSql & " ORDER BY nRangoCod"
'***ALPA**END****************************************************************************
    
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
'            If lnFila = 45 Or lnFila = 48 Then
'                xlHoja1.Range(xlHoja1.Cells(lnFila, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila, (Val(psMoneda) * 3))).Formula = "=((" & _
'                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, (Val(psMoneda) * 3))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3))).Address & ") + (" & _
'                    xlHoja1.Range(xlHoja1.Cells(lnFila + 2, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, (Val(psMoneda) * 3))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3))).Address & "))/(" & _
'                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3))).Address & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3))).Address & ")"
'            Else
                xlHoja1.Cells(lnFila, (Val(psMoneda) * 3)) = rs!nTasaInt_Anio
'            End If
            xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 3)) = rs!nSaldo 'rs!PromedioGasto
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 3)) = rs!nTCEA  'rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 8)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00"
    End If
    
'    xlHoja1.Range("C17").FormulaR1C1 = "=+((R18C3*R18C4)+(R19C3*R19C4))/IF((R18C4+R19C4)=0,1,(R18C4+R19C4))"
'    xlHoja1.Range("D17").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E17").FormulaR1C1 = "=+((R18C4*R18C5)+(R19C4*R19C5))/IF((R18C4+R19C4)=0,1,(R18C4+R19C4))"
'    '--
'    xlHoja1.Range("F17").FormulaR1C1 = "=+((R18C6*R18C7)+(R19C6*R19C7))/IF((R18C7+R19C7)=0,1,(R18C7+R19C7))"
'    xlHoja1.Range("G17").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H17").FormulaR1C1 = "=+((R18C7*R18C8)+(R19C7*R19C8))/IF((R18C7+R19C7)=0,1,(R18C7+R19C7))"
'
'
'    xlHoja1.Range("C20").FormulaR1C1 = "=+((R21C3*R21C4)+(R22C3*R22C4))/IF((R21C4+R22C4)=0,1,(R21C4+R22C4))"
'    xlHoja1.Range("D20").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E20").FormulaR1C1 = "=+((R21C4*R21C5)+(R22C4*R22C5))/IF((R21C4+R22C4)=0,1,(R21C4+R22C4))"
'    '--
'    xlHoja1.Range("F20").FormulaR1C1 = "=+((R21C6*R21C7)+(R22C6*R22C7))/IF((R21C7+R22C7)=0,1,(R21C7+R22C7))"
'    xlHoja1.Range("G20").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H20").FormulaR1C1 = "=+((R21C7*R21C8)+(R22C7*R22C8))/IF((R21C7+R22C7)=0,1,(R21C7+R22C7))"
'
'
'    xlHoja1.Range("C23").FormulaR1C1 = "=+((R24C3*R24C4)+(R25C3*R25C4))/IF((R24C4+R25C4)=0,1,(R24C4+R25C4))"
'    xlHoja1.Range("D23").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E23").FormulaR1C1 = "=+((R24C4*R24C5)+(R25C4*R25C5))/IF((R24C4+R25C4)=0,1,(R24C4+R25C4))"
'    '--
'    xlHoja1.Range("F23").FormulaR1C1 = "=+((R24C6*R24C7)+(R25C6*R25C7))/IF((R24C7+R25C7)=0,1,(R24C7+R25C7))"
'    xlHoja1.Range("G23").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H23").FormulaR1C1 = "=+((R24C7*R24C8)+(R25C7*R25C8))/IF((R24C7+R25C7)=0,1,(R24C7+R25C7))"
'
'
'    xlHoja1.Range("C26").FormulaR1C1 = "=+((R27C3*R27C4)+(R28C3*R28C4))/IF((R27C4+R28C4)=0,1,(R27C4+R28C4))"
'    xlHoja1.Range("D26").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E26").FormulaR1C1 = "=+((R27C4*R27C5)+(R28C4*R28C5))/IF((R27C4+R28C4)=0,1,(R27C4+R28C4))"
'    '--
'    xlHoja1.Range("F26").FormulaR1C1 = "=+((R27C6*R27C7)+(R28C6*R28C7))/IF((R27C7+R28C7)=0,1,(R27C7+R28C7))"
'    xlHoja1.Range("G26").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H26").FormulaR1C1 = "=+((R27C7*R27C8)+(R28C7*R28C8))/IF((R27C7+R28C7)=0,1,(R27C7+R28C7))"
'
'
'    xlHoja1.Range("C29").FormulaR1C1 = "=+((R30C3*R30C4)+(R31C3*R31C4))/IF((R30C4+R31C4)=0,1,(R30C4+R31C4))"
'    xlHoja1.Range("D29").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E29").FormulaR1C1 = "=+((R30C4*R30C5)+(R31C4*R31C5))/IF((R30C4+R31C4)=0,1,(R30C4+R31C4))"
'    '--
'    xlHoja1.Range("F29").FormulaR1C1 = "=+((R30C6*R30C7)+(R31C6*R31C7))/IF((R30C7+R31C7)=0,1,(R30C7+R31C7))"
'    xlHoja1.Range("G29").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H29").FormulaR1C1 = "=+((R30C7*R30C8)+(R31C7*R31C8))/IF((R30C7+R31C7)=0,1,(R30C7+R31C7))"
''------------
'    xlHoja1.Range("C45").FormulaR1C1 = "=+((R46C3*R46C4)+(R47C3*R47C4))/IF((R46C4+R47C4)=0,1,(R46C4+R47C4))"
'    xlHoja1.Range("D45").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E45").FormulaR1C1 = "=+((R46C4*R46C5)+(R47C4*R47C5))/IF((R46C4+R47C4)=0,1,(R46C4+R47C4))"
'    '--
'    xlHoja1.Range("F45").FormulaR1C1 = "=+((R46C6*R46C7)+(R47C6*R47C7))/IF((R46C7+R47C7)=0,1,(R46C7+R47C7))"
'    xlHoja1.Range("G45").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H45").FormulaR1C1 = "=+((R46C7*R46C8)+(R47C7*R47C8))/IF((R46C7+R47C7)=0,1,(R46C7+R47C7))"
'
'
'    xlHoja1.Range("C48").FormulaR1C1 = "=+((R49C3*R49C4)+(R50C3*R50C4))/IF((R49C4+R50C4)=0,1,(R49C4+R50C4))"
'    xlHoja1.Range("D48").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E48").FormulaR1C1 = "=+((R49C4*R49C5)+(R50C4*R50C5))/IF((R49C4+R50C4)=0,1,(R49C4+R50C4))"
'    '--
'    xlHoja1.Range("F48").FormulaR1C1 = "=+((R49C6*R49C7)+(R50C6*R50C7))/IF((R49C7+R50C7)=0,1,(R49C7+R50C7))"
'    xlHoja1.Range("G48").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H48").FormulaR1C1 = "=+((R49C7*R49C8)+(R50C7*R50C8))/IF((R49C7+R50C7)=0,1,(R49C7+R50C7))"
    
    xlHoja1.Range("C17:H17").Font.Bold = True
    xlHoja1.Range("C20:H20").Font.Bold = True
    xlHoja1.Range("C23:H23").Font.Bold = True
    xlHoja1.Range("C26:H26").Font.Bold = True
    xlHoja1.Range("C29:H29").Font.Bold = True
    xlHoja1.Range("C45:H45").Font.Bold = True
    xlHoja1.Range("C48:H48").Font.Bold = True
    
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub

'peac 20071121
Private Sub Genera6D1_D(psMoneda As String, pdFecha As Date)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Long, lnFilaIni As Long
    Dim lsProd As String
    Dim cJudicial As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    cJudicial = gColocEstRecVigJud & ", " & gColocEstRecVigCast & ", " & gColocEstSolic & ", " & gColocEstSug & ", " & gColocEstRetirado & ", " & gColocEstRech
    
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue

'If psMoneda = "1" Then
    xlHoja1.PageSetup.Zoom = 80
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Reporte 6D", 6)
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES ACTIVAS"
    

    If psMoneda = "1" Then
        xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
        xlHoja1.Cells(lnFila + 1, 3) = "CREDITO"
        xlHoja1.Cells(lnFila + 1, 4) = "TASA EFECTIVA ANUAL PROMEDIO   (%)"
        xlHoja1.Cells(lnFila + 1, 5) = "MONTO DESEMBOLSADO (en nuevos soles)  "
        xlHoja1.Cells(lnFila + 1, 6) = "TASA DE COSTO EFECTIVO ANUAL PROMEDIO    (%)"
    
    Else
        xlHoja1.Cells(lnFila, 6) = "MONEDA EXTRANJERA"
        xlHoja1.Cells(lnFila + 1, 3) = "CREDITO"
        xlHoja1.Cells(lnFila + 1, 4) = "TASA EFECTIVA ANUAL PROMEDIO   (%)"
        xlHoja1.Cells(lnFila + 1, 5) = "MONTO DESEMBOLSADO (en nuevos soles)  "
        xlHoja1.Cells(lnFila + 1, 6) = "TASA DE COSTO EFECTIVO ANUAL PROMEDIO    (%)"
    End If
    
    
       
    
    xlHoja1.Range("A1").ColumnWidth = 23
    xlHoja1.Range("B1").ColumnWidth = 14
    xlHoja1.Range("C1").ColumnWidth = 10
    xlHoja1.Range("D1").ColumnWidth = 14
    xlHoja1.Range("E1").ColumnWidth = 17
    xlHoja1.Range("F1").ColumnWidth = 13
    xlHoja1.Range("G1").ColumnWidth = 12
    xlHoja1.Range("H1").ColumnWidth = 12
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 6)).MergeCells = True
    'xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 8)).MergeCells = True
'Else
'    lnFila = 10
'End If

    oBarra.Progress 1, "REPORTE 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2
    
'declare @fecha char(8),@moneda char(1)
'set @fecha='20071031'
'set @moneda='1'
'***ALPA**2008/06/03*********************************************************************
sSql = " SELECT  C.cCtaCod,r.nRangoCod,"
sSql = sSql & "    r.nRangoDes,"
sSql = sSql & "     r.cProdCod Prod,"
sSql = sSql & "     round(SUM(C.nTasaInt)/count(*),2) nTasaInt_Mes,"
sSql = sSql & "     round(SUM(C.ntasCosEfeAnu)/count(*),2) nTCEA,"
'sSql = sSql & "     Round(SUM((power(1+ C.nTasaInt/100 ,12)-1)*100)/Count(*),2) nTasaInt_Anio,"
sSql = sSql & "     (SUM((power(1+(convert(decimal(12,2),C.nTasaInt)/100.00),12) -1) * 100.00*col.nMontoCol)/sum(col.nMontoCol)) nTasaInt_Anio, "
sSql = sSql & "     Sum(C.nMontoDesemb) nSaldo,"
sSql = sSql & "     round(SUM(C.nGasto)/count(*),2) PromedioGasto"
sSql = sSql & " FROM DBConsolidada..Rango1 r"
sSql = sSql & " LEFT JOIN (SELECT   W.cCtaCod,"
sSql = sSql & "             W.nTasaInt,"
sSql = sSql & "             isnull(cc.ntasCosEfeAnu,0) ntasCosEfeAnu,"
sSql = sSql & "             W.nMontoDesemb,"
sSql = sSql & "             W.nCuotasApr * CASE WHEN W.nPlazoApr = 0 THEN 30 ELSE W.NPLAZOAPR END nDias,"
sSql = sSql & "             0 as  nGasto"
sSql = sSql & "         FROM DBConsolidada..CreditoConsolTotal W"
sSql = sSql & "         left JOIN ColocacCred cc on W.cctacod=cc.cctacod"
sSql = sSql & "         WHERE datediff(m,W.dFecVig,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and not SubString(W.cCtaCod,6,3) IN ('320','423')"
sSql = sSql & "         and SubString(W.cCtaCod,9,1) = " & psMoneda & " and W.nPrdEstado"
sSql = sSql & "         IN ('2020', '2021', '2022', '2030', '2031', '2032', '2101', '2104', '2106', '2107')"
sSql = sSql & "         and W.cRefinan = 'N') C"
sSql = sSql & "     ON C.nDias BETWEEN r.nRangoIni and r.nRangoFin"
sSql = sSql & "     and r.cProdCod = SubString(C.cCtaCod,6,1)"
sSql = sSql & "     And nMontoDesemb BETWEEN r.nMontoIni and r.nMontoFin"
sSql = sSql & "     And ((r.bPrdIn = 1"
sSql = sSql & "         And Substring(C.cCtaCod,6,3) In (Select cPrdIn From DBConsolidada..RangoDet1 RD"
sSql = sSql & "                         Where rd.cTipoAnx = R.cTipoAnx And rd.cProdCod = R.cProdCod"
sSql = sSql & "                         And RD.nRangoCod = r.nRangoCod))"
sSql = sSql & "         Or (r.bPrdOut = 1 And Substring(C.cCtaCod,6,3)"
sSql = sSql & "             Not In (Select cPrdOut From DBConsolidada..RangoDet1 RD"
sSql = sSql & "                     Where rd.cTipoAnx = R.cTipoAnx And rd.cProdCod = R.cProdCod"
sSql = sSql & "                     And RD.nRangoCod = r.nRangoCod)))"
sSql = sSql & "left join colocaciones col on (C.cCtaCod=col.cCtaCod) "
sSql = sSql & " WHERE r.cTipoAnx = 'D' And r.nMoneda = " & psMoneda
sSql = sSql & " GROUP BY C.cCtaCod,r.nRangoCod, r.nRangoDes, r.cProdCod"
sSql = sSql & " ORDER BY nRangoCod"
'***ALPA**END****************************************************************************
    
    lsProd = "": lnFilaIni = lnFila
    Set rs = oCon.CargaRecordSet(sSql)
    oBarra.Progress 2, "REPORTE 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        'If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        'End If
        If rs!nSaldo <> 0 Then
            If lnFila = 45 Or lnFila = 48 Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila, (Val(psMoneda) * 3))).Formula = "=((" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, (Val(psMoneda) * 3))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3))).Address & ") + (" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 2, (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, (Val(psMoneda) * 3))).Address & "*" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3))).Address & "))/(" & _
                    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 1, 2 + (Val(psMoneda) * 3))).Address & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3)), xlHoja1.Cells(lnFila + 2, 2 + (Val(psMoneda) * 3))).Address & ")"
            Else
                xlHoja1.Cells(lnFila, 4) = rs!nTasaInt_Anio
            End If
            xlHoja1.Cells(lnFila, 3) = rs!cCtaCod
            xlHoja1.Cells(lnFila, 5) = rs!nSaldo 'rs!PromedioGasto
            xlHoja1.Cells(lnFila, 6) = rs!nTCEA  'rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6D (nuevo): TASAS DE INTERES ACTIVAS DE OPERACIONES DIARIAS*", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    'If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        'xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 4), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    'End If
    
'    xlHoja1.Range("C17").FormulaR1C1 = "=+((R18C3*R18C4)+(R19C3*R19C4))/IF((R18C4+R19C4)=0,1,(R18C4+R19C4))"
'    xlHoja1.Range("D17").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E17").FormulaR1C1 = "=+((R18C4*R18C5)+(R19C4*R19C5))/IF((R18C4+R19C4)=0,1,(R18C4+R19C4))"
'    '--
'    xlHoja1.Range("F17").FormulaR1C1 = "=+((R18C6*R18C7)+(R19C6*R19C7))/IF((R18C7+R19C7)=0,1,(R18C7+R19C7))"
'    xlHoja1.Range("G17").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H17").FormulaR1C1 = "=+((R18C7*R18C8)+(R19C7*R19C8))/IF((R18C7+R19C7)=0,1,(R18C7+R19C7))"
'
'
'    xlHoja1.Range("C20").FormulaR1C1 = "=+((R21C3*R21C4)+(R22C3*R22C4))/IF((R21C4+R22C4)=0,1,(R21C4+R22C4))"
'    xlHoja1.Range("D20").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E20").FormulaR1C1 = "=+((R21C4*R21C5)+(R22C4*R22C5))/IF((R21C4+R22C4)=0,1,(R21C4+R22C4))"
'    '--
'    xlHoja1.Range("F20").FormulaR1C1 = "=+((R21C6*R21C7)+(R22C6*R22C7))/IF((R21C7+R22C7)=0,1,(R21C7+R22C7))"
'    xlHoja1.Range("G20").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H20").FormulaR1C1 = "=+((R21C7*R21C8)+(R22C7*R22C8))/IF((R21C7+R22C7)=0,1,(R21C7+R22C7))"
'
'
'    xlHoja1.Range("C23").FormulaR1C1 = "=+((R24C3*R24C4)+(R25C3*R25C4))/IF((R24C4+R25C4)=0,1,(R24C4+R25C4))"
'    xlHoja1.Range("D23").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E23").FormulaR1C1 = "=+((R24C4*R24C5)+(R25C4*R25C5))/IF((R24C4+R25C4)=0,1,(R24C4+R25C4))"
'    '--
'    xlHoja1.Range("F23").FormulaR1C1 = "=+((R24C6*R24C7)+(R25C6*R25C7))/IF((R24C7+R25C7)=0,1,(R24C7+R25C7))"
'    xlHoja1.Range("G23").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H23").FormulaR1C1 = "=+((R24C7*R24C8)+(R25C7*R25C8))/IF((R24C7+R25C7)=0,1,(R24C7+R25C7))"
'
'
'    xlHoja1.Range("C26").FormulaR1C1 = "=+((R27C3*R27C4)+(R28C3*R28C4))/IF((R27C4+R28C4)=0,1,(R27C4+R28C4))"
'    xlHoja1.Range("D26").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E26").FormulaR1C1 = "=+((R27C4*R27C5)+(R28C4*R28C5))/IF((R27C4+R28C4)=0,1,(R27C4+R28C4))"
'    '--
'    xlHoja1.Range("F26").FormulaR1C1 = "=+((R27C6*R27C7)+(R28C6*R28C7))/IF((R27C7+R28C7)=0,1,(R27C7+R28C7))"
'    xlHoja1.Range("G26").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H26").FormulaR1C1 = "=+((R27C7*R27C8)+(R28C7*R28C8))/IF((R27C7+R28C7)=0,1,(R27C7+R28C7))"
'
'
'    xlHoja1.Range("C29").FormulaR1C1 = "=+((R30C3*R30C4)+(R31C3*R31C4))/IF((R30C4+R31C4)=0,1,(R30C4+R31C4))"
'    xlHoja1.Range("D29").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E29").FormulaR1C1 = "=+((R30C4*R30C5)+(R31C4*R31C5))/IF((R30C4+R31C4)=0,1,(R30C4+R31C4))"
'    '--
'    xlHoja1.Range("F29").FormulaR1C1 = "=+((R30C6*R30C7)+(R31C6*R31C7))/IF((R30C7+R31C7)=0,1,(R30C7+R31C7))"
'    xlHoja1.Range("G29").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H29").FormulaR1C1 = "=+((R30C7*R30C8)+(R31C7*R31C8))/IF((R30C7+R31C7)=0,1,(R30C7+R31C7))"
''------------
'    xlHoja1.Range("C45").FormulaR1C1 = "=+((R46C3*R46C4)+(R47C3*R47C4))/IF((R46C4+R47C4)=0,1,(R46C4+R47C4))"
'    xlHoja1.Range("D45").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E45").FormulaR1C1 = "=+((R46C4*R46C5)+(R47C4*R47C5))/IF((R46C4+R47C4)=0,1,(R46C4+R47C4))"
'    '--
'    xlHoja1.Range("F45").FormulaR1C1 = "=+((R46C6*R46C7)+(R47C6*R47C7))/IF((R46C7+R47C7)=0,1,(R46C7+R47C7))"
'    xlHoja1.Range("G45").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H45").FormulaR1C1 = "=+((R46C7*R46C8)+(R47C7*R47C8))/IF((R46C7+R47C7)=0,1,(R46C7+R47C7))"
'
'
'    xlHoja1.Range("C48").FormulaR1C1 = "=+((R49C3*R49C4)+(R50C3*R50C4))/IF((R49C4+R50C4)=0,1,(R49C4+R50C4))"
'    xlHoja1.Range("D48").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("E48").FormulaR1C1 = "=+((R49C4*R49C5)+(R50C4*R50C5))/IF((R49C4+R50C4)=0,1,(R49C4+R50C4))"
'    '--
'    xlHoja1.Range("F48").FormulaR1C1 = "=+((R49C6*R49C7)+(R50C6*R50C7))/IF((R49C7+R50C7)=0,1,(R49C7+R50C7))"
'    xlHoja1.Range("G48").FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'    xlHoja1.Range("H48").FormulaR1C1 = "=+((R49C7*R49C8)+(R50C7*R50C8))/IF((R49C7+R50C7)=0,1,(R49C7+R50C7))"
    
'    xlHoja1.Range("C17:H17").Font.Bold = True
'    xlHoja1.Range("C20:H20").Font.Bold = True
'    xlHoja1.Range("C23:H23").Font.Bold = True
'    xlHoja1.Range("C26:H26").Font.Bold = True
'    xlHoja1.Range("C29:H29").Font.Bold = True
'    xlHoja1.Range("C45:H45").Font.Bold = True
'    xlHoja1.Range("C48:H48").Font.Bold = True
    
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
Private Sub Genera6E1(psMoneda As String, pdFecha As Date, psAgeCod As String, psConsol As String)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Integer, lnFilaIni As Integer
    Dim lsProd As String
    Dim oConecta As DConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    Dim lsFechaIni As String
    Dim lsFechaFin As String
        
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
   
   lsFechaIni = "01/" & Mid(CStr(pdFecha), 4, 2) & "/" & Mid(CStr(pdFecha), 7, 4)
   lsFechaFin = CStr(pdFecha)

If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Reporte 6E", 6)
    xlHoja1.Range("A1").ColumnWidth = 19
    xlHoja1.Range("B1").ColumnWidth = 18
    xlHoja1.Range("C1").ColumnWidth = 11
    xlHoja1.Range("D1").ColumnWidth = 13
    xlHoja1.Range("E1").ColumnWidth = 11
    xlHoja1.Range("F1").ColumnWidth = 14
    
    xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
    xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 5) = "MONEDA EXTRANJERA"
    
    xlHoja1.Cells(lnFila + 1, 3) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
    '''xlHoja1.Cells(lnFila + 1, 4) = "MONTO RECIBIDO 2/                                                      (en nuevos soles)" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 4) = "MONTO RECIBIDO 2/                                                      (en " & StrConv(gcPEN_PLURAL, vbLowerCase) & ")" 'MARG ERS044-2016
    xlHoja1.Cells(lnFila + 1, 5) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
    xlHoja1.Cells(lnFila + 1, 6) = "MONTO RECIBIDO 2/                                                      (en dólares de N.A.)"
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
Else
    lnFila = 10
End If

    oBarra.Progress 1, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2

'DECLARE @INI CHAR(8), @FIN CHAR(8),@MONEDA CHAR(1)
'@INI='20071001'
'@FIN='20071031'
'@MONEDA='1'

Set oConecta = New DConecta
oConecta.AbreConexion
'***ALPA**2008/06/03*********************************************************************
sSql = " SELECT  r.cProdCod Prod,"
sSql = sSql & "     r.nRangoCod,"
sSql = sSql & "     r.nRangoDes,"
sSql = sSql & "     nTasaInt_Anio,"
sSql = sSql & "     SUM(nSaldo) nSaldo"
sSql = sSql & " FROM DBConsolidada..Rango1 R"
sSql = sSql & " LEFT JOIN ( SELECT 31 nRango,-10 nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntCTS/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldoContable)/sum(nSaldoContable) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN sum(nSaldoContable) = 0 THEN 0 ELSE SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldoContable)/sum(nSaldoContable) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         sum(nSaldoContable) nSaldo"
sSql = sSql & "         from Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..CTSConsol ah on ah.cctacod = MVC.cCtaCod"
sSql = sSql & "         where MVC.cOpeCod like '220[1-2]%' and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and  '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag =0 and substring(MVC.cctacod,6,3) = '234' and substring(MVC.cctacod,9,1) = '" & psMoneda & "'"
If psConsol = "N" Then
    sSql = sSql & "         and SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "     Union"
sSql = sSql & "     SELECT 2 nRango,-5 nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntAC/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldoContable)/sum(nSaldoContable) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN sum(nSaldoContable) = 0 THEN 0 ELSE SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldoContable)/sum(nSaldoContable) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         sum(nSaldoContable) nSaldo"
sSql = sSql & "         FROM  Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..ahorrocConsol ah on ah.cctacod = MVC.cctacod"
sSql = sSql & "         where (MVC.cOpeCod like '2001%' or MVC.cOpeCod = '100102')"
sSql = sSql & "         And substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and (nMovflag =0 or nMovflag <> 1 )"
sSql = sSql & "         and substring(MVC.cctacod,6,3) = '232' and substring(MVC.cctacod,9,1) = '" & psMoneda & "'"
If psConsol = "N" Then
    sSql = sSql & "         and SubString(isnull(MVC.cCtaCod,'10901'),4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "     Union"
sSql = sSql & "     SELECT max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN SUM(nSaldCntPF) = 0 THEN 0 ELSE SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria=1"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and r1.nRangoCod between 14 and 18"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
If psConsol = "N" Then
    sSql = sSql & "         And SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "         GROUP BY R1.nRangoIni"
sSql = sSql & "     Union"
sSql = sSql & "     SELECT max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN SUM(nSaldCntPF) = 0 THEN 0 ELSE SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria=2"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and nRangoCod between 20 and 24"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
If psConsol = "N" Then
    sSql = sSql & "         And SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "         GROUP BY R1.nRangoIni"
sSql = sSql & "     Union"
sSql = sSql & "     SELECT max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
'sSql = sSql & "         ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         CASE WHEN SUM(nSaldCntPF) = 0 THEN 0 ELSE SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/SUM(nSaldCntPF) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria not in (1,2)"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and nRangoCod between 26 and 30"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
If psConsol = "N" Then
    sSql = sSql & "         And SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "         GROUP BY R1.nRangoIni) Dat"
sSql = sSql & "         ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin and dat.nRango=r.nRangoCod"
sSql = sSql & " WHERE cTipoAnx = 'E' and r.nRangoCod <> 32"
sSql = sSql & " GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio"
sSql = sSql & " Union"
sSql = sSql & " SELECT  Prod,"
sSql = sSql & "     nRangoCod,"
sSql = sSql & "     nRangoDes,"
sSql = sSql & "     SUM(nTasaInt_Anio) AS nTasaInt_Anio ,"
sSql = sSql & "     SUM(nSaldo) As nSaldo"
sSql = sSql & "     From  (SELECT r.cProdCod Prod,"
sSql = sSql & "             r.nRangoCod,"
sSql = sSql & "             r.nRangoDes,"
sSql = sSql & "             ISNULL(nTasaInt_Anio,000000000.00) nTasaInt_Anio,"
sSql = sSql & "             ISNULL(SUM(nSaldo),00000000.00000) nSaldo"
sSql = sSql & "         FROM DBConsolidada..Rango1 R"
sSql = sSql & "         JOIN (SELECT  -1 nDias,"
sSql = sSql & "                 ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual,"
'sSql = sSql & "                  ROUND(SUM((power(1+nTasaIntPF/36000,360) -1) * 100)/count(*),2) nTasaInt_Anio,"
'sSql = sSql & "                  SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldoContable)/SUM(nSaldoContable) nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "                 CASE WHEN SUM(nSaldoContable) = 0 THEN 0 ELSE SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldoContable)/SUM(nSaldoContable) END nTasaInt_Anio," 'FRHU 20151102 INC1511020005
sSql = sSql & "                 SUM(nSaldoContable) nSaldo"
sSql = sSql & "             FROM Mov MV"
sSql = sSql & "             INNER JOIN  MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "             left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod"
sSql = sSql & "             where MVC.cOpeCod LIKE '210[16]%'"
sSql = sSql & "             and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "             and mv.nmovflag=0 and substring(MVC.cctacod,6,3) = '233'"
sSql = sSql & "             and substring(MVC.cctacod,9,1) = '" & psMoneda & "'"
sSql = sSql & "             and PF.cCtaCod in (select P.cCtaCod"
sSql = sSql & "                         from DBConsolidada..PlazoFijoConsol P"
sSql = sSql & "                         join DBConsolidada..ProductoBloqueosConsol PB on P.cCtaCod=PB.cCtaCod"
sSql = sSql & "                         where nBlqMotivo=3 and substring(p.cctacod,6,3)='233'"
If psConsol = "N" Then
    sSql = sSql & "             And SubString(MVC.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto('" & psAgeCod & "')) "
End If
sSql = sSql & "                         and nEstCtaPF in(1100,1200))) Dat"
sSql = sSql & "                     ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin"
sSql = sSql & "         WHERE cTipoAnx = 'E' AND nRangoCod = 32"
sSql = sSql & "         GROUP BY r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio)X"
sSql = sSql & " GROUP BY Prod, nRangoCod, nRangoDes"
sSql = sSql & " ORDER BY Prod, nRangoCod, nRangoDes" 'EJVG20131202
'***ALPA**END****************************************************************************
    lsProd = "": lnFilaIni = lnFila
    
    If gbBitCentral = True Then
        Set rs = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Else
        Set rs = oCon.CargaRecordSet(sSql)
    End If
    oBarra.Progress 2, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 1 + (Val(psMoneda) * 2)) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 2 + (Val(psMoneda) * 2)) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
            
            
    xlHoja1.Range("C24").FormulaR1C1 = "=+((R25C3*R25C4)+(R26C3*R26C4)+(R27C3*R27C4)+(R28C3*R28C4)+(R29C3*R29C4))/IF(SUM(R25C4:R29C4)=0,1,SUM(R25C4:R29C4))"
    xlHoja1.Range("D24").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E24").FormulaR1C1 = "=+((R25C5*R25C6)+(R26C5*R26C6)+(R27C5*R27C6)+(R28C5*R28C6)+(R29C5*R29C6))/IF(SUM(R25C6:R29C6)=0,1,SUM(R25C6:R29C6))"
    xlHoja1.Range("F24").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"

    xlHoja1.Range("C30").FormulaR1C1 = "=+((R31C3*R31C4)+(R32C3*R32C4)+(R33C3*R33C4)+(R34C3*R34C4)+(R35C3*R35C4))/IF(SUM(R31C4:R35C4)=0,1,SUM(R31C4:R35C4))"
    xlHoja1.Range("D30").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E30").FormulaR1C1 = "=+((R31C5*R31C6)+(R32C5*R32C6)+(R33C5*R33C6)+(R34C5*R34C6)+(R35C5*R35C6))/IF(SUM(R31C6:R35C6)=0,1,SUM(R31C6:R35C6))"
    xlHoja1.Range("F30").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"

    xlHoja1.Range("C36").FormulaR1C1 = "=+((R37C3*R37C4)+(R38C3*R38C4)+(R39C3*R39C4)+(R40C3*R40C4)+(R41C3*R41C4))/IF(SUM(R37C4:R41C4)=0,1,SUM(R37C4:R41C4))"
    xlHoja1.Range("D36").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E36").FormulaR1C1 = "=+((R37C5*R37C6)+(R38C5*R38C6)+(R39C5*R39C6)+(R40C5*R40C6)+(R41C5*R41C6))/IF(SUM(R37C6:R41C6)=0,1,SUM(R37C6:R41C6))"
    xlHoja1.Range("F36").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    
    xlHoja1.Range("C24:H24").Font.Bold = True
    xlHoja1.Range("C30:H30").Font.Bold = True
    xlHoja1.Range("C36:H36").Font.Bold = True
    
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub
Private Sub Genera6E1_D(psMoneda As String, pdFecha As Date)
    Dim lbExisteHoja  As Boolean
    Dim i  As Long
    Dim lnFila As Long, lnFilaIni As Long
    Dim lsProd As String
    Dim oConecta As DConecta
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    Dim lsFechaIni As String
    Dim lsFechaFin As String
        
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 2
   oBarra.Progress 0, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
   
   lsFechaIni = "01/" & Mid(CStr(pdFecha), 4, 2) & "/" & Mid(CStr(pdFecha), 7, 4)
   lsFechaFin = CStr(pdFecha)

'If psMoneda = "1" Then
    lnFila = CabeceraReporte(pdFecha, "TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Reporte 6E", 6)
    xlHoja1.Range("A1").ColumnWidth = 19
    xlHoja1.Range("B1").ColumnWidth = 18
    xlHoja1.Range("C1").ColumnWidth = 11
    xlHoja1.Range("D1").ColumnWidth = 13
    xlHoja1.Range("E1").ColumnWidth = 11
    xlHoja1.Range("F1").ColumnWidth = 14
     xlHoja1.Cells(lnFila, 1) = "OPERACIONES PASIVAS"
     xlHoja1.Cells(lnFila + 1, 3) = "CREDITOS"
    If psMoneda = "1" Then
        xlHoja1.Cells(lnFila, 3) = "MONEDA NACIONAL"
        xlHoja1.Cells(lnFila + 1, 4) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
        xlHoja1.Cells(lnFila + 1, 5) = "MONTO RECIBIDO 2/                                                      (en nuevos soles)"
    Else
        xlHoja1.Cells(lnFila, 3) = "MONEDA EXTRANJERA"
        xlHoja1.Cells(lnFila + 1, 4) = "TASA EFECTIVA ANUAL  PROMEDIO 1/                                                           (%)"
        xlHoja1.Cells(lnFila + 1, 5) = "MONTO RECIBIDO 2/                                                      (en dólares de N.A.)"
    End If
    
    
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 6)).MergeCells = True
'Else
'    lnFila = 10
'End If

    oBarra.Progress 1, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS SOBRE SALDOS", "Cargando Datos", "", vbBlue
    lnFila = lnFila + 2

'DECLARE @INI CHAR(8), @FIN CHAR(8),@MONEDA CHAR(1)
'@INI='20071001'
'@FIN='20071031'
'@MONEDA='1'

Set oConecta = New DConecta
oConecta.AbreConexion
'***ALPA**2008/06/03*********************************************************************
sSql = " SELECT  cCtaCod,r.cProdCod Prod,"
sSql = sSql & "     r.nRangoCod,"
sSql = sSql & "     r.nRangoDes,"
sSql = sSql & "     nTasaInt_Anio,"
sSql = sSql & "     SUM(nSaldo) nSaldo"
sSql = sSql & " FROM DBConsolidada..Rango1 R"
sSql = sSql & " LEFT JOIN ( SELECT MVC.cCtaCod,31 nRango,-10 nDias,"
sSql = sSql & "         SUM((power(1+nTasaIntCTS/36000,360) -1) * 100*nSaldoContable)/case sum(nSaldoContable) when 0 then 1 else sum(nSaldoContable) end nTasaInt_Anio,"
sSql = sSql & "         sum(nSaldoContable) nSaldo"
sSql = sSql & "         from Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..CTSConsol ah on ah.cctacod = MVC.cCtaCod"
sSql = sSql & "         where MVC.cOpeCod like '220[1-2]%' and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and  '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag =0 and substring(MVC.cctacod,6,3) = '234' and substring(MVC.cctacod,9,1) = '" & psMoneda & "' group by MVC.cCtaCod "
sSql = sSql & "     Union"
sSql = sSql & "     SELECT MVC.cCtaCod,2 nRango,-5 nDias,"
sSql = sSql & "         SUM((power(1+nTasaIntAC/36000,360) -1) * 100*nSaldoContable)/case sum(nSaldoContable) when 0 then 1 else sum(nSaldoContable) end  nTasaInt_Anio,"
sSql = sSql & "         sum(nSaldoContable) nSaldo"
sSql = sSql & "         FROM  Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..ahorrocConsol ah on ah.cctacod = MVC.cctacod"
sSql = sSql & "         where (MVC.cOpeCod like '2001%' or MVC.cOpeCod = '100102')"
sSql = sSql & "         And substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and (nMovflag =0 or nMovflag <> 1 )"
sSql = sSql & "         and substring(MVC.cctacod,6,3) = '232' and substring(MVC.cctacod,9,1) = '" & psMoneda & "' group by MVC.cCtaCod"
sSql = sSql & "     Union"
sSql = sSql & "     SELECT MVC.cCtaCod,max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case sum(nSaldCntPF) when 0 then 1 else sum(nSaldCntPF) end  nTasaInt_Anio,"
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria=1"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and r1.nRangoCod between 14 and 18"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
sSql = sSql & "         GROUP BY MVC.cCtaCod,R1.nRangoIni"
sSql = sSql & "     Union"
sSql = sSql & "     SELECT MVC.cCtaCod,max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case sum(nSaldCntPF) when 0 then 1 else sum(nSaldCntPF) end nTasaInt_Anio,"
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria=2"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and nRangoCod between 20 and 24"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
sSql = sSql & "         GROUP BY MVC.cCtaCod,R1.nRangoIni"
sSql = sSql & "     Union"
sSql = sSql & "     SELECT MVC.cCtaCod,max(r1.nRangoCod) nRango, R1.nRangoIni nDias,"
sSql = sSql & "         SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldCntPF)/case sum(nSaldCntPF) when 0 then 1 else sum(nSaldCntPF) end nTasaInt_Anio,"
sSql = sSql & "         SUM(nSaldCntPF) nSaldo"
sSql = sSql & "         FROM Mov MV"
sSql = sSql & "         Inner Join MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "         left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod and npersoneria not in (1,2)"
sSql = sSql & "     JOIN  DBConsolidada..Rango1 R1 ON pf.nPlazo BETWEEN R1.nRangoIni and r1.nRangoFin and nRangoCod between 26 and 30"
sSql = sSql & "         where MVC.cOpeCod LIKE '210[16]%' and substring(cMovNro, 1,8)between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "         and nMovflag = 0 AND nRangoCod <> 32  and substring(MVC.cctacod,6,3) = '233' and nEstCtaPF"
sSql = sSql & "         not in (1300,1400)  and Substring(MVC.cctacod,9,1) = '" & psMoneda & "' and cTipoAnx = 'E'"
sSql = sSql & "         GROUP BY MVC.cCtaCod,R1.nRangoIni    ) Dat"
sSql = sSql & "         ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin and dat.nRango=r.nRangoCod"
sSql = sSql & " WHERE cTipoAnx = 'E' and r.nRangoCod <> 32"
sSql = sSql & " GROUP BY cCtaCod,r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio"
sSql = sSql & " Union"
sSql = sSql & " SELECT  x.cCtaCod,Prod,"
sSql = sSql & "     nRangoCod,"
sSql = sSql & "     nRangoDes,"
sSql = sSql & "     SUM(nTasaInt_Anio) AS nTasaInt_Anio ,"
sSql = sSql & "     SUM(nSaldo) As nSaldo"
sSql = sSql & "     From  (SELECT dat.cCtaCod,r.cProdCod Prod,"
sSql = sSql & "             r.nRangoCod,"
sSql = sSql & "             r.nRangoDes,"
sSql = sSql & "             ISNULL(nTasaInt_Anio,000000000.00) nTasaInt_Anio,"
sSql = sSql & "             ISNULL(SUM(nSaldo),00000000.00000) nSaldo"
sSql = sSql & "         FROM DBConsolidada..Rango1 R"
sSql = sSql & "         JOIN (SELECT  MVC.cCtaCod,-1 nDias,"
sSql = sSql & "                 ROUND(SUM(nTasaIntPF)/count(*),2) nTasaNAnual,"
sSql = sSql & "                 SUM((power(1+nTasaIntPF/36000,360) -1) * 100*nSaldoContable)/case sum(nSaldoContable) when 0 then 1 else sum(nSaldoContable) end nTasaInt_Anio,"
sSql = sSql & "                 SUM(nSaldoContable) nSaldo"
sSql = sSql & "             FROM Mov MV"
sSql = sSql & "             INNER JOIN  MovCap MVC ON MV.nMovNro=MVC.nMovNro"
sSql = sSql & "             left join DBConsolidada..PlazoFijoConsol pf ON pf.cctacod = MVC.cctacod"
sSql = sSql & "             where MVC.cOpeCod LIKE '210[16]%'"
sSql = sSql & "             and substring(cMovNro, 1,8) between  '" & Format(lsFechaIni, "YYYYMMdd") & "' and '" & Format(lsFechaFin, "YYYYMMdd") & "'"
sSql = sSql & "             and mv.nmovflag=0 and substring(MVC.cctacod,6,3) = '233'"
sSql = sSql & "             and substring(MVC.cctacod,9,1) = '" & psMoneda & "'"
sSql = sSql & "             and PF.cCtaCod in (select P.cCtaCod"
sSql = sSql & "                         from Producto P"
sSql = sSql & "                         join ProductoBloqueos PB on P.cCtaCod=PB.cCtaCod"
sSql = sSql & "                         where nBlqMotivo=3 and substring(p.cctacod,6,3)='233'"
sSql = sSql & "                         and nPrdEstado in(1100,1200)) group by MVC.cCtaCod) Dat"
sSql = sSql & "                     ON Dat.nDias BETWEEN R.nRangoIni and nRangoFin"
sSql = sSql & "         WHERE cTipoAnx = 'E' AND nRangoCod = 32"
sSql = sSql & "         GROUP BY dat.cCtaCod,r.cProdCod, r.nRangoCod, r.nRangoDes, nTasaInt_Anio)X"
sSql = sSql & " GROUP BY x.cCtaCod,Prod, nRangoCod, nRangoDes"
'***ALPA**END****************************************************************************
    lsProd = "": lnFilaIni = lnFila
    
    If gbBitCentral = True Then
        Set rs = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Else
        Set rs = oCon.CargaRecordSet(sSql)
    End If
    oBarra.Progress 2, "ANEXO 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Cargando Datos", "", vbBlue
    oBarra.Max = rs.RecordCount
    Do While Not rs.EOF
        'If psMoneda = "1" Then
            xlHoja1.Cells(lnFila, 1) = rs!nRangoDes
            If lsProd <> rs!Prod Then
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
                lsProd = rs!Prod
            End If
        'End If
        If rs!nSaldo <> 0 Then
            xlHoja1.Cells(lnFila, 3) = rs!cCtaCod
            xlHoja1.Cells(lnFila, 4) = rs!nTasaInt_Anio
            xlHoja1.Cells(lnFila, 5) = rs!nSaldo
        End If
        lnFila = lnFila + 1
        oBarra.Progress rs.Bookmark, "Reporte 6E (nuevo): TASAS DE INTERES PASIVAS DE OPERACIONES DIARIAS", "Generando Reporte", "", vbBlue
        rs.MoveNext
    Loop
    If psMoneda = "1" Then
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 1), xlHoja1.Cells(lnFila, 6)).BorderAround xlContinuous, xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 3), xlHoja1.Cells(lnFila, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnFilaIni, 4), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00"
    End If
            
            
    xlHoja1.Range("C24").FormulaR1C1 = "=+((R25C3*R25C4)+(R26C3*R26C4)+(R27C3*R27C4)+(R28C3*R28C4)+(R29C3*R29C4))/IF(SUM(R25C4:R29C4)=0,1,SUM(R25C4:R29C4))"
    xlHoja1.Range("D24").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E24").FormulaR1C1 = "=+((R25C5*R25C6)+(R26C5*R26C6)+(R27C5*R27C6)+(R28C5*R28C6)+(R29C5*R29C6))/IF(SUM(R25C6:R29C6)=0,1,SUM(R25C6:R29C6))"
    xlHoja1.Range("F24").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"

    xlHoja1.Range("C30").FormulaR1C1 = "=+((R31C3*R31C4)+(R32C3*R32C4)+(R33C3*R33C4)+(R34C3*R34C4)+(R35C3*R35C4))/IF(SUM(R31C4:R35C4)=0,1,SUM(R31C4:R35C4))"
    xlHoja1.Range("D30").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E30").FormulaR1C1 = "=+((R31C5*R31C6)+(R32C5*R32C6)+(R33C5*R33C6)+(R34C5*R34C6)+(R35C5*R35C6))/IF(SUM(R31C6:R35C6)=0,1,SUM(R31C6:R35C6))"
    xlHoja1.Range("F30").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"

    xlHoja1.Range("C36").FormulaR1C1 = "=+((R37C3*R37C4)+(R38C3*R38C4)+(R39C3*R39C4)+(R40C3*R40C4)+(R41C3*R41C4))/IF(SUM(R37C4:R41C4)=0,1,SUM(R37C4:R41C4))"
    xlHoja1.Range("D36").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    xlHoja1.Range("E36").FormulaR1C1 = "=+((R37C5*R37C6)+(R38C5*R38C6)+(R39C5*R39C6)+(R40C5*R40C6)+(R41C5*R41C6))/IF(SUM(R37C6:R41C6)=0,1,SUM(R37C6:R41C6))"
    xlHoja1.Range("F36").FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
    
'    xlHoja1.Range("C24:H24").Font.Bold = True
'    xlHoja1.Range("C30:H30").Font.Bold = True
'    xlHoja1.Range("C36:H36").Font.Bold = True
    
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    RSClose rs
End Sub

Private Sub Form_Load()
Set oCon = New DConecta
oCon.AbreConexion
CentraForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub



