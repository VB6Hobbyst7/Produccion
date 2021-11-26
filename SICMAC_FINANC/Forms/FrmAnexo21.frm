VERSION 5.00
Begin VB.Form FrmAnexo21 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmAnexo21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim R As New ADODB.Recordset
Dim oCon As DConecta

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim oBarra As clsProgressBar
Public Sub GeneraSUCAVEAnx21Soles(pfecha As Date)
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New scripting.FileSystemObject
Dim nFil As Integer
Dim nCol As Integer
Dim pdFecha As String
Dim i As Integer
Dim J As Integer
Dim m As Integer
Dim sCad As String
Dim Fecha As Date


On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Format(pfecha, "MM") & "01MN.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    
    psArchivoAGrabar = App.path & "\SPOOLER\01" & Format(pfecha, "YYMMdd") & ".121"
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlHoja In xlLibro.Worksheets
        If UCase(xlHoja.Name) = UCase("Anx1") Then
            bEncontrado = True
            xlHoja.Activate
            Exit For
        End If
    Next

    If bEncontrado = False Then
        MsgBox "No existen la hoja Anx_1", vbExclamation, "Aviso!!!"
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
        Exit Sub
    End If
    '''''''''''''''''''''''''''''
    
    Dim matArreglo(1 To 33, 1 To 21) As Currency
    Dim matCodigo(1 To 33) As String * 5
    Dim nTemp1 As Integer
    Dim nTemp2 As Integer
    Dim EncajeExigible
    Dim nDias As Integer
    nDias = Format(pfecha, "DD")
    For i = 1 To nDias
        matCodigo(i) = CStr(i)
    Next i
    matCodigo(32) = "100"
    matCodigo(33) = "200"
    

For i = 16 To 16 + nDias - 1
        m = 0
        nCol = i - 15
        For J = Asc("B") To Asc("V")
            m = m + 1
            matArreglo(nCol, m) = xlHoja.Range(Chr(J) & i)
        Next J
Next i

'matCodigo (32)
'matCodigo (33)
m = 0
For J = Asc("B") To Asc("V")
    m = m + 1
    matArreglo(32, m) = xlHoja.Range(Chr(J) & i)
    matArreglo(33, m) = xlHoja.Range(Chr(J) & i) / nDias
Next J

'Creacion del Archivo
Open psArchivoAGrabar For Output As #1
'0121010011220040630012
Print #1, "01210100109" & Format(pfecha, "YYYYMMDD") & "012" & Space(15) & LlenaCerosSUCAVE(xlHoja.Range("D" & i + 4)) & LlenaCerosSUCAVE(xlHoja.Range("D" & i + 6))
sCad = ""
For i = 1 To nDias
    sCad = ""
    For J = 1 To 21
        sCad = sCad & LlenaCerosSUCAVE(matArreglo(i, J))
    Next J
    'Print #1, IIf(matCodigo(i) < 100, Space(2), Space(1)) & IIf(matCodigo(i) < 10, matCodigo(i) & Space(1), matCodigo(i)) & sCad
    Print #1, Trim(matCodigo(i)) & Space(4 - Len(Trim(matCodigo(i)))) & sCad
Next i

For i = 32 To 33
    sCad = ""
    For J = 1 To 21
        sCad = sCad & LlenaCerosSUCAVE(matArreglo(i, J))
    Next J
    Print #1, Trim(matCodigo(i)) & Space(4 - Len(Trim(matCodigo(i)))) & sCad
Next i

Close #1

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
MsgBox "Reporte SUCAVE Generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub
ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Public Sub GeneraSUCAVEAnx21Dolares(pfecha As Date)
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New scripting.FileSystemObject
Dim nFil As Integer
Dim nCol As Integer
Dim pdFecha As String
Dim i As Integer
Dim J As Integer
Dim m As Integer
Dim sCad As String
Dim Fecha As Date


On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Format(pfecha, "MM") & "01ME.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    
    psArchivoAGrabar = App.path & "\SPOOLER\02" & Format(pfecha, "YYMMdd") & ".121"
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlHoja In xlLibro.Worksheets
        If UCase(xlHoja.Name) = UCase("Anx1") Then
            bEncontrado = True
            xlHoja.Activate
            Exit For
        End If
    Next

    If bEncontrado = False Then
        MsgBox "No existen la hoja anexo01", vbExclamation, "Aviso!!!"
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
        Exit Sub
    End If
    '''''''''''''''''''''''''''''
    
    Dim matArreglo(1 To 33, 1 To 18) As Currency
    Dim matCodigo(1 To 33) As String * 5
    Dim nTemp1 As Integer
    Dim nTemp2 As Integer
    Dim EncajeExigible
    Dim nDias As Integer
    nDias = Format(pfecha, "DD")
    For i = 1 To nDias
        matCodigo(i) = CStr(i)
    Next i
    matCodigo(32) = "100"
    matCodigo(33) = "200"
    

For i = 16 To 16 + nDias - 1
        m = 0
        nCol = i - 15
        For J = Asc("B") To Asc("S")
            m = m + 1
            matArreglo(nCol, m) = xlHoja.Range(Chr(J) & i)
        Next J
Next i

'matCodigo (32)
'matCodigo (33)
m = 0
For J = Asc("B") To Asc("S")
    m = m + 1
    matArreglo(32, m) = xlHoja.Range(Chr(J) & i)
    matArreglo(33, m) = xlHoja.Range(Chr(J) & i) / nDias
Next J

'Creacion del Archivo
Open psArchivoAGrabar For Output As #1
'0121020011220040531012
Print #1, "01210200109" & Format(pfecha, "YYYYMMDD") & "012" & Space(15) & LlenaCerosSUCAVE(xlHoja.Range("C" & i + 4)) & LlenaCerosSUCAVE(xlHoja.Range("C" & i + 6))
sCad = ""
For i = 1 To nDias
    sCad = ""
    For J = 1 To 18
        sCad = sCad & LlenaCerosSUCAVE(matArreglo(i, J))
    Next J
    Print #1, IIf(matCodigo(i) < 100, IIf(matCodigo(i) < 10, "   ", "  "), " ") & Trim(CStr(matCodigo(i))) & sCad
Next i

For i = 32 To 33
    sCad = ""
    For J = 1 To 18
        sCad = sCad & LlenaCerosSUCAVE(matArreglo(i, J))
    Next J
    Print #1, IIf(matCodigo(i) < 100, "  ", " ") & Trim(CStr(matCodigo(i))) & sCad
Next i

Close #1

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
MsgBox "Reporte SUCAVE Generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub
ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub







