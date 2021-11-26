VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCOMExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Public Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional pbActivaHoja As Boolean = True)
Dim lbExisteHoja As Boolean
Dim lbBorrarRangos As Boolean
'On Error Resume Next
lbExisteHoja = False
lbBorrarRangos = False
activaHoja:
For Each xlHoja1 In xlLibro.Worksheets
    If UCase(xlHoja1.Name) = UCase(psHojName) Then
        If Not pbActivaHoja Then
            SendKeys "{ENTER}"
            xlHoja1.Delete
        Else
            xlHoja1.Activate
            If lbBorrarRangos Then xlHoja1.Range("A1:BZ1").EntireColumn.Delete
            lbExisteHoja = True
        End If
       Exit For
    End If
Next
If Not lbExisteHoja Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = psHojName
    If Err Then
        Err.Clear
        pbActivaHoja = True
        lbBorrarRangos = True
        GoTo activaHoja
    End If
End If
End Sub

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Public Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Dim cObj As FCOMObjetos
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  Err.Raise Err.Number, "Aviso", Err.Description
  ExcelBegin = False
End Function
'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Public Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
Dim cObj As New FCOMObjetos
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   Err.Raise Err.Number, "Aviso", Err.Description
End Sub

Public Function ExcelColumnaString(pnCol As Integer) As String
Dim sTexto As String
Dim nLetra As Integer
   If pnCol + 64 <= 90 Then
      sTexto = Chr(pnCol + 64)
   ElseIf pnCol + 64 <= 740 Then
      nLetra = Int((pnCol - 26) / 26) + IIf((pnCol - 26) Mod 26 = 0, 0, 1)
      sTexto = Chr(nLetra + 64) & Chr(((pnCol - 26) Mod (26 + IIf((pnCol - 26) Mod 26 = 0, 1, 0))) + IIf((pnCol - 26) Mod 26 = 0, nLetra, 1) + 63)
   End If
   ExcelColumnaString = sTexto
End Function

Public Sub ExcelCuadro(xlHoja1 As Excel.Worksheet, ByVal X1 As Currency, ByVal Y1 As Currency, ByVal X2 As Currency, ByVal Y2 As Currency, Optional lbLineasVert As Boolean = True, Optional lbLineasHoriz As Boolean = False)
xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).BorderAround xlContinuous, xlThin
If lbLineasVert Then
   If X2 <> X1 Then
     xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End If
End If
If lbLineasHoriz Then
    If Y1 <> Y2 Then
        xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End If
End If
End Sub

