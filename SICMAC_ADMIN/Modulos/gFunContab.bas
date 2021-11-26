Attribute VB_Name = "gFunContab"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A837285037A"
Option Base 0
Option Explicit

' Para declarar en MODULO */
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Sub CargaArchivo(lsArchivo As String, lsRutaArchivo As String)
    Dim x As Long
    Dim Temp As String
    Temp = GetActiveWindow()
    x = ShellExecute(Temp, "open", lsArchivo, "", lsRutaArchivo, 1)
    If x <= 32 Then
        If x = 2 Then
            MsgBox "No se encuentra el Archivo adjunto, " & vbCr & " verifique el servidor de archivos", vbInformation, " Aviso "
        ElseIf x = 8 Then
            MsgBox "Memoria insuficiente ", vbInformation, " Aviso "
        Else
            MsgBox "No se pudo abrir el Archivo adjunto", vbInformation, " Aviso "
        End If
    End If
  
End Sub

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Public Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional pbActivaHoja As Boolean = False)
Dim lbExisteHoja As Boolean
On Error Resume Next
    lbExisteHoja = False
activaHoja:
    For Each xlHoja1 In xlLibro.Worksheets
        If UCase(xlHoja1.Name) = UCase(psHojName) Then
            If Not pbActivaHoja Then
                SendKeys "{ENTER}"
                xlHoja1.Delete
            Else
                xlHoja1.Activate
                xlHoja1.Range("A1:BZ1").EntireColumn.Delete
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
    '*** PEAC 20120613
  'MsgBox TextErr(Err.Description), vbInformation, "Aviso"
  MsgBox TextErr(Err.Description + " o el archivo Excel ya fue ebierto."), vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Public Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
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
   MsgBox Err.Description, vbInformation, "Aviso"
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
Public Function LeeConstanteSist(psConst As ConstSistemas) As String
Dim oFun As NConstSistemas
Set oFun = New NConstSistemas
LeeConstanteSist = oFun.LeeConstSistema(psConst)
Set oFun = Nothing
End Function

