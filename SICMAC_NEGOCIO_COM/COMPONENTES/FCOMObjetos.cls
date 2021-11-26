VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCOMObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub CentraForm(frmCentra As Object)
    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2, frmCentra.Width, frmCentra.Height
End Sub
'##ModelId=3A7EDEA302BF
Public Sub CargaCombo(ByRef Combo As Object, rs As ADODB.Recordset)
Dim Campo As ADODB.Field
Dim lsDato As String
If rs Is Nothing Then Exit Sub
Combo.Clear
Do While Not rs.EOF
    lsDato = ""
    For Each Campo In rs.Fields
        lsDato = lsDato & Campo.Value & Space(50)
    Next
    lsDato = Mid(lsDato, 1, Len(lsDato) - 50)
    Combo.AddItem lsDato
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
Public Sub fEnfoque(ctrControl As Object)
    ctrControl.SelStart = 0
    ctrControl.SelLength = Len(ctrControl.Text)
End Sub
Public Function BuscaDato(ByVal Criterio As String, rsAdo As ADODB.Recordset, ByVal start As Long, ByVal lMsg As Boolean) As Boolean
Dim Pos As Variant
On Error GoTo Errbusq
   BuscaDato = False
   Pos = rsAdo.Bookmark
   rsAdo.Find Criterio, IIf(start = 1, 0, start + 1), adSearchForward, 1
   If rsAdo.EOF Then
      rsAdo.Bookmark = Pos
      If lMsg Then
         MsgBox " ! Dato no encontrado... ! ", vbExclamation, "Error de Busqueda"
         BuscaDato = False
      End If
   Else
      BuscaDato = True
   End If
Exit Function
Errbusq:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Function

Public Sub RSClose(rs As ADODB.Recordset)
If Not rs Is Nothing Then
   If rs.State = adStateOpen Then
       rs.Close
       Set rs = Nothing
   End If
End If
End Sub
Public Function RSVacio(rs1 As ADODB.Recordset) As Boolean
 RSVacio = (rs1.BOF And rs1.EOF)
End Function
Public Function TextErr(sMsg As String) As String
Dim nLen As Integer
nLen = InStr(1, sMsg, "*", vbTextCompare)
TextErr = Mid(sMsg, nLen + 1, Len(sMsg))
End Function

Public Sub CargaArchivo(lsArchivo As String, lsRutaArchivo As String)
    Dim X As Long
    Dim Temp As String
    Temp = GetActiveWindow()
    X = ShellExecute(Temp, "open", lsArchivo, "", lsRutaArchivo, 1)
    If X <= 32 Then
        If X = 2 Then
            MsgBox "No se encuentra el Archivo adjunto, " & vbCr & " verifique el servidor de archivos", vbInformation, " Aviso "
        ElseIf X = 8 Then
            MsgBox "Memoria insuficiente ", vbInformation, " Aviso "
        Else
            MsgBox "No se pudo abrir el Archivo adjunto", vbInformation, " Aviso "
        End If
    End If
  
End Sub

Public Sub CambiaTamañoCombo(ByRef cboCombo As Object, Optional nTamaño As Long = 200)
Dim oVarPublica As New FCOMVarPublicas
    SendMessage cboCombo.hwnd, oVarPublica.CB_SETDROPPEDWIDTH, nTamaño, 0
End Sub

Public Sub RSLlenaCombo(prs As Recordset, psCombo As Object)
If Not prs Is Nothing Then
   If Not prs.EOF Then
      psCombo.Clear
      Do While Not prs.EOF
         psCombo.AddItem Trim(prs(1)) & Space(100) & Trim(prs(0))
         prs.MoveNext
      Loop
   End If
End If
End Sub

Public Function AdicionaRecordSet(ByRef prsDat As ADODB.Recordset, ByVal prs As ADODB.Recordset)
Dim nCol As Integer
Do While Not prs.EOF
    If Not prsDat Is Nothing Then
        If prsDat.State = adStateClosed Then
            For nCol = 0 To prs.Fields.Count - 1
                With prs.Fields(nCol)
                    prsDat.Fields.Append .Name, .Type, .DefinedSize, .Attributes
                End With
            Next
            prsDat.Open
        End If
        prsDat.AddNew
        For nCol = 0 To prs.Fields.Count - 1
            prsDat.Fields(nCol).Value = prs.Fields(nCol).Value
        Next
        prsDat.Update
    End If
    prs.MoveNext
Loop
If Not prsDat Is Nothing Then
    If prsDat.RecordCount > 0 Then
        prsDat.MoveFirst
    End If
End If
End Function

Public Function RecordSetAdiciona(ByRef prsDat As ADODB.Recordset, ByVal prs As ADODB.Recordset)
Dim nCol As Integer
RecordSetDefineCampos prsDat, prs
Do While Not prs.EOF
    prsDat.AddNew
    For nCol = 0 To prs.Fields.Count - 1
        prsDat.Fields(nCol).Value = prs.Fields(nCol).Value
    Next
    prsDat.Update
    prs.MoveNext
Loop
End Function


Public Function RecordSetDefineCampos(ByRef prsDat As ADODB.Recordset, ByVal prs As ADODB.Recordset)
Dim nCol As Integer

If prsDat Is Nothing Then
    Set prsDat = New ADODB.Recordset
End If
If prsDat.State = adStateClosed Then
    For nCol = 0 To prs.Fields.Count - 1
        With prs.Fields(nCol)
            prsDat.Fields.Append .Name, .Type, .DefinedSize, .Attributes
        End With
    Next
    prsDat.Open
End If
End Function

Public Sub UbicaCombo(pCombo As Object, psDato As String, Optional pbBuscaFinal As Boolean = True, Optional pnNumComp As Integer = 7)
    Dim i As Integer
    Dim lbBan As Boolean
    lbBan = False
    
    If pbBuscaFinal Then
        For i = 0 To pCombo.ListCount - 1
            If Trim(Right(pCombo.List(i), pnNumComp)) = Trim(Right(psDato, pnNumComp)) Then
                lbBan = True
                pCombo.ListIndex = i
                i = pCombo.ListCount
            End If
        Next i
    Else
        For i = 0 To pCombo.ListCount - 1
            If Trim(Left(pCombo.List(i), pnNumComp)) = Trim(Left(psDato, pnNumComp)) Then
                lbBan = True
                pCombo.ListIndex = i
                i = pCombo.ListCount
            End If
        Next i
    End If
    
    If Not lbBan Then pCombo.ListIndex = -1
End Sub

