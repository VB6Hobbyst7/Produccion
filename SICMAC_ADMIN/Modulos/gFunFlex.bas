Attribute VB_Name = "gFunFlex"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A80990902BF"
Option Base 0
Option Explicit

'##ModelId=3A80991402FD
Public Sub FormatoFlex()
    On Error GoTo FormatoFlexErr

    'your code goes here...

    Exit Sub
FormatoFlexErr:
    Call RaiseError(MyUnhandledError, "iFunFlex:FormatoFlex Method")
End Sub

'##ModelId=3A80991903D8
Public Sub AdicionaRow(fg As MSHFlexGrid, Optional nItem As Integer = 0)
    On Error GoTo AdicionaRowErr

    If nItem = 0 Then
         nItem = fg.Rows
    Else
    If nItem > fg.Rows Then
        Exit Sub
    End If
    End If
    If fg.Rows = 2 And Len(Trim(fg.TextMatrix(1, 1)) & "") = 0 Then
         nItem = 1
    End If
    If nItem > 1 Then
        fg.AddItem " ", nItem
    End If
    fg.RowHeight(nItem) = 285
    fg.Col = 1
    fg.Row = nItem
    fg.TopRow = nItem
    fg.TextMatrix(nItem, 0) = nItem
Exit Sub
AdicionaRowErr:
    Call RaiseError(MyUnhandledError, "iFunFlex:AdicionaRow Method")
End Sub
Public Sub EnumeraItems(fg As MSHFlexGrid)
Dim nPos
For nPos = 1 To fg.Rows - 1
    fg.TextMatrix(nPos, 0) = nPos
Next
End Sub

'##ModelId=3A809921000F
Public Sub EliminaRow(fg As MSHFlexGrid, nItem As Integer)
    On Error GoTo EliminaRowErr
    Dim nPos As Integer
    
    nPos = nItem
    If fg.Rows > 2 Then
        fg.RemoveItem nPos
    Else
        For nPos = 0 To fg.Cols - 1
            fg.TextMatrix(1, nPos) = ""
        Next
    End If
    EnumeraItems fg
    Exit Sub
EliminaRowErr:
    Call RaiseError(MyUnhandledError, "iFunFlex:EliminaRow Method")
End Sub
'##ModelId=3A80B11F00AB
 
 Public Sub EnfocaTexto(txtCelda As Variant, KeyAscii As Integer, fg As MSHFlexGrid, Optional pnTopIni As Integer = 0, Optional pnLeftIni As Integer = 0, Optional psTipoControl As String = "T", Optional pnMasHeigth As Integer = 0)
Dim nx, ny As Integer
Dim R As Control
   On Error GoTo EnfocaTextoErr
   If InStr("13-8-27", Format(KeyAscii, "0")) > 0 Then
      KeyAscii = 0
   End If
   nx = fg.Left + fg.ColPos(fg.Col) + pnLeftIni
   ny = fg.Top + fg.RowPos(fg.Row) + pnTopIni
   txtCelda.Height = fg.CellHeight + 30 + pnMasHeigth
   txtCelda.Width = fg.CellWidth + 30
   txtCelda.Tag = fg.Row
   txtCelda.Visible = True
   If KeyAscii = 0 Then
      txtCelda.Text = fg.Text
      txtCelda.SelStart = 0
      txtCelda.SelLength = Len(txtCelda.Text)
   Else
      If psTipoControl = "T" Then
         txtCelda.Text = Chr(KeyAscii)
      ElseIf psTipoControl = "F" Then
         txtCelda.Text = Chr(KeyAscii) + " /  /    "
      End If
      txtCelda.SelStart = 1
   End If
   txtCelda.Left = nx
   txtCelda.Top = ny
   
   txtCelda.SetFocus
Exit Sub
EnfocaTextoErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

'##ModelId=3A80B1670271
Public Sub PresionaKey(flex As MSHFlexGrid, KeyCode As Integer, Shift As Integer)
    On Error GoTo PresionaKeyErr
    Select Case KeyCode
        Case vbKeyC And Shift = 2   '   Copiar  [Ctrl+C]
            Clipboard.Clear
            Clipboard.SetText flex.Text
            KeyCode = 0
        Case vbKeyV And Shift = 2   '   Pegar  [Ctrl+V]
            flex.Text = Clipboard.GetText
            KeyCode = 0
        Case vbKeyX And Shift = 2   '   Cortar  [Ctrl+X]
            Clipboard.Clear
            Clipboard.SetText flex.Text
            flex.Text = ""
            KeyCode = 0
        Case vbKeyDelete            '   Borrar [Delete]
            flex.Text = ""
            KeyCode = 0
    End Select

Exit Sub
PresionaKeyErr:
    Call RaiseError(MyUnhandledError, "iFunFlex:PresionaKey Method")
End Sub

'##ModelId=3A9422530157
Public Sub CalculaTotales()
    On Error GoTo CalculaTotalesErr

    'your code goes here...

    Exit Sub
CalculaTotalesErr:
    Call RaiseError(MyUnhandledError, "iFunFlex:CalculaTotales Method")
End Sub
Public Sub MSHFlex(ByRef pFlexGrid As MSHFlexGrid, ByVal pCol As Integer, _
ByVal pEncabezado As String, ByVal pAnchoCol As String, ByVal pAlineaCol As String)
Dim x As Integer, vPos As Integer
Dim vAli As String
With pFlexGrid
    .Cols = pCol
    .Clear
    For x = .Rows To 1 Step -1
        If x <= 2 Then Exit For
        .RemoveItem (x - 1)
    Next x
    
    For x = 0 To pCol - 1
        vPos = InStr(1, pEncabezado, "-", vbTextCompare)
        .TextMatrix(0, x) = Mid(pEncabezado, 1, IIf(vPos > 0, vPos - 1, Len(pEncabezado)))
        pEncabezado = Mid(pEncabezado, IIf(vPos > 0, vPos + 1, Len(pEncabezado)))
    Next x
    
    For x = 0 To pCol - 1
        vPos = InStr(1, pAnchoCol, "-", vbTextCompare)
        .ColWidth(x) = Val(Mid(pAnchoCol, 1, IIf(vPos > 0, vPos - 1, Len(pAnchoCol))))
        pAnchoCol = Mid(pAnchoCol, IIf(vPos > 0, vPos + 1, Len(pAnchoCol)))
    Next x
    
    For x = 0 To pCol - 1
        .ColAlignmentFixed(x) = 4
        vPos = InStr(1, pAlineaCol, "-", vbTextCompare)
        vAli = UCase(Mid(pAlineaCol, 1, IIf(vPos > 0, vPos - 1, Len(pAlineaCol))))
        If Len(vAli) > 0 And (vAli = "L" Or vAli = "R" Or vAli = "C") Then
            .ColAlignment(x) = Switch(vAli = "L", 1, vAli = "R", 7, vAli = "C", 4)
            pAlineaCol = Mid(pAlineaCol, IIf(vPos > 0, vPos + 1, Len(pAlineaCol)))
        Else
            .ColAlignment(x) = 4
        End If
    Next x
    .Row = 1
    .Col = 1
    .RowHeight(-1) = 285
End With
End Sub
'EJVG20120914 ***
Public Sub FormateaFlex(ByVal pflex As FlexEdit)
    pflex.Clear
    pflex.FormaCabecera
    pflex.Rows = 2
End Sub
Public Function FlexVacio(ByVal pflex As FlexEdit) As Boolean
    If (pflex.Rows - 1 = 1 And pflex.TextMatrix(1, 0) = "") Then
        FlexVacio = True
    Else
        FlexVacio = False
    End If
End Function
'END EJVG
