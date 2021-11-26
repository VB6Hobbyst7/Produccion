Attribute VB_Name = "gFunFlex"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A80990902BF"
Option Base 0
Option Explicit

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
    fg.row = nItem
    fg.TopRow = nItem
    fg.TextMatrix(nItem, 0) = nItem
Exit Sub
AdicionaRowErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Public Sub BackColorFg(fg As MSHFlexGrid, psColor As String, Optional pbBold As Boolean = False)
Dim N As Integer
Dim nCol As Integer
nCol = fg.Col
For N = 1 To fg.Cols - 1
   fg.Col = N
   fg.CellBackColor = psColor
   fg.CellFontBold = pbBold
Next
fg.Col = nCol
End Sub

Public Sub EnumeraItems(fg As MSHFlexGrid, Optional nRowIni As Integer = 1)
Dim nPos
Dim nCont As Integer
nCont = 0
For nPos = nRowIni To fg.Rows - 1
    If nPos = nRowIni And fg.TextMatrix(nPos, 1) = "" Then
       Exit Sub
    End If
    nCont = nCont + 1
    fg.TextMatrix(nPos, 0) = nCont
Next
End Sub

'##ModelId=3A809921000F
Public Sub EliminaRow(fg As MSHFlexGrid, nItem As Integer, Optional nRowIni As Integer = 1)
    On Error GoTo EliminaRowErr
    Dim nPos As Integer
    
    nPos = nItem
    If fg.Rows > nRowIni + 1 Then
        fg.RemoveItem nPos
    Else
        For nPos = 0 To fg.Cols - 1
            fg.TextMatrix(nRowIni, nPos) = ""
        Next
    End If
    EnumeraItems fg, nRowIni
    Exit Sub
EliminaRowErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
'##ModelId=3A80B11F00AB
Public Sub EnfocaTexto(txtCelda As Variant, KeyAscii As Integer, fg As MSHFlexGrid, Optional pnTopIni As Integer = 0, Optional pnLeftIni As Integer = 0)
Dim nx, ny As Integer
   On Error GoTo EnfocaTextoErr
   txtCelda.Text = ""
   nx = fg.Left + fg.ColPos(fg.Col) + 10 + pnLeftIni
   ny = fg.Top + fg.RowPos(fg.row) + 10 + pnTopIni
   txtCelda.Height = fg.CellHeight
   txtCelda.Width = fg.CellWidth + 30
   
   txtCelda.Tag = fg.row
   txtCelda.Visible = True
   If KeyAscii = 0 Then
      txtCelda.Text = fg.Text
      txtCelda.SelStart = 0
      txtCelda.SelLength = Len(txtCelda.Text)
   Else
      txtCelda.Text = Chr(KeyAscii)
      txtCelda.SelStart = 1
   End If
   txtCelda.Left = nx
   txtCelda.Top = ny
   
   txtCelda.SetFocus
Exit Sub
EnfocaTextoErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
'##ModelId=3A80B1670271
Public Sub Flex_PresionaKey(Flex As MSHFlexGrid, KeyCode As Integer, Shift As Integer)
   On Error GoTo PresionaKeyErr
    Select Case KeyCode
        Case vbKeyC And Shift = 2   '   Copiar  [Ctrl+C]
            Clipboard.Clear
            Clipboard.SetText Flex.Text
            KeyCode = 0
        Case vbKeyV And Shift = 2   '   Pegar  [Ctrl+V]
            Flex.Text = Clipboard.GetText
            KeyCode = 0
        Case vbKeyX And Shift = 2   '   Cortar  [Ctrl+X]
            Clipboard.Clear
            Clipboard.SetText Flex.Text
            Flex.Text = ""
            KeyCode = 0
        Case vbKeyDelete            '   Borrar [Delete]
            Flex.Text = ""
            KeyCode = 0
    End Select
   Exit Sub
PresionaKeyErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub


'##ModelId=3A9422530157
Public Sub CalculaTotales()
    On Error GoTo CalculaTotalesErr

    'your code goes here...

    Exit Sub
CalculaTotalesErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
Public Sub MSHFlex(ByRef pFlexGrid As Object, ByVal pCol As Integer, _
ByVal pEncabezado As String, ByVal pAnchoCol As String, ByVal pAlineaCol As String)
Dim X As Integer, vPos As Integer
Dim vAli As String
With pFlexGrid
    .Cols = pCol
    .Clear
    For X = .Rows To 1 Step -1
        If X <= 2 Then Exit For
        .RemoveItem (X - 1)
    Next X
    
    For X = 0 To pCol - 1
        vPos = InStr(1, pEncabezado, "-", vbTextCompare)
        .TextMatrix(0, X) = Mid(pEncabezado, 1, IIf(vPos > 0, vPos - 1, Len(pEncabezado)))
        pEncabezado = Mid(pEncabezado, IIf(vPos > 0, vPos + 1, Len(pEncabezado)))
    Next X
    
    For X = 0 To pCol - 1
        vPos = InStr(1, pAnchoCol, "-", vbTextCompare)
        .ColWidth(X) = Val(Mid(pAnchoCol, 1, IIf(vPos > 0, vPos - 1, Len(pAnchoCol))))
        pAnchoCol = Mid(pAnchoCol, IIf(vPos > 0, vPos + 1, Len(pAnchoCol)))
    Next X
    
    For X = 0 To pCol - 1
        .ColAlignmentFixed(X) = 4
        vPos = InStr(1, pAlineaCol, "-", vbTextCompare)
        vAli = UCase(Mid(pAlineaCol, 1, IIf(vPos > 0, vPos - 1, Len(pAlineaCol))))
        If Len(vAli) > 0 And (vAli = "L" Or vAli = "R" Or vAli = "C") Then
            .ColAlignment(X) = Switch(vAli = "L", 1, vAli = "R", 7, vAli = "C", 4)
            pAlineaCol = Mid(pAlineaCol, IIf(vPos > 0, vPos + 1, Len(pAlineaCol)))
        Else
            .ColAlignment(X) = 4
        End If
    Next X
    .row = 1
    .Col = 1
    .RowHeight(-1) = 285
End With
End Sub

Public Sub MoveFlex(fg As MSHFlexGrid, KeyCode As Integer)
If fg.Rows > 2 Then
   If KeyCode = 38 Then
      If fg.row > 1 Then
         fg.row = fg.row - 1
         If fg.TopRow > fg.row Then
            fg.TopRow = fg.row
         End If
      End If
   Else
      If fg.row < fg.Rows - 1 Then
         fg.row = fg.row + 1
         If fg.TopRow + (fg.Height / fg.RowHeight(fg.row)) < fg.row Then
            fg.TopRow = fg.TopRow + 1
         End If
      End If
   End If
End If
End Sub
'EJVG20131118 ***
Public Sub FormateaFlex(ByVal pflex As FlexEdit)
    pflex.Clear
    pflex.FormaCabecera
    pflex.Rows = 2
End Sub
'END EJVG *******
