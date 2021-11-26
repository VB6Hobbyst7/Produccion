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
    fg.row = nItem
    fg.TopRow = nItem
    fg.TextMatrix(nItem, 0) = nItem
Exit Sub
AdicionaRowErr:
    Call RaiseError(MyUnhandledError, "iFunFlex:AdicionaRow Method")
End Sub
Public Sub EnumeraItems(fg As MSHFlexGrid)
Dim nPos
For nPos = 1 To fg.Rows - 1
    If nPos = 1 And fg.TextMatrix(nPos, 1) = "" Then
       Exit Sub
    End If
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
Public Sub EnfocaTexto(txtCelda As Variant, KeyAscii As Integer, fg As MSHFlexGrid, Optional pnTopIni As Integer = 0, Optional pnLeftIni As Integer = 0)
Dim nX, nY As Integer
Dim R As Control
   'On Error GoTo EnfocaTextoErr
   
   If txtCelda.Name = "txtFecha" Then
      txtCelda.Text = "  /  /    "
   Else
      txtCelda.Text = ""
   End If
   nX = fg.Left + fg.ColPos(fg.Col) + 10 + pnLeftIni
   nY = fg.Top + fg.RowPos(fg.row) + 10 + pnTopIni
   If txtCelda.Name = "txtFecha" Then
        txtCelda.Height = fg.CellHeight + 30
   Else
        txtCelda.Height = fg.CellHeight + 30
   End If
   txtCelda.Width = fg.CellWidth + 30
   txtCelda.Tag = fg.row
   txtCelda.Visible = True
   If txtCelda.Name <> "txtFecha" Then
        Select Case fg.ColAlignment(fg.Col)
            Case 0, 1, 2
                txtCelda.Alignment = 0
            Case 6, 7, 8
                txtCelda.Alignment = 1
            Case 3, 4, 5
                txtCelda.Alignment = 2
        End Select
   End If
   If KeyAscii = 0 Then
        If txtCelda.Name = "txtFecha" And fg.Text = "" Then
             txtCelda.Text = "  /  /    "
        Else
            If IsDate(fg.Text) = False And txtCelda.Name = "txtFecha" Then
                txtCelda.Text = "  /  /    "
            Else
                txtCelda.Text = fg.Text
            End If
        End If
        txtCelda.SelStart = 0
        txtCelda.SelLength = Len(txtCelda.Text)
   Else
        If txtCelda.Name = "txtFecha" Then
            If IsNumeric(Chr(KeyAscii)) Then
                If IsNumeric(Chr(KeyAscii)) Then
                    txtCelda.Text = Chr(KeyAscii) & " /  /    "
                End If
            End If
        Else
            txtCelda.Text = Chr(KeyAscii)
        End If
        txtCelda.SelStart = 1
   End If
   txtCelda.Left = nX
   txtCelda.Top = nY
   
   txtCelda.SetFocus
Exit Sub
EnfocaTextoErr:
   MsgBox err.Description, vbInformation, "¡Aviso!"
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
   MsgBox err.Description, vbInformation, "¡Aviso!"
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
        .ColWidth(x) = val(Mid(pAnchoCol, 1, IIf(vPos > 0, vPos - 1, Len(pAnchoCol))))
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
    .row = 1
    .Col = 1
    .RowHeight(-1) = 285
End With
End Sub

Public Sub BackColorFg(fg As MSHFlexGrid, psColor As String, Optional pbBold As Boolean = False)
Dim n As Integer
Dim nCol As Integer
nCol = fg.Col
For n = 1 To fg.Cols - 1
   fg.Col = n
   fg.CellBackColor = psColor
   fg.CellFontBold = pbBold
Next
fg.Col = nCol
End Sub
'WIOR 20120709 SEGUN OYP-RFC090-2012**************************************************
Public Function SumarCampo(ByVal pFE As FlexEdit, ByVal pnCampo As Integer) As Double
On Error GoTo ErrorSumarCampo
    Dim nTotal As Double
    Dim nConteo As Integer
    nTotal = 0
    If pFE.Rows - 1 > 0 Then
        For nConteo = 1 To pFE.Rows - 1
            If pFE.TextMatrix(nConteo, pnCampo) <> "0.00" Then
                nTotal = nTotal + CDbl(IIf(pFE.TextMatrix(nConteo, pnCampo) = "", 0, pFE.TextMatrix(nConteo, pnCampo)))
            End If
        Next
    End If

    SumarCampo = nTotal
    
    Exit Function
ErrorSumarCampo:
    MsgBox err.Description, vbCritical, "Error"
End Function
'WIOR FIN *****************************************************************************

Public Function ValidaDatosGrid(ByVal Flex As FlexEdit, ByVal psMsjVacio As String, ByVal psMsjFaltaDatos As String, ByVal pnCantColumn As Integer) As Boolean
    Dim i As Integer, J As Integer
    ValidaDatosGrid = False
    If Flex.Rows - 1 = 1 And Flex.TextMatrix(1, 0) = "" Then
        MsgBox psMsjVacio, vbInformation, "Aviso"
        ValidaDatosGrid = False
        Exit Function
    End If
    
    For i = 1 To Flex.Rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            For J = 1 To pnCantColumn - 1
                If Trim(Flex.TextMatrix(i, J)) = "" Then
                    MsgBox psMsjFaltaDatos & " fila " & J, vbInformation, "Aviso"
                    ValidaDatosGrid = False
                    Exit Function
                End If
            Next J
        End If
    Next i
    ValidaDatosGrid = True
End Function
'WIOR 20130921 **********************************************************
Public Function ValidaFlex(ByVal Flex As FlexEdit, ByVal pnCol As Long) As Boolean
Dim sColumnas() As String
ValidaFlex = True
sColumnas = Split(Flex.ColumnasAEditar, "-")
If sColumnas(pnCol) = "X" Then
   ValidaFlex = False
   SendKeys "{Tab}", True
   Exit Function
End If
End Function
'WIOR FIN **********************************************************

'EJVG20131213 ***
Public Sub FormateaFlex(ByVal pflex As FlexEdit)
    pflex.Clear
    pflex.FormaCabecera
    pflex.Rows = 2
End Sub
'END EJVG *******
