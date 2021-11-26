Attribute VB_Name = "CSimon"
Option Explicit

Public gcMovNro As String
Public gcDocTpo As String
Public gcPersona As String
Public gnSaldo As Currency

Public Sub FlexBackColor(Flex As MSHFlexGrid, pnFil As Integer, pnColor As Double)
    Dim K     As Integer
    Dim lnCol As Integer
    Dim lnFil As Integer
    lnCol = Flex.Col
    lnFil = Flex.Row
    Flex.Row = pnFil
    For K = 1 To Flex.Cols - 1
       Flex.Col = K
       Flex.CellBackColor = pnColor
    Next
    Flex.Row = lnFil
    Flex.Col = lnCol
End Sub

Public Sub KeyUp_Flex(Flex As MSHFlexGrid, KeyCode As Integer, Shift As Integer)
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
End Sub

Public Sub EliminaRow2(fg As MSHFlexGrid, nItem As Integer, Optional nPrimerRow As Integer = 1)
Dim nPos As Integer
nPos = nItem
If fg.Rows > nPrimerRow + 1 Then
   fg.RemoveItem nPos
Else
   For nPos = 0 To fg.Cols - 1
       fg.TextMatrix(nPrimerRow, nPos) = ""
   Next
End If
End Sub

