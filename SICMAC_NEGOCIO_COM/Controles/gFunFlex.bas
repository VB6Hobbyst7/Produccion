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
Dim nx, ny As Integer
Dim r As Control
   'On Error GoTo EnfocaTextoErr
   
   If txtCelda.Name = "txtFecha" Then
      txtCelda.Text = "  /  /    "
   Else
      txtCelda.Text = ""
   End If
   nx = fg.Left + fg.ColPos(fg.Col) + 10 + pnLeftIni
   ny = fg.Top + fg.RowPos(fg.Row) + 10 + pnTopIni
   If txtCelda.Name = "txtFecha" Then
        txtCelda.Height = fg.CellHeight + 30
   Else
        txtCelda.Height = fg.CellHeight + 30
   End If
   txtCelda.Width = fg.CellWidth + 30
   txtCelda.Tag = fg.Row
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
   txtCelda.Left = nx
   txtCelda.Top = ny
   
   txtCelda.SetFocus
Exit Sub
EnfocaTextoErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
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
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub


'##ModelId=3A941B9A0251
Public Sub LimpiaFlex()
    On Error GoTo LimpiaFlexErr

    'your code goes here...

    Exit Sub
LimpiaFlexErr:
    Call RaiseError(MyUnhandledError, "iFunFlex:LimpiaFlex Method")
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


Public Function Letras(intTecla As Integer, Optional lbMayusculas As Boolean = True) As Integer
If lbMayusculas Then
    Letras = Asc(UCase(Chr(intTecla)))
Else
    Letras = Asc(LCase(Chr(intTecla)))
End If
End Function
Public Function SoloLetras(intTecla As Integer) As Integer
Dim cValidar As String
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$%^&()*"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    SoloLetras = intTecla
End Function
Public Function NumerosDecimales(cTexto As TextBox, intTecla As Integer, _
    Optional nLongitud As Integer = 8, Optional nDecimal As Integer = 2) As Integer
    Dim cValidar As String
    Dim cCadena As String
    cCadena = cTexto
    cValidar = "-0123456789."
    
    If InStr(".", Chr(intTecla)) <> 0 Then
        If InStr(cCadena, ".") <> 0 Then
            intTecla = 0
            Beep
        ElseIf intTecla > 26 Then
            If InStr(cValidar, Chr(intTecla)) = 0 Then
                intTecla = 0
                Beep
            End If
        End If
    ElseIf intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    Dim vPosCur As Byte
    Dim vPosPto As Byte
    Dim vNumDec As Byte
    Dim vNumLon As Byte
    
    vPosPto = InStr(cTexto.Text, ".")
    vPosCur = cTexto.SelStart
    vNumLon = Len(cTexto)
    If vPosPto > 0 Then
        vNumDec = Len(Mid(cTexto, vPosPto + 1))
    End If
    If vPosPto > 0 Then
        If cTexto.SelLength <> Len(cTexto) Then
        If ((vNumDec >= nDecimal And cTexto.SelStart >= vPosPto) Or _
        (vNumLon >= nLongitud)) _
        And intTecla <> vbKeyBack And intTecla <> vbKeyDecimal And intTecla <> vbKeyReturn Then
            intTecla = 0
            Beep
        End If
        End If
    Else
        If vNumLon >= nLongitud And intTecla <> vbKeyBack _
        And intTecla <> vbKeyReturn Then
            intTecla = 0
            Beep
        End If
        If (vNumLon - cTexto.SelStart) > nDecimal And intTecla = 46 Then
            intTecla = 0
            Beep
        End If
    End If
    NumerosDecimales = intTecla
End Function
Public Function NumerosEnteros(intTecla As Integer, Optional pbNegativos As Boolean = False) As Integer
Dim cValidar As String
    If pbNegativos = False Then
        cValidar = "0123456789"
    Else
        cValidar = "0123456789-"
    End If
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    NumerosEnteros = intTecla
End Function

Public Sub fEnfoque(ctrControl As Control)
    ctrControl.SelStart = 0
    ctrControl.SelLength = Len(ctrControl.Text)
End Sub
Public Function ValFecha(lsControl As Control) As Boolean
   If Mid(lsControl, 1, 2) > 0 And Mid(lsControl, 1, 2) <= 31 Then
        If Mid(lsControl, 4, 2) > 0 And Mid(lsControl, 4, 2) <= 12 Then
            If Mid(lsControl, 7, 4) >= 1900 And Mid(lsControl, 7, 4) <= 9999 Then
               If IsDate(lsControl) = False Then
                    ValFecha = False
                    MsgBox "Formato de fecha no es válido", vbInformation, "Aviso"
                    lsControl.SetFocus
                    Exit Function
               Else
                    ValFecha = True
               End If
            Else
                ValFecha = False
                MsgBox "Año de Fecha no es válido", vbInformation, "Aviso"
                lsControl.SetFocus
                lsControl.SelStart = 6
                lsControl.SelLength = 4
                Exit Function
            End If
        Else
            ValFecha = False
            MsgBox "Mes de Fecha no es válido", vbInformation, "Aviso"
            lsControl.SetFocus
            lsControl.SelStart = 3
            lsControl.SelLength = 2
            Exit Function
        End If
    Else
        ValFecha = False
        MsgBox "Dia de Fecha no es válido", vbInformation, "Aviso"
        lsControl.SetFocus
        lsControl.SelStart = 0
        lsControl.SelLength = 2
        Exit Function
    End If
End Function

