Attribute VB_Name = "JPNZ"
Option Explicit

Public gcOpeCod As String
Public gcOpeDesc As String

Public Sub CargaComboLog(rs As ADODB.Recordset, pCombo As ComboBox, Optional pnEspacio As Long = 100, Optional pnColumna0 As Integer = 0, Optional pnColumna1 As Integer = 1, Optional pnFiltro As Integer = -1)
    pCombo.Clear
    If Not rs.EOF Then
        If pnFiltro = -1 Then
            If rs.Fields.Count = 1 Then
                While Not rs.EOF
                    pCombo.AddItem Trim(rs.Fields(pnColumna0))
                    rs.MoveNext
                Wend
            Else
                While Not rs.EOF
                    pCombo.AddItem Trim(rs.Fields(pnColumna0)) & Space(pnEspacio) & Trim(rs.Fields(pnColumna1))
                    rs.MoveNext
                Wend
            End If
        Else
            If rs.Fields.Count = 1 Then
                While Not rs.EOF
                    If pnFiltro <> rs.Fields(pnColumna1) Then
                        pCombo.AddItem Trim(rs.Fields(pnColumna0)) & Space(pnEspacio) & Trim(rs.Fields(pnColumna1))
                    End If
                    rs.MoveNext
                Wend
            Else
                While Not rs.EOF
                    If pnFiltro <> rs.Fields(pnColumna1) Then
                        pCombo.AddItem Trim(rs.Fields(pnColumna0)) & Space(pnEspacio) & Trim(rs.Fields(pnColumna1))
                    End If
                    rs.MoveNext
                Wend
            End If
        End If
        If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
    End If
End Sub

Public Sub GeneraReporte(pflex As MSHFlexGrid, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency

    For i = 0 To pflex.Rows - 1
        If pnColFiltroVacia = 0 Then
            For j = 0 To pflex.Cols - 1
                pxlHoja1.Cells(i + 1, j + 1) = pflex.TextMatrix(i, j)
            Next j
        Else
            If pflex.TextMatrix(i, pnColFiltroVacia) <> "" Then
                For j = 0 To pflex.Cols - 1
                    pxlHoja1.Cells(i + 1, j + 1) = pflex.TextMatrix(i, j)
                Next j
            End If
        End If
    Next i

End Sub

Public Function GetProveedorRUC(psPersCod As String) As ADODB.Recordset
    Dim Sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion

    Sql = " Select isnull(cPersIDnro,'')cPersIDnro From PersID" _
        & " Where cPersCod = '" & psPersCod & "' And  cPersIDTpo=2 "

    Set GetProveedorRUC = oCon.CargaRecordSet(Sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
