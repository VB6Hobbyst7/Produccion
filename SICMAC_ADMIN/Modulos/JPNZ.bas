Attribute VB_Name = "JPNZ"
Option Explicit

Public gcOpeCod As String
Public gcOpeDesc As String
Global Const gsPermisoHorario = "629990"
Global Const gnGradoMaxAut = 9


Public Type Penalidad
    nombre As String
    ruc As String
    direccion As String
    monto As String
    Concepto As String
    bit As Integer
End Type
Public datos() As Penalidad

Public Sub CargaCombo(rs As ADODB.Recordset, pCombo As ComboBox, Optional pnEspacio As Long = 100, Optional pnColumna0 As Integer = 0, Optional pnColumna1 As Integer = 1, Optional pnFiltro As Integer = -1)
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

Public Sub UbicaCombo(pCombo As ComboBox, psDato As String, Optional pbBuscaFinal As Boolean = True, Optional pnNumComp As Integer = 7)
    Dim I As Integer
    Dim lbBan As Boolean
    lbBan = False
    
    If pbBuscaFinal Then
        For I = 0 To pCombo.ListCount - 1
            If Trim(Right(pCombo.List(I), pnNumComp)) = Trim(Right(psDato, pnNumComp)) Then
                lbBan = True
                pCombo.ListIndex = I
                I = pCombo.ListCount
            End If
        Next I
    Else
        For I = 0 To pCombo.ListCount - 1
            If Trim(Left(pCombo.List(I), pnNumComp)) = Trim(Left(psDato, pnNumComp)) Then
                lbBan = True
                pCombo.ListIndex = I
                I = pCombo.ListCount
            End If
        Next I
    End If
    
    If Not lbBan Then pCombo.ListIndex = -1
End Sub

Public Function FiltroCadena(psCadena As String) As String
    Dim lsCadena As String
    lsCadena = psCadena
    lsCadena = Replace(lsCadena, Chr(39), " ")
    lsCadena = Replace(lsCadena, Chr(34), " ")
    FiltroCadena = lsCadena
End Function

'***************************************************
'**********************************************
'MSHFlexGridEdit: Procedimeinto que controla el desplazamiento
'del control Edit en la cuadricula.
'**********************************************
Public Sub MSHFlexGridEdit(ByRef MSHFlexGrid As Control, _
Edt As Control, KeyAscii As Integer, Optional pbTextBox As Boolean = True, Optional pbFechaHora As Boolean = False, Optional pbSoloFecha As Boolean = False, Optional pbEsComboBox As Boolean = False)

   ' Usar el carácter escrito.
   Select Case KeyAscii
   ' Un espacio significa modificar el texto actual.
   Case 0 To 31
      Edt = MSHFlexGrid
      If Not pbEsComboBox Then Edt.SelStart = 1000

   Case 32
    If pbEsComboBox Then
      Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
               MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
               MSHFlexGrid.CellWidth - 8 _
               
      UbicaCombo Edt, Right(MSHFlexGrid, 5)
      Edt.Visible = True
      Edt.SetFocus

    Else
      Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
               MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
               MSHFlexGrid.CellWidth - 8, _
               MSHFlexGrid.CellHeight - 8
      If MSHFlexGrid = "" Then Exit Sub
      Edt = MSHFlexGrid
      Edt.Visible = True
      Edt.SetFocus
      Edt.SelStart = 0
      Edt.SelLength = Len(Edt)
    End If
   
   ' Otro carácter reemplaza el texto actual.
   Case Else
    If pbEsComboBox Then
        Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
           MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
           MSHFlexGrid.CellWidth - 8 _
           
    Else
        Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
           MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
           MSHFlexGrid.CellWidth - 8, _
           MSHFlexGrid.CellHeight - 8
    End If
    
   Edt.Visible = True
   
   Edt.SetFocus
   If pbTextBox Then
        Edt = Chr(KeyAscii)
        Edt.SelStart = 1
   ElseIf pbEsComboBox Then
        Edt.ListIndex = 0
   Else
        If pbFechaHora Then
            Edt = Chr(KeyAscii) & "_/__/____ __:__:__"
        ElseIf pbSoloFecha Then
            Edt = Chr(KeyAscii) & "_/__/____"
        Else
            Edt = Chr(KeyAscii) & "_:__:__"
        End If
        Edt.SelStart = 1
   End If
   
   If Not pbEsComboBox Then
        If Not pbTextBox Then Edt.SelLength = 20
   End If
   
   End Select

   ' Mostrar Edt en la posición correcta.
   ' Y hacer que funcione.
   'Edt.SelStart = 0
   'Edt.SelLength = Len(Edt.Text)
   
End Sub

'**********************************************
'EditKeyCode: Procedimeinto que controla el desempeño de las
'teclas del control dado como parametro, la teclas ESC, ENTER,
'fechas de arriba y abajo
'**********************************************

Public Sub EditKeyCode(MSHFlexGrid As Control, ByVal Edt As _
Control, KeyCode As Integer, Shift As Integer)

    ' Procesamiento del control de edición estándar.
   Select Case KeyCode

   Case 27   ' ESC: ocultar, devuelve el enfoque a
             ' MSFlexGrid.
      Edt.Visible = False
      MSHFlexGrid.SetFocus

   Case 13   ' ENTRAR devuelve el enfoque a MSHFlexGrid.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.row < MSHFlexGrid.Rows - 1 Then
         MSHFlexGrid.row = MSHFlexGrid.row + 1
      End If
      'MSHFlexGrid.SetFocus

   Case 38      ' Arriba.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.row > MSHFlexGrid.FixedRows Then
         MSHFlexGrid.row = MSHFlexGrid.row - 1
      End If

   Case 40      ' Abajo.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.row < MSHFlexGrid.Rows - 1 Then
         MSHFlexGrid.row = MSHFlexGrid.row + 1
      End If
   End Select
End Sub

'*******************************************************************

Public Sub SetFlexEdit(flex As FlexEdit, rs As ADODB.Recordset)
    If rs.EOF And rs.BOF Then
        flex.Clear
        flex.Rows = 2
        flex.FormaCabecera
    Else
        Set flex.Recordset = rs
    End If
End Sub

Public Function FlexARecordSet(pflex As MSHFlexGrid) As ADODB.Recordset
    Dim lnI As Integer
    Dim lnJ As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lnTipoDato As DataTypeEnum
        
    For lnJ = 0 To pflex.Cols - 1
        If Len(Trim(pflex.TextMatrix(0, lnJ))) >= 16 Then
            lnTipoDato = adVarChar
        Else
            If IsNumeric(pflex.TextMatrix(0, lnJ)) And pflex.ColAlignment(lnJ) >= 7 Then
                lnTipoDato = adDouble
            Else
                If IsDate(pflex.TextMatrix(0, lnJ)) Then
                    lnTipoDato = adDate
                Else
                    lnTipoDato = adVarChar
                End If
            End If
        End If
        If lnTipoDato = adVarChar Then
            rs.Fields.Append pflex.TextMatrix(0, lnJ), lnTipoDato, Val(adVarChar), adFldMayBeNull
        Else
            rs.Fields.Append pflex.TextMatrix(0, lnJ), lnTipoDato, , adFldMayBeNull
        End If
    Next lnJ
    
    rs.Open
    For lnI = 1 To pflex.Rows - 1
        rs.AddNew
        For lnJ = 0 To pflex.Cols - 1
            If IsNumeric(pflex.TextMatrix(lnI, lnJ)) Then
                If lnJ <> 6 Then
                    rs.Fields(lnJ) = CCur(IIf(pflex.TextMatrix(lnI, lnJ) = "", "0", pflex.TextMatrix(lnI, lnJ)))
                End If
            Else
                If IsDate(pflex.TextMatrix(lnI, lnJ)) Then
                    rs.Fields(lnJ) = IIf(pflex.TextMatrix(lnI, lnJ) = "", Null, pflex.TextMatrix(lnI, lnJ))
                Else
                    rs.Fields(lnJ) = pflex.TextMatrix(lnI, lnJ)
                End If
            End If
        Next lnJ
        rs.Update
    Next lnI
    rs.MoveFirst
    Set FlexARecordSet = rs
End Function

Public Function nombre()
    nombre = "Jomark"
End Function

Public Function SetSumaFlexCol(ByVal flex As FlexEdit, pnCol As Integer) As Currency
    Dim I As Integer
    Dim lnSuma As Currency
    
    lnSuma = 0
    For I = 1 To flex.Rows - 1
        lnSuma = lnSuma + IIf(IsNumeric(flex.TextMatrix(I, pnCol)), flex.TextMatrix(I, pnCol), 0)
    Next I
    
    SetSumaFlexCol = lnSuma
End Function


Public Function TieneAcceso(psOpeCod As String) As Boolean
    Dim I As Integer

    For I = 0 To NroRegOpe - 1
        If MatOperac(I, 0) = psOpeCod Then
            TieneAcceso = True
            Exit Function
        End If
    Next I

    TieneAcceso = False
    Exit Function
End Function

Public Function ValidaConfiguracionRegional() As Boolean
Dim nMoneda As Currency
Dim nMonto As Double
Dim sNumero As String, sFecha As String
Dim nPosPunto As Integer, nPosComa As Integer

'Inicializamos las variables
ValidaConfiguracionRegional = True
nMoneda = 1234567
nMonto = 1234567
'Validamos Configuración de punto y Coma de Moneda
sNumero = Format$(nMoneda, "#,##0.00")
nPosPunto = InStr(1, sNumero, ".", vbTextCompare)
nPosComa = InStr(1, sNumero, ",", vbTextCompare)

If nPosPunto < nPosComa Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
'Validamos la configuración del punto y coma de los números
sNumero = Format$(nMonto, "#,##0.00")
nPosPunto = InStr(1, sNumero, ".", vbTextCompare)
nPosComa = InStr(1, sNumero, ",", vbTextCompare)
If nPosPunto < nPosComa Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
'Validamos la fecha y la configuración de la hora
If Date <> Format$(Date, "dd/MM/yyyy") Then 'Validar el formato de la fecha
    ValidaConfiguracionRegional = False
    Exit Function
End If

sFecha = Format$(Date & " " & Time, "dd/mm/yyyy hh:mm:ss AMPM")
If InStr(1, sFecha, "A.M.", vbTextCompare) > 0 Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
If InStr(1, sFecha, "P.M.", vbTextCompare) > 0 Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
sFecha = Trim(Date)
If Day(Date) <> CInt(Mid(sFecha, 1, 2)) Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
If Month(Date) <> CInt(Mid(sFecha, 4, 2)) Then
    ValidaConfiguracionRegional = False
    Exit Function
End If
If Year(Date) <> CInt(Mid(sFecha, 7, 4)) Then
    ValidaConfiguracionRegional = False
    Exit Function
End If

End Function


Public Function fgActualizaUltVersionEXE(psAgenciaCod As String) As Boolean
Dim fs As Scripting.FileSystemObject
Dim fCurrent As Scripting.Folder
Dim fi As Scripting.File
Dim fd As Scripting.File

Dim lsRutaUltActualiz As String
Dim lsRutaSICMACT As String
Dim lsFecUltModifLOCAL As String
Dim lsFecUltModifORIGEN As String
Dim lsFlagActualizaEXE As String

On Error GoTo ERROR
    fgActualizaUltVersionEXE = False
    Dim oCons As NConstSistemas
    Set oCons = New NConstSistemas
    
    lsRutaUltActualiz = oCons.GetRutaAcceso(psAgenciaCod)
    lsRutaSICMACT = App.path & "\"
    lsFlagActualizaEXE = oCons.LeeConstSistema(49)
    
    If lsFlagActualizaEXE = "0" Then  ' No Actualiza Ejecutable
        Exit Function
    End If
    
    If Dir(lsRutaSICMACT & "*.*") = "" Then
        Exit Function
    End If
    If Dir(lsRutaUltActualiz & "*.*") = "" Then
        Exit Function
    End If
 
    Set fs = New Scripting.FileSystemObject
    Set fCurrent = fs.GetFolder(lsRutaUltActualiz)
    For Each fi In fCurrent.Files
          If Right(UCase(fi.Name), 3) = "EXE" Or Right(UCase(fi.Name), 3) = "INI" Or Right(UCase(fi.Name), 3) = "DLL" Then
             lsFecUltModifORIGEN = Format(fi.DateLastModified, "yyyy/mm/dd hh:mm:ss")
             If Dir(lsRutaSICMACT & fi.Name) <> "" Then
                Set fd = fs.GetFile(lsRutaSICMACT & fi.Name)
                lsFecUltModifLOCAL = Format(fd.DateLastModified, "yyyy/mm/dd hh:mm:ss")
                If lsFecUltModifLOCAL < lsFecUltModifORIGEN And lsFecUltModifORIGEN <> "" Then ' ACTUALIZA
                    fgActualizaUltVersionEXE = True
                End If
             Else
                fgActualizaUltVersionEXE = True
             End If
             If fgActualizaUltVersionEXE = True Then
                Exit For
             End If
          End If
    Next
    If fgActualizaUltVersionEXE = True Then
        frmHerActualizaSicmact.IniciaVariables True
        frmHerActualizaSicmact.Show 1
    End If
    Exit Function

ERROR:
    MsgBox "No se puede acceder a la ruta de origen, de la Ultima Actualizacion. - " & lsRutaUltActualiz, vbInformation, "Aviso"
    fgActualizaUltVersionEXE = False
End Function

Public Sub GeneraReporte(pflex As MSHFlexGrid, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
    Dim I As Integer
    Dim k As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    For I = 0 To pflex.Rows - 1
        If pnColFiltroVacia = 0 Then
            For j = 0 To pflex.Cols - 1
                pxlHoja1.Cells(I + 1, j + 1) = pflex.TextMatrix(I, j)
            Next j
        Else
            If pflex.TextMatrix(I, pnColFiltroVacia) <> "" Then
                For j = 0 To pflex.Cols - 1
                    pxlHoja1.Cells(I + 1, j + 1) = pflex.TextMatrix(I, j)
                Next j
            End If
        End If
    Next I
    
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
                prsDat.Fields(nCol).value = prs.Fields(nCol).value
            Next
            prsDat.Update
        End If
        prs.MoveNext
    Loop
End Function

Public Function ProcesoEjecucionValida(psProceso As String, psNombreOpcion As String) As Boolean
    Dim oProceso As DProcesosEjecucion
    Set oProceso = New DProcesosEjecucion
    
    ProcesoEjecucionValida = oProceso.ProcesoEjecucionValida(gdFecSis & " " & Time, 2, psProceso)
    
    If Not ProcesoEjecucionValida Then
        MsgBox "Ud. no puede ingresar a " & psNombreOpcion & ", pues no tiene permiso de acceso, en este momento.", vbInformation, "Aviso"
    End If
    
    Set oProceso = Nothing
End Function

Public Sub BitacoraSistema(psProceso As String, Optional psComentario As String)
    Dim oBitacora As DBitacoraSistema
    Set oBitacora = New DBitacoraSistema

    oBitacora.BitarcoraSistemaInserta gdFecSis & " " & Time, gsCodUser, 2, gsCodAge, gsCodArea, gcPC, psProceso
    
    Set oBitacora = Nothing
End Sub

Public Function GetProveedorRUC(psPersCod As String) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Dim prs  As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    
    sql = " Select isnull(cPersIDnro,'')cPersIDnro From PersID" _
        & " Where cPersCod = '" & psPersCod & "' And  cPersIDTpo=2 "
    
    Set GetProveedorRUC = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

