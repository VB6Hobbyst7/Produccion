VERSION 5.00
Begin VB.Form frmRHPrestAdmAge 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   ForeColor       =   &H80000008&
   Icon            =   "frmPrestAdmAge.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2010
      TabIndex        =   4
      Top             =   3450
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   930
      TabIndex        =   3
      Top             =   3450
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Seleccione agencia de búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3330
      Left            =   75
      TabIndex        =   5
      Top             =   60
      Width           =   3825
      Begin VB.OptionButton optAge 
         Appearance      =   0  'Flat
         Caption         =   "&Ninguno"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optAge 
         Appearance      =   0  'Flat
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.ListBox lstAge 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   2505
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   705
         Width           =   3570
      End
   End
End
Attribute VB_Name = "frmRHPrestAdmAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function ExisteEnLista(ByVal sCta As String) As Boolean
    Dim L As ListItem
    Dim bExiste As Boolean
    bExiste = False
    For Each L In frmRHPrestamosAdm.lstCredAdm.ListItems
        If L.SubItems(1) = sCta Then
            bExiste = True
            Exit For
        End If
    Next
    ExisteEnLista = bExiste
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
Dim L As ListItem
Dim sAge As String
Dim sCta As String
Dim nMonto As Double, nTC As Double
Dim rsEmp As ADODB.Recordset
Dim nItem As Integer
Dim VSQL As String
Dim oCon As DConecta
Set oCon = New DConecta
nItem = 0

fgITFParametros

Dim sqlP As String
Dim rsP As ADODB.Recordset
Set rsP = New ADODB.Recordset
Dim oConC As DConecta
Set oConC = New DConecta
 
If gbBitCentral Then
    'CENTRALIZADO
    
    oConC.AbreConexion

    sqlP = "Select cPersCod Cod From RRHH Where cRHCod Like 'E%' And nRHEstado Not Like '[78]%'"
    Set rsP = oConC.CargaRecordSet(sqlP)
    
    sqlP = ""
    
    While Not rsP.EOF
        If sqlP = "" Then
            sqlP = rsP!Cod
        Else
            sqlP = sqlP & "','" & rsP!Cod
        End If
        rsP.MoveNext
    Wend
    
    Me.MousePointer = 11
    If oCon.AbreConexion Then
        For i = 0 To lstAge.ListCount - 1
            If lstAge.Selected(i) Then
                sAge = Right(lstAge.List(i), 2)
    
                    Set rsEmp = New ADODB.Recordset
                    rsEmp.CursorLocation = adUseClient
    
                    If frmRHPrestamosAdm.optTipo(0).value Then
                        VSQL = " Select PRD.nPrdEstado, P.cPersCod, P.cPersNombre, C.cCtaCod , PP.dVenc, PP.nCuota cNroCuo," _
                             & "    (Select ISNull(Sum(nMonto),0) From ColocCalendDet COD" _
                             & "        Where COD.cCtaCod = C.cCtaCod And COD.nNroCalen = PP.nNroCalen" _
                             & "            And COD.nColocCalendApl = PP.nColocCalendApl And COD.nCuota = PP.nCuota" _
                             & "    ) nMonto" _
                             & "    FROM Persona P" _
                             & "    INNER JOIN ProductoPersona PC ON P.cPersCod = PC.cPersCod" _
                             & "    INNER JOIN Colocaciones C ON C.cCtaCod = PC.cCtaCod" _
                             & "    INNER JOIN ColocCalendario PP ON PP.cCtaCod = C.cCtaCod" _
                             & "    INNER JOIN Producto PRD ON PC.cCtaCod = PRD.cCtaCod" _
                             & "    WHERE Substring(C.cCtaCod,6,3) In ('" & gColConsuPrestAdm & "','" & gColHipoCaja & "','" & gColHipoMiVivienda & "') And PRD.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ") And PP.nColocCalendEstado = " & gColocCalendEstadoPendiente & "" _
                             & "    And PP.nNroCalen = (Select Max(nNroCalen) From ColocCalendario P2 Where P2.cCtaCod = PP.cCtaCod)" _
                             & "    And PC.nPrdPersRelac = " & gColRelPersTitular & "" _
                             & "    And PP.dVenc = ( Select MIN(dVenc) From ColocCalendario P1 Where P1.cCtaCod = PP.cCtaCod And P1.nColocCalendEstado = " & gColocCalendEstadoPendiente & "  And  P1.nColocCalendApl = " & gColocCalendAplCuota & "  And P1.nNroCalen = PP.nNroCalen)" _
                             & "    And P.cPersCod = '" & Right(frmRHPrestamosAdm.sCodEmp, 13) & "'" _
                             & "    Order by P.cPersNombre, C.cCtaCod"
    
                        'VSQL = "Select P.cCodPers, P.cNomPers, C.cCodCta, PP.dFecVenc, cNroCuo, " _
                            & "(PP.nCapital+PP.nInteres+ISNULL(PP.nMora,0)+ISNULL(PP.nIntGra,0)) Monto " _
                            & "FROM " & oCon.sCentralPers & "Persona P INNER JOIN " _
                            & "PersCredito PC INNER JOIN Credito C INNER JOIN PlandesPag PP ON C.cCodCta = " _
                            & "PP.cCodCta ON PC.cCodCta = C.cCodCta ON P.cCodPers = PC.cCodPers WHERE " _
                            & "C.cCodCta like '___" & sAge & "320%' And C.cEstado = 'F' And PP.cEstado = 'P' " _
                            & "And PP.dFecVenc IN (Select MIN(dFecVenc) From PlandesPag P1 Where " _
                            & "P1.cCodCta = PP.cCodCta And cEstado = 'P') And PC.cRelaCta = 'TI' " _
                            & "And PC.cCodPers = '" & Right(frmRHPrestamosAdm.sCodEmp, 10) & "' Order by P.cNomPers, C.cCodCta"
                    Else
                        VSQL = " Select P.cPersCod , P.cPersNombre , C.cCtaCod , PP.dVenc , PP.nCuota cNroCuo ," _
                             & "    (Select ISNull(Sum(nMonto),0) From ColocCalendDet COD" _
                             & "        Where COD.cCtaCod = C.cCtaCod And COD.nNroCalen = PP.nNroCalen" _
                             & "            And COD.nColocCalendApl = PP.nColocCalendApl And COD.nCuota = PP.nCuota" _
                             & "    ) nMonto" _
                             & "    FROM RRHH RH INNER JOIN Persona P ON P.CPERSCOD = RH.CPERSCOD " _
                             & "    INNER JOIN ProductoPersona PC ON P.cPersCod = PC.cPersCod" _
                             & "    INNER JOIN Colocaciones C ON C.cCtaCod = PC.cCtaCod" _
                             & "    INNER JOIN ColocCalendario PP ON PP.cCtaCod = C.cCtaCod" _
                             & "    INNER JOIN Producto PRD ON PC.cCtaCod = PRD.cCtaCod " _
                             & "    WHERE NrheSTADO < 700 And  Substring(C.cCtaCod,6,3) In ('" & gColConsuPrestAdm & "','" & gColHipoCaja & "','" & gColHipoMiVivienda & "','" & gColConsuDctoPlan & "') And PRD.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & ") And PP.nColocCalendEstado = " & gColocCalendEstadoPendiente & "" _
                             & "    And PP.nNroCalen = (Select Max(nNroCalen) From ColocCalendario P2 Where P2.cCtaCod = PP.cCtaCod)" _
                             & "    And PC.nPrdPersRelac = " & gColRelPersTitular & "" _
                             & "    And PP.dVenc = ( Select MIN(dVenc) From ColocCalendario P1 Where P1.cCtaCod = PP.cCtaCod And P1.nColocCalendEstado = " & gColocCalendEstadoPendiente & " And  P1.nColocCalendApl = " & gColocCalendAplCuota & "  And P1.nNroCalen = PP.nNroCalen)" _
                             & "    Order by P.cPersNombre, C.cCtaCod"
                    End If
    
                    Set rsEmp = oCon.CargaRecordSet(VSQL)
                    
                    Do While Not rsEmp.EOF
                        sCta = rsEmp("cCtaCod")
                        If Not ExisteEnLista(sCta) Then
                            nItem = nItem + 1
                            Set L = frmRHPrestamosAdm.lstCredAdm.ListItems.Add(, , PstaNombre(rsEmp("cPersNombre")))
                            nMonto = rsEmp("nMonto") + Format(fgITFCalculaImpuesto(rsEmp("nMonto")), "0.00")
                            L.SubItems(1) = rsEmp("cCtaCod")
                            L.SubItems(2) = rsEmp("cNroCuo")
                            L.SubItems(3) = Format$(rsEmp("dVenc"), "dd/mm/yyyy")
                            L.SubItems(6) = rsEmp("cPersCod")
                            If Mid(sCta, 9, 1) = Moneda.gMonedaNacional Then
                                L.SubItems(4) = Format$(nMonto, "#,##0.00")
                                L.SubItems(5) = "0.00"
                            Else
                                nTC = CDbl(frmRHPrestamosAdm.lblTCC)
                                L.SubItems(4) = Format$(nMonto * nTC, "#,##0.00")
                                L.SubItems(5) = Format$(nMonto, "#,##0.00")
                                L.ForeColor = &H808000
                                L.ListSubItems(1).ForeColor = &H808000
                                L.ListSubItems(2).ForeColor = &H808000
                                L.ListSubItems(3).ForeColor = &H808000
                                L.ListSubItems(4).ForeColor = &H808000
                                L.ListSubItems(5).ForeColor = &H808000
                            End If
                        End If
                        rsEmp.MoveNext
                    Loop
                    rsEmp.Close
                    Set rsEmp = Nothing
            End If
        Next i
    End If
    
    If nItem = 0 Then
        frmRHPrestamosAdm.txtEmp = ""
        MsgBox "No se encontraron créditos en la agencia seleccionada.", vbInformation, "Aviso"
    End If
    Me.MousePointer = 0
    Unload Me
Else
    'DISTRIBUIDO
    'Dim i As Integer
    'Dim l As ListItem
    'Dim sAge As String
    'Dim sCta As String
    'Dim nMonto As Double, nTC As Double
    'Dim rsEmp As ADODB.Recordset
    'Dim nItem As Integer
    'Dim VSQL As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    'nItem = 0
    
    oConC.AbreConexion
    
    sqlP = "Select right(cPersCod,10) Cod From RRHH Where cRHCod Like 'E%' And nRHEstado Not Like '[78]%'"
    Set rsP = oConC.CargaRecordSet(sqlP)
    
    sqlP = ""
    
    While Not rsP.EOF
        If sqlP = "" Then
            sqlP = rsP!Cod
        Else
            sqlP = sqlP & "','" & rsP!Cod
        End If
        rsP.MoveNext
    Wend
    
    Me.MousePointer = 11
    For i = 0 To lstAge.ListCount - 1
        If lstAge.Selected(i) Then
            sAge = Right(lstAge.List(i), 2)
            sAge = sAge & "0000000000"
            If oCon.AbreConexion Then 'Remota(sAge, True, False) Then
                Set rsEmp = New ADODB.Recordset
                rsEmp.CursorLocation = adUseClient
    
                If frmRHPrestamosAdm.optTipo(0).value Then
                    VSQL = " Select P.cCodPers, P.cNomPers, C.cCodCta, PP.dFecVenc, cNroCuo, " _
                        & " (PP.nCapital+PP.nInteres+ISNULL(PP.nMora,0)+ISNULL(PP.nIntGra,0)) Monto " _
                        & " FROM  Persona P INNER JOIN " _
                        & " PersCredito PC INNER JOIN Credito C INNER JOIN PlandesPag PP ON C.cCodCta = " _
                        & " PP.cCodCta ON PC.cCodCta = C.cCodCta ON P.cCodPers = PC.cCodPers WHERE " _
                        & " Substring(C.cCodCta,3,3) IN ('" & gColConsuPrestAdm & "','" & gColHipoCaja & "','" & gColHipoMiVivienda & "') And C.cEstado = 'F' And PP.cEstado = 'P' " _
                        & " And PP.dFecVenc IN (Select MIN(dFecVenc) From PlandesPag P1 Where " _
                        & " P1.cCodCta = PP.cCodCta And cEstado = 'P') And PC.cRelaCta = 'TI' And PC.cCodPers IN ('" & sqlP & "')" _
                        & " And PC.cCodPers = '" & Right(frmRHPrestamosAdm.sCodEmp, 10) & "' Order by P.cNomPers, C.cCodCta"
                Else
                    VSQL = " Select P.cCodPers, P.cNomPers, C.cCodCta, PP.dFecVenc, cNroCuo, " _
                        & " (PP.nCapital+PP.nInteres+ISNULL(PP.nMora,0)+ISNULL(PP.nIntGra,0)) Monto " _
                        & " FROM  Persona P INNER JOIN " _
                        & " PersCredito PC INNER JOIN Credito C INNER JOIN PlandesPag PP ON C.cCodCta = " _
                        & " PP.cCodCta ON PC.cCodCta = C.cCodCta ON P.cCodPers = PC.cCodPers WHERE " _
                        & " Substring(C.cCodCta,3,3) IN ('" & gColConsuPrestAdm & "','" & gColHipoCaja & "','" & gColHipoMiVivienda & "') And C.cEstado = 'F' And PP.cEstado = 'P' " _
                        & " And PP.dFecVenc IN (Select MIN(dFecVenc) From PlandesPag P1 Where " _
                        & " P1.cCodCta = PP.cCodCta And cEstado = 'P') And PC.cRelaCta = 'TI' And PC.cCodPers IN ('" & sqlP & "')" _
                       & " Order by P.cNomPers, C.cCodCta"
                End If
    
                Set rsEmp = oCon.CargaRecordSet(VSQL)
                Set rsEmp.ActiveConnection = Nothing
                Do While Not rsEmp.EOF
                    sCta = rsEmp("cCodCta")
                    If Not ExisteEnLista(sCta) Then
                        nItem = nItem + 1
                        Set L = frmRHPrestamosAdm.lstCredAdm.ListItems.Add(, , PstaNombre(rsEmp("cNomPers"), True))
                        nMonto = rsEmp("Monto")
                        L.SubItems(1) = rsEmp("cCodCta")
                        L.SubItems(2) = rsEmp("cNroCuo")
                        L.SubItems(3) = Format$(rsEmp("dFecVenc"), "dd/mm/yyyy")
                        L.SubItems(6) = rsEmp("cCodPers")
                        If Mid(sCta, 6, 1) = Moneda.gMonedaNacional Then
                            L.SubItems(4) = Format$(nMonto, "#,##0.00")
                            L.SubItems(5) = "0.00"
                        Else
                            nTC = CDbl(frmRHPrestamosAdm.lblTCC)
                            L.SubItems(4) = Format$(nMonto * nTC, "#,##0.00")
                            L.SubItems(5) = Format$(nMonto, "#,##0.00")
                            L.ForeColor = &H808000
                            L.ListSubItems(1).ForeColor = &H808000
                            L.ListSubItems(2).ForeColor = &H808000
                            L.ListSubItems(3).ForeColor = &H808000
                            L.ListSubItems(4).ForeColor = &H808000
                            L.ListSubItems(5).ForeColor = &H808000
                        End If
                    End If
                    rsEmp.MoveNext
                Loop
                rsEmp.Close
                Set rsEmp = Nothing
            End If
        End If
    Next i
    If nItem = 0 Then
        frmRHPrestamosAdm.txtEmp = ""
        MsgBox "No se encontraron créditos en la agencia seleccionada.", vbInformation, "Aviso"
    End If
    Me.MousePointer = 0
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim rsAge As ADODB.Recordset
Dim oAge As DActualizaDatosArea
Set oAge = New DActualizaDatosArea
Set rsAge = New ADODB.Recordset

Set rsAge = oAge.GetAgencias

Me.Caption = "Selecciona Agencia"
'rsAge.CursorLocation = adUseClient
'VSQL = "Select S.cCodAge, T.cNomTab from Servidor S INNER JOIN " & gcCentralCom & "Tablacod T " _
    & "ON S.cCodAge = RTRIM(T.cValor) Where T.cCodtab like '47%' And S.cNroSer = '01' " _
    & "Order by S.cCodAge"
'rsAge.Open VSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'Set rsAge.ActiveConnection = Nothing
Do While Not rsAge.EOF
    If rsAge!codigo <> "0" Then lstAge.AddItem Trim(rsAge("Descripcion")) & Space(50) & rsAge("Codigo")
    rsAge.MoveNext
Loop
cmdOK.Enabled = False
End Sub

Private Sub lstAge_Click()
Dim i As Integer
Dim bHab As Boolean
bHab = False
For i = 0 To lstAge.ListCount - 1
    If lstAge.Selected(i) Then
        bHab = True
        Exit For
    End If
Next i
cmdOK.Enabled = bHab
End Sub

Private Sub lstAge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdOK.Enabled Then cmdOK.SetFocus
End If
End Sub

Private Sub optAge_Click(Index As Integer)
Dim i As Integer
For i = 0 To lstAge.ListCount - 1
    If Index = 0 Then
        lstAge.Selected(i) = True
    ElseIf Index = 1 Then
        lstAge.Selected(i) = False
    End If
Next i
End Sub


