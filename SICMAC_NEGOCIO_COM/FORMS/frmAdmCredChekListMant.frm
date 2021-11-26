VERSION 5.00
Begin VB.Form frmAdmCredChekListMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CheckList "
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmAdmCredChekListMant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   300
      Left            =   5380
      TabIndex        =   8
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   300
      Left            =   6480
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   300
      Left            =   4290
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerarExcel 
      Caption         =   "Generar Excel"
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " CheckList "
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   690
      Width           =   7935
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   6735
      End
      Begin SICMACT.FlexEdit feRequisitos 
         Height          =   4185
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   7382
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Requisito-Estado-idRequisito"
         EncabezadosAnchos=   "300-5900-800-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X"
         ListaControles  =   "0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label1 
         Caption         =   "Categoría:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   400
         Width           =   735
      End
   End
   Begin SICMACT.ActXCodCta_New ActXCodCta 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Texto           =   "Crédito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmAdmCredChekListMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmAdmCredChekListMant
'** Descripción : Formulario que permite realizar el mantenimiento del CheckList
'** Creación    : RECO, 20150421 - ERS010-2015
'**********************************************************************************************
Option Explicit
Dim sCtaCod As String
Dim sTpoCred As String
Dim nTpoOpe As Integer
Dim vMatriz() As Variant
Dim nInicializa As Integer

Public Sub Inicio(ByVal psTitulo As String, ByVal pnTpoOpe As Integer)
    Me.Caption = Me.Caption & " - " & psTitulo
    LimpiarFormulario
    nTpoOpe = pnTpoOpe
    Me.Show 1
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    Dim obj As New COMNCredito.NCOMCredito
    Dim rs As New ADODB.Recordset
    Dim rsCred As New ADODB.Recordset
    Dim sMensaje As String
    Dim sTpoCredCod As String
    
    If KeyAscii = vbKeyReturn Then
        Set rsCred = obj.DevuelveDatosCred(ActXCodCta.NroCuenta)
        If Not (rsCred.EOF And rsCred.BOF) Then sTpoCredCod = rsCred!cTpoCredCod
                
        Set rs = obj.DevuelveDatosAutorizacion(ActXCodCta.NroCuenta)
        sMensaje = ValidaAutorizacion(rs)
        If sMensaje = "" Then
            nInicializa = 1
            sCtaCod = ActXCodCta.NroCuenta
            sTpoCred = sTpoCredCod
            If CargarCombo(sTpoCredCod) = 1 Then
                feRequisitos.Clear
                FormateaFlex feRequisitos
                nInicializa = 2
                Call CargarMatriz
                Call CargaDatoMatriz(cboCategoria.ItemData(cboCategoria.ListIndex))
                cmdGrabar.Enabled = True
                cmdGenerarExcel.Enabled = True
                ActXCodCta.Enabled = False
                cmdGrabar.Visible = IIf(nTpoOpe = 1, True, False)
            End If
        Else
            MsgBox sMensaje, vbInformation, "Alerta"
        End If
    End If
End Sub

Private Function ValidaAutorizacion(ByVal pRs As ADODB.Recordset) As String
    If Not (pRs.EOF And pRs.BOF) Then
        If pRs!cCtaCod <> ActXCodCta.NroCuenta Then
            ValidaAutorizacion = "No se encontró autorización de mantenimiento para esta cuenta. Consulte con su Jefatura"
            Exit Function
        End If
        If pRs!cPersCodAuto <> gsCodPersUser Then
            ValidaAutorizacion = "Usted no tiene autorización para dar mantenimiento a esta cuenta. Consulte con su Jefatura"
            Exit Function
        End If
        If Format(pRs!dFecAuto, "dd/MM/yyyy") <> gdFecSis Then
            ValidaAutorizacion = "La fecha de autorización se encuentra caducada. Consulte con su Jefatura"
            Exit Function
        End If
        If pRs!nEstado <> 1 Then
            ValidaAutorizacion = "La autorización ya fue utilizada. Consulte con su Jefatura"
            Exit Function
        End If
    Else
        ValidaAutorizacion = "No se encontró datos del crédito."
    End If
End Function

Private Sub LimpiarFormulario()
    ActXCodCta.NroCuenta = ""
    'cboCategoria.ListIndex = 0
    cboCategoria.Clear
    feRequisitos.Clear
    FormateaFlex feRequisitos
    cmdGrabar.Enabled = False
    cmdGrabar.Visible = True
    ActXCodCta.Enabled = True
End Sub

Private Function CargarCombo(ByVal psTpoCred As String) As Integer
    Dim obj As New COMNCredito.NCOMCredito
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    Set rs = obj.ListaCredAdmSecciones(Mid(psTpoCred, 1, 2) & "0")
    If Not (rs.BOF And rs.EOF) Then
        For i = 1 To rs.RecordCount
                cboCategoria.AddItem "" & rs!cDescripcion
                cboCategoria.ItemData(cboCategoria.NewIndex) = "" & rs!nIdSeccion
                rs.MoveNext
        Next
        cboCategoria.ListIndex = 0
        CargarCombo = 1
    Else
        CargarCombo = 0
    End If
End Function
        
Private Sub CargarRequisitos()
    Dim obj As New COMNCredito.NCOMCredito
    Dim oCon As New COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    feRequisitos.Clear
    FormateaFlex feRequisitos
    If nInicializa = 1 Then
        Set rs = obj.ListaCredAdmRequisitos(cboCategoria.ItemData(cboCategoria.ListIndex))
        feRequisitos.CargaCombo oCon.RecuperaConstantes(10064)
        If Not (rs.EOF And rs.BOF) Then
            For i = 1 To rs.RecordCount
                feRequisitos.AdicionaFila
                feRequisitos.TextMatrix(i, 3) = rs!nIdRequisito
                feRequisitos.TextMatrix(i, 2) = "N/A" & Space(75) & "3"
                feRequisitos.TextMatrix(i, 1) = rs!cDescripcion
                rs.MoveNext
            Next
        End If
    Else
        Call CargaDatoMatriz(cboCategoria.ItemData(cboCategoria.ListIndex))
    End If
    feRequisitos.TopRow = 1
End Sub

Private Sub cboCategoria_Click()
    If nInicializa = 2 Then
        Call ActualizaMatriz
    End If
    Call CargarRequisitos
End Sub

Private Sub CargarMatriz()
    Dim obj As New COMNCredito.NCOMCredito
    Dim rs As New ADODB.Recordset
    Dim i As Integer, nCantidad As Integer
    
    Set rs = obj.ObtieneRequisitosTpoCred(Mid(sTpoCred, 1, 2) & "0", sCtaCod)
    If Not (rs.BOF And rs.EOF) Then
        nCantidad = rs.RecordCount
        ReDim Preserve vMatriz(3, 0 To nCantidad - 1)
        For i = 0 To nCantidad - 1
            vMatriz(1, i) = rs!nIdSeccion
            vMatriz(2, i) = rs!cRequisitoDesc & Space(50) & rs!nIdRequisito
            vMatriz(3, i) = rs!cEstado & Space(75) & rs!nEstado
            rs.MoveNext
        Next
    End If
End Sub

Private Sub ActualizaMatriz()
    Dim i As Integer, J As Integer
    For i = 1 To feRequisitos.Rows - 1
        For J = 0 To UBound(vMatriz, 2)
            If feRequisitos.TextMatrix(i, 3) = Trim(Right(vMatriz(2, J), 4)) Then
                vMatriz(3, J) = feRequisitos.TextMatrix(i, 2)
            End If
        Next
    Next
End Sub

Private Sub CargaDatoMatriz(ByVal pnSeccion As Integer)
    Dim i As Integer, J As Integer
    For J = 0 To UBound(vMatriz, 2)
        If pnSeccion = vMatriz(1, J) Then
            feRequisitos.AdicionaFila
            feRequisitos.TextMatrix(feRequisitos.Rows - 1, 1) = Trim(Mid(vMatriz(2, J), 1, Len(vMatriz(2, J)) - 4))
            feRequisitos.TextMatrix(feRequisitos.Rows - 1, 2) = vMatriz(3, J)
            feRequisitos.TextMatrix(feRequisitos.Rows - 1, 3) = Trim(Right(vMatriz(2, J), 4))
        End If
    Next
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarFormulario
End Sub

Private Sub cmdGenerarExcel_Click()
    Call ActualizaMatriz
    Call Imprimir
End Sub

Private Sub cmdGrabar_Click()
    Dim obj As New COMNCredito.NCOMCredito
    If MsgBox("Esta seguro que desea guardar la información.", vbQuestion + vbYesNo) = vbYes Then
        Dim i As Integer, J As Integer
        Dim idEstado As Integer, idrequisito As Integer
        Screen.MousePointer = 11
        cmdGrabar.Enabled = False
        Call ActualizaMatriz
        Call obj.EliminaCheckListCta(sCtaCod)
        For J = 0 To UBound(vMatriz, 2)
            Call obj.RegistraRequisitoCta(sCtaCod, Trim(Right(vMatriz(2, J), 4)), Trim(Right(vMatriz(3, J), 4)), nTpoOpe)
        Next
        Screen.MousePointer = 0
        cmdGenerarExcel.Enabled = True
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Imprimir()
    Dim fs As New Scripting.FileSystemObject
    Dim xlsAplicacion As New Excel.Application
    Dim obj As New COMNCredito.NCOMCredito
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim rs As New ADODB.Recordset
    Dim rsObs As New ADODB.Recordset
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo As String, lsFile As String, lsNomHoja As String
    Dim i As Integer, J As Integer, IniTablas As Integer
    Dim lbExisteHoja As Boolean
    Dim rsCred As New ADODB.Recordset
    Set rs = obj.ObtenerDatosCredChekList(sCtaCod)
    lsNomHoja = "REQUISITOS-LIST"
    lsFile = "CHECK_LIST_REQUISITOS"
    
    lsArchivo = "\spooler\" & "CheckList" & "_" & sCtaCod & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"  'COMENTADO POR pti1 20/07/2018
    'lsArchivo = "\spooler\" & "CheckList" & "_" & sCtaCod & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xlsx"  'AGREGADO POR pti1 20/07/2018
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia" 'COMENTADO POR pti1 20/07/2018
        'MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xlsx), Consulte con el Area de TI", vbInformation, "Advertencia" 'AGREGADO POR pti1 20/07/2018
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    Dim nContador As Integer
    Dim nFilaCan As Integer
    Dim nFilaAnt As Integer
    IniTablas = 13
    nContador = IniTablas
    
    If Not (rs.EOF And rs.BOF) Then
        xlHoja1.Cells(5, 4) = rs!cPersNombre
        xlHoja1.Cells(5, 6) = rs!cPersCodSbs
        xlHoja1.Cells(5, 11) = rs!cCtaCod
        xlHoja1.Cells(7, 4) = rs!cUser
        xlHoja1.Cells(7, 6) = rs!cTpoPrd
        xlHoja1.Cells(7, 13) = rs!dVigencia
        xlHoja1.Cells(9, 4) = rs!cAgeDescripcion
        xlHoja1.Cells(9, 6) = rs!nMontoCol
        xlHoja1.Cells(9, 13) = rs!cModalidad
    End If
    
    For i = 0 To cboCategoria.ListCount - 1
        cboCategoria.ListIndex = i
        'nContador = IIf(nContador = 8, 13, nContador) agregado por pti1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
        If ValidaExisteItem(cboCategoria.ListIndex) = True Then
            xlHoja1.Cells(nContador, 3) = i + 1 & " " & cboCategoria.Text
            xlHoja1.Range(xlHoja1.Cells(nContador, 3), xlHoja1.Cells(nContador, 18)).Merge
            xlHoja1.Range(xlHoja1.Cells(nContador, 3), xlHoja1.Cells(nContador, 18)).Style = "Salida"
            nContador = nContador + 2
            nFilaAnt = nContador
            nFilaCan = 1
        End If
          '********************Agregado por pti1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
        Dim nContadorColumna As Integer
        Dim nContadorGrupo As Integer
        Dim nContadorGrupoEstado As Boolean
        
        nContadorGrupoEstado = False
        nContadorGrupo = 0
        nContadorColumna = 0
        
        For J = 0 To UBound(vMatriz, 2)
            If Trim(cboCategoria.ItemData(cboCategoria.ListIndex)) = Trim(vMatriz(1, J)) Then
          
              If Right(vMatriz(3, J), 1) <> 3 Then
               nContadorColumna = nContadorColumna + 1
               nContadorGrupo = nContadorGrupo + 1
               If nContadorColumna = 1 Then
                        xlHoja1.Cells(nContador, 2) = nContadorGrupo
                        xlHoja1.Cells(nContador, 3) = Trim(Mid(vMatriz(2, J), 1, Len(vMatriz(2, J)) - 4))
                        xlHoja1.Range(xlHoja1.Cells(nContador, 3), xlHoja1.Cells(nContador, 3 + 2)).Merge
                        xlHoja1.Cells(nContador, 6) = Trim(Mid(vMatriz(3, J), 1, Len(vMatriz(3, J)) - 1))
                        xlHoja1.Range(xlHoja1.Cells(nContador, 6), xlHoja1.Cells(nContador, 6)).Borders.LineStyle = 1
                        nContadorGrupoEstado = True
               End If
        
               If nContadorColumna = 2 Then
                        xlHoja1.Cells(nContador, 8) = nContadorGrupo
                        xlHoja1.Cells(nContador, 9) = Trim(Mid(vMatriz(2, J), 1, Len(vMatriz(2, J)) - 4))
                        xlHoja1.Range(xlHoja1.Cells(nContador, 9), xlHoja1.Cells(nContador, 9 + 2)).Merge
                        xlHoja1.Cells(nContador, 12) = Trim(Mid(vMatriz(3, J), 1, Len(vMatriz(3, J)) - 1))
                        xlHoja1.Range(xlHoja1.Cells(nContador, 12), xlHoja1.Cells(nContador, 12)).Borders.LineStyle = 1
                        nContadorGrupoEstado = True
               End If
               If nContadorColumna = 3 Then
                        xlHoja1.Cells(nContador, 14) = nContadorGrupo
                        xlHoja1.Cells(nContador, 15) = Trim(Mid(vMatriz(2, J), 1, Len(vMatriz(2, J)) - 4))
                        xlHoja1.Range(xlHoja1.Cells(nContador, 15), xlHoja1.Cells(nContador, 15 + 2)).Merge
                        xlHoja1.Cells(nContador, 18) = Trim(Mid(vMatriz(3, J), 1, Len(vMatriz(3, J)) - 1))
                        xlHoja1.Range(xlHoja1.Cells(nContador, 18), xlHoja1.Cells(nContador, 18)).Borders.LineStyle = 1
                        nContadorColumna = 0
                        nContador = nContador + 2
                        nContadorGrupoEstado = False
               End If
               
             End If
        ' ***********************Fin agregado
            ' If nFilaCan <= IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 6, 4) Then
                '    If Right(vMatriz(3, J), 1) <> 3 Then
                 '       xlHoja1.Cells(nContador, 2) = J + 1
                 '       xlHoja1.Cells(nContador, 3) = Trim(Mid(vMatriz(2, J), 1, Len(vMatriz(2, J)) - 4))
                  '      xlHoja1.Range(xlHoja1.Cells(nContador, 3), xlHoja1.Cells(nContador, 3 + 2)).Merge
                 '       xlHoja1.Cells(nContador, 6) = Trim(Mid(vMatriz(3, J), 1, Len(vMatriz(3, J)) - 1))
                 '       xlHoja1.Range(xlHoja1.Cells(nContador, 6), xlHoja1.Cells(nContador, 6)).Borders.LineStyle = 1
                 '       nContador = nContador + 2
                 '       nFilaCan = nFilaCan + 1
                 '   End If
               ' ElseIf nFilaCan <= IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8) Then
                   ' If Right(vMatriz(3, J), 1) <> 3 Then
                '        xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 8) = J + 1
                     '   xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 9) = Trim(Mid(vMatriz(2, J), 1, Len(vMatriz(2, J)) - 4))
                     '   xlHoja1.Range(xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 9), xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 9 + 2)).Merge
                     '   xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 12) = Trim(Mid(vMatriz(3, J), 1, Len(vMatriz(3, J)) - 1))
                    '    xlHoja1.Range(xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 12), xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 12)).Borders.LineStyle = 1
                   '     nContador = nContador + 2
                 '       nFilaCan = nFilaCan + 1
                '    End If
              '  Else
                    'If Right(vMatriz(3, J), 1) <> 3 Then
                       ' xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 14) = J + 1
                        'xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 15) = Trim(Mid(vMatriz(2, J), 1, Len(vMatriz(2, J)) - 4))
                        'xlHoja1.Range(xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 15), xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 15 + 2)).Merge
                        'xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 18) = Trim(Mid(vMatriz(3, J), 1, Len(vMatriz(3, J)) - 1))
                       ' xlHoja1.Range(xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 18), xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 18)).Borders.LineStyle = 1
                       ' nContador = nContador + 2
                       ' nFilaCan = nFilaCan + 1
                   ' End If
                'End If
                ' Fin Comentado *****************************
            End If
        Next J
         If nContadorGrupoEstado = True Then 'AGREGADO POR PTI1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
            nContador = nContador + 2 'AGREGADO POR PTI1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
         End If 'AGREGADO POR PTI1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
      Next i
    xlHoja1.Cells(nContador + 4, 3) = "RECOMENDACIÓN Y/O OBSERVACIONES"
    Dim nContadorObs As Integer
    nContadorObs = 1
    nContador = nContador + 4
    Set rsCred = oCred.BuscaObsAdmCred(sCtaCod)
    If Not (rsCred.EOF And rsCred.BOF) Then
        Do While Not rsCred.EOF
            xlHoja1.Cells(nContador + nContadorObs, 3) = nContadorObs
            xlHoja1.Cells(nContador + nContadorObs, 4) = Trim(rsCred!cDescripcion)
            xlHoja1.Range(xlHoja1.Cells(nContador + nContadorObs, 4), xlHoja1.Cells(nContador + nContadorObs, 18)).Merge
            xlHoja1.Range(xlHoja1.Cells(nContador + nContadorObs, 3), xlHoja1.Cells(nContador + nContadorObs, 18)).Borders.LineStyle = 5
            nContadorObs = nContadorObs + 1
            rsCred.MoveNext
        Loop
        rsCred.Close
        Set rsCred = Nothing
    End If
    Set oCred = Nothing
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.Path & lsArchivo
    psArchivoAGrabarC = App.Path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Private Function ValidaExisteItem(ByVal pnIndex As Integer) As Boolean
    Dim J As Integer
    ValidaExisteItem = False
    For J = 0 To UBound(vMatriz, 2)
        If Trim(cboCategoria.ItemData(pnIndex)) = Trim(vMatriz(1, J)) Then
            If Right(vMatriz(3, J), 1) <> 3 Then
                ValidaExisteItem = True
            End If
        End If
    Next J
End Function

