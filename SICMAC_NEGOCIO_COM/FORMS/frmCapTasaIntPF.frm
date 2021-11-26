VERSION 5.00
Begin VB.Form frmCapTasaIntPF 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmCapTasaIntPF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Cancel          =   -1  'True
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdNuevaBus 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   75
      Picture         =   "frmCapTasaIntPF.frx":030A
      TabIndex        =   10
      Top             =   4860
      Width           =   1035
   End
   Begin VB.Frame fraTasa 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   90
      TabIndex        =   7
      Top             =   45
      Width           =   9735
      Begin VB.ComboBox cboTpoPrograma 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4275
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   472
         Width           =   2295
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   7920
         TabIndex        =   3
         Top             =   360
         Width           =   1035
      End
      Begin VB.ComboBox cboTipoTasa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   472
         Width           =   1935
      End
      Begin VB.ComboBox cboMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   472
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sub Producto:"
         Height          =   195
         Left            =   4275
         TabIndex        =   12
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Tasa:"
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imrpimir"
      Height          =   375
      Left            =   1140
      TabIndex        =   4
      Top             =   4860
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   4920
      Width           =   1035
   End
   Begin VB.Frame fraTarifa 
      Caption         =   "Tasas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3735
      Left            =   90
      TabIndex        =   6
      Top             =   1065
      Width           =   9765
      Begin SICMACT.FlexEdit grdTasas 
         Height          =   2895
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9495
         _extentx        =   16748
         _extenty        =   5106
         cols0           =   11
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Monto Ini-Monto Fin-Plazo Ini-Plazo Fin-Ord Pag-Tasa Int-nTasaCod-Cambio-Activa-bEdit"
         encabezadosanchos=   "300-1500-1500-900-900-800-1300-0-1300-600-1"
         font            =   "frmCapTasaIntPF.frx":064C
         font            =   "frmCapTasaIntPF.frx":0674
         font            =   "frmCapTasaIntPF.frx":069C
         font            =   "frmCapTasaIntPF.frx":06C4
         font            =   "frmCapTasaIntPF.frx":06EC
         fontfixed       =   "frmCapTasaIntPF.frx":0714
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-1-2-3-4-X-6-X-X-9-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0-0-0-0-0-4-0"
         encabezadosalineacion=   "C-R-R-R-R-C-R-L-L-C-C"
         formatosedit    =   "0-2-2-3-3-0-2-0-0-1-0"
         cantentero      =   12
         cantdecimales   =   4
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbformatocol    =   -1
         lbbuscaduplicadotext=   -1
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapTasaIntPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim nProducto As COMDConstantes.Producto
'Dim nmoneda As COMDConstantes.Moneda
'Dim nTipoTasa As COMDConstantes.CaptacTipoTasa
'Dim nTpoPrograma As Integer
'Dim bConsulta As Boolean
'Dim sTitProd As String
'Dim nNroReg As String
''By capi 21012009
'Dim objPista As COMManejador.Pista
'
'
'
'Private Function ValidaTasas() As Boolean
'Dim i As Long, J As Long
'Dim nMontoIni As Double, nMontoFin As Double
'Dim nPlazoIni As Long, nPlazoFin As Long
'Dim nMontoIniAux As Double, nMontoFinAux As Double
'Dim nPlazoIniAux As Long, nPlazoFinAux As Long
'Dim bOrdPag As String, bOrdPagAux As Boolean
'Dim nTpoPrograma As Integer, nTpoProgramaAux As Integer
'Dim nTasa As Double
'For i = 1 To grdTasas.Rows - 1
'    nMontoIni = CDbl(grdTasas.TextMatrix(i, 1))
'    nMontoFin = CDbl(grdTasas.TextMatrix(i, 2))
'    nPlazoIni = CDbl(grdTasas.TextMatrix(i, 3))
'    nPlazoFin = CDbl(grdTasas.TextMatrix(i, 4))
'    nTasa = CDbl(grdTasas.TextMatrix(i, 6))
'    bOrdPag = IIf(grdTasas.TextMatrix(i, 5) = "Si", True, False)
'
'    If nTasa = 0 Then
'        MsgBox "Monto de Tasa no válido, debe ser mayor a cero.", vbInformation, "Aviso"
'        grdTasas.row = i
'        grdTasas.Col = 6
'        grdTasas.SetFocus
'        Exit Function
'    End If
'
'    If nPlazoIni = 0 And nMontoFin = 0 Then
'        ValidaTasas = False
'        MsgBox "Tasa tiene rangos de plazos no válidos, deben ser mayor a cero.", vbInformation, "Aviso"
'        grdTasas.row = i
'        grdTasas.SetFocus
'        Exit Function
'    End If
'
'    For J = 1 To nNroReg - 1
'        If J <> i Then
'            nMontoIniAux = CDbl(grdTasas.TextMatrix(J, 1))
'            nMontoFinAux = CDbl(grdTasas.TextMatrix(J, 2))
'            nPlazoIniAux = CDbl(grdTasas.TextMatrix(J, 3))
'            nPlazoFinAux = CDbl(grdTasas.TextMatrix(J, 4))
'            bOrdPagAux = IIf(grdTasas.TextMatrix(J, 5) = "Si", True, False)
'            'By Capi 20012008
'            If grdTasas.TextMatrix(J, 8) = "." Then
'                If ((nMontoIni >= nMontoIniAux And nMontoIni < nMontoFinAux) Or (nMontoFin <= nMontoFinAux And nMontoFin > nMontoIniAux)) And _
'                    ((nPlazoIni >= nPlazoIniAux And nPlazoIni <= nPlazoFinAux) Or (nPlazoFin <= nPlazoFinAux And nPlazoFin >= nPlazoIniAux)) Then
'                    ValidaTasas = False
'                    MsgBox "Tasa tiene rangos de montos, plazos, o la opción de Orden Pago No Válida no válidos.", vbInformation, "Aviso"
'                    grdTasas.row = J
'                    grdTasas.SetFocus
'                    Exit Function
'                End If
'           End If
'        End If
'    Next J
'Next i
'ValidaTasas = True
'End Function
'
'Private Sub IniciaCombos(ByRef combo As ComboBox, ByVal nConstante As ConstanteCabecera)
'    Dim clsGen As COMDConstSistema.DCOMGeneral
'    Dim rsGen As ADODB.Recordset
'    Set clsGen = New COMDConstSistema.DCOMGeneral
'    Set rsGen = clsGen.GetConstante(nConstante)
'    Set clsGen = Nothing
'
'    Do While Not rsGen.EOF
'        combo.AddItem rsGen("cDescripcion") & Space(100) & rsGen("nConsValor")
'        rsGen.MoveNext
'    Loop
'
'    combo.ListIndex = 0
'    rsGen.Close
'    Set rsGen = Nothing
'End Sub
'
'Public Sub Inicia(Optional nProd As Producto = gCapAhorros, Optional bCons As Boolean = False)
'    bConsulta = bCons
'    IniciaCombos cboMoneda, gMoneda
'    IniciaCombos cboTipoTasa, gCaptacTipoTasa
'    cboMoneda.ListIndex = 0
'    cboTipoTasa.ListIndex = 0
'    nProducto = nProd
'
'    sTitProd = "PLAZO FIJO"
'    grdTasas.ColWidth(5) = 0
'    IniciaCombos cboTpoPrograma, gCaptacSubProdPlazoFijo
'    'by capi 21092009
'    gsOpeCod = gPFMantTasaInteres
'    '
'
'    If bConsulta Then
'        'cmdNuevo.Visible = False
'       ' cmdEliminar.Visible = False
'        cmdGrabar.Visible = False
'        grdTasas.lbEditarFlex = False
'        Me.Caption = "Captaciones - " & sTitProd & " - Tasas Interés - Consulta"
'    Else
'        'cmdNuevo.Visible = True
'        'cmdEliminar.Visible = True
'        cmdGrabar.Visible = True
'        Me.Caption = "Captaciones - " & sTitProd & " - Tasas Interés - Mantenimiento"
'        'cmdNuevo.Enabled = False
'        cmdGrabar.Enabled = False
'        'cmdEliminar.Enabled = False
'        grdTasas.lbEditarFlex = True
'    End If
'    cmdImprimir.Enabled = False
'    cmdNuevaBus.Enabled = False
'
'    Me.Show 1
'End Sub
'
'Private Sub cboMoneda_Click()
'If cboMoneda.ListIndex = 0 Then
'    nmoneda = gMonedaNacional
'    grdTasas.BackColor = &HC0FFFF
'Else
'    nmoneda = gMonedaExtranjera
'    grdTasas.BackColor = &HC0FFC0
'End If
'End Sub
'
'Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cboTipoTasa.SetFocus
'End If
'End Sub
'
'Private Sub cboTipoTasa_Click()
'    nTipoTasa = CLng(Trim(Right(cboTipoTasa, 4)))
'End Sub
'
'Private Sub cboTipoTasa_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cmdAplicar.SetFocus
'End If
'End Sub
'
'Private Sub cboTpoPrograma_Click()
'nTpoPrograma = CInt(Trim(Right(cboTpoPrograma, 4)))
'sTitProd = sTitProd & " - " & Trim(Left(cboTpoPrograma.Text, 20))
'End Sub
'
'Private Sub cmdAplicar_Click()
'    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
'    Dim rsTasa As ADODB.Recordset
'
'    Dim L As ListItem
'
'
'    Set rsTasa = New ADODB.Recordset
'
'    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
'    Set rsTasa = clsDef.GetTarifarioPF(nProducto, nmoneda, nTipoTasa, gsCodAge, nTpoPrograma)
'    Set clsDef = Nothing
'    nNroReg = 0
'    If Not (rsTasa.EOF And rsTasa.BOF) Then
'        Set grdTasas.Recordset = rsTasa
'        nNroReg = grdTasas.Rows
'        If Not bConsulta Then
'            'cmdNuevo.Enabled = True
'            'cmdEliminar.Enabled = True
'            cmdGrabar.Enabled = True
'            grdTasas.lbEditarFlex = True
'        End If
'    Else
'        If Not bConsulta Then
'            'cmdNuevo.Enabled = True
'            'cmdEliminar.Enabled = False
'            cmdGrabar.Enabled = True
'            grdTasas.lbEditarFlex = True
'        End If
'    End If
'grdTasas.FormateaColumnas
'cmdAplicar.Enabled = False
'cboMoneda.Enabled = False
'cboTipoTasa.Enabled = False
'cboTpoPrograma.Enabled = False
'cmdNuevaBus.Enabled = True
'cmdImprimir.Enabled = True
'End Sub
''By Capi 15012008 comentado porque ahora solo se activara o desactivara
''Private Sub cmdEliminar_Click()
''Dim nFila As Long
''nFila = grdTasas.Row
''If grdTasas.TextMatrix(nFila, 7) = "" Then
''    grdTasas.EliminaFila nFila
''Else
''    MsgBox "No es posible eliminar una tasa ya registrada en el Tarifario", vbInformation, "Aviso"
''End If
''End Sub
'
'Private Sub CmdGrabar_Click()
'If Trim(grdTasas.TextMatrix(1, 1)) = "" Then
'    MsgBox "Ingrese las Tasas", vbInformation, "Aviso"
'    Exit Sub
'End If
'If Not ValidaTasas() Then Exit Sub
'
'If MsgBox("¿Desea grabar la información actualizada?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
'    Exit Sub
'End If
'
'Dim clsTasa As COMNCaptaGenerales.NCOMCaptaDefinicion
'Dim i As Integer
'Dim nCodTasa As Long
'Dim nMontoIni As Double, nMontoFin As Double
'Dim nPlazoIni As Long, nPlazoFin As Long
'Dim bOrdPag As Boolean
''By Capi 15012008
'Dim bActiva As Boolean
'Dim oMov As COMDMov.DCOMMov
'
'
'Dim VCMovNro As String
''
'Dim nValorTasa As Double
'Dim nUltFila As Long
'
'nUltFila = grdTasas.Rows - 1
'
'Set clsTasa = New COMNCaptaGenerales.NCOMCaptaDefinicion
'For i = 1 To grdTasas.Rows - 1
'    If CDbl(grdTasas.TextMatrix(nUltFila, 6)) = 0 Then
'        MsgBox "No es posible agregar un nuevo registro si no completa los datos anteriores", vbInformation, "Aviso"
'        grdTasas.Col = 6
'        grdTasas.row = nUltFila
'        grdTasas.SetFocus
'        Exit Sub
'    Else
'        If grdTasas.TextMatrix(i, 10) <> "" Or grdTasas.TextMatrix(i, 8) <> "." Then
'            nMontoIni = grdTasas.TextMatrix(i, 1)
'            nMontoFin = grdTasas.TextMatrix(i, 2)
'            nPlazoIni = grdTasas.TextMatrix(i, 3)
'            nPlazoFin = grdTasas.TextMatrix(i, 4)
'            bOrdPag = IIf(grdTasas.TextMatrix(i, 5) = "Si", True, False)
'            'By Capi 15012008
'            bActiva = IIf(grdTasas.TextMatrix(i, 8) = ".", True, False)
'            Set oMov = New COMDMov.DCOMMov
'            VCMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'            '
'            nValorTasa = CDbl(grdTasas.TextMatrix(i, 6))
'            If grdTasas.TextMatrix(i, 10) = "M" Then
'                nCodTasa = grdTasas.TextMatrix(i, 7)
'                'By Capi 15012008
'                clsTasa.ActualizaTasaPF nCodTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, nValorTasa, "", False
'                'If grdTasas.TextMatrix(i, 8) = "." Then
'                'clsTasa.ActualizaTasa nCodTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, nValorTasa, vcMovNro, True
'                clsTasa.NuevaTasaPF nProducto, nmoneda, nTipoTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, gsCodAge, nValorTasa, nTpoPrograma, VCMovNro, True, nCodTasa
'                'End If
'                '
'            ElseIf grdTasas.TextMatrix(i, 10) = "A" Then
'                nCodTasa = grdTasas.TextMatrix(i, 7)
'                'By Capi 15012008
'                clsTasa.ActualizaTasaPF nCodTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, nValorTasa, "", False
'                'If grdTasas.TextMatrix(i, 8) = "." Then
'                'clsTasa.ActualizaTasa nCodTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, nValorTasa, vcMovNro, True
'                If grdTasas.TextMatrix(i, 9) = "." Then
'                    clsTasa.NuevaTasaPF nProducto, nmoneda, nTipoTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, gsCodAge, nValorTasa, nTpoPrograma, VCMovNro, True, nCodTasa
'                Else
'                    clsTasa.NuevaTasaPF nProducto, nmoneda, nTipoTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, gsCodAge, nValorTasa, nTpoPrograma, VCMovNro, False, nCodTasa
'                End If
'                'End If
'
'            ElseIf grdTasas.TextMatrix(i, 10) = "N" Then
'                'By Capi 15012008 se agrego 2 nuevos parametros
'                clsTasa.NuevaTasaPF nProducto, nmoneda, nTipoTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, gsCodAge, nValorTasa, nTpoPrograma, VCMovNro, True, nCodTasa
'
'            End If
'
'            'By Capi 21012009
'            If grdTasas.TextMatrix(i, 10) = "M" Or grdTasas.TextMatrix(i, 10) = "A" Or grdTasas.TextMatrix(i, 10) = "N" Then
'                objPista.InsertarPista gsOpeCod, VCMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar
'            End If
'
'        End If
'    End If
'Next i
'
'Set clsTasa = Nothing
'cmdNuevaBus_Click
'End Sub
'
'Private Sub cmdImprimir_Click()
'    Dim P As previo.clsprevio
'    Dim sCad As String
'    Dim rs As New ADODB.Recordset
'    Dim LsCapImp As COMNCaptaGenerales.NCOMCaptaImpresion
'    Dim i As Integer
'    With rs
'    'Crear RecordSet
'        .Fields.Append "nMontoI", adCurrency
'        .Fields.Append "nMontoF", adCurrency
'        .Fields.Append "nPlazoI", adCurrency
'        .Fields.Append "nPlazoF", adCurrency
'        .Fields.Append "sOrdPago", adVarChar, 100
'        .Fields.Append "sTasaI", adVarChar, 100
'        .Fields.Append "sActiva", adVarChar, 100
'        .Fields.Append "LogCambio", adVarChar, 100
'        .Open
'    'Llenar Recordset
'        For i = 1 To grdTasas.Rows - 1
'            .AddNew
'            .Fields("nMontoI") = CDbl(Me.grdTasas.TextMatrix(i, 1))
'            .Fields("nMontoF") = CDbl(Me.grdTasas.TextMatrix(i, 2))
'            .Fields("nPlazoI") = CDbl(Me.grdTasas.TextMatrix(i, 3))
'            .Fields("nPlazoF") = CDbl(Me.grdTasas.TextMatrix(i, 4))
'            .Fields("sOrdPago") = IIf(Me.grdTasas.TextMatrix(i, 5) = ".", "SI", "NO")
'            .Fields("sTasaI") = Me.grdTasas.TextMatrix(i, 6)
'            .Fields("sActiva") = IIf(Me.grdTasas.TextMatrix(i, 9) = ".", "SI", "NO")
'            .Fields("LogCambio") = Me.grdTasas.TextMatrix(i, 8)
'        Next i
'    End With
'
'    Set LsCapImp = New COMNCaptaGenerales.NCOMCaptaImpresion
'        sCad = LsCapImp.ImprimirTasaInt(rs, sTitProd, gMonedaNacional, gsNomAge, gdFecSis, gsNomCmac, nmoneda)
'    Set LsCapImp = Nothing
'
'    Set P = New previo.clsprevio
'        P.Show sCad, "TASAS DE INTERES", False, , gImpresora
'    Set P = Nothing
'End Sub
'
'Private Sub cmdNuevaBus_Click()
'cmdAplicar.Enabled = True
'cboMoneda.Enabled = True
'cboTipoTasa.Enabled = True
'cboTpoPrograma.Enabled = True
'cmdNuevaBus.Enabled = False
'cboMoneda.SetFocus
'If Not bConsulta Then
'    'cmdNuevo.Enabled = False
'    'cmdEliminar.Enabled = False
'End If
'cmdImprimir.Enabled = False
'grdTasas.Clear
'grdTasas.Rows = 2
'grdTasas.FormaCabecera
'End Sub
'
'Private Sub cmdNuevo_Click()
'Dim nUltFila As Long
'nUltFila = grdTasas.Rows - 1
'If grdTasas.TextMatrix(nUltFila, 1) <> "" Then
'    If CDbl(grdTasas.TextMatrix(nUltFila, 6)) = 0 Then
'        MsgBox "No es posible agregar un nuevo registro si no completa los datos anteriores", vbInformation, "Aviso"
'        grdTasas.Col = 6
'        grdTasas.row = nUltFila
'        grdTasas.SetFocus
'        Exit Sub
'    End If
'End If
'grdTasas.AdicionaFila
'nUltFila = grdTasas.Rows - 1
'grdTasas.lbEditarFlex = True
'grdTasas.TextMatrix(nUltFila, 1) = "0.00"
'grdTasas.TextMatrix(nUltFila, 2) = "0.00"
'grdTasas.TextMatrix(nUltFila, 3) = "0"
'grdTasas.TextMatrix(nUltFila, 4) = "0"
'grdTasas.TextMatrix(nUltFila, 6) = "0.00"
'grdTasas.SetFocus
'cmdGrabar.Enabled = True
''cmdEliminar.Enabled = True
'End Sub
'
'Private Sub cmdsalir_Click()
'    Unload Me
'End Sub
'
'
'
'Private Sub Form_Load()
'    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
'    grdTasas.ColWidth(10) = 0
'
'    'By Capi 20012009
'    Set objPista = New COMManejador.Pista
'
'    'End By
'
'End Sub
'
'Private Sub grdTasas_OnCellChange(pnRow As Long, pnCol As Long)
'If grdTasas.TextMatrix(pnRow, 10) = "" Then grdTasas.TextMatrix(pnRow, 10) = "M"
'If pnCol = 6 Then
'    If grdTasas.TextMatrix(pnRow, pnCol) <> "" Then
'        If CDbl(grdTasas.TextMatrix(pnRow, pnCol)) < 0 Then
'            grdTasas.TextMatrix(pnRow, pnCol) = "0.00"
'        End If
'    End If
'End If
'cmdGrabar.Enabled = True
'End Sub
'
'Private Sub grdTasas_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'If grdTasas.TextMatrix(pnRow, 10) = "" Then grdTasas.TextMatrix(pnRow, 10) = "A"
'cmdGrabar.Enabled = True
'End Sub
'
'Private Sub grdTasas_OnRowAdd(pnRow As Long)
'grdTasas.TextMatrix(pnRow, 10) = "N"
'cmdGrabar.Enabled = True
'End Sub
