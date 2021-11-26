VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPRangoMoraPerm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Rangos de Mora Permitidos"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   Icon            =   "frmCredBPPRangoMoraPerm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Rangos de Mora"
      TabPicture(0)   =   "frmCredBPPRangoMoraPerm.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feCategorias"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGuardar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Factor de Rendimiento"
      TabPicture(1)   =   "frmCredBPPRangoMoraPerm.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdGuardarFacRend"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdCancelarFacRend"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "flxCategorias"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame2 
         Caption         =   "Filtro de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   6735
         Begin VB.ComboBox cmbNiveles 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1080
            Width           =   1455
         End
         Begin VB.ComboBox cmbAgencias 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox cmbMeses 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   5520
            TabIndex        =   13
            Top             =   480
            Width           =   1095
         End
         Begin SICMACT.uSpinner uspAnio 
            Height          =   315
            Left            =   2040
            TabIndex        =   14
            Top             =   480
            Width           =   975
            _extentx        =   1720
            _extenty        =   556
            max             =   9999
            min             =   1900
            maxlength       =   4
            min             =   1900
            font            =   "frmCredBPPRangoMoraPerm.frx":0342
            fontname        =   "MS Sans Serif"
            fontsize        =   8.25
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Agencia :"
            Height          =   195
            Left            =   3120
            TabIndex        =   18
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label2 
            Caption         =   "Mes - Año :"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filtro de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   -74760
         TabIndex        =   7
         Top             =   360
         Width           =   6735
         Begin VB.ComboBox cmbAgenciasFR 
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox cmbMesesFacRend 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton cmdMostrarFacRend 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   5640
            TabIndex        =   8
            Top             =   480
            Width           =   975
         End
         Begin SICMACT.uSpinner uspAnioFacRend 
            Height          =   315
            Left            =   2040
            TabIndex        =   10
            Top             =   480
            Width           =   975
            _extentx        =   1720
            _extenty        =   556
            max             =   9999
            min             =   1900
            maxlength       =   4
            min             =   1900
            font            =   "frmCredBPPRangoMoraPerm.frx":036E
            fontname        =   "MS Sans Serif"
            fontsize        =   8.25
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agencia :"
            Height          =   195
            Left            =   3240
            TabIndex        =   20
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label4 
            Caption         =   "Mes - Año :"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdGuardarFacRend 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   -70680
         TabIndex        =   6
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarFacRend 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   -69360
         TabIndex        =   5
         Top             =   4320
         Width           =   1335
      End
      Begin SICMACT.FlexEdit flxCategorias 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   4
         Top             =   1440
         Width           =   6735
         _extentx        =   11880
         _extenty        =   3625
         cols0           =   6
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "-Categoria-Nivel I(%)-Nivel II(%)-Nivel III(%)-Aux"
         encabezadosanchos=   "0-1600-1600-1600-1600-0"
         font            =   "frmCredBPPRangoMoraPerm.frx":039A
         font            =   "frmCredBPPRangoMoraPerm.frx":03C6
         font            =   "frmCredBPPRangoMoraPerm.frx":03F2
         font            =   "frmCredBPPRangoMoraPerm.frx":041E
         font            =   "frmCredBPPRangoMoraPerm.frx":044A
         fontfixed       =   "frmCredBPPRangoMoraPerm.frx":0476
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-2-3-4-X"
         listacontroles  =   "0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-R-R-R-L"
         formatosedit    =   "0-0-2-2-2-0"
         cantentero      =   7
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   4320
         Width           =   1335
      End
      Begin SICMACT.FlexEdit feCategorias 
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   6735
         _extentx        =   11880
         _extenty        =   3625
         cols0           =   5
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "-Categoria-(%) Mora 8 a 30-(%) Mora mayor 30-cCodAge"
         encabezadosanchos=   "0-2400-2000-2000-0"
         font            =   "frmCredBPPRangoMoraPerm.frx":04A4
         font            =   "frmCredBPPRangoMoraPerm.frx":04D0
         font            =   "frmCredBPPRangoMoraPerm.frx":04FC
         font            =   "frmCredBPPRangoMoraPerm.frx":0528
         font            =   "frmCredBPPRangoMoraPerm.frx":0554
         fontfixed       =   "frmCredBPPRangoMoraPerm.frx":0580
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-2-3-X"
         listacontroles  =   "0-0-0-0-0"
         encabezadosalineacion=   "C-C-R-R-C"
         formatosedit    =   "0-0-2-2-0"
         cantentero      =   7
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCredBPPRangoMoraPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmbNiveles_Click()
'If ValidaDatos(0) Then
'    CargaGridRangosMora Right(cmbMeses.Text, 2), uspAnio.valor, Trim(Right(cmbAgencias.Text, 4)), CInt(Trim(Right(cmbNiveles.Text, 5)))
'End If
'End Sub
'
'Private Sub cmdCancelar_Click()
'    LimpiaFlex feCategorias
'End Sub
'
'Private Sub cmdCancelarFacRend_Click()
'    LimpiaFlex flxCategorias
'    CargaGridCategorias
'End Sub
'
'Private Sub CmdGuardar_Click()
'On Error GoTo Error
'If ValidaDatos(0, 1) Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Dim lsCodAge As String
'    Dim lsCategoria As String
'    Dim lnMes As Integer
'    Dim lnAnio As Integer
'    Dim lnNivel As Integer
'    Dim lsFecha As String
'    Dim lMora8a30 As Double
'    Dim lMoraMaya30 As Double
'
'        lnMes = Right(cmbMeses.Text, 2)
'        lnAnio = uspAnio.valor
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'        lsCodAge = Trim(Right(cmbAgencias.Text, 4))
'        lnNivel = CInt(Trim(Right(cmbNiveles.Text, 4)))
'        lnFila = feCategorias.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        Call oBPP.EliminaRangosMora(lnMes, lnAnio, lsCodAge, lnNivel)
'
'        For i = 1 To lnFila
'            lMora8a30 = feCategorias.TextMatrix(i, 2)
'            lMoraMaya30 = feCategorias.TextMatrix(i, 3)
'            lsCategoria = Trim(feCategorias.TextMatrix(i, 1))
'
'            oBPP.InsertaRangosdeMora lnMes, lnAnio, lsCodAge, lMora8a30, lMoraMaya30, gsCodUser, lsFecha, lnNivel, lsCategoria
'
'        Next
'        MsgBox "Se registraron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarFacRend_Click()
'On Error GoTo Error
'If ValidaDatos(1, 1) Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Dim lsCategoria As String
'    Dim lnMes As Integer
'    Dim lnAnio As Integer
'    Dim lsFecha As String
'    Dim lsCodAge As String
'    Dim lnNivel1 As Double
'    Dim lnNivel2 As Double
'    Dim lnNivel3 As Double
'
'        lnMes = Right(cmbMesesFacRend.Text, 2)
'        lnAnio = uspAnioFacRend.valor
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'        lsCodAge = Trim(Right(cmbAgenciasFR.Text, 4))
'        lnFila = flxCategorias.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        Call oBPP.EliminaFactorRend(lnMes, lnAnio, lsCodAge)
'
'        For i = 1 To lnFila
'            lsCategoria = flxCategorias.TextMatrix(i, 1)
'            lnNivel1 = flxCategorias.TextMatrix(i, 2)
'            lnNivel2 = flxCategorias.TextMatrix(i, 3)
'            lnNivel3 = flxCategorias.TextMatrix(i, 4)
'
'            oBPP.InsertaFactorRend lnMes, lnAnio, lsCategoria, lnNivel1, lnNivel2, lnNivel3, gsCodUser, lsFecha, lsCodAge
'
'        Next
'        MsgBox "Se registraron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdMostrar_Click()
'If ValidaDatos(0) Then
'    CargaGridRangosMora Right(cmbMeses.Text, 2), uspAnio.valor, Trim(Right(cmbAgencias.Text, 4)), CInt(Trim(Right(cmbNiveles.Text, 5)))
'End If
'End Sub
'
'Private Sub cmdMostrarFacRend_Click()
'If ValidaDatos(1) Then
'    CargaFactorRendimiento Right(cmbMesesFacRend.Text, 2), uspAnioFacRend.valor, Trim(Right(cmbAgenciasFR.Text, 4))
'End If
'End Sub
'
'Private Sub Form_Load()
'    CargaComboMeses cmbMeses
'    CargaComboMeses cmbMesesFacRend
'
'    CargaComboAgencias cmbAgencias
'    CargaComboAgencias cmbAgenciasFR
'
'    'CargaGridAgencias
'    CargaGridCategorias
'    uspAnio.valor = Year(gdFecSis)
'    uspAnioFacRend.valor = Year(gdFecSis)
'
'    Dim oConst As COMDConstantes.DCOMConstantes
'    Dim rsConst As ADODB.Recordset
'
'    'CARGA NIVELES
'    Set oConst = New COMDConstantes.DCOMConstantes
'    Set rsConst = oConst.RecuperaConstantes(7064)
'
'    'CARGA COMBO DE NIVELES
'    Call Llenar_Combo_con_Recordset(rsConst, cmbNiveles)
'End Sub
'
'Private Sub CargaGridAgencias()
'Dim oConst As COMDConstantes.DCOMAgencias
'Dim R As ADODB.Recordset
'Dim i As Integer
'    On Error GoTo ERRORCargaGridAgencias
'
'    Set oConst = New COMDConstantes.DCOMAgencias
'    Set R = oConst.ObtieneAgencias()
'    Set oConst = Nothing
'    i = 1
'    Do While Not R.EOF
'        flxAgenciasMora.AdicionaFila
'        flxAgenciasMora.TextMatrix(i, 1) = R!cConsDescripcion
'        flxAgenciasMora.TextMatrix(i, 4) = R!nConsValor
'        R.MoveNext
'        i = i + 1
'    Loop
'    R.Close
'    Set R = Nothing
'    flxAgenciasMora.TopRow = 1
'    Exit Sub
'
'ERRORCargaGridAgencias:
'    MsgBox err.Description, vbCritical, "Aviso"
'End Sub
'Private Sub CargaCategoriaMora(ByVal pnNivel As Integer)
'Dim oDBPP As COMDCredito.DCOMBPPR
'LimpiaFlex feCategorias
'Set oDBPP = New COMDCredito.DCOMBPPR
'Set rsDBPP = oDBPP.ObtenerCatOParamCabXNivel(1, pnNivel)
'
'Set oDBPPDatos = New COMDCredito.DCOMBPPR
'
'If Not (rsDBPP.BOF And rsDBPP.EOF) Then
'    For i = 0 To rsDBPP.RecordCount - 1
'        feCategorias.AdicionaFila
'        feCategorias.TextMatrix(i + 1, 0) = i + 1
'        feCategorias.TextMatrix(i + 1, 1) = Trim(rsDBPP!cCategoria)
'
'        rsDBPP.MoveNext
'    Next i
'Else
'    MsgBox "No Hay Datos.", vbInformation, "Aviso"
'End If
'
'End Sub
'
'
'Private Sub CargaGridRangosMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnNivel As Integer)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverRangosMora(pnMes, pnAnio, psCodAge, pnNivel)
'
'    LimpiaFlex feCategorias
'
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            feCategorias.AdicionaFila
'
'            feCategorias.TextMatrix(i, 1) = rs!cCategoria
'            feCategorias.TextMatrix(i, 2) = Format(rs!nMora8a30, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feCategorias.TextMatrix(i, 3) = Format(rs!nMoraMayora30, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feCategorias.TextMatrix(i, 4) = rs!cCategoria
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    Else
'        Call CargaCategoriaMora(pnNivel)
'    End If
'     feCategorias.TopRow = 1
'    Set rs = Nothing
'    feCategorias.lbEditarFlex = True
'End Sub
'
'Private Sub CargaFactorRendimiento(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverFactorRend(pnMes, pnAnio, psCodAge)
'
'    LimpiaFlex flxCategorias
'
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            flxCategorias.AdicionaFila
'
'            flxCategorias.TextMatrix(i, 1) = rs!cCategoria
'            flxCategorias.TextMatrix(i, 2) = Format(rs!nNivel1, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxCategorias.TextMatrix(i, 3) = Format(rs!nNivel2, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxCategorias.TextMatrix(i, 4) = Format(rs!nNivel3, "###," & String(15, "#") & "#0." & String(2, "0"))
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    Else
'        CargaGridCategorias
'    End If
'
'    flxCategorias.TopRow = 1
'    Set rs = Nothing
'    flxCategorias.lbEditarFlex = True
'End Sub
'
'Private Sub CargaGridCategorias()
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverCategorias()
'
'    LimpiaFlex flxCategorias
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            flxCategorias.AdicionaFila
'
'            flxCategorias.TextMatrix(i, 1) = rs!cCategoria
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    End If
'
'    Set rs = Nothing
'End Sub
'
'Private Function ValidaDatos(ByVal nTpoFrame As Integer, Optional pnTipo As Integer = 0) As Boolean
'If nTpoFrame = 0 Then
'       If Trim(cmbMeses.Text) = "" Then
'            MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'        If Trim(uspAnio.valor) = "" Or CDbl(uspAnio.valor) = 0 Then
'            MsgBox "Ingrese el Año", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'        If Trim(cmbAgencias.Text) = "" Then
'            MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'        If Trim(cmbNiveles.Text) = "" Then
'            MsgBox "Seleccione el Nivel", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'    If pnTipo = 1 Then
'        For i = 0 To feCategorias.Rows - 2
'            If feCategorias.TextMatrix(i + 1, 1) = "" Or feCategorias.TextMatrix(i + 1, 2) = "" Or feCategorias.TextMatrix(i + 1, 3) = "" Then
'                MsgBox "Ingrese los datos Correctamente", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If IsNumeric(feCategorias.TextMatrix(i + 1, 2)) Then
'                If CDbl(feCategorias.TextMatrix(i + 1, 2)) < 0 Or CDbl(feCategorias.TextMatrix(i + 1, 2)) > 100 Then
'                    MsgBox "Ingrese correctamente los valores del (%)Mora 8 -30 (0.00% -100.00%) de la categoria ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'            If IsNumeric(feCategorias.TextMatrix(i + 1, 3)) Then
'                If CDbl(feCategorias.TextMatrix(i + 1, 3)) < 0 Or CDbl(feCategorias.TextMatrix(i + 1, 3)) > 100 Then
'                    MsgBox "Ingrese correctamente los valores del (%)Mora > 30 (0.00% -100.00%) de la categoria ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'        Next i
'    End If
'ElseIf nTpoFrame = 1 Then
'        If Trim(cmbMesesFacRend.Text) = "" Then
'            MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'        If Trim(uspAnioFacRend.valor) = "" Or CDbl(uspAnio.valor) = 0 Then
'            MsgBox "Ingrese el Año", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'        If Trim(cmbAgenciasFR.Text) = "" Then
'            MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'    If pnTipo = 1 Then
'        For i = 0 To flxCategorias.Rows - 2
'            If flxCategorias.TextMatrix(i + 1, 1) = "" Or flxCategorias.TextMatrix(i + 1, 2) = "" Or flxCategorias.TextMatrix(i + 1, 3) = "" Then
'                MsgBox "Ingrese los datos Correctamente", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'
'            If IsNumeric(flxCategorias.TextMatrix(i + 1, 2)) Then
'                If CDbl(flxCategorias.TextMatrix(i + 1, 2)) < 0 Or CDbl(flxCategorias.TextMatrix(i + 1, 2)) > 100 Then
'                    MsgBox "Ingrese correctamente los valores del Nivel I (0.00% -100.00%) de la categoria ''" & Trim(flxCategorias.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'            If IsNumeric(flxCategorias.TextMatrix(i + 1, 3)) Then
'                If CDbl(flxCategorias.TextMatrix(i + 1, 3)) < 0 Or CDbl(flxCategorias.TextMatrix(i + 1, 3)) > 100 Then
'                    MsgBox "Ingrese correctamente los valores del  Nivel II (0.00% -100.00%) de la categoria ''" & Trim(flxCategorias.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'            If IsNumeric(flxCategorias.TextMatrix(i + 1, 4)) Then
'                If CDbl(flxCategorias.TextMatrix(i + 1, 4)) < 0 Or CDbl(flxCategorias.TextMatrix(i + 1, 4)) > 100 Then
'                    MsgBox "Ingrese correctamente los valores del  Nivel III (0.00% -100.00%) de la categoria ''" & Trim(flxCategorias.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'        Next i
'    End If
'End If
'
'ValidaDatos = True
'End Function
