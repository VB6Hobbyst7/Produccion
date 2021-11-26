VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredBPPParamCumGeren 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Parámetros de Cumplimiento - Gerencia"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   Icon            =   "frmCredBPPParamCumGeren.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Parámetros de cumplimiento"
      TabPicture(0)   =   "frmCredBPPParamCumGeren.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSTab2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbMeses"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "uspAnio"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbAgencias"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdMostrar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   7320
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbAgencias 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   870
         Width           =   2175
      End
      Begin Spinner.uSpinner uspAnio 
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Top             =   870
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Max             =   9999
         Min             =   1900
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.ComboBox cmbMeses 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   870
         Width           =   1815
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5953
         _Version        =   393216
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "Nivel I"
         TabPicture(0)   =   "frmCredBPPParamCumGeren.frx":0326
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "flxNivelI"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdCancelarNivelI"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdGuardarNivelI"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Nivel II"
         TabPicture(1)   =   "frmCredBPPParamCumGeren.frx":0342
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdCancelarNivelII"
         Tab(1).Control(1)=   "cmdGuardarNivelII"
         Tab(1).Control(2)=   "flxNivelII"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Nivel III"
         TabPicture(2)   =   "frmCredBPPParamCumGeren.frx":035E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdCancelarNivelIII"
         Tab(2).Control(1)=   "cmdGuardarNivelIII"
         Tab(2).Control(2)=   "flxNivelIII"
         Tab(2).ControlCount=   3
         Begin VB.CommandButton cmdCancelarNivelIII 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   -67920
            TabIndex        =   16
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdGuardarNivelIII 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   -69120
            TabIndex        =   15
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelarNivelII 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   -67920
            TabIndex        =   13
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdGuardarNivelII 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   -69120
            TabIndex        =   12
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdGuardarNivelI 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   5880
            TabIndex        =   11
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelarNivelI 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   7080
            TabIndex        =   10
            Top             =   2760
            Width           =   1095
         End
         Begin SICMACT.FlexEdit flxNivelI 
            Height          =   2055
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   7935
            _extentx        =   13996
            _extenty        =   3625
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "-Factores-D-C-B-A-Cod Factor"
            encabezadosanchos=   "0-3000-1200-1200-1200-1200-0"
            font            =   "frmCredBPPParamCumGeren.frx":037A
            font            =   "frmCredBPPParamCumGeren.frx":03A6
            font            =   "frmCredBPPParamCumGeren.frx":03D2
            font            =   "frmCredBPPParamCumGeren.frx":03FE
            font            =   "frmCredBPPParamCumGeren.frx":042A
            fontfixed       =   "frmCredBPPParamCumGeren.frx":0456
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-2-3-4-5-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-2-2-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit flxNivelII 
            Height          =   2055
            Left            =   -74760
            TabIndex        =   14
            Top             =   480
            Width           =   7935
            _extentx        =   13996
            _extenty        =   3625
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "-Factores-D-C-B-A-cod factor"
            encabezadosanchos=   "0-3000-1200-1200-1200-1200-0"
            font            =   "frmCredBPPParamCumGeren.frx":0484
            font            =   "frmCredBPPParamCumGeren.frx":04B0
            font            =   "frmCredBPPParamCumGeren.frx":04DC
            font            =   "frmCredBPPParamCumGeren.frx":0508
            font            =   "frmCredBPPParamCumGeren.frx":0534
            fontfixed       =   "frmCredBPPParamCumGeren.frx":0560
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-2-3-4-5-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-2-2-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit flxNivelIII 
            Height          =   2055
            Left            =   -74760
            TabIndex        =   17
            Top             =   480
            Width           =   7935
            _extentx        =   13996
            _extenty        =   3625
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "-Factores-D-C-B-A-cod factor"
            encabezadosanchos=   "0-3000-1200-1200-1200-1200-0"
            font            =   "frmCredBPPParamCumGeren.frx":058E
            font            =   "frmCredBPPParamCumGeren.frx":05BA
            font            =   "frmCredBPPParamCumGeren.frx":05E6
            font            =   "frmCredBPPParamCumGeren.frx":0612
            font            =   "frmCredBPPParamCumGeren.frx":063E
            fontfixed       =   "frmCredBPPParamCumGeren.frx":066A
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-2-3-4-5-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-2-2-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         Height          =   195
         Left            =   4200
         TabIndex        =   6
         Top             =   930
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes - Año :"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   930
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Filtro de Registro"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   540
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCredBPPParamCumGeren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim lnNivel As Integer
'Dim lnMes As Integer
'Dim lnAnio As Integer
'Dim lsCodAge As String
'Dim lsCodFact As String
'Dim lnCatA As Double
'Dim lnCatB As Double
'Dim lnCatC As Double
'Dim lnCatD As Double
'Dim lsFecha As String
'Dim i As Integer, j As Integer
'Dim lnFila As Integer
'Dim lnTipOpe As Integer
'
'Public Sub Inicia(ByVal pnTipOpe As Integer)
'    lnTipOpe = pnTipOpe
'    flxNivelI.EncabezadosNombres = "-Factores-D-C-B-A-cod factor"
'    flxNivelII.EncabezadosNombres = "-Factores-D-C-B-A-cod factor"
'    flxNivelIII.EncabezadosNombres = "-Factores-D-C-B-A-cod factor"
'
'    If lnTipOpe = 2 Then
'        Me.Caption = "BPP - Porcentaje de Cumplimiento - Gerencia"
'        flxNivelI.EncabezadosNombres = "-Factores-D(%)-C(%)-B(%)-A(%)-cod factor"
'        flxNivelII.EncabezadosNombres = "-Factores-D(%)-C(%)-B(%)-A(%)-cod factor"
'        flxNivelIII.EncabezadosNombres = "-Factores-D(%)-C(%)-B(%)-A(%)-cod factor"
'    End If
'
'    CargaCombos
'    CargaFactoresNivelI
'    CargaFactoresNivelII
'    CargaFactoresNivelIII
'
'    uspAnio.valor = Year(gdFecSis)
'
'    Me.Show 1
'End Sub
'
'Private Sub CargaCombos()
'    CargaComboAgencias cmbAgencias
'    CargaComboMeses cmbMeses
'End Sub
'
'Private Sub CargaFactoresNivelI()
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverFactoresCumplimiento(1, lnTipOpe)
'
'    i = 1
'
'    Do While Not rs.EOF
'        flxNivelI.AdicionaFila
'
'        flxNivelI.TextMatrix(i, 1) = rs!cDescFactor
'        flxNivelI.TextMatrix(i, 6) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'
'    Set rs = Nothing
'    flxNivelI.TopRow = 1
'End Sub
'
'Private Sub CargaFactoresNivelII()
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverFactoresCumplimiento(2, lnTipOpe)
'
'    i = 1
'
'    Do While Not rs.EOF
'        flxNivelII.AdicionaFila
'
'        flxNivelII.TextMatrix(i, 1) = rs!cDescFactor
'        flxNivelII.TextMatrix(i, 6) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'    flxNivelII.TopRow = 1
'    Set rs = Nothing
'End Sub
'
'Private Sub CargaFactoresNivelIII()
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverFactoresCumplimiento(3, lnTipOpe)
'    i = 1
'
'    Do While Not rs.EOF
'        flxNivelIII.AdicionaFila
'
'        flxNivelIII.TextMatrix(i, 1) = rs!cDescFactor
'        flxNivelIII.TextMatrix(i, 6) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'    flxNivelIII.TopRow = 1
'End Sub
'
'Private Sub cmdCancelarNivelI_Click()
'    LimpiaFlex flxNivelI
'    CargaFactoresNivelI
'End Sub
'
'Private Sub cmdCancelarNivelII_Click()
'    LimpiaFlex flxNivelII
'    CargaFactoresNivelII
'End Sub
'
'Private Sub cmdCancelarNivelIII_Click()
'    LimpiaFlex flxNivelIII
'    CargaFactoresNivelIII
'End Sub
'
'Private Sub cmdGuardarNivelI_Click()
'On Error GoTo Error
'If ValidaDatos(1) Then
'    Dim oBPP As COMNCredito.NCOMBPPR
'        If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'        lnNivel = 1
'        lnMes = Right(cmbMeses.Text, 2)
'        lnAnio = uspAnio.valor
'        lsCodAge = Right(cmbAgencias.Text, 2)
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        lnFila = flxNivelI.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'
'        Call oBPP.EliminaParametrosCumplimiento(lnNivel, lnMes, lnAnio, lsCodAge, lnTipOpe)
'
'        For i = 1 To lnFila
'            lsCodFact = flxNivelI.TextMatrix(i, 6)
'            lnCatA = flxNivelI.TextMatrix(i, 5)
'            lnCatB = flxNivelI.TextMatrix(i, 4)
'            lnCatC = flxNivelI.TextMatrix(i, 3)
'            lnCatD = flxNivelI.TextMatrix(i, 2)
'
'            oBPP.InsertaParametroCumplimiento lnNivel, lsCodFact, lnCatA, lnCatB, lnCatC, lnCatD, lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha, lnTipOpe
'
'        Next
'        MsgBox "Se guardaron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarNivelII_Click()
'On Error GoTo Error
'If ValidaDatos(2) Then
'    Dim oBPP As COMNCredito.NCOMBPPR
'        If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'        lnNivel = 2
'        lnMes = Right(cmbMeses.Text, 2)
'        lnAnio = uspAnio.valor
'        lsCodAge = Right(cmbAgencias.Text, 2)
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        lnFila = flxNivelII.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        Call oBPP.EliminaParametrosCumplimiento(lnNivel, lnMes, lnAnio, lsCodAge, lnTipOpe)
'        For i = 1 To lnFila
'            lsCodFact = flxNivelII.TextMatrix(i, 6)
'            lnCatA = flxNivelII.TextMatrix(i, 5)
'            lnCatB = flxNivelII.TextMatrix(i, 4)
'            lnCatC = flxNivelII.TextMatrix(i, 3)
'            lnCatD = flxNivelII.TextMatrix(i, 2)
'
'            oBPP.InsertaParametroCumplimiento lnNivel, lsCodFact, lnCatA, lnCatB, lnCatC, lnCatD, lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha, lnTipOpe
'
'        Next
'        MsgBox "Se guardaron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarNivelIII_Click()
'On Error GoTo Error
'If ValidaDatos(3) Then
'    Dim oBPP As COMNCredito.NCOMBPPR
'        If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'
'        lnNivel = 3
'        lnMes = Right(cmbMeses.Text, 2)
'        lnAnio = uspAnio.valor
'        lsCodAge = Right(cmbAgencias.Text, 2)
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        lnFila = flxNivelIII.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        Call oBPP.EliminaParametrosCumplimiento(lnNivel, lnMes, lnAnio, lsCodAge, lnTipOpe)
'        For i = 1 To lnFila
'            lsCodFact = flxNivelIII.TextMatrix(i, 6)
'            lnCatA = flxNivelIII.TextMatrix(i, 5)
'            lnCatB = flxNivelIII.TextMatrix(i, 4)
'            lnCatC = flxNivelIII.TextMatrix(i, 3)
'            lnCatD = flxNivelIII.TextMatrix(i, 2)
'
'            oBPP.InsertaParametroCumplimiento lnNivel, lsCodFact, lnCatA, lnCatB, lnCatC, lnCatD, lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha, lnTipOpe
'
'        Next
'        MsgBox "Se guardaron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdMostrar_Click()
'If ValidaDatos Then
'    CargaGridNivelI 1, Right(cmbMeses.Text, 2), uspAnio.valor, Right(cmbAgencias.Text, 2)
'    CargaGridNivelII 2, Right(cmbMeses.Text, 2), uspAnio.valor, Right(cmbAgencias.Text, 2)
'    CargaGridNivelIII 3, Right(cmbMeses.Text, 2), uspAnio.valor, Right(cmbAgencias.Text, 2)
'End If
'End Sub
'
'Private Sub CargaGridNivelI(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverParametrosCumplimiento(pnNivel, pnMes, pnAnio, psCodAge, lnTipOpe)
'
'    LimpiaFlex flxNivelI
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            flxNivelI.AdicionaFila
'
'            flxNivelI.TextMatrix(i, 1) = rs!cDescFactor
'            flxNivelI.TextMatrix(i, 2) = Format(rs!nCategoriaD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 3) = Format(rs!nCategoriaC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 4) = Format(rs!nCategoriaB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 5) = Format(rs!nCategoriaA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 6) = rs!cCodFactor
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    Else
'        CargaFactoresNivelI
'    End If
'
'    Set rs = Nothing
'End Sub
'Private Sub CargaGridNivelII(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverParametrosCumplimiento(pnNivel, pnMes, pnAnio, psCodAge, lnTipOpe)
'
'    LimpiaFlex flxNivelII
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            flxNivelII.AdicionaFila
'
'            flxNivelII.TextMatrix(i, 1) = rs!cDescFactor
'            flxNivelII.TextMatrix(i, 2) = Format(rs!nCategoriaD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 3) = Format(rs!nCategoriaC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 4) = Format(rs!nCategoriaB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 5) = Format(rs!nCategoriaA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 6) = rs!cCodFactor
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    Else
'        CargaFactoresNivelII
'    End If
'
'    Set rs = Nothing
'End Sub
'Private Sub CargaGridNivelIII(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverParametrosCumplimiento(pnNivel, pnMes, pnAnio, psCodAge, lnTipOpe)
'
'    LimpiaFlex flxNivelIII
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            flxNivelIII.AdicionaFila
'
'            flxNivelIII.TextMatrix(i, 1) = rs!cDescFactor
'            flxNivelIII.TextMatrix(i, 2) = Format(rs!nCategoriaD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 3) = Format(rs!nCategoriaC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 4) = Format(rs!nCategoriaB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 5) = Format(rs!nCategoriaA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 6) = rs!cCodFactor
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    Else
'        CargaFactoresNivelIII
'    End If
'
'    Set rs = Nothing
'End Sub
'
'Private Function ValidaDatos(Optional ByVal pnNivel As Integer = 4) As Boolean
'
'   If Trim(cmbMeses.Text) = "" Then
'        MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'    If Trim(uspAnio.valor) = "" Or CDbl(uspAnio.valor) = 0 Then
'        MsgBox "Ingrese el Año", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'    If Trim(cmbAgencias.Text) = "" Then
'        MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'If pnNivel = 1 Then
'    For i = 0 To flxNivelI.Rows - 2
'        For j = 2 To 5
'
'            If Trim(flxNivelI.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If lnTipOpe = 2 Then
'                If IsNumeric(flxNivelI.TextMatrix(i + 1, j)) Then
'                    If CDbl(flxNivelI.TextMatrix(i + 1, j)) < 0 Or CDbl(flxNivelI.TextMatrix(i + 1, j)) > 100 Then
'                        MsgBox "Ingrese correctamente los valores (0.00% - 100.00%) del Factor ''" & Trim(flxNivelI.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                        ValidaDatos = False
'                        Exit Function
'                    End If
'                End If
'            Else
'                If IsNumeric(flxNivelI.TextMatrix(i + 1, j)) Then
'                    If CDbl(flxNivelI.TextMatrix(i + 1, j)) < 0 Then
'                        MsgBox "Ingrese correctamente los valores del Factor ''" & Trim(flxNivelI.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                        ValidaDatos = False
'                        Exit Function
'                    End If
'                End If
'            End If
'        Next j
'
'    Next i
'ElseIf pnNivel = 2 Then
'    For i = 0 To flxNivelII.Rows - 2
'
'        For j = 2 To 5
'
'            If Trim(flxNivelII.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If lnTipOpe = 2 Then
'                If IsNumeric(flxNivelII.TextMatrix(i + 1, j)) Then
'                    If CDbl(flxNivelII.TextMatrix(i + 1, j)) < 0 Or CDbl(flxNivelII.TextMatrix(i + 1, j)) > 100 Then
'                        MsgBox "Ingrese correctamente los valores (0.00% - 100.00%) del Factor ''" & Trim(flxNivelII.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                        ValidaDatos = False
'                        Exit Function
'                    End If
'                End If
'            Else
'                If IsNumeric(flxNivelII.TextMatrix(i + 1, j)) Then
'                    If CDbl(flxNivelII.TextMatrix(i + 1, j)) < 0 Then
'                        MsgBox "Ingrese correctamente los valores del Factor ''" & Trim(flxNivelII.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                        ValidaDatos = False
'                        Exit Function
'                    End If
'                End If
'            End If
'        Next j
'
'    Next i
'ElseIf pnNivel = 3 Then
'    For i = 0 To flxNivelIII.Rows - 2
'
'        For j = 2 To 5
'
'            If Trim(flxNivelIII.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If lnTipOpe = 2 Then
'                If IsNumeric(flxNivelIII.TextMatrix(i + 1, j)) Then
'                    If CDbl(flxNivelIII.TextMatrix(i + 1, j)) < 0 Or CDbl(flxNivelIII.TextMatrix(i + 1, j)) > 100 Then
'                        MsgBox "Ingrese correctamente los valores (0.00% - 100.00%) del Factor ''" & Trim(flxNivelIII.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                        ValidaDatos = False
'                        Exit Function
'                    End If
'                End If
'            Else
'                If IsNumeric(flxNivelIII.TextMatrix(i + 1, j)) Then
'                    If CDbl(flxNivelIII.TextMatrix(i + 1, j)) < 0 Then
'                        MsgBox "Ingrese correctamente los valores del Factor ''" & Trim(flxNivelIII.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                        ValidaDatos = False
'                        Exit Function
'                    End If
'                End If
'            End If
'        Next j
'
'
'    Next i
'End If
'
'ValidaDatos = True
'End Function
