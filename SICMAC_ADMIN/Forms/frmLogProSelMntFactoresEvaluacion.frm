VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogProSelMntFactoresEvaluacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de Factores de Evaluación por Proceso de Selección"
   ClientHeight    =   6870
   ClientLeft      =   1485
   ClientTop       =   1125
   ClientWidth     =   9810
   Icon            =   "frmLogProSelMntFactoresEvaluacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CboObj 
      Height          =   315
      ItemData        =   "frmLogProSelMntFactoresEvaluacion.frx":08CA
      Left            =   120
      List            =   "frmLogProSelMntFactoresEvaluacion.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ComboBox cboGrupoBS 
      Height          =   315
      ItemData        =   "frmLogProSelMntFactoresEvaluacion.frx":08CE
      Left            =   2760
      List            =   "frmLogProSelMntFactoresEvaluacion.frx":08D5
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2520
      Width           =   6975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483630
      Cols            =   4
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483624
      BackColorBkg    =   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Factores de Evaluacion"
      TabPicture(0)   =   "frmLogProSelMntFactoresEvaluacion.frx":08E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameFactores"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameLista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Rango de Valores"
      TabPicture(1)   =   "frmLogProSelMntFactoresEvaluacion.frx":08FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexValoresVer"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexValoresVer 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   5530
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         Cols            =   4
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   -2147483647
         ForeColorSel    =   -2147483624
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Frame FrameLista 
         BorderStyle     =   0  'None
         Height          =   3435
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   9435
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            Height          =   375
            Left            =   8040
            TabIndex        =   33
            Top             =   2940
            Width           =   1155
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   2940
            Width           =   1155
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   2940
            Width           =   1155
         End
         Begin VB.CommandButton cmdmodificar 
            Caption         =   "Modificar Formula y/o Puntaje"
            Height          =   375
            Left            =   2520
            TabIndex        =   6
            Top             =   2940
            Visible         =   0   'False
            Width           =   2475
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexFactores 
            Height          =   2415
            Left            =   0
            TabIndex        =   3
            Top             =   360
            Width           =   9315
            _ExtentX        =   16431
            _ExtentY        =   4260
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   -2147483630
            Cols            =   6
            FixedCols       =   0
            ForeColorFixed  =   -2147483646
            BackColorSel    =   -2147483647
            ForeColorSel    =   -2147483624
            BackColorBkg    =   16777215
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            GridColorUnpopulated=   -2147483633
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.Label Label2 
            Caption         =   "Factores de Evaluacion por Proceso, Objeto y Grupo de Bienes o Servicios"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   7365
         End
      End
      Begin VB.Frame frameFactores 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3435
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   9435
         Begin VB.TextBox txttipo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6900
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtunidades 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton cmdGrabarApelacion 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   5955
            TabIndex        =   21
            Top             =   3000
            Width           =   1275
         End
         Begin VB.TextBox txtpuntaje 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1455
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelarApelacion 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   7275
            TabIndex        =   19
            Top             =   3000
            Width           =   1275
         End
         Begin VB.ComboBox cboFactores 
            Height          =   315
            Left            =   1455
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   120
            Width           =   7215
         End
         Begin VB.ComboBox CboFormula 
            Height          =   315
            Left            =   1455
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   480
            Width           =   7215
         End
         Begin VB.Frame FrameValores 
            Caption         =   "Rangos de Valores"
            Height          =   1695
            Left            =   795
            TabIndex        =   15
            Top             =   1200
            Visible         =   0   'False
            Width           =   7815
            Begin VB.TextBox txtMayor 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6000
               MaxLength       =   10
               MultiLine       =   -1  'True
               TabIndex        =   34
               Top             =   240
               Width           =   1695
            End
            Begin VB.CommandButton CmdMenos 
               Caption         =   "Quitar"
               Height          =   255
               Left            =   480
               TabIndex        =   27
               Top             =   960
               Width           =   1035
            End
            Begin VB.CommandButton CmdMas 
               Caption         =   "Agregar"
               Height          =   255
               Left            =   480
               TabIndex        =   26
               Top             =   660
               Width           =   1035
            End
            Begin VB.TextBox TxtMenor 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2820
               MaxLength       =   10
               MultiLine       =   -1  'True
               TabIndex        =   25
               Top             =   285
               Width           =   1815
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexValores 
               Height          =   975
               Left            =   1920
               TabIndex        =   16
               Top             =   600
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   1720
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   -2147483630
               Cols            =   4
               FixedCols       =   0
               ForeColorFixed  =   -2147483646
               BackColorSel    =   -2147483647
               ForeColorSel    =   -2147483624
               BackColorBkg    =   16777215
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               GridColorUnpopulated=   -2147483633
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   6.75
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
               _NumberOfBands  =   1
               _Band(0).Cols   =   4
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "y Menor o Igual a "
               Height          =   195
               Left            =   4695
               TabIndex        =   35
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Mayor o Igual a "
               Height          =   195
               Left            =   1560
               TabIndex        =   28
               Top             =   360
               Width           =   1140
            End
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   6480
            TabIndex        =   32
            Top             =   900
            Width           =   315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Unidades"
            Height          =   195
            Left            =   2955
            TabIndex        =   30
            Top             =   900
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Puntaje"
            Height          =   195
            Left            =   795
            TabIndex        =   24
            Top             =   900
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Factor"
            Height          =   195
            Left            =   915
            TabIndex        =   23
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Formula"
            Height          =   195
            Left            =   795
            TabIndex        =   22
            Top             =   600
            Width           =   555
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Grupo de Bienes o Servicios"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7365
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Objeto"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label Label18 
      Caption         =   "Grupo de Bienes o Servicios"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   2280
      Width           =   2085
   End
End
Attribute VB_Name = "frmLogProSelMntFactoresEvaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub FormaFlex()
With MSFlex
    .Clear
    .Rows = 2
    .RowHeight(0) = 320
    .RowHeight(1) = 10
    .ColWidth(0) = 0:     .TextMatrix(0, 0) = "Cod Tipo":        '.ColAlignment(1) = 4
    .ColWidth(1) = 3000:    .TextMatrix(0, 1) = "Tipo"
    .ColWidth(2) = 0:     .TextMatrix(0, 2) = "Cod Sub Tipo":        '.ColAlignment(2) = 4
    .ColWidth(3) = 3000:     .TextMatrix(0, 3) = "SubTipo":
End With
End Sub

Sub FormaFlexValores()
With MSFlexValores
    .Clear
    .Rows = 2
    .RowHeight(0) = 320
    .RowHeight(1) = 10
    .ColWidth(0) = 800:     .TextMatrix(0, 0) = "Item":        '.ColAlignment(1) = 4
    .ColWidth(1) = 1000:    .TextMatrix(0, 1) = "Minimo"
    .ColWidth(2) = 1000:     .TextMatrix(0, 2) = "Maximo":        '.ColAlignment(2) = 4
    .ColWidth(3) = 1000:     .TextMatrix(0, 3) = "Puntaje":        '.ColAlignment(2) = 4
End With
End Sub

Sub FormaFlexValoresVer()
With MSFlexValoresVer
    .Clear
    .Rows = 2
    .RowHeight(0) = 320
    .RowHeight(1) = 10
    .ColWidth(0) = 800:     .TextMatrix(0, 0) = "Item":        '.ColAlignment(1) = 4
    .ColWidth(1) = 1000:    .TextMatrix(0, 1) = "Minimo"
    .ColWidth(2) = 1000:     .TextMatrix(0, 2) = "Maximo":        '.ColAlignment(2) = 4
    .ColWidth(3) = 1000:     .TextMatrix(0, 3) = "Puntaje":        '.ColAlignment(2) = 4
End With
End Sub

Sub FormaFlexFactores()
With MSFlexFactores
    .Clear
    .Rows = 2
    .Cols = 6
    .RowHeight(0) = 320
    .RowHeight(1) = 10
    .ColWidth(0) = 3000:     .TextMatrix(0, 0) = "Factor":        '.ColAlignment(1) = 4
    .ColWidth(1) = 700:    .TextMatrix(0, 1) = "Puntaje"
    .ColWidth(2) = 2500:     .TextMatrix(0, 2) = "Formula":        '.ColAlignment(2) = 4
    .ColWidth(3) = 0:     .TextMatrix(0, 3) = "NroFactor":
    .ColWidth(4) = 1700:     .TextMatrix(0, 4) = "Unidades":
    .ColWidth(5) = 1200:     .TextMatrix(0, 5) = "Propuesta":
End With
End Sub

Private Sub CargarProceso()
    On Error GoTo CargarProcesoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select distinct p.cProSelTpoDescripcion, p.nProSelTpoCod, r.cProSelSubTpo, r.nProSelSubTpo " & _
                "from LogProSeltpo p inner join LogProSeltpoRangos r on p.nProSelTpoCod=r.nProSelTpoCod"
        FormaFlex
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            i = i + 1
            InsRow MSFlex, i
            MSFlex.TextMatrix(i, 0) = Rs!nProSelTpoCod
            MSFlex.TextMatrix(i, 1) = Rs!cProSelTpoDescripcion
            MSFlex.TextMatrix(i, 2) = Rs!nProSelSubTpo
            MSFlex.TextMatrix(i, 3) = Rs!cProSelSubTpo
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
CargarProcesoErr:
    MsgBox Err.Number & Err.Description, vbInformation
End Sub

Private Sub CargarFactoresProceso(ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pnObjeto As Integer, ByVal pcBSGrupoCod As String)
    On Error GoTo CargarEtapasProcesoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select f.nFactorNro,x.cFactorDescripcion, f.nPuntaje, c.cConsDescripcion, x.cUnidades, x.nTipo " & _
                "from LogProSelEvalTpoFactor f " & _
                "inner join constante c on f.nFormula=c.nConsValor and c.nConsCod=9084 " & _
                "inner join LogProSelFactor x on f.nFactorNro=x.nFactorNro " & _
                "Where f.nVigente = 1 And f.nProSelTpoCod = " & pnProSelTpoCod & _
                " And f.nProSelSubTpo =" & pnProSelSubTpo & " and nObjeto= " & pnObjeto & " and cBSGrupoCod= '" & pcBSGrupoCod & "'"
                '"inner join constante c on c.nConsValor=f.nobjeto and nConsCod=9044 "
        Set Rs = oCon.CargaRecordSet(sSQL)
        FormaFlexFactores
        i = 0
        Do While Not Rs.EOF
            With MSFlexFactores
                i = i + 1
                InsRow MSFlexFactores, i
                .TextMatrix(i, 0) = Rs!cFactorDescripcion
                .TextMatrix(i, 1) = Rs!npuntaje
                .TextMatrix(i, 2) = Rs!cConsDescripcion
                .TextMatrix(i, 3) = Rs!nFactorNro
                .TextMatrix(i, 4) = Rs!cUnidades
                .TextMatrix(i, 5) = IIf(Rs!nTipo, "Economica", "Tecnica")
                End With
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    MSFlexFactores_GotFocus
    Exit Sub
CargarEtapasProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cboFactores_Click()
    Dim Pos As Integer, postipo As Integer
    txtunidades.Text = ""
    txttipo.Text = ""
    Pos = InStr(1, cboFactores.Text, "/")
    postipo = InStr(1, cboFactores.Text, "#")
    If Pos + 1 <> postipo Then txtunidades.Text = Mid(cboFactores.Text, Pos + 1, postipo - Pos - 1)
    txttipo.Text = Mid(cboFactores.Text, postipo + 1, 100)
End Sub

Private Sub CboFormula_Click()
    If cboFactores.ListCount = 0 Then Exit Sub
    If CboFormula.ListIndex = 2 Then
        FrameValores.Visible = True
    Else
        FrameValores.Visible = False
    End If
End Sub

Private Sub cboGrupoBS_Click()
    MSFlex_SelChange
End Sub

Private Sub CboObj_Click()
    CargarGrupoBS (CboObj.ItemData(CboObj.ListIndex))
    MSFlex_SelChange
End Sub

Private Sub cmdAgregar_Click()
    Dim i As Integer, Cadena As String
    i = 1
    Do While i < MSFlexFactores.Rows
        If Cadena = "" Then
            Cadena = MSFlexFactores.TextMatrix(i, 3)
        Else
            Cadena = Cadena & "," & MSFlexFactores.TextMatrix(i, 3)
        End If
        i = i + 1
    Loop
    MSFlex.Enabled = False
    CboObj.Enabled = False
    cboGrupoBS.Enabled = False
    FormaFlexValores
    CargarFactores Cadena
    FrameLista.Visible = False
    frameFactores.Visible = True
    SSTab1.TabEnabled(1) = False
End Sub

Private Sub cmdCancelarApelacion_Click()
    MSFlex.Enabled = True
    CboObj.Enabled = True
    cboGrupoBS.Enabled = True
    FrameLista.Visible = True
    frameFactores.Visible = False
    CboFormula.ListIndex = 0
    txtpuntaje.Text = ""
    txtMayor.Text = ""
    TxtMenor.Text = ""
    SSTab1.TabEnabled(1) = True
    MSFlex_SelChange
End Sub

Private Sub cmdGrabarApelacion_Click()
    On Error GoTo cmdGrabarApelacionErr
    Dim oCon As DConecta, sSQL As String, i As Integer, nFactor As Integer
    Set oCon = New DConecta
    If Len(Trim(txtpuntaje.Text)) = 0 Then Exit Sub
    If cboFactores.ListCount = 0 Then Exit Sub
    
    If FrameValores.Visible Then
        For i = 2 To MSFlexValores.Rows - 1
            If Len(MSFlexValores.TextMatrix(i, 1)) = 0 Then
               MsgBox "No Existen Valores...", vbInformation, "Aviso"
               Exit Sub
            End If
        Next
    End If
    
    If Not ValidarTotalPuntos(IIf(CboFormula.ListIndex = 2, MaxValor, txtpuntaje.Text)) Then
        MsgBox "Error ha Sobrepasado el total Permitido de Puntos", vbInformation, "Aviso"
        cmdCancelarApelacion_Click
        Exit Sub
    End If
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
        sSQL = "declare @tmp int "
        sSQL = sSQL & "set @tmp=(select count(*) from LogProSelEvalTpoFactor Where nVigente=0 and nProSelTpoCod = " & MSFlex.TextMatrix(MSFlex.row, 0) & " And nProSelSubTpo = " & MSFlex.TextMatrix(MSFlex.row, 2) & " And cBSGrupoCod = '" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & "' and nObjeto=" & CboObj.ItemData(CboObj.ListIndex) & " and nFactorNro= " & cboFactores.ItemData(cboFactores.ListIndex) & ") "
        sSQL = sSQL & " if @tmp=0 "
        sSQL = sSQL & " insert into LogProSelEvalTpoFactor(nFactorNro,nProSelTpoCod,nProSelSubTpo,nVigente,nPuntaje,nFormula,nObjeto,cBSGrupoCod) " & _
                "values(" & cboFactores.ItemData(cboFactores.ListIndex) & "," & MSFlex.TextMatrix(MSFlex.row, 0) & "," & _
                MSFlex.TextMatrix(MSFlex.row, 2) & "," & 1 & "," & IIf(CboFormula.ListIndex = 2, MaxValor, txtpuntaje.Text) & _
                "," & CboFormula.ItemData(CboFormula.ListIndex) & "," & CboObj.ItemData(CboObj.ListIndex) & ",'" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & _
                "')" ' & ChkTecnica.Value & ") "
        sSQL = sSQL & "else "
        sSQL = sSQL & "update LogProSelEvalTpoFactor " & _
                "Set nVigente = 1, nPuntaje=" & IIf(CboFormula.ListIndex = 2, MaxValor, txtpuntaje.Text) & "," & _
                " nFormula=" & CboFormula.ItemData(CboFormula.ListIndex) & _
                " Where nVigente=0 and nProSelTpoCod = " & MSFlex.TextMatrix(MSFlex.row, 0) & " And nProSelSubTpo = " & MSFlex.TextMatrix(MSFlex.row, 2) & " And cBSGrupoCod = '" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & "' and nObjeto=" & CboObj.ItemData(CboObj.ListIndex) & " and nFactorNro=" & cboFactores.ItemData(cboFactores.ListIndex)
        
        oCon.Ejecutar sSQL
        sSQL = ""
        If CboFormula.ListIndex = 2 Then
            i = 2
            Do While i < MSFlexValores.Rows
                sSQL = "insert into LogProSelEvalTpoFactorRangos(nFactorNro,nProSelTpoCod,nProSelsubTpo,nObjeto,cBSGrupoCod,nRangoItem,nRangoMin,nRangoMax,nVigente,nPuntaje) " & _
                        "values(" & cboFactores.ItemData(cboFactores.ListIndex) & "," & MSFlex.TextMatrix(MSFlex.row, 0) & "," & MSFlex.TextMatrix(MSFlex.row, 2) & "," & _
                            CboObj.ItemData(CboObj.ListIndex) & ",'" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & "'," & MSFlexValores.TextMatrix(i, 0) & "," & _
                            MSFlexValores.TextMatrix(i, 1) & "," & MSFlexValores.TextMatrix(i, 2) & ",1," & MSFlexValores.TextMatrix(i, 3) & ")"
                If sSQL <> "" Then oCon.Ejecutar sSQL
                i = i + 1
            Loop
        End If
        oCon.CommitTrans
        MsgBox "Factor Asignado Correctamente...", vbInformation
        oCon.CierraConexion
        cmdCancelarApelacion_Click
    End If
    Exit Sub
cmdGrabarApelacionErr:
    oCon.RollbackTrans
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Function ValidarTotalPuntos(ByVal pnPuntos As Integer) As Boolean
On Error GoTo CalcularTotalPuntosErr
    Dim i As Integer, nTotal As Integer, nRow As Integer, nCol As Integer
    With MSFlexFactores
        Do While i < .Rows
            nTotal = nTotal + Val(.TextMatrix(i, 1))
            i = i + 1
        Loop
    End With
    If nTotal + pnPuntos <= 100 Then
        ValidarTotalPuntos = True
    Else
        ValidarTotalPuntos = False
    End If
    Exit Function
CalcularTotalPuntosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Private Sub CmdMas_Click()
On Error GoTo CmdMasErr
    Dim i As Integer
        If Val(txtMayor.Text) < Val(TxtMenor.Text) Then
            MsgBox "Rango de Factores Incorrecto...", vbInformation
            Exit Sub
        End If
        If cboFactores.ListCount = 0 Then
            MsgBox "No Existen Factores...", vbInformation
            Exit Sub
        End If
        If Len(Trim(txtpuntaje.Text)) <= 0 Then
            MsgBox "Puntaje Incorrecto...", vbInformation
            Exit Sub
        End If
        If Len(Trim(txtMayor.Text)) = 0 Then
            MsgBox "Rango Mayor Incorrecto...", vbInformation
            Exit Sub
        End If
        If Len(Trim(TxtMenor.Text)) < 0 Then
            MsgBox "Rango Menor Incorrecto...", vbInformation
            Exit Sub
        End If
        With MSFlexValores
            i = .Rows
            If .Rows > 2 And Val(.TextMatrix(i - 1, 3)) = Val(txtpuntaje.Text) Then
                MsgBox "Puntaje Incorrecto Incorrecto...", vbInformation
                Exit Sub
            End If
            If .Rows > 2 And Val(.TextMatrix(i - 1, 1)) >= Val(TxtMenor.Text) Then
                MsgBox "Rango Menor Incorrecto...", vbInformation
                Exit Sub
            End If
            If .Rows > 2 And Val(.TextMatrix(i - 1, 2)) >= Val(txtMayor.Text) Then
                MsgBox "Rango Mayor Incorrecto...", vbInformation
                Exit Sub
            End If
            If .Rows > 2 And Val(.TextMatrix(i - 1, 2)) + 1 <> Val(TxtMenor.Text) Then
                If MsgBox("Esta Omitiendo Numeros con el Rango Especificado, " & vbCrLf & "Desea Continuar...?", vbQuestion + vbYesNo) = vbNo Then
                    Exit Sub
                    TxtMenor.Text = Val(.TextMatrix(i - 1, 2)) + 1
                End If
            End If
'            If .Rows > 2 And Val(.TextMatrix(i - 1, 3)) <= Val(txtpuntaje.Text) Then Exit Sub
            InsRow MSFlexValores, i
            .TextMatrix(i, 0) = i - 1
'            If .Rows = 3 Then
'                .TextMatrix(i, 1) = 0
'            Else
'                .TextMatrix(i, 1) = .TextMatrix(i - 1, 2) + 1
'            End If
            .TextMatrix(i, 1) = TxtMenor.Text
            .TextMatrix(i, 2) = txtMayor.Text
            .TextMatrix(i, 3) = txtpuntaje.Text
            TxtMenor.Text = Val(txtMayor.Text) + 1
            txtMayor.Text = Val(TxtMenor.Text) + 1
            
        End With
    Exit Sub
CmdMasErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CmdMenos_Click()
Dim i As Integer
Dim K As Integer

i = MSFlexValores.Rows - 1
If Len(Trim(MSFlexValores.TextMatrix(i, 2))) = 0 Then
   Exit Sub
End If

If MsgBox("¿ está seguro de quitar el elemento ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If MSFlexValores.Rows - 1 > 1 Then
      MSFlexValores.RemoveItem i
   Else
      For K = 0 To MSFlexValores.Cols - 1
          MSFlexValores.TextMatrix(i, K) = ""
      Next
      MSFlexValores.RowHeight(i) = 8
   End If
End If
End Sub

Private Sub cmdModificar_Click()
    MSFlex.Enabled = False
    CboObj.Enabled = False
    cboGrupoBS.Enabled = False
    FrameLista.Visible = False
    frameFactores.Visible = True
    cboFactores.ListIndex = MSFlexFactores.TextMatrix(MSFlexFactores.row, 3)
    cboFactores.Enabled = False
End Sub

Private Sub cmdQuitar_Click()
    On Error GoTo cmdQuitarErr
    Dim oCon As DConecta, sSQL As String
    Set oCon = New DConecta
    If MsgBox("Seguro que Desea Quitar el Factor...", vbQuestion + vbYesNo) = vbYes Then
        If oCon.AbreConexion Then
            oCon.BeginTrans
            
            sSQL = "delete LogProSelEvalTpoFactorRangos " & _
                    " where nProSelTpoCod = " & MSFlex.TextMatrix(MSFlex.row, 0) & _
                    " and nProSelSubTpo= " & MSFlex.TextMatrix(MSFlex.row, 2) & _
                    " And nFactorNro = " & MSFlexFactores.TextMatrix(MSFlexFactores.row, 3) & _
                    " and cBSGrupoCod='" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & _
                    "' and nObjeto=" & CboObj.ItemData(CboObj.ListIndex)
            oCon.Ejecutar sSQL
            
            sSQL = "update LogProSelEvalTpoFactor " & _
                    "Set nVigente = 0 " & _
                    " Where nProSelTpoCod = " & MSFlex.TextMatrix(MSFlex.row, 0) & _
                    " and nProSelSubTpo= " & MSFlex.TextMatrix(MSFlex.row, 2) & _
                    " And nFactorNro = " & MSFlexFactores.TextMatrix(MSFlexFactores.row, 3) & _
                    " and cBSGrupoCod='" & Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4) & _
                    "' and nObjeto=" & CboObj.ItemData(CboObj.ListIndex)
            oCon.Ejecutar sSQL
            
            oCon.CommitTrans
            oCon.CierraConexion
        End If
        MsgBox "Factor Eliminado"
        CargarFactoresProceso MSFlex.TextMatrix(MSFlex.row, 0), MSFlex.TextMatrix(MSFlex.row, 2), CboObj.ItemData(CboObj.ListIndex), Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
    End If
    Exit Sub
cmdQuitarErr:
    oCon.RollbackTrans
    oCon.CierraConexion
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    CargarProceso
    CargarObjeto
    FormaFlexFactores
    CargarFormula
    FormaFlexValoresVer
    MSFlex_SelChange
End Sub

Private Sub CargarFactores(Cadena As String)
    On Error GoTo CargarFactoresErr
    Dim oCon As New DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If Cadena <> "" Then Cadena = "and not nFactorNro in (" & Cadena & ")"
    If oCon.AbreConexion Then
        sSQL = "select * from LogProSelFactor where nFactorEstado=1 " & Cadena
        Set Rs = oCon.CargaRecordSet(sSQL)
        cboFactores.Clear
        Do While Not Rs.EOF
            cboFactores.AddItem Rs!cFactorDescripcion & Space(200) & "/" & Rs!cUnidades & "#" & IIf(Rs!nTipo, "Economica", "Tecnica"), cboFactores.ListCount
            cboFactores.ItemData(cboFactores.ListCount - 1) = Rs!nFactorNro
            Rs.MoveNext
            cboFactores.ListIndex = 0
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
CargarFactoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarFormula()
    On Error GoTo CargarFormulaErr
    Dim oCon As New DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select * from constante where nConsCod=9084 and nConsCod<>nConsValor"
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            CboFormula.AddItem Rs!cConsDescripcion, CboFormula.ListCount
            CboFormula.ItemData(CboFormula.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    CboFormula.ListIndex = 0
    Exit Sub
CargarFormulaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelMntFactoresEvaluacion = Nothing
End Sub

Private Sub MSFlex_SelChange()
    If cboGrupoBS.ListCount > 0 And CboObj.ListCount > 0 And MSFlex.Rows > 2 Then
        CargarFactoresProceso MSFlex.TextMatrix(MSFlex.row, 0), MSFlex.TextMatrix(MSFlex.row, 2), CboObj.ItemData(CboObj.ListIndex), Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
    End If
End Sub

Private Sub txtformula_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub MSFlexFactores_GotFocus()
    CargarFactoresVer Val(MSFlexFactores.TextMatrix(MSFlexFactores.row, 3)), Val(MSFlex.TextMatrix(MSFlex.row, 0)), Val(MSFlex.TextMatrix(MSFlex.row, 2)), CboObj.ItemData(CboObj.ListIndex), Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
End Sub

Private Sub MSFlexFactores_RowColChange()
    CargarFactoresVer Val(MSFlexFactores.TextMatrix(MSFlexFactores.row, 3)), Val(MSFlex.TextMatrix(MSFlex.row, 0)), Val(MSFlex.TextMatrix(MSFlex.row, 2)), CboObj.ItemData(CboObj.ListIndex), Right(cboGrupoBS.List(cboGrupoBS.ListIndex), 4)
End Sub

Private Sub txtMayor_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtMayor, KeyAscii)
End Sub

Private Sub TxtMenor_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(TxtMenor, KeyAscii)
End Sub

Private Sub txtpuntaje_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtpuntaje, KeyAscii)
End Sub

Private Sub CargarObjeto()
On Error GoTo CargarModalidadErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String
    Set oCon = New DConecta
    sSQL = "select * from Constante where nConsCod =9044 and nConsValor<>nConsCod"
    CboObj.Clear
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            CboObj.AddItem Rs!cConsDescripcion, CboObj.ListCount
            CboObj.ItemData(CboObj.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        oCon.CierraConexion
        If CboObj.ListCount > 0 Then CboObj.ListIndex = 0
    End If
Exit Sub
CargarModalidadErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub CargarGrupoBS(psBSCod As Integer)
    On Error GoTo CargarGrupoBSErr
    Dim oCon As New DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    sSQL = "select * from BSGrupos where nObjetoCod= " & psBSCod & "and len(cBSGrupoCod)=4"
    If oCon.AbreConexion Then
        cboGrupoBS.Clear
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            cboGrupoBS.AddItem Rs!cBSGrupoDescripcion & Space(150) & Rs!cBSGrupoCod, cboGrupoBS.ListCount
            'cboGrupoBS.ItemData(cboGrupoBS.ListCount - 1) = Rs!cBSGrupoCod
            Rs.MoveNext
        Loop
        oCon.CierraConexion
        If cboGrupoBS.ListCount > 0 Then cboGrupoBS.ListIndex = 0
    End If
    Exit Sub
CargarGrupoBSErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub CargarFactoresVer(pnFactorNro As Integer, pnProSelTpoCod As Integer, pnProSelSubTpo As Integer, pnObjeto As Integer, pcBSGrupoCod As String)
    On Error GoTo CargarFactoresErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    FormaFlexValoresVer
    sSQL = "select * from LogProSelEvalTpoFactorRangos where nFactorNro=" & pnFactorNro & " and  nProSelTpoCod=" & pnProSelTpoCod & " and nProSelsubTpo=" & pnProSelSubTpo & _
            " and  nObjeto=" & pnObjeto & " and cBSGrupoCod='" & pcBSGrupoCod & "'"
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            With MSFlexValoresVer
                i = i + 1
                InsRow MSFlexValoresVer, i
                .TextMatrix(i, 0) = Rs!nRangoItem
                .TextMatrix(i, 1) = Rs!nRangoMin
                .TextMatrix(i, 2) = Rs!nRangoMax
                .TextMatrix(i, 3) = Rs!npuntaje
            End With
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
CargarFactoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Function MaxValor() As Integer
    Dim i As Integer, max As Integer
    i = 2
    Do While i < MSFlexValores.Rows
        If max < MSFlexValores.TextMatrix(i, 3) Then
            max = MSFlexValores.TextMatrix(i, 3)
        End If
        i = i + 1
    Loop
    MaxValor = max
End Function






