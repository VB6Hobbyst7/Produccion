VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLimiteEfectivoAdm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Límite de Efectivo "
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmLimiteEfectivoAdm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Registro"
      TabPicture(0)   =   "frmLimiteEfectivoAdm.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FEDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdLimpiar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGuardar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame4 
         Caption         =   "Historial de Registros"
         Height          =   1335
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   7695
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "Eliminar"
            Height          =   375
            Left            =   6915
            TabIndex        =   20
            ToolTipText     =   "Eliminar"
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   6915
            TabIndex        =   19
            ToolTipText     =   "Eliminar"
            Top             =   240
            Width           =   735
         End
         Begin VB.ListBox lstHistorial 
            Height          =   1035
            Left            =   70
            TabIndex        =   18
            Top             =   190
            Width           =   6825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Grupo"
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   3735
         Begin VB.ComboBox cboGrupo 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Hasta"
         Height          =   855
         Left            =   3840
         TabIndex        =   10
         Top             =   1800
         Width           =   3975
         Begin VB.ComboBox cboMes2 
            Height          =   315
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   420
            Width           =   1575
         End
         Begin VB.ComboBox cboAnio2 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Año"
            Height          =   255
            Left            =   2160
            TabIndex        =   14
            Top             =   225
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Mes"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   225
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Desde"
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   3735
         Begin VB.ComboBox cboAnio1 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   420
            Width           =   1215
         End
         Begin VB.ComboBox cboMes1 
            Height          =   315
            ItemData        =   "frmLimiteEfectivoAdm.frx":0326
            Left            =   360
            List            =   "frmLimiteEfectivoAdm.frx":0328
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Año"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   210
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Mes"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   210
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   7200
         Width           =   975
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "&Limpiar"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   7200
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6840
         TabIndex        =   1
         Top             =   7200
         Width           =   975
      End
      Begin SICMACT.FlexEdit FEDatos 
         Height          =   3855
         Left            =   120
         TabIndex        =   4
         Top             =   3315
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6800
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Agencia-PÓLIZA $-MÁXIMO $-MÍNIMO $-cAgeCod"
         EncabezadosAnchos=   "500-2200-1500-1500-1500-0"
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
         ColumnasAEditar =   "X-X-2-3-4-X"
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "0-L-R-R-R-0"
         FormatosEdit    =   "0-0-2-2-2-0"
         TextArray0      =   "Nro"
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmLimiteEfectivoAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmLimiteEfectivoAdm
'** Descripción : Formulario para realizar la administracion del límite de efectivo
'** Creación    : RECO, 20141011 - ERS022-2014
'**********************************************************************************************

Option Explicit

Private Sub cargarAnios()
    Dim i As Integer
    
    cboAnio1.Clear
    cboAnio2.Clear
    
    For i = 2012 To 2050
        cboAnio1.AddItem i
        cboAnio2.AddItem i
    Next i
    
    cboAnio1.ListIndex = 0
    cboAnio2.ListIndex = 0
End Sub

Private Sub cargarMeses()
    Dim i As Integer
    
    cboMes1.Clear
    cboMes2.Clear
    
    For i = 1 To 12
        cboMes1.AddItem IIf(i = 1, "Enero", IIf(i = 2, "Febrero", IIf(i = 3, "Marzo", IIf(i = 4, "Abril", IIf(i = 5, "Mayo", IIf(i = 6, "Junio", IIf(i = 7, "Julio", IIf(i = 8, "Agosto", IIf(i = 9, "Setiembre", IIf(i = 10, "Octubre", IIf(i = 11, "Noviembre", IIf(i = 12, "Diciembre", "")))))))))))) & Space(50) & i
        cboMes2.AddItem IIf(i = 1, "Enero", IIf(i = 2, "Febrero", IIf(i = 3, "Marzo", IIf(i = 4, "Abril", IIf(i = 5, "Mayo", IIf(i = 6, "Junio", IIf(i = 7, "Julio", IIf(i = 8, "Agosto", IIf(i = 9, "Setiembre", IIf(i = 10, "Octubre", IIf(i = 11, "Noviembre", IIf(i = 12, "Diciembre", "")))))))))))) & Space(50) & i
    Next i
    
    cboMes1.ListIndex = 0
    cboMes2.ListIndex = 0
End Sub
Private Sub Registrar()
    Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    Set rs = oDCOMGeneral.RegistraLimiteEfectivo(cboGrupo.ItemData(cboGrupo.ListIndex), CInt(Right(cboAnio1, 4)), CInt(Right(cboMes1, 2)), CInt(Right(cboAnio2, 4)), CInt(Right(cboMes2, 2)))
    If Not (rs.EOF And rs.BOF) Then
        Dim nCodigoID As Integer
        nCodigoID = rs!nCodigoID
        For i = 1 To FEDatos.Rows - 2
            Call oDCOMGeneral.RegistraLimiteEfectivoDet(nCodigoID, FEDatos.TextMatrix(i, 5), Replace(FEDatos.TextMatrix(i, 2), ",", ""), Replace(FEDatos.TextMatrix(i, 3), ",", ""), Replace(FEDatos.TextMatrix(i, 4), ",", ""))
        Next
    End If
    Call CargarHistorial
End Sub

Private Sub cmdAceptar_Click()
    Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Set rs = oDCOMGeneral.ObtieneDatosLimiteEfectivo(CInt(Right(lstHistorial.Text, 4)))
    
    FEDatos.Clear
    FormateaFlex FEDatos
    If Not (rs.EOF And rs.BOF) Then
        For i = 1 To rs.RecordCount - 1
            FEDatos.AdicionaFila
            FEDatos.TextMatrix(i, 1) = rs!cAgeDescripcion
            FEDatos.TextMatrix(i, 2) = Format(rs!nPoliza, "#,##0.00")
            FEDatos.TextMatrix(i, 3) = Format(rs!nMaximo, "#,##0.00")
            FEDatos.TextMatrix(i, 4) = Format(rs!nMinimo, "#,##0.00")
            FEDatos.TextMatrix(i, 5) = rs!cAgeCod
            rs.MoveNext
        Next
    End If
    Set rs = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
    Call oDCOMGeneral.EliminaLimiteEfectivo(CInt(Right(lstHistorial.Text, 4)))
    Call CmdLimpiar_Click
End Sub

Private Sub cmdGuardar_Click()
    Call Registrar
End Sub

Private Sub CargarAgencias()
Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
Dim rsAgencias As New ADODB.Recordset
Dim i As Integer

Set rsAgencias = oDCOMGeneral.devolverAgenciasParaCoberturar

If Not rsAgencias.BOF And Not rsAgencias.EOF Then
i = 1
Call LimpiaFlex(FEDatos)
FEDatos.lbEditarFlex = True
    Do While Not rsAgencias.EOF
            FEDatos.AdicionaFila
            FEDatos.TextMatrix(i, 1) = rsAgencias!cAgeDescripcion
            FEDatos.TextMatrix(i, 2) = Format(0#, "#,##0.00")
            FEDatos.TextMatrix(i, 3) = Format(0#, "#,##0.00")
            FEDatos.TextMatrix(i, 4) = Format(0#, "#,##0.00")
            FEDatos.TextMatrix(i, 5) = rsAgencias!cAgeCod
        i = i + 1
        rsAgencias.MoveNext
    Loop
Else
    MsgBox "No se registraron las Agencias.", vbInformation, "Aviso"
End If

End Sub

Public Sub CargarCombo()
    Dim rs As New ADODB.Recordset
    Dim oConst As New COMDConstantes.DCOMConstantes
    Set rs = oConst.RecuperaConstantes(10050)
    cboGrupo.Clear
    Do Until rs.EOF
        cboGrupo.AddItem "" & rs!cConsDescripcion
        cboGrupo.ItemData(cboGrupo.NewIndex) = "" & rs!nConsValor
        rs.MoveNext
    Loop
    Set rs = Nothing
    cboGrupo.ListIndex = 0
End Sub

Public Sub CargarHistorial()
    Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim sCadena As String
    
    Set rs = oDCOMGeneral.ObtieneHistorialLimiteEfectivo
    lstHistorial.Clear
    If Not (rs.EOF And rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            Dim nLen As Integer, nLenResul As Integer, nLenSuma As Integer
            nLen = Len(rs!cConsDescripcion)
            nLenResul = 38 - nLen
            nLenSuma = nLen + nLenResul
            
            sCadena = rs!cConsDescripcion
            'sCadena = sCadena & Space(nLenResul)
            sCadena = sCadena & " - Desde " & DevulveMes(rs!nMesDesde) & " " & rs!nAnioDesde
            sCadena = sCadena & " Hasta " & DevulveMes(rs!nMesHasta) & " " & rs!nAnioHasta
            sCadena = sCadena & Space(50) & rs!nCodigoID
            'lstHistorial.AddItem (UCase(sCadena))
            lstHistorial.AddItem (sCadena)
            rs.MoveNext
        Next
    End If
    Set rs = Nothing
End Sub

Public Function DevulveMes(ByVal nMES As Integer) As String
    DevulveMes = IIf(nMES = 1, "Enero", IIf(nMES = 2, "Febrero", IIf(nMES = 3, "Marzo", IIf(nMES = 4, "Abril", IIf(nMES = 5, "Mayo", IIf(nMES = 6, "Junio", IIf(nMES = 7, "Julio", IIf(nMES = 8, "Agosto", IIf(nMES = 9, "Setiembre", IIf(nMES = 10, "Octubre", IIf(nMES = 11, "Noviembre", IIf(nMES = 12, "Diciembre", ""))))))))))))
End Function

Public Sub Inicio(ByVal pnTpoOpe As Integer, ByVal psTitulo As String)
    Me.Caption = psTitulo
    SSTab1.TabCaption(0) = IIf(pnTpoOpe = 1, "Registro", "Mantenimiento")
    If pnTpoOpe = 1 Then
        Call CargarHistorial
        Call CargarCombo
        Call CargarAgencias
        Call cargarAnios
        Call cargarMeses
        Call HabilitaControles(False)
    Else
        Call CargarHistorial
        Call HabilitaControles(True)
    End If
    Me.Show 1
End Sub

Public Sub HabilitaControles(ByVal pbValor As Boolean)
    cboGrupo.Enabled = Not pbValor
    cboAnio1.Enabled = Not pbValor
    cboAnio2.Enabled = Not pbValor
    cboMes1.Enabled = Not pbValor
    cboMes2.Enabled = Not pbValor
    cmdAceptar.Enabled = pbValor
    cmdEliminar.Enabled = pbValor
End Sub

Private Sub CmdLimpiar_Click()
    Call CargarCombo
    Call CargarHistorial
    Call CargarAgencias
    Call cargarAnios
    Call cargarMeses
End Sub
