VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRHConceptoAsigRRHH 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmRHConceptoAsigRRHH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.ctrRRHHGen RRHH 
      Height          =   1200
      Left            =   60
      TabIndex        =   16
      Top             =   15
      Width           =   7470
      _extentx        =   13176
      _extenty        =   2117
      font            =   "frmRHConceptoAsigRRHH.frx":030A
   End
   Begin VB.PictureBox Pic 
      Height          =   315
      Left            =   1230
      Picture         =   "frmRHConceptoAsigRRHH.frx":0336
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   6285
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5025
      Left            =   45
      TabIndex        =   4
      Top             =   1215
      Width           =   7515
      Begin Sicmact.TxtBuscar TxtPlanilla 
         Height          =   300
         Left            =   1125
         TabIndex        =   17
         Top             =   210
         Width           =   1410
         _extentx        =   2487
         _extenty        =   529
         appearance      =   0
         appearance      =   0
         backcolor       =   -2147483624
         font            =   "frmRHConceptoAsigRRHH.frx":0678
         appearance      =   0
         stitulo         =   ""
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4365
         Left            =   75
         TabIndex        =   5
         Top             =   615
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   7699
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "&Ingresos"
         TabPicture(0)   =   "frmRHConceptoAsigRRHH.frx":06A4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FlexIng"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "TextEdit"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "&Descuentos"
         TabPicture(1)   =   "frmRHConceptoAsigRRHH.frx":06C0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FlexDes"
         Tab(1).Control(1)=   "TextEditD"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "&Aportaciones"
         TabPicture(2)   =   "frmRHConceptoAsigRRHH.frx":06DC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FlexApo"
         Tab(2).Control(1)=   "TextEditA"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "&Variables"
         TabPicture(3)   =   "frmRHConceptoAsigRRHH.frx":06F8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FlexVar"
         Tab(3).Control(1)=   "TextEditV"
         Tab(3).ControlCount=   2
         Begin VB.TextBox TextEditV 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   -72975
            MaxLength       =   30
            TabIndex        =   9
            Top             =   645
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TextEditA 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   -71760
            MaxLength       =   30
            TabIndex        =   8
            Top             =   1380
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TextEditD 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   -71400
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1260
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TextEdit 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   4800
            MaxLength       =   30
            TabIndex        =   6
            Top             =   900
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexIng 
            CausesValidation=   0   'False
            Height          =   3825
            Left            =   75
            TabIndex        =   10
            Top             =   405
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   6747
            _Version        =   393216
            BackColorSel    =   12502347
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexDes 
            CausesValidation=   0   'False
            Height          =   3885
            Left            =   -74925
            TabIndex        =   12
            Top             =   390
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   6853
            _Version        =   393216
            BackColorSel    =   12502347
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexApo 
            CausesValidation=   0   'False
            Height          =   3855
            Left            =   -74910
            TabIndex        =   13
            Top             =   405
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   6800
            _Version        =   393216
            BackColorSel    =   12502347
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexVar 
            CausesValidation=   0   'False
            Height          =   3840
            Left            =   -74910
            TabIndex        =   14
            Top             =   405
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   6773
            _Version        =   393216
            BackColorSel    =   12502347
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Label lblPlanilla 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2580
         TabIndex        =   18
         Top             =   195
         Width           =   4830
      End
      Begin VB.Label lblPlaCod 
         Caption         =   "Cod.Planilla:"
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   380
      Left            =   6510
      TabIndex        =   3
      Top             =   6285
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   380
      Left            =   5415
      TabIndex        =   2
      Top             =   6285
      Width           =   1050
   End
   Begin VB.CommandButton cmdTodosEmp 
      Caption         =   "&Aplica T."
      Height          =   380
      Left            =   4320
      TabIndex        =   1
      Top             =   6285
      Width           =   1050
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   380
      Left            =   3225
      TabIndex        =   0
      Top             =   6285
      Visible         =   0   'False
      Width           =   1050
   End
End
Attribute VB_Name = "frmRHConceptoAsigRRHH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnColorB As Long
Dim lnColorP As Long
Dim lsCadenaValida As String
Dim lnTipo As TipoOpe
Dim lbConceptoFijo As Boolean

Dim lsCaption As String

Private Sub cmdAplicar_Click()
    'AsignaValEmp RRHH.psCodigoPersona
    'IniPan RRHH.psCodigoPersona
End Sub

Private Sub cmdImprimir_Click()
    'frmPrevio.Previo r, "Conceptos por Persona", True, 66
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTodosEmp_Click()
    Dim lsCon As String
    Dim lbBan As Boolean
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    If Me.SSTab1.Tab = 0 Then
        lsCon = Trim(Me.FlexIng.TextMatrix(Me.FlexIng.Row, 0))
        If FlexIng.TextMatrix(Me.FlexIng.Row, 2) = "." Then
            lbBan = False
        Else
            lbBan = True
        End If
    ElseIf Me.SSTab1.Tab = 1 Then
        lsCon = Trim(Me.FlexDes.TextMatrix(Me.FlexDes.Row, 0))
        If FlexDes.TextMatrix(Me.FlexDes.Row, 2) = "." Then
            lbBan = False
        Else
            lbBan = True
        End If
    ElseIf Me.SSTab1.Tab = 2 Then
        lsCon = Trim(Me.FlexApo.TextMatrix(Me.FlexApo.Row, 0))
        If FlexApo.TextMatrix(Me.FlexApo.Row, 2) = "." Then
            lbBan = False
        Else
            lbBan = True
        End If
    ElseIf Me.SSTab1.Tab = 3 Then
        lsCon = Trim(Me.FlexVar.TextMatrix(Me.FlexVar.Row, 0))
        If FlexVar.TextMatrix(Me.FlexVar.Row, 2) = "." Then
            lbBan = False
        Else
            lbBan = True
        End If
    End If
    
    If lsCon = "" Then Exit Sub
    
    
    If MsgBox("Si Pulsa yes,Se Aplicaran los cambios A Todos los Trabajadores ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
      Exit Sub
    
    End If
    

        
    If lbBan Then
        oCon.EliminaConceptoRRHH "", TxtPlanilla.Text, lsCon, True
        oCon.AgregaConceptoRRHH "", TxtPlanilla.Text, lsCon, "0", GetMovNro(gsCodUser, gsCodAge), True
    Else
        oCon.EliminaConceptoRRHH "", TxtPlanilla.Text, lsCon, True
    End If
    
    Flex_EnterCell
End Sub

Private Sub Flex_EnterCell()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    IniFlexCon FlexIng, RHConceptosTpoVisible.RHConceptosTpoVIngreso
    IniFlexCon FlexDes, RHConceptosTpoVisible.RHConceptosTpoVEgreso
    IniFlexCon FlexApo, RHConceptosTpoVisible.RHConceptosTpoVAportacion
    IniFlexCon FlexVar, RHConceptosTpoVisible.RHConceptosTpoVVarUsuario

    Set rsE = oCon.GetConceptosPlanillaRRHH(Me.TxtPlanilla.Text, Me.RRHH.psCodigoPersona)
    If Not RSVacio(rsE) Then
        While Not rsE.EOF
            If Left(rsE.Fields(0), 1) = RHConceptosTpoVisible.RHConceptosTpoVIngreso Then
                SetCon FlexIng, Trim(rsE.Fields(0)), rsE!Monto & ""
            ElseIf Left(rsE.Fields(0), 1) = RHConceptosTpoVisible.RHConceptosTpoVEgreso Then
                SetCon FlexDes, Trim(rsE.Fields(0)), rsE!Monto & ""
            ElseIf Left(rsE.Fields(0), 1) = RHConceptosTpoVisible.RHConceptosTpoVAportacion Then
                SetCon FlexApo, Trim(rsE.Fields(0)), rsE!Monto & ""
            ElseIf Left(rsE.Fields(0), 1) = RHConceptosTpoVisible.RHConceptosTpoVVarUsuario Then
                SetCon FlexVar, Trim(rsE.Fields(0)), rsE!Monto & ""
            End If
            rsE.MoveNext
        Wend
    End If
    
    Set oCon = Nothing
    rsE.Close
    Set rsE = Nothing
End Sub

Private Sub FlexApo_RowColChange()
    FlexApo.Refresh
End Sub

Private Sub FlexDes_RowColChange()
    FlexDes.Refresh
End Sub

Private Sub FlexIng_RowColChange()
    FlexIng.Refresh
End Sub

Private Sub FlexVar_RowColChange()
    FlexVar.Refresh
End Sub

'Private Sub RRHH1_Click()
'    Dim oPersona As UPersona
'    Dim oRRHH As DActualizaDatosRRHH
'    Set oRRHH = New DActualizaDatosRRHH
'    Set oPersona = New UPersona
'    Set oPersona = frmBuscaPersona.Inicio(True)
'    If Not oPersona Is Nothing Then
'        Limpia
'        Me.RRHH.psCodigoPersona = oPersona.sPersCod
'        Me.RRHH.psDNIPersona = oPersona.sPersIdnroDNI
'        Me.RRHH.psFechaNacimiento = oPersona.sPersCod
'        Me.RRHH.psNombreEmpledo = oPersona.sPersNombre
'        Me.RRHH.psDireccionPersona = oPersona.sPersDireccDomicilio
'        Me.RRHH.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(oPersona.sPersCod)
'    End If
'End Sub

Private Sub RRHH_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        Limpia
        Me.RRHH.psCodigoPersona = oPersona.sPersCod
        Me.RRHH.psNombreEmpledo = oPersona.sPersNombre
        Me.RRHH.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.RRHH.psCodigoPersona)
        TxtPlanilla_EmiteDatos
        'CargaData Me.ctrRRHHGen.psCodigoPersona
    End If
End Sub

Private Sub RRHH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oRRHH As DActualizaDatosRRHH
        Dim rsR As ADODB.Recordset
        Set oRRHH = New DActualizaDatosRRHH
        RRHH.psCodigoEmpleado = Left(RRHH.psCodigoEmpleado, 1) & Format(Trim(Mid(RRHH.psCodigoEmpleado, 2)), "00000")
        Dim oCon As DActualizaDatosContrato
        Set oCon = New DActualizaDatosContrato
        
        Set rsR = oRRHH.GetRRHH(RRHH.psCodigoEmpleado, gPersIdDNI)
           
        If Not (rsR.EOF And rsR.BOF) Then
            RRHH.SpinnerValor = CInt(Right(RRHH.psCodigoEmpleado, 5))
            RRHH.psCodigoPersona = rsR.Fields("Codigo")
            RRHH.psNombreEmpledo = rsR.Fields("Nombre")
            rsR.Close
            Set rsR = oRRHH.GetRRHHGeneralidades(RRHH.psCodigoEmpleado)
            TxtPlanilla_EmiteDatos
            'CargaData Me.RRHH.psCodigoPersona
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            Limpia
            RRHH.SetFocus
        End If
        
        rsR.Close
        Set rsR = Nothing
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub FlexApo_DblClick()
    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
    SetEstCon FlexApo
    If FlexApo.Col <> 3 Then
        Exit Sub
    Else
       MSHFlexGridEdit FlexApo, TextEditA, 32 ' Simula un espacio.
    End If
End Sub

Private Sub FlexApo_EnterCell()
     FlexApo.Col = 1
     FlexApo.CellBackColor = lnColorP
     FlexApo.Col = 2
     FlexApo.CellBackColor = lnColorP
     FlexApo.Col = 3
End Sub

Private Sub FlexApo_GotFocus()
    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
    If Me.FlexApo.Col <> 3 Then
        Exit Sub
    Else
       If TextEditA.Visible = False Then Exit Sub
        FlexApo = TextEditA
        SetEstCon FlexApo, True
        TextEditA.Visible = False
    End If
End Sub

Private Sub FlexApo_KeyPress(KeyAscii As Integer)
    If lnTipo <> gTipoOpeMantenimiento Or ((Not lbConceptoFijo) And KeyAscii <> 13) Then Exit Sub

    If NumerosEnteros(KeyAscii) = 0 Then
        Exit Sub
    ElseIf KeyAscii = 13 Then
        FlexApo_DblClick
    End If
    
    If Me.FlexApo.Col <> 3 Then
        Exit Sub
    Else
       MSHFlexGridEdit FlexApo, TextEditA, KeyAscii
    End If
End Sub

Private Sub FlexApo_LeaveCell()
     FlexApo.Col = 1
     FlexApo.CellBackColor = lnColorB
     FlexApo.Col = 2
     FlexApo.CellBackColor = lnColorB
     FlexApo.Col = 3
End Sub

Private Sub FlexDes_DblClick()
    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
    
    SetEstCon FlexDes
    If FlexDes.Col <> 3 Then
        Exit Sub
    Else
       MSHFlexGridEdit FlexDes, TextEditD, 32 ' Simula un espacio.
    End If
End Sub

Private Sub FlexDes_EnterCell()
     FlexDes.Col = 1
     FlexDes.CellBackColor = lnColorP
     FlexDes.Col = 2
     FlexDes.CellBackColor = lnColorP
     FlexDes.Col = 3
End Sub

Private Sub FlexDes_GotFocus()
    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
    
    If Me.FlexDes.Col <> 3 Then
        Exit Sub
    Else
       If TextEditD.Visible = False Then Exit Sub
        FlexDes = TextEditD
        SetEstCon FlexDes, True
        TextEditD.Visible = False
    End If
End Sub

Private Sub FlexDes_KeyPress(KeyAscii As Integer)
    If lnTipo <> gTipoOpeMantenimiento Or ((Not lbConceptoFijo) And KeyAscii <> 13) Then Exit Sub

    If NumerosEnteros(KeyAscii) = 0 Then
        Exit Sub
    ElseIf KeyAscii = 13 Then
        FlexDes_DblClick
    End If
    
    If Me.FlexDes.Col <> 3 Then
        Exit Sub
    Else
       MSHFlexGridEdit FlexDes, TextEditD, KeyAscii
    End If
End Sub

Private Sub FlexDes_LeaveCell()
     FlexDes.Col = 1
     FlexDes.CellBackColor = lnColorB
     FlexDes.Col = 2
     FlexDes.CellBackColor = lnColorB
     FlexDes.Col = 3
End Sub

Private Sub FlexIng_DblClick()
    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
    SetEstCon FlexIng
    If FlexIng.Col <> 3 Then
        Exit Sub
    Else
       MSHFlexGridEdit FlexIng, TextEdit, 32  ' Simula un espacio.
    End If
End Sub

Private Sub FlexIng_EnterCell()
     FlexIng.Col = 1
     FlexIng.CellBackColor = lnColorP
     FlexIng.Col = 2
     FlexIng.CellBackColor = lnColorP
     FlexIng.Col = 3
End Sub

Private Sub FlexIng_LeaveCell()
     FlexIng.Col = 1
     FlexIng.CellBackColor = lnColorB
     FlexIng.Col = 2
     FlexIng.CellBackColor = lnColorB
     FlexIng.Col = 3
End Sub

Private Sub FlexVar_DblClick()
    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
    SetEstCon FlexVar
    If FlexVar.Col <> 3 Then
        Exit Sub
    Else
       MSHFlexGridEdit FlexVar, TextEditV, 32 ' Simula un espacio.
    End If
End Sub

Private Sub FlexVar_EnterCell()
     FlexVar.Col = 1
     FlexVar.CellBackColor = lnColorP
     FlexVar.Col = 2
     FlexVar.CellBackColor = lnColorP
     FlexVar.Col = 3
End Sub

Private Sub FlexVar_GotFocus()
    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
    If Me.FlexVar.Col <> 3 Then
        Exit Sub
    Else
       If TextEditV.Visible = False Then Exit Sub
        FlexVar = TextEditV
        SetEstCon FlexVar, True
        TextEditV.Visible = False
    End If
End Sub

Private Sub FlexVar_KeyPress(KeyAscii As Integer)
    If lnTipo <> gTipoOpeMantenimiento Or ((Not lbConceptoFijo) And KeyAscii <> 13) Then Exit Sub
    
    If NumerosEnteros(KeyAscii) = 0 Then
        Exit Sub
    ElseIf KeyAscii = 13 Then
        FlexVar_DblClick
    End If
    Me.FlexVar.Col = 3
    If Me.FlexVar.Col <> 3 Then
        Exit Sub
    Else
       MSHFlexGridEdit FlexVar, TextEditV, KeyAscii
    End If
End Sub

Private Sub FlexVar_LeaveCell()
     FlexVar.Col = 1
     FlexVar.CellBackColor = lnColorB
     FlexVar.Col = 2
     FlexVar.CellBackColor = lnColorB
     FlexVar.Col = 3
End Sub

Private Sub Form_Load()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    
    Caption = lsCaption
    
    Interprete_InI
    
    TxtPlanilla.rs = oPla.GetPlanillas(, True)
    
    Set rsP = oCon.GetConceptos(Trim(Str(RHConceptosTpoVisible.RHConceptosTpoVIngreso)))
    
    If Not (rsP.EOF And rsP.BOF) Then
        While Not rsP.EOF
            lsCadenaValida = lsCadenaValida & rsP.Fields(0) & ";"
            rsP.MoveNext
        Wend
    End If

    Flex_EnterCell
    lnColorB = -2147483643
    lnColorP = &HC0C000
    
    rsP.Close
    Set rsP = Nothing
    
    If lnTipo = gTipoOpeConsulta Then
        'Me.cmdAplicar.Enabled = False
        Me.cmdImprimir.Enabled = False
        Me.cmdTodosEmp.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        'Me.cmdAplicar.Enabled = False
        Me.cmdTodosEmp.Enabled = False
    ElseIf lnTipo = gTipoOpeMantenimiento Then
        Me.cmdImprimir.Enabled = False
    End If
End Sub

Private Sub IniFlexCon(f As MSHFlexGrid, psIni As String)
    Dim rsI As New ADODB.Recordset
    Dim oCon As DRHConcepto
    Dim I As Integer
    Set oCon = New DRHConcepto
    
    f.Rows = 1
    f.Rows = 2
    f.FixedRows = 1
    f.Cols = 4
    f.FixedCols = 0
    
    f.TextMatrix(0, 0) = "Codigo"
    f.TextMatrix(0, 1) = "Nombre"
    f.TextMatrix(0, 2) = "Estado"
    f.TextMatrix(0, 3) = "Monto.Ref."
    
    f.ColWidth(0) = 1
    f.ColWidth(2) = 800
    
    If lbConceptoFijo Then
        Set rsI = oCon.GetConceptosPlanilla(TxtPlanilla.Text, psIni, , 1)
        f.ColWidth(1) = 4300
        f.ColWidth(3) = 1600
    Else
        Set rsI = oCon.GetConceptosPlanilla(TxtPlanilla.Text, psIni, , 2)
        f.ColWidth(1) = 5900
        f.ColWidth(3) = 1
    End If
    
    
    While Not rsI.EOF
        If f.TextMatrix(f.Rows - 1, 0) <> "" Then f.Rows = f.Rows + 1
        f.TextMatrix(f.Rows - 1, 0) = rsI.Fields(0)
        f.TextMatrix(f.Rows - 1, 1) = rsI.Fields(1)
        rsI.MoveNext
    Wend
    
    Set oCon = Nothing
    rsI.Close
    Set rsI = Nothing
End Sub

'Private Sub IniFlexEmp(f As MSHFlexGrid)
'    Dim sqlI As String
'    Dim rsI As New ADODB.Recordset
'
'    sqlI = " Select E.cCodPers,E.cEmpCod,PE.cNomPers from Empleado E" _
'         & " Inner Join " & gcCentralPers & "Persona PE On E.cCodPers = PE.cCodPers where cEmpEst <> '3' ORDER BY E.cEmpCod"
'    rsI.Open sqlI, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    f.Rows = 2
'    f.FixedRows = 1
'    f.Cols = 3
'    f.FixedCols = 0
'
'    f.TextMatrix(0, 0) = "Nombre"
'    f.TextMatrix(0, 1) = "CodEmp"
'    f.TextMatrix(0, 2) = "CodPers"
'
'    f.ColWidth(0) = 4170
'    f.ColWidth(1) = 750
'    f.ColWidth(2) = 1000
'
'    If Not RSVacio(rsI) Then
'        While Not rsI.EOF
'            If f.TextMatrix(f.Rows - 1, 0) <> "" Then f.Rows = f.Rows + 1
'            f.TextMatrix(f.Rows - 1, 0) = Trim(rsI!cNomPers)
'            f.TextMatrix(f.Rows - 1, 1) = Trim(rsI!cEmpCod)
'            f.TextMatrix(f.Rows - 1, 2) = Trim(rsI!cCodPers)
'            rsI.MoveNext
'        Wend
'    End If
'
'    rsI.Close
'    Set rsI = Nothing
'
'End Sub

Private Sub SetCon(f As MSHFlexGrid, psCon As String, pnMonto As String)
    Dim I As Integer
    f.Col = 2
    For I = 1 To f.Rows - 1
        If f.TextMatrix(I, 0) = psCon Then
            f.TextMatrix(I, 2) = "."
            f.Row = I
            Set f.CellPicture = Pic.Picture
            f.TextMatrix(I, 3) = Format(pnMonto, "#,##0.00")
            I = f.Rows - 1
        End If
    Next I
   f.Col = 0
End Sub

Private Sub SetEstCon(f As MSHFlexGrid, Optional pbActualiza As Boolean = False)
    Dim sqlR As String
    Dim oCon As NRHConcepto
    Set oCon = New NRHConcepto
    
    If Me.RRHH.psCodigoEmpleado = "" Then
        MsgBox "Debe Ingresar un empleado valido.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If f.TextMatrix(f.Row, 0) = "" Then Exit Sub
    f.Col = 2
    If f.Text = "" Then
        oCon.AgregaConceptoRRHH Me.RRHH.psCodigoPersona, TxtPlanilla.Text, f.TextMatrix(f.Row, 0), f.TextMatrix(f.Row, 3), GetMovNro(gsCodUser, gsCodAge)
        f.Text = "."
        Set f.CellPicture = Pic.Picture
    Else
        If Not pbActualiza Then
            oCon.EliminaConceptoRRHH Me.RRHH.psCodigoPersona, TxtPlanilla.Text, f.TextMatrix(f.Row, 0)
            f.Text = ""
            f.TextMatrix(f.Row, 3) = ""
            Set f.CellPicture = LoadPicture
        Else
            oCon.ModificaConceptorrhh Me.RRHH.psCodigoPersona, TxtPlanilla.Text, f.TextMatrix(f.Row, 0), f.TextMatrix(f.Row, 3), GetMovNro(gsCodUser, gsCodAge)
        End If
    End If
    f.Col = 0
End Sub

Sub Flexing_GotFocus()
    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
    If Me.FlexIng.Col <> 3 Then
        Exit Sub
    Else
       If TextEdit.Visible = False Then Exit Sub
        FlexIng = TextEdit
        SetEstCon FlexIng, True
        TextEdit.Visible = False
    End If
End Sub

Sub Flexing_KeyPress(KeyAscii As Integer)
    If lnTipo <> gTipoOpeMantenimiento Or ((Not lbConceptoFijo) And KeyAscii <> 13) Then Exit Sub
    
    If NumerosEnteros(KeyAscii) = 0 Then
        Exit Sub
    ElseIf KeyAscii = 13 Then
        FlexIng_DblClick
    End If

    If Me.FlexIng.Col <> 3 Then
        Exit Sub
    Else
       MSHFlexGridEdit FlexIng, TextEdit, KeyAscii
    End If
End Sub


Private Sub TextEdit_Change()
    If TextEdit = "" Then
        TextEdit = "0"
    End If
End Sub

Sub textEdit_KeyDown(KeyCode As Integer, _
Shift As Integer)
    If Me.FlexIng.Col <> 3 Then
        Exit Sub
    Else
        EditKeyCode FlexIng, TextEdit, KeyCode, Shift
    End If
End Sub

Private Sub textEdit_KeyPress(KeyAscii As Integer)
   ' Elimina los retornos para quitar los pitidos.
    If KeyAscii = 13 Then TextEdit = CalEvaluaBil(TextEdit)
       
    If Me.FlexIng.Col <> 3 Then
        KeyAscii = 0
        Exit Sub
    Else
        If Me.FlexIng.Col = 1 Then
            If KeyAscii = Asc(vbCr) And NumerosEnteros(KeyAscii) = 0 Then
               KeyAscii = 0
            End If
        ElseIf Me.FlexIng.Col = 2 Then
               KeyAscii = CalValidaIngBil(KeyAscii)
        End If
    End If
End Sub

Private Sub TextEditA_Change()
    If TextEditA = "" Then
        TextEditA = "0"
    End If
End Sub

Private Sub TextEditA_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.FlexApo.Col <> 3 Then
        Exit Sub
    Else
        EditKeyCode FlexApo, TextEditA, KeyCode, Shift
    End If
End Sub

Private Sub TextEditA_KeyPress(KeyAscii As Integer)
   ' Elimina los retornos para quitar los pitidos.
    If KeyAscii = 13 Then TextEditA = CalEvaluaBil(TextEditA)
       
    If Me.FlexApo.Col <> 3 Then
        KeyAscii = 0
        Exit Sub
    Else
        If Me.FlexApo.Col = 1 Then
            If KeyAscii = Asc(vbCr) And NumerosEnteros(KeyAscii) = 0 Then
               KeyAscii = 0
            End If
        ElseIf Me.FlexApo.Col = 2 Then
               KeyAscii = CalValidaIngBil(KeyAscii)
        End If
    End If
End Sub

Private Sub TextEditD_Change()
    If TextEditD = "" Then
        TextEditD = "0"
    End If
End Sub

Private Sub TextEditD_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.FlexDes.Col <> 3 Then
        Exit Sub
    Else
        EditKeyCode FlexDes, TextEditD, KeyCode, Shift
    End If
End Sub

Private Sub TextEditD_KeyPress(KeyAscii As Integer)
   ' Elimina los retornos para quitar los pitidos.
    If KeyAscii = 13 Then TextEditD = CalEvaluaBil(TextEditD)
       
    If Me.FlexDes.Col <> 3 Then
        KeyAscii = 0
        Exit Sub
    Else
        If Me.FlexDes.Col = 1 Then
            If KeyAscii = Asc(vbCr) And NumerosEnteros(KeyAscii) = 0 Then
               KeyAscii = 0
            End If
        ElseIf Me.FlexDes.Col = 2 Then
               KeyAscii = CalValidaIngBil(KeyAscii)
        End If
    End If
End Sub

Private Sub TextEditV_Change()
    If TextEditV = "" Then
        TextEditV = "0"
    End If
End Sub

Private Sub TextEditV_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.FlexVar.Col <> 3 Then
        Exit Sub
    Else
        EditKeyCode FlexVar, TextEditV, KeyCode, Shift
    End If
End Sub

Private Sub TextEditV_KeyPress(KeyAscii As Integer)
   ' Elimina los retornos para quitar los pitidos.
    If KeyAscii = 13 Then TextEditV = CalEvaluaBil(TextEditV)
       
    If Me.FlexVar.Col <> 3 Then
        KeyAscii = 0
        Exit Sub
    Else
        If Me.FlexVar.Col = 1 Then
            If KeyAscii = Asc(vbCr) And NumerosEnteros(KeyAscii) = 0 Then
               KeyAscii = 0
            End If
        ElseIf Me.FlexVar.Col = 2 Then
               KeyAscii = CalValidaIngBil(KeyAscii)
        End If
    End If
End Sub

Private Sub IniPan(psCodPers As String)
    If Trim(psCodPers) <> "" Then Flex_EnterCell
End Sub

Public Sub Ini(pnTipoOpe As TipoOpe, pbConceptoFijo As Boolean, psCaption As String)
    lnTipo = pnTipoOpe
    lsCaption = psCaption
    lbConceptoFijo = pbConceptoFijo
    Me.Show 1
End Sub

'Private Sub ValidaConSue()
'    Dim I As Integer
'    Dim lnMonto As Currency
'
'    If lnTipo <> gTipoOpeMantenimiento Then Exit Sub
'    If Trim(cmbPlaCod) = "" Then Exit Sub
'    If Right(cmbPlaCod, 3) <> "E01" Then Exit Sub
'
'    lnMonto = 0
'    For I = 1 To FlexIng.Rows - 1
'        If 0 <> InStr(1, lsCadenaValida, FlexIng.TextMatrix(I, 0)) Then
'            If Trim(FlexIng.TextMatrix(I, 3)) <> "" And Trim(FlexIng.TextMatrix(I, 2)) <> "" Then
'                lnMonto = lnMonto + CCur(FlexIng.TextMatrix(I, 3))
'            End If
'        End If
'    Next I
'
'    If CCur(Me.RRHH.psSueldoContrato) <> lnMonto Then
'        If MsgBox("Existe una distribución errada del sueldo en los conceptos basicos. Desea Re-distribuirla ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'            AsignaValEmp RRHH.psCodigoPersona
'            IniPan RRHH.psCodigoPersona
'        End If
'    End If
'End Sub

'Private Sub AsignaValEmp(psCodPers As String)
'    Dim lnMontoSueBas As Currency
'    Dim lnCarFam As Currency
'    Dim lnMontoBonCon As Currency
'    Dim lsNum As String
'    Dim lnAux As Integer
'    Dim lnMontoTotal As Currency
'
'    Dim sql As String
'
'    If Left(Me.RRHH.psCodigoEmpleado, 1) = "E" Then
'        If Right(cmbPlaCod.Text, 3) <> "E01" Then Exit Sub
'
'        lnMontoSueBas = GetValorActualConcep(RRHH.psCodigoEmpleado, "I_SUE_BAS")
'        'If lnMontoSueBas = 0 Then lnMontoSueBas = GetSueBasNivEmp(RRHH.psCodigoEmpleado, psCodPers)
'
'        'If GetNumHijosEmp(psCodPers) = 0 Then
'         '   lnCarFam = 0
'        'Else
'        '    lsNum = GetValorFunLog("I_BON_CAR_FAM", psCodEmp, psCodPers, gdFecSis, gdFecSis, False, "")
'        '    lsNum = Left(lsNum, InStr(1, lsNum, "*") - 1)
'        '    lnCarFam = CCur(lsNum)
'
'            'Form1.Show 1
'        'End If
'
''        lnMontoTotal = 0
''        For I = 1 To FlexIng.Rows - 1
''            If FlexIng.TextMatrix(I, 0) <> "I_BON_CAR_FAM" And FlexIng.TextMatrix(I, 0) <> "I_SUE_BAS" And FlexIng.TextMatrix(I, 0) <> "I_BON_CONSOL" Then
''                If 0 <> InStr(1, lsCadenaValida, FlexIng.TextMatrix(I, 0)) Then
''                    If Trim(FlexIng.TextMatrix(I, 3)) <> "" Then
''                        lnMontoTotal = lnMontoTotal + CCur(FlexIng.TextMatrix(I, 3))
''                    End If
''                End If
''            End If
''        Next I
''
''        lnMontoTotal = lnMontoTotal + lnMontoSueBas + lnCarFam
''
''        lnMontoBonCon = CCur(lblMonSueCon) - lnMontoTotal
''
''
''        If lnMontoBonCon < 0 Then lnMontoBonCon = 0
''
''        dbCmact.BeginTrans
''            'sql = "Delete EmpCon Where cCodPers = '" & psCodPers & "' And cEmpCod = '" & psCodEmp & "' And cConcepCod In ('I_SUE_BAS','I_BON_CAR_FAM','I_BON_CONSOL')"
''            'dbCmact.Execute sql
''            sql = "Update EmpCon Set nMonto = " & lnMontoSueBas & " where cCodPers = '" & psCodPers & "' And cEmpCod = '" & psCodEmp & "' And cConcepCod = 'I_SUE_BAS'"
''            dbCmact.Execute sql
''            sql = "Update EmpCon Set nMonto = " & lnCarFam & " where cCodPers = '" & psCodPers & "' And cEmpCod = '" & psCodEmp & "' And cConcepCod = 'I_BON_CAR_FAM'"
''            dbCmact.Execute sql
''            sql = "Update EmpCon Set nMonto = " & lnMontoBonCon & " where cCodPers = '" & psCodPers & "' And cEmpCod = '" & psCodEmp & "' And cConcepCod = 'I_BON_CONSOL'"
''            dbCmact.Execute sql
''        dbCmact.CommitTrans
''
''    ElseIf Left(psCodEmp, 1) = "L" Then
''
''    ElseIf Left(psCodEmp, 1) = "S" Then
''
''    ElseIf Left(psCodEmp, 1) = "P" Then
''
'    Else
'        MsgBox "Tipo de Empleado no reconocido.", vbInformation, "Aviso"
'    End If
'End Sub

'Private Function GetValorActualConcep(psCodEmp As String, psCodCon As String) As Currency
'    Dim sqlC As String
'    Dim rsC As New ADODB.Recordset
'
'    sqlC = "Select nMonto from EmpCon Where cEmpCod = '" & psCodEmp & "' And cConcepCod = '" & psCodCon & "'"
'    'rsC.Open sqlC, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'   'If RSVacio(rsC) Then
'        GetValorActualConcep = 0
'   'Else
'   '     GetValorActualConcep = rsC!nMonto
'   ' End If
'
'    'RSCierra rsC
'
'End Function

Private Sub TxtPlanilla_EmiteDatos()
    Me.lblPlanilla.Caption = Me.TxtPlanilla.psDescripcion
    IniPan Me.RRHH.psCodigoPersona
    SSTab1.Tab = 0
End Sub

Private Sub oRRHH()

End Sub

Private Sub Limpia()
    RRHH.ClearScreen
    
End Sub
