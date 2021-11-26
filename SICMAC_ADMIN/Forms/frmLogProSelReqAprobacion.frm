VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogProSelReqAprobacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobacion"
   ClientHeight    =   6360
   ClientLeft      =   1095
   ClientTop       =   1605
   ClientWidth     =   10215
   Icon            =   "frmLogProSelReqAprobacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   120
      TabIndex        =   4
      Top             =   -60
      Width           =   9975
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Requerimientos para el Mes"
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
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   2370
      End
   End
   Begin VB.CommandButton CmndAprobar 
      Caption         =   "Aprobar"
      Height          =   375
      Left            =   7020
      TabIndex        =   2
      Top             =   5895
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   8550
      TabIndex        =   1
      Top             =   5895
      Width           =   1455
   End
   Begin VB.Frame FrmAreasAgencias 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   10035
      Begin TabDlg.SSTab SSTab1 
         Height          =   2445
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   4313
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "    Agencias             "
         TabPicture(0)   =   "frmLogProSelReqAprobacion.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSFlexAA"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "  Requerimientos       "
         TabPicture(1)   =   "frmLogProSelReqAprobacion.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSFlexR"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexAA 
            Height          =   1830
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   3228
            _Version        =   393216
            FixedCols       =   0
            BackColorSel    =   15269875
            ForeColorSel    =   8388608
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483633
            WordWrap        =   -1  'True
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BandDisplay     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexR 
            Height          =   1815
            Left            =   -74880
            TabIndex        =   10
            Top             =   480
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   3201
            _Version        =   393216
            BackColor       =   -2147483633
            FixedCols       =   0
            BackColorSel    =   14545407
            ForeColorSel    =   8388608
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483633
            WordWrap        =   -1  'True
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BandDisplay     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista para Aprobacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   9975
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   2055
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorSel    =   16775645
         ForeColorSel    =   8388608
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         ScrollBars      =   2
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
   End
   Begin VB.Image imgNN 
      Height          =   240
      Left            =   7320
      Picture         =   "frmLogProSelReqAprobacion.frx":0902
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   7680
      Picture         =   "frmLogProSelReqAprobacion.frx":0C44
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLogProSelReqAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nTipo As Integer

Public Sub Inicio(ByVal pnTipo As Integer)
    nTipo = pnTipo
    Me.Show 1
End Sub

Private Sub cboMes_Click()
    Select Case nTipo
        Case 1, 2
            CargarRequerimiento
        Case 3
            CargarDatosProcesos
    End Select
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmndAprobar_Click()
On Error GoTo CmndAprobarErr
    Dim oCon As DConecta, sSQL As String, i As Integer, nBan As Boolean
    Set oCon = New DConecta
    Select Case nTipo
        Case 1, 2
            If Val(MSFlex.TextMatrix(1, 2)) = 0 Then
                MsgBox "No Existen Requerimientos", vbInformation, "Aviso"
                Exit Sub
            End If
        Case 3
            If Val(MSFlex.TextMatrix(1, 4)) = 0 Then
                MsgBox "No Existen Requerimientos", vbInformation, "Aviso"
                Exit Sub
            End If
    End Select
    
If MsgBox(" ¿ Aprobar los requerimientos indicados ? " + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
    If oCon.AbreConexion Then
        i = 1
        With MSFlex
            Do While i < .Rows
                .Col = 1
                .row = i
                If .CellPicture = imgOK Then
                    nBan = True
                    Select Case nTipo
                        Case 1
                            sSQL = "update LogProSelReq set nVoBoPresupuesto = 1 where nProSelReqNro in ( select nProSelReqNro from LogProSelReqDetalle  where nAnio = " & Year(gdFecSis) & " and nPrecioUnitario > 0 and nEstado = 1 and nMesEje = " & cboMes.ListIndex + 1 & " ) "
                        Case 2
                            sSQL = "update LogProSelReq set nSubGerenciaAdm = 1 where nProSelReqNro in ( select nProSelReqNro from LogProSelReqDetalle  where nAnio = " & Year(gdFecSis) & " and nPrecioUnitario > 0 and nEstado = 1 and nMesEje = " & cboMes.ListIndex + 1 & " ) "
                        Case 3
                            sSQL = "update LogProcesoSeleccion set nVoBoPresupuesto=1 where nProSelNro = " & .TextMatrix(i, 4)
                    End Select
                    oCon.Ejecutar sSQL
                End If
                i = i + 1
            Loop
            .Col = 0
            .row = 1
            .ColSel = .Cols - 1
        End With
        oCon.CierraConexion
    End If
End If
    If nBan Then
        MsgBox "Aprobacion Registrada", vbInformation
        Select Case nTipo
            Case 1, 2
                CargarRequerimiento
            Case 3
                CargarDatosProcesos
        End Select
    Else
        MsgBox "Debe Seleccionar algun Item para Aprobar", vbInformation, "Aviso"
    End If
Exit Sub

CmndAprobarErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To 12
        cboMes.AddItem Format("01/" & i & "/" & Year(gdFecSis), "mmmm")
    Next
    cboMes.ListIndex = Month(gdFecSis) - 1
    CentraForm Me
    Select Case nTipo
        Case 1
            Caption = "Visto Bueno de Presupuesto"
            CmndAprobar.Top = 3240
            cmdSalir.Top = 3240
            FrmAreasAgencias.Visible = False
            Height = 4095
            CargarRequerimiento
        Case 2
            Caption = "Aprobacion de SubGerencia"
            CmndAprobar.Top = 5895
            cmdSalir.Top = 5895
            FrmAreasAgencias.Visible = True
            Height = 6735
            CargarRequerimiento
        Case 3
            Caption = "Visto Bueno de Presupuesto al Proceso de Seleccion"
            CmndAprobar.Top = 3240
            cmdSalir.Top = 3240
            FrmAreasAgencias.Visible = False
            Height = 4095
            CargarDatosProcesos
    End Select
End Sub

Private Sub CargarDatosProcesos()
    On Error GoTo CargarDatosProcesosErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, i As Integer, sSQL As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select s.cArchivoBases, t.cProSelTpoDescripcion, s.nProSelNro, s.nPlanAnualNro, s.nPlanAnualAnio, " & _
                " s.nPlanAnualMes, s.nProSelTpoCod, s.nProSelSubTpo, nNroProceso, c.cConsDescripcion, " & _
                " s.nObjetoCod , s.nMoneda, s.nProSelMonto, s.nProSelEstado, cSintesis, nModalidadCompra " & _
                " from LogProcesoSeleccion s " & _
                " inner join LogProSelTpo t on s.nProSelTpoCod = t.nProSelTpoCod " & _
                " left outer join constante c on s.nObjetoCod=c.nConsValor and c.nConsCod = 9048" & _
                " where nVoBoPresupuesto = 0 and s.nProSelEstado > -1 and s.nPlanAnualMes=" & cboMes.ListIndex + 1 & " and nPlanAnualAnio = " & Year(gdFecSis) & _
                " order by nNroProceso "
        Set Rs = oCon.CargaRecordSet(sSQL)
        
        FormaFlex
        i = 0
        Do While Not Rs.EOF
            i = i + 1
            InsRow MSFlex, i
            With MSFlex
'                .RowHeight(i) = 800
                .Col = 1
                .row = i
                Set .CellPicture = imgNN
                .TextMatrix(i, 2) = Rs!cProSelTpoDescripcion
                .TextMatrix(i, 3) = Rs!cSintesis
                .TextMatrix(i, 4) = Rs!nProselNro
                .TextMatrix(i, 6) = FNumero(Rs!nProSelMonto)
            End With
            Rs.MoveNext
        Loop
        MSFlex.row = 1
        MSFlex.Col = 0
        MSFlex.ColSel = MSFlex.Cols - 1
        oCon.CierraConexion
    End If
    Exit Sub
CargarDatosProcesosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarRequerimiento()
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nNroReq As Integer
Dim sSQL As String

sSQL = ""
FormaFlex

If oConn.AbreConexion Then

'    sSQL = "select x.nProSelReqNro, d.cBSCod, b.cBSDescripcion,u.cUnidad, nCantidad, x.nMesEje, x.nAnio, d.nPrecioUnitario, x.cSustento from LogProSelReqDetalle d " & _
           " inner join LogProSelReq x on d.nProSelReqNro = x.nProSelReqNro " & _
           " inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
           " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
           " where d.nAnio = " & Year(gdFecSis) & " and d.nPrecioUnitario > 0 and d.nEstado = 1 and x.nMesEje = " & cboMes.ListIndex + 1 & _
           " and (select count(*) from LogProSelAprobacion where nEstadoAprobacion=0 and nProSelReqNro = x.nProSelReqNro) = 0 "
    sSQL = " select d.nProSelReqNro,  d.cBSCod, b.cBSDescripcion,u.cUnidad, nCantidad=sum(nCantidad), x.nMesEje, x.nAnio, d.nPrecioUnitario " & _
           " from LogProSelReqDetalle d " & _
           " inner join LogProSelReq x on d.nProSelReqNro = x.nProSelReqNro " & _
           " inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
           " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
           " where d.nAnio = " & Year(gdFecSis) & " and d.nPrecioUnitario > 0 and x.nMesEje = " & cboMes.ListIndex + 1 & " and " & _
           " (select count(*) from LogProSelAprobacion where nEstadoAprobacion=0 and nProSelReqNro = x.nProSelReqNro) = 0  "
    Select Case nTipo
        Case 1
            sSQL = sSQL & " and d.nEstado =1 and nVoBoPresupuesto = 0 "
        Case 2
            sSQL = sSQL & "  and d.nEstado =1  and nVoBoPresupuesto =1 and nSubGerenciaAdm = 0 "
    End Select
    'sSQL = sSQL & " group by d.cBSCod, b.cBSDescripcion,u.cUnidad, nCantidad, x.nMesEje, x.nAnio, d.nPrecioUnitario, x.cSustento "
    'sSQL = sSQL & " group by d.cBSCod, b.cBSDescripcion,u.cUnidad, x.nMesEje, x.nAnio, d.nPrecioUnitario "
    sSQL = sSQL & " group by d.nProSelReqNro, d.cBSCod, b.cBSDescripcion,u.cUnidad, x.nMesEje, x.nAnio, d.nPrecioUnitario "
    If oConn.AbreConexion Then
       Set Rs = oConn.CargaRecordSet(sSQL)
       oConn.CierraConexion
    End If
    If Not Rs.EOF Then
       i = 0
          Do While Not Rs.EOF
             i = i + 1
             If nNroReq <> Rs!nProSelReqNro Then
                nNroReq = Rs!nProSelReqNro
                With MSFlex
                    InsRow MSFlex, i
                    .Col = 0
                    .row = i
                    .RowHeight(i) = 300
                    .CellFontSize = 10
                    .CellFontBold = True
                    .TextMatrix(i, 0) = "+" ' "-"
                    .Col = 1
                    Set .CellPicture = imgNN
                    .TextMatrix(i, 2) = Rs!nProSelReqNro
                    .TextMatrix(i, 3) = Rs!cBSDescripcion
                    '.TextMatrix(i, 3) = Rs!cSustento
                    .TextMatrix(i, 6) = FNumero(CalculaTotalReq(Rs!nProSelReqNro))
                End With
                
                i = i + 1
                InsRow MSFlex, i
                With MSFlex
                    .RowHeight(i) = 0
                    .TextMatrix(i, 2) = Rs!cBSCod
                    .TextMatrix(i, 3) = Rs!cBSDescripcion
                    .TextMatrix(i, 4) = Rs!cUnidad
                    .TextMatrix(i, 5) = Rs!nCantidad
                    .TextMatrix(i, 6) = FNumero(Rs!nPrecioUnitario)
                End With

            Else
                InsRow MSFlex, i
                With MSFlex
                    .RowHeight(i) = 0
                    .row = i
                    .Col = 1
                    Set .CellPicture = imgNN
                    .TextMatrix(i, 2) = Rs!nProSelReqNro
                    .TextMatrix(i, 3) = Rs!cBSDescripcion
                    .TextMatrix(i, 4) = Rs!cUnidad
                    .TextMatrix(i, 5) = Rs!nCantidad
                    .TextMatrix(i, 6) = FNumero(Rs!nPrecioUnitario)
                End With
            End If
            Rs.MoveNext
          Loop
          MSFlex.row = 1
          MSFlex.ColSel = MSFlex.Cols - 1
    End If
'    MSFlex.SetFocus
End If
End Sub

Sub FormaFlex()
Dim i As Integer
Select Case nTipo
    Case 1, 2
        MSFlex.Clear
        MSFlex.Rows = 2
        MSFlex.RowHeight(-1) = 280
        MSFlex.RowHeight(0) = 300
        MSFlex.RowHeight(1) = 8
'        Select Case nTipo
'            Case 1
'                MSFlex.ColWidth(0) = 0
'            Case 2
                MSFlex.ColWidth(0) = 300
'        End Select
        MSFlex.ColWidth(1) = 300:                           MSFlex.TextMatrix(0, 1) = ""
        MSFlex.ColWidth(2) = 1000:                          MSFlex.TextMatrix(0, 2) = "Cod"
        MSFlex.ColWidth(3) = IIf(nTipo = 1, 4400, 4100):    MSFlex.TextMatrix(0, 3) = "Descripción"
        MSFlex.ColWidth(4) = 1000:                          MSFlex.TextMatrix(0, 4) = "   U. Medida" ':   MSFlex.ColAlignment(3) = 4
        MSFlex.ColWidth(5) = 1200:                          MSFlex.TextMatrix(0, 5) = "Cantidad"
        MSFlex.ColWidth(6) = 1500:                          MSFlex.TextMatrix(0, 6) = "P. Ref"
    Case 3
        With MSFlex
            .Clear
            .Rows = 2
            .TextMatrix(0, 0) = "":                        .ColWidth(0) = 0
            .TextMatrix(0, 1) = "    ":                    .ColWidth(1) = 250
            .TextMatrix(0, 2) = "Proceso":                 .ColWidth(2) = 3000
            .TextMatrix(0, 3) = "Sintesis":                .ColWidth(3) = 4900
            .TextMatrix(0, 4) = "U. Medida":               .ColWidth(4) = 0
            .TextMatrix(0, 5) = "Cantidad":                .ColWidth(5) = 0:
            .TextMatrix(0, 6) = "Monto":                   .ColWidth(6) = 1200
            .WordWrap = True
        End With
End Select

    With MSFlexAA
        .Clear
        .Rows = 2
        .TextMatrix(0, 0) = "Codigo":           .ColWidth(0) = 600
        .TextMatrix(0, 1) = "Descripcion":      .ColWidth(1) = 9000
    End With
    
    With MSFlexR
        .Clear
        .Rows = 2
        .TextMatrix(0, 0) = "Nro Req":            .ColWidth(0) = 600
        .TextMatrix(0, 1) = "Responsable":        .ColWidth(1) = 9000
    End With

End Sub

Private Function CalculaTotalReq(ByVal pnProSelReqNro As Integer)
On Error GoTo CalculaTotalReqErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, nSuma As Currency, sSQL As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nPrecioUnitario from LogProSelReqDetalle where nProSelReqNro = " & pnProSelReqNro
        Set Rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
    End If
    Do While Not Rs.EOF
        nSuma = nSuma + Rs!nPrecioUnitario
        Rs.MoveNext
    Loop
    CalculaTotalReq = nSuma
    Exit Function
CalculaTotalReqErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function


Private Sub MSFlex_Click()
If MSFlex.Col = 1 Then
    With MSFlex
        .Col = 1
        If .CellPicture = imgOK Then
           Set .CellPicture = imgNN
           '&H80000005
           .CellBackColor = "&H80000005"
        ElseIf .CellPicture = imgNN Then
           Set .CellPicture = imgOK
           .BackColorBand(.Col) = "&H00EAFFFF"
           '&H00EAFFFF
           
        End If
        '.Col = 0
        '.ColSel = .Cols - 1
    End With
End If
End Sub

Private Sub MSFlex_DblClick()
On Error GoTo MSItemErr
    Dim i As Integer, bTipo As Boolean
    
    'If nTipo = 1 Then Exit Sub
    
If MSFlex.Col = 0 Then
    With MSFlex
        If Trim(.TextMatrix(.row, 0)) = "-" Then
           .TextMatrix(.row, 0) = "+"
           i = .row + 1
           bTipo = True
        ElseIf Trim(.TextMatrix(.row, 0)) = "+" Then
           .TextMatrix(.row, 0) = "-"
           i = .row + 1
           bTipo = False
        Else
            Exit Sub
        End If
        Do While i < .Rows
            If Trim(.TextMatrix(i, 0)) = "+" Or Trim(.TextMatrix(i, 0)) = "-" Then
                Exit Sub
            End If
            
            If bTipo Then
                .RowHeight(i) = 0
            Else
                .RowHeight(i) = 300
            End If
            i = i + 1
        Loop
    End With
End If
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarDatosAA(ByVal pnProSelReqNro As Integer)
On Error GoTo CargarDatosAAErr
    Dim sSQL As String, Rs As ADODB.Recordset, oCon As DConecta, i As Integer
    Set oCon = New DConecta
    sSQL = " select a.cAgeCod, a.cAgeDescripcion, x.careaDescripcion " & _
           " from LogProSelReqDetalle d " & _
           " inner join rrhh z on d.cPersCod = z.cPersCod " & _
           " inner join agencias a on z.cAgenciaAsig = a.cAgeCod " & _
           " inner join areas x on z.careacod = x.careacod " & _
           " where d.nProSelReqNro = '" & pnProSelReqNro & "' and d.nEstado = 1 and " & _
           " (select count(*) from LogProSelAprobacion where nEstadoAprobacion=0 and nProSelReqNro = d.nProSelReqNro) = 0 " & _
           " group by a.cAgeCod, a.cAgeDescripcion, x.careaDescripcion "
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
    End If
    MSFlexAA.Rows = 2
    Do While Not Rs.EOF
        With MSFlexAA
            i = i + 1
            InsRow MSFlexAA, i
            .TextMatrix(i, 0) = Rs!cAgeCod
            .TextMatrix(i, 1) = Rs!cAgeDescripcion
        End With
        Rs.MoveNext
    Loop
    Exit Sub
CargarDatosAAErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarDatosR(ByVal pnProSelReqNro As Integer)
On Error GoTo CargarDatosRErr
    Dim sSQL As String, Rs As ADODB.Recordset, oCon As DConecta, i As Integer
    Set oCon = New DConecta
    sSQL = " select d.nProSelReqNro, cPersNombre=replace(p.cPersNombre,'/',' ') " & _
           " from LogProSelReqDetalle d " & _
           " inner join Persona p on d.cPersCod = p.cPersCod " & _
           " where d.nProSelReqNro = '" & pnProSelReqNro & "' and nEstado = 1 and " & _
           " (select count(*) from LogProSelAprobacion where nEstadoAprobacion=0 and nProSelReqNro = d.nProSelReqNro) = 0"
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
    End If
    MSFlexR.Rows = 2
    Do While Not Rs.EOF
        With MSFlexR
            i = i + 1
            InsRow MSFlexR, i
            .TextMatrix(i, 0) = Rs!nProSelReqNro
            .TextMatrix(i, 1) = Rs!cPersNombre
        End With
        Rs.MoveNext
    Loop
    Exit Sub
CargarDatosRErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub MSFlex_GotFocus()
    If Len(MSFlex.TextMatrix(MSFlex.row, 2)) = 10 Then Exit Sub
    CargarDatosAA Val(MSFlex.TextMatrix(MSFlex.row, 2))
    CargarDatosR Val(MSFlex.TextMatrix(MSFlex.row, 2))
End Sub

Private Sub MSFlex_RowColChange()
    If Len(MSFlex.TextMatrix(MSFlex.row, 2)) = 10 Then Exit Sub
    CargarDatosAA Val(MSFlex.TextMatrix(MSFlex.row, 2))
    CargarDatosR Val(MSFlex.TextMatrix(MSFlex.row, 2))
End Sub

Private Sub MSFlex_SelChange()
    If Len(MSFlex.TextMatrix(MSFlex.row, 2)) = 10 Then Exit Sub
    CargarDatosAA Val(MSFlex.TextMatrix(MSFlex.row, 2))
    CargarDatosR Val(MSFlex.TextMatrix(MSFlex.row, 2))
End Sub
