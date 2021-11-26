VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingenciaReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contingencias: Registro"
   ClientHeight    =   8535
   ClientLeft      =   8400
   ClientTop       =   5085
   ClientWidth     =   10095
   Icon            =   "frmContingenciaReg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   8040
      Width           =   1050
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   8040
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTabConting 
      Height          =   8535
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Contingente Activo"
      TabPicture(0)   =   "frmContingenciaReg.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraActivo"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Contingente Pasivo"
      TabPicture(1)   =   "frmContingenciaReg.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FraActivo 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   8
         Top             =   480
         Width           =   5775
         Begin VB.ComboBox cboMontoA 
            Height          =   315
            ItemData        =   "frmContingenciaReg.frx":0342
            Left            =   1680
            List            =   "frmContingenciaReg.frx":034C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1110
            Width           =   615
         End
         Begin VB.ComboBox cboOrigenA 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtDescA 
            Height          =   735
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   1520
            Width           =   3015
         End
         Begin Sicmact.EditMoney txtMontoA 
            Height          =   315
            Left            =   2400
            TabIndex        =   3
            Top             =   1110
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblNroRegistroA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1680
            TabIndex        =   0
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Monto Aproximado:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1150
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Origen: "
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   760
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Nº de Registro:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   400
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7455
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   9855
         Begin VB.ComboBox cboMontoPerdidaSinGastos 
            Height          =   315
            ItemData        =   "frmContingenciaReg.frx":0375
            Left            =   1680
            List            =   "frmContingenciaReg.frx":037F
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   2280
            Width           =   735
         End
         Begin Sicmact.FlexEdit feMontos 
            Height          =   1755
            Left            =   1680
            TabIndex        =   41
            Top             =   2760
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   3096
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Concepto-Moneda-Monto"
            EncabezadosAnchos=   "385-3800-1200-2100"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3"
            TextStyleFixed  =   3
            ListaControles  =   "0-3-3-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R"
            FormatosEdit    =   "0-0-0-2"
            CantEntero      =   15
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   390
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "-"
            Height          =   360
            Left            =   9240
            TabIndex        =   43
            Top             =   3240
            Width           =   495
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "+"
            Height          =   360
            Left            =   9240
            TabIndex        =   42
            Top             =   2760
            Width           =   495
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   5040
            TabIndex        =   38
            Top             =   1130
            Visible         =   0   'False
            Width           =   3015
            Begin VB.OptionButton optReportado 
               Caption         =   "No Reportado"
               Height          =   255
               Index           =   2
               Left            =   1560
               TabIndex        =   40
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton optReportado 
               Caption         =   "Reportado"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.ComboBox cboMontoPRecuperado 
            Height          =   315
            ItemData        =   "frmContingenciaReg.frx":03B4
            Left            =   7080
            List            =   "frmContingenciaReg.frx":03BE
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   2280
            Width           =   735
         End
         Begin VB.ComboBox cboEvPerdidaP 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1800
            Width           =   4335
         End
         Begin VB.ComboBox cboMontoPProvision 
            Height          =   315
            ItemData        =   "frmContingenciaReg.frx":03F3
            Left            =   1680
            List            =   "frmContingenciaReg.frx":03FD
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   5160
            Width           =   735
         End
         Begin VB.ComboBox cboOrigenP 
            Height          =   315
            ItemData        =   "frmContingenciaReg.frx":0432
            Left            =   1680
            List            =   "frmContingenciaReg.frx":0434
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1320
            Width           =   3135
         End
         Begin Sicmact.EditMoney txtMontoPProvision 
            Height          =   315
            Left            =   2520
            TabIndex        =   32
            Top             =   5160
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
         End
         Begin MSMask.MaskEdBox txtFecOcurrencia 
            Height          =   285
            Left            =   1680
            TabIndex        =   22
            Top             =   840
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   15794175
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFecDescubrimiento 
            Height          =   285
            Left            =   4920
            TabIndex        =   24
            Top             =   840
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   15794175
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFecRegContable 
            Height          =   285
            Left            =   7920
            TabIndex        =   26
            Top             =   840
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   15794175
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.TextBox txtDescP 
            Height          =   975
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   5760
            Width           =   7575
         End
         Begin Sicmact.EditMoney txtMontoPRecuperado 
            Height          =   315
            Left            =   7920
            TabIndex        =   35
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin Sicmact.EditMoney txtMontoPerdidaSinGastos 
            Height          =   315
            Left            =   2520
            TabIndex        =   50
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin VB.Label Label20 
            Caption         =   "Monto de Perdida:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   2325
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Monto Recuperado:"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Monto Recuperado:"
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label lblPerdidaNeta 
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   46
            Top             =   6960
            Width           =   1455
         End
         Begin VB.Label lblMontoBruto 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7800
            TabIndex        =   45
            Top             =   4560
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Monto Bruto (MN) :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   44
            Top             =   4560
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Monto Provisionado:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   5160
            Width           =   1575
         End
         Begin VB.Label Label15 
            Caption         =   "Perdida Neta (MN) :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   6960
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Monto Recuperado:"
            Height          =   255
            Index           =   0
            Left            =   5520
            TabIndex        =   27
            Top             =   2325
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha Reg. Contable:"
            Height          =   375
            Left            =   6240
            TabIndex        =   25
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha Descubrimiento:"
            Height          =   375
            Left            =   3240
            TabIndex        =   23
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Fecha Ocurrencia:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Registro"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblNroRegistroP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1680
            TabIndex        =   10
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label9 
            Caption         =   "Evento de Perdida: "
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1845
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Montos:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Origen: "
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1365
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Nº de Registro:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   400
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   5760
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmContingenciaReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingenciaReg
'** Descripción : Registro de Contingencias creado segun RFC056-2012
'** Creación : JUEZ, 20120615 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim nTipoReg As Integer
Dim nValorOption As Integer 'CROB20170721
Dim oContTM As DContingenciaTipoMonto 'CROB20170724
Dim oContM As DContingenciaMontos 'CROB20170726


Public Function RegistrarContingencia(ByVal TipoRegistro As Integer)
    nTipoReg = TipoRegistro
    Set oConting = New DContingencia
    Set oGen = New DGeneral
    If TipoRegistro = gActivoContingente Then
        SSTabConting.TabVisible(0) = True
        SSTabConting.TabVisible(1) = False
        
        Set rs = oConting.ObtenerNroRegistroContingencia(gActivoContingente)
        Set oConting = Nothing
        
        lblNroRegistroA.Caption = rs!cNumRegProx
        
        Set rs = oGen.GetConstante(5080)
        Call CargaCombo(rs, cboOrigenA, 0, 1)
        Call CambiaTamañoCombo(cboOrigenA, 200)
    Else
        SSTabConting.TabVisible(0) = False
        SSTabConting.TabVisible(1) = True
        
        Set rs = oConting.ObtenerNroRegistroContingencia(gPasivoContingente)
        Set oConting = Nothing
        
        lblNroRegistroP.Caption = rs!cNumRegProx
    
        Set rs = oGen.GetConstante(5081)
        Call CargaCombo(rs, cboOrigenP, 0, 1)
        cboOrigenP.RemoveItem (0) 'TORE --> Quitamos la constante "DEMANDAS CIVILES, LABORALES Y ADMINISTRATIVAS" a solicitud del usuario
        nValorOption = 1 'CROB21070721 Reportados
        txtFecRegContable.Text = gdFecSis 'CROB21070727
    End If
    Me.Show 1
End Function

Private Sub txtDescA_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtDescA.Text <> "" Then
            If Len(txtDescA.Text) = 2 Then
                txtDescA.Text = Mid(txtDescA.Text, 1, IIf(Len(txtDescA.Text) = 0, 2, Len(txtDescA.Text)) - 2)
            End If
        End If
    End If
End Sub

Private Sub txtDescP_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtDescP.Text <> "" Then
            If Len(txtDescP.Text) = 2 Then
                txtDescP.Text = Mid(txtDescP.Text, 1, IIf(Len(txtDescP.Text) = 0, 2, Len(txtDescP.Text)) - 2)
            End If
        End If
    End If
End Sub

Private Sub txtMontoA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescA.SetFocus
    End If
End Sub
Private Sub txtMontoP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescP.SetFocus
    End If
End Sub

Private Sub txtMontoPerdidaSinGastos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMontoPRecuperado.ListIndex = 0
        cboMontoPProvision.ListIndex = 0
        cboMontoPRecuperado.SetFocus
        lblMontoBruto.Caption = FormatNumber(SumaPersonalizada(3) + IIf(Right(cboMontoPerdidaSinGastos.Text, 2) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text) * txtMontoPerdidaSinGastos.Text, txtMontoPerdidaSinGastos.Text), 2)
        lblPerdidaNeta.Caption = FormatNumber((CDbl(lblMontoBruto.Caption) - CDbl(txtMontoPRecuperado.Text)), 2)
    End If
End Sub

Private Sub txtMontoPerdidaSinGastos_LostFocus()
    cboMontoPRecuperado.ListIndex = 0
    
    cboMontoPRecuperado.SetFocus
    lblMontoBruto.Caption = FormatNumber(SumaPersonalizada(3) + IIf(Right(cboMontoPerdidaSinGastos.Text, 2) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text) * txtMontoPerdidaSinGastos.Text, txtMontoPerdidaSinGastos.Text), 2)
    lblPerdidaNeta.Caption = FormatNumber((CDbl(lblMontoBruto.Caption) - CDbl(txtMontoPRecuperado.Text)), 2)
End Sub

Private Sub txtMontoPProvision_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescP.SetFocus
    End If
End Sub



Public Sub CargaCombo(ByVal prsCombo As ADODB.Recordset, ByVal CtrlCombo As ComboBox, ByVal pnFiel1 As Integer, ByVal pnFiel2 As Integer)
    CtrlCombo.Clear
    While Not rs.EOF
        CtrlCombo.AddItem prsCombo.Fields(pnFiel1) & space(100) & prsCombo.Fields(pnFiel2) 'CROB20170721
        rs.MoveNext
    Wend
End Sub

Private Sub cboMontoA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoA.SetFocus
    End If
End Sub
Private Sub cboMontoPBruto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoPProvision.SetFocus
    End If
End Sub


Private Sub cboOrigenA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMontoA.SetFocus
    End If
End Sub

Private Sub cboOrigenP_Click()
    If Right(cboOrigenP.Text, 1) = gOrigenEventoPerdidas Then
        'Frame2.Visible = True 'CROB20170722
        HabilitarBotones (True) 'CROB20170727
        Set oConting = New DContingencia
        cboEvPerdidaP.Enabled = True
        
        'CROB20170722
        If nValorOption = 1 Then
            Set rs = oConting.ListarCtaContPasivosEventosPerdidaReportados
            Call CargaCombo(rs, cboEvPerdidaP, 0, 1)
        ElseIf nValorOption = 2 Then
            Set rs = oConting.ListarCtaContPasivosEventosPerdidaNOReportados
            Call CargaCombo(rs, cboEvPerdidaP, 0, 1)
        End If 'CROB20170722
        
        'Set rs = oGen.GetConstante(5082)
        'Call CargaCombo(rs, cboEvPerdidaP, 0, 1)
        'cboEvPerdidaP.SetFocus
    Else
        cboEvPerdidaP.Enabled = False
        cboEvPerdidaP.Clear
        optReportado.Item(1).value = True
        
        HabilitarBotones (False)
        'Call LimpiaPantalla
        'cboOrigenP.ListIndex = 0
        
        'cboMontoP.SetFocus
    End If
End Sub

Private Sub cboOrigenP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Right(cboOrigenP.Text, 1) = gOrigenEventoPerdidas Then
            HabilitarBotones (True) 'CROB20170727
            Set oConting = New DContingencia
            cboEvPerdidaP.Enabled = True
            
            'CROB20170722
            If nValorOption = 1 Then
                Set rs = oConting.ListarCtaContPasivosEventosPerdidaReportados
                Call CargaCombo(rs, cboEvPerdidaP, 0, 1)
            ElseIf nValorOption = 2 Then
                Set rs = oConting.ListarCtaContPasivosEventosPerdidaNOReportados
                Call CargaCombo(rs, cboEvPerdidaP, 0, 1)
            End If 'CROB20170722
        Else
            cboEvPerdidaP.Enabled = False
            cboEvPerdidaP.Clear
            optReportado.Item(1).value = True
            
            HabilitarBotones (False)
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
End Sub

Private Sub cmdRegistrar_Click()
    If ValidaDatos Then
        If MsgBox("Está seguro de registrar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        Set oConting = New DContingencia
        Set oContM = New DContingenciaMontos
        Dim cReg As String
        Dim reportado As String
        'Dim Validar As Boolean
        
        'Dim oMov As DMov
        'Set oMov = New DMov
        'gdFecha = gdFecSis
        'gsMovNro = oMov.GeneraMovNro(gdFecha, gsCodAge, gsCodUser)
        
        
        If nTipoReg = gActivoContingente Then
            'Call oConting.RegistraNuevaContigencia(Trim(lblNroRegistroA.Caption), CInt(Trim(Right(cboOrigenA.Text, 2))), CInt(Trim(Right(cboMontoA.Text, 2))), Format(txtMontoA.value, "#0.00"), Trim(txtDescA.Text), gsMovNro)
            cReg = "Activo"
            gsOpeCod = gRegistroActivoContingente
            gsGlosa = "Registro de Activo Contingente"
        Else
            If Trim(Right(cboOrigenP.Text, 2)) = gOrigenEventoPerdidas Then
                'Call oConting.RegistraNuevaContigencia(Trim(lblNroRegistroP.Caption), CInt(Trim(Right(cboOrigenP.Text, 2))), CInt(Trim(Right(.Text, 2))), Format(txtMontoP.value, "#0.00"), Trim(txtDescP.Text), gsMovNro, CInt(Trim(Right(cboEvPerdidaP.Text, 2))))
                'Validar =
                Call RegistrarMontos(lblNroRegistroP.Caption)  'CROB20170726
                
                'If Validar = False Then Exit Sub 'TORE 06/03/2018
                
                reportado = IIf(optReportado.Item(1).value = True, 1, 2)
                Set rs = oConting.RegistraNuevaContigencia(lblNroRegistroP.Caption, 2, Trim(Right(cboOrigenP.Text, 2)), Trim(Replace(Left(cboEvPerdidaP.Text, 2), ".", " ")) _
                                                     , reportado, txtDescP.Text, lblMontoBruto.Caption, cboMontoPRecuperado.ItemData(cboMontoPRecuperado.ListIndex) _
                                                     , IIf(txtMontoPRecuperado.Text = "0", "0", txtMontoPRecuperado.Text), cboMontoPProvision.ItemData(cboMontoPProvision.ListIndex), IIf(txtMontoPProvision.Text = "0", "0", txtMontoPProvision.Text) _
                                                     , lblPerdidaNeta.Caption, Trim(Right(cboEvPerdidaP.List(cboEvPerdidaP.ListIndex), 25)), txtFecOcurrencia.Text, txtFecDescubrimiento.Text, txtFecRegContable.Text _
                                                     , gsCodUser, gsCodAge, cboMontoPerdidaSinGastos.ItemData(cboMontoPerdidaSinGastos.ListIndex), txtMontoPerdidaSinGastos.Text)
            Else
                'Call oConting.RegistraNuevaContigencia(Trim(lblNroRegistroP.Caption), CInt(Trim(Right(cboOrigenP.Text, 2))), CInt(Trim(Right(cboMontoP.Text, 2))), Format(txtMontoP.value, "#0.00"), Trim(txtDescP.Text), gsMovNro)
            End If
            cReg = "Pasivo"
            gsOpeCod = gRegistroPasivoContingente
            gsGlosa = "Registro de Pasivo Contingente"
        End If
        
        'oMov.BeginTrans
        'oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa
        'gnMovNro = oMov.GetnMovNro(gsMovNro)
        'oMov.CommitTrans
        'Set oMov = Nothing
        
        If rs!nResultado = 1 Then
            MsgBox "El Contigente " & cReg & " se ha registrado exitosamente", vbInformation, "Aviso"
        Else
            MsgBox "No se pudo realizar el Registro Correctamente, Comuniquese con TI", vbInformation, "Aviso"
        End If
        Call LimpiaPantalla
    End If
End Sub

Private Sub feMontos_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
If feMontos.col = 1 Then
    lblMontoBruto.Caption = FormatNumber(SumaPersonalizada(3) + IIf(Right(cboMontoPerdidaSinGastos.Text, 2) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text) * txtMontoPerdidaSinGastos.Text, txtMontoPerdidaSinGastos.Text), 2)
End If
End Sub


Private Sub Form_Load()
    CentraForm Me
End Sub


'CROB20170721
Private Sub optReportado_Click(index As Integer)
    Set oConting = New DContingencia
    
    If Right(cboOrigenP.Text, 1) = "2" Then
        If index = 1 Then
            Set rs = oConting.ListarCtaContPasivosEventosPerdidaReportados
            Call CargaCombo(rs, cboEvPerdidaP, 0, 1)
        ElseIf index = 2 Then
            Set rs = oConting.ListarCtaContPasivosEventosPerdidaNOReportados
            Call CargaCombo(rs, cboEvPerdidaP, 0, 1)
        End If
    End If
End Sub 'CROB20170721


Public Sub LimpiaPantalla()
    Dim i As Integer
    lblNroRegistroA.Caption = ""
    lblNroRegistroP.Caption = ""
    txtFecOcurrencia.Text = "  /  /    "
    txtFecDescubrimiento.Text = "  /  /    "
    txtFecRegContable.Text = gdFecSis 'CROB21070727
    txtMontoA.value = 0
    txtMontoPProvision.value = 0
    txtMontoPProvision.Enabled = False
    txtMontoPRecuperado.value = 0
    txtMontoPRecuperado.Enabled = False
    txtDescA.Text = ""
    txtDescP.Text = ""
    cboOrigenA.ListIndex = -1
    cboOrigenP.ListIndex = -1
    cboEvPerdidaP.ListIndex = -1
    cboMontoA.ListIndex = -1
    cboMontoPProvision.ListIndex = -1
    cboMontoPRecuperado.ListIndex = -1
    txtMontoPerdidaSinGastos.value = 0
    
    'TORE 06-03-2018
    txtMontoPerdidaSinGastos.Enabled = False
    cboMontoPerdidaSinGastos.ListIndex = -1
    'TORE END
    
    For i = (feMontos.Rows - 1) To 0 Step -1
        feMontos.EliminaFila (i)
    Next i
    
    lblMontoBruto.Caption = "0.00"
    lblPerdidaNeta.Caption = "0.00"
    Call NuevaContingencia
End Sub


Public Function ValidaDatos() As Boolean
    'Contigente Activo
    If nTipoReg = gActivoContingente Then
        If cboOrigenA.ListIndex = -1 Then
            MsgBox "Falta seleccionar el Origen del Activo", vbInformation, "Aviso"
            cboOrigenA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If cboMontoA.ListIndex = -1 Then
            MsgBox "Falta seleccionar el tipo de moneda", vbInformation, "Aviso"
            cboMontoA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtMontoA.value = 0 Then
            MsgBox "Falta ingresar el monto aprox.", vbInformation, "Aviso"
            txtMontoA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Len(Trim(Me.txtDescA)) = 0 Then
            MsgBox "Falta ingresar una descripcion", vbInformation, "Aviso"
            txtDescA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    'Contingente Pasivo
    Else
        If cboOrigenP.ListIndex = -1 Then
            MsgBox "Falta seleccionar el Origen del Activo", vbInformation, "Aviso"
            cboOrigenP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(Right(cboOrigenP.Text, 2)) = gOrigenEventoPerdidas Then
            If cboEvPerdidaP.ListIndex = -1 Then
                MsgBox "Falta seleccionar el evento de pérdida", vbInformation, "Aviso"
                cboEvPerdidaP.SetFocus
                ValidaDatos = False
                Exit Function
            End If
        End If
        
        'TORE 06-03-2018
        If cboMontoPerdidaSinGastos.ListIndex = -1 Then
            MsgBox "El Monto de Perdida es obligatorio, seleccione el tipo de moneda del monto de perdida.", vbInformation, "Aviso"
            cboMontoPerdidaSinGastos.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If txtMontoPerdidaSinGastos.value = 0 Then
            MsgBox "Ingrese un valor en el Monto de Perdida.", vbInformation, "Aviso"
            txtMontoPerdidaSinGastos.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        'END TORE
        
        If cboMontoPProvision.ListIndex = -1 Then
            MsgBox "Seleccione el tipo de moneda del Monto Provisionado", vbInformation, "Aviso"
            cboMontoPProvision.SetFocus
            ValidaDatos = False
            Exit Function
        End If

        If cboMontoPRecuperado.ListIndex = -1 Then
            MsgBox "El Monto de Perdida Recuperado es obligatorio, seleccione el tipo de moneda del Monto de Perdida Recuperado.", vbInformation, "Aviso"
            cboMontoPRecuperado.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        'Comentado por el usuario TORE17032018
'         If txtMontoPRecuperado.value = 0 Then
'            MsgBox "Ingrese un valor en Monto de Perdida Recuperado.", vbInformation, "Aviso"
'            txtMontoPRecuperado.SetFocus
'            ValidaDatos = False
'            Exit Function
'        End If
        'END TORE
        
        
        If lblPerdidaNeta.Caption = "0" Then
            MsgBox "El monto recuperado debe ser inferior al monto bruto", vbInformation, "Aviso"
            txtMontoPRecuperado.value = 0
            txtMontoPRecuperado.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Len(Trim(Me.txtDescP)) = 0 Then
            MsgBox "Falta ingresar una descripcion", vbInformation, "Aviso"
            txtDescP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    ValidaDatos = True
End Function

Public Sub NuevaContingencia()
    Set oConting = New DContingencia
    Set oGen = New DGeneral
    If nTipoReg = gActivoContingente Then
        SSTabConting.TabVisible(0) = True
        SSTabConting.TabVisible(1) = False
        
        Set rs = oConting.ObtenerNroRegistroContingencia(gActivoContingente)
        Set oConting = Nothing
        
        lblNroRegistroA.Caption = rs!cNumRegProx
        
        Set rs = oGen.GetConstante(5080)
        Call CargaCombo(rs, cboOrigenA, 0, 1)
        Call CambiaTamañoCombo(cboOrigenA, 200)
    Else
        SSTabConting.TabVisible(0) = False
        SSTabConting.TabVisible(1) = True
        
        Set rs = oConting.ObtenerNroRegistroContingencia(gPasivoContingente)
        Set oConting = Nothing
        
        lblNroRegistroP.Caption = rs!cNumRegProx
        
        Set rs = oGen.GetConstante(5081)
        Call CargaCombo(rs, cboOrigenP, 0, 1)
        cboOrigenP.RemoveItem (0) 'TORE --> Quitamos la constante "DEMANDAS CIVILES, LABORALES Y ADMINISTRATIVAS" a solicitud del usuario
    End If
End Sub

Private Sub txtMontoPRecuperado_Change()
    Dim nMoneda As Integer
    Dim nMontoRecup As Double
    
    nMoneda = CInt(cboMontoPRecuperado.ItemData(cboMontoPRecuperado.ListIndex))
    
    If nMoneda = 1 Then
        nMontoRecup = CDbl(txtMontoPRecuperado.Text)
    Else
         nMontoRecup = CDbl(txtMontoPRecuperado.Text) * ObtenerTipoCambioFecha(txtFecRegContable.Text)
    End If
    
    If nMontoRecup <= CDbl(lblMontoBruto.Caption) Then
        lblPerdidaNeta.Caption = FormatNumber((CDbl(lblMontoBruto.Caption) - nMontoRecup), 2)
    Else
        lblPerdidaNeta.Caption = "0"
        MsgBox "El monto recuperado debe ser inferior al monto bruto", vbInformation, "Aviso"
    End If
End Sub

Private Sub CargarTiposDeMontos()
    Set oContTM = New DContingenciaTipoMonto
    Set rs = oContTM.ListarTipoMontoPasivoContingente
    feMontos.CargaCombo rs
    Set oContTM = Nothing
    Set rs = Nothing
End Sub

Private Sub CargarTipoDeMoneda()
    Set oContTM = New DContingenciaTipoMonto
    Set rs = oContTM.ListarTipoMoneda
    feMontos.CargaCombo rs
    Set oContTM = Nothing
    Set rs = Nothing
End Sub

Private Sub feMontos_KeyPress(KeyAscii As Integer)
    If feMontos.col = 1 Then
        CargarTiposDeMontos
    ElseIf feMontos.col = 2 Then
        CargarTipoDeMoneda
    End If
    'lblMontoBruto.Caption = FormatNumber(SumaPersonalizada(3), 2)
End Sub

Private Sub feMontos_Click()
    If feMontos.col = 1 Then
        CargarTiposDeMontos
    ElseIf feMontos.col = 2 Then
        CargarTipoDeMoneda
    End If
End Sub

Private Sub cmdNuevo_Click()
    If txtMontoPerdidaSinGastos.Text = 0 Then
        MsgBox "Proporcione el Monto Perdida", vbInformation, "Aviso"
    Else
        feMontos.AdicionaFila , Val(feMontos.TextMatrix(feMontos.Rows - 1, 0)) + 1
        feMontos.SetFocus
        lblMontoBruto.Caption = FormatNumber(SumaPersonalizada(3) + IIf(Right(cboMontoPerdidaSinGastos.Text, 2) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text) * txtMontoPerdidaSinGastos.Text, txtMontoPerdidaSinGastos.Text), 2)
        lblPerdidaNeta.Caption = FormatNumber((CDbl(lblMontoBruto.Caption) - CDbl(txtMontoPRecuperado.Text)), 2)
        
    End If
End Sub

Private Sub cmdEliminar_Click()
If feMontos.TextMatrix(feMontos.Row, 0) <> "" Then
        feMontos.EliminaFila (feMontos.Row)
        lblMontoBruto.Caption = FormatNumber(SumaPersonalizada(3) + IIf(Right(cboMontoPerdidaSinGastos.Text, 2) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text) * txtMontoPerdidaSinGastos.Text, txtMontoPerdidaSinGastos.Text), 2)
        lblPerdidaNeta.Caption = FormatNumber((CDbl(lblMontoBruto.Caption) - CDbl(txtMontoPRecuperado.Text)), 2)
        txtMontoPRecuperado.value = 0
        'txtMontoPRecuperado.SetFocus
End If
End Sub

Private Sub feMontos_OnCellChange(pnRow As Long, pnCol As Long)
If feMontos.col = 3 Then
    lblMontoBruto.Caption = FormatNumber(SumaPersonalizada(3) + IIf(Right(cboMontoPerdidaSinGastos.Text, 2) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text) * txtMontoPerdidaSinGastos.Text, txtMontoPerdidaSinGastos.Text), 2)
    lblPerdidaNeta.Caption = FormatNumber((CDbl(lblMontoBruto.Caption) - CDbl(txtMontoPRecuperado.Text)), 2)
End If
End Sub

Private Sub feMontos_LostFocus()
    lblPerdidaNeta.Caption = FormatNumber((CDbl(lblMontoBruto.Caption) - CDbl(txtMontoPRecuperado.Text)), 2)
    lblMontoBruto.Caption = FormatNumber(SumaPersonalizada(3) + IIf(Right(cboMontoPerdidaSinGastos.Text, 2) = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text) * txtMontoPerdidaSinGastos.Text, txtMontoPerdidaSinGastos.Text), 2) 'aqui
    lblPerdidaNeta.Caption = FormatNumber((CDbl(lblMontoBruto.Caption) - CDbl(txtMontoPRecuperado.Text)), 2)
End Sub

'Cambiar de Function a Sub de ser requerido (TORE 07/03/2018)
Private Sub RegistrarMontos(ByVal cNumReg As String) 'As Boolean
    Dim i As Integer
    Dim rowCount As Integer
    Dim nTipoMontoPCID As Integer
    Dim nMoneda As Integer
    Dim nMonto As Double
    rowCount = feMontos.Rows
    
    If feMontos.TextMatrix(1, 1) <> "" And feMontos.TextMatrix(1, 2) <> "" And feMontos.TextMatrix(1, 3) <> "" Then
        Set oContM = New DContingenciaMontos
        For i = 1 To rowCount - 1 Step 1
            nTipoMontoPCID = CInt(RTrim(Right(feMontos.TextMatrix(i, 1), 3)))
            nMoneda = CInt(RTrim(Right(feMontos.TextMatrix(i, 2), 1)))
            nMonto = CDbl(feMontos.TextMatrix(i, 3))
            
            Call oContM.RegistrarMontoContingencia(cNumReg, nTipoMontoPCID, nMonto, nMoneda, IIf(nMoneda = 2, ObtenerTipoCambioFecha(txtFecRegContable.Text), 0))
        Next i
        Set oContM = Nothing
        'RegistrarMontos = True 'TORE07032018
    'Else
        'MsgBox "Ingresar datos en la tabla Montos", vbInformation, "Aviso" 'Paso el
        'RegistrarMontos = True 'TORE07032018
    End If
    
End Sub

Private Sub HabilitarBotones(ByVal pbHabilitar As Boolean)
    Frame2.Visible = pbHabilitar
    cmdNuevo.Enabled = pbHabilitar
    cmdEliminar.Enabled = pbHabilitar
    cmdRegistrar.Enabled = pbHabilitar
    cmdCancelar.Enabled = pbHabilitar
    Dim i As Integer
    For i = (feMontos.Rows - 1) To 0 Step -1
        feMontos.EliminaFila (i)
    Next i
    lblMontoBruto.Caption = "0.00"
    lblPerdidaNeta.Caption = "0.00"
End Sub

Public Function SumaPersonalizada(ByVal pnCol As Long) As Currency
Dim i As Integer
Dim nMoneda As Integer
Dim lnSuma As Currency
lnSuma = 0

For i = 1 To feMontos.Rows - 1
    If IsNumeric(feMontos.TextMatrix(i, pnCol)) Then
        nMoneda = CInt(RTrim(Right(feMontos.TextMatrix(i, 2), 1)))
        If nMoneda = 1 Then
            lnSuma = lnSuma + CCur(IIf(feMontos.TextMatrix(i, pnCol) = "", "0", feMontos.TextMatrix(i, pnCol))) ' + CCur(IIf(txtMontoPerdidaSinGastos.value = 0, 0, txtMontoPerdidaSinGastos.Text)) 'TORE20032018
        Else
            lnSuma = lnSuma + (CCur(IIf(feMontos.TextMatrix(i, pnCol) = "", "0", feMontos.TextMatrix(i, pnCol))) * ObtenerTipoCambioFecha(txtFecRegContable.Text)) ' + txtMontoPerdidaSinGastos.Text
            'lnSuma = lnSuma + (CCur(IIf(feMontos.TextMatrix(i, pnCol) = "", "0", feMontos.TextMatrix(i, pnCol))) + CCur(IIf(txtMontoPerdidaSinGastos.value = 0, 0, txtMontoPerdidaSinGastos.Text)) * ObtenerTipoCambioFecha(txtFecRegContable.Text))
        End If
    End If
Next
SumaPersonalizada = lnSuma '+ CCur(IIf(txtMontoPerdidaSinGastos.value = 0, 0, txtMontoPerdidaSinGastos.Text)) 'TORE20032018
End Function


Private Sub cboMontoPRecuperado_Click()
    If cboMontoPRecuperado.ListIndex = -1 Then
        txtMontoPRecuperado.Enabled = False
    Else
        txtMontoPRecuperado.Enabled = True
        txtMontoPRecuperado.Text = "0"
    End If
End Sub

'TORE 06-03-2018
Private Sub cboMontoPerdidaSinGastos_Click()
    If cboMontoPerdidaSinGastos.ListIndex = -1 Then
        txtMontoPerdidaSinGastos.Enabled = False
    Else
        txtMontoPerdidaSinGastos.Enabled = True
        txtMontoPerdidaSinGastos.Text = "0"
    End If
End Sub

'END TORE

Private Sub cboMontoPProvision_Click()
    If cboMontoPProvision.ListIndex = -1 Then
        txtMontoPProvision.Enabled = False
    Else
        txtMontoPProvision.Enabled = True
        txtMontoPProvision.Text = "0"
    End If
End Sub

Private Sub txtFecDescubrimiento_LostFocus()
    If Not IsDate(txtFecDescubrimiento) Then
        MsgBox "Verifique Dia, Mes, Año , Fecha Incorrecta", vbInformation, "Aviso"
    End If
End Sub

Private Sub txtFecOcurrencia_LostFocus()
    If Not IsDate(txtFecOcurrencia) Then
        MsgBox "Verifique Dia, Mes, Año , Fecha Incorrecta", vbInformation, "Aviso"
    End If
End Sub

Private Sub txtFecRegContable_LostFocus()
    If Not IsDate(txtFecRegContable) Then
        MsgBox "Verifique Dia, Mes, Año , Fecha Incorrecta", vbInformation, "Aviso"
    End If
End Sub

Private Function ObtenerTipoCambioFecha(ByVal psFecha As String) As Double
    Set oContTM = New DContingenciaTipoMonto
    Set rs = oContTM.ObtenerTipoCambioFecha(psFecha)
    ObtenerTipoCambioFecha = rs!nValFijo
End Function



'CROB20170724


