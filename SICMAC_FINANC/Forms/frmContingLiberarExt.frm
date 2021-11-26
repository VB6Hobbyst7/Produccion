VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingLiberarExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contingencias: Extorno de Liberación"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   Icon            =   "frmContingLiberarExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabConting 
      Height          =   5460
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   9631
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   14
      TabHeight       =   520
      TabCaption(0)   =   "Filtro"
      TabPicture(0)   =   "frmContingLiberarExt.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkTodosAreas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBuscar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdExtLiberacion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCerrar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   345
         Left            =   9840
         TabIndex        =   14
         Top             =   4920
         Width           =   1050
      End
      Begin VB.CommandButton cmdExtLiberacion 
         Caption         =   "Extornar Liberación"
         Height          =   345
         Left            =   7920
         TabIndex        =   13
         Top             =   4920
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9840
         TabIndex        =   10
         Top             =   720
         Width           =   1050
      End
      Begin VB.Frame Frame3 
         Caption         =   " Fechas "
         Height          =   700
         Left            =   6960
         TabIndex        =   7
         Top             =   480
         Width           =   2700
         Begin MSMask.MaskEdBox txtFechaIni 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaFin 
            Height          =   315
            Left            =   1440
            TabIndex        =   9
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.CheckBox chkTodosAreas 
         Caption         =   "Todos"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo "
         Height          =   700
         Left            =   5040
         TabIndex        =   2
         Top             =   495
         Width           =   1860
         Begin VB.ComboBox cboTipoConting 
            Height          =   315
            ItemData        =   "frmContingLiberarExt.frx":0326
            Left            =   210
            List            =   "frmContingLiberarExt.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Area/Agencia "
         Height          =   700
         Left            =   120
         TabIndex        =   4
         Top             =   495
         Width           =   4815
         Begin Sicmact.TxtBuscar txtBuscarArea 
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblAreaDesc 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1290
            TabIndex        =   6
            Top             =   240
            Width           =   3420
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resultado"
         Height          =   3495
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   10815
         Begin MSDataGridLib.DataGrid DBGrdConting 
            Height          =   3135
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   0   'False
            ColumnHeaders   =   -1  'True
            HeadLines       =   2
            RowHeight       =   17
            RowDividerStyle =   4
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "cNumRegistro"
               Caption         =   "Nro Registro"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "dFechaReg"
               Caption         =   "Fecha Reg."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "cTpoConting"
               Caption         =   "Tipo Contingencia"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "cOrigen"
               Caption         =   "Origen"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "cTpoEvPerdida"
               Caption         =   "Tipo Evento de Perdida"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "cMoneda"
               Caption         =   "Moneda"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "nMonto"
               Caption         =   "Monto"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   " #,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "nCantInformesTecs"
               Caption         =   "Nº Informes"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "cCalif"
               Caption         =   "Calificación"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               SizeMode        =   1
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Size            =   800
               BeginProperty Column00 
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1904.882
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1695.118
               EndProperty
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmContingLiberarExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingLiberar
'** Descripción : Extorno Liberacion Contingencias creado segun RFC056-2012
'** Creación : JUEZ, 20120622 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim sNumRegistro As String

Private Sub cboTipoConting_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaIni.SetFocus
    End If
End Sub

Private Sub chkTodosAreas_Click()
    If chkTodosAreas.value = 1 Then
        txtBuscarArea.Enabled = False
        txtBuscarArea.BackColor = &H80000000
        lblAreaDesc.BackColor = &H80000000
        txtBuscarArea.Text = ""
        lblAreaDesc.Caption = ""
    Else
        txtBuscarArea.Enabled = True
        txtBuscarArea.BackColor = &H80000005
        lblAreaDesc.BackColor = &H80000005
    End If
End Sub

Private Sub cmdbuscar_Click()
    Set oConting = New DContingencia
    If cboTipoConting.Text <> "" And txtFechaIni.Text <> "" And txtFechaFin.Text <> "" Then
       If (chkTodosAreas.value = 0 And txtBuscarArea.Text <> "") Or chkTodosAreas.value = 1 Then
            Set rs = oConting.BuscaContigenciasLiberadas(CInt(Right(cboTipoConting.Text, 1)), txtFechaIni.Text, txtFechaFin.Text, IIf(chkTodosAreas.value = 0, Left(txtBuscarArea.Text, 3), ""))
        
            Set DBGrdConting.DataSource = rs
            DBGrdConting.Refresh
            Screen.MousePointer = 0
            If rs.RecordCount = 0 Then
              MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
              cmdExtLiberacion.Visible = False
            Else
              cmdExtLiberacion.Visible = True
            End If
        Else
            MsgBox "Faltan datos para la busqueda", vbInformation, "Aviso"
        End If
    Else
        MsgBox "Faltan datos para la busqueda", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdExtLiberacion_Click()
    sNumRegistro = DBGrdConting.Columns(0)
    If sNumRegistro <> "" Then
        Set oConting = New DContingencia
        
        If MsgBox("Está seguro de extornar la Liberacion de la Contingencia? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        Call oConting.ExtornaLiberacionContingencia(sNumRegistro)
        MsgBox "Se ha extornado la liberacion de la Contingencia", vbInformation, "Aviso"
        Call cmdbuscar_Click
    End If
End Sub

Private Sub Form_Load()
    Dim oAreas As DActualizaDatosArea
    Set oAreas = New DActualizaDatosArea
    txtBuscarArea.lbUltimaInstancia = False
    txtBuscarArea.psRaiz = "AREAS PARA SEGUIMIENTO DE CONTINGENCIAS"
    txtBuscarArea.rs = oAreas.GetAgenciasAreas()
    
    txtFechaIni.Text = gdFecSis
    txtFechaFin.Text = gdFecSis
End Sub

Private Sub txtBuscarArea_EmiteDatos()
    If txtBuscarArea = "" Then Exit Sub
    lblAreaDesc = txtBuscarArea.psDescripcion
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CDate(txtFechaFin.Text) < CDate(txtFechaIni.Text) Then
        MsgBox "Fecha Final No Puede Ser Menor que la fecha Inicial", vbInformation, "Aviso"
        txtFechaFin.SetFocus
        Exit Sub
        Else
            cmdBuscar.SetFocus
        End If
    End If
End Sub

Private Sub txtFechaFin_LostFocus()
    Dim sMsj As String
    sMsj = ValidaFecha(txtFechaFin.Text)
    If Not Trim(sMsj) = "" Then
        MsgBox sMsj, vbInformation, "Aviso"
        If txtFechaFin.Enabled Then txtFechaFin.SetFocus
        Exit Sub
    End If
    If CDate(txtFechaFin.Text) > gdFecSis Then
        MsgBox "Fecha Final No Puede Ser Mayor que la Fecha del Sistema", vbInformation, "Aviso"
        txtFechaFin.SetFocus
        Exit Sub
    End If
    If CDate(txtFechaFin.Text) < CDate(txtFechaIni.Text) Then
        MsgBox "Fecha Final No Puede Ser Menor que la fecha Inicial", vbInformation, "Aviso"
        txtFechaFin.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CDate(txtFechaIni.Text) > CDate(txtFechaFin.Text) Then
            MsgBox "Fecha de Inicio No Puede Ser Mayor que la fecha Final", vbInformation, "Aviso"
            txtFechaIni.SetFocus
            Exit Sub
        Else
            txtFechaFin.SetFocus
        End If
    End If
End Sub

Private Sub txtFechaIni_LostFocus()
    Dim sMsj As String
    sMsj = ValidaFecha(txtFechaIni.Text)
    If Not Trim(sMsj) = "" Then
        MsgBox sMsj, vbInformation, "Aviso"
        If txtFechaIni.Enabled Then txtFechaIni.SetFocus
        Exit Sub
    End If
    If CDate(txtFechaIni.Text) > gdFecSis Then
        MsgBox "Fecha de Inicio No Puede Ser Mayor que la Fecha del Sistema", vbInformation, "Aviso"
        txtFechaIni.SetFocus
        Exit Sub
    End If
End Sub
