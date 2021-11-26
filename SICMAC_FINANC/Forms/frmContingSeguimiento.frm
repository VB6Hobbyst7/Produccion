VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingSeguimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contingencias: Seguimiento"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   Icon            =   "frmContingSeguimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabConting 
      Height          =   5340
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   9419
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Seguimiento de Contingencias"
      TabPicture(0)   =   "frmContingSeguimiento.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGrdConting"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBuscar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCerrar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdExportar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdLiberarConting"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdVerInformeTec"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdRegInformeTec"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkTodosAreas"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdDesestimar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CommandButton cmdDesestimar 
         Caption         =   "Desestimar"
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
         Left            =   4200
         TabIndex        =   16
         Top             =   4800
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CheckBox chkTodosAreas 
         Caption         =   "Todos"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   470
         Width           =   855
      End
      Begin VB.CommandButton cmdRegInformeTec 
         Caption         =   "Registrar IT"
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
         Left            =   120
         TabIndex        =   14
         Top             =   4800
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdVerInformeTec 
         Caption         =   "Ver ITs"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   4800
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdLiberarConting 
         Caption         =   "Liberar"
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
         Left            =   2760
         TabIndex        =   12
         Top             =   4800
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
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
         Left            =   5640
         TabIndex        =   11
         Top             =   4800
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
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
         Left            =   9240
         TabIndex        =   10
         Top             =   4800
         Width           =   1050
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
         Left            =   9120
         TabIndex        =   8
         Top             =   720
         Width           =   1050
      End
      Begin VB.Frame Frame3 
         Caption         =   " Estado "
         Height          =   700
         Left            =   7110
         TabIndex        =   6
         Top             =   480
         Width           =   1860
         Begin VB.ComboBox cboEstado 
            Height          =   315
            ItemData        =   "frmContingSeguimiento.frx":0326
            Left            =   210
            List            =   "frmContingSeguimiento.frx":0336
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo "
         Height          =   700
         Left            =   5160
         TabIndex        =   4
         Top             =   480
         Width           =   1860
         Begin VB.ComboBox cboTipoConting 
            Height          =   315
            ItemData        =   "frmContingSeguimiento.frx":03C7
            Left            =   210
            List            =   "frmContingSeguimiento.frx":03D4
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Area/Agencia "
         Height          =   700
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4935
         Begin Sicmact.TxtBuscar txtBuscarArea 
            Height          =   315
            Left            =   240
            TabIndex        =   2
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
            TabIndex        =   3
            Top             =   240
            Width           =   3420
         End
      End
      Begin MSDataGridLib.DataGrid DBGrdConting 
         Height          =   3135
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   10275
         _ExtentX        =   18124
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
Attribute VB_Name = "frmContingSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingSeguimiento
'** Descripción : Seguimiento de Contingencias creado segun RFC056-2012
'** Creación : JUEZ, 20120621 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim sNumRegistro As String

Private Sub chkTodosAreas_Click()
    If chkTodosAreas.value = 1 Then
        txtBuscarArea.Enabled = False
        txtBuscarArea.BackColor = &H80000000
        lblAreaDesc.BackColor = &H80000000
    Else
        txtBuscarArea.Enabled = True
        txtBuscarArea.BackColor = &H80000005
        lblAreaDesc.BackColor = &H80000005
    End If
End Sub

Private Sub cmdbuscar_Click()
    Set oConting = New DContingencia
    If cboTipoConting.Text <> "" And cboEstado.Text <> "" Then
       If (chkTodosAreas.value = 0 And txtBuscarArea.Text <> "") Or chkTodosAreas.value = 1 Then
            Set rs = oConting.BuscaContigenciasSeguimiento(IIf(chkTodosAreas.value = 0, txtBuscarArea.Text, ""), Right(cboTipoConting.Text, 5), Right(cboEstado.Text, 8))
        
            Set DBGrdConting.DataSource = rs
            DBGrdConting.Refresh
            Screen.MousePointer = 0
            If rs.RecordCount = 0 Then
              MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
              VisualizarBotones (False)
            Else
              VisualizarBotones (True)
            End If
        Else
            MsgBox "Faltan datos para la busqueda", vbInformation, "Aviso"
        End If
    Else
        MsgBox "Faltan datos para la busqueda", vbInformation, "Aviso"
    End If
End Sub

Private Sub VisualizarBotones(ByVal phHabilita As Boolean)
    Dim MatGrupos() As String
    Dim i As Integer
    If phHabilita = False Then
        cmdRegInformeTec.Visible = phHabilita
        cmdVerInformeTec.Visible = phHabilita
        cmdLiberarConting.Visible = phHabilita
        cmdDesestimar.Visible = phHabilita
        cmdExportar.Visible = phHabilita
    Else
        MatGrupos = Split(gsGrupoUsu, ",")
        For i = 0 To CStr(UBound(MatGrupos))
            If MatGrupos(i) = "GRUPO CONTABILIDAD I" Or MatGrupos(i) = "GRUPO CONTABILIDAD II" Then
                cmdRegInformeTec.Visible = phHabilita
                cmdVerInformeTec.Visible = phHabilita
                cmdLiberarConting.Visible = False
                cmdDesestimar.Visible = phHabilita
                cmdExportar.Visible = phHabilita
                cmdDesestimar.Left = 2760
                cmdExportar.Left = 4200
                Exit For
            ElseIf MatGrupos(i) = "GRUPO TESORERIA I" Or MatGrupos(i) = "GRUPO TESORERIA II" Then
                cmdRegInformeTec.Visible = False
                cmdVerInformeTec.Visible = False
                cmdLiberarConting.Visible = phHabilita
                cmdDesestimar.Visible = False
                cmdExportar.Visible = False
                cmdLiberarConting.Left = 120
                Exit For
            End If
        Next i
        If cmdRegInformeTec.Visible = False And cmdVerInformeTec.Visible = False And cmdLiberarConting.Visible = False And cmdDesestimar.Visible = False And cmdExportar.Visible = False Then
            If gsCodArea = "021" Then
                cmdRegInformeTec.Visible = phHabilita
                cmdVerInformeTec.Visible = phHabilita
                cmdLiberarConting.Visible = False
                cmdDesestimar.Visible = phHabilita
                cmdExportar.Visible = phHabilita
                cmdDesestimar.Left = 2760
                cmdExportar.Left = 4200
            ElseIf gsCodArea = "025" Then
                cmdRegInformeTec.Visible = False
                cmdVerInformeTec.Visible = False
                cmdLiberarConting.Visible = phHabilita
                cmdDesestimar.Visible = False
                cmdExportar.Visible = False
                cmdLiberarConting.Left = 120
            End If
        End If
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdDesestimar_Click()
    sNumRegistro = DBGrdConting.Columns(0)
    If sNumRegistro <> "" Then
        Set oConting = New DContingencia
        Set rs = oConting.BuscaContigenciaSeleccionada(sNumRegistro)
        If rs!nestado = 1 Then
            MsgBox "La contigencia seleccionada no tiene ningún informe técnico", vbInformation, "Aviso"
            Exit Sub
        ElseIf rs!nestado = 3 Then
            MsgBox "La contigencia está liberada", vbInformation, "Aviso"
            Exit Sub
        ElseIf rs!nestado = 4 Then
            MsgBox "La contigencia ya está desestimada", vbInformation, "Aviso"
            Exit Sub
        Else
            If oConting.VerificaContingenciaSiPuedeDesestimarse(sNumRegistro) Then 'Para Activos que su calificación sea diferente de Cierto,Para Pasivo que esté como Remoto
                frmContingDesestimar.Inicio sNumRegistro
            Else
                Dim sMensaje As String
                If Left(sNumRegistro, 1) = gActivoContingente Then
                    sMensaje = "Sólo pueden desestimarse las Contingencias Activas que tengan calificación diferente de CIERTA"
                Else
                    sMensaje = "Sólo pueden desestimarse las Contingencias Pasivas calificadas como REMOTO"
                End If
                MsgBox sMensaje, vbInformation, "Aviso!"
            End If
        End If
    End If
    'Call cmdbuscar_Click
End Sub

Private Sub cmdExportar_Click()
    Call ExportaContingenciasExcel
End Sub

Private Sub cmdLiberarConting_Click()
    sNumRegistro = DBGrdConting.Columns(0)
    If sNumRegistro <> "" Then
        Set oConting = New DContingencia
        Set rs = oConting.BuscaContigenciaSeleccionada(sNumRegistro)
        If rs!nestado = 1 Then
            MsgBox "La contigencia seleccionada no tiene ningun informe técnico", vbInformation, "Aviso"
            Exit Sub
        ElseIf rs!nestado = 3 Then
            MsgBox "La contigencia ya está liberada", vbInformation, "Aviso"
            Exit Sub
        ElseIf rs!nestado = 4 Then
            MsgBox "La contigencia está desestimada", vbInformation, "Aviso"
            Exit Sub
        Else
            If oConting.VerificaContingenciaSiPuedeLiberarse(sNumRegistro) Then 'Para Activos que este calificado como Cierta,Para Pasivo como Probable
                frmContingLiberar.Liberar sNumRegistro
            Else
                Dim sMensaje As String
                If Left(sNumRegistro, 1) = gActivoContingente Then
                    sMensaje = "Sólo pueden liberarse las Contingencias Activas calificadas como CIERTA"
                Else
                    sMensaje = "Sólo pueden liberarse las Contingencias Pasivas calificadas como PROBABLE"
                End If
                MsgBox sMensaje, vbInformation, "Aviso!"
            End If
        End If
    End If
    'Call cmdbuscar_Click
End Sub

Private Sub cmdRegInformeTec_Click()
    sNumRegistro = DBGrdConting.Columns(0)
    If sNumRegistro <> "" Then
        Set oConting = New DContingencia
        Set rs = oConting.BuscaContigenciaSeleccionada(sNumRegistro)
        If rs!nestado <> 3 Then
            If rs!nestado <> 4 Then
                If Left(sNumRegistro, 1) = "1" Then
                    frmContingInformeTecReg.RegistroActivo sNumRegistro
                Else
                    frmContingInformeTecReg.RegistroPasivo sNumRegistro
                End If
            Else
                MsgBox "La contigencia ya está desestimada", vbInformation, "Aviso!"
            End If
        Else
            MsgBox "La contigencia ya está liberada", vbInformation, "Aviso!"
        End If
    End If
    'Call cmdbuscar_Click
End Sub

Private Sub cmdVerInformeTec_Click()
    sNumRegistro = DBGrdConting.Columns(0)
    If sNumRegistro <> "" Then
        Set oConting = New DContingencia
        Set rs = oConting.BuscaContigenciaSeleccionada(sNumRegistro)
        If rs!nestado = 1 Then
            MsgBox "La contigencia seleccionada no tiene ningun informe técnico", vbInformation, "Aviso"
            Exit Sub
        End If
        frmContingInformeTecCons.Extorno sNumRegistro, 1
    End If
    'Call cmdbuscar_Click
End Sub

Private Sub Form_Load()
    Dim oAreas As DActualizaDatosArea
    Set oAreas = New DActualizaDatosArea
    txtBuscarArea.lbUltimaInstancia = False
    txtBuscarArea.psRaiz = "AREAS PARA SEGUIMIENTO DE CONTINGENCIAS"
    txtBuscarArea.rs = oAreas.GetAgenciasAreas()
End Sub

Private Sub txtBuscarArea_EmiteDatos()
    If txtBuscarArea = "" Then Exit Sub
    lblAreaDesc = txtBuscarArea.psDescripcion
End Sub

Private Sub ExportaContingenciasExcel()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim PrimeraLinea As Integer
    Dim nFilaSig As Integer
    Dim i As Integer
    
    Set oConting = New DContingencia
    Set rs = oConting.BuscaContigenciasSeguimiento(IIf(chkTodosAreas.value = 0, txtBuscarArea.Text, ""), Right(cboTipoConting.Text, 5), Right(cboEstado.Text, 8))
    
    On Error GoTo ErrorExportaContingencias
    
    If rs.RecordCount > 0 Then
        Set fs = New Scripting.FileSystemObject
        Set xlsAplicacion = New Excel.Application
        
        lsNomHoja = "Seguimiento Contingencias"
        lsFile = "SeguimientoContingencias"
        
        lsArchivo = "\spooler\" & lsFile & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & ".xls"
        If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
            Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
        Else
            MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
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
        
        xlHoja1.Cells(3, 1) = gdFecSis
        xlHoja1.Cells(5, 2) = IIf(txtBuscarArea.Text = "", "Todos", lblAreaDesc.Caption)
        xlHoja1.Cells(5, 5) = Trim(Left(cboTipoConting.Text, Len(cboTipoConting.Text) - 3))
        xlHoja1.Cells(5, 8) = Trim(Left(cboEstado.Text, Len(cboEstado.Text) - 7))
        
        PrimeraLinea = 8
        
        For i = 0 To rs.RecordCount - 1
            nFilaSig = PrimeraLinea + i
            xlHoja1.Cells(nFilaSig, 1) = rs!dFechaReg
            xlHoja1.Cells(nFilaSig, 2) = rs!cTpoConting
            xlHoja1.Cells(nFilaSig, 3) = rs!cOrigen
            xlHoja1.Cells(nFilaSig, 4) = rs!cTpoEvPerdida
            xlHoja1.Cells(nFilaSig, 5) = rs!cmoneda
            xlHoja1.Cells(nFilaSig, 6) = rs!nMonto
            xlHoja1.Cells(nFilaSig, 7) = rs!nCantInformesTecs
            xlHoja1.Cells(nFilaSig, 8) = rs!cCalif
            rs.MoveNext
        Next i
        
        Dim psArchivoAGrabarC As String
        
        xlHoja1.SaveAs App.path & lsArchivo
         psArchivoAGrabarC = App.path & lsArchivo
         xlsAplicacion.Visible = True
         xlsAplicacion.Windows(1).Visible = True
         Set xlsAplicacion = Nothing
         Set xlsLibro = Nothing
         Set xlHoja1 = Nothing
        MsgBox "Reporte Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
    Else
        MsgBox "No hay Datos", vbInformation, "Aviso"
    End If
    
    Exit Sub
ErrorExportaContingencias:
    MsgBox Err.Description, vbInformation, "Error!!"
End Sub
