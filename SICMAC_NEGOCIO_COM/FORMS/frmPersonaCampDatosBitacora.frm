VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersonaCampDatosBitacora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualización de Datos"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "frmPersonaCampDatosBitacora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos de Clientes"
      TabPicture(0)   =   "frmPersonaCampDatosBitacora.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pgbExcel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "feBitacora"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRestaurar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdExportar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSalir"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9120
         TabIndex        =   10
         Top             =   3840
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7920
         TabIndex        =   9
         Top             =   3840
         Width           =   1050
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
         Height          =   360
         Left            =   1320
         TabIndex        =   8
         Top             =   3840
         Width           =   1050
      End
      Begin VB.CommandButton cmdRestaurar 
         Caption         =   "Restaurar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   1050
      End
      Begin SICMACT.FlexEdit feBitacora 
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4471
         Cols0           =   34
         HighLight       =   1
         EncabezadosNombres=   $"frmPersonaCampDatosBitacora.frx":0326
         EncabezadosAnchos=   $"frmPersonaCampDatosBitacora.frx":04E7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-C-C-L-C-C-L-C-L-L-L-L-L-L-C-L-L-L-L-L-L-L-L-L-L-L-L-L-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         SelectionMode   =   1
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin VB.Frame Frame1 
         Caption         =   " Cliente: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10095
         Begin VB.TextBox lblPersNombre 
            Height          =   300
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox lblDOITipo 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "Mostrar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8880
            TabIndex        =   5
            Top             =   240
            Width           =   1050
         End
         Begin SICMACT.TxtBuscar TxtBCodPers 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1740
            _ExtentX        =   3281
            _ExtentY        =   503
            Appearance      =   1
            Text            =   "."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin VB.Label lblDOINro 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7320
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "DOI:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   3
            Top             =   300
            Width           =   375
         End
      End
      Begin ComctlLib.ProgressBar pgbExcel 
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   3900
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmPersonaCampDatosBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmPersonaCampDatosBitacora
'** Descripción : Formulario para visualizar y administrar la actualizacion de datos de las personas
'**               que pertenecen a la Campaña "Actualiza tus Datos" según TI-ERS134-2013
'** Creación : JUEZ, 20131016 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Dim oDPers As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset

Public Sub InicioActualizar()
    cmdRestaurar.Visible = True
    HabilitaControles False
    Me.Show 1
End Sub
Public Sub InicioConsultar()
    cmdRestaurar.Visible = False
    cmdExportar.Left = 180
    HabilitaControles False
    Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    HabilitaControles False
    cmdRestaurar.Enabled = False
End Sub

Private Sub cmdRestaurar_Click()
    If Trim(TxtBCodPers.Text) = "" Or feBitacora.TextMatrix(feBitacora.row, 0) = "" Then
        MsgBox "No existen datos para restaurar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Se va a proceder a restaurar los datos de la Persona, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Dim rsPersona As ADODB.Recordset, rsDocs As ADODB.Recordset
    
    Set rsPersona = LlenaRecordSet_rsPersona
    Set rsDocs = LlenaRecordSet_rsDocs
    
    Set oDPers = New COMDPersona.DCOMPersonas
    Call oDPers.GrabarDatosPersonaCampDatosBitacora(TxtBCodPers.Text, rsPersona, rsDocs, feBitacora.TextMatrix(feBitacora.row, 32), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
    Set oDPers = Nothing
    MsgBox "Los datos fueron restaurados", vbInformation, "Aviso"
    'Call cmdCancelar_Click
    Call cmdMostrar_Click
    cmdRestaurar.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub feBitacora_Click()
    cmdRestaurar.Enabled = True
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    If Trim(TxtBCodPers.Text) = "" Then
        Call LimpiaFlex(feBitacora)
        Exit Sub
    End If
    
    Set oDPers = New COMDPersona.DCOMPersonas
    Set rs = oDPers.RecuperaDatosPersonaCampDatosBitacora(TxtBCodPers.Text)
    If rs.RecordCount > 0 Then
        Set rs = oDPers.RecuperaDatosPersona_Basic(TxtBCodPers.Text)
        lblPersNombre.Text = rs("cPersNombre")
        lblDOITipo.Text = rs("cDOITipo")
        lblDOINro.Caption = rs("nDOINro")
        Set rs = Nothing
        Call LimpiaFlex(feBitacora)
        Call HabilitaControles(True)
        cmdMostrar.SetFocus
    Else
        MsgBox "El cliente no tiene actualizaciones para mostrar", vbInformation, "Aviso"
        Limpiar
    End If
    Set oDPers = Nothing
End Sub

Private Sub Limpiar()
    TxtBCodPers.Text = ""
    lblPersNombre.Text = ""
    lblDOITipo.Text = ""
    lblDOINro.Caption = ""
    TxtBCodPers.Enabled = True
    Call LimpiaFlex(feBitacora)
End Sub

Private Sub cmdMostrar_Click()
    Set oDPers = New COMDPersona.DCOMPersonas
    Set rs = oDPers.RecuperaDatosPersonaCampDatosBitacora(TxtBCodPers.Text)
    If rs.RecordCount > 0 Then
        feBitacora.rsFlex = rs
        feBitacora.TopRow = 1
        Set rs = Nothing
        cmdRestaurar.Enabled = False
    End If
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
TxtBCodPers.Enabled = Not pbHabilita
cmdMostrar.Enabled = pbHabilita
'cmdRestaurar.Enabled = pbHabilita
cmdExportar.Enabled = pbHabilita
End Sub

Private Sub cmdExportar_Click()

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
    
    If Trim(TxtBCodPers.Text) = "" Or feBitacora.TextMatrix(1, 0) = "" Then
        MsgBox "No existen datos para exportar", vbInformation, "Aviso"
        Exit Sub
    End If

    pgbExcel.Visible = True
    pgbExcel.Min = 0
    pgbExcel.value = 0

    lsArchivo = App.path & "\SPOOLER\ActualizacionDatos_" & TxtBCodPers.Text & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbLibroOpen Then
        Exit Sub
    End If
    nLin = 1
    lsHoja = "Hoja1"
    gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
    
    xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 18
    xlHoja1.Range("A1:B1").ColumnWidth = 12
    xlHoja1.Range("C1:E1").ColumnWidth = 14
    xlHoja1.Range("F1:S1").ColumnWidth = 19
    xlHoja1.Range("F1:S1").ColumnWidth = 19
    xlHoja1.Range("AA1:AA1").ColumnWidth = 12
    xlHoja1.Range("AB1:AB1").ColumnWidth = 18
    
    xlHoja1.Cells(nLin, 1) = "CLIENTE: " & lblPersNombre.Text
    xlHoja1.Cells(nLin, 6) = "DOI: " & lblDOITipo.Text
    xlHoja1.Cells(nLin, 7) = "'" & lblDOINro.Caption
    
    nLin = nLin + 2
    
    xlHoja1.Cells(nLin, 1) = "Doc Primario"
    xlHoja1.Cells(nLin, 2) = "Doc Primario"
    xlHoja1.Cells(nLin, 3) = "Doc Secundario"
    xlHoja1.Cells(nLin, 4) = "Doc Secundario"
    xlHoja1.Cells(nLin, 5) = "Pais Residencia"
    xlHoja1.Cells(nLin, 6) = "Departamento Dom."
    xlHoja1.Cells(nLin, 7) = "Provincia Dom."
    xlHoja1.Cells(nLin, 8) = "Distrito Dom."
    xlHoja1.Cells(nLin, 9) = "Zona Dom."
    xlHoja1.Cells(nLin, 10) = "Direccion Dom."
    xlHoja1.Cells(nLin, 11) = "Referencia Dom."
    xlHoja1.Cells(nLin, 12) = "Departamento Neg."
    xlHoja1.Cells(nLin, 13) = "Provincia Neg."
    xlHoja1.Cells(nLin, 14) = "Distrito Neg."
    xlHoja1.Cells(nLin, 15) = "Zona Neg."
    xlHoja1.Cells(nLin, 16) = "Dirección Neg."
    xlHoja1.Cells(nLin, 17) = "Referencia Neg."
    xlHoja1.Cells(nLin, 18) = "Actividad"
    xlHoja1.Cells(nLin, 19) = "Email"
    xlHoja1.Cells(nLin, 20) = "Tel. Fijo 1"
    xlHoja1.Cells(nLin, 21) = "Tel. Fijo 2"
    xlHoja1.Cells(nLin, 22) = "Celular 1"
    xlHoja1.Cells(nLin, 23) = "Celular 2"
    xlHoja1.Cells(nLin, 24) = "Celular 3"
    xlHoja1.Cells(nLin, 25) = "Fecha Act."
    xlHoja1.Cells(nLin, 26) = "Usuario Reg."
    xlHoja1.Cells(nLin, 27) = "Usuario Resp."
    xlHoja1.Cells(nLin, 28) = "Aut.Remisión Email"
    
    xlHoja1.Range("A" & nLin & ":AB" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":AB" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":AB" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":AB" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":AB" & nLin).Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("A" & nLin & ":AB" & nLin).Interior.Color = RGB(255, 50, 50)
    xlHoja1.Range("A" & nLin & ":AB" & nLin).Font.Color = RGB(255, 255, 255)
    
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
    
    nItem = 1
    nLin = nLin + 1
    pgbExcel.Max = feBitacora.Rows - 1
    For nItem = 1 To feBitacora.Rows - 1
        xlHoja1.Range("A" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignLeft
        xlHoja1.Cells(nLin, 1) = feBitacora.TextMatrix(nItem, 2) 'Doc Primario
        xlHoja1.Cells(nLin, 2) = "'" & feBitacora.TextMatrix(nItem, 3) 'Doc Primario
        xlHoja1.Cells(nLin, 3) = feBitacora.TextMatrix(nItem, 5) 'Doc Secundario
        xlHoja1.Cells(nLin, 4) = "'" & feBitacora.TextMatrix(nItem, 6) 'Doc Secundario
        xlHoja1.Cells(nLin, 5) = feBitacora.TextMatrix(nItem, 8) 'Pais Residencia
        xlHoja1.Cells(nLin, 6) = feBitacora.TextMatrix(nItem, 10) 'Departamento Dom.
        xlHoja1.Cells(nLin, 7) = feBitacora.TextMatrix(nItem, 11) 'Provincia Dom.
        xlHoja1.Cells(nLin, 8) = feBitacora.TextMatrix(nItem, 12) 'Distrito Dom.
        xlHoja1.Cells(nLin, 9) = feBitacora.TextMatrix(nItem, 13) 'Zona Dom.
        xlHoja1.Cells(nLin, 10) = feBitacora.TextMatrix(nItem, 14) 'Direccion Dom.
        xlHoja1.Cells(nLin, 11) = feBitacora.TextMatrix(nItem, 15) 'Referencia Dom.
        xlHoja1.Cells(nLin, 12) = feBitacora.TextMatrix(nItem, 17) 'Departamento Neg.
        xlHoja1.Cells(nLin, 13) = feBitacora.TextMatrix(nItem, 18) 'Provincia Neg.
        xlHoja1.Cells(nLin, 14) = feBitacora.TextMatrix(nItem, 19) 'Distrito Neg.
        xlHoja1.Cells(nLin, 15) = feBitacora.TextMatrix(nItem, 20) 'Zona Neg.
        xlHoja1.Cells(nLin, 16) = feBitacora.TextMatrix(nItem, 21) 'Dirección Neg.
        xlHoja1.Cells(nLin, 17) = feBitacora.TextMatrix(nItem, 22) 'Referencia Neg.
        xlHoja1.Cells(nLin, 18) = feBitacora.TextMatrix(nItem, 23) 'Actividad
        xlHoja1.Cells(nLin, 19) = feBitacora.TextMatrix(nItem, 24) 'Email
        xlHoja1.Cells(nLin, 20) = "'" & feBitacora.TextMatrix(nItem, 25) 'Tel. Fijo 1
        xlHoja1.Cells(nLin, 21) = "'" & feBitacora.TextMatrix(nItem, 26) 'Tel. Fijo 2
        xlHoja1.Cells(nLin, 22) = "'" & feBitacora.TextMatrix(nItem, 27) 'Celular 1
        xlHoja1.Cells(nLin, 23) = "'" & feBitacora.TextMatrix(nItem, 28) 'Celular 2
        xlHoja1.Cells(nLin, 24) = "'" & feBitacora.TextMatrix(nItem, 29) 'Celular 3
        xlHoja1.Cells(nLin, 25) = feBitacora.TextMatrix(nItem, 30) 'Fecha Act.
        xlHoja1.Cells(nLin, 26) = feBitacora.TextMatrix(nItem, 31) 'Usuario Reg.
        xlHoja1.Cells(nLin, 27) = feBitacora.TextMatrix(nItem, 32) 'Usuario Resp.
        xlHoja1.Cells(nLin, 28) = feBitacora.TextMatrix(nItem, 33) 'Aut.Remisión Email

        pgbExcel.value = pgbExcel.value + 1
        nLin = nLin + 1
    Next nItem
    
    gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    gFunGeneral.CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    pgbExcel.Min = 0
    pgbExcel.value = 0
    pgbExcel.Visible = False
End Sub

Private Function LlenaRecordSet_rsPersona() As ADODB.Recordset
Dim rsAux As ADODB.Recordset

Set rsAux = New ADODB.Recordset
With rsAux
    'Crear RecordSet
    .fields.Append "cPaisReside", adVarChar, 12
    .fields.Append "cPersDireccUbiGeo", adVarChar, 12
    .fields.Append "cPersDireccDomicilio", adVarChar, 100
    .fields.Append "cPersRefDomicilio", adVarChar, 100
    .fields.Append "cPersNegocioUbiGeo", adVarChar, 12
    .fields.Append "cPersNegocioDireccion", adVarChar, 100
    .fields.Append "cPersNegocioRef", adVarChar, 200
    .fields.Append "cActiGiro", adVarChar, 100
    .fields.Append "cPersEmail", adVarChar, 50
    .fields.Append "cPersTelefono", adVarChar, 100
    .fields.Append "cPersTelefono2", adVarChar, 100
    .fields.Append "cPersCelular", adVarChar, 100
    .fields.Append "cPersCelular2", adVarChar, 100
    .fields.Append "cPersCelular3", adVarChar, 100
    .fields.Append "nRemisionInfoEmail", adInteger
    .Open
    
    .AddNew
    .fields("cPaisReside") = feBitacora.TextMatrix(feBitacora.row, 7)
    .fields("cPersDireccUbiGeo") = feBitacora.TextMatrix(feBitacora.row, 9)
    .fields("cPersDireccDomicilio") = feBitacora.TextMatrix(feBitacora.row, 14)
    .fields("cPersRefDomicilio") = feBitacora.TextMatrix(feBitacora.row, 15)
    .fields("cPersNegocioUbiGeo") = feBitacora.TextMatrix(feBitacora.row, 16)
    .fields("cPersNegocioDireccion") = feBitacora.TextMatrix(feBitacora.row, 21)
    .fields("cPersNegocioRef") = feBitacora.TextMatrix(feBitacora.row, 22)
    .fields("cActiGiro") = feBitacora.TextMatrix(feBitacora.row, 23)
    .fields("cPersEmail") = feBitacora.TextMatrix(feBitacora.row, 24)
    .fields("cPersTelefono") = feBitacora.TextMatrix(feBitacora.row, 25)
    .fields("cPersTelefono2") = feBitacora.TextMatrix(feBitacora.row, 26)
    .fields("cPersCelular") = feBitacora.TextMatrix(feBitacora.row, 27)
    .fields("cPersCelular2") = feBitacora.TextMatrix(feBitacora.row, 28)
    .fields("cPersCelular3") = feBitacora.TextMatrix(feBitacora.row, 29)
    .fields("nRemisionInfoEmail") = IIf(feBitacora.TextMatrix(feBitacora.row, 33) = "SI", 1, 0)
End With
Set LlenaRecordSet_rsPersona = rsAux
End Function
Private Function LlenaRecordSet_rsDocs() As ADODB.Recordset
Dim rsAux As ADODB.Recordset

Set rsAux = New ADODB.Recordset
With rsAux
    'Crear RecordSet
    .fields.Append "cPersIDTpo", adInteger
    .fields.Append "cPersIDNro", adVarChar, 12
    .Open
    
    If feBitacora.TextMatrix(feBitacora.row, 1) <> 0 Then
        .AddNew
        .fields("cPersIDTpo") = feBitacora.TextMatrix(feBitacora.row, 1)
        .fields("cPersIDNro") = feBitacora.TextMatrix(feBitacora.row, 3)
    End If
    If feBitacora.TextMatrix(feBitacora.row, 4) <> 0 Then
        .AddNew
        .fields("cPersIDTpo") = feBitacora.TextMatrix(feBitacora.row, 4)
        .fields("cPersIDNro") = feBitacora.TextMatrix(feBitacora.row, 6)
    End If
End With

Set LlenaRecordSet_rsDocs = rsAux
End Function
