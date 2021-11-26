VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogBienAjusteVidaUtil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activos Fijos: Ajuste de Vida Útil"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15735
   Icon            =   "frmLogBienAjusteVidaUtil.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAreaAgeNombre 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   885
      Width           =   3060
   End
   Begin VB.TextBox txtTipoBienNombre 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1250
      Width           =   3060
   End
   Begin VB.CommandButton cmdGuardar 
      Cancel          =   -1  'True
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   13200
      TabIndex        =   20
      Top             =   7520
      Width           =   1050
   End
   Begin TabDlg.SSTab TabBien 
      Height          =   7965
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   14049
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Selección de Bienes"
      TabPicture(0)   =   "frmLogBienAjusteVidaUtil.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkTiempoDepr"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExportar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "feBien"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdHistorico"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkPorcDepr"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chkPorcDepr 
         Caption         =   "Porc.Depr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   10920
         TabIndex        =   21
         Top             =   1830
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdHistorico 
         Caption         =   "&Ver Histórico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   1320
         TabIndex        =   7
         Top             =   7395
         Width           =   1290
      End
      Begin Sicmact.FlexEdit feBien 
         Height          =   5190
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   9155
         Cols0           =   12
         HighLight       =   1
         EncabezadosNombres=   "#-nMovNro-Código-Nombre-Marca-Modelo-Serie/caract.-Fecha Adq.-Tmp.Depr.Cont (mes)-Tmp.Depr.Cont (mes)2-Porc. Depr-Motivo"
         EncabezadosAnchos=   "350-0-1500-1700-1300-1300-1500-1100-2000-0-1200-3000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-8-X-10-11"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-L-L-C-L-C-C-C-C-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-3-3-2-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filtro de Busqueda"
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
         Height          =   1545
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   8355
         Begin VB.TextBox txtSerieNombre 
            Height          =   285
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1000
            Width           =   3060
         End
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "&Mostrar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6240
            TabIndex        =   3
            Top             =   960
            Width           =   1050
         End
         Begin VB.ComboBox cboMostrar 
            Height          =   315
            ItemData        =   "frmLogBienAjusteVidaUtil.frx":0326
            Left            =   6960
            List            =   "frmLogBienAjusteVidaUtil.frx":0328
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin Sicmact.TxtBuscar txtTipoBienCod 
            Height          =   255
            Left            =   1560
            TabIndex        =   1
            Top             =   645
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   450
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Sicmact.TxtBuscar txtAreaAgeCod 
            Height          =   255
            Left            =   1560
            TabIndex        =   0
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   450
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Sicmact.TxtBuscar txtSerieCod 
            Height          =   255
            Left            =   1560
            TabIndex        =   2
            Top             =   1005
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   450
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label24 
            Caption         =   "Serie:"
            Height          =   255
            Left            =   1080
            TabIndex        =   23
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Agencia:"
            Height          =   255
            Left            =   840
            TabIndex        =   19
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Mostrar:"
            Height          =   255
            Left            =   6240
            TabIndex        =   18
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Activo Fijo:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   660
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   120
         TabIndex        =   6
         Top             =   7395
         Width           =   1050
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   14280
         TabIndex        =   8
         Top             =   7395
         Width           =   1050
      End
      Begin MSComctlLib.ListView lstAhorros 
         Height          =   2790
         Left            =   -74910
         TabIndex        =   10
         Top             =   495
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   4921
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Producto"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agencia"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Nro. Cuenta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nro. Cta Antigua"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Participación"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "SaldoCont"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "SaldoDisp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Motivo de Bloque"
            Object.Width           =   7231
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Moneda"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox chkTiempoDepr 
         Caption         =   "Tmp. Depr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   22
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblDolaresAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -67680
         TabIndex        =   15
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label lblSolesAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -70815
         TabIndex        =   14
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   13
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73185
         TabIndex        =   12
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   11
         Top             =   3465
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmLogBienAjusteVidaUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmLogBienMnt
'** Descripción : Ajuste de Vida Útil Bienes creado segun ERS059-2013
'** Creación : EJVG, 20130621 09:00:00 AM
'***************************************************************************
Option Explicit
Dim oBien As DBien
Dim fdEjercicioAnioActIni As Date
Dim nEnterCell As Integer 'NAGL 20191023
Dim sCambioColumn As String 'NAGL 20191023
Dim checkTmpDepPorcDep As String 'NAGL 20191023
Dim checKActivado As String 'NAGL 20191023

Private Sub Form_Load()
    Set oBien = New DBien
    '*********************
    nEnterCell = 0
    sCambioColumn = ""
    checkTmpDepPorcDep = ""
    '*****NAGL 20191024***
    CentraForm Me
    CargarControles
    feBien.Enabled = False
    feBien.lbEditarFlex = False 'NAGL 20191024
    
    fdEjercicioAnioActIni = CDate(CStr(Year(gdFecSis)) & "-01-01")
    'Se podra editar si y solo si no realizaron depreciación en enero
    'If oBien.RealizaronDepreciacion(Year(fdEjercicioAnioActIni), Month(fdEjercicioAnioActIni)) Then
        'feBien.lbEditarFlex = False
    'Else
        'feBien.lbEditarFlex = True
    'End If 'Comentado by NAGL 20191024
    
    If oBien.HabilitaAjusteVidaUtil = True Then
        feBien.lbEditarFlex = True
    Else
        feBien.lbEditarFlex = False
    End If 'NAGL 20191223 Según RFC1910190001
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oBien = Nothing
End Sub

Private Sub CargarControles()
    Dim obj As New DActualizaDatosArea
    'Me.txtAreaAgeCod.rs = obj.GetAgenciasAreas 'Comentado by NAGL 20191019
    Me.txtAreaAgeCod.rs = obj.GetAgencias 'Agregado by NAGL 20191019 Según RFC1910190001
    cboMostrar.Clear
    cboMostrar.AddItem "Todos" & Space(200) & "1"
    cboMostrar.AddItem "Pendientes" & Space(200) & "2"
    txtTipoBienCod.rs = oBien.RecuperaCategoriasBienPaObjeto(True, "")
    txtSerieCod.rs = oBien.RecuperaSeriesPaObjeto("", "")
    Set obj = Nothing
End Sub

Private Sub cmdMostrar_Click()
    Dim rs As New ADODB.Recordset
    Dim fila As Long
    Dim lsAreaAgeCod As String
    
    If txtTipoBienCod.Text = "" And txtSerieCod.Text = "" Then
        MsgBox "Ud. debe seleccionar un Tipo de Activo Fijo", vbInformation, "Aviso"
        txtTipoBienCod.SetFocus
        Exit Sub
    ElseIf txtSerieCod.Text = "" Then
        If cboMostrar.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar la opción de muestra", vbInformation, "Aviso"
            cboMostrar.SetFocus
            Exit Sub
        End If
    Else
        cboMostrar.ListIndex = 0
    End If 'NAGL 20191222 Agregó Condicional
    
    If txtAreaAgeCod.Text <> "" Then
        'lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2)) 'Comentado by NAGL 20191019
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) 'NAGL 20191019 Según RFC1910190001
    End If
    
    'Set rs = oBien.RecuperaBienxAjusteVidaUtil(txtTipoBienCod.Text, CInt(Trim(Right(cboMostrar.Text, 2))), lsAreaAgeCod, gdFecSis) 'NAGL 20191023 Agregó gdFecSis
    'LimpiaFlex feBien
    
    'Do While Not rs.EOF
        'feBien.AdicionaFila
        'Fila = feBien.row
        'feBien.TextMatrix(Fila, 1) = rs!nMovNro
        'feBien.TextMatrix(Fila, 2) = rs!cInventarioCod
        'feBien.TextMatrix(Fila, 3) = rs!cNombre
        'feBien.TextMatrix(Fila, 4) = rs!cMarca
        'feBien.TextMatrix(Fila, 5) = rs!cModelo
        'feBien.TextMatrix(Fila, 6) = rs!cSerie
        'feBien.TextMatrix(Fila, 7) = Format(rs!dCompra, "dd/mm/yyyy") 'NAGL Cambió de dActivacion a dCompra
        'feBien.TextMatrix(Fila, 8) = rs!nBSPerDeprecia
        'feBien.col = 8
        'feBien.CellBackColor = vbGreen
        '****************************
        'feBien.TextMatrix(Fila, 9) = rs!nBSPerDeprecia 'Para tener el valor anterior
        'feBien.TextMatrix(Fila, 10) = Format(IIf(rs!nBSPerDeprecia = 0, 0, (12 / rs!nBSPerDeprecia) * 100), "###,##0.00")
        'rs.MoveNext
    'Loop 'Comentado by NAGL 20191015 para trasladarlo en el Sgte Función: MuestraBienesVidaUtil
    
    If MuestraBienesVidaUtil(txtTipoBienCod.Text, CInt(Trim(Right(cboMostrar.Text, 2))), lsAreaAgeCod, gdFecSis, "Show", txtSerieCod.Text) = True Then
        '***NAGL 20191023
        checKActivado = "1" '***NAGL 20191023
        feBien.Enabled = True
        chkTiempoDepr.Visible = True
        chkPorcDepr.Visible = True
        chkTiempoDepr.value = "1"
        chkPorcDepr.value = "0"
        checKActivado = "0" '***NAGL 20191023
        '*****
        feBien.TopRow = 1
        feBien.row = 1
        SendKeys "{Tab}", True
    End If 'NAGL 20191024
  
    If FlexVacio(feBien) Then
        MsgBox "No se encontraron resultados de la Búsqueda realizada", vbInformation, "Aviso"
        feBien.lbEditarFlex = False
    Else
        'fdEjercicioAnioActIni = CDate(CStr(Year(gdFecSis)) & "-01-01")
        'If oBien.RealizaronDepreciacion(Year(fdEjercicioAnioActIni), Month(fdEjercicioAnioActIni)) Then
            'feBien.lbEditarFlex = False
        'Else
            'feBien.lbEditarFlex = True
        'End If 'Comentado by NAGL 20191223
        If oBien.HabilitaAjusteVidaUtil = True Then
        feBien.lbEditarFlex = True
        Else
            feBien.lbEditarFlex = False
        End If 'NAGL 20191223 Según RFC1910190001
    End If
    Set rs = Nothing
End Sub

Private Function MuestraBienesVidaUtil(psBienCod As String, pnTpoSel As Integer, psAgencia As String, pdFecha As Date, Optional psSombrear As String = "", Optional psCodSerie As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim fila As Long
    Dim lsAreaAgeCod As String
    Set rs = oBien.RecuperaBienxAjusteVidaUtil(psBienCod, pnTpoSel, psAgencia, pdFecha, psCodSerie)
    LimpiaFlex feBien
    If rs.RecordCount > 0 Then
       txtAreaAgeCod.Enabled = False
       txtAreaAgeNombre.Enabled = False
       txtTipoBienCod.Enabled = False
       txtTipoBienNombre.Enabled = False
       txtSerieCod.Enabled = False
       txtSerieNombre.Enabled = False
       cboMostrar.Enabled = False
       cmdMostrar.Enabled = False
       cmdExportar.Enabled = False
       cmdHistorico.Enabled = False
       cmdGuardar.Enabled = False
       cmdSalir.Enabled = False
    End If
    If rs.RecordCount <= 0 Then
        chkTiempoDepr.Visible = False
        chkPorcDepr.Visible = False
        FormateaFlex feBien
        MuestraBienesVidaUtil = False
        Exit Function
    End If
    Do While Not rs.EOF
        feBien.AdicionaFila
        fila = feBien.row
        feBien.TextMatrix(fila, 1) = rs!nMovNro
        feBien.TextMatrix(fila, 2) = rs!cInventarioCod
        feBien.TextMatrix(fila, 3) = rs!cNombre
        feBien.TextMatrix(fila, 4) = rs!cMarca
        feBien.TextMatrix(fila, 5) = rs!cModelo
        feBien.TextMatrix(fila, 6) = rs!cSerie
        feBien.TextMatrix(fila, 7) = Format(rs!dActivacion, "dd/mm/yyyy")
        feBien.TextMatrix(fila, 8) = rs!nBSPerDeprecia
        If psSombrear = "Show" Then
            feBien.col = 8
            feBien.CellBackColor = vbGreen
        End If
        '****************************
        feBien.TextMatrix(fila, 9) = rs!nBSPerDeprecia 'Para tener el valor anterior
        If rs!nBSPerDeprecia = 0 Then
            feBien.TextMatrix(fila, 10) = 0
        Else
            feBien.TextMatrix(fila, 10) = Format((12 / rs!nBSPerDeprecia) * 100, "###,##0.00")
        End If
        feBien.TextMatrix(fila, 11) = rs!cMotivo
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       txtAreaAgeCod.Enabled = True
       txtAreaAgeNombre.Enabled = True
       txtTipoBienCod.Enabled = True
       txtTipoBienNombre.Enabled = True
       txtSerieCod.Enabled = True
       txtSerieNombre.Enabled = True
       cboMostrar.Enabled = True
       cmdMostrar.Enabled = True
       cmdExportar.Enabled = True
       cmdHistorico.Enabled = True
       cmdGuardar.Enabled = True
       cmdSalir.Enabled = True
    End If
    MuestraBienesVidaUtil = True
End Function 'NAGL 20191023 Según RFC1910190001

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub feBien_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    On Error GoTo ErrValidate
    sColumnas = Split(feBien.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}"
        Exit Sub
    Else
        '*****************************
        If Len(feBien.TextMatrix(pnRow, pnCol)) > 5 And pnCol <> 11 Then
           Cancel = False
           MsgBox "El valor ingresado es incorrecto", vbInformation, "Aviso"
           SendKeys "{Tab}"
           Exit Sub
        Else
            If pnRow < feBien.Rows - 1 Then
                feBien.AvanceCeldas = Vertical
            Else
                sCambioColumn = "Col"
            End If
        End If 'Agregado by NAGL 20191023
    End If
    'If pnCol = 8 Then
        ''Graba el ajuste de vida util
        'If Not GrabaAjusteVidaUtil(pnRow) Then
            'Cancel = False
            'SendKeys "{Tab}", True
            'Exit Sub
        'End If
    'End If
    'Comentado by NAGL 20191023
    Exit Sub
ErrValidate:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Function ValidaTiempoDepreciacion() As Boolean
Dim fila As Integer
Dim lsNombreInventario As String
Dim lnBSDepreciaCont As Integer, lnBSDepreciaContAnt As Integer
Dim lnMesesTranscFecha As Integer
Dim ldFechaActv As Date
Dim ldEjercicioAnioAntFin As Date
Dim lnMesesTranscurridos As Integer
Dim psMotivo As String

For fila = 1 To feBien.Rows - 1
    lsNombreInventario = CStr(Trim(feBien.TextMatrix(fila, 2)))
    lnBSDepreciaCont = CInt(Trim(feBien.TextMatrix(fila, 8))) 'Nuevo
    lnBSDepreciaContAnt = CInt(Trim(feBien.TextMatrix(fila, 9))) 'Anterior
    ldFechaActv = CDate(Trim(feBien.TextMatrix(fila, 7)))
    ldEjercicioAnioAntFin = DateAdd("D", -1, fdEjercicioAnioActIni) 'Diciembre año anterior
    lnMesesTranscurridos = DateDiff("M", ldFechaActv, IIf(Year(ldEjercicioAnioAntFin) < Year(ldFechaActv), gdFecSis, ldEjercicioAnioAntFin))
    psMotivo = CStr(Trim(feBien.TextMatrix(fila, 11)))
    If lnBSDepreciaContAnt <> lnBSDepreciaCont Then
        'Tiempo debe ser mayor que cero
        If lnBSDepreciaCont <= 0 Then
            MsgBox "El tiempo de depreciación debe ser mayor que cero en el Bien " & lsNombreInventario, vbInformation, "Aviso"
            Exit Function
        End If
        
        'Tiempo que edita debe ser mayor al tiempo transcurrido desde la fecha de activación hasta el ejercicio anterior
        If lnBSDepreciaCont < lnMesesTranscurridos Then
            MsgBox "El tiempo de depreciación que se esta ingresando es menor al tiempo transcurrido en el Bien " & lsNombreInventario, vbInformation, "Aviso"
            Exit Function
        End If
        If psMotivo = "" Then
            MsgBox "Por favor ingresar el motivo del Cambio en la Depreciación del Bien " & lsNombreInventario, vbInformation, "Aviso"
            Exit Function
        End If
    End If
Next fila
ValidaTiempoDepreciacion = True
End Function 'NAGL 20191023 Según RFC1910190001

Private Sub cmdGuardar_Click()
Dim bien As New DBien
Dim mov As New DMov
Dim bTransBien As Boolean, bTransMov As Boolean
Dim pnFilas As Integer
'**********NAGL 20191223********
Dim lnBSDepreciaCont As Integer, lnBSDepreciaContAnt As Integer
Dim lsMovNro As String, lsInventarioCod As String
Dim lnMovNro As Long, lnMovNroAF As Long
Dim ValDepr As Boolean
Dim psMotivo As String
'*******************************
On Error GoTo ErrGrabaAjusteVidaUtil
ValDepr = False
psMotivo = ""

If FlexVacio(feBien) Then
        MsgBox "No hay información disponible para guardar", vbInformation, "Aviso"
        Exit Sub
End If 'NAGL 20191222 Según RFC1910190001

If MsgBox("Se va a realizar el Ajuste de la Vida Util del Bien" & Chr(10) & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
If ValidaTiempoDepreciacion = True Then
    ValDepr = True
    cmdGuardar.Width = 1150
    cmdGuardar.Caption = "Procesando"
    cmdGuardar.Enabled = False
    cmdExportar.Enabled = False
    cmdHistorico.Enabled = False
    cmdSalir.Enabled = False
    chkTiempoDepr.Enabled = False
    chkPorcDepr.Enabled = False
    Frame1.Enabled = False
    
    For pnFilas = 1 To feBien.Rows - 1
        lnMovNroAF = CLng(Trim(feBien.TextMatrix(pnFilas, 1))) 'Movimiento del Registro del Bien
        lsInventarioCod = CStr(Trim(feBien.TextMatrix(pnFilas, 2)))
        lnBSDepreciaCont = CInt(Trim(feBien.TextMatrix(pnFilas, 8))) 'Nuevo
        lnBSDepreciaContAnt = CInt(Trim(feBien.TextMatrix(pnFilas, 9))) 'Anterior
        psMotivo = CStr(Trim(feBien.TextMatrix(pnFilas, 11))) 'Descripción por Cambio en el Tiempo de Depreciación
        If lnBSDepreciaCont <> lnBSDepreciaContAnt Then
            bien.dBeginTrans
            mov.BeginTrans
            bTransBien = True
            bTransMov = True
            
            lsMovNro = mov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            mov.InsertaMov lsMovNro, gnAjusteVidaUtilAF, "Ajuste de Vida Útil Bien: " & lsInventarioCod, gMovEstContabMovContable, gMovFlagVigente
            lnMovNro = mov.GetnMovNro(lsMovNro)
            bien.InsertaVidaUtilAF lnMovNroAF, lnMovNro, lnBSDepreciaCont, psMotivo 'NAGL Agregó psMotivo 20191223
            bien.ActualizarAF lnMovNroAF, , , , , lnBSDepreciaCont
            
            feBien.TextMatrix(pnFilas, 9) = CInt(Trim(feBien.TextMatrix(pnFilas, 8)))
            
            mov.CommitTrans
            bien.dCommitTrans
            bTransBien = False
            bTransMov = False
        Else
            bien.ActualizaVidaUtilAF_NewGlosa lnMovNroAF, gdFecSis, lnBSDepreciaCont, psMotivo 'NAGL 20191223
        End If
    Next pnFilas
    
     MsgBox "Se ha realizado el Ajuste de la Vida Util del Bien satisfactoriamente", vbInformation, "Aviso"
     cmdGuardar.Width = 1050
     cmdGuardar.Caption = "Guardar"
     cmdGuardar.Enabled = True
     cmdExportar.Enabled = True
     cmdHistorico.Enabled = True
     cmdSalir.Enabled = True
     chkTiempoDepr.Enabled = True
     chkPorcDepr.Enabled = True
     Frame1.Enabled = True
     Set mov = Nothing
     Set bien = Nothing
     Exit Sub
End If
ErrGrabaAjusteVidaUtil:
    If ValDepr = True Then
        MsgBox Err.Description, vbCritical, "Aviso"
        If bTransBien Then
            bien.dRollbackTrans
            Set bien = Nothing
        End If
        If bTransMov Then
            mov.RollbackTrans
            Set mov = Nothing
        End If
        cmdGuardar.Width = 1050
        cmdGuardar.Caption = "Guardar"
        cmdGuardar.Enabled = True
        cmdExportar.Enabled = True
        cmdHistorico.Enabled = True
        cmdSalir.Enabled = True
        chkTiempoDepr.Enabled = True
        chkPorcDepr.Enabled = True
        Frame1.Enabled = True
    End If
End Sub 'NAGL 20191023 Según RFC1910190001

Private Sub chkTiempoDepr_Click()
Dim I As Integer
    If chkTiempoDepr = 1 Then
        chkPorcDepr.value = 0
        If checKActivado = "0" Then
            If MuestraBienesVidaUtil(txtTipoBienCod.Text, CInt(Trim(Right(cboMostrar.Text, 2))), Left(txtAreaAgeCod.Text, 3), gdFecSis, "", txtSerieCod.Text) = True Then
                checkTmpDepPorcDep = "Temp"
            End If
        End If
        checkTmpDepPorcDep = "Temp"
        For I = 1 To feBien.Rows - 1
            feBien.row = I
            feBien.col = 8
            feBien.CellBackColor = vbGreen
            feBien.ColumnasAEditar = "X-X-X-X-X-X-X-X-8-X-X-11"
            feBien.row = I
            feBien.col = 10
            feBien.CellBackColor = vbWhite
        Next I
    ElseIf chkTiempoDepr = 0 And chkPorcDepr.value = 0 Then
       chkTiempoDepr.value = 1
    End If
End Sub 'NAGL 20191023 Según RFC1910190001

Private Sub chkPorcDepr_Click()
Dim I As Integer
    If chkPorcDepr = 1 Then
        chkTiempoDepr.value = 0
        If checKActivado = "0" Then
            If MuestraBienesVidaUtil(txtTipoBienCod.Text, CInt(Trim(Right(cboMostrar.Text, 2))), Left(txtAreaAgeCod.Text, 3), gdFecSis, "", txtSerieCod.Text) = True Then
                checkTmpDepPorcDep = "Porc"
            End If
        End If
        checkTmpDepPorcDep = "Porc"
        For I = 1 To feBien.Rows - 1
            feBien.row = I
            feBien.col = 10
            feBien.CellBackColor = vbGreen
            feBien.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-10-11"
            feBien.row = I
            feBien.col = 8
            feBien.CellBackColor = vbWhite
        Next I
    ElseIf chkPorcDepr = 0 And chkTiempoDepr.value = 0 Then
       chkPorcDepr.value = 1
    End If
End Sub 'NAGL 20191023 Según RFC1910190001

Private Sub feBien_RowColChange()
If nEnterCell = 0 Then
    If feBien.col = 8 Or feBien.col = 10 Then
        feBien.AvanceCeldas = Vertical
    End If
    If sCambioColumn = "" Then
        CalculaPorcentajeDeprec feBien.col, feBien.row - 1
    Else
       If sCambioColumn = "Col" And feBien.col = 10 And checkTmpDepPorcDep = "Temp" Then
            CalculaPorcentajeDeprec feBien.Cols - 4, feBien.Rows - 1
       ElseIf sCambioColumn = "Col" And feBien.col = 11 And checkTmpDepPorcDep = "Porc" Then
            CalculaPorcentajeDeprec feBien.Cols - 2, feBien.Rows - 1
       End If
       sCambioColumn = ""
    End If
End If
nEnterCell = 0
End Sub 'NAGL 20191023 Según RFC1910190001

Private Sub feBien_EnterCell()
    nEnterCell = 1
End Sub 'NAGL 20191023 Según RFC1910190001

Private Sub CalculaPorcentajeDeprec(pnCol As Integer, pnRow As Integer)
Dim nPorcDepr As Double
Dim nTpoMesDep As Double
nPorcDepr = 0
    If pnCol = 8 And checkTmpDepPorcDep = "Temp" Then
        If CDbl(feBien.TextMatrix(pnRow, 8)) = 0 Then
            nPorcDepr = 0
        Else
            nPorcDepr = (12 / CDbl(feBien.TextMatrix(pnRow, 8))) * 100
        End If
        feBien.TextMatrix(pnRow, 10) = IIf(nPorcDepr = 0, 0, Format(Round(nPorcDepr, 2), "###,##0.00"))
        If pnRow <> 0 Then
            If CDbl(feBien.TextMatrix(pnRow, 8)) <> CDbl(feBien.TextMatrix(pnRow, 9)) Then
               feBien.TextMatrix(pnRow, 11) = ""
            End If
        End If
    ElseIf pnCol = 10 And checkTmpDepPorcDep = "Porc" Then
        If CDbl(feBien.TextMatrix(pnRow, 10)) = 0 Then
            nTpoMesDep = 0
        Else
            nTpoMesDep = Round((1 / (CDbl(feBien.TextMatrix(pnRow, 10)) / 100)) * 12, 0)
        End If
        feBien.TextMatrix(pnRow, 8) = IIf(nTpoMesDep = 0, 0, Format(Round(nTpoMesDep, 2), "###,##"))
        If pnRow <> 0 Then
            If CDbl(feBien.TextMatrix(pnRow, 8)) <> CDbl(feBien.TextMatrix(pnRow, 9)) Then
               feBien.TextMatrix(pnRow, 11) = ""
            End If
        End If
    End If
End Sub 'NAGL 20191023 Según RFC1910190001

Private Sub cmdHistorico_Click()
    If FlexVacio(feBien) Then
        MsgBox "Ud. primero debe de seleccionar el Activo Fijo", vbInformation, "Aviso"
        feBien.SetFocus
        Exit Sub
    End If
    frmLogBienHistoVidaUtil.Inicio CLng(feBien.TextMatrix(feBien.row, 1))
End Sub

Private Sub txtTipoBienCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cboMostrar.SetFocus
        txtSerieCod.SetFocus 'NAGL 20191222
    End If
End Sub
Private Sub cboMostrar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAreaAgeCod.SetFocus
    End If
End Sub
Private Sub txtAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cmdMostrar.SetFocus
        txtTipoBienCod.SetFocus 'NAGL 20191222
    End If
End Sub
Private Sub cmdExportar_Click()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim lnFila As Long, lnColumna As Long, lnColumnaMax As Long
    Dim I As Long, j As Long
    Dim lsArchivo As String
    
On Error GoTo ErrExportar
    
    If FlexVacio(feBien) Then
        MsgBox "No hay información para exportar a formato Excel", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    lsArchivo = "\spooler\RptAjusteVidaUtilBien" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    Set xlsLibro = xlsAplicacion.Workbooks.Add

    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "Reporte Ajuste Vida Util"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    
    lnFila = 2
    
    For I = 0 To feBien.Rows - 1
        lnColumna = 2
        For j = 0 To feBien.Cols - 1
            If feBien.ColWidth(j) > 0 Then
                xlsHoja.Cells(lnFila, lnColumna) = "'" & feBien.TextMatrix(I, j)
                lnColumna = lnColumna + 1
                lnColumnaMax = lnColumna
            End If
        Next
        lnFila = lnFila + 1
    Next

    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Interior.Color = RGB(191, 191, 191)
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).Borders.Weight = xlThin

    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).EntireColumn.AutoFit
    
    MsgBox "Se ha exportado satisfactoriamente la información", vbInformation, "Aviso"
    
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Screen.MousePointer = 0
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    Exit Sub
ErrExportar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub txtAreaAgeCod_EmiteDatos()
   txtAreaAgeNombre.Text = ""
   If txtAreaAgeCod.Text <> "" Then
        txtAreaAgeNombre.Text = txtAreaAgeCod.psDescripcion
        txtTipoBienCod.Text = ""
        txtTipoBienNombre.Text = ""
        txtSerieCod.Text = ""
        txtSerieNombre.Text = ""
        Call LimpiaFlex(feBien)
        feBien.Enabled = False
        chkTiempoDepr.Visible = False
        chkPorcDepr.Visible = False
   End If
End Sub 'NAGL 20191222 Según RFC1910190001
Private Sub txtTipoBienCod_EmiteDatos()
    Dim oBien As New DBien
    Dim lsAgeCod As String
    Dim lsCategoria As String
    txtTipoBienNombre.Text = ""
    If txtTipoBienCod.Text <> "" Then
       lsCategoria = txtTipoBienCod.Text
       txtTipoBienNombre.Text = txtTipoBienCod.psDescripcion
    End If
    If txtAreaAgeCod.Text <> "" Then
        lsAgeCod = Left(txtAreaAgeCod.Text, 3)
    End If
    txtSerieCod.Text = ""
    txtSerieNombre.Text = ""
    Call LimpiaFlex(feBien)
    feBien.Enabled = False
    chkTiempoDepr.Visible = False
    chkPorcDepr.Visible = False
    txtSerieCod.rs = oBien.RecuperaSeriesPaObjeto("", lsCategoria, , lsAgeCod)
    txtSerieCod_EmiteDatos
    Set oBien = Nothing
End Sub 'NAGL 20191222 Según RFC1910190001

Private Sub txtSerieCod_EmiteDatos()
   txtSerieNombre.Text = ""
   If txtSerieCod.Text <> "" Then
        txtSerieNombre.Text = txtSerieCod.psDescripcion
   End If
End Sub 'NAGL 20191222 Según RFC1910190001




'Private Function GrabaAjusteVidaUtil(ByVal fila As Long) As Boolean
'    Dim bien As New DBien
'    Dim mov As New DMov
'    Dim bTransBien As Boolean, bTransMov As Boolean
'    Dim lsMovNro As String, lsInventarioCod As String
'    Dim lnMovNro As Long, lnMovNroAF As Long
'    Dim lnBSDepreciaCont As Integer, lnBSDepreciaContAnt As Integer
'    Dim lnMesesTranscFecha As Integer
'    Dim ldFechaActv As Date
'    Dim ldEjercicioAnioAntFin As Date
'    Dim lnMesesTranscurridos As Integer
'
'    On Error GoTo ErrGrabaAjusteVidaUtil
'
'    lnMovNroAF = CLng(Trim(feBien.TextMatrix(fila, 1)))
'    lsInventarioCod = CStr(Trim(feBien.TextMatrix(fila, 2)))
'    lnBSDepreciaCont = CInt(Trim(feBien.TextMatrix(fila, 8))) 'Nuevo
'    lnBSDepreciaContAnt = CInt(Trim(feBien.TextMatrix(fila, 9))) 'Anterior
'    ldFechaActv = CDate(Trim(feBien.TextMatrix(fila, 7)))
'    ldEjercicioAnioAntFin = DateAdd("D", -1, fdEjercicioAnioActIni) 'Diciembre año anterior
'    lnMesesTranscurridos = DateDiff("M", ldFechaActv, IIf(Year(ldEjercicioAnioAntFin) < Year(ldFechaActv), gdFecSis, ldEjercicioAnioAntFin))
'
'    'No realizó edición, continuará en el recorrido pero no grabará cambios
'    If lnBSDepreciaContAnt = lnBSDepreciaCont Then
'        GrabaAjusteVidaUtil = True
'        Exit Function
'    End If
'    'Tiempo debe ser mayor que cero
'    If lnBSDepreciaCont <= 0 Then
'        MsgBox "El tiempo de total de depreciación debe ser mayor que cero", vbInformation, "Aviso"
'        Exit Function
'    End If
'    'Tiempo que edita debe ser mayor al tiempo transcurrido desde la fecha de activación hasta el ejercicio anterior
'    If lnBSDepreciaCont < lnMesesTranscurridos Then
'        MsgBox "El tiempo de total de depreciación que se esta ingresando es menor al tiempo transcurrido", vbInformation, "Aviso"
'        Exit Function
'    End If
'
'    If MsgBox("Se va a realizar el Ajuste de la Vida Util del Bien" & Chr(10) & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Function
'
'    bien.dBeginTrans
'    mov.BeginTrans
'    bTransBien = True
'    bTransMov = True
'
'    lsMovNro = mov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
'    mov.InsertaMov lsMovNro, gnAjusteVidaUtilAF, "Ajuste de Vida Útil Bien: " & lsInventarioCod, gMovEstContabMovContable, gMovFlagVigente
'    lnMovNro = mov.GetnMovNro(lsMovNro)
'    bien.InsertaVidaUtilAF lnMovNroAF, lnMovNro, lnBSDepreciaCont
'    bien.ActualizarAF lnMovNroAF, , , , , lnBSDepreciaCont
'
'    feBien.TextMatrix(fila, 9) = CInt(Trim(feBien.TextMatrix(fila, 8)))
'
'    mov.CommitTrans
'    bien.dCommitTrans
'    bTransBien = False
'    bTransMov = False
'
'    GrabaAjusteVidaUtil = True
'    MsgBox "Se ha realizado el Ajuste de la Vida Util del Bien satisfactoriamente", vbInformation, "Aviso"
'
'    Set mov = Nothing
'    Set bien = Nothing
'    Exit Function
'ErrGrabaAjusteVidaUtil:
'    GrabaAjusteVidaUtil = False
'    MsgBox Err.Description, vbCritical, "Aviso"
'    If bTransBien Then
'        bien.dRollbackTrans
'        Set bien = Nothing
'    End If
'    If bTransMov Then
'        mov.RollbackTrans
'        Set mov = Nothing
'    End If
'End Function
'Comentado by NAGL 20191222
