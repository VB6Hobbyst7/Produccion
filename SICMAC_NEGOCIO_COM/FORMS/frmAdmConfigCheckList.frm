VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdmConfigCheckList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de CheckList por tipo de crédito"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   Icon            =   "frmAdmConfigCheckList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5385
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9499
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Configuración"
      TabPicture(0)   =   "frmAdmConfigCheckList.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   9720
         TabIndex        =   8
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   " Requisitos "
         ForeColor       =   &H00FF0000&
         Height          =   3495
         Left            =   4080
         TabIndex        =   4
         Top             =   1320
         Width           =   6735
         Begin VB.CommandButton cmdReqQuitar 
            Caption         =   "Quitar"
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   3070
            Width           =   1095
         End
         Begin VB.CommandButton cmdReqAgregar 
            Caption         =   "Agregar"
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   3070
            Width           =   1095
         End
         Begin SICMACT.FlexEdit feRequisitos 
            Height          =   2820
            Left            =   120
            TabIndex        =   5
            Top             =   195
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   4974
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Requisitos-IdRequisito"
            EncabezadosAnchos=   "350-6000-0"
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
            ColumnasAEditar =   "X-X-X"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Secciones "
         ForeColor       =   &H00FF0000&
         Height          =   3495
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         Begin VB.CommandButton cmdSecQuitar 
            Caption         =   "Quitar"
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   3070
            Width           =   1095
         End
         Begin VB.CommandButton cmdSecAgregar 
            Caption         =   "Agregar"
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   3070
            Width           =   1095
         End
         Begin SICMACT.FlexEdit feSecciones 
            Height          =   2775
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   4895
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Sección-IdSeccion"
            EncabezadosAnchos=   "350-3500-0"
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
            ColumnasAEditar =   "X-X-X"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Tipo de Crédito a Configurar "
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   6735
         Begin VB.ComboBox cboTipoCred 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   6255
         End
      End
   End
End
Attribute VB_Name = "frmAdmConfigCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmAdmConfigChekList
'** Descripción : Formulario de configuración de lista de requisitos por crédito
'** Creación    : RECO, 20150421 - ERS010-2015
'**********************************************************************************************
Option Explicit
Dim lsUltActulizacion As String


Private Sub cboTipoCred_Click()
    Call CargarSecciones
    Call CargarRequisitos
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdReqAgregar_Click()
    Dim obj As New COMNCredito.NCOMCredito
    Dim sSeccion As String

    sSeccion = InputBox("Ingrese nuevo requisito", "Registro de Requisito")
    If StrPtr(sSeccion) = vbEmpty Then

    ElseIf Trim(sSeccion) <> "" Then
        Call obj.RegistroCredAdmRequisito(feSecciones.TextMatrix(feSecciones.row, 2), sSeccion, lsUltActulizacion, 1)
        Call CargarRequisitos
    Else
        MsgBox "No se puede ingresar un dato vacío", vbInformation, "Alerta"
    End If
End Sub

Private Sub cmdReqQuitar_Click()
    Dim obj As New COMNCredito.NCOMCredito
    If feRequisitos.TextMatrix(1, 1) <> "" Then
        Call obj.CredAdmActualizaRequisito(feRequisitos.TextMatrix(feRequisitos.row, 2), 2, lsUltActulizacion)
    End If
    Call CargarRequisitos
End Sub

Private Sub cmdSecAgregar_Click()
    Dim obj As New COMNCredito.NCOMCredito
    Dim sSeccion As String

    sSeccion = InputBox("Ingrese nueva sección", "Registro de Sección")
    If StrPtr(sSeccion) = vbEmpty Then
    ElseIf Trim(sSeccion) <> "" Then
        Call obj.RegistroCredAdmSeccion(cboTipoCred.ItemData(cboTipoCred.ListIndex), sSeccion, lsUltActulizacion, 1)
        Call CargarSecciones
    Else
        MsgBox "No se puede ingresar un dato vacío", vbInformation, "Alerta"
    End If
End Sub

Private Sub cmdSecQuitar_Click()
    Dim obj As New COMNCredito.NCOMCredito
    If feSecciones.TextMatrix(1, 1) <> "" Then
        Call obj.CredAdmActualizaSeccion(feSecciones.TextMatrix(feSecciones.row, 2), 2, lsUltActulizacion)
    End If
    Call CargarSecciones
    Call CargarRequisitos
End Sub

Private Sub feSecciones_OnRowChange(pnRow As Long, pnCol As Long)
    Call CargarRequisitos
End Sub

Private Sub Form_Load()
    Call CargarCombo
    If CargarSecciones = 1 Then Call CargarRequisitos
    lsUltActulizacion = Format(gdFecSis, "yyyyMMdd") & gsCodUser
End Sub

Private Sub CargarCombo()
    Dim obj As New COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Set rs = obj.ObtieneConstanteFiltroXCodValor(3034, "__0")
    If Not (rs.BOF And rs.EOF) Then
        cboTipoCred.Clear
        For i = 1 To rs.RecordCount
            cboTipoCred.AddItem "" & rs!cConsDescripcion
            cboTipoCred.ItemData(cboTipoCred.NewIndex) = "" & rs!nConsValor
            rs.MoveNext
        Next
        cboTipoCred.ListIndex = 0
    End If
End Sub

Private Function CargarSecciones() As Integer
    Dim obj As New COMNCredito.NCOMCredito
    Dim rs As ADODB.Recordset
    Dim i As Integer

    feSecciones.Clear
    FormateaFlex feSecciones
    Set rs = obj.ListaCredAdmSecciones(cboTipoCred.ItemData(cboTipoCred.ListIndex))
    If Not (rs.EOF And rs.BOF) Then
        For i = 1 To rs.RecordCount
            feSecciones.AdicionaFila
            feSecciones.TextMatrix(i, 2) = rs!nIdSeccion
            feSecciones.TextMatrix(i, 1) = rs!cDescripcion
            rs.MoveNext
        Next
        CargarSecciones = 1
    Else
        CargarSecciones = 0
    End If
End Function
Private Sub CargarRequisitos()
    Dim obj As New COMNCredito.NCOMCredito
    Dim rs As ADODB.Recordset
    Dim i As Integer

    feRequisitos.Clear
    FormateaFlex feRequisitos
    If feSecciones.TextMatrix(feSecciones.row, 2) <> "" Then
    Set rs = obj.ListaCredAdmRequisitos(feSecciones.TextMatrix(feSecciones.row, 2))
        If Not (rs.EOF And rs.BOF) Then
            For i = 1 To rs.RecordCount
                feRequisitos.AdicionaFila
                feRequisitos.TextMatrix(i, 2) = rs!nIdRequisito
                feRequisitos.TextMatrix(i, 1) = rs!cDescripcion
                rs.MoveNext
            Next
        End If
    End If
End Sub
