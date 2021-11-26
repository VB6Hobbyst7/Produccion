VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredVerDocsPendiente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos pendientes"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   Icon            =   "frmCredVerDocsPendiente.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      Height          =   345
      Left            =   8160
      TabIndex        =   4
      Top             =   5600
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9210
      TabIndex        =   3
      Top             =   5600
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Por vencerse"
      TabPicture(0)   =   "frmCredVerDocsPendiente.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TabDocumento"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Vencidos"
      TabPicture(1)   =   "frmCredVerDocsPendiente.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SSTab2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TabDlg.SSTab TabDocumento 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   2
         Top             =   840
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   7858
         _Version        =   393216
         Style           =   1
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Garantías no constituidas"
         TabPicture(0)   =   "frmCredVerDocsPendiente.frx":0342
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "feGarantiaXVenc"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Pólizas"
         TabPicture(1)   =   "frmCredVerDocsPendiente.frx":035E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fePolizaXVenc"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tasaciones"
         TabPicture(2)   =   "frmCredVerDocsPendiente.frx":037A
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "feTasacXVenc"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin NegForms.FlexEdit fePolizaXVenc 
            Height          =   3810
            Left            =   -74880
            TabIndex        =   5
            Top             =   480
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6720
            Cols0           =   9
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Agencia-N° Garantía-Ult. Crédito-Titular Garantía-Ult. Analista-N° Poliza-Fecha Vencimiento-Días"
            EncabezadosAnchos=   "0-1200-1200-1200-1800-1800-1200-1700-600"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-C-L-L-L-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin NegForms.FlexEdit feTasacXVenc 
            Height          =   3810
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6720
            Cols0           =   8
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Agencia-N° Garantía-Ult. Crédito-Titular Garantía-Ult. Analista-Fecha Vencimiento-Días"
            EncabezadosAnchos=   "0-1200-1200-1200-1800-1800-1700-600"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-C-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin NegForms.FlexEdit feGarantiaXVenc 
            Height          =   3810
            Left            =   -74880
            TabIndex        =   14
            Top             =   480
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6720
            Cols0           =   8
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Agencia-N° Garantía-Ult. Crédito-Titular Garantía-Ult. Analista-Fecha Vencimiento-Días"
            EncabezadosAnchos=   "0-1200-1200-1200-1800-1800-1700-600"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-C-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4455
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   7858
         _Version        =   393216
         Style           =   1
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
         TabCaption(0)   =   "Garantías no constituidas"
         TabPicture(0)   =   "frmCredVerDocsPendiente.frx":0396
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "feGarantiaVenc"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Pólizas"
         TabPicture(1)   =   "frmCredVerDocsPendiente.frx":03B2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fePolizaVenc"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tasaciones"
         TabPicture(2)   =   "frmCredVerDocsPendiente.frx":03CE
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "feTasacVenc"
         Tab(2).ControlCount=   1
         Begin NegForms.FlexEdit feGarantiaVenc 
            Height          =   3810
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6720
            Cols0           =   8
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Agencia-N° Garantía-Ult. Crédito-Titular Garantía-Ult. Analista-Fecha Vencimiento-Días"
            EncabezadosAnchos=   "0-1200-1200-1200-1800-1800-1700-600"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-C-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin NegForms.FlexEdit FlexEdit2 
            Height          =   3810
            Left            =   -74880
            TabIndex        =   9
            Top             =   480
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6720
            Cols0           =   8
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Agencia-N° Garantía-Ult. Crédito-Titular Garantía-Ult. Analista-Fecha Vencimiento-Días"
            EncabezadosAnchos=   "0-1200-1200-1200-1800-1800-1700-600"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-L-L-L-L-L-L"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin NegForms.FlexEdit FlexEdit3 
            Height          =   3810
            Left            =   -74880
            TabIndex        =   10
            Top             =   480
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6720
            Cols0           =   8
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Agencia-N° Garantía-Ult. Crédito-Titular Garantía-Ult. Analista-Fecha Vencimiento-Días"
            EncabezadosAnchos=   "0-1200-1200-1200-1800-1800-1700-600"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-L-L-L-L-L-L"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin NegForms.FlexEdit fePolizaVenc 
            Height          =   3810
            Left            =   -74880
            TabIndex        =   12
            Top             =   480
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6720
            Cols0           =   9
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Agencia-N° Garantía-Ult. Crédito-Titular Garantía-Ult. Analista-N° Poliza-Fecha Vencimiento-Días"
            EncabezadosAnchos=   "0-1200-1200-1200-1800-1800-1200-1700-600"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-C-L-L-L-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin NegForms.FlexEdit feTasacVenc 
            Height          =   3810
            Left            =   -74880
            TabIndex        =   13
            Top             =   480
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6720
            Cols0           =   8
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Agencia-N° Garantía-Ult. Crédito-Titular Garantía-Ult. Analista-Fecha Vencimiento-Días"
            EncabezadosAnchos=   "0-1200-1200-1200-1800-1800-1700-600"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-C-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            TextArray0      =   "N°"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Label Label2 
         Caption         =   "En esta opción podrá visualizar las garantías, tasaciones y pólizas vencidas a la fecha."
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
         Left            =   195
         TabIndex        =   11
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label Label1 
         Caption         =   "En esta opción podrá visualizar las garantías, tasaciones y pólizas por vencerse hasta fines del próximo mes."
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
         Left            =   -74805
         TabIndex        =   1
         Top             =   480
         Width           =   8055
      End
   End
End
Attribute VB_Name = "frmCredVerDocsPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
'** Nombre : frmCredVerDocsPendiente
'** Descripción : Para visualizar documentos pendientes creado segun TI-ERS034-2014
'** Creación : EJVG, 20130416 07:26:00 PM
'**********************************************************************************
Option Explicit
Dim fbFiltraAge As Boolean

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Function CargarDatos() As Boolean
    Dim oNGar As New COMNCredito.NCOMGarantia
    Dim oRSGarXVenc As ADODB.Recordset
    Dim oRSGarVenc As ADODB.Recordset
    Dim oRSPolXVenc As ADODB.Recordset
    Dim oRSPolVenc As ADODB.Recordset
    Dim oRSTasacXVenc As ADODB.Recordset
    Dim oRSTasacVenc As ADODB.Recordset
    
    On Error GoTo ErrCargarDatos
    Screen.MousePointer = 11
    oNGar.CargarDatosDocsPendiente fbFiltraAge, gsCodUser, Right(gsCodAge, 2), gdFecSis, oRSGarXVenc, oRSGarVenc, oRSPolXVenc, oRSPolVenc, oRSTasacXVenc, oRSTasacVenc
    If oRSGarXVenc.RecordCount > 0 Or oRSGarVenc.RecordCount > 0 _
        Or oRSPolXVenc.RecordCount > 0 Or oRSPolVenc.RecordCount > 0 _
        Or oRSTasacXVenc.RecordCount > 0 Or oRSTasacVenc.RecordCount > 0 Then
        LlenarGrilla feGarantiaXVenc, oRSGarXVenc
        LlenarGrilla feGarantiaVenc, oRSGarVenc
        LlenarGrillaPoliza fePolizaXVenc, oRSPolXVenc
        LlenarGrillaPoliza fePolizaVenc, oRSPolVenc
        LlenarGrilla feTasacXVenc, oRSTasacXVenc
        LlenarGrilla feTasacVenc, oRSTasacVenc
        CargarDatos = True
    End If
    Screen.MousePointer = 0

    Set oRSGarXVenc = Nothing
    Set oRSGarVenc = Nothing
    Set oRSPolXVenc = Nothing
    Set oRSPolVenc = Nothing
    Set oRSTasacXVenc = Nothing
    Set oRSTasacVenc = Nothing
    Set oNGar = Nothing
    Exit Function
ErrCargarDatos:
    Screen.MousePointer = 0
    CargarDatos = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Public Sub Inicio(ByVal pbFiltraAge As Boolean)
    fbFiltraAge = pbFiltraAge
    Caption = "Documentos pendientes"
    If CargarDatos Then
        Show 1
    End If
End Sub
Private Sub LlenarGrilla(Flex As FlexEdit, pRs As ADODB.Recordset)
    Dim fila As Long
    'FormateaFlex Flex
    Do While Not pRs.EOF
        Flex.AdicionaFila
        fila = Flex.row
        Flex.TextMatrix(fila, 1) = IIf(IsNull(pRs!cAgeDescripcion), "", pRs!cAgeDescripcion)
        Flex.TextMatrix(fila, 2) = IIf(IsNull(pRs!cNumGarant), "", pRs!cNumGarant)
        Flex.TextMatrix(fila, 3) = IIf(IsNull(pRs!cCtaCodVer), "", pRs!cCtaCodVer)
        Flex.TextMatrix(fila, 4) = IIf(IsNull(pRs!cPersNombreTit), "", pRs!cPersNombreTit)
        Flex.TextMatrix(fila, 5) = IIf(IsNull(pRs!cPersNombreAnal), "", pRs!cPersNombreAnal)
        Flex.TextMatrix(fila, 6) = Format(IIf(IsNull(pRs!dFechaVencimiento), CDate("01/01/1950"), pRs!dFechaVencimiento), gsFormatoFechaView)
        Flex.TextMatrix(fila, 7) = IIf(IsNull(pRs!nDias), 0, pRs!nDias)
        pRs.MoveNext
    Loop
End Sub
Private Sub LlenarGrillaPoliza(Flex As FlexEdit, pRs As ADODB.Recordset)
    Dim fila As Long
    FormateaFlex Flex
    Do While Not pRs.EOF
        Flex.AdicionaFila
        fila = Flex.row
        Flex.TextMatrix(fila, 1) = IIf(IsNull(pRs!cAgeDescripcion), "", pRs!cAgeDescripcion)
        Flex.TextMatrix(fila, 2) = IIf(IsNull(pRs!cNumGarant), "", pRs!cNumGarant)
        Flex.TextMatrix(fila, 3) = IIf(IsNull(pRs!cCtaCodVer), "", pRs!cCtaCodVer)
        Flex.TextMatrix(fila, 4) = IIf(IsNull(pRs!cPersNombreTit), "", pRs!cPersNombreTit)
        Flex.TextMatrix(fila, 5) = IIf(IsNull(pRs!cPersNombreAnal), "", pRs!cPersNombreAnal)
        Flex.TextMatrix(fila, 6) = IIf(IsNull(pRs!cNumPoliza), "", pRs!cNumPoliza)
        Flex.TextMatrix(fila, 7) = Format(IIf(IsNull(pRs!dFechaVencimiento), CDate("01/01/1950"), pRs!dFechaVencimiento), gsFormatoFechaView)
        Flex.TextMatrix(fila, 8) = IIf(IsNull(pRs!nDias), 0, pRs!nDias)
        pRs.MoveNext
    Loop
End Sub
Private Sub cmdExportar_Click()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim xlHoja1 As Excel.Worksheet
    Dim rsData As ADODB.Recordset
    Dim lnFila As Long, lnColumna As Long
    Dim lsArchivo As String

    On Error GoTo ErrExportar
    Screen.MousePointer = 11

    lsArchivo = "\spooler\RptDocsPendiente_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time(), "HHMMSS") & ".xls"
    Set xlsLibro = xlsAplicacion.Workbooks.Add

    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "Por Vencerse"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    
    lnFila = 1
    ImprimirRecorsetAExcel xlsHoja, feGarantiaXVenc.GetRsNew, "GARANTIAS POR VENCERSE", lnFila, 1, RGB(216, 216, 216), True
    lnFila = lnFila + 1
    ImprimirRecorsetAExcel xlsHoja, fePolizaXVenc.GetRsNew, "POLIZAS POR VENCERSE", lnFila, 1, RGB(216, 216, 216), True
    lnFila = lnFila + 1
    ImprimirRecorsetAExcel xlsHoja, feTasacXVenc.GetRsNew, "TASACIONES POR VENCERSE", lnFila, 1, RGB(216, 216, 216), True
    
    xlsHoja.Cells.EntireColumn.AutoFit
    
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "Vencidos"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    
    lnFila = 1
    ImprimirRecorsetAExcel xlsHoja, feGarantiaVenc.GetRsNew, "GARANTIAS VENCIDAS", lnFila, 1, RGB(216, 216, 216), True
    lnFila = lnFila + 1
    ImprimirRecorsetAExcel xlsHoja, fePolizaVenc.GetRsNew, "POLIZAS VENCIDAS", lnFila, 1, RGB(216, 216, 216), True
    lnFila = lnFila + 1
    ImprimirRecorsetAExcel xlsHoja, feTasacVenc.GetRsNew, "TASACIONES VENCIDAS", lnFila, 1, RGB(216, 216, 216), True
    
    xlsHoja.Cells.EntireColumn.AutoFit
    
    For Each xlHoja1 In xlsLibro.Worksheets
        If UCase(xlHoja1.Name) = "HOJA1" Or UCase(xlHoja1.Name) = "HOJA2" Or UCase(xlHoja1.Name) = "HOJA3" Then
            xlHoja1.Delete
        End If
    Next
    
    MsgBox "Se ha exportado satisfactoriamente la información", vbInformation, "Aviso"
    
    xlsHoja.SaveAs App.Path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Screen.MousePointer = 0
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    Exit Sub
ErrExportar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub ImprimirRecorsetAExcel(ByRef xlsHoja As Excel.Worksheet, ByVal poRS As ADODB.Recordset, _
                                    Optional ByVal psTitulo As String = "", _
                                    Optional ByRef pnRowExcel As Long = 1, Optional ByVal pnColExcel As Long = 1, _
                                    Optional ByVal pColorRGB As Long = -1, Optional ByVal pbLineas As Boolean = True)
    Dim Col As Long
    Dim ColExcel As Long
    Dim rowIni As Long
    
    If Not poRS Is Nothing Then
        If psTitulo <> "" Then
            xlsHoja.Cells(pnRowExcel, pnColExcel) = UCase(psTitulo)
            xlsHoja.Cells(pnRowExcel, pnColExcel).Font.Bold = True
            pnRowExcel = pnRowExcel + 1
        End If
        
        rowIni = pnRowExcel
        ColExcel = pnColExcel
        For Col = 0 To poRS.Fields.Count - 1
            xlsHoja.Cells(pnRowExcel, ColExcel) = poRS.Fields(Col).Name
            ColExcel = ColExcel + 1
        Next
        
        xlsHoja.Range(xlsHoja.Cells(pnRowExcel, pnColExcel), xlsHoja.Cells(pnRowExcel, ColExcel - 1)).HorizontalAlignment = xlCenter
        xlsHoja.Range(xlsHoja.Cells(pnRowExcel, pnColExcel), xlsHoja.Cells(pnRowExcel, ColExcel - 1)).Font.Bold = True
        If pColorRGB <> -1 Then
            xlsHoja.Range(xlsHoja.Cells(pnRowExcel, pnColExcel), xlsHoja.Cells(pnRowExcel, ColExcel - 1)).Interior.Color = pColorRGB
        End If
        pnRowExcel = pnRowExcel + 1
        xlsHoja.Cells(pnRowExcel, pnColExcel).CopyFromRecordset poRS
        pnRowExcel = pnRowExcel + poRS.RecordCount - 1
        If pbLineas Then
            xlsHoja.Range(xlsHoja.Cells(rowIni, pnColExcel), xlsHoja.Cells(pnRowExcel, ColExcel - 1)).Borders.Weight = xlThin
        End If
        pnRowExcel = pnRowExcel + 1
    End If
End Sub
