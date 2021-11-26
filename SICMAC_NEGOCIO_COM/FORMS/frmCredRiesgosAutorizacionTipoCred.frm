VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgosAutorizacionTipoCred 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de limites de Autorizaciones"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredRiesgosAutorizacionTipoCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   15675
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab stabAutorizaciones 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Registro Mensual Autorizaciones"
      TabPicture(0)   =   "frmCredRiesgosAutorizacionTipoCred.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Registro Anual Autorizaciones"
      TabPicture(1)   =   "frmCredRiesgosAutorizacionTipoCred.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frameCabecera"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameAutorizaciones"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdMostrar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdNuevo"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdEditar"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdGuardar"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdCancelar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame1 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   15015
         Begin VB.Frame frameParametros 
            Caption         =   "Parametros de registro"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   5055
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   7455
            Begin VB.Frame Frame2 
               Height          =   855
               Left            =   240
               TabIndex        =   36
               Top             =   3720
               Width           =   7095
               Begin VB.CommandButton Command7 
                  Caption         =   "Guardar"
                  Height          =   400
                  Left            =   4320
                  TabIndex        =   40
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "Editar"
                  Height          =   400
                  Left            =   1560
                  TabIndex        =   39
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "Nuevo"
                  Height          =   400
                  Left            =   240
                  TabIndex        =   38
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.CommandButton Command6 
                  Caption         =   "Cancelar"
                  Height          =   400
                  Left            =   5640
                  TabIndex        =   37
                  Top             =   240
                  Width           =   1200
               End
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   1200
               Width           =   1425
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   1680
               Width           =   2625
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   2160
               Width           =   4785
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Año:"
               Height          =   195
               Left            =   1920
               TabIndex        =   34
               Top             =   1200
               Width           =   345
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mes :"
               Height          =   195
               Left            =   1920
               TabIndex        =   33
               Top             =   1680
               Width           =   390
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Autoriazación / Exoneración : "
               Height          =   195
               Left            =   240
               TabIndex        =   32
               Top             =   2160
               Width           =   2160
            End
         End
         Begin VB.Frame frameAgencias 
            Caption         =   "Agencias / Factor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   5055
            Left            =   7920
            TabIndex        =   27
            Top             =   360
            Width           =   6975
            Begin SICMACT.FlexEdit FlexEdit2 
               Height          =   4215
               Left            =   240
               TabIndex        =   28
               Top             =   480
               Width           =   6285
               _ExtentX        =   11086
               _ExtentY        =   7435
               Cols0           =   5
               HighLight       =   1
               EncabezadosNombres=   "Aux1-Agencia-Factor-Estado-Aux2"
               EncabezadosAnchos=   "0-3000-1800-1200-0"
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
               ColumnasAEditar =   "X-1-2-3-X"
               ListaControles  =   "0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-L-C-C"
               FormatosEdit    =   "0-0-0-0-0"
               TextArray0      =   "Aux1"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   6
               lbBuscaDuplicadoText=   -1  'True
               RowHeight0      =   300
            End
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   14280
         TabIndex        =   25
         Top             =   6160
         Width           =   1200
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   360
         Left            =   12960
         TabIndex        =   24
         Top             =   6160
         Width           =   1200
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   360
         Left            =   1440
         TabIndex        =   23
         Top             =   6160
         Width           =   1200
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   360
         Left            =   120
         TabIndex        =   22
         Top             =   6160
         Width           =   1200
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   360
         Left            =   9360
         TabIndex        =   20
         Top             =   540
         Width           =   1215
      End
      Begin VB.Frame frameAutorizaciones 
         Caption         =   "Registro por meses de Autorizaciones /  Producto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   5055
         Left            =   50
         TabIndex        =   3
         Top             =   1080
         Width           =   15480
         Begin SICMACT.FlexEdit feAutorizacion 
            Height          =   1935
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   15285
            _ExtentX        =   26961
            _ExtentY        =   3413
            Cols0           =   20
            HighLight       =   1
            EncabezadosNombres=   "-Autorizacion / Tpo.Cred-Min-Prom-Max-Min-Prom-Max-Min-Prom-Max-Min-Prom-Max-Min-Prom-Max-Min-Prom-Max"
            EncabezadosAnchos=   "0-2500-700-700-700-700-700-700-700-700-700-700-700-700-700-700-700-700-700-700"
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
            ColumnasAEditar =   "X-X-2-3-4-5-6-7-8-9-10-11-12-13-14-15-16-17-18-19"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R"
            FormatosEdit    =   "0-0-2-2-2-2-2-2-2-2-2-2-2-2-2-2-2-2-2-2"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin SICMACT.FlexEdit feAutorizacion2 
            Height          =   2055
            Left            =   120
            TabIndex        =   5
            Top             =   2880
            Width           =   15280
            _ExtentX        =   26961
            _ExtentY        =   3625
            Cols0           =   20
            HighLight       =   1
            EncabezadosNombres=   "-Autorizacion / Tpo.Cred-Min-Prom-Max-Min-Prom-Max-Min-Prom-Max-Min-Prom-Max-Min-Prom-Max-Min-Prom-Max"
            EncabezadosAnchos=   "0-2500-700-700-700-700-700-700-700-700-700-700-700-700-700-700-700-700-700-700"
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
            ColumnasAEditar =   "X-X-2-3-4-5-6-7-8-9-10-11-12-13-14-15-16-17-18-19"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R"
            FormatosEdit    =   "0-0-2-2-2-2-2-2-2-2-2-2-2-2-2-2-2-2-2-2"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Julio"
            Height          =   255
            Left            =   2645
            TabIndex        =   17
            Top             =   2640
            Width           =   2130
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Agosto"
            Height          =   255
            Left            =   4775
            TabIndex        =   16
            Top             =   2640
            Width           =   2130
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Setiembre"
            Height          =   255
            Left            =   6865
            TabIndex        =   15
            Top             =   2640
            Width           =   2130
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Octubre"
            Height          =   255
            Left            =   8995
            TabIndex        =   14
            Top             =   2640
            Width           =   2130
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Noviembre"
            Height          =   255
            Left            =   11110
            TabIndex        =   13
            Top             =   2640
            Width           =   2130
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diciembre"
            Height          =   255
            Left            =   13225
            TabIndex        =   12
            Top             =   2640
            Width           =   2130
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Junio"
            Height          =   255
            Left            =   13225
            TabIndex        =   11
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mayo"
            Height          =   255
            Left            =   11110
            TabIndex        =   10
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abril"
            Height          =   255
            Left            =   8995
            TabIndex        =   9
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Marzo"
            Height          =   255
            Left            =   6865
            TabIndex        =   8
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Febrero"
            Height          =   255
            Left            =   4775
            TabIndex        =   7
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Enero"
            Height          =   255
            Left            =   2645
            TabIndex        =   6
            Top             =   240
            Width           =   2130
         End
      End
      Begin VB.Frame frameCabecera 
         Caption         =   "Buscar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9135
         Begin VB.ComboBox cmbAutorizacion 
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   195
            Width           =   4785
         End
         Begin VB.ComboBox cboAnio 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   195
            Width           =   1545
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Autorización / Exoneración:"
            Height          =   195
            Left            =   2280
            TabIndex        =   2
            Top             =   240
            Width           =   1980
         End
      End
   End
End
Attribute VB_Name = "frmCredRiesgosAutorizacionTipoCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre      : frmCredRiesgosAutorizacionTipoCred                                            *
'** Descripción : Formulario para Registro/ Mantenimiento de Autorizaciones x Producto          *
'** Referencia  : ER0S65-2016                                                                   *
'** Creación    : ARLO, 20170601 09:00:00 AM                                                    *
'************************************************************************************************
Option Explicit
Dim cols As Integer
Dim rows As Integer
Dim i, j As Integer

Public Function Inicio()
    Call CargaAutorizaciones
    Call CargarAño
    Call CargaTpoCredito
    Call CargaTpoCredito2
    stabAutorizaciones.TabVisible(0) = False
    feAutorizacion.Enabled = False
    cmdGuardar.Enabled = False
    Me.cmdNuevo.Enabled = False
    Me.cmdEditar.Enabled = False
    frmCredRiesgosAutorizacionTipoCred.Show 1
End Function
Private Sub CargarAño()
    
    Dim nAño As String
    Dim nAño2 As Integer
    Dim nAñoIncio As Integer
    
    nAñoIncio = 2016
    nAño = Year(gdFecSis)
    nAño2 = CInt(nAño)
    Do
    cboAnio.AddItem "" & nAñoIncio
    cboAnio.ItemData(cboAnio.NewIndex) = "" & nAñoIncio
    nAñoIncio = nAñoIncio + 1
    Loop While (nAñoIncio <= nAño2 + 1)
    cboAnio.ListIndex = 0

End Sub
Public Function CargaAutorizaciones() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim lnAgeCodAct As Integer
    Set rs = New ADODB.Recordset
    
    
    Set rs = DevuelveAutorizaciones()
    
    CargarComboBox rs, cmbAutorizacion
    
    rs.Close
    Set rs = Nothing
    End Function
Public Function DevuelveAutorizaciones() As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaAutorizacionAnual"
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveAutorizaciones = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function CargaTpoCredito() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim nNumFila As Integer
    
    Set rs = DevuelveTpoCredito()
    
    Do While Not rs.EOF
            
            feAutorizacion.AdicionaFila
            nNumFila = feAutorizacion.rows - 1
                       
            feAutorizacion.TextMatrix(nNumFila, 1) = rs!cConsDescripcion
            rs.MoveNext
      Loop
    rs.Close
    Set rs = Nothing
    End Function
    Public Function CargaTpoCredito2() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim nNumFila As Integer
    
    Set rs = DevuelveTpoCredito()
    
    Do While Not rs.EOF
                feAutorizacion2.AdicionaFila
                nNumFila = feAutorizacion2.rows - 1
                feAutorizacion2.TextMatrix(nNumFila, 1) = rs!cConsDescripcion
                rs.MoveNext
      Loop
    rs.Close
    Set rs = Nothing
    End Function
    Public Function DevuelveTpoCredito() As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaTpoCreditoAutorizacion"
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveTpoCredito = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Private Sub Cmdguardar_Click()

Dim nFacMin As Double
Dim nFacPro As Double
Dim nFacMax As Double
Dim nMes As Integer
Dim lsAño As String
Dim ntipoCred As Integer
Dim nExoneraCod As Integer
Dim lsMovNro As String
Dim lsMes As String



lsAño = cboAnio.ItemData(cboAnio.ListIndex)
nExoneraCod = cmbAutorizacion.ItemData(cmbAutorizacion.ListIndex)
lsMovNro = GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser)
nMes = 0
ntipoCred = 0

If FlexVacio(feAutorizacion) Then
        MsgBox " Por favor, llene todos los datos. El Factor debe ser Mayor que cero (0).", vbInformation, "Aviso"
Exit Sub
End If

If FlexVacio1(feAutorizacion2) Then
        MsgBox " Por favor, llene todos los datos. El Factor debe ser Mayor que cero (0).", vbInformation, "Aviso"
Exit Sub
End If

If cboAnio.ListIndex = -1 Then
        MsgBox "Eliga un Año, Por Favor", vbInformation, "Aviso"
Exit Sub
End If

If cmbAutorizacion.ListIndex = -1 Then
        MsgBox "Eliga una Autorización, Por Favor", vbInformation, "Aviso"
Exit Sub
End If

If MsgBox("Seguro Desea Grabar los Datos ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub


'Primera Matriz
cols = feAutorizacion.cols - 1
rows = feAutorizacion.rows - 1
For j = 1 To rows
ntipoCred = ntipoCred + 1
    For i = 2 To cols
        nFacMin = feAutorizacion.TextMatrix(j, i)
        nFacPro = feAutorizacion.TextMatrix(j, i + 1)
        nFacMax = feAutorizacion.TextMatrix(j, i + 2)
        i = i + 2
        nMes = nMes + 1
        If Len(CStr(nMes)) = 1 Then
        lsMes = "0" + CStr(nMes)
        Else
        lsMes = nMes
        End If
        Call GrabaDatosAnuales(lsMovNro, lsAño, lsMes, nExoneraCod, ntipoCred, nFacMin, nFacPro, nFacMax)
        If (nMes = 6) Then
        nMes = 0
        End If
    Next i
Next j

'Segunda Matriz
nMes = 6
ntipoCred = 0
cols = feAutorizacion2.cols - 1
rows = feAutorizacion2.rows - 1
For j = 1 To rows
ntipoCred = ntipoCred + 1
    For i = 2 To cols
        nFacMin = feAutorizacion2.TextMatrix(j, i)
        nFacPro = feAutorizacion2.TextMatrix(j, i + 1)
        nFacMax = feAutorizacion2.TextMatrix(j, i + 2)
        i = i + 2
        nMes = nMes + 1
        If Len(CStr(nMes)) = 1 Then
        lsMes = "0" + CStr(nMes)
        Else
        lsMes = CStr(nMes)
        End If
        Call GrabaDatosAnuales(lsMovNro, lsAño, lsMes, nExoneraCod, ntipoCred, nFacMin, nFacPro, nFacMax)
        If (nMes = 12) Then
        nMes = 6
        End If
    Next i
Next j

MsgBox "Se Registraron los Datos Correctamente", vbInformation, "Aviso"

    Me.feAutorizacion.Clear
    Me.feAutorizacion.rows = 2
    Me.feAutorizacion.FormaCabecera
    CargaTpoCredito
    Me.feAutorizacion2.Clear
    Me.feAutorizacion2.rows = 2
    Me.feAutorizacion2.FormaCabecera
    CargaTpoCredito2
    Me.cmdNuevo.Enabled = False
    Me.cmdEditar.Enabled = False
    Me.cmdGuardar.Enabled = False
    Me.feAutorizacion.Enabled = False
    Me.feAutorizacion2.Enabled = False


End Sub

Public Sub GrabaDatosAnuales(ByVal lsMovNro As String, ByVal lsAño As String, ByVal lsMes As String, ByVal nExoneraCod As Integer, ByVal ntipoCred As String, _
                                ByVal nFacMin As Double, ByVal nFacPro As Double, ByVal nFacMax As Double)
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim pbTran As Boolean
    
    lsSQL = "exec stp_ins_ERS0652016_AutorizacionConFigFactor '" & lsMovNro & "','" & lsAño & "','" & lsMes & "'," & nExoneraCod & "," & ntipoCred & "," & nFacMin & "," & nFacPro & "," & nFacMax
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.Ejecutar (lsSQL)
    loReg.CierraConexion
End Sub

Private Sub cmdMostrar_Click()
    Dim nAño As Integer
    Dim nExoneraCod As String
    Dim rs As ADODB.Recordset
    
    nAño = cboAnio.ItemData(cboAnio.ListIndex)
    nExoneraCod = cmbAutorizacion.ItemData(cmbAutorizacion.ListIndex)
    
    Call CargaTpoFactor(nAño, nExoneraCod)
    
    Set rs = ValidaRegistro(nAño, nExoneraCod)
    
        If (rs.RecordCount > 0) Then
                If (DateDiff("m", gdFecSis, rs!dFechaReg) = 0) Then
                    Me.cmdEditar.Enabled = True
                Else
                    Me.cmdEditar.Enabled = False
                End If
        End If

End Sub
Public Function CargaTpoFactor(ByVal nAño As Integer, ByVal nExoneraCod As Integer) As ADODB.Recordset
    Dim rs, rs1 As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Dim nNumFila As Integer
    
    Set rs = DevuelveTpoFactor(nAño, nExoneraCod)
    Set rs1 = DevuelveTpoFactor1(nAño, nExoneraCod)
    
    cols = feAutorizacion.cols - 1
    rows = feAutorizacion.rows - 1
    
    If (rs.RecordCount > 0) Then

    
            For j = 1 To rows
            For i = 2 To cols
            feAutorizacion.TextMatrix(j, i) = rs!nFactorMin
            feAutorizacion.TextMatrix(j, i + 1) = rs!nFactorProm
            feAutorizacion.TextMatrix(j, i + 2) = rs!nFactorMax
            i = i + 2
            rs.MoveNext
            Next i
            Next j
            Me.cmdNuevo.Enabled = False
            Me.cmdEditar.Enabled = True
            Me.feAutorizacion.Enabled = False
     End If
     If (rs1.RecordCount > 0) Then
            For j = 1 To rows
            For i = 2 To cols
            feAutorizacion2.TextMatrix(j, i) = rs1!nFactorMin
            feAutorizacion2.TextMatrix(j, i + 1) = rs1!nFactorProm
            feAutorizacion2.TextMatrix(j, i + 2) = rs1!nFactorMax
            i = i + 2
            rs1.MoveNext
            Next i
            Next j
            Me.cmdNuevo.Enabled = False
            Me.cmdEditar.Enabled = True
            Me.feAutorizacion2.Enabled = False
            
    Else
    MsgBox "No se encontraron Registros", vbInformation, "Aviso"
    Me.feAutorizacion.Clear
    Me.feAutorizacion.rows = 2
    Me.feAutorizacion.FormaCabecera
    CargaTpoCredito
    Me.feAutorizacion2.Clear
    Me.feAutorizacion2.rows = 2
    Me.feAutorizacion2.FormaCabecera
    CargaTpoCredito2
    Me.cmdNuevo.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.feAutorizacion.Enabled = False
    Me.feAutorizacion2.Enabled = False
    End If
    rs.Close
    Set rs = Nothing
    End Function
    Public Function DevuelveTpoFactor(ByVal nAño As Integer, ByVal lsExoneraCod As String) As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaConfTpoFactorAutorizacion " & nAño & "," & lsExoneraCod
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveTpoFactor = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
    Public Function DevuelveTpoFactor1(ByVal nAño As Integer, ByVal lsExoneraCod As String) As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaConfTpoFactorAutorizacion1 " & nAño & "," & lsExoneraCod
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveTpoFactor1 = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Private Sub cmdNuevo_Click()
        Me.cmdEditar.Enabled = False
        Me.cmdGuardar.Enabled = True
        Me.cmdNuevo.Enabled = False
        feAutorizacion.Enabled = True
        feAutorizacion2.Enabled = True
End Sub
Private Sub CmdEditar_Click()
        Me.cmdNuevo.Enabled = False
        Me.cmdGuardar.Enabled = True
        Me.cmdEditar.Enabled = False
        feAutorizacion.Enabled = True
        feAutorizacion2.Enabled = True
End Sub
Private Sub cmdCancelar_Click()
        Unload Me
End Sub
Public Function FlexVacio(ByVal pflex As FlexEdit) As Boolean
    
cols = feAutorizacion.cols - 1
rows = feAutorizacion.rows - 1

For j = 1 To rows
    For i = 2 To cols
    
        If (pflex.TextMatrix(j, i) = "" Or IIf(pflex.TextMatrix(j, i) = "", 0, pflex.TextMatrix(j, i)) = 0) Then
        FlexVacio = True
        ElseIf (pflex.TextMatrix(j, i + 1) = "" Or IIf(pflex.TextMatrix(j, i + 1) = "", 0, pflex.TextMatrix(j, i + 1)) = 0) Then
        FlexVacio = True
        ElseIf (pflex.TextMatrix(j, i + 2) = "" Or IIf(pflex.TextMatrix(j, i + 2) = "", 0, pflex.TextMatrix(j, i + 2)) = 0) Then
        FlexVacio = True
        End If
        i = i + 2
    Next i
Next j

End Function

Public Function FlexVacio1(ByVal pflex As FlexEdit) As Boolean
    
cols = feAutorizacion2.cols - 1
rows = feAutorizacion2.rows - 1

For j = 1 To rows
    For i = 2 To cols
        If (pflex.TextMatrix(j, i) = "" Or IIf(pflex.TextMatrix(j, i) = "", 0, pflex.TextMatrix(j, i)) = 0) Then
        FlexVacio1 = True
        ElseIf (pflex.TextMatrix(j, i + 1) = "" Or IIf(pflex.TextMatrix(j, i + 1) = "", 0, pflex.TextMatrix(j, i + 1)) = 0) Then
        FlexVacio1 = True
        ElseIf (pflex.TextMatrix(j, i + 2) = "" Or IIf(pflex.TextMatrix(j, i + 2) = "", 0, pflex.TextMatrix(j, i + 2)) = 0) Then
        FlexVacio1 = True
        End If
        i = i + 2
    Next i
Next j

End Function
    Public Function ValidaRegistro(ByVal nAño As Integer, ByVal nExoneraCod As String) As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_Sel_ERS0652016_ValidaRegistroAnualAutorizaciones " & nAño & "," & nExoneraCod
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set ValidaRegistro = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'AGREGRADO POR ARLO20170718 INICIO****************************
Private Sub feAutorizacion_OnCellChange(pnRow As Long, pnCol As Long)
            
        If IsNumeric(feAutorizacion.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feAutorizacion.TextMatrix(pnRow, pnCol) < 0 Then
            feAutorizacion.TextMatrix(pnRow, pnCol) = 0
            MsgBox "Por favor, ingrese un número Mayor que cero (0).", vbInformation, "Aviso"
        End If
        Else
        feAutorizacion.TextMatrix(pnRow, pnCol) = 0
        End If
                
End Sub
Private Sub feAutorizacion2_OnCellChange(pnRow As Long, pnCol As Long)

        If IsNumeric(feAutorizacion2.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feAutorizacion2.TextMatrix(pnRow, pnCol) < 0 Then
            feAutorizacion2.TextMatrix(pnRow, pnCol) = 0
            MsgBox "Por favor, ingrese un número Mayor que cero (0).", vbInformation, "Aviso"
        End If
        Else
        feAutorizacion2.TextMatrix(pnRow, pnCol) = 0
        End If
End Sub
'FIN***************************************************
