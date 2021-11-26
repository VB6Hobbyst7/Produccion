VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRegVisitaAnalista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Visita Analista"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   Icon            =   "frmCredRegVisitaAnalista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   375
      Left            =   6000
      TabIndex        =   35
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9120
      TabIndex        =   11
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   6000
      Width           =   1455
   End
   Begin TabDlg.SSTab sstDatos 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos de la Visita"
      TabPicture(0)   =   "frmCredRegVisitaAnalista.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraComentariosAnalista"
      Tab(0).Control(1)=   "fraDatosGenerales"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Créditos Vigentes(CMACM)"
      TabPicture(1)   =   "frmCredRegVisitaAnalista.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDestinoObservacion"
      Tab(1).Control(1)=   "fraCreditosEvaluacion"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Desarrollo Crediticio"
      TabPicture(2)   =   "frmCredRegVisitaAnalista.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDesarrolloCrediticio"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Deuda Central de Riesgo"
      TabPicture(3)   =   "frmCredRegVisitaAnalista.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fraDeudasCentral"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraDeudasCentral 
         Caption         =   "3. Deudas en Central de Riesgos SBS al: "
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
         Height          =   3495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   10215
         Begin SICMACT.FlexEdit feDeudasCentral 
            Height          =   3015
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   5318
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Entidad-Moneda-Saldo-Calificación-%"
            EncabezadosAnchos=   "500-4000-1200-1200-1200-1200"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C-R-C-C"
            FormatosEdit    =   "0-0-0-2-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraDesarrolloCrediticio 
         Caption         =   "2. Desarrollo Crediticio (Días de Atraso de últimas 06 Cuotas)"
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
         Height          =   3495
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   10215
         Begin SICMACT.FlexEdit feDesarrolloCred 
            Height          =   3015
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   5318
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Nº Crédito-6-5-4-3-2-1"
            EncabezadosAnchos=   "500-1800-1250-1250-1250-1250-1250-1250"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraDestinoObservacion 
         Caption         =   "1.1. Destino/Observación"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   5
         Top             =   2280
         Width           =   10215
         Begin SICMACT.FlexEdit feObservaciones 
            Height          =   1335
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2355
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Nº Crédito-Destino-Observaciones-Aux"
            EncabezadosAnchos=   "500-1800-2500-5000-0"
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
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-C"
            FormatosEdit    =   "0-0-0-0-0"
            CantEntero      =   15
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraCreditosEvaluacion 
         Caption         =   "1. Créditos a la fecha de Evaluación (en su moneda origen)"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   10215
         Begin SICMACT.FlexEdit feCreditosVigentes 
            Height          =   1335
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2355
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Nº Crédito-Moneda-Fecha Desem.-Monto Desem.-Saldo Capital-Cuotas-Avance Cuota-Fecha Venc."
            EncabezadosAnchos=   "500-1800-800-1200-1200-1200-800-1200-1200"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-5-0-0-0-0-5"
            CantEntero      =   15
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraComentariosAnalista 
         Caption         =   "Comentarios Analista"
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
         Height          =   1575
         Left            =   -74880
         TabIndex        =   3
         Top             =   2280
         Width           =   10215
         Begin VB.TextBox txtComentarios 
            Height          =   645
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   720
            Width           =   6975
         End
         Begin VB.TextBox txtNVisita 
            Height          =   285
            Left            =   1800
            TabIndex        =   28
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblNVisita 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Visita:"
            Height          =   195
            Left            =   480
            TabIndex        =   27
            Top             =   360
            Width           =   870
         End
         Begin VB.Label lblComentarios 
            AutoSize        =   -1  'True
            Caption         =   "Comentarios:"
            Height          =   195
            Left            =   480
            TabIndex        =   26
            Top             =   720
            Width           =   915
         End
      End
      Begin VB.Frame fraDatosGenerales 
         Caption         =   "Datos Generales"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   10215
         Begin VB.TextBox txtRelacion 
            Height          =   285
            Left            =   6600
            TabIndex        =   25
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtEntrevistado 
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label lblAnalista 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1800
            TabIndex        =   23
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblGiroNegocio 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1800
            TabIndex        =   22
            Top             =   960
            Width           =   6975
         End
         Begin VB.Label lblDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1800
            TabIndex        =   21
            Top             =   240
            Width           =   6975
         End
         Begin VB.Label lblAnalistaDesc 
            AutoSize        =   -1  'True
            Caption         =   "Analista:"
            Height          =   195
            Left            =   480
            TabIndex        =   20
            Top             =   1320
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Relación:"
            Height          =   195
            Left            =   5880
            TabIndex        =   19
            Top             =   600
            Width           =   675
         End
         Begin VB.Label lblGiroNegocioDesc 
            AutoSize        =   -1  'True
            Caption         =   "Giro del Negocio:"
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   960
            Width           =   1230
         End
         Begin VB.Label lblDireccionDesc 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   480
            TabIndex        =   17
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Entrevistado:"
            Height          =   195
            Left            =   480
            TabIndex        =   16
            Top             =   600
            Width           =   930
         End
      End
   End
   Begin VB.Frame fraClienteSobreendeudado 
      Caption         =   "Cliente Sobreendeudado"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   7560
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin SICMACT.TxtBuscar txtPersona 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         Appearance      =   0
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin MSMask.MaskEdBox txtFecVisita 
         Height          =   300
         Left            =   5760
         TabIndex        =   14
         Top             =   720
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "DD/MM/YYYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFechaVisita 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Visita:"
         Height          =   195
         Left            =   4800
         TabIndex        =   13
         Top             =   765
         Width           =   915
      End
      Begin VB.Label lblNumDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblNombrePersona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   6975
      End
   End
End
Attribute VB_Name = "frmCredRegVisitaAnalista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredRegVisitaAnalista
'***     Descripcion:      Registro de Visita Analista
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     10/09/2013 01:00:00 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Private fsPersCod As String
Private i As Integer

Private Sub cmdCancelar_Click()
txtPersona.Enabled = True
cmdCargar.Enabled = True
txtFecVisita.Enabled = True
txtPersona.Text = ""
lblNombrePersona.Caption = ""
lblNumDoc.Caption = ""
lblDireccion.Caption = ""
lblGiroNegocio.Caption = ""
txtEntrevistado.Text = ""
txtRelacion.Text = ""
txtNVisita.Text = ""
txtcomentarios.Text = ""
txtFecVisita.Text = "__/__/____"
LimpiaFlex feCreditosVigentes
LimpiaFlex feObservaciones
LimpiaFlex feDesarrolloCred
LimpiaFlex feDeudasCentral
fraDatosGenerales.Enabled = True
fraComentariosAnalista.Enabled = True
fraCreditosEvaluacion.Enabled = True
feObservaciones.Enabled = True
feDesarrolloCred.Enabled = True
feDeudasCentral.Enabled = True
cmdImprimir.Enabled = False
fraDeudasCentral.Caption = "3. Deudas en Central de Riesgos SBS al: "
End Sub

Private Sub cmdCargar_Click()
If ValidaDatos(1) Then
    Dim oCredito As COMDCredito.DCOMCredito
    Dim rsCredito As ADODB.Recordset
    Set oCredito = New COMDCredito.DCOMCredito
    
    Set rsCredito = oCredito.ExisteVisitaAnalista(fsPersCod, CDate(txtFecVisita.Text))
    
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        If MsgBox("Existe una Visita Registrada al Cliente con la misma Fecha, Desea Continuar con la Carga?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    
    txtPersona.Enabled = False
    cmdCargar.Enabled = False
    txtFecVisita.Enabled = False
    cmdRegistrar.Enabled = True
    CargarDatos
End If
End Sub

Private Sub cmdImprimir_Click()
GenerarExcel
End Sub

Private Sub cmdRegistrar_Click()
If ValidaDatos(2) Then
    If MsgBox("Estas Seguro de guardar los Datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    
    Dim oCredito As COMDCredito.DCOMCredito
    Dim nCodigo As Long
    Set oCredito = New COMDCredito.DCOMCredito
    
     nCodigo = oCredito.RegistrarVisitaAnalista(fsPersCod, CDate(txtFecVisita.Text), Trim(txtEntrevistado.Text), Trim(txtRelacion.Text), _
                    Trim(lblAnalista.Caption), CLng(txtNVisita.Text), Trim(txtcomentarios.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        
        If Trim(feObservaciones.TextMatrix(1, 0)) <> "" Then
            For i = 1 To feObservaciones.Rows - 1
                Call oCredito.RegistrarVisitaAnalistadetalle(nCodigo, Trim(feObservaciones.TextMatrix(i, 1)), Trim(feObservaciones.TextMatrix(i, 3)))
            Next i
        End If
        MsgBox "Se registro correctamente los Datos", vbInformation, "Aviso"
        fraDatosGenerales.Enabled = False
        fraComentariosAnalista.Enabled = False
        fraCreditosEvaluacion.Enabled = False
        feObservaciones.Enabled = False
        feDesarrolloCred.Enabled = False
        feDeudasCentral.Enabled = False
        cmdRegistrar.Enabled = False
        cmdImprimir.Enabled = True
    End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub feCreditosVigentes_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Cancel = ValidaFlex(feCreditosVigentes, pnCol)
End Sub
Private Sub feDesarrolloCred_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Cancel = ValidaFlex(feDesarrolloCred, pnCol)
End Sub


Private Sub feDeudasCentral_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Cancel = ValidaFlex(feDeudasCentral, pnCol)
End Sub

Private Sub feObservaciones_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Cancel = ValidaFlex(feObservaciones, pnCol)
End Sub

Private Sub Form_Load()
lblAnalista.Caption = UCase(gsCodUser)
cmdImprimir.Enabled = False
cmdRegistrar.Enabled = False
End Sub

Private Sub txtNVisita_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtPersona_EmiteDatos()
Dim rs As ADODB.Recordset
Dim oPers As COMDPersona.DCOMPersonas
On Error GoTo ErrorPersona

fsPersCod = txtPersona.psCodigoPersona

If Trim(fsPersCod) <> "" Then
    Set oPers = New COMDPersona.DCOMPersonas
    Set rs = oPers.RecuperaDatosPersona_Basic(fsPersCod)
    If Not (rs.EOF And rs.BOF) Then
        lblGiroNegocio.Caption = Trim(rs!cActiGiro)
    Else
        lblGiroNegocio.Caption = ""
    End If
    lblNombrePersona.Caption = txtPersona.psDescripcion
    lblNumDoc.Caption = txtPersona.sPersNroDoc
    lblDireccion.Caption = txtPersona.sPersDireccion
End If

    Exit Sub
ErrorPersona:
    MsgBox err.Description, vbInformation, "Error"
End Sub

Private Sub CargarDatos()
Dim oCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Set oCredito = New COMDCredito.DCOMCredito
Set rsCredito = oCredito.MostrarDatosCreditosVisita(fsPersCod, CDate(txtFecVisita.Text))

LimpiaFlex feCreditosVigentes
LimpiaFlex feObservaciones
LimpiaFlex feDesarrolloCred
If Not (rsCredito.EOF And rsCredito.BOF) Then
    For i = 1 To rsCredito.RecordCount
        'Creditos Vigentes
        feCreditosVigentes.AdicionaFila
        feCreditosVigentes.TextMatrix(i, 1) = Trim(rsCredito!cCtaCod)
        feCreditosVigentes.TextMatrix(i, 2) = Trim(rsCredito!Moneda)
        feCreditosVigentes.TextMatrix(i, 3) = Trim(rsCredito!FDesem)
        feCreditosVigentes.TextMatrix(i, 4) = Format(Trim(rsCredito!MontoDesem), "###," & String(15, "#") & "#0.00")
        feCreditosVigentes.TextMatrix(i, 5) = Format(Trim(rsCredito!SalCap), "###," & String(15, "#") & "#0.00")
        feCreditosVigentes.TextMatrix(i, 6) = Trim(rsCredito!nCuotas)
        feCreditosVigentes.TextMatrix(i, 7) = Trim(rsCredito!CuotasPagadas)
        feCreditosVigentes.TextMatrix(i, 8) = Trim(rsCredito!FVecimiento)
        'Destino/Observaciones
        feObservaciones.AdicionaFila
        feObservaciones.TextMatrix(i, 1) = Trim(rsCredito!cCtaCod)
        feObservaciones.TextMatrix(i, 2) = Trim(rsCredito!Destino)
        'Desarrollo Crediticio
        feDesarrolloCred.AdicionaFila
        feDesarrolloCred.TextMatrix(i, 1) = Trim(rsCredito!cCtaCod)
        feDesarrolloCred.TextMatrix(i, 2) = Trim(rsCredito!DiasAtraso6)
        feDesarrolloCred.TextMatrix(i, 3) = Trim(rsCredito!DiasAtraso5)
        feDesarrolloCred.TextMatrix(i, 4) = Trim(rsCredito!DiasAtraso4)
        feDesarrolloCred.TextMatrix(i, 5) = Trim(rsCredito!DiasAtraso3)
        feDesarrolloCred.TextMatrix(i, 6) = Trim(rsCredito!DiasAtraso2)
        feDesarrolloCred.TextMatrix(i, 7) = Trim(rsCredito!DiasAtraso1)
        
        rsCredito.MoveNext
    Next i
Else
    MsgBox "No hay datos sobre Creditos", vbInformation, "Aviso"
End If
Set rsCredito = Nothing

Set rsCredito = oCredito.MostrarDeudaCentralRiesgo(txtPersona.sPersNroDoc, IIf(txtPersona.PersPersoneria = "1", "1", "2"))
LimpiaFlex feDeudasCentral
If Not (rsCredito.EOF And rsCredito.BOF) Then
fraDeudasCentral.Caption = fraDeudasCentral.Caption & Trim(rsCredito!Fecha)
    For i = 1 To rsCredito.RecordCount
        feDeudasCentral.AdicionaFila
        feDeudasCentral.TextMatrix(i, 1) = rsCredito!Entidad
        feDeudasCentral.TextMatrix(i, 2) = rsCredito!Moneda
        feDeudasCentral.TextMatrix(i, 3) = Format(Trim(rsCredito!Saldo), "###," & String(15, "#") & "#0.00")
        feDeudasCentral.TextMatrix(i, 4) = rsCredito!Clasificacion
        feDeudasCentral.TextMatrix(i, 5) = Trim(rsCredito!Porcentaje) & "%"
        rsCredito.MoveNext
    Next i
Else
    MsgBox "No hay datos de Deudas en la Central de Riesgo", vbInformation, "Aviso"
End If
Set rsCredito = Nothing
End Sub

Private Function ValidaDatos(ByVal pnIndex As Integer) As Boolean
ValidaDatos = True
    
If pnIndex = 1 Then
    If Trim(txtPersona.Text) = "" Then
        MsgBox "Seleccione el Cliente", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    If InStr(Trim(txtFecVisita.Text), "_") > 0 Then
        MsgBox "Ingrese correctamente la Fecha de la Visita", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
ElseIf pnIndex = 2 Then
    If Trim(txtEntrevistado.Text) = "" Then
        MsgBox "Ingrese al Entrevistado", vbInformation, "Aviso"
        ValidaDatos = False
        sstDatos.Tab = 0
        txtEntrevistado.SetFocus
        Exit Function
    End If
    If Trim(txtRelacion.Text) = "" Then
        MsgBox "Ingrese la Relación del Entrevistado", vbInformation, "Aviso"
        ValidaDatos = False
        sstDatos.Tab = 0
        txtRelacion.SetFocus
        Exit Function
    End If
    If Trim(txtNVisita.Text) = "" Or Trim(txtNVisita.Text) = "0" Then
        MsgBox "Ingrese el Nº de Visita", vbInformation, "Aviso"
        ValidaDatos = False
        sstDatos.Tab = 0
        txtNVisita.SetFocus
        Exit Function
    End If
    
    If Trim(txtcomentarios.Text) = "" Then
        MsgBox "Ingrese los Comentarios", vbInformation, "Aviso"
        ValidaDatos = False
        sstDatos.Tab = 0
        txtcomentarios.SetFocus
        Exit Function
    End If

    For i = 1 To feObservaciones.Rows - 1
        If feObservaciones.TextMatrix(i, 3) = "" Then
            MsgBox "Ingrese la observación del Crédito Nº " & Trim(feObservaciones.TextMatrix(i, 1)), vbInformation, "Aviso"
            ValidaDatos = False
            sstDatos.Tab = 1
            feDesarrolloCred.SetFocus
            Exit Function
        End If
    Next i

End If
End Function

Private Sub GenerarExcel()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    
    Dim lnExcel As Long
    
    On Error GoTo ErrorGeneraExcelFormato
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    lsNomHoja = "Visita"
    lsFile = "ClienteSobreendeudado"
    
    lsArchivo = "\spooler\" & "ClienteSobreendeudado_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    'Activar Hoja
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
    
    
    xlHoja1.Cells(3, 2) = Trim(UCase(lblDireccion.Caption))
    xlHoja1.Cells(4, 2) = Trim(UCase(lblNombrePersona.Caption))
    xlHoja1.Cells(5, 2) = Trim(UCase(txtEntrevistado.Text)) & "(" & Trim(UCase(Me.txtRelacion.Text)) & ")"
    xlHoja1.Cells(6, 2) = Trim(UCase(lblGiroNegocio.Caption))
    xlHoja1.Cells(6, 8) = Trim(UCase(lblAnalista.Caption))
        
    xlHoja1.Cells(82, 7) = Trim(txtNVisita.Text)
    xlHoja1.Cells(82, 12) = Trim(txtFecVisita.Text)
    xlHoja1.Cells(83, 1) = Trim(UCase(txtcomentarios.Text))
    
    lnExcel = 11
    If Trim(feCreditosVigentes.TextMatrix(1, 0)) <> "" Then
        For i = 1 To feCreditosVigentes.Rows - 1
            xlHoja1.Cells(lnExcel, 2).NumberFormat = "@"
            xlHoja1.Cells(lnExcel, 2) = feCreditosVigentes.TextMatrix(i, 1)
            xlHoja1.Cells(lnExcel, 3) = feCreditosVigentes.TextMatrix(i, 2)
            xlHoja1.Cells(lnExcel, 4) = feCreditosVigentes.TextMatrix(i, 3)
            xlHoja1.Cells(lnExcel, 5) = feCreditosVigentes.TextMatrix(i, 4)
            xlHoja1.Cells(lnExcel, 6) = feCreditosVigentes.TextMatrix(i, 5)
            xlHoja1.Cells(lnExcel, 7) = feCreditosVigentes.TextMatrix(i, 6)
            xlHoja1.Cells(lnExcel, 8) = feCreditosVigentes.TextMatrix(i, 7)
            xlHoja1.Cells(lnExcel, 9) = feCreditosVigentes.TextMatrix(i, 8)
            lnExcel = lnExcel + 1
        Next i
    End If
    xlHoja1.Range(xlHoja1.Cells(lnExcel, 1), xlHoja1.Cells(25, 13)).Delete
    
    lnExcel = lnExcel + 3
    If Trim(feObservaciones.TextMatrix(1, 0)) <> "" Then
        For i = 1 To feObservaciones.Rows - 1
            xlHoja1.Cells(lnExcel, 2).NumberFormat = "@"
            xlHoja1.Cells(lnExcel, 2) = feObservaciones.TextMatrix(i, 1)
            xlHoja1.Cells(lnExcel, 3) = feObservaciones.TextMatrix(i, 2)
            xlHoja1.Cells(lnExcel, 5) = feObservaciones.TextMatrix(i, 3)
            lnExcel = lnExcel + 1
        Next i
    End If
    xlHoja1.Range(xlHoja1.Cells(lnExcel, 1), xlHoja1.Cells(lnExcel + (15 - i), 13)).Delete
    
    lnExcel = lnExcel + 3
    If Trim(feDesarrolloCred.TextMatrix(1, 0)) <> "" Then
        For i = 1 To feDesarrolloCred.Rows - 1
            xlHoja1.Cells(lnExcel, 2).NumberFormat = "@"
            xlHoja1.Cells(lnExcel, 2) = feDesarrolloCred.TextMatrix(i, 1)
            xlHoja1.Cells(lnExcel, 3) = feDesarrolloCred.TextMatrix(i, 2)
            xlHoja1.Cells(lnExcel, 4) = feDesarrolloCred.TextMatrix(i, 3)
            xlHoja1.Cells(lnExcel, 5) = feDesarrolloCred.TextMatrix(i, 4)
            xlHoja1.Cells(lnExcel, 6) = feDesarrolloCred.TextMatrix(i, 5)
            xlHoja1.Cells(lnExcel, 7) = feDesarrolloCred.TextMatrix(i, 6)
            xlHoja1.Cells(lnExcel, 8) = feDesarrolloCred.TextMatrix(i, 7)
            lnExcel = lnExcel + 1
        Next i
    End If
    
    xlHoja1.Range(xlHoja1.Cells(lnExcel, 1), xlHoja1.Cells(lnExcel + (15 - i), 13)).Delete
    
    lnExcel = lnExcel + 3
    If Trim(feDeudasCentral.TextMatrix(1, 0)) <> "" Then
    xlHoja1.Cells(lnExcel - 2, 1) = xlHoja1.Cells(lnExcel - 2, 1) & Trim(Right(fraDeudasCentral.Caption, 10))
        For i = 1 To feDeudasCentral.Rows - 1
            xlHoja1.Cells(lnExcel, 2) = feDeudasCentral.TextMatrix(i, 1)
            xlHoja1.Cells(lnExcel, 3) = feDeudasCentral.TextMatrix(i, 2)
            xlHoja1.Cells(lnExcel, 4) = feDeudasCentral.TextMatrix(i, 3)
            xlHoja1.Cells(lnExcel, 5) = feDeudasCentral.TextMatrix(i, 4)
            xlHoja1.Cells(lnExcel, 6) = feDeudasCentral.TextMatrix(i, 5)
            lnExcel = lnExcel + 1
        Next i
    End If
    xlHoja1.Range(xlHoja1.Cells(lnExcel, 1), xlHoja1.Cells(lnExcel + (15 - i), 13)).Delete

    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
    
    Exit Sub
ErrorGeneraExcelFormato:
    MsgBox err.Description, vbCritical, "Error a Generar El Formato Excel"
End Sub

