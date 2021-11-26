VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajaArqueoVentBov 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arqueos: "
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11070
   Icon            =   "frmCajaArqueoVentBov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   10815
      Begin VB.TextBox txtGlosa 
         Height          =   585
         IMEMode         =   3  'DISABLE
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3360
         Width           =   9135
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9480
         TabIndex        =   17
         Top             =   3600
         Width           =   1170
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9480
         TabIndex        =   16
         Top             =   3190
         Width           =   1170
      End
      Begin SICMACT.FlexEdit fgBilletes 
         Height          =   2250
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   3969
         Cols0           =   6
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Denominación-Cant-Monto-cEfectivoCod-nEfectivoValor"
         EncabezadosAnchos=   "350-1350-600-900-0-0"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-R-C-C"
         FormatosEdit    =   "0-0-3-4-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         CellBackColor   =   -2147483633
      End
      Begin SICMACT.FlexEdit fgMonedas 
         Height          =   2250
         Index           =   0
         Left            =   3480
         TabIndex        =   20
         Top             =   360
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   3969
         Cols0           =   6
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Denominación-Cant-Monto-cEfectivoCod-nEfectivoValor"
         EncabezadosAnchos=   "350-1350-600-900-0-0"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-R-C-C"
         FormatosEdit    =   "0-0-3-4-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         CellBackColor   =   -2147483633
      End
      Begin SICMACT.FlexEdit fgBilletes 
         Height          =   2250
         Index           =   1
         Left            =   7080
         TabIndex        =   21
         Top             =   360
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   3969
         Cols0           =   6
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Denominación-Cant-Monto-cEfectivoCod-nEfectivoValor"
         EncabezadosAnchos=   "350-1350-600-900-0-0"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-R-C-C"
         FormatosEdit    =   "0-0-3-4-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         CellBackColor   =   -2147483633
      End
      Begin MSComctlLib.ProgressBar Pg 
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   3680
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblTotalBilletes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
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
         Height          =   300
         Index           =   0
         Left            =   1710
         TabIndex        =   32
         Top             =   2655
         Width           =   1635
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL BILLETES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Index           =   0
         Left            =   210
         TabIndex        =   31
         Top             =   2730
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL MONEDAS :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Index           =   0
         Left            =   3570
         TabIndex        =   30
         Top             =   2730
         Width           =   1440
      End
      Begin VB.Label lblTotMoneda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   0
         Left            =   5040
         TabIndex        =   29
         Top             =   2655
         Width           =   1845
      End
      Begin VB.Label Label6 
         Caption         =   "Billetes Soles :"
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
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Monedas Soles :"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL BILLETES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Index           =   1
         Left            =   7185
         TabIndex        =   26
         Top             =   2730
         Width           =   1410
      End
      Begin VB.Label lblTotalBilletes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "0.00"
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
         Height          =   300
         Index           =   1
         Left            =   8625
         TabIndex        =   25
         Top             =   2655
         Width           =   1950
      End
      Begin VB.Label Label8 
         Caption         =   "Billetes Dólares :"
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
         Left            =   7080
         TabIndex        =   24
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Glosa"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   0
         Left            =   3480
         Top             =   2640
         Width           =   3435
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   0
         Left            =   120
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   1
         Left            =   7080
         Top             =   2640
         Width           =   3525
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Personal Involucrado "
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   10815
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1050
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   345
         Left            =   1200
         TabIndex        =   7
         Top             =   1800
         Width           =   1050
      End
      Begin Spinner.uSpinner SpnVentanilla 
         Height          =   315
         Left            =   5760
         TabIndex        =   10
         Top             =   1800
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Max             =   99
         Min             =   1
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin SICMACT.FlexEdit grdPersInvolucra 
         Height          =   1455
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   10560
         _ExtentX        =   18627
         _ExtentY        =   2566
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Codigo-Usuario-Nombre-Cargo-Relacion con Arqueo-Aux-nIdArqueador-nEstado"
         EncabezadosAnchos=   "300-1400-850-3500-3200-2100-0-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-5-X-X-X"
         ListaControles  =   "0-1-0-0-0-3-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-L-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin VB.Label lblAgeCod 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10320
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblAgeDesc 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         TabIndex        =   12
         Top             =   1800
         Width           =   2835
      End
      Begin VB.Label Label4 
         Caption         =   "Agencia :"
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   11
         Top             =   1845
         Width           =   855
      End
      Begin VB.Label lblVent 
         Caption         =   "Ventanilla :"
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   1845
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Datos del Arqueo "
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar Relaciones"
         Height          =   315
         Left            =   7950
         TabIndex        =   33
         Top             =   260
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.ComboBox cboTipoArqueo 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   260
         Width           =   1695
      End
      Begin VB.TextBox txtOtros 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         TabIndex        =   1
         Top             =   260
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha/Hora Inicio :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblFechaHora 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   255
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Arqueo : "
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   300
         Width           =   1455
      End
   End
   Begin SICMACT.Usuario oUser 
      Left            =   120
      Top             =   360
      _ExtentX        =   820
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmCajaArqueoVentBov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmOpeArqueoVentBov
'** Descripción : Proceso de registro de Arqueos de Ventanillas y Bóvedas creado segun RFC081-2012
'** Creación : JUEZ, 20120807 06:00:00 PM
'********************************************************************

Option Explicit
Dim bResultadoVisto As Boolean
Dim oVisto As frmVistoElectronico
Dim lnTipoArqVentBov As Integer
Dim cIdArqueo As String
Dim lsMovNroIni As String
Dim lsMovNroFin As String
Dim lnMovNro As Long
Dim NumVentanilla As Integer
Dim psUserPersArqueado As String
Dim cUsuVisto As String

' *** RIRO20140710 ERS072
Private sPersCodArqueado As String
Private sPersCodArqueador As String
Private sCargosArqueoAgenciaPrincipal As String
Private sCargosArqueoOtrasAgencias As String
' *** END RIRO

' *** TORE20180712 ERS033
Dim clsMov As COMNContabilidad.NCOMContFunciones
Private sCargoArquedorSR As String
Dim sMovNro As String
Dim nTipoProceso As Integer
' *** END TORE


Public Sub Inicia(ByVal pnTipoArqVentBov As Integer)
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    Dim psMovNro As String
    ' RIRO20140701 ERS072 *****************************
    Dim oUsuariosArea As COMDConstSistema.DCOMGeneral
    Dim ClsPersona As New COMDPersona.DCOMPersonas
    Dim bUsuario As Boolean
    Dim sMensaje As String
    Dim nTipoAge As Integer
    Dim nMostrar As Integer
    Dim rsTmp As New ADODB.Recordset
    ' END RIRO ****************************************
    
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        
    Dim i As Integer, nEstadoArqueador As Integer
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsPrsInvol As New ADODB.Recordset
    
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set oVisto = New frmVistoElectronico
    
    'bResultadoVisto = oVisto.Inicio(6) 'RIRO20140705 ER072 COM
    
    ' RIRO20140705 ERS072 ADD ****************
    If pnTipoArqVentBov <> 3 Then
        bResultadoVisto = oVisto.Inicio(6)
        If Not bResultadoVisto Then
            Exit Sub
        End If
        cUsuVisto = oVisto.ObtieneUsuarioVisto
    End If
    ' END RIRO *******************************
    
    'cUsuVisto = oVisto.ObtieneUsuarioVisto 'RIRO20140705 COM
    'Set rsPrsInvol = clsGen.GetConstante(4051) ' RIRO20140710 COM
    
    cIdArqueo = ""
    lnTipoArqVentBov = pnTipoArqVentBov
    Set oUsuariosArea = New COMDConstSistema.DCOMGeneral
    sCargosArqueoAgenciaPrincipal = oUsuariosArea.LeeConstSistema("493")
    sCargosArqueoOtrasAgencias = oUsuariosArea.LeeConstSistema("494")
    sCargoArquedorSR = oUsuariosArea.LeeConstSistema("496")
    Set oUsuariosArea = Nothing
    
    If lnTipoArqVentBov = 1 Then
        Me.Caption = Me.Caption + "Ventanilla"
        lblVent.Visible = True
        SpnVentanilla.Visible = True
    'RIRO20140701 ERS072 *****************************************************************************
        txtOtros.Visible = True
        Call CargaCombo(cboTipoArqueo, 4050, "4")
    ElseIf lnTipoArqVentBov = 3 Then
        Me.Caption = Me.Caption + "Entre Ventanilla"
        lblVent.Visible = True
        SpnVentanilla.Visible = True
        txtOtros.Visible = False
        bUsuario = False
        cboTipoArqueo.Width = 2500 'RIRO20140710 ERS072
        sPersCodArqueado = ""
        sPersCodArqueador = ""
        If gsCodAge = "01" Then
            nTipoAge = 1
        Else
            nTipoAge = 2
        End If
        Call CargaCombo(cboTipoArqueo, 4050, "", "4")
        
        'Supervisor de Operaciones y RFIII
       If InStr(1, "006005,007026", gsCodCargo) > 0 Then
            Set oUsuariosArea = New COMDConstSistema.DCOMGeneral
            If gsCodCargo = "006005" Then
                cmdGenerar.Visible = True
                Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser2(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADO
            Else
                cmdGenerar.Visible = True
                Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADO
            End If
            If Not rsTmp.EOF And Not rsTmp.BOF Then
                If rsTmp!nEstado = 0 Then     ' Sin Generar Arqueo
                    sMensaje = "Deben generarse las parejas de arqueo para continuar con el proceso" & vbNewLine & "Consultar con el supervisor de operaciones correspondiente"
                    nEstadoArqueador = 0
                ElseIf rsTmp!nEstado = 1 Then ' Generado
                    bUsuario = True
                    nEstadoArqueador = 1
                ElseIf rsTmp!nEstado = 2 Then ' Arqueado
                    sMensaje = "El usuario <<" & UCase(gsCodUser) & ">> ya participó de un arqueo entre ventanillas" 'Comentado by jato 20210105 'JATO 20210105
                    'bUsuario = True 'COMENTADO BY JATO 20210105
                    nEstadoArqueador = 2
                ElseIf rsTmp!nEstado = 3 Then ' Pendiente de Arqueo
                    bUsuario = True
                    'sMensaje = "El usuario <<" & UCase(gsCodUser) & ">> ya participó de un arqueo entre ventanillas" 'Comentado by jato 20210105
                    nEstadoArqueador = 3
                Else
                    bUsuario = True
                    nEstadoArqueador = 1
                    'sMensaje = "No se ha identificado el estado del usuario, consultar con el Area de T.I."
                End If
            Else
                sMensaje = "Usuario no autorizado para usar esta opción"
            End If
            
            If Not bUsuario Then
                If nEstadoArqueador = 0 Then
                    
                    If gsCodCargo = "006005" Or gsCodCargo = "007026" Then
                    Set rsTmp = ClsPersona.BuscaCliente(gsCodPersUser, BusquedaCodigo)
                    If gsCodCargo = "006005" Then
                        cmdGenerar.Visible = True
                    End If
                    If Not rsTmp Is Nothing Then
                        If Not rsTmp.BOF And Not rsTmp.EOF Then
                            grdPersInvolucra.AdicionaFila
                            grdPersInvolucra.TextMatrix(1, 1) = gsCodPersUser
                            grdPersInvolucra.TextMatrix(1, 2) = gsCodUser
                            grdPersInvolucra.TextMatrix(1, 3) = rsTmp!cPersNombre
                            grdPersInvolucra.TextMatrix(1, 4) = "SUPERVISOR DE OPERACIONES"
                            grdPersInvolucra.TextMatrix(1, 5) = "ARQUEADOR/AUDITOR                                                                           1"
                            sPersCodArqueador = gsCodPersUser
                        End If
                    End If
                    'Representante Financiero III
                    If gsCodCargo = "007026" Then
                        Set rsTmp = ValidarRFIII
                        If Not rsTmp.BOF And Not rsTmp.EOF Then
                            If rsTmp!bModoSupervisor = True Then
                                cmdGenerar.Visible = True
                            End If
                        End If
                    End If
                End If
                    
                    
                Else
                    MsgBox sMensaje, vbExclamation, "Aviso"
                    Exit Sub
                End If
                
            Else
                lblVent.Left = 4200
                lblVent.Width = 1500
                lblVent.Caption = "Ventanilla a arquear:"
                cmdNuevo.Enabled = False
                cmdEliminar.Enabled = False
                Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser2(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADOR
                If Not rsTmp.EOF And Not rsTmp.BOF Then
                    
                    ' *** Validando que los dos miembros de la pareja ya esten arqueados
                    If rsTmp!nEstado = 2 And rsTmp!nEstadoArqueador = 2 Then
                        MsgBox "Ya se efectuó el proceso de arqueo entre ventanillas", vbInformation, "Aviso"
                        Set rsTmp = Nothing
                        Exit Sub
                    ElseIf rsTmp!nEstadoArqueador = 1 Then
                        '1ra Fila ***
                        grdPersInvolucra.AdicionaFila
                        grdPersInvolucra.TextMatrix(1, 1) = rsTmp!cPersCodArqueado
                        grdPersInvolucra.TextMatrix(1, 2) = rsTmp!cUserArqueado
                        grdPersInvolucra.TextMatrix(1, 3) = rsTmp!cPersNombreArqueado
                        grdPersInvolucra.TextMatrix(1, 4) = rsTmp!cRHCargoArqueado
                        grdPersInvolucra.TextMatrix(1, 5) = "PERSONAL ARQUEADO                                                                           2"
                        grdPersInvolucra.TextMatrix(1, 7) = rsTmp!nIdArqueado
                        grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstadoArqueador
                        '2da Fila ***
                        grdPersInvolucra.AdicionaFila
                        grdPersInvolucra.TextMatrix(2, 1) = rsTmp!cPersCodArqueador
                        grdPersInvolucra.TextMatrix(2, 2) = rsTmp!cUserArqueador
                        grdPersInvolucra.TextMatrix(2, 3) = rsTmp!cPersNombreArqueador
                        grdPersInvolucra.TextMatrix(2, 4) = rsTmp!cRHCargoArqueador
                        grdPersInvolucra.TextMatrix(2, 5) = "ARQUEADOR/AUDITOR                                                                           1"
                        grdPersInvolucra.TextMatrix(2, 7) = rsTmp!nIdArqueador
                        grdPersInvolucra.TextMatrix(2, 8) = rsTmp!nEstadoArqueador
                        sPersCodArqueador = rsTmp!cPersCodArqueador
                        
                    Else
                        grdPersInvolucra.AdicionaFila
                        grdPersInvolucra.TextMatrix(1, 1) = rsTmp!cPersCodArqueado
                        grdPersInvolucra.TextMatrix(1, 2) = rsTmp!cUserArqueado
                        grdPersInvolucra.TextMatrix(1, 3) = rsTmp!cPersNombreArqueado
                        grdPersInvolucra.TextMatrix(1, 4) = rsTmp!cRHCargoArqueado
                        grdPersInvolucra.TextMatrix(1, 5) = "PERSONAL ARQUEADO                                                                           2"
                        grdPersInvolucra.TextMatrix(1, 7) = rsTmp!nIdArqueado
                        grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstado
                        
                        cmdNuevo.Enabled = True
                        cmdEliminar.Enabled = True
                    End If
                    
                Else
                    If nEstadoArqueador = 2 Then
                        MsgBox "El usuario " & UCase(gsCodUser) & " ya fue arqueado", vbExclamation, "Aviso"
                    Else
                        MsgBox "El usuario " & UCase(gsCodUser) & " será arqueado por el Supervisor de Operaciones o por el RFIII", vbExclamation, "Aviso"
                        Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADOR
                        If Not rsTmp.EOF And Not rsTmp.BOF Then
                       
                            '1ra Fila ***
                            grdPersInvolucra.AdicionaFila
                            grdPersInvolucra.TextMatrix(1, 1) = rsTmp!cPersCodArqueado
                            grdPersInvolucra.TextMatrix(1, 2) = rsTmp!cUserArqueado
                            grdPersInvolucra.TextMatrix(1, 3) = rsTmp!cPersNombreArqueado
                            grdPersInvolucra.TextMatrix(1, 4) = rsTmp!cRHCargoArqueado
                            grdPersInvolucra.TextMatrix(1, 5) = "PERSONAL ARQUEADO                                                                           2"
                            grdPersInvolucra.TextMatrix(1, 7) = rsTmp!nIdArqueado
                            grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstadoArqueador
                            
                            
                            '2da Fila ***
                            grdPersInvolucra.AdicionaFila
                            grdPersInvolucra.TextMatrix(2, 1) = rsTmp!cPersCodArqueador
                            grdPersInvolucra.TextMatrix(2, 2) = rsTmp!cUserArqueador
                            grdPersInvolucra.TextMatrix(2, 3) = rsTmp!cPersNombreArqueador
                            grdPersInvolucra.TextMatrix(2, 4) = rsTmp!cRHCargoArqueador
                            grdPersInvolucra.TextMatrix(2, 5) = "ARQUEADOR/AUDITOR                                                                           1"
                            grdPersInvolucra.TextMatrix(2, 7) = rsTmp!nIdArqueador
                            grdPersInvolucra.TextMatrix(2, 8) = rsTmp!nEstadoArqueador
                            sPersCodArqueador = rsTmp!cPersCodArqueador

                        End If
                
                    End If
                    
                    Set oUsuariosArea = Nothing
                    Set rsTmp = Nothing
                    'Exit Sub
                    
                End If
            End If
          
        Else ' RF, Tasadores y Asesores en caso de agencias que no sean la principal.
            Set oUsuariosArea = New COMDConstSistema.DCOMGeneral
            Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADO
            If Not rsTmp.EOF And Not rsTmp.BOF Then
                If rsTmp!nEstado = 0 Then     ' Sin Generar Arqueo
                    sMensaje = "Deben generarse las parejas de arqueo para continuar con el proceso" & vbNewLine & "Consultar con el supervisor de operaciones correspondiente"
                    nEstadoArqueador = 0
                ElseIf rsTmp!nEstado = 1 Then ' Generado
                    bUsuario = True
                    nEstadoArqueador = 1
                ElseIf rsTmp!nEstado = 2 Then ' Arqueado
                    sMensaje = "El usuario <<" & UCase(gsCodUser) & ">> ya participó de un arqueo entre ventanillas"
                    'bUsuario = True ' comentado JATO 20210105
                ElseIf rsTmp!nEstado = 3 Then ' Pendiente de Arqueo
                    bUsuario = True 'add by jato 20210105
                    'sMensaje = "El usuario <<" & UCase(gsCodUser) & ">> ya participó de un arqueo entre ventanillas" ' comentado JATO 20210105
                    nEstadoArqueador = 3
                Else
                    bUsuario = True
                    nEstadoArqueador = 1
                    'sMensaje = "No se ha identificado el estado del usuario, consultar con el Area de T.I."
                End If
            Else
                sMensaje = "Usuario no autorizado para usar esta opción"
            End If
            
            If Not bUsuario Then
                MsgBox sMensaje, vbExclamation, "Aviso"
                Exit Sub
            Else
                lblVent.Left = 4200
                lblVent.Width = 1500
                lblVent.Caption = "Ventanilla a arquear:"
                Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser2(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADOR
                If Not rsTmp.EOF And Not rsTmp.BOF Then
                    
                    ' *** Validando que los dos miembros de la pareja ya esten arqueados
                    If rsTmp!nEstado = 2 And rsTmp!nEstadoArqueador = 2 Then
                        MsgBox "Ya se efectuó el proceso de arqueo entre ventanillas", vbInformation, "Aviso"
                        Set rsTmp = Nothing
                        Exit Sub
                    End If
                    ' *** end validacion
                    
                    grdPersInvolucra.AdicionaFila
                    grdPersInvolucra.TextMatrix(1, 1) = rsTmp!cPersCodArqueado
                    grdPersInvolucra.TextMatrix(1, 2) = rsTmp!cUserArqueado
                    grdPersInvolucra.TextMatrix(1, 3) = rsTmp!cPersNombreArqueado
                    grdPersInvolucra.TextMatrix(1, 4) = rsTmp!cRHCargoArqueado
                    grdPersInvolucra.TextMatrix(1, 5) = "PERSONAL ARQUEADO                                                                           2"
                    grdPersInvolucra.TextMatrix(1, 7) = rsTmp!nIdArqueado
                    grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstado
                    sPersCodArqueado = rsTmp!cPersCodArqueado
                    
                    '2da Fila ***
                    If rsTmp!cPersCodArqueador <> "" Then
                        grdPersInvolucra.AdicionaFila
                        grdPersInvolucra.TextMatrix(2, 1) = rsTmp!cPersCodArqueador
                        grdPersInvolucra.TextMatrix(2, 2) = rsTmp!cUserArqueador
                        grdPersInvolucra.TextMatrix(2, 3) = rsTmp!cPersNombreArqueador
                        grdPersInvolucra.TextMatrix(2, 4) = rsTmp!cRHCargoArqueador
                        grdPersInvolucra.TextMatrix(2, 5) = "ARQUEADOR/AUDITOR                                                                           1"
                        grdPersInvolucra.TextMatrix(2, 7) = rsTmp!nIdArqueador
                        grdPersInvolucra.TextMatrix(2, 8) = rsTmp!nEstadoArqueador
                        sPersCodArqueador = rsTmp!cPersCodArqueador
                    End If
                    
'                    grdPersInvolucra.AdicionaFila
'                    grdPersInvolucra.TextMatrix(2, 1) = rsTmp!cPersCodArqueador
'                    grdPersInvolucra.TextMatrix(2, 2) = rsTmp!cUserArqueador
'                    grdPersInvolucra.TextMatrix(2, 3) = rsTmp!cPersNombreArqueador
'                    grdPersInvolucra.TextMatrix(2, 4) = rsTmp!cRHCargoArqueador
'                    grdPersInvolucra.TextMatrix(2, 5) = "ARQUEADOR/AUDITOR                                                                           1"
'                    grdPersInvolucra.TextMatrix(2, 7) = rsTmp!nIdArqueador
'                    grdPersInvolucra.TextMatrix(2, 8) = rsTmp!nEstadoArqueador
'                    sPersCodArqueador = rsTmp!cPersCodArqueador
                    
                Else
                    If nEstadoArqueador = 2 Then
                        MsgBox "El usuario " & UCase(gsCodUser) & " ya fue arqueado", vbExclamation, "Aviso"
                    Else
                        MsgBox "El usuario " & UCase(gsCodUser) & " será arqueado por el Supervisor de Operaciones o por el RFIII", vbExclamation, "Aviso"
                        Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADOR
                        If Not rsTmp.EOF And Not rsTmp.BOF Then
                            ' *** Validando que los dos miembros de la pareja ya esten arqueados
                            ''If rsTmp!nEstado = 2 And rsTmp!nEstadoArqueador = 2 Then
                              ''  MsgBox "Ya se efectuó el proceso de arqueo entre ventanillas", vbInformation, "Aviso"
                              ''  Set rsTmp = Nothing
                              ''  Exit Sub
                            ''End If
                            ' *** end validacion
                            
                            '1ra Fila ***
                            grdPersInvolucra.AdicionaFila
                            grdPersInvolucra.TextMatrix(1, 1) = rsTmp!cPersCodArqueado
                            grdPersInvolucra.TextMatrix(1, 2) = rsTmp!cUserArqueado
                            grdPersInvolucra.TextMatrix(1, 3) = rsTmp!cPersNombreArqueado
                            grdPersInvolucra.TextMatrix(1, 4) = rsTmp!cRHCargoArqueado
                            grdPersInvolucra.TextMatrix(1, 5) = "PERSONAL ARQUEADO                                                                           2"
                            grdPersInvolucra.TextMatrix(1, 7) = rsTmp!nIdArqueado
                            grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstadoArqueador
                            'sPersCodArqueador = rsTmp!cPersCodArqueado
                        End If
                
                    End If
                    
                    Set oUsuariosArea = Nothing
                    Set rsTmp = Nothing
                    'Exit Sub
                    
                End If
            End If
          
        End If
       
    'END RIRO ****************************************************************************************
    Else
        Me.Caption = Me.Caption + "Bóveda"
        lblVent.Visible = False
        SpnVentanilla.Visible = False
        Call CargaCombo(cboTipoArqueo, 4050, "'4'") ' RIRO20140705 ERS072 ADD
        txtOtros.Visible = True
    End If
    'Set clsGen = Nothing 'RIRO20140710 ERS072
    
    lsMovNroIni = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Me.lblFechaHora = fgFechaHoraGrab(lsMovNroIni)
    'Call CargaCombo(cboTipoArqueo, 4050) 'RIRO20140706 ERS072 Comentado
    
    lblAgeDesc.Caption = UCase(gsNomAge)
    lblAgeCod.Caption = gsCodAge
    
    CargaBilletajes gMonedaNacional
    CargaBilletajes gMonedaExtranjera
    
    If lnTipoArqVentBov <> 3 Then
        Set rsPrsInvol = clsGen.GetConstante(4051)
    Else
        ' Si es Supervisor o RFIII
        If InStr(1, "006005,007026", gsCodCargo) > 0 Then
            Set rsPrsInvol = clsGen.GetConstante(4051, , "'[23]'")
        ' Si es RF, Tasador o Asesor
        Else
            Set rsPrsInvol = clsGen.GetConstante(4051, , "'3'")
        End If
        'grdPersInvolucra_RowColChange
    End If
    Set clsGen = Nothing
    grdPersInvolucra.CargaCombo rsPrsInvol
    
    psUserPersArqueado = ""
    
    Me.Show 1
End Sub

'Public Sub inicia(ByVal pnTipoArqVentBov As Integer)
'    Dim loContFunct As COMNContabilidad.NCOMContFunciones
'    Dim psMovNro As String
'    ' RIRO20140701 ERS072 *****************************
'    Dim oUsuariosArea As COMDConstSistema.DCOMGeneral
'    Dim ClsPersona As New COMDPersona.DCOMPersonas
'    Dim bUsuario As Boolean
'    Dim sMensaje As String
'    Dim nTipoAge As Integer
'    Dim rsTmp As New ADODB.Recordset
'    ' END RIRO ****************************************
'
'    Set loContFunct = New COMNContabilidad.NCOMContFunciones
'
'    Dim i As Integer, nEstadoArqueador As Integer
'    Dim clsGen As COMDConstSistema.DCOMGeneral
'    Dim rsPrsInvol As New ADODB.Recordset
'
'    Set clsGen = New COMDConstSistema.DCOMGeneral
'    Set oVisto = New frmVistoElectronico
'
'    'bResultadoVisto = oVisto.Inicio(6) 'RIRO20140705 ER072 COM
'
'    ' RIRO20140705 ERS072 ADD ****************
'    If pnTipoArqVentBov <> 3 Then
'        bResultadoVisto = oVisto.Inicio(6)
'        If Not bResultadoVisto Then
'            Exit Sub
'        End If
'        cUsuVisto = oVisto.ObtieneUsuarioVisto
'    End If
'    ' END RIRO *******************************
'
'    'cUsuVisto = oVisto.ObtieneUsuarioVisto 'RIRO20140705 COM
'    'Set rsPrsInvol = clsGen.GetConstante(4051) ' RIRO20140710 COM
'
'    cIdArqueo = ""
'    lnTipoArqVentBov = pnTipoArqVentBov
'    Set oUsuariosArea = New COMDConstSistema.DCOMGeneral
'    sCargosArqueoAgenciaPrincipal = oUsuariosArea.LeeConstSistema("493")
'    sCargosArqueoOtrasAgencias = oUsuariosArea.LeeConstSistema("494")
'    Set oUsuariosArea = Nothing
'
'    If lnTipoArqVentBov = 1 Then
'        Me.Caption = Me.Caption + "Ventanilla"
'        lblVent.Visible = True
'        SpnVentanilla.Visible = True
'    'RIRO20140701 ERS072 *****************************************************************************
'        txtOtros.Visible = True
'        Call CargaCombo(cboTipoArqueo, 4050, "4")
'    ElseIf lnTipoArqVentBov = 3 Then
'        Me.Caption = Me.Caption + "Entre Ventanilla"
'        lblVent.Visible = True
'        SpnVentanilla.Visible = True
'        txtOtros.Visible = False
'        bUsuario = False
'        cboTipoArqueo.Width = 2500 'RIRO20140710 ERS072
'        sPersCodArqueado = ""
'        sPersCodArqueador = ""
'        If gsCodAge = "01" Then
'            nTipoAge = 1
'        Else
'            nTipoAge = 2
'        End If
'        Call CargaCombo(cboTipoArqueo, 4050, "", "4")
'
'        'Supervisor de Operaciones y RFIII
'        If InStr(1, "006005,007026", gsCodCargo) > 0 Then
'
'            Set rsTmp = ClsPersona.BuscaCliente(gsCodPersUser, BusquedaCodigo)
'            If gsCodCargo = "006005" Then cmdGenerar.Visible = True
'            If Not rsTmp Is Nothing Then
'                If Not rsTmp.BOF And Not rsTmp.EOF Then
'                    grdPersInvolucra.AdicionaFila
'                    grdPersInvolucra.TextMatrix(1, 1) = gsCodPersUser
'                    grdPersInvolucra.TextMatrix(1, 2) = gsCodUser
'                    grdPersInvolucra.TextMatrix(1, 3) = rsTmp!cPersNombre
'                    If gsCodCargo = "006005" Then     ' Supervisor
'                        grdPersInvolucra.TextMatrix(1, 4) = "SUPERVISOR DE OPERACIONES"
'                    ElseIf gsCodCargo = "007026" Then ' RFIII
'                        grdPersInvolucra.TextMatrix(1, 4) = "REPRESENTANTE FINANCIERO III"
'                    End If
'                    grdPersInvolucra.TextMatrix(1, 5) = "ARQUEADOR/AUDITOR                                                                           1"
'                    sPersCodArqueador = gsCodPersUser
'                End If
'            End If
'            'Representante Financiero III
'            If gsCodCargo = "007026" Then
'                Set rsTmp = ValidarRFIII
'                If Not rsTmp.BOF And Not rsTmp.EOF Then
'                    If rsTmp!bModoSupervisor = True Then
'                        cmdGenerar.Visible = True
'                    End If
'                End If
'            End If
'            bUsuario = True
'
'        Else ' RF, Tasadores y Asesores en caso de agencias que no sean la principal.
'
'            Set oUsuariosArea = New COMDConstSistema.DCOMGeneral
'            Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADO
'            If Not rsTmp.EOF And Not rsTmp.BOF Then
'                If rsTmp!nEstado = 0 Then     ' Sin Generar Arqueo
'                    sMensaje = "Deben generarse las parejas de arqueo para continuar con el proceso" & vbNewLine & "Consultar con el supervisor de operaciones correspondiente"
'                    nEstadoArqueador = 0
'                ElseIf rsTmp!nEstado = 1 Then ' Generado
'                    bUsuario = True
'                    nEstadoArqueador = 1
'                ElseIf rsTmp!nEstado = 2 Then ' Arqueado
'                    bUsuario = True
'                    nEstadoArqueador = 2
'                ElseIf rsTmp!nEstado = 3 Then ' Pendiente de Arqueo
'                    sMensaje = "El usuario <<" & UCase(gsCodUser) & ">> ya participó de un arqueo entre ventanillas"
'                    nEstadoArqueador = 3
'                Else
'                    sMensaje = "No se ha identificado el estado del usuario, consultar con el Area de T.I."
'                End If
'            Else
'                sMensaje = "Usuario no autorizado para usar esta opción"
'            End If
'            If Not bUsuario Then
'                MsgBox sMensaje, vbExclamation, "Aviso"
'                Exit Sub
'            Else
'                lblVent.Left = 4200
'                lblVent.Width = 1500
'                lblVent.Caption = "Ventanilla a arquear:"
'                Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 2) ' BUSCAR POR USUARIO ARQUEADOR
'                If Not rsTmp.EOF And Not rsTmp.BOF Then
'
'                    ' *** Validando que los dos miembros de la pareja ya esten arqueados
'                    If rsTmp!nEstado = 2 And rsTmp!nEstadoArqueador = 2 Then
'                        MsgBox "Ya se efectuó el proceso de arqueo entre ventanillas", vbInformation, "Aviso"
'                        Set rsTmp = Nothing
'                        Exit Sub
'                    End If
'                    ' *** end validacion
'
'                    '1ra Fila ***
'                    grdPersInvolucra.AdicionaFila
'                    grdPersInvolucra.TextMatrix(1, 1) = rsTmp!cPersCodArqueador
'                    grdPersInvolucra.TextMatrix(1, 2) = rsTmp!cUserArqueador
'                    grdPersInvolucra.TextMatrix(1, 3) = rsTmp!cPersNombreArqueador
'                    grdPersInvolucra.TextMatrix(1, 4) = rsTmp!cRHCargoArqueador
'                    grdPersInvolucra.TextMatrix(1, 5) = "ARQUEADOR/AUDITOR                                                                           1"
'                    grdPersInvolucra.TextMatrix(1, 7) = rsTmp!nIdArqueador
'                    grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstadoArqueador
'                    sPersCodArqueador = rsTmp!cPersCodArqueador
'
'                    '2da Fila ***
'                    grdPersInvolucra.AdicionaFila
'                    grdPersInvolucra.TextMatrix(2, 1) = rsTmp!cPersCodArqueado
'                    grdPersInvolucra.TextMatrix(2, 2) = rsTmp!cUserArqueado
'                    grdPersInvolucra.TextMatrix(2, 3) = rsTmp!cPersNombreArqueado
'                    grdPersInvolucra.TextMatrix(2, 4) = rsTmp!cRHCargoArqueado
'                    grdPersInvolucra.TextMatrix(2, 5) = "PERSONAL ARQUEADO                                                                           2"
'                    grdPersInvolucra.TextMatrix(2, 7) = rsTmp!nIdArqueado
'                    grdPersInvolucra.TextMatrix(2, 8) = rsTmp!nEstado
'                    sPersCodArqueado = rsTmp!cPersCodArqueado
'
'                Else
'                    If nEstadoArqueador = 2 Then
'                        MsgBox "El usuario " & UCase(gsCodUser) & " ya fue arqueado", vbExclamation, "Aviso"
'                    Else
'                        MsgBox "El usuario " & UCase(gsCodUser) & " será arqueado por el Supervisor de Operaciones o por el RFIII", vbExclamation, "Aviso"
'                    End If
'
'                    Set oUsuariosArea = Nothing
'                    Set rsTmp = Nothing
'                    Exit Sub
'
'                End If
'            End If
'        End If
'
'    'END RIRO ****************************************************************************************
'    Else
'        Me.Caption = Me.Caption + "Bóveda"
'        lblVent.Visible = False
'        SpnVentanilla.Visible = False
'        Call CargaCombo(cboTipoArqueo, 4050, "'4'") ' RIRO20140705 ERS072 ADD
'        txtOtros.Visible = True
'    End If
'    'Set clsGen = Nothing 'RIRO20140710 ERS072
'
'    lsMovNroIni = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'    Me.lblFechaHora = fgFechaHoraGrab(lsMovNroIni)
'    'Call CargaCombo(cboTipoArqueo, 4050) 'RIRO20140706 ERS072 Comentado
'
'    lblAgeDesc.Caption = UCase(gsNomAge)
'    lblAgeCod.Caption = gsCodAge
'
'    CargaBilletajes gMonedaNacional
'    CargaBilletajes gMonedaExtranjera
'
'    If lnTipoArqVentBov <> 3 Then
'        Set rsPrsInvol = clsGen.GetConstante(4051)
'    Else
'        ' Si es Supervisor o RFIII
'        If InStr(1, "006005,007026", gsCodCargo) > 0 Then
'            Set rsPrsInvol = clsGen.GetConstante(4051, , "'[23]'")
'        ' Si es RF, Tasador o Asesor
'        Else
'            Set rsPrsInvol = clsGen.GetConstante(4051, , "'3'")
'        End If
'    End If
'    Set clsGen = Nothing
'    grdPersInvolucra.CargaCombo rsPrsInvol
'
'    psUserPersArqueado = ""
'
'    Me.Show 1
'End Sub

'RIRO20140709 ERS072 ADD "psFiltro"
Public Sub CargaCombo(ByVal CtrlCombo As ComboBox, ByVal psConst As String, Optional ByVal psFiltro As String = "", Optional ByVal psFiltro2 As String = "")
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(psConst, psFiltro, psFiltro2)
    Set clsGen = Nothing
    
    CtrlCombo.Clear
    While Not rsConst.EOF
        CtrlCombo.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
End Sub

Private Sub cboTipoArqueo_Click()
    If Trim(Right(cboTipoArqueo, 2)) = "3" Then
        txtOtros.Enabled = True
        txtOtros.BackColor = &H80000005
        txtOtros.SetFocus
    Else
        txtOtros.Enabled = False
        txtOtros.Text = ""
        txtOtros.BackColor = &H8000000F
        'cmdNuevo.SetFocus
    End If
End Sub

Private Sub cmdEliminar_Click()
    
    Dim nRow As Integer
    Dim bResult As Boolean
    
    nRow = grdPersInvolucra.row
        
    If grdPersInvolucra.TextMatrix(nRow, 1) = sPersCodArqueado And sPersCodArqueado <> "" Then
        bResult = False
    ElseIf grdPersInvolucra.TextMatrix(nRow, 1) = sPersCodArqueador And sPersCodArqueador <> "" Then
        bResult = False
    Else
        bResult = True
    End If
'    If Not bResult Then
'        MsgBox "No es posible retirar el registro seleccionado", vbInformation, "Aviso"
'        Exit Sub
'    End If
        
    If MsgBox("¿Está seguro de eliminar a la persona?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdPersInvolucra.EliminaFila grdPersInvolucra.row
    End If
End Sub

Private Sub cmdGenerar_Click()
Dim oRelacion As New frmGeneraRelacionArqueo
oRelacion.Show 1
End Sub

Private Sub cmdNuevo_Click()
    'Comentado - TORE-ERS033-2018
    grdPersInvolucra.AdicionaFila
    grdPersInvolucra.SetFocus
    ' *** TORE ERS033-2018 ***
    
'    If grdPersInvolucra.Rows = 2 And gsCodCargo <> "007026" Then
'        MsgBox "Ud. ", vbInformation, "Aviso"
'        Exit Sub
'    Else
'        grdPersInvolucra.AdicionaFila
'        grdPersInvolucra.SetFocus
'    End If
    ' *** END TORE ***
End Sub

'Private Sub cmdNuevo_Click()
'    grdPersInvolucra.AdicionaFila
'    grdPersInvolucra.SetFocus
'End Sub

Private Sub CargaBilletajes(ByVal nmoneda As Moneda)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim oContFunct As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim oEfec As COMDCajaGeneral.DCOMEfectivo   'Defectivo
Dim lnFila As Long
Dim i As Integer

Set oContFunct = New COMNContabilidad.NCOMContFunciones
Set oEfec = New COMDCajaGeneral.DCOMEfectivo

Set rs = New ADODB.Recordset
'If lbRegistro Then
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    'If sMovNro <> "" Then
    '    Set rs = oCajero.GetBilletajeCajero(sMovNro, txtBuscarUser, nmoneda, "B")
    'Else
        Set rs = oEfec.EmiteBilletajes(nmoneda, "B")
    'End If
     Set oCajero = Nothing
'Else
'    If lbMuestra = False Then
'        Set rs = oEfec.EmiteBilletajes(nmoneda, "B")
'    Else
'        Set rs = oEfec.GetBilletajesMov(sMovNro, nmoneda, "B")
'    End If
'End If

i = IIf(nmoneda = gMonedaNacional, 0, 1)

fgBilletes(i).FontFixed.Bold = True
fgBilletes(i).Clear
fgBilletes(i).FormaCabecera
fgBilletes(i).Rows = 2
Do While Not rs.EOF
    fgBilletes(i).AdicionaFila
    lnFila = fgBilletes(i).row
    fgBilletes(i).TextMatrix(lnFila, 1) = Replace(rs!Descripcion, "BILLETAJE", "")
    
    'Modificado por gitu 23-10-2009
    'If bLimBille Then
        fgBilletes(i).TextMatrix(lnFila, 2) = Format(0, "#,##0")
        fgBilletes(i).TextMatrix(lnFila, 3) = Format(0, "#,##0.00")
    'Else
    '    fgBilletes(i).TextMatrix(lnFila, 2) = Format(rs!Cantidad, "#,##0")
    '    fgBilletes(i).TextMatrix(lnFila, 3) = Format(rs!Monto, "#,##0.00")
    'End If
    'End Gitu
    fgBilletes(i).TextMatrix(lnFila, 4) = rs!cEfectivoCod
    fgBilletes(i).TextMatrix(lnFila, 5) = rs!nEfectivoValor
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
fgBilletes(i).Col = 2

Set rs = New ADODB.Recordset
Set oCajero = New COMNCajaGeneral.NCOMCajero
'If lbRegistro Then
'    If sMovNro <> "" Then
'        Set rs = oCajero.GetBilletajeCajero(sMovNro, txtBuscarUser, nmoneda, "M")
'    Else
        Set rs = oEfec.EmiteBilletajes(nmoneda, "M")
'    End If
'Else
'    If lbMuestra = False Then
'        Set rs = oEfec.EmiteBilletajes(nmoneda, "M")
'    Else
'        Set rs = oEfec.GetBilletajesMov(sMovNro, nmoneda, "M")
'    End If
'End If

If i <> 1 And nmoneda <> gMonedaExtranjera Then
    fgMonedas(i).FontFixed.Bold = True
    fgMonedas(i).Clear
    fgMonedas(i).FormaCabecera
    fgMonedas(i).Rows = 2
    Do While Not rs.EOF
        fgMonedas(i).AdicionaFila
        lnFila = fgMonedas(i).row
        fgMonedas(i).TextMatrix(lnFila, 1) = Replace(rs!Descripcion, "MONEDA   ", "")
        
        'comentado por gitu 23-10-2009
        'If bLimBille Then
            fgMonedas(i).TextMatrix(lnFila, 2) = Format(0, "#,##0")
            fgMonedas(i).TextMatrix(lnFila, 3) = Format(0, "#,##0.00")
        'Else
        '    fgMonedas(i).TextMatrix(lnFila, 2) = Format(rs!Cantidad, "#,##0")
        '    fgMonedas(i).TextMatrix(lnFila, 3) = Format(rs!Monto, "#,##0.00")
        'End If
        
        fgMonedas(i).TextMatrix(lnFila, 4) = rs!cEfectivoCod
        fgMonedas(i).TextMatrix(lnFila, 5) = rs!nEfectivoValor
        rs.MoveNext
    Loop
End If

    rs.Close
    Set rs = Nothing
    Set oContFunct = Nothing
    Set oEfec = Nothing
    
    lblTotalBilletes(i) = Format(fgBilletes(i).SumaRow(3), "#,##0.00")
If i <> 1 And nmoneda <> gMonedaExtranjera Then
    fgMonedas(i).Col = 2
    lblTotMoneda(i) = Format(fgMonedas(i).SumaRow(3), "#,##0.00")
    'lblTotal(i) = Format(CDbl(lblTotalBilletes(i)) + CDbl(lblTotMoneda(i)), "#,##0.00")
End If
End Sub

Private Sub cmdProcesar_Click()
    Dim oCajaGen As COMNCajaGeneral.NCOMCajaGeneral
    Dim oContFunct As COMNContabilidad.NCOMContFunciones
    Dim oCajero As COMNCajaGeneral.NCOMCajero
    Dim oEfectivo As COMDCajaGeneral.DCOMEfectivo

    Set oCajaGen = New COMNCajaGeneral.NCOMCajaGeneral
    Set oContFunct = New COMNContabilidad.NCOMContFunciones
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    Set oEfectivo = New COMDCajaGeneral.DCOMEfectivo
    
    Dim i As Integer
    Dim ValidarUsrVisto As Boolean
    Dim rs As ADODB.Recordset, rsBillMN As ADODB.Recordset
    Dim rsMonMN As ADODB.Recordset, rsBillME As ADODB.Recordset
    Dim nResultSolSist As Double, nResultSolBill As Double
    Dim nResultDolSist As Double, nResultDolBill As Double
    Dim nSaldoEfectivoAyer As Double, nSaldoMovimientos As Double
    Dim nIngresos As Double, nEgresos As Double
    Dim nHabilitacion As Double, nDevolucion As Double, nDevolucionBilletaje As Double
    Dim bConforme As Integer
    Dim PersRel() As String
    Dim MatDatos() As String
    
    'RIRO20140710 ERS072 *******************************
    Dim bArqueoSup As Boolean, bEsPar As Boolean
    Dim oUsuariosArea As COMDConstSistema.DCOMGeneral
    Dim rsTmp As Recordset
    'END RIRO ******************************************
    
    Dim oTipoCamb As COMDConstSistema.NCOMTipoCambio
    Dim nTiCamb As Double
    Set oTipoCamb = New COMDConstSistema.NCOMTipoCambio
    nTiCamb = oTipoCamb.EmiteTipoCambio(gdFecSis, TCFijoDia)
    
    If Trim(Right(cboTipoArqueo.Text, 2)) = "" Then
        MsgBox "Seleccione el tipo de arqueo", vbInformation, "Aviso"
        Exit Sub
    End If
    
    bArqueoSup = False
    bEsPar = False
   If lnTipoArqVentBov = 1 Then
        If Trim(Right(cboTipoArqueo.Text, 2)) = "2" Then
            If oCajaGen.VerificaSiRealizoArqueoVentBov(gdFecSis, Trim(lblAgeCod.Caption), CInt(SpnVentanilla.valor)) Then
                MsgBox "Ya se realizó el arqueo a la ventanilla " + CStr(SpnVentanilla.valor) + ". No se puede realizar más 1 arqueo por dia y por ventanilla", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    ' Else 'RIRO20140710 ERS072 COM
    ElseIf lnTipoArqVentBov = 2 Then ' RIRO20140710 ERS072 ADD
        If oCajaGen.VerificaSiRealizoArqueoVentBov(gdFecSis, Trim(lblAgeCod.Caption), 0) Then
            MsgBox "Ya se realizó el arqueo a la Bóveda de la " + lblAgeDesc.Caption + ". No se puede realizar más 1 arqueo por dia", vbInformation, "Aviso"
            Exit Sub
        End If
    ' RIRO20140710 ERS072 ADD ******
    ElseIf lnTipoArqVentBov = 3 Then
        Dim nEstado As Integer
        For i = 1 To grdPersInvolucra.Rows - 1
            nEstado = Val(grdPersInvolucra.TextMatrix(i, 8))
            
            'Si la persona involucrada esta arqueada
            If Val(Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 5))) = 2 And nEstado = 2 Then
                MsgBox "El usuario <<" & grdPersInvolucra.TextMatrix(i, 2) & ">> ya fue arqueado", vbExclamation, "Aviso"
                Exit Sub
            End If
            'Si el arqueador es supervisor o rf3, considerar las siguientes validaciones
            If Val(Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 5))) = 2 And InStr(1, "006005,007026", gsCodCargo) > 0 Then
                If nEstado = 0 Or nEstado = 6 Then ' Sin generar, Excluido
                    bArqueoSup = True
                ElseIf nEstado = 1 Then ' Generado
                    'Verificando que el usuario no tenga arqueador, por lo tanto si debe ser arqueado por el supervisor de operaciones o rf3
                    bArqueoSup = True
                    Set oUsuariosArea = New COMDConstSistema.DCOMGeneral
                    Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(grdPersInvolucra.TextMatrix(i, 2), gdFecSis, gsCodArea, gsCodAge, IIf(gsCodAge = "01", 1, 2), 1)   ' BUSCAR POR USUARIO ARQUEADO
                    If Not rsTmp Is Nothing Then
                        If Not rsTmp.BOF And Not rsTmp.EOF Then
                            If Len(Trim(rsTmp!cUserArqueador)) <> 0 Then
                                bEsPar = True
                            End If
                        End If
                    End If
                    If bEsPar Then
                        If MsgBox("El usuario <<" & grdPersInvolucra.TextMatrix(i, 2) & ">> tiene actualmente una " & vbNewLine & "pareja de arqueo, si continua dicha pareja se eliminará" & _
                                   vbNewLine & vbNewLine & "¿Desea continuar?", vbExclamation + vbYesNo, "Aviso") = vbNo Then
                            Exit Sub
                        End If
                    End If
                ElseIf nEstado = 2 Then ' Arqueado
                    MsgBox "El usuario <<" & grdPersInvolucra.TextMatrix(i, 2) & ">> ya fue arqueado", vbExclamation, "Aviso"
                    Exit Sub
                ElseIf nEstado = 3 Then ' Pendiente de Arqueo
                    MsgBox "El usuario <<" & grdPersInvolucra.TextMatrix(i, 2) & ">> Debe ser arqueado por su respectiva pareja de arqueo", vbExclamation, "Aviso"
                    Exit Sub
                End If
            End If
        Next i
    ' END RIRO *********************
    End If

    If ValidaDatos Then
        For i = 1 To grdPersInvolucra.Rows - 1
            If Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 2)) = 1 Then
                If Trim(grdPersInvolucra.TextMatrix(i, 2)) = cUsuVisto Then ' Validamos que el Arqueador/Auditor sea el que emita el visto
                    ValidarUsrVisto = True
                    Exit For
                End If
            End If
        Next i
        
        NumVentanilla = 0
        If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710 ERS072 ADD "Or lnTipoArqVentBov = 3"
            NumVentanilla = CInt(SpnVentanilla.valor)
        End If
        
        'If ValidarUsrVisto = False Then 'RIRO20140705 ERS072 COMENTADO
        If (ValidarUsrVisto = False) And (lnTipoArqVentBov <> 3) Then 'RIRO20140705 ERS072 COMENTADO
            MsgBox "Una de las personas con relacion Arqueador/Auditor debe ser el que emita el Visto Electrónico", vbInformation, "Aviso"
            cmdNuevo.SetFocus
            Exit Sub
        End If
        
        If MsgBox("Se va a procesar el arqueo, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        cIdArqueo = Format(gdFecSis, "yyyymmdd") + Format(Time, "hhmmss") + gsCodAge + gsCodUser
        
        ReDim PersRel(grdPersInvolucra.Rows - 1, 5) 'RIRO20140710 ERS072 Se cambio Nro de columnas 3 por 5
        For i = 1 To grdPersInvolucra.Rows - 1
            PersRel(i - 1, 0) = grdPersInvolucra.TextMatrix(i, 2) 'Usuario
            PersRel(i - 1, 1) = Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 2)) 'Relacion
            PersRel(i - 1, 2) = grdPersInvolucra.TextMatrix(i, 3) 'Nombre
            ' *** RIRO20140710 ERS072
            PersRel(i - 1, 3) = grdPersInvolucra.TextMatrix(i, 7) 'Id Pareja Vetanilla
            PersRel(i - 1, 4) = IIf(IsNumeric(grdPersInvolucra.TextMatrix(i, 8)), grdPersInvolucra.TextMatrix(i, 8), 9) 'nEstado de Pareja de Arqueo.
            ' *** END RIRO
        Next i
        
        Set rsBillMN = fgBilletes(0).GetRsNew()
        Set rsMonMN = fgMonedas(0).GetRsNew()
        Set rsBillME = fgBilletes(1).GetRsNew()
        
        Dim nEfec As Double
        Dim J As Integer
        
        If lnTipoArqVentBov = 2 Then
            psUserPersArqueado = "BOVE"
        End If
        J = 1
        i = 1
        'RIRO20140710 ERS072 ADD "lnTipoArqVentBov = 3"
        For i = 1 To 2
            Set rs = oCajero.GetOpeIngEgreCaptaciones(gGruposIngEgreAgeLocal, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            Set rs = oCajero.GetOpeIngEgreColocaciones(gGruposIngEgreAgeLocal, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            Set rs = oCajero.GetOpeIngEgreCaptaciones(gGruposIngEgreOtraAgencia, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            Set rs = oCajero.GetOpeIngEgreColocaciones(gGruposIngEgreOtraAgencia, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser, gsCodCMAC)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            Set rs = oCajero.GetOpeIngEgreOpeCMACs(gGruposIngEgreOtraCMAC, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            Set rs = oCajero.GetOpeIngEgreOtrasOpe(gGruposIngEgreOtrasOpe, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            Set rs = oCajero.GetOpeIngEgreServicios(gGruposIngEgreServicios, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            Set rs = oCajero.GetOpeIngEgreCompraVenta(gGruposIngEgreCompraVenta, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            Set rs = oCajero.GetOpeIngEgreSobranteFaltante(gGruposIngEgreSobFalt, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            
            ' *** Agregado Por RIRO 20130917
            Set rs = oCajero.GetOpeIngEgreRecaudo(CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis)
            For J = 0 To rs.RecordCount - 1
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then 'RIRO20140710
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                Else
                    If rs!Efectivo >= 0 Then
                        nIngresos = nIngresos + rs!Efectivo + rs!OrdenPago
                    Else
                        nEgresos = nEgresos + rs!Efectivo + rs!OrdenPago
                    End If
                End If
                rs.MoveNext
            Next J
            ' *** Fin RIRO
            
            If psUserPersArqueado <> "BOVE" Then
                Set rs = oCajero.GetOpeIngEgreHabDevCajero(gGruposIngEgreHabDev, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
                For J = 0 To rs.RecordCount - 1
                    nEfec = nEfec + rs!Efectivo + rs!OrdenPago
                    rs.MoveNext
                Next J
            Else
                Set rs = oCajero.GetOpeIngEgreHabDevCajero(gGruposIngEgreHabDev, gdFecSis, lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
                For J = 0 To rs.RecordCount - 1
                    nEfec = rs!Efectivo + rs!OrdenPago
                    If rs!cGrupoCod = gsCajHabDeBove Then
                        nHabilitacion = nEfec
                    ElseIf rs!cGrupoCod = gsCajDevABove Then
                        nDevolucion = nEfec
                    ElseIf rs!cGrupoCod = gsCajDevBilletaje Then
                        nDevolucionBilletaje = nEfec
                    Else
                        If rs!Efectivo >= 0 Then
                            nIngresos = nIngresos + rs!Efectivo
                        Else
                            nEgresos = nEgresos + rs!Efectivo
                        End If
                    End If
                    rs.MoveNext
                Next J
                Set rs = oCajero.GetOpeIngEgreHabDev(gGruposIngEgreHabDev, CDate(gdFecSis), lblAgeCod.Caption, i, gdFecSis, gsCodUser)
                For J = 0 To rs.RecordCount - 1
                    nEfec = rs!Efectivo + rs!OrdenPago
                    If rs!cGrupoCod = gsCajHabDeBove Then
                        nHabilitacion = nEfec
                    ElseIf rs!cGrupoCod = gsCajDevABove Then
                        nDevolucion = nEfec
                    ElseIf rs!cGrupoCod = gsCajDevBilletaje Then
                        nDevolucionBilletaje = nEfec
                    Else
                        If rs!Efectivo >= 0 Then
                            nIngresos = nIngresos + rs!Efectivo
                        Else
                            nEgresos = nEgresos + rs!Efectivo
                        End If
                    End If
                    rs.MoveNext
                Next J
            End If
            
            'Set rs = oCajero.GetOpeIngEgreHabDevCajero(gGruposIngEgreHabDev, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
            'For j = 0 To rs.RecordCount - 1
            '    nEfec = nEfec + rs!Efectivo
            '    rs.MoveNext
            'Next j
            'Set rs = oCajero.GetOpeIngEgreHabDevBillCajero(gGruposIngEgreHabDev, CDate(gdFecSis), lblAgeCod.Caption, psUserPersArqueado, i, gdFecSis, gsCodUser)
            'For j = 0 To rs.RecordCount - 1
            '    nEfec = nEfec + rs!Efectivo
            '    rs.MoveNext
            'Next j
            If i = 1 Then
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then ' RIRO20140710
                    nResultSolSist = Format(nEfec, "#,##0.00")
                Else
                    nSaldoEfectivoAyer = oEfectivo.GetSaldoEfectivo(lblAgeCod.Caption, DateAdd("d", -1, CDate(gdFecSis)), i, gdFecSis, gsCodUser, , True)
                    nSaldoMovimientos = nIngresos + nEgresos + nHabilitacion + nDevolucion
                    
                    nResultSolSist = Format(nSaldoEfectivoAyer + nSaldoMovimientos, "#,##0.00")
                End If
            Else
                If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then ' RIRO20140710
                    nResultDolSist = Format(nEfec, "#,##0.00")
                Else
                    nSaldoEfectivoAyer = oEfectivo.GetSaldoEfectivo(lblAgeCod.Caption, DateAdd("d", -1, CDate(gdFecSis)), i, gdFecSis, gsCodUser, , True)
                    nSaldoMovimientos = nIngresos + nEgresos + nHabilitacion + nDevolucion
                    
                    nResultDolSist = Format(nSaldoEfectivoAyer + nSaldoMovimientos, "#,##0.00")
                End If
            End If
            nEfec = 0
            nIngresos = 0
            nEgresos = 0
        Next i
    
        nResultSolBill = CDbl(lblTotalBilletes(0)) + CDbl(lblTotMoneda(0))
        nResultDolBill = CDbl(lblTotalBilletes(1))
        
        If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 3 Then ' RIRO20140710
            If nResultSolBill - nResultSolSist <= 2 And nResultSolBill - nResultSolSist >= -2 Then
                If nResultDolBill - nResultDolSist <= (2 / nTiCamb) And nResultDolBill - nResultDolSist >= -(2 / nTiCamb) Then
                    bConforme = 1
                Else
                    bConforme = 0
                End If
            Else
                bConforme = 0
            End If
        Else
            If nResultSolBill - nResultSolSist = 0 And nResultDolBill - nResultDolSist = 0 Then
                bConforme = 1
            Else
                bConforme = 0
            End If
        End If
        
        lsMovNroFin = oContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Call oCajaGen.GrabaProcesoArqueo(cIdArqueo, fgFechaHoraGrab(lsMovNroIni), lnTipoArqVentBov, CInt(Trim(Right(cboTipoArqueo.Text, 2))), _
                                        Trim(txtOtros.Text), lblAgeCod.Caption, NumVentanilla, lnMovNro, bConforme, txtGlosa.Text, nResultSolSist, _
                                        nResultSolBill, nResultDolSist, nResultDolBill, fgFechaHoraGrab(lsMovNroFin), lsMovNroFin, PersRel, bArqueoSup)
        'RIRO20140710 ERS072 Se agrego "bArqueoSup"
        
        'MARG ERS052-2017----
        Dim oMov As COMDMov.DCOMMov
        Set oMov = New COMDMov.DCOMMov
        lnMovNro = oMov.GetnMovNro(lsMovNroFin)
        
        oVisto.RegistraVistoElectronico lnMovNro, , psUserPersArqueado, lnMovNro
        'END MARG- -------------
        
        ReDim MatDatos(1, 4)
        MatDatos(0, 0) = cIdArqueo
        MatDatos(0, 1) = lblAgeDesc.Caption
        MatDatos(0, 2) = fgFechaHoraGrab(lsMovNroFin)
        MatDatos(0, 3) = lblAgeCod.Caption
        MatDatos(0, 4) = fgFechaHoraGrab(lsMovNroIni)
        
        frmCajaArqueoVentBovResult.Inicio MatDatos, PersRel, lnTipoArqVentBov, IIf(bConforme = 1, True, False), nResultSolSist, _
                                        nResultSolBill, nResultDolSist, nResultDolBill, rsBillMN, rsMonMN, rsBillME
        Unload Me
    End If
End Sub

Private Function ValidaDatos() As Boolean
    Dim i As Integer
    Dim ValidaArqueador As Boolean
    Dim ValidaPersArq As Boolean
    Dim ValidaVeedor As Boolean
    Dim ValidaPersArqIniSesion As Boolean
    Dim MsgValidaVeedor As String
    
    ValidaArqueador = False
    ValidaPersArq = False
    ValidaVeedor = False
    ValidaPersArqIniSesion = False
    
    For i = 1 To grdPersInvolucra.Rows - 1
        If Trim(Right(grdPersInvolucra.TextMatrix(i, 1), 2)) = "" Or Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 2)) = "" Then
            MsgBox "Faltan datos en la lista de personas involucradas", vbInformation, "Aviso"
            cmdNuevo.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    Next i
    
    If cboTipoArqueo.ListIndex = -1 Then
        MsgBox "Falta seleccionar el Tipo de Arqueo", vbInformation, "Aviso"
        cboTipoArqueo.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Trim(Right(cboTipoArqueo.Text, 2)) = "3" Then
        If Trim(txtOtros.Text) = "" Then
            MsgBox "No detalló el tipo de arqueo", vbInformation, "Aviso"
            txtOtros.SetFocus
            ValidaDatos = False
        Exit Function
        End If
    End If
    
    If grdPersInvolucra.TextMatrix(1, 1) = "" Then
        MsgBox "No ingresó ningún personal involucrado", vbInformation, "Aviso"
        cmdNuevo.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    For i = 1 To grdPersInvolucra.Rows - 1
        If Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 2)) = 1 Then ' Validamos que haya un Arqueador/Auditor
            ValidaArqueador = True
        End If
        'If lnTipoArqVentBov = 1 Then
            If Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 2)) = 2 Then ' Validamos que haya un Personal Arqueado
                psUserPersArqueado = Trim(grdPersInvolucra.TextMatrix(i, 2))
                ValidaPersArq = True
            End If
        'End If
'        If lnTipoArqVentBov = 2 Then
'            If Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 2)) = 3 Then  ' Validamos que haya un Veedor y tambien que el veedor sea personal de Contabilidad
'                oUser.Inicio Trim(grdPersInvolucra.TextMatrix(i, 2))
'                If oUser.AreaCod = "021" Then
'                    ValidaVeedor = True
'                Else
'                    MsgValidaVeedor = "La persona registrada como Veedor debe ser personal de Contabilidad"
'                End If
'            Else
'                MsgValidaVeedor = "Necesita registrar a una persona de Contabilidad como Veedor"
'            End If
'        End If
    Next i
    
    If ValidaArqueador = False Then
        MsgBox "Nesecita registrar a una persona como Arqueador/Auditor para proceder", vbInformation, "Aviso"
        cmdNuevo.SetFocus
        ValidaDatos = False
        Exit Function
    End If
'    If lnTipoArqVentBov = 2 Then
'        If ValidaVeedor = False Then
'            MsgBox MsgValidaVeedor, vbInformation, "Aviso"
'            cmdNuevo.SetFocus
'            ValidaDatos = False
'            Exit Function
'        End If
'    End If
    'If lnTipoArqVentBov = 1 Then
        If ValidaPersArq = False Then
            MsgBox "Nesecita registrar a una persona como Personal Arqueado para proceder", vbInformation, "Aviso"
            cmdNuevo.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        For i = 1 To grdPersInvolucra.Rows - 1
            If Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 2)) = 2 Then
                If Trim(grdPersInvolucra.TextMatrix(i, 2)) = gsCodUser Then ' Validamos que el Personal Arqueado sea el que inicie la sesion
                    ValidaPersArqIniSesion = True
                    Exit For
                End If
            End If
        Next i
        If ValidaPersArqIniSesion = False And lnTipoArqVentBov <> 3 Then
            MsgBox "El Personal Arqueado debe ser el usuario que inicio sesión", vbInformation, "Aviso"
            cmdNuevo.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
    'End If
    
    If CDbl(lblTotalBilletes(0).Caption) = 0 Or CDbl(lblTotMoneda(0)) = 0 Or CDbl(lblTotalBilletes(1)) = 0 Then
        If MsgBox("Hay totales que estan en 0.00, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
             ValidaDatos = False
            Exit Function
        End If
    End If
    
    ' *** RIRO20140710 ERS072
    Dim oCls As New COMNCajaGeneral.NCOMCajaGeneral
    Dim nNroArqueos As Integer
    If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 2 Then
        If oCls.ObtenerNroArqueos(gdFecSis, gsCodAge, lnTipoArqVentBov, Val((SpnVentanilla.valor)), nNroArqueos) Then
            If nNroArqueos > 1 Then
                If lnTipoArqVentBov = 1 Then
                    MsgBox "No es posible efectuar mas de dos veces al día el arqueo a una misma ventanilla", vbExclamation, "Aviso"
                Else
                    MsgBox "No es posible efectuar mas de dos veces al día el arqueo a una misma Bóveda", vbExclamation, "Aviso"
                End If
                ValidaDatos = False
                Exit Function
            End If
        Else
            MsgBox "Se presentó un error durante el proceso ode validación.", vbExclamation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
    ElseIf lnTipoArqVentBov = 3 Then
        Dim oUsuariosArea As New COMDConstSistema.DCOMGeneral
        Dim rsTmp As New ADODB.Recordset
                        
        If InStr(1, "006005,007026", gsCodCargo) = 0 Then
            Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(gsCodUser, gdFecSis, gsCodArea, gsCodAge, IIf(gsCodAge = "01", 1, 2), 2)  ' BUSCAR POR USUARIO ARQUEADOR
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF And Not rsTmp.BOF Then
                    For i = 1 To grdPersInvolucra.Rows - 1
                        If Val(Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 5))) = 2 Then 'Si la relacion es de "Arqueado"
                            If rsTmp!cUserArqueado <> grdPersInvolucra.TextMatrix(i, 2) Then
                                MsgBox "La pareja de arqueo no es la correcta, se procederá a cerrar el formulario" & vbNewLine & "Ingrese nuevamente.", vbExclamation, "Aviso"
                                ValidaDatos = False
                                Unload Me
                                Exit Function
                            End If
                        End If
                    Next i
                End If
            End If
        Else
            'For i = 1 To grdPersInvolucra.Rows - 1
            '    If val(Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 5))) = 2 Then 'Si la relacion es de "Arqueado"
            '        Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(grdPersInvolucra.TextMatrix(2, 2), gdFecSis, gsCodArea, gsCodAge, IIf(gsCodAge = "01", 1, 2), 1)  ' BUSCAR POR USUARIO ARQUEADO
            '        If Not rsTmp Is Nothing Then
            '            If Not rsTmp.EOF And Not rsTmp.BOF Then
            '
            '            End If
            '        End If
            '    End If
            'Next i
        End If
        If oCls.ObtenerNroArqueos(gdFecSis, gsCodAge, lnTipoArqVentBov, Val((SpnVentanilla.valor)), nNroArqueos) Then
            If nNroArqueos > 0 Then
                MsgBox "La ventanilla seleccionada ya fue arqueada", vbExclamation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Else
            MsgBox "Se presentó un error durante el proceso ode validación.", vbExclamation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
        
    End If
    ' *** END RIRO
    
    ValidaDatos = True
End Function


Private Sub cmdSalir_Click()
    If MsgBox("Desea salir del proceso de Arqueo??", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Unload Me
End Sub



Private Sub grdPersInvolucra_OnChangeCombo()
    Dim i As Integer, J As Integer
    'If lnTipoArqVentBov = 1 Then
        For i = 1 To grdPersInvolucra.Rows - 1
            If Trim(Right(grdPersInvolucra.TextMatrix(i, 5), 2)) = 2 Then
                For J = 1 To grdPersInvolucra.Rows - 1
                    If Trim(grdPersInvolucra.TextMatrix(i, 2)) <> Trim(grdPersInvolucra.TextMatrix(J, 2)) Then
                        If Trim(grdPersInvolucra.TextMatrix(i, 5)) = Trim(grdPersInvolucra.TextMatrix(J, 5)) Then
                            MsgBox "Sólo puede escoger una persona para cada tipo de relación con el arqueo", vbInformation, "Aviso"
                            grdPersInvolucra.TextMatrix(i, 5) = " "
                            Exit Sub
                        End If
                    End If
                Next J
            End If
        Next i
    'End If
End Sub

Private Sub grdPersInvolucra_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim oAcceso As COMDPersona.UCOMAcceso
    Dim bResult As Boolean 'RIRO20140710 ERS072
    Dim nTipoAge As Integer
    Set oAcceso = New COMDPersona.UCOMAcceso
    bResult = True 'RIRO20140710 ERS072
    Dim i As Integer
    
    If gsCodAge = "01" Then
        nTipoAge = 1
    Else
        nTipoAge = 2
    End If
    
    If pbEsDuplicado Then
        MsgBox "La persona ya esta registrada en la relación.", vbInformation, "Aviso"
        grdPersInvolucra.EliminaFila grdPersInvolucra.row
    End If
    
    For i = 1 To grdPersInvolucra.Rows - 2
        If Trim(grdPersInvolucra.TextMatrix(i, 1)) = Trim(grdPersInvolucra.TextMatrix(grdPersInvolucra.Rows - 1, 1)) Then
            MsgBox "La persona ya esta registrada en la relación.", vbInformation, "Aviso"
            grdPersInvolucra.EliminaFila grdPersInvolucra.row
            Exit Sub
        End If
    Next i
    
'    If grdPersInvolucra.Rows = 2 Then
'        If grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) <> "" And grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) <> "" _
'        And grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) <> "" Then
'            MsgBox "Ud. ya cuenta con arqueador, no es posible realizar la acción", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
    
    
    
    If grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = "" Then
        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = ""
        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = ""
    Else
        Dim ClsPersona As COMDPersona.DCOMPersonas
        Dim R As New ADODB.Recordset
        Dim rsTmp As New ADODB.Recordset 'RIRO20140710 ERS072
        Dim oUsuariosArea As New COMDConstSistema.DCOMGeneral 'RIRO20140710 ERS072
        Set ClsPersona = New COMDPersona.DCOMPersonas
        
        Set R = ClsPersona.BuscaCliente(grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1), BusquedaEmpleadoCodigo)
        If Not (R.EOF And R.BOF) Then
            oUser.Inicio R!cUser
            ' *** RIRO20140710 ERS072
            If lnTipoArqVentBov = 3 Then
                'Comentado: TORE - ERS033-2018
                'If gsCodAge = "01" Then
                '    If InStr(1, sCargosArqueoAgenciaPrincipal, Trim(oUser.PersCargoCod)) = 0 Then
                '        bResult = False
                '    End If
                'Else
                '    If InStr(1, sCargosArqueoOtrasAgencias, Trim(oUser.PersCargoCod)) = 0 Then
                '        bResult = False
                '    End If
                'End If
                'If Not bResult Then
                '    If gsCodAge = "01" Then
                '        MsgBox "Solo podrán acceder a esta opcion los usuarios con cargo de Representante financiero y tasador", vbExclamation, "Aviso"
                '    Else
                '        MsgBox "Solo podrán acceder a esta opcion los usuarios con cargo de Representante financiero, tasador o asesor", vbExclamation, "Aviso"
                '    End If
                '    grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
                '    Exit Sub
                'End If
                'END TORE
                
            ' *** TORE - ERS033-2018 ***
                Set oUsuariosArea = New COMDConstSistema.DCOMGeneral
                'Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser2(gsCodUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADO
                Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser2(R!cUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1) ' BUSCAR POR USUARIO ARQUEADO
            
                If Not rsTmp.EOF And Not rsTmp.BOF Then
                    If rsTmp!nEstado = 1 Then ' Generado
                        MsgBox "El usuario <<" & UCase(R!cUser) & ">> ya formó pareja para el arqueo entre ventanillas" & vbNewLine & _
                        "Consultar con el supervisor de operaciones correspondiente", vbInformation, "Aviso"
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 5) = ""
                        Exit Sub
                    ElseIf rsTmp!nEstado = 2 Then ' Arqueado
                        MsgBox "El usuario <<" & UCase(R!cUser) & ">> ya participó de un arqueo entre ventanillas" & vbNewLine & _
                        "Consultar con el supervisor de operaciones correspondiente", vbInformation, "Aviso"
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 5) = ""
                        Exit Sub
                    ElseIf rsTmp!nEstado = 3 Then ' Pendiente de Arqueo
                        MsgBox "El usuario <<" & UCase(R!cUser) & ">> ya participó de un arqueo entre ventanillas" & _
                        "Consultar con el supervisor de operaciones correspondiente", vbInformation, "Aviso"
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 5) = ""
                        Exit Sub
                    End If
                End If
                
            
               If gsCodCargo = "007026" Then 'REPRESENTANTE FINANCIERO III
                    If oUser.PersCargoCod = "006005" Then 'SUPERVISOR DE OPERACIONES
                        'grdPersInvolucra.AdicionaFila
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = Trim(oUser.PersCod) 'R!cPersCod
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = UCase(oUser.UserCod) 'R!cUser
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = UCase(oUser.UserNom) 'R!cPersNombre
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = UCase(oUser.PersCargo) 'R!cRHCargoDescripcion
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 5) = "ARQUEADOR/AUDITOR                                                                           1"
                        sPersCodArqueador = Trim(oUser.PersCod)
                        
                        Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(grdPersInvolucra.TextMatrix(1, 2), gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1)
                        If Not rsTmp.EOF And Not rsTmp.BOF Then
                            'Indicamos que se tiene que registrar al arqueador
                            If rsTmp!nEstadoArqueador = 0 Then
                                nTipoProceso = 1
                            Else
                                nTipoProceso = 2
                            End If
                            
                            grdPersInvolucra.TextMatrix(1, 7) = rsTmp!nIdArqueado
                            grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstado
                        
                        End If
                        'Set rsTmp = Nothing
                        'Actualizar arqueador
                        'Set ClsMov = New COMNContabilidad.NCOMContFunciones
                        'sMovNro = ClsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                        If nTipoProceso = 1 Then
                            Set rsTmp = oUsuariosArea.ActualizaUsuarioArquedor(CInt(grdPersInvolucra.TextMatrix(1, 7)), CStr(grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2)), sMovNro, 1)
                            'Datos del usuario arqueador
                            If Not rsTmp.EOF And Not rsTmp.BOF Then
                                grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstadoArqueado
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 7) = rsTmp!nIdArqueos
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 8) = rsTmp!nEstado
                            End If
                            Set rsTmp = Nothing
                        Else
                            Set rsTmp = oUsuariosArea.ActualizaUsuarioArquedor(CInt(grdPersInvolucra.TextMatrix(1, 7)), CStr(grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2)), "", 2)
                             If Not rsTmp.EOF And Not rsTmp.BOF Then
                                grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstadoArqueado
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 7) = rsTmp!nIdArqueos
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 8) = rsTmp!nEstado
                            End If
                            Set rsTmp = Nothing
                        End If
                        
                        Set rsTmp = Nothing
                        
                    Else
                        MsgBox "Los usuarios con perfil de RFIII, solo pueden ser arquedo por el SUPERVISOR DE OPERACIONES", vbInformation, "Aviso"
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
                        grdPersInvolucra.row = i
                        grdPersInvolucra.Col = 0
                        cmdNuevo.SetFocus
                        Exit Sub
                    End If
                Else 'NO RFIII
                     If InStr(1, sCargoArquedorSR, Trim(oUser.PersCargoCod)) >= 1 Then 'SUPERVISOR DE OPERACIONES/RFIII
                        nTipoProceso = 1
                        'grdPersInvolucra.AdicionaFila
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = Trim(oUser.PersCod) 'R!cPersCod
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = UCase(oUser.UserCod) 'R!cUser
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = UCase(oUser.UserNom)  'R!cPersNombre
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = UCase(oUser.PersCargo)  'R!cRHCargoDescripcion
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 5) = "ARQUEADOR/AUDITOR                                                                           1"
                        sPersCodArqueador = Trim(oUser.PersCod)
                        
                        
                        If nTipoProceso = 1 Then
                            If Right(grdPersInvolucra.TextMatrix(1, 5), 10) = 1 Then
                                MsgBox "La acción no esta permitida, no se admite doble arqueador para la misma persona. Asegurese que la " _
                                & "persona que quedo impar lo asigne como arquedor.", vbInformation, "Aviso"
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = ""
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = ""
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = ""
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 5) = ""
                        
                                grdPersInvolucra.row = i
                                grdPersInvolucra.Col = 0
                                
                                Exit Sub
                            End If
                            Set rsTmp = oUsuariosArea.ActualizaUsuarioArquedor(CInt(grdPersInvolucra.TextMatrix(1, 7)), CStr(grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2)), sMovNro, 1)
                            'Datos del usuario arqueador
                            If Not rsTmp.EOF And Not rsTmp.BOF Then
                                grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstadoArqueado
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 7) = rsTmp!nIdArqueos
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 8) = rsTmp!nEstado
                            End If
                            Set rsTmp = Nothing
                        Else
                            Set rsTmp = oUsuariosArea.ActualizaUsuarioArquedor(CInt(grdPersInvolucra.TextMatrix(1, 7)), CStr(grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2)), "", 2)
                             If Not rsTmp.EOF And Not rsTmp.BOF Then
                                 grdPersInvolucra.TextMatrix(1, 8) = rsTmp!nEstadoArqueado
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 7) = rsTmp!nIdArqueado
                                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 8) = rsTmp!nEstado
                            End If
                            Set rsTmp = Nothing
                        End If
                        
                    Else
                        MsgBox "El personal ingresado no cumple con el perfil de RFIII o SUPERVISOR DE OPERACIONES", vbInformation, "Aviso"
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = ""
                        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = ""
                 
                        grdPersInvolucra.row = i
                        grdPersInvolucra.Col = 0
                        cmdNuevo.SetFocus
                        Exit Sub
                    End If
                End If
              ' *** END TORE ***
                ''Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(R!cUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1)
                ''If Not rsTmp.EOF And Not rsTmp.BOF Then
                    ''grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 7) = rsTmp!nIdArqueado
                    ''grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 8) = rsTmp!nEstado
                ''End If
                ''Set rsTmp = Nothing
            End If
            ' *** END RIRO
            
            If lnTipoArqVentBov = 1 Or lnTipoArqVentBov = 2 Then
                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = UCase(R!cUser)
                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = R!cPersNombre
                grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = oUser.PersCargo
            End If
            
        Else
            MsgBox "Debe ingresar a un empleado de la Institución", vbInformation, "Aviso"
            'grdPersInvolucra.EliminaFila grdPersInvolucra.Row
            grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
            grdPersInvolucra.row = i
            grdPersInvolucra.Col = 0
            cmdNuevo.SetFocus
            Exit Sub
        End If
    End If
    Set ClsPersona = Nothing
End Sub

'Private Sub grdPersInvolucra_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
'    Dim oAcceso As COMDPersona.UCOMAcceso
'    Dim bResult As Boolean 'RIRO20140710 ERS072
'    Dim nTipoAge As Integer
'    Set oAcceso = New COMDPersona.UCOMAcceso
'    bResult = True 'RIRO20140710 ERS072
'
'    Dim i As Integer
'
'    If gsCodAge = "01" Then
'        nTipoAge = 1
'    Else
'        nTipoAge = 2
'    End If
'
'    If pbEsDuplicado Then
'        MsgBox "La persona ya esta registrada en la relación.", vbInformation, "Aviso"
'        grdPersInvolucra.EliminaFila grdPersInvolucra.row
'    End If
'
'    For i = 1 To grdPersInvolucra.Rows - 2
'        If Trim(grdPersInvolucra.TextMatrix(i, 1)) = Trim(grdPersInvolucra.TextMatrix(grdPersInvolucra.Rows - 1, 1)) Then
'            MsgBox "La persona ya esta registrada en la relación.", vbInformation, "Aviso"
'            grdPersInvolucra.EliminaFila grdPersInvolucra.row
'            Exit Sub
'        End If
'    Next i
'
'    If grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1) = "" Then
'        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
'        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = ""
'        grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = ""
'    Else
'        Dim ClsPersona As COMDPersona.DCOMPersonas
'        Dim R As New ADODB.Recordset
'        Dim rsTmp As New ADODB.Recordset 'RIRO20140710 ERS072
'        Dim oUsuariosArea As New COMDConstSistema.DCOMGeneral 'RIRO20140710 ERS072
'        Set ClsPersona = New COMDPersona.DCOMPersonas
'        Set R = ClsPersona.BuscaCliente(grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 1), BusquedaEmpleadoCodigo)
'        If Not (R.EOF And R.BOF) Then
'            oUser.Inicio R!cUser
'            ' *** RIRO20140710 ERS072
'            If lnTipoArqVentBov = 3 Then
'                If gsCodAge = "01" Then
'                    If InStr(1, sCargosArqueoAgenciaPrincipal, Trim(oUser.PersCargoCod)) = 0 Then
'                        bResult = False
'                    End If
'                Else
'                    If InStr(1, sCargosArqueoOtrasAgencias, Trim(oUser.PersCargoCod)) = 0 Then
'                        bResult = False
'                    End If
'                End If
'                If Not bResult Then
'                    If gsCodAge = "01" Then
'                        MsgBox "Solo podrán acceder a esta opcion los usuarios con cargo de Representante financiero y tasador", vbExclamation, "Aviso"
'                    Else
'                        MsgBox "Solo podrán acceder a esta opcion los usuarios con cargo de Representante financiero, tasador o asesor", vbExclamation, "Aviso"
'                    End If
'                    grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
'                    Exit Sub
'                End If
'                Set rsTmp = oUsuariosArea.GetUserAreaAgenciaRelacionXuser(R!cUser, gdFecSis, gsCodArea, gsCodAge, nTipoAge, 1)
'                If Not rsTmp.EOF And Not rsTmp.BOF Then
'                    grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 7) = rsTmp!nIdArqueado
'                    grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 8) = rsTmp!nEstado
'                End If
'                Set rsTmp = Nothing
'            End If
'            ' *** END RIRO
'            grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = UCase(R!cUser)
'            grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 3) = R!cPersNombre
'            grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 4) = oUser.PersCargo
'        Else
'            MsgBox "Debe ingresar a un empleado de la Institución", vbInformation, "Aviso"
'            'grdPersInvolucra.EliminaFila grdPersInvolucra.Row
'            grdPersInvolucra.TextMatrix(grdPersInvolucra.row, 2) = ""
'            grdPersInvolucra.row = i
'            grdPersInvolucra.Col = 0
'            cmdNuevo.SetFocus
'            Exit Sub
'        End If
'    Set ClsPersona = Nothing
'    End If
'
'    'grdPersInvolucra.SetFocus
'End Sub

Private Sub fgBilletes_OnValidate(Index As Integer, ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnTotal As Currency
Dim lnValor As Currency
Select Case pnCol
    Case 3
        lnValor = CCur(IIf(fgBilletes(Index).TextMatrix(pnRow, pnCol) = "", "0", fgBilletes(Index).TextMatrix(pnRow, pnCol)))
        If Residuo(lnValor, CCur(fgBilletes(Index).TextMatrix(pnRow, 5))) Then
            fgBilletes(Index).TextMatrix(pnRow, 2) = Format(Round(lnValor / fgBilletes(Index).TextMatrix(pnRow, 5), 0), "#,##0")
        Else
            Cancel = False
            Exit Sub
        End If
    Case 2
        lnValor = CCur(IIf(fgBilletes(Index).TextMatrix(pnRow, pnCol) = "", "0", fgBilletes(Index).TextMatrix(pnRow, pnCol)))
        fgBilletes(Index).TextMatrix(pnRow, 3) = Format(lnValor * CCur(IIf(fgBilletes(Index).TextMatrix(pnRow, 5) = "", "0", fgBilletes(Index).TextMatrix(pnRow, 5))), "#,##0.00")
End Select
lblTotalBilletes(Index) = Format(fgBilletes(Index).SumaRow(3), "#,##0.00")
'lnTotal = Format(CCur(lblTotalBilletes(Index)) + CCur(lblTotMoneda(Index)), "#,##0.00")
lnTotal = Format(CCur(lblTotalBilletes(Index)), "#,##0.00")
'lblTotal(Index) = Format(lnTotal, "#,##0.00")
End Sub
    
Private Sub fgMonedas_OnValidate(Index As Integer, ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnTotal As Currency
Dim lnValor As Currency
Select Case pnCol
    Case 3
        lnValor = CCur(IIf(fgMonedas(Index).TextMatrix(pnRow, pnCol) = "", "0", fgMonedas(Index).TextMatrix(pnRow, pnCol)))
        If Residuo(lnValor, CCur(fgMonedas(Index).TextMatrix(pnRow, 5))) Then
            fgMonedas(Index).TextMatrix(pnRow, 2) = Format(Round(lnValor / fgMonedas(Index).TextMatrix(pnRow, 5), 0), "#,##0")
        Else
            Cancel = False
            Exit Sub
        End If
    Case 2
        lnValor = CCur(IIf(fgMonedas(Index).TextMatrix(pnRow, pnCol) = "", "0", fgMonedas(Index).TextMatrix(pnRow, pnCol)))
        fgMonedas(Index).TextMatrix(pnRow, 3) = Format(lnValor * CCur(IIf(fgMonedas(Index).TextMatrix(pnRow, 5) = "", "0", fgMonedas(Index).TextMatrix(pnRow, 5))), "#,##0.00")
End Select
lblTotMoneda(Index) = Format(fgMonedas(Index).SumaRow(3), "#,##0.00")
lnTotal = Format(CCur(lblTotalBilletes(Index)) + CCur(lblTotMoneda(Index)), "#,##0.00")
'lblTotal(Index) = Format(lnTotal, "#,##0.00")
End Sub
Private Sub TxtOtros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdNuevo.SetFocus
    End If
End Sub
' *** RIRO20140710
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub
Private Sub grdPersInvolucra_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdPersInvolucra.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub

Private Sub grdPersInvolucra_RowColChange()

    Dim nRow As Integer
    Dim nCol As Integer
        
    If lnTipoArqVentBov <> 3 Then
        Exit Sub
    End If
    
    nRow = grdPersInvolucra.row
    nCol = grdPersInvolucra.Col
    
'    If grdPersInvolucra.TextMatrix(nRow, 1) = sPersCodArqueado And sPersCodArqueado <> "" Then
'        grdPersInvolucra.lbEditarFlex = False
'    ElseIf grdPersInvolucra.TextMatrix(nRow, 1) = sPersCodArqueador And sPersCodArqueador <> "" Then
'        grdPersInvolucra.lbEditarFlex = False
'    Else
'        'grdPersInvolucra.lbEditarFlex = True 'Comentado TORE - ERS033-2018
'        ' *** TORE - ERS033-2018 ***
''        If Right(grdPersInvolucra.TextMatrix(nRow, 5), 15) = 2 Then
''            grdPersInvolucra.lbEditarFlex = False
''        Else
'            grdPersInvolucra.lbEditarFlex = True
''        End If
'        ' *** END TORE ***
'    End If
    
    
    
    If grdPersInvolucra.TextMatrix(nRow, 1) = sPersCodArqueado And sPersCodArqueado <> "" Then
        grdPersInvolucra.lbEditarFlex = False
    End If
    If grdPersInvolucra.Rows > 2 Then
        If grdPersInvolucra.TextMatrix(2, 1) = "" Or grdPersInvolucra.TextMatrix(2, 2) = "" Or grdPersInvolucra.TextMatrix(2, 3) = "" Then
            grdPersInvolucra.lbEditarFlex = True
        Else
            grdPersInvolucra.lbEditarFlex = False
        End If
    Else
        'grdPersInvolucra.lbEditarFlex = True 'Comentado TORE - ERS033-2018
        ' *** TORE - ERS033-2018 ***
'        If Right(grdPersInvolucra.TextMatrix(nRow, 5), 15) = 2 Then
'            grdPersInvolucra.lbEditarFlex = False
'        Else
            grdPersInvolucra.lbEditarFlex = True
'        End If
        ' *** END TORE ***
    End If
    
        
End Sub



'Private Sub grdPersInvolucra_RowColChange()
'
'    Dim nRow As Integer
'    Dim nCol As Integer
'
'    If lnTipoArqVentBov <> 3 Then
'        Exit Sub
'    End If
'
'    nRow = grdPersInvolucra.row
'    nCol = grdPersInvolucra.Col
'
'    If grdPersInvolucra.TextMatrix(nRow, 1) = sPersCodArqueado And sPersCodArqueado <> "" Then
'        grdPersInvolucra.lbEditarFlex = False
'    ElseIf grdPersInvolucra.TextMatrix(nRow, 1) = sPersCodArqueador And sPersCodArqueador <> "" Then
'        grdPersInvolucra.lbEditarFlex = False
'    Else
'        grdPersInvolucra.lbEditarFlex = True
'    End If
'
'End Sub


' *** END RIRO









