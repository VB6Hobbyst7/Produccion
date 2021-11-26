VERSION 5.00
Begin VB.Form frmCapCargos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6960
   ClientLeft      =   1320
   ClientTop       =   1980
   ClientWidth     =   9030
   Icon            =   "frmCapCargos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctNotaAbono 
      Height          =   300
      Left            =   2310
      Picture         =   "frmCapCargos.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   29
      Top             =   7410
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pctCheque 
      Height          =   345
      Left            =   2040
      Picture         =   "frmCapCargos.frx":088C
      ScaleHeight     =   285
      ScaleWidth      =   120
      TabIndex        =   28
      Top             =   7410
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1275
      Left            =   45
      TabIndex        =   16
      Top             =   0
      Width           =   8910
      Begin SICMACMOPEINTER.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   90
         TabIndex        =   0
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Frame fraDatos 
         Height          =   585
         Left            =   45
         TabIndex        =   17
         Top             =   630
         Width           =   8700
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lblApertura 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   900
            TabIndex        =   20
            Top             =   195
            Width           =   1965
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Contacto :"
            Height          =   195
            Left            =   3435
            TabIndex        =   19
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lblUltContacto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4740
            TabIndex        =   18
            Top             =   195
            Width           =   1995
         End
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   3960
      End
   End
   Begin VB.Frame fraCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3195
      Left            =   45
      TabIndex        =   9
      Top             =   1215
      Width           =   8925
      Begin VB.CommandButton cmdMostrarFirma 
         Caption         =   "Mostrar Firma"
         Height          =   315
         Left            =   7350
         TabIndex        =   48
         Top             =   2025
         Width           =   1440
      End
      Begin SICMACMOPEINTER.FlexEdit grdCliente 
         Height          =   1755
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3096
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Direccion-ID-Firma Oblig"
         EncabezadosAnchos=   "250-1700-3800-1500-0-0-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   255
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin VB.Label LblTituloExoneracion 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Exoneración :"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   2835
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblExoneracion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXONERADA POR  ..."
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
         Height          =   300
         Left            =   1620
         TabIndex        =   46
         Top             =   2790
         Visible         =   0   'False
         Width           =   7170
      End
      Begin VB.Label lblMinFirmas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   6330
         TabIndex        =   45
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mínimo Firmas :"
         Height          =   195
         Left            =   5145
         TabIndex        =   44
         Top             =   2130
         Width           =   1110
      End
      Begin VB.Label lblAlias 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1620
         TabIndex        =   43
         Top             =   2430
         Width           =   7185
      End
      Begin VB.Label Label3 
         Caption         =   "Alias de la Cuenta:"
         Height          =   225
         Left            =   180
         TabIndex        =   42
         Top             =   2490
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   2123
         Width           =   960
      End
      Begin VB.Label lblTipoCuenta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1620
         TabIndex        =   14
         Top             =   2070
         Width           =   1800
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "# Firmas :"
         Height          =   195
         Left            =   3645
         TabIndex        =   13
         Top             =   2123
         Width           =   690
      End
      Begin VB.Label lblFirmas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   4395
         TabIndex        =   12
         Top             =   2070
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   45
      TabIndex        =   11
      Top             =   6525
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   6525
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      Top             =   6525
      Width           =   1000
   End
   Begin VB.Frame fraMonto 
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2040
      Left            =   5040
      TabIndex        =   25
      Top             =   4410
      Width           =   3930
      Begin VB.CheckBox chkITFEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   37
         Top             =   1125
         Width           =   705
      End
      Begin SICMACMOPEINTER.EditMoney txtMonto 
         Height          =   375
         Left            =   1455
         TabIndex        =   7
         Top             =   540
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         ForeColor       =   192
         Text            =   "0"
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1410
         TabIndex        =   41
         Top             =   1470
         Width           =   1905
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   2220
         TabIndex        =   40
         Top             =   1065
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   855
         TabIndex        =   39
         Top             =   1530
         Width           =   450
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   855
         TabIndex        =   38
         Top             =   1125
         Width           =   330
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3390
         TabIndex        =   27
         Top             =   555
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   855
         TabIndex        =   26
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2040
      Left            =   45
      TabIndex        =   23
      Top             =   4410
      Width           =   4920
      Begin VB.TextBox txtCtaBanco 
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   35
         Top             =   330
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboMonedaBanco 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   705
         Visible         =   0   'False
         Width           =   1050
      End
      Begin SICMACMOPEINTER.TxtBuscar txtBanco 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   330
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
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
         psRaiz          =   "BANCOS"
         sTitulo         =   ""
      End
      Begin VB.TextBox txtGlosa 
         Height          =   690
         Left            =   1095
         TabIndex        =   6
         Top             =   1080
         Width           =   3600
      End
      Begin VB.TextBox txtOrdenPago 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   330
         Width           =   1155
      End
      Begin VB.CommandButton cmdDocumento 
         Height          =   350
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   315
         Width           =   475
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2565
      End
      Begin VB.Label lblCtaBanco 
         Caption         =   "Cta Banco :"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblEtqBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   2520
         TabIndex        =   33
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblBanco 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   1110
         TabIndex        =   32
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblOrdenPago 
         AutoSize        =   -1  'True
         Caption         =   "Orden Pago :"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   393
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmCapCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PbPreIngresado As Boolean
Private pnMovNro As Long

Public nProducto As COMDConstantes.Producto
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim nMoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim bDocumento As Boolean
Dim nDocumento As COMDConstantes.TpoDoc
Dim dFechaValorizacion As Date
Dim sPersCodCMAC As String
Dim sNombreCMAC As String, sTipoCuenta As String
Dim nPersoneria As COMDConstantes.PersPersoneria
Dim pbOrdPag As Boolean
Dim sOperacion As String
Dim nSaldoCuenta As Double
Dim nSaldoRetiro As Double

'Variables para la impresion de la boleta de Lavado de Dinero
Dim sPersCod As String, sDocId As String, sDireccion As String
Dim sPersCodRea As String, sNombreRea As String, sDocIdRea As String, sDireccionRea As String
Dim sNombre As String

'Variable para la autorización de retiros y Cancelaciones
Dim sMovNroAut As String
'Variables para el ITF
Dim lbITFCtaExonerada As Boolean

'Variable para obtener el SubProduto de Ahorros
Dim lnTpoPrograma As Integer
Dim lsDescTpoPrograma As String

'Funcion de Impresion de Boletas

Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    'Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

'********************************************
Private Function EsTitularExoneradoLavDinero() As Boolean
Dim bExito As Boolean
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim sPersCod As String
bExito = True
For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    If nRelacion = gCapRelPersTitular Then
        sPersCod = grdCliente.TextMatrix(i, 1)
        
        Exit For
    End If
Next i
EsTitularExoneradoLavDinero = bExito
End Function

Private Sub IniciaLavDinero(poLavDinero As SICMACMOPEINTER.frmMovLavDinero)
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nMonto As Double
Dim sCuenta As String
Dim oDatos As COMDPersona.DCOMPersonas
Dim rsPersona As New ADODB.Recordset

For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    If nPersoneria = gPersonaNat Then
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.TitPersLavDineroDir = grdCliente.TextMatrix(i, 4)
            poLavDinero.TitPersLavDineroDoc = grdCliente.TextMatrix(i, 5)
            Exit For
        End If
    Else
        'By Capi 08072008
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.TitPersLavDineroDir = grdCliente.TextMatrix(i, 4)
            poLavDinero.TitPersLavDineroDoc = grdCliente.TextMatrix(i, 5)
            'Exit For
        End If
        '
        If nRelacion = gCapRelPersRepTitular Then
            poLavDinero.ReaPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.ReaPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.ReaPersLavDineroDir = grdCliente.TextMatrix(i, 4)
            poLavDinero.ReaPersLavDineroDoc = grdCliente.TextMatrix(i, 5)
            Exit For
        End If
    End If
Next i
nMonto = txtMonto.value
sCuenta = txtCuenta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPerscod, sNombre, sDireccion, sDocId, False, False, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
'    If txtCuenta.Prod = Producto.gCapCTS Then
'        IniciaLavDinero = frmMovLavDinero.Inicia(sPerscod, sNombre, sDireccion, sDocId, False, False, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'    Else
'        IniciaLavDinero = frmMovLavDinero.Inicia(sPerscod, sNombre, sDireccion, sDocId, False, False, nMonto, sCuenta, sOperacion, True, sTipoCuenta)
'    End If
'End If
End Sub

Sub MuestraDataTarj(ByVal sNumTar As String, ByVal sClaveTar As String, ByVal sCaption As String)
Dim lnResult As ResultVerificacionTarjeta
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim nEstado As Integer
Set clsGen = New COMDConstSistema.DCOMGeneral
        
Select Case clsGen.ValidaTarjeta(sNumTar, sClaveTar)
    Case gClaveValida
            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
            Dim rsTarj As New ADODB.Recordset
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
            If rsTarj.EOF And rsTarj.BOF Then
                MsgBox "Tarjeta no posee ninguna relación con cuentas de activas   o Tarjeta no activa.", vbInformation, "Aviso"
                Me.Caption = sCaption
                Set clsGen = Nothing
                Exit Sub
            Else
                nEstado = rsTarj("nEstado")
                If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
                    If nEstado = gCapTarjEstBloqueada Then
                        MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                    ElseIf nEstado = gCapTarjEstCancelada Then
                        MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                    End If
                    Me.Caption = sCaption
                    Set clsGen = Nothing
                    Exit Sub
                End If
                Dim rsPers As New ADODB.Recordset
                Dim sCta As String, sProducto As String, sMoneda As String
                Dim clsCuenta As UCapCuenta
                Set rsPers = clsMant.GetTarjetaCuentas(sNumTar, nProducto)
                Set clsMant = Nothing
                If Not (rsPers.EOF And rsPers.EOF) Then
                    Do While Not rsPers.EOF
                        sCta = rsPers("cCtaCod")
                        sProducto = rsPers("Producto")
                        sMoneda = Trim(rsPers("Moneda"))
                        frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sProducto & Space(2) & sMoneda
                        rsPers.MoveNext
                    Loop
                    Set clsCuenta = New UCapCuenta
                    Set clsCuenta = frmCapMantenimientoCtas.Inicia
                    If clsCuenta.sCtaCod <> "" Then
                        txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                        txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
                        txtCuenta.cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                        txtCuenta.SetFocusCuenta
                        Call txtCuenta_KeyPress(13)
                    End If
                    Set clsCuenta = Nothing
                Else
                    MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
                End If
                rsPers.Close
                Set rsPers = Nothing
            End If
            Set rsTarj = Nothing
            Set clsMant = Nothing
    Case gTarjNoRegistrada
        'ppoa Modificacion
'        If Not WriteToLcd("Espere Por Favor") Then
'            FinalizaPinPad
'            MsgBox "No se Realizó Envío", vbInformation, "Aviso"
'            Set clsGen = Nothing
'            Exit Sub
'        End If
        MsgBox "Tarjeta no Registrada", vbInformation, "Aviso"
    Case gClaveNOValida
        'ppoa Modificacion
 '       If Not WriteToLcd("Clave Incorrecta") Then
 '           MsgBox "No se Realizó Envío", vbInformation, "Aviso"
 '           Set clsGen = Nothing
 '           Exit Sub
 '       End If
        MsgBox "Clave Incorrecta", vbInformation, "Aviso"
End Select
Set clsGen = Nothing
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
Set clsCap = Nothing
If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        If bDocumento And nDocumento = TpoDocOrdenPago Then
            If Not rsCta("bOrdPag") Then
                rsCta.Close
                Set rsCta = Nothing
                MsgBox "Cuenta NO fue aperturada con ORDEN DE PAGO", vbInformation, "Aviso"
                txtCuenta.cuenta = ""
                txtCuenta.SetFocus
                Exit Sub
            End If
        End If
        'ITF INICIO
        lbITFCtaExonerada = fgITFVerificaExoneracion(sCuenta)
        fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
        grdCliente.lbEditarFlex = False
        If sPersCodCMAC = "" Then
            If nProducto = gCapAhorros Then
                If gbITFAsumidoAho Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.value = 0
                End If
            ElseIf nProducto = gCapPlazoFijo Then
                If gbITFAsumidoPF Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.Enabled = True
                    Me.chkITFEfectivo.value = 1
                End If
            ElseIf nProducto = gCapCTS Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
            End If
        Else
            chkITFEfectivo.Enabled = False
            If nProducto = gCapAhorros Then
                If gbITFAsumidoAho Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.value = 0
                End If
            ElseIf nProducto = gCapPlazoFijo Then
                If gbITFAsumidoPF Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.value = 1
                End If
            ElseIf nProducto = gCapCTS Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
            End If
        End If
        'ITF FIN
        
        nSaldoCuenta = rsCta("nSaldoDisp")
        nEstado = rsCta("nPrdEstado")
        nPersoneria = rsCta("nPersoneria")
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm:ss")
        nMoneda = CLng(Mid(sCuenta, 9, 1))
        
        If nMoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            txtMonto.BackColor = &HC0FFFF
            lblMon.Caption = "S/."
        Else
            sMoneda = "MONEDA EXTRANJERA"
            txtMonto.BackColor = &HC0FFC0
            lblMon.Caption = "US$"
        End If
        
        lblITF.BackColor = txtMonto.BackColor
        lblTotal.BackColor = txtMonto.BackColor
        
        Select Case nProducto
            Case gCapAhorros
                If rsCta("bOrdPag") Then
                    lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
                    pbOrdPag = True
                Else
                    lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
                    pbOrdPag = False
                End If
                lblUltContacto = Format$(rsCta("dUltContacto"), "dd mmm yyyy hh:mm:ss")
                Me.lblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.lblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
                
                If lbITFCtaExonerada Then
                    Dim nTipoExo As String, sDescripcion As String
                    nTipoExo = fgITFTipoExoneracion(sCuenta, sDescripcion)
                    LblTituloExoneracion.Visible = True
                    lblExoneracion.Visible = True
                    lblExoneracion.Caption = sDescripcion
                End If
                
               If nProducto = gCapAhorros Then
                    Dim oCons As COMDConstantes.DCOMConstantes
                    Set oCons = New COMDConstantes.DCOMConstantes
                    lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
                    lsDescTpoPrograma = Trim(oCons.DameDescripcionConstante(2030, lnTpoPrograma))
                    Set oCons = Nothing
               End If

                
            Case gCapPlazoFijo
                lblUltContacto = rsCta("nPlazo")
                Me.lblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.lblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
            
            Case gCapCTS
                lblUltContacto = rsCta("cInstitucion")
                Dim nDiasTranscurridos As Long
                Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
                Dim nSaldoMinRet As Double
                    
                Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
                nSaldoMinRet = clsDef.GetSaldoMinimoPersoneria(nProducto, nMoneda, nPersoneria, False)
                Set clsDef = Nothing
                If rsCta("nSaldoDisp") - rsCta("nSaldRetiro") >= nSaldoMinRet Then
                    txtMonto.Text = Format(rsCta("nSaldRetiro"), "#,##0.00")
                Else
                    txtMonto.Text = IIf((rsCta("nSaldRetiro") - nSaldoMinRet) > 0, Format(rsCta("nSaldRetiro") - nSaldoMinRet, "### ###,##0.00"), "0.00")
                End If
                nSaldoRetiro = rsCta("nSaldRetiro")
        End Select
        
        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
        sTipoCuenta = lblTipoCuenta
        nTipoCuenta = rsCta("nPrdCtaTpo")
        lblFirmas = Format$(rsCta("nFirmas"), "#0")
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        
        sPersona = ""
        
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales 'DCapMantenimiento
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
                
        Do While Not rsRel.EOF
            If rsRel("cPersCod") = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                grdCliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & Space(50) & Trim(rsRel("nPrdPersRelac"))
                grdCliente.TextMatrix(nRow, 4) = rsRel("Direccion") & ""
                grdCliente.TextMatrix(nRow, 5) = rsRel("ID N°")
                grdCliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("cobligatorio")) Or rsRel("cobligatorio") = "N", "NO", IIf(rsRel("cobligatorio") = "S", "SI", "OPCIONAL"))
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
        
        rsRel.Close
        Set rsRel = Nothing
        fraCliente.Enabled = True
        fraDocumento.Enabled = True
        fraMonto.Enabled = True
        
        If cboDocumento.Visible Then
            cboDocumento.Enabled = True
        ElseIf txtOrdenPago.Visible Then
            txtOrdenPago.SetFocus
        Else
            txtGlosa.SetFocus
        End If
        fraCuenta.Enabled = False
        
        cmdGrabar.Enabled = True
       
        cmdCancelar.Enabled = True
    End If
    
'    MuestraFirmas sCuenta
    
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
Set clsMant = Nothing
End Sub

Private Sub LimpiaControles()
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
txtGlosa = ""
txtMonto.value = 0
If bDocumento Then

End If
txtMonto.BackColor = &HC0FFFF
lblITF.BackColor = txtMonto.BackColor
lblTotal.BackColor = txtMonto.BackColor

lblMon.Caption = "S/."
lblMensaje = ""
cmdGrabar.Enabled = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = ""
txtCuenta.cuenta = ""
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
lblApertura = ""
lblUltContacto = ""
lblFirmas = ""
lblTipoCuenta = ""
fraCliente.Enabled = False
fraDatos.Enabled = False
fraDocumento.Enabled = False
fraMonto.Enabled = False
fraCuenta.Enabled = True
txtOrdenPago.Text = ""

txtCuenta.SetFocus

txtBanco.Text = ""
txtCtaBanco.Text = ""
Me.cboMonedaBanco.ListIndex = -1
Me.lblBanco.Caption = ""

lblAlias.Caption = ""
lblMinFirmas.Caption = ""

nSaldoCuenta = 0
sMovNroAut = ""
If nProducto = Producto.gCapAhorros Then
        Label3.Visible = True
        Label5.Visible = True
           
        lblAlias.Visible = True
        lblMinFirmas.Visible = True
ElseIf nProducto = Producto.gCapCTS Then
        Label3.Visible = False
        Label5.Visible = False
        lblAlias.Visible = False
        lblMinFirmas.Visible = False
End If
lblExoneracion.Visible = False
LblTituloExoneracion.Visible = False
End Sub

Public Sub Inicia(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, _
        ByVal sDescOperacion As String, Optional sCodCmac As String = "", _
        Optional sNomCmac As String, Optional lcCtaCod As String, Optional pnMonto As Double, Optional lnMovNro As Long)

nProducto = nProd
nOperacion = nOpe
sPersCodCMAC = sCodCmac
sNombreCMAC = sNomCmac
sOperacion = sDescOperacion

Select Case nProd
    Case gCapAhorros
        lblEtqUltCnt = "Ult. Contacto :"
        lblUltContacto.Width = 2000
        txtCuenta.Prod = Trim(Str(gCapAhorros))
        If sPersCodCMAC = "" Then
            Me.Caption = "Captaciones - Cargo - Ahorros " & sDescOperacion
        Else
            Me.Caption = "Captaciones - Cargo - Ahorros " & sDescOperacion & " - " & sNombreCMAC
        End If
        Label3.Visible = True
        Label5.Visible = True
        
        lblAlias.Visible = True
        lblMinFirmas.Visible = True
        
        grdCliente.ColWidth(6) = 1200
        
    Case gCapPlazoFijo
    
        lblEtqUltCnt = "Plazo :"
        lblUltContacto.Width = 1000
        txtCuenta.Prod = Trim(Str(gCapPlazoFijo))
        Me.Caption = "Captaciones - Cargo - Plazo Fijo " & sDescOperacion
        Label3.Visible = True
        Label5.Visible = True
        
        lblAlias.Visible = True
        lblMinFirmas.Visible = True
        
        grdCliente.ColWidth(6) = 1200
               
    Case gCapCTS
        lblEtqUltCnt = "Institución :"
        lblUltContacto.Width = 4250
        lblUltContacto.Left = lblUltContacto.Left - 275
        txtCuenta.Prod = Trim(Str(gCapCTS))
        If sPersCodCMAC = "" Then
            Me.Caption = "Captaciones - Cargo - CTS " & sDescOperacion
        Else
            Me.Caption = "Captaciones - Cargo - CTS " & sDescOperacion & " - " & sNombreCMAC
        End If
        fraMonto.Visible = True
        Label3.Visible = False
        Label5.Visible = False
        
        lblAlias.Visible = False
        lblMinFirmas.Visible = False
        grdCliente.ColWidth(6) = 0
End Select


'Verifica si la operacion necesita algun documento
Dim clsOpe As COMDConstSistema.DCOMOperacion 'DOperacion
Dim rsDoc As ADODB.Recordset
Set clsOpe = New COMDConstSistema.DCOMOperacion
Set rsDoc = clsOpe.CargaOpeDoc(Trim(nOperacion))
Set clsOpe = Nothing
If Not (rsDoc.EOF And rsDoc.BOF) Then
    nDocumento = rsDoc("nDocTpo")
    If nDocumento = TpoDocOrdenPago Then
        lblDocumento.Visible = False
        cboDocumento.Visible = False
        cmdDocumento.Visible = False
        lblOrdenPago.Visible = True
        txtOrdenPago.Visible = True
        
        If nOperacion = gAhoRetOPCanje Then
            txtBanco.Visible = True
            lblBanco.Visible = True
            lblEtqBanco.Visible = True
            Dim clsBanco As COMNCajaGeneral.NCOMCajaCtaIF 'NCajaCtaIF
            Dim rsBanco As New ADODB.Recordset
            Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
                Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "_1%", MuestraInstituciones, "1")
            Set clsBanco = Nothing
            txtBanco.rs = rsBanco
             
        Else
            txtBanco.Visible = False
            lblBanco.Visible = False
            lblEtqBanco.Visible = False
            
        End If
    End If
    fraDocumento.Caption = Trim(rsDoc("cDocDesc"))
    bDocumento = True
Else
    If gAhoRetTransf = nOperacion Or gAhoRetEmiChq = nOperacion Or "220302" = nOperacion Or "210202" = nOperacion Then
        Dim oCon As COMDConstantes.DCOMConstantes 'DConstante
        Set oCon = New COMDConstantes.DCOMConstantes
        Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
            Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "0[123]%", MuestraInstituciones)
        Set clsBanco = Nothing
        
        txtBanco.rs = rsBanco
        
        CargaCombo Me.cboMonedaBanco, oCon.RecuperaConstantes(gMoneda)
        Set oCon = Nothing
        
        txtBanco.Visible = True
        lblBanco.Visible = True
        lblEtqBanco.Visible = True
        
        lblDocumento.Visible = True
        cmdDocumento.Visible = True
        
        cboMonedaBanco.Visible = True
        txtCtaBanco.Visible = True
        lblCtaBanco.Visible = True
        
    Else
        txtBanco.Visible = False
        lblBanco.Visible = False
        lblEtqBanco.Visible = False
    
        lblDocumento.Visible = False
        cmdDocumento.Visible = False
        
    End If
    
    lblDocumento.Visible = False
    cboDocumento.Visible = False
    cmdDocumento.Visible = False

    bDocumento = False
    lblOrdenPago.Visible = False
    txtOrdenPago.Visible = False
    
End If
rsDoc.Close
Set rsDoc = Nothing
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledProd = False
txtCuenta.EnabledCMAC = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraCliente.Enabled = False
fraDocumento.Enabled = False
fraMonto.Enabled = False

sMovNroAut = ""

PbPreIngresado = False
If lcCtaCod <> "" Then
    txtCuenta.NroCuenta = lcCtaCod
    txtMonto.value = pnMonto
    PbPreIngresado = True
    pnMovNro = lnMovNro
End If
Me.Show 1
End Sub

Private Sub cboDocumento_Click()
If nDocumento = TpoDocNotaCargo Then
    If cboDocumento.Text = "<Nuevo>" Then
        cmdDocumento.Enabled = True
        txtMonto.Text = "0.00"
    Else
        cmdDocumento.Enabled = False
        Dim nMonto As Double
        nMonto = CDbl(Trim(Right(cboDocumento.Text, 15)))
        txtMonto.Text = Format$(nMonto, "#,##0.00")
    End If
End If
End Sub

Private Sub cboDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

Private Sub cboMonedaBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtGlosa.SetFocus
    End If
    
End Sub

Private Sub cmdCancelar_Click()
LimpiaControles
End Sub


Private Sub cmdGrabar_Click()
    Dim sNroDoc As String, sCodIF As String
    Dim nMonto As Double
    Dim sCuenta As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Dim sMovNro As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim lbProcesaAutorizacion As Boolean
    Dim lsmensaje As String
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    Dim loLavDinero As SICMACMOPEINTER.frmMovLavDinero
    Set loLavDinero = New SICMACMOPEINTER.frmMovLavDinero


    sCuenta = txtCuenta.NroCuenta

    nMonto = txtMonto.value
           
    If nMonto = 0 Then
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        If txtMonto.Enabled Then txtMonto.SetFocus
        Exit Sub
    End If

    'Validar que no se realice retiros de Ctas de Aho Pandero/Panderito/Destino
    If nProducto = gCapAhorros Then
        'By capi 19012009 para que no permita retiros de ahorro ñañito.
        'If lnTpoPrograma = 2 Or lnTpoPrograma = 3 Or lnTpoPrograma = 4 Then
        If lnTpoPrograma = 1 Or lnTpoPrograma = 2 Or lnTpoPrograma = 3 Or lnTpoPrograma = 4 Then
            MsgBox "No se Puede realizar un Retiro de una Cuenta de " & lsDescTpoPrograma, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    '--------------------------------------------------------------------------
    If nProducto = gCapAhorros And nOperacion < 200310 Then
        Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion, nMontoMinRet As Double
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        If pbOrdPag Then
            nMontoMinRet = clsDef.GetMontoMinimoRetOPPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag)
            If nMontoMinRet > nMonto Then
                MsgBox "El Monto de Retiro Cta con Ord. Pago es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, "S/. ", "US$. ") & CStr(nMontoMinRet), vbOKOnly + vbInformation, "Aviso"
                Exit Sub
            End If
        Else
            nMontoMinRet = clsDef.GetMontoMinimoRetPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag)
            If nMontoMinRet > nMonto Then
                MsgBox "El Monto de Retiro es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, "S/. ", "US$. ") & CStr(nMontoMinRet), vbOKOnly + vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        Set clsDef = Nothing
    End If

    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    clsCap.IniciaImpresora gImpresora

    'Valida documento
    If bDocumento Then
        If nDocumento = TpoDocOrdenPago Then
            sNroDoc = Trim(txtOrdenPago)
            If sNroDoc = "" Then
                MsgBox "Debe digitar un N° de Orden de Pago Válido", vbInformation, "Aviso"
                Exit Sub
            End If
            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
            Dim rsOP As New ADODB.Recordset
            Dim nEstadoOP As COMDConstantes.CaptacOrdPagoEstado
            Dim bOPExiste As Boolean
            If clsCap.EsOrdenPagoEmitida(sCuenta, CLng(sNroDoc)) Then
                Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                Set rsOP = clsMant.GetDatosOrdenPago(sCuenta, CLng(sNroDoc))  'DocRecOP
                Set clsMant = Nothing
                If Not (rsOP.EOF And rsOP.BOF) Then
                    bOPExiste = True
                    nEstadoOP = rsOP("nEstado")
                    If nEstadoOP = gCapOPEstAnulada Or nEstadoOP = gCapOPEstCobrada Or nEstadoOP = gCapOPEstExtraviada Then
                        MsgBox "Orden de Pago N° " & sNroDoc & " " & rsOP("cDescripcion"), vbInformation, "Aviso"
                        rsOP.Close
                        Set rsOP = Nothing
                        txtOrdenPago.SetFocus
                        Exit Sub
                    ElseIf rsOP("nEstado") = gCapOPEstCertifiCada Then
                        If nMonto <> rsOP("nMonto") Then
                            MsgBox "Orden de Pago Certificada. Monto No Coincide con Monto de Certificación", vbInformation, "Aviso"
                            txtMonto.Text = Format$(rsOP("nMonto"), "#,##0.00")
                            rsOP.Close
                            Set rsOP = Nothing
                            Exit Sub
                        Else
                            If nOperacion = gAhoRetOP Then
                                nOperacion = gAhoRetOPCert
                            ElseIf nOperacion = gAhoRetOPCanje Then
                                nOperacion = gAhoRetOPCertCanje
                            End If
                        End If
                    End If
                Else
                    bOPExiste = False
                End If
                rsOP.Close
                Set rsOP = Nothing
                If txtBanco.Visible Then
                    If txtBanco.Text = "" Then
                        MsgBox "Debe Seleccionar un Banco.", vbInformation, "Aviso"
                        Exit Sub
                    Else
                        sCodIF = txtBanco.Text
                    End If
                Else
                    sCodIF = ""
                End If
            Else
                MsgBox "Orden de Pago No ha sido emitida para esta cuenta", vbInformation, "Aviso"
                Exit Sub
            End If
        ElseIf nDocumento = TpoDocNotaCargo Then
            sNroDoc = Trim(Left(cboDocumento.Text, 8))
            If InStr(1, sNroDoc, "<Nuevo>", vbTextCompare) > 0 Then
                MsgBox "Debe seleccionar un documento (" & fraDocumento.Caption & ") válido para la operacion.", vbInformation, "Aviso"
                cboDocumento.SetFocus
                Exit Sub
            End If
            sCodIF = ""
        End If
    End If

'Valida Saldo de la Cuenta a Retirar
    If nOperacion <> gAhoRetOPCert And nOperacion <> gAhoRetOPCertCanje Then
        Dim nitf As Double
        nitf = 0
        If gbITFAplica And lbITFCtaExonerada = False Then
            If (nProducto = gCapAhorros And gbITFAsumidoAho = False) Or (nProducto = gCapPlazoFijo And gbITFAsumidoPF = False) Then
                If chkITFEfectivo.value = vbUnchecked Then
                    nitf = CDbl(lblITF.Caption)
                End If
            End If
        End If
    
        Dim nComixMov As Double
        Dim nComixRet As Double
        'By Capi 05032008
        Dim lbCuentaRRHH As Boolean
        Dim loCptGen As COMDCaptaGenerales.DCOMCaptaGenerales
        
        If nOperacion = gCMACOAAhoRetEfec Then
            nComixMov = 0
            nComixRet = 0
        ElseIf nOperacion = gCMACOAAhoRetOP Then
            nComixMov = 0
            nComixRet = 0
        Else
            If nProducto = gCapAhorros Then
                '*****  VERIFICAR MAX MOVIMIENTOS Y CALCULAR COMISION  AVMM 03-06-2006 *****
                nComixMov = Round(CalcularComisionxMaxOpeRet(), 2)
                '***************************************************************************
            
                '******                 CALCULAR COMISION  AVMM-03-2006                *****
                nComixRet = Round(CalcularComisionRetOtraAge(), 2)
                '***************************************************************************
                'By Capi 05032008
                Set loCptGen = New COMDCaptaGenerales.DCOMCaptaGenerales
                lbCuentaRRHH = loCptGen.ObtenerSiEsCuentaRRHH(sCuenta)
                'By capi 19082008 se modifico para la cuenta soñada
                'If lbCuentaRRHH Then
                If lbCuentaRRHH Then
                    nComixMov = 0
                    nComixRet = 0
                End If
                If lnTpoPrograma = 5 Then
                    nComixMov = 0
                    'By Capi 15092008 para que no tome en cuenta las operaciones en otras plazas
                    nComixRet = 0
                    '
                End If
                               
                'End By
                
                
            Else
                nComixMov = 0
               nComixRet = 0
            End If
        End If
    
        If Not clsCap.ValidaSaldoCuenta(sCuenta, nMonto + nitf, , nComixRet, nComixMov) Then
            If nOperacion = gAhoRetOPCanje Then
                'Por ahora en esta validacion no se da un trato especial esperar cambios
                '            If MsgBox("Cuenta NO posee SALDO SUFICIENTE. ¿Desea registrarla como Devuelta?", vbInformation, "Aviso") = vbYes Then
                '                'Registrar Orden de Pago Devuelta
                '
                '            End If
            ElseIf nOperacion = gAhoRetOP Then
                Dim nMaxSobregiro As Long, nSobregiro As Long
                Dim nMontoDescuento As Double, nSaldoMinimo As Double
                Dim nBloqueoParcial As Double
                Dim clsGen As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
                Dim nEstado As COMDConstantes.CaptacEstado
                Dim oMov As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
                Dim sGlosa As String
                Dim nSaldoDisponible As Double
        
                Set clsGen = New COMNCaptaGenerales.NCOMCaptaDefinicion
                nMaxSobregiro = clsGen.GetCapParametro(gNumVecesMinRechOP)
                If nMoneda = gMonedaNacional Then
                    nMontoDescuento = clsGen.GetCapParametro(gMonDctoMNRechOP)
                Else
                    nMontoDescuento = clsGen.GetCapParametro(gMonDctoMERechOP)
                End If
                                   
                Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
                Set rsOP = clsMant.GetDatosCuenta(sCuenta)
                Set clsMant = Nothing
                nSobregiro = rsOP("nSobregiro")
                nEstado = rsOP("nPrdEstado")
                nPersoneria = rsOP("nPersoneria")
                nSaldoDisponible = rsOP("nSaldoDisp")
                nBloqueoParcial = rsOP("nBloqueoParcial")
                nSaldoMinimo = clsGen.GetSaldoMinimoPersoneria(gCapAhorros, nMoneda, nPersoneria, True)
                Set clsGen = Nothing
            
                nPersoneria = rsOP("nPersoneria")
            
                'Determinar el monto de descuento Real de acuerdo a su saldo
                If nSaldoDisponible - nSaldoMinimo <= 0 Then
                    nMontoDescuento = 0
                ElseIf nSaldoDisponible - nMontoDescuento - nSaldoMinimo - nBloqueoParcial < 0 Then
                    nMontoDescuento = nSaldoDisponible - nSaldoMinimo - nBloqueoParcial
                End If
        
                Set clsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
                sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Set clsMov = Nothing
            
                Set oMov = New COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
                oMov.IniciaImpresora gImpresora
            
                If nSobregiro + 1 >= nMaxSobregiro Then 'El ultimo sobregiro, se descuenta y luego se cancela
                    MsgBox "NO POSSE SALDO SUFICIENTE..." & Chr$(13) _
                       & "La Orden de Pago N° " & sNroDoc & " ha sido Sobregirada " & nMaxSobregiro & " VECES!." & Chr$(13) _
                       & "Se procederá a bloquear la cuenta y hacer el descuento por Orden de Pago Rechazada.", vbInformation, "Aviso"
                
                    'Hacer el descuento
                    sGlosa = "OP Rechazada " & sNroDoc & ". Cuenta " & sCuenta
                
                    oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nMoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF
                    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"

                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    clsMant.ActualizaSobregiro sCuenta, nSobregiro + 1, sMovNro, nEstado, False
                    clsMant.BloqueCuentaTotal sCuenta, gCapMotBlqTotOrdenPagoRechazada, sGlosa, sMovNro
                    Set clsMant = Nothing
                
                'COMENTADO HASTA DEFINIR PROCESO 0609-2006-AVMM
    '                oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nMoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF

    '                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
    '                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
    '                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
    '
    '                Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    '                clsMant.ActualizaSobregiro sCuenta, nMaxSobregiro, sMovNro, nEstado, True
    '                Set clsMant = Nothing
    '                'Cancelar la cuenta
    '                oMov.GetSaldoCancelacion sCuenta, gdFecSis, gsCodAge, 0  'x Midificar GetSaldoCancelacion Raul
    '
    '                Set clsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
    '                    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    '                Set clsMov = Nothing
    '
    '                oMov.CapCancelaCuentaAho sCuenta, sMovNro, sGlosa, gAhoCancSobregiroOP, gCapEstCancelada, gsNomAge, sLpt, , gsCodCMAC, False, , , , , lsmensaje, lsBoleta, lsBoletaITF
    '
    '                Set oMov = Nothing
    '
    '                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
    '                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
    '                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
                                
                    cmdCancelar_Click
                    Exit Sub
            
                ElseIf nSobregiro + 1 = nMaxSobregiro - 1 Then 'Una menos que la ultima se descuenta y se bloquea
                    MsgBox "NO POSEE SALDO SUFICIENTE" & Chr$(13) _
                      & "La Cuenta " & sCuenta & " ha sido Sobregirada " & nSobregiro + 1 & " VECES!." & Chr$(13) _
                      & "Se procederá a bloquear la cuenta y hacer el descuento por Orden de Pago Rechazada.", vbInformation, "Aviso"
                    'Hacer el descuento
                    sGlosa = "OP Rechazada " & sNroDoc & ". Cuenta " & sCuenta

                    oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nMoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF
                    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"

                    'Se actualiza el sobregiro y se bloquea la cuenta
                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    clsMant.ActualizaSobregiro sCuenta, nSobregiro + 1, sMovNro, nEstado, False
                    clsMant.BloqueCuentaTotal sCuenta, gCapMotBlqTotOrdenPagoRechazada, sGlosa, sMovNro
                    Set clsMant = Nothing
                
                    cmdCancelar_Click
                    Exit Sub
                Else 'Se descuenta noma
                
                    MsgBox "NO POSEE SALDO SUFICIENTE" & Chr$(13) _
                       & "La Cuenta " & sCuenta & " ha sido Sobregirada " & nSobregiro + 1 & " VECES!." & Chr$(13) _
                       & "Se procederá a hacer el descuento por Orden de Pago Rechazada.", vbInformation, "Aviso"
            
                    'Hacer el descuento
                    'oMov.CapCargoCuentaAho sCuenta, nMontoDescuento, gAhoRetComOrdPagDev, sMovNro, sGlosa, TpoDocOrdenPago, sNroDoc, , True, , , , , , sLpt, , , , nMoneda, gsCodCMAC, , gsCodAge, False, , , , , , , lsmensaje, lsBoleta, lsBoletaITF

                    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
                
                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    clsMant.ActualizaSobregiro sCuenta, nSobregiro + 1, sMovNro, nEstado, False
                    Set clsMant = Nothing
                
                    cmdCancelar_Click
                    Exit Sub
                End If
                Set oMov = Nothing
            Else
                MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
                txtMonto.SetFocus
                Exit Sub
            End If
        End If
    End If

    'Valida que la transaccion no se pueda realizar porque la cuenta no posee firmas
    If Not gbRetiroSinFirma Then
        If Not clsCap.CtaConFirmas(txtCuenta.NroCuenta) Then
            MsgBox "No puede retirar, porque la cuenta no cuenta con las firmas de las personas relacionadas a ella.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

    'Valida la operacion de Retiro por Transferencias
    If nOperacion = gAhoRetTransf Or nOperacion = gCTSRetTransf Then
        If Trim(txtBanco.Text) = "" Then
            MsgBox "Debe seleccionar el Banco a Transferir", vbInformation, "Aviso"
            txtBanco.SetFocus
            Exit Sub
        End If
        If cboMonedaBanco.Text = "" Then
            MsgBox "Debe seleccionar la Moneda de la Cuenta de Banco a Transferir", vbInformation, "Aviso"
            cboMonedaBanco.SetFocus
            Exit Sub
        End If
        If Trim(txtCtaBanco.Text) = "" Then
            MsgBox "Debe digitar la Cuenta de Banco a Transferir", vbInformation, "Aviso"
            txtCtaBanco.SetFocus
            Exit Sub
        End If
    End If

    '----------- Verificar Autorizacion -- AVMM -- 18/04/2004 -------
    If VerificarAutorizacion = False Then Exit Sub
    '----------------------------------------------------------------

    '----------- Verificar Creditos -- CAAU -- 15-12-2006----
    If nProducto = gCapAhorros Then
        Dim k As Integer
        Dim clsCapD As COMDCaptaGenerales.DCOMCaptaGenerales
        Set clsCapD = New COMDCaptaGenerales.DCOMCaptaGenerales
        Dim oCred As COMDCredito.DCOMCredito
        Dim lsMsgCred As String
        Set oCred = New COMDCredito.DCOMCredito
    
        For k = 1 To Me.grdCliente.Rows - 1
            '10= Titular
            If Right(Me.grdCliente.TextMatrix(k, 3), 2) = "10" Then
                If oCred.VerificarClienteCredMorosos(Me.grdCliente.TextMatrix(k, 1)) Then
                    lsMsgCred = "Cliente posee pagos de Creditos Pendientes..."
                End If
            End If
        Next k
        Set clsCapD = Nothing
    Else
        lsMsgCred = ""
    End If
    '---------------------------------------------------------

    If MsgBox(lsMsgCred & "¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Dim sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
        Dim nSaldo As Double, nPorcDisp As Double
        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
        Dim nMontoLavDinero As Double, nTC As Double

        'Realiza la Validación para el Lavado de Dinero
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
            Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
            If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
                Set clsExo = Nothing
                sPersLavDinero = ""
                nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
                Set clsLav = Nothing

                If nMoneda = gMonedaNacional Then
                    Dim clsTC As COMDConstSistema.NCOMTipoCambio 'nTipoCambio
                    Set clsTC = New COMDConstSistema.NCOMTipoCambio
                    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set clsTC = Nothing
                Else
                    nTC = 1
                End If
                If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                    'By Capi 1402208
                    Call IniciaLavDinero(loLavDinero)
                    'ALPA
                    'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, sTipoCuenta, , , , , nmoneda)
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, sTipoCuenta, , , , , nMoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                    'End
                End If
            Else
                Set clsExo = Nothing
            End If
        'Else
        '    Set clsLav = Nothing
        'End If
        Set clsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set clsMov = Nothing
    '    On Error GoTo ErrGraba
        Select Case nProducto
            Case gCapAhorros
                If bDocumento Then
                    If nDocumento = TpoDocOrdenPago Then
                        If nOperacion = gAhoRetOPCert Or nOperacion = gAhoRetOPCertCanje Then
                            nSaldo = clsCap.CapCargoAhoOPCertifcada(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, sNroDoc, sCodIF, , , gsNomAge, sLpt, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, lsBoleta, , , gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                        Else
                            nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, sCodIF, bOPExiste, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , , , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                        End If
                    ElseIf nDocumento = TpoDocNotaCargo Then
                        nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, sCodIF, bOPExiste, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , , , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                    End If
                Else
                    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    If Me.cboMonedaBanco.Text <> "" Then
                        nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , Right(Me.txtBanco.Text, 13), , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , Me.txtCtaBanco.Text, CInt(Right(Me.cboMonedaBanco.Text, 3)), gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                    Else
                        nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , Right(Me.txtBanco.Text, 13), , , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sReaPersLavDinero, , Me.txtCtaBanco.Text, , gsCodCMAC, , gsCodAge, lbProcesaAutorizacion, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), sOperacion, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, nComixRet, nComixMov, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                    End If
                    'If nSaldo = -69 Then frmCapAutorizacion.Inicio
                End If

                '-------------------------------------------------
            
                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"

            Case gCapCTS
                If nOperacion = gCTSRetEfec Or nOperacion = gCMACOACTSRetEfec Then
                    nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                ElseIf nOperacion = gCTSRetTransf Then
                    nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, Right(Me.txtBanco.Text, 13), Me.txtCtaBanco.Text, CInt(Right(Me.cboMonedaBanco.Text, 3)), 0, 0, 0, 0, 0, lsmensaje, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
    '            ElseIf nOperacion = "220303" Then
    '                nMonto = CDbl(Me.TxtMonto2.Text)
    '                nSaldo = clsCap.CapCargoCuentaCTS(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, nDocumento, sNroDoc, , , sPersCodCMAC, sNombreCMAC, gsNomAge, sLpt, sPersLavDinero, sReaPersLavDinero, , , , CDbl(Val(Me.lblSalIntaD.Caption)), CDbl(Val(Me.lblIntIntaD.Caption)), CDbl(Val(Me.lblIntDisD.Caption)))
                End If
                
                If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
                If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
                If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
        End Select
        If gnMovNro > 0 Then
            Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
        End If
        Set clsLav = Nothing
        Set clsCap = Nothing
        Set loLavDinero = Nothing
        gVarPublicas.LimpiaVarLavDinero
        cmdCancelar_Click
    End If
    Exit Sub
ErrGraba:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub cmdMostrarFirma_Click()
With grdCliente
    If .TextMatrix(.Row, 1) = "" Then Exit Sub
    Call frmPersonaFirma.Inicio(Trim(.TextMatrix(.Row, 1)), Trim(txtCuenta.Age), False)
End With
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

'Cargos
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim sNumTar As String
    Dim sClaveTar As String
    Dim nErr As Integer
    Dim sCaption As String
    Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
    Dim nEstado As COMDConstantes.CaptacTarjetaEstado
    Dim nCOM As Integer
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    sCaption = Me.Caption
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        Set clsGen = New COMDConstSistema.DCOMGeneral
            bRetSinTarjeta = clsGen.GetPermisoEspecialUsuario(gCapPermEspRetSinTarj, gsCodUser, gsDominio)
        Set clsGen = Nothing
        sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta)
        If Val(Mid(sCuenta, 6, 3)) <> nProducto Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    
If KeyCode = vbKeyF11 And txtCuenta.Enabled = True Then 'F11
'        Dim nPuerto As TipoPuertoSerial
        Dim sMaquina As String
        sMaquina = GetComputerName
        sCaption = Me.Caption
        Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
        MsgBox "Pase la tarjeta.", vbInformation, "AVISO"
        'ppoa Modificacion
        'sNumTar = GetNumTarjeta
        sNumTar = ""
        sNumTar = GetNumTarjeta_ACS
        
        If Len(sNumTar) <> 16 Then
            MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
            Me.Caption = sCaption
            'Set clsGen = Nothing
            Exit Sub
        End If
        
        Me.Caption = "Ingrese la Clave de la Tarjeta."
        MsgBox "Ingrese la Clave de la Tarjeta.", vbInformation, "AVISO"
        'ppoa Modificacion
        'sClaveTar = GetClaveTarjeta
        Select Case GetClaveTarjeta_ACS(sNumTar, 1)
            Case gClaveValida
                    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                    Dim rsTarj As ADODB.Recordset
                    Set rsTarj = New ADODB.Recordset
                    Dim ObjTarj As COMNCaptaServicios.NCOMCaptaTarjeta
                    Set ObjTarj = New COMNCaptaServicios.NCOMCaptaTarjeta
                    Set rsTarj = ObjTarj.Get_Datos_Tarj(sNumTar)
                    
                    If rsTarj.EOF And rsTarj.BOF Then
                        MsgBox "Tarjeta no posee ninguna relación con cuentas de activas   o Tarjeta no activa.", vbInformation, "Aviso"
                        Me.Caption = sCaption
                        Set ObjTarj = Nothing
                        Exit Sub
                    Else
                        nEstado = rsTarj("nEstado")
                        If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
                            If nEstado = gCapTarjEstBloqueada Then
                                MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                            ElseIf nEstado = gCapTarjEstCancelada Then
                                MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                            End If
                            Set ObjTarj = Nothing
                            Me.Caption = sCaption
                            Exit Sub
                        End If
                        Dim rsPers As New ADODB.Recordset
                        Dim sCta As String, sProducto As String, sMoneda As String
                        Dim clsCuenta As UCapCuenta
                        Set rsPers = clsMant.GetCuentasPersona(rsTarj("cPersCod"), nProducto)
                        Set clsMant = Nothing
                        If Not (rsPers.EOF And rsPers.EOF) Then
                            Do While Not rsPers.EOF
                                sCta = rsPers("cCtaCod")
                                sProducto = rsPers("cDescripcion")
                                sMoneda = Trim(rsPers("cMoneda"))
                                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sProducto & Space(2) & sMoneda
                                rsPers.MoveNext
                            Loop
                            Set clsCuenta = New UCapCuenta
                            Set clsCuenta = frmCapMantenimientoCtas.Inicia
                            If clsCuenta.sCtaCod <> "" Then
                                txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                                txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
                                txtCuenta.cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                                txtCuenta.SetFocusCuenta
                                Call txtCuenta_KeyPress(13)
                                'SendKeys "{Enter}"
                            End If
                            Set clsCuenta = Nothing
                        Else
                            MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
                        End If
                        rsPers.Close
                        Set rsPers = Nothing
                    End If
                    Set rsTarj = Nothing
                    Set clsMant = Nothing
                              
            Case gClaveNOValida
                MsgBox "Clave Incorrecta", vbInformation, "Aviso"
            Case Else
            
        End Select
        Me.Caption = "Captaciones - Cargo - Ahorros " & sOperacion
    End If


    '**DAOR 20081125, Para tarjetas ***********************
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion))
        If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    '*******************************************************
End Sub
Public Function ObtieneDatosTarjeta(ByVal psCodTarj As String) As Boolean
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsTarj As New ADODB.Recordset
Dim sTarjeta As String, sPersona As String
Dim nEstado As COMDConstantes.CaptacTarjetaEstado

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsTarj = clsMant.GetTarjetaCuentas(psCodTarj)
If rsTarj.EOF And rsTarj.BOF Then
    MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
    ObtieneDatosTarjeta = False
Else
    ObtieneDatosTarjeta = True
End If
Set rsTarj = Nothing
Set clsMant = Nothing
End Function

Private Sub grdCliente_DblClick()
Dim sPersCod As String

sPersCod = grdCliente.TextMatrix(grdCliente.Row, 1)
If sPersCod = "" Then Exit Sub
MuestraFirma sPersCod, gsCodAge
End Sub

Private Sub txtBanco_EmiteDatos()
lblBanco = Trim(txtBanco.psDescripcion)
If lblBanco <> "" Then
    If Me.cboMonedaBanco.Visible And Me.cboMonedaBanco.Enabled Then
        cboMonedaBanco.SetFocus
    Else
        txtGlosa.SetFocus
    End If
End If
End Sub

Private Sub txtCtaBanco_GotFocus()
    txtCtaBanco.SelStart = 0
    txtCtaBanco.SelLength = 50
End Sub

Private Sub txtCtaBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtBanco.Enabled And txtBanco.Visible Then Me.txtBanco.SetFocus
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    txtMonto.Enabled = True
    txtMonto.SetFocus
End If
End Sub

Private Sub txtMonto_Change()
    If gbITFAplica And nProducto <> gCapCTS Then       'Filtra para CTS
        If txtMonto.value > gnITFMontoMin Then
            If Not lbITFCtaExonerada Then
                'If nOperacion = gAhoRetTransf Or nOperacion = gAhoDepTransf Or nOperacion = gAhoDepPlanRRHH Or nOperacion = gAhoDepPlanRRHHAdelantoSueldos Then
                If nOperacion = gAhoDepPlanRRHH Or nOperacion = gAhoDepPlanRRHHAdelantoSueldos Then
                    Me.lblITF.Caption = Format(0, "#,##0.00")
                ElseIf nProducto = gCapAhorros And (nOperacion <> gAhoDepChq Or nOperacion <> gCMACOAAhoDepChq Or nOperacion <> gCMACOTAhoDepChq) Then
                    Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                Else
                    Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                End If
            Else
                Me.lblITF.Caption = "0.00"
            End If
            If nOperacion = gAhoRetOPCanje Or nOperacion = gAhoRetOPCertCanje Or nOperacion = gAhoRetFondoFijoCanje Then
                Me.lblTotal.Caption = Format(0, "#,##0.00")
            ElseIf nOperacion = gAhoDepChq Then
                Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption), "#,##0.00")
            Else
                If nProducto = gCapAhorros And gbITFAsumidoAho Then
                    Me.lblITF.Caption = "0.00"
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                    Me.lblITF.Caption = "0.00"
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                Else
                    If chkITFEfectivo.value = 1 Then
                        Me.lblTotal.Caption = Format(txtMonto.value + CDbl(Me.lblITF.Caption), "#,##0.00")
                    Else
                        Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                    End If
                End If
            End If
        End If
    Else
        Me.lblITF.Caption = Format(0, "#,##0.00")
        
        If nProducto = gCTSDepChq Then
            Me.lblTotal.Caption = Format(0, "#,##0.00")
        Else
            Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
        End If
    End If
    
    If txtMonto.value = 0 Then
        Me.lblITF.Caption = "0.00"
        Me.lblTotal.Caption = "0.00"
    End If
    
    chkITFEfectivo_Click
End Sub

Private Sub chkITFEfectivo_Click()
Dim nMonto As Double
nMonto = txtMonto.value
If chkITFEfectivo.value = 1 Then
    lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
Else
    If nProducto = gCapAhorros And gbITFAsumidoAho Then
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    Else
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    End If
End If
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
txtMonto.SelStart = 0
txtMonto.SelLength = Len(txtMonto.Text)
End Sub


Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And CDbl(txtMonto.Text) > 0 Then
      cmdGrabar.Enabled = True
      Call cmdGrabar_Click
End If
End Sub

Private Sub txtOrdenPago_GotFocus()
With txtOrdenPago
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtOrdenPago_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Trim(txtOrdenPago.Text) <> "" And IsNumeric(txtOrdenPago) Then
        Dim oMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
        Dim rsOP As New ADODB.Recordset
        Dim nEstadoOP As COMDConstantes.CaptacOrdPagoEstado
        Dim sCuenta As String
        sCuenta = txtCuenta.NroCuenta
        Set oMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsOP = oMant.GetDatosOrdenPago(sCuenta, CLng(Val(Trim(txtOrdenPago.Text))))
        Set oMant = Nothing
        If Not (rsOP.EOF And rsOP.BOF) Then
            nEstadoOP = rsOP("nEstado")
            If rsOP("nEstado") = gCapOPEstCertifiCada Then
                txtMonto.Text = Format$(rsOP("nMonto"), "#,##0.00")
            End If
        End If
        rsOP.Close
        Set rsOP = Nothing
    Else
        MsgBox "Debe ingresar un número válido para la Orden de Pago", vbOKOnly + vbExclamation, App.Title
        Exit Sub
    End If
    If txtBanco.Visible Then
        txtBanco.SetFocus
    Else
        txtGlosa.SetFocus
    End If
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub


Private Sub MuestraFirmas(ByVal sCuenta As String)
Dim i As Integer
Dim sPersona As String
    
If nPersoneria <> PersPersoneria.gPersonaNat Then
    For i = 1 To Me.grdCliente.Rows - 1
        If Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersRepSuplente Or Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersRepTitular Then
            sPersona = grdCliente.TextMatrix(i, 1)
            MuestraFirma sPersona, gsCodAge
        End If
    Next i
Else
    For i = 1 To Me.grdCliente.Rows - 1
        If Trim(Right(grdCliente.TextMatrix(i, 3), 5)) = gCapRelPersTitular Then
            sPersona = grdCliente.TextMatrix(i, 1)
            MuestraFirma sPersona, gsCodAge
        End If
    Next i
End If
End Sub

Private Function Cargousu(ByVal NomUser As String) As String
Dim rs As New ADODB.Recordset
Dim oCons As COMDConstSistema.DCOMUAcceso
Set oCons = New COMDConstSistema.DCOMUAcceso
 
Set rs = oCons.Cargousu(NomUser)
If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
End If
rs.Close
Set rs = Nothing
Set oCons = Nothing
End Function
 
Private Function VerificarAutorizacion() As Boolean
Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim oCapAutN  As COMNCaptaGenerales.NCOMCaptAutorizacion
Dim oPers As COMDPersona.UCOMAcceso
Dim rs As New ADODB.Recordset
Dim lnMonTopD As Double
Dim lnMonTopS As Double
Dim lsmensaje As String
Dim gsGrupo As String
Dim sCuenta As String, sNivel As String
Dim lbEstadoApr As Boolean
Dim nMonto As Double
Dim nMoneda As Moneda

sCuenta = txtCuenta.NroCuenta
nMonto = txtMonto.value
nMoneda = CLng(Mid(sCuenta, 9, 1))
'Obtiene los grupos al cual pertenece el usuario
Set oPers = New COMDPersona.UCOMAcceso
    gsGrupo = oPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
Set oPers = Nothing
 
'Verificar Montos
Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    Set rs = oCapAut.ObtenerMontoTopNivAutRetCan(gsGrupo, "1", gsCodAge)
Set oCapAut = Nothing
 
If Not (rs.EOF And rs.BOF) Then
    lnMonTopD = rs("nTopDol")
    lnMonTopS = rs("nTopSol")
    sNivel = rs("cNivCod")
Else
    MsgBox "Usuario no Autorizado para realizar Operacion", vbInformation, "Aviso"
    VerificarAutorizacion = False
    Exit Function
End If

If nMoneda = gMonedaNacional Then
    If nMonto <= lnMonTopS Then
        VerificarAutorizacion = True
        Exit Function
    End If
Else
    If nMonto <= lnMonTopD Then
        VerificarAutorizacion = True
        Exit Function
    End If
End If
   
Set oCapAutN = New COMNCaptaGenerales.NCOMCaptAutorizacion
If sMovNroAut = "" Then 'Si es nueva, registra
    oCapAutN.NuevaSolicitudAutorizacion sCuenta, "1", nMonto, gdFecSis, gsCodAge, gsCodUser, nMoneda, gOpeAutorizacionRetiro, sNivel, sMovNroAut
    MsgBox "Solicitud Registrada, comunique a su Admnistrador para la Aprobación..." & Chr$(10) & _
        " No salir de esta operación mientras se realice el proceso..." & Chr$(10) & _
        " Porque sino se procedera a grabar otra Solicitud...", vbInformation, "Aviso"
    VerificarAutorizacion = False
Else
    'Valida el estado de la Solicitud
    If Not oCapAutN.VerificarAutorizacion(sCuenta, "1", nMonto, sMovNroAut, lsmensaje) Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        VerificarAutorizacion = False
    Else
        VerificarAutorizacion = True
    End If
End If
Set oCapAutN = Nothing
End Function

'******                 CALCULAR COMISION  AVMM-03-2006                *****
Public Function CalcularComisionRetOtraAge() As Double
    Dim nComixRet As Double
    Dim nPorComixRet As Double
    Dim nMontoLimiRetAge As Double
    Dim oCons As COMDConstantes.DCOMAgencias
    Dim clsGen As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set oCons = New COMDConstantes.DCOMAgencias
    
        '*** Verificar Ubicacion de la Cuenta ***
            
        If oCons.VerficaZonaAgencia(gsCodAge, Mid(txtCuenta.NroCuenta, 4, 2)) Then
        
            Set clsGen = New COMNCaptaGenerales.NCOMCaptaDefinicion
                If nMoneda = gMonedaNacional Then
                    nMontoLimiRetAge = clsGen.GetCapParametro(gMontoLimiMNRetOtraAge)
                Else
                    nMontoLimiRetAge = clsGen.GetCapParametro(gMontoLimiMERetOtraAge)
                End If
                ' *** Verificar Monto de Retiro en Otra Agencia ***
                If CDbl(txtMonto.Text) >= nMontoLimiRetAge Then
                ' *** Calcular Comisión ***
                    nPorComixRet = clsGen.GetCapParametro(gComisionRetOtraAge)
                    nComixRet = CDbl(txtMonto.Text) * (Val(nPorComixRet) / 100)
                Else
                    nComixRet = 0
                End If
                
            Set clsGen = Nothing
        Else
            nComixRet = 0
        End If
    Set oCons = Nothing
    CalcularComisionRetOtraAge = nComixRet
End Function

'******   VERIFICAR MAX MOVIMIENTOS Y CALCULAR COMISION  AVMM 03-06-2006  *****
Public Function CalcularComisionxMaxOpeRet() As Double
    Dim nNroMaxOpe As Double
    Dim nNroOpeRet As Double
    Dim nMontoxMaxOpe As Double
    Dim nMontoTope As Double
    Dim sFecha As String
    Dim oCapG As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim clsGen As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set oCapG = New COMDCaptaGenerales.DCOMCaptaGenerales
    
        Set clsGen = New COMNCaptaGenerales.NCOMCaptaDefinicion
            nNroMaxOpe = clsGen.GetCapParametro(gNroMaxRet)
            ' *** Obtener Nro de Operaciones ***
            sFecha = Mid(Format(gdFecSis, "yyyymmdd"), 1, 6)
            nNroOpeRet = oCapG.ObtenerNroMovimientosRet(txtCuenta.NroCuenta, Trim(sFecha))
            ' *** Verificar Nro de Operaciones ***
            If nNroOpeRet >= nNroMaxOpe Then
            ' *** Calcular Monto de Comisión ***
                 If nMoneda = gMonedaNacional Then
                     nMontoTope = clsGen.GetCapParametro(2097)
                     If nSaldoCuenta <= nMontoTope Then
                        nMontoxMaxOpe = clsGen.GetCapParametro(gMontoMNx31Ope)
                     Else
                        nMontoxMaxOpe = 0
                     End If
                 Else
                     nMontoTope = clsGen.GetCapParametro(2098)
                     If nSaldoCuenta <= nMontoTope Then
                        nMontoxMaxOpe = clsGen.GetCapParametro(gMontoMEx31Ope)
                     Else
                        nMontoxMaxOpe = 0
                     End If
                 End If
            Else
                nMontoxMaxOpe = 0
            End If
            
        Set clsGen = Nothing
        
    Set oCapG = Nothing
    CalcularComisionxMaxOpeRet = nMontoxMaxOpe
End Function

Sub Finaliza_Verifone5000()
        If Not GmyPSerial Is Nothing Then
            GmyPSerial.Disconnect
            Set GmyPSerial = Nothing
        End If
End Sub


