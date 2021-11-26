VERSION 5.00
Begin VB.Form frmGarantiaCred 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Créditos asociados a la Garantía"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   Icon            =   "frmGarantiaCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRapiFlash 
      Caption         =   "Máximo disponible para Créditos RapiFlash"
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
      Height          =   540
      Left            =   80
      TabIndex        =   18
      Top             =   5295
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Con Campaña :"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   230
         Width           =   1215
      End
      Begin VB.Label lblDispRapiFlashCamp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Left            =   4560
         TabIndex        =   21
         Top             =   200
         Width           =   1605
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Sin Campaña :"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   230
         Width           =   1095
      End
      Begin VB.Label lblDispRapiFlash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Left            =   1440
         TabIndex        =   19
         Top             =   200
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de la Garantía "
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
      Height          =   1035
      Left            =   80
      TabIndex        =   3
      Top             =   0
      Width           =   11175
      Begin VB.Label lblTitular 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6840
         TabIndex        =   15
         Top             =   240
         Width           =   4235
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Titular :"
         Height          =   255
         Left            =   6120
         TabIndex        =   14
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblMoneda 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   13
         Top             =   620
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   650
         Width           =   765
      End
      Begin VB.Label lblValorGarantia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Left            =   9480
         TabIndex        =   11
         Top             =   620
         Width           =   1605
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Garantía :"
         Height          =   255
         Left            =   8160
         TabIndex        =   10
         Top             =   650
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Gravamen :"
         Height          =   255
         Left            =   4905
         TabIndex        =   8
         Top             =   650
         Width           =   855
      End
      Begin VB.Label lblDescripcion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblGravamen 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5880
         TabIndex        =   6
         Top             =   620
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "V.R.M.:"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   650
         Width           =   615
      End
      Begin VB.Label lblVRM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3000
         TabIndex        =   4
         Top             =   620
         Width           =   1605
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Créditos Asociados a la Garantía"
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
      Height          =   4215
      Left            =   80
      TabIndex        =   1
      Top             =   1080
      Width           =   11175
      Begin SICMACT.FlexEdit feCreditos 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5530
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nº Crédito-Titular-Mon.-Monto-Monto Garantía-Liberado-Ratio Cob.-Cobertura-Estado"
         EncabezadosAnchos=   "350-1700-2400-660-1200-1200-1200-850-1270-1500"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-R-R-R-R-R-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   6650
         X2              =   11120
         Y1              =   3765
         Y2              =   3765
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   6650
         X2              =   11120
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Label lblDisponible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Left            =   9480
         TabIndex        =   24
         Top             =   3840
         Width           =   1605
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Disponible para cálculo de Cobertura :"
         Height          =   255
         Left            =   6600
         TabIndex        =   23
         Top             =   3870
         Width           =   2775
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Left            =   9480
         TabIndex        =   17
         Top             =   3435
         Width           =   1605
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Cobertura Total :"
         Height          =   255
         Left            =   8160
         TabIndex        =   16
         Top             =   3495
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   350
      Left            =   10180
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   5385
      Width           =   1095
   End
End
Attribute VB_Name = "frmGarantiaCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmGarantiaCred
'** Descripción : Formulario que muestra la lista de créditos relacionados a la garantía
'** Creación    : EJVG, 20151011 - ERS063-2014
'**********************************************************************************************

Option Explicit

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Public Sub Inicio(ByVal psNumGarant As String)
    If CargarDatos(psNumGarant) Then
        Show 1
    End If
End Sub
Private Function CargarDatos(ByVal psNumGarant As String) As Boolean
    Dim obj As New COMDCredito.DCOMGarantia
    Dim rsGar As New ADODB.Recordset
    Dim rsCred As New ADODB.Recordset
    Dim nTotal As Double
    Dim nCobertura As Double
    Dim i As Integer
    
    On Error GoTo ErrCargar
    
    Screen.MousePointer = 11
    
    Set rsGar = obj.RecuperaGarantiaxVerCredito(psNumGarant)
    If Not (rsGar.EOF And rsGar.BOF) Then
        lblDescripcion.Caption = rsGar!cdescripcion
        lblTitular.Caption = rsGar!cPersNombre
        lblMoneda.Caption = rsGar!cMoneda
        lblVRM.Caption = Format(rsGar!nRealizacion, gsFormatoNumeroView)
        lblGravamen.Caption = Format(rsGar!nGravamen, gsFormatoNumeroView)
        lblValorGarantia.Caption = Format(rsGar!nPorGravar, gsFormatoNumeroView)
        lblDisponible.Caption = Format(Round(rsGar!nDisponible, 2), gsFormatoNumeroView)
        fraRapiFlash.Visible = rsGar!bAutoLiquidable
        
        If rsGar!bAutoLiquidable Then
            lblDispRapiFlash.Caption = Format(Round(rsGar!nDisponiblePF, 2), gsFormatoNumeroView)
            lblDispRapiFlashCamp.Caption = Format(Round(rsGar!nDisponiblePFCamp, 2), gsFormatoNumeroView)
        End If
        
        FormateaFlex feCreditos
        Set rsCred = obj.RecuperaCreditosGarantiaxVerCredito(psNumGarant)
        Do While Not rsCred.EOF
            feCreditos.AdicionaFila
            i = feCreditos.row
            feCreditos.TextMatrix(i, 1) = rsCred!cCtaCod
            feCreditos.TextMatrix(i, 2) = rsCred!cPersNombre
            feCreditos.TextMatrix(i, 3) = rsCred!cMoneda
            feCreditos.TextMatrix(i, 4) = Format(rsCred!nMontoCol, gsFormatoNumeroView)
            feCreditos.TextMatrix(i, 5) = Format(rsCred!nGravado, gsFormatoNumeroView)
            feCreditos.TextMatrix(i, 6) = Format(rsCred!nLiberado, gsFormatoNumeroView)
            feCreditos.TextMatrix(i, 7) = rsCred!nRatioCobertura
            nCobertura = IIf(rsCred!nGravado - rsCred!nLiberado < 0#, 0#, rsCred!nGravado - rsCred!nLiberado) * rsCred!nRatioCobertura
            feCreditos.TextMatrix(i, 8) = Format(nCobertura, "##,###,##0.00")
            nTotal = nTotal + nCobertura
            feCreditos.TextMatrix(i, 9) = rsCred!cEstado 'EJVG20160424
            rsCred.MoveNext
        Loop
        CargarDatos = True
    Else
        MsgBox "No se ha podido cargar los datos de la Garantía", vbInformation, "Aviso"
    End If
    Screen.MousePointer = 0
    
    lblTotal.Caption = Format(nTotal, "##,###,##0.00")
    
    RSClose rsGar
    RSClose rsCred
    Set obj = Nothing
    
    Exit Function
ErrCargar:
    Screen.MousePointer = 0
    CargarDatos = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

