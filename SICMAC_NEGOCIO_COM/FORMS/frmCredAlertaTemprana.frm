VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredAlertaTemprana 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indicadores de Alertas Tempranas"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredAlertaTemprana.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstabIndicador 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Indicadores"
      TabPicture(0)   =   "frmCredAlertaTemprana.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feAlertas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin SICMACT.FlexEdit feAlertas 
         Height          =   2055
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   3625
         Cols0           =   11
         HighLight       =   1
         EncabezadosNombres=   "-Indicadores-Detalle-cFormula-cDetalleValor-nValor-nValorLimite-nEstado-cUnidad-cLeyenda-Aux"
         EncabezadosAnchos=   "0-5000-1250-0-0-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-1-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-C-C-C-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-3-0-0-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmCredAlertaTemprana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'** Nombre : frmCredAlertaTemprana                                                        ****
'** Descripción : Muestra las alertas tempranas por crédito                               ****
'** Referencia : ERS001-2017 - Metodología para monitoreo de Alertas Tempranas de Crédito ****
'** Creación : LUCV, 20170131 13:51:01 PM                                                 ****
'*********************************************************************************************
Option Explicit
Dim ofrmAlertaDet As frmCredAlertaTempranaDet
Dim fsTitulo As String
Dim fsFormula As String
Dim fsDetalle As String
Dim fsLimite As Integer
Dim fsCtaCod As String

Private Sub Form_Load()
    fsFormula = "Formula"
    fsDetalle = "detalle"
    fsLimite = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
    If KeyCode = 113 And Shift = 0 Then
        KeyCode = 10
    End If
    If KeyCode = 27 And Shift = 0 Then
        Unload Me
    End If
End Sub
Private Sub feAlertas_EnterCell()
    If feAlertas.Col = 2 Or (feAlertas.Col = 2 And feAlertas.row = 1) Then
        feAlertas.AvanceCeldas = Vertical
         Me.feAlertas.ColumnasAEditar = "X-X-2-X-X-X-X-X-X-X-X"
         Me.feAlertas.ListaControles = "0-0-1-0-0-0-0-0-0-0-0"
    Else
        feAlertas.AvanceCeldas = Horizontal
         Me.feAlertas.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X"
    End If
End Sub
Private Sub feAlertas_Click()
    Set ofrmAlertaDet = New frmCredAlertaTempranaDet
    If feAlertas.Col = 2 Then
        If Not ofrmAlertaDet.Inicio(feAlertas.TextMatrix(feAlertas.row, 1), feAlertas.TextMatrix(feAlertas.row, 3), feAlertas.TextMatrix(feAlertas.row, 4), feAlertas.TextMatrix(feAlertas.row, 5), feAlertas.TextMatrix(feAlertas.row, 6), feAlertas.TextMatrix(feAlertas.row, 7), feAlertas.TextMatrix(feAlertas.row, 8), feAlertas.TextMatrix(feAlertas.row, 9)) Then
            MsgBox "No se puede visualizar el detalle del ratio, favor coordinar con el area correspondiente", vbInformation, "Aviso"
        End If
    End If
End Sub
Private Sub feAlertas_DblClick()
    Set ofrmAlertaDet = New frmCredAlertaTempranaDet
    If Not ofrmAlertaDet.Inicio(feAlertas.TextMatrix(feAlertas.row, 1), feAlertas.TextMatrix(feAlertas.row, 3), feAlertas.TextMatrix(feAlertas.row, 4), feAlertas.TextMatrix(feAlertas.row, 5), feAlertas.TextMatrix(feAlertas.row, 6), feAlertas.TextMatrix(feAlertas.row, 7), feAlertas.TextMatrix(feAlertas.row, 8), feAlertas.TextMatrix(feAlertas.row, 9)) Then
        MsgBox "No se puede visualizar el detalle del ratio, favor coordinar con el area correspondiente", vbInformation, "Aviso"
    End If
    Set ofrmAlertaDet = Nothing
End Sub
Private Sub feAlertas_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Set ofrmAlertaDet = New frmCredAlertaTempranaDet
    psDescripcion = ""
    psCodigo = 0
    psDescripcion = feAlertas.TextMatrix(feAlertas.row, 1) 'Detalle
    psCodigo = feAlertas.TextMatrix(feAlertas.row, 2) 'Formula
    If Not ofrmAlertaDet.Inicio(feAlertas.TextMatrix(feAlertas.row, 1), feAlertas.TextMatrix(feAlertas.row, 3), feAlertas.TextMatrix(feAlertas.row, 4), feAlertas.TextMatrix(feAlertas.row, 5), feAlertas.TextMatrix(feAlertas.row, 6), feAlertas.TextMatrix(feAlertas.row, 7), feAlertas.TextMatrix(feAlertas.row, 8), feAlertas.TextMatrix(feAlertas.row, 9)) Then
        MsgBox "No se puede visualizar el detalle del ratio, favor coordinar con el area correspondiente", vbInformation, "Aviso"
    End If
    Set ofrmAlertaDet = Nothing
End Sub
Public Sub Inicio(ByVal psCtaCod As String)
    fsCtaCod = psCtaCod
    If Not CargarDatos(fsCtaCod) Then
        Exit Sub
    End If
    Show 1
End Sub
Public Function CargarDatos(ByVal psCtaCod As String) As Boolean
On Error GoTo ErrorCargaDatos
    Call CargaFlexAlertasTempranas(psCtaCod)
    CargarDatos = True
    Exit Function
ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Private Sub CargaFlexAlertasTempranas(ByVal pcCtaCod As String)
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    Dim rsDatosAlerta As New ADODB.Recordset
    Set rsDatosAlerta = oDCOMFormatosEval.ObtieneListadoAlertasTempranas(pcCtaCod)
    Dim lnFila As Integer
        Call LimpiaFlex(feAlertas)
            Do While Not rsDatosAlerta.EOF
                feAlertas.AdicionaFila
                lnFila = feAlertas.row
                feAlertas.TextMatrix(lnFila, 1) = UCase(rsDatosAlerta!cAlerta)
                feAlertas.TextMatrix(lnFila, 2) = "..."
                feAlertas.TextMatrix(lnFila, 3) = rsDatosAlerta!cFormula
                feAlertas.TextMatrix(lnFila, 4) = rsDatosAlerta!cDetalle
                feAlertas.TextMatrix(lnFila, 5) = rsDatosAlerta!nValor
                feAlertas.TextMatrix(lnFila, 6) = rsDatosAlerta!nValorLimite
                feAlertas.TextMatrix(lnFila, 7) = rsDatosAlerta!nEstado
                feAlertas.TextMatrix(lnFila, 8) = rsDatosAlerta!cUnidad
                feAlertas.TextMatrix(lnFila, 9) = rsDatosAlerta!cLeyenda
                
                   If rsDatosAlerta!nEstado = 0 Then
                        feAlertas.BackColorRow (&HC0FFC0)
                    Else
                        feAlertas.BackColorRow (&HC0C0FF)
                    End If
                
                rsDatosAlerta.MoveNext
            Loop
        rsDatosAlerta.Close
        Set rsDatosAlerta = Nothing
End Sub

