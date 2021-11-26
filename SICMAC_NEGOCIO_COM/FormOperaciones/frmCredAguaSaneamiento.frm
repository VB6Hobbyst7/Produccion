VERSION 5.00
Begin VB.Form frmCredAguaSaneamiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Agua y Saneamiento"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "frmCredAguaSaneamiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Datos de Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3470
      Left            =   0
      TabIndex        =   0
      Top             =   40
      Width           =   9000
      Begin VB.CommandButton btnSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   7680
         TabIndex        =   6
         Top             =   3020
         Width           =   1000
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   2475
         TabIndex        =   5
         ToolTipText     =   "Eliminar"
         Top             =   3020
         Width           =   1000
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   6480
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   3020
         Width           =   1000
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Nuevo"
         Height          =   345
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Nuevo"
         Top             =   3020
         Width           =   1000
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Editar"
         Height          =   345
         Left            =   1160
         TabIndex        =   1
         ToolTipText     =   "Editar"
         Top             =   3020
         Width           =   1000
      End
      Begin SICMACT.FlexEdit feAguaSaneamiento 
         Height          =   2730
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4815
         Cols0           =   6
         HighLight       =   2
         EncabezadosNombres=   "N°-Destino Principal-N° Crédito-Sub Destino-Monto Sub Destino-"
         EncabezadosAnchos=   "300-2000-1800-2750-1800-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C-L"
         FormatosEdit    =   "0-0-0-0-0-0"
         CantEntero      =   12
         TextArray0      =   "N°"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmCredAguaSaneamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre : frmCredAguaSaneamiento
'** Descripción : Paraincluir el destino de agua y saneamiento para discrimar que porcentaje
'** del monto solicitado apunta hacia un destino determinado creado segun TI-ERS052-2018
'** Creación : EAAS, 20180727 04:30:00 PM
'************************************************************************************************
Option Explicit

Dim fvListaAguaSaneamiento() As TAguaSaneamiento
Dim fnTpoProducto As Integer
Dim sDestCred As String
Dim dMontoSol As Double
Dim dMontoSolT As Double
Dim sNroCredito As String
Dim sCtaCod As String
Dim nCentinela As Integer 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nSumaAguaSaneamiento As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nMontoSubDestinoEditar As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM



Private Sub btnSalir_Click()
'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
    Dim nSuma As Double
    nSuma = SumarCampo(feAguaSaneamiento, 4)
    nSumaAguaSaneamiento = nSuma
'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    cmdQuitar.Enabled = False
    cmdModificar.Enabled = False
    
    If UBound(fvListaAguaSaneamiento) > 0 Then
        For i = 1 To UBound(fvListaAguaSaneamiento)
            Call AdicionaFila(fvListaAguaSaneamiento(i))
        Next

        cmdQuitar.Enabled = True
        cmdModificar.Enabled = True
    End If
End Sub
Public Sub Inicio(ByRef pvListaAguaSaneamiento() As TAguaSaneamiento, ByVal psnTpoProducto As Integer, ByVal pDestCred As String, ByRef pMontoSol As Double, ByVal psCtaCod As String, ByRef pnCentinela As Integer, ByRef pnSumaAguaSaneamiento As Double) 'EAAS20191401 SEGUN 018-GM-DI_CMACM pnCentinela pnSumaAguaSaneamiento
    fvListaAguaSaneamiento = pvListaAguaSaneamiento
    fnTpoProducto = psnTpoProducto
    sDestCred = pDestCred
    dMontoSol = pMontoSol
    dMontoSolT = dMontoSol
    sCtaCod = psCtaCod
    nSumaAguaSaneamiento = pnSumaAguaSaneamiento 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    nCentinela = pnCentinela 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    Show 1
    pvListaAguaSaneamiento = fvListaAguaSaneamiento
    pnSumaAguaSaneamiento = nSumaAguaSaneamiento 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    
End Sub
Private Sub cmdAgregar_Click()
    Dim frm As New frmCredAguaSaneamientoDet
    Dim lvAguaSaneamiento As TAguaSaneamiento
    Dim lvTemp() As TAguaSaneamiento
    Dim bOK As Boolean
    Dim Index As Integer
    Dim ixCD As Integer
    Dim nSuma As Double 'joep
    Dim nbeneficiarios As Integer
    lvTemp = fvListaAguaSaneamiento 'Temporal para no modificar el actual array
    
    If (UBound(fvListaAguaSaneamiento)) Then
        For ixCD = 1 To 1
            nbeneficiarios = fvListaAguaSaneamiento(ixCD).nBeneficia
        Next
    End If
    
    nSuma = SumarCampo(feAguaSaneamiento, 4)
    dMontoSol = dMontoSolT 'EAAS20191401 SEGUN 018-GM-DI_CMACM se quito - nSuma
    bOK = frm.Registrar(lvAguaSaneamiento, lvTemp, fnTpoProducto, sDestCred, dMontoSol, sCtaCod, nbeneficiarios, nCentinela) 'EAAS20191401 SEGUN 018-GM-DI_CMACM nCentinela
    dMontoSolT = dMontoSol 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    If bOK Then
        Index = UBound(fvListaAguaSaneamiento) + 1
        ReDim Preserve fvListaAguaSaneamiento(Index)
        fvListaAguaSaneamiento(Index) = lvAguaSaneamiento
        
        AdicionaFila lvAguaSaneamiento
        
        cmdQuitar.Enabled = True
        cmdModificar.Enabled = True
    End If
    Set frm = Nothing
End Sub
Private Sub CmdModificar_Click()
    Dim frm As frmCredAguaSaneamientoDet
    Dim lvAguaSaneamiento As TAguaSaneamiento
    Dim lvTemp() As TAguaSaneamiento
    Dim bOK As Boolean
    Dim Index As Integer
    Dim nSuma As Double
    If feAguaSaneamiento.TextMatrix(1, 0) = "" Then Exit Sub
    Index = feAguaSaneamiento.row
    lvTemp = fvListaAguaSaneamiento 'Temporal para no modificar el actual array
    lvAguaSaneamiento = fvListaAguaSaneamiento(Index)
    nMontoSubDestinoEditar = lvAguaSaneamiento.nMontoS 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    Set frm = New frmCredAguaSaneamientoDet
    nSuma = SumarCampo(feAguaSaneamiento, 4)
    dMontoSol = dMontoSolT
    bOK = frm.Modificar(lvAguaSaneamiento, Index, lvTemp, fnTpoProducto, dMontoSol, sCtaCod, nSuma, nMontoSubDestinoEditar, nCentinela) 'EAAS20191401 SEGUN 018-GM-DI_CMACM nMontoSubDestinoEditar
    fvListaAguaSaneamiento = lvTemp
    dMontoSolT = dMontoSol 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    If bOK Then
        fvListaAguaSaneamiento(Index) = lvAguaSaneamiento
        ModificaFila Index, lvAguaSaneamiento
    End If
    Set frm = Nothing
End Sub
Private Sub cmdQuitar_Click()
    Dim lvListaTemp() As TAguaSaneamiento
    Dim Index As Integer
    Dim i As Integer
    Dim j As Integer
        
    If feAguaSaneamiento.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = feAguaSaneamiento.row
    
    If MsgBox("Se va a eliminar el Sub Destino. ¿Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    ReDim lvListaTemp(0)
    
    For i = 1 To UBound(fvListaAguaSaneamiento)
        If i <> Index Then
            j = UBound(lvListaTemp) + 1
            ReDim Preserve lvListaTemp(j)
            lvListaTemp(j) = fvListaAguaSaneamiento(i)
        End If
    Next
    
    fvListaAguaSaneamiento = lvListaTemp
    feAguaSaneamiento.EliminaFila Index
    
    If UBound(fvListaAguaSaneamiento) = 0 Then
        cmdQuitar.Enabled = False
        cmdModificar.Enabled = False
    End If
End Sub

Private Sub AdicionaFila(ByRef pvListaAguaSaneamiento As TAguaSaneamiento)

    Dim i As Integer
    feAguaSaneamiento.AdicionaFila
    i = feAguaSaneamiento.row
    feAguaSaneamiento.TextMatrix(i, 1) = pvListaAguaSaneamiento.sDestinoDescripcion
    feAguaSaneamiento.TextMatrix(i, 2) = sCtaCod 'pvListaAguaSaneamiento.sNroCredito
    feAguaSaneamiento.TextMatrix(i, 3) = pvListaAguaSaneamiento.sSubDestinoDescripcion
    feAguaSaneamiento.TextMatrix(i, 4) = Format(pvListaAguaSaneamiento.nMontoS, "#,##0.00")
    feAguaSaneamiento.TextMatrix(i, 5) = pvListaAguaSaneamiento.nBeneficia
    feAguaSaneamiento.BackColorRow vbWhite
    


End Sub
Private Sub ModificaFila(ByVal pnIndex As Integer, ByRef pvAguaSaneamiento As TAguaSaneamiento)
    Dim i As Integer
    feAguaSaneamiento.TextMatrix(pnIndex, 1) = pvAguaSaneamiento.sDestinoDescripcion
    feAguaSaneamiento.TextMatrix(pnIndex, 2) = pvAguaSaneamiento.sNroCredito
    feAguaSaneamiento.TextMatrix(pnIndex, 3) = pvAguaSaneamiento.sSubDestinoDescripcion
    'feAguaSaneamiento.TextMatrix(pnIndex, 4) = pvAguaSaneamiento.nMontoP
    feAguaSaneamiento.TextMatrix(pnIndex, 4) = Format(pvAguaSaneamiento.nMontoS, "#,##0.00")
    For i = 1 To feAguaSaneamiento.rows - 1
    feAguaSaneamiento.TextMatrix(i, 5) = pvAguaSaneamiento.nBeneficia
    Next
End Sub
Private Sub cmdsalir_Click()
    Dim nSuma As Double
    nSuma = SumarCampo(feAguaSaneamiento, 4)
    If (feAguaSaneamiento.TextMatrix(1, 1) <> "") Then
        If (UBound(fvListaAguaSaneamiento) = 0 And Right(sDestCred, 2) = 26) Then
            MsgBox "Debe ingresar el detalle de Agua y Saneamiento, no podrá grabar la solicitud", vbInformation, "Aviso"
        End If
    
        If (dMontoSolT <> 0 And Right(sDestCred, 2) = 26) Then 'EAAS20190410 SEGUN 018-GM-DI_CMACM
            MsgBox "La suma de los subdestinos de agua y saneamiento debe ser igual al monto solicitado", vbInformation, "Aviso"
            Exit Sub
        End If
    
        If nSuma > 8000 Then
        MsgBox "El monto total de Agua y Saneamiento no debe superar más de S/ 8000.00", vbInformation, "Aviso"
        Else
        nSumaAguaSaneamiento = nSuma 'EAAS20191401 SEGUN 018-GM-DI_CMACM
            Unload Me
        End If
    
    Else
    nSumaAguaSaneamiento = nSuma 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    Unload Me
    End If
End Sub

