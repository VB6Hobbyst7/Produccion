VERSION 5.00
Begin VB.Form frmCredVerde 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Crédito Verde"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "frmCredVerde.frx":0000
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
Attribute VB_Name = "frmCredVerde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre : frmCredAguaSaneamiento
'** Descripción : Para incluir el destino de credito verde para discrimar que porcentaje
'** del monto solicitado apunta hacia un destino determinado creado segun EAAS20191401 SEGUN 018-GM-DI_CMACM
'** Creación : EAAS, 20191401 04:30:00 PM
'************************************************************************************************
Option Explicit
Dim fvListaCreditoVerde() As TCreditoVerde
Dim fnTpoProducto As Integer
Dim sDestCred As String
Dim dMontoSol As Double
Dim dMontoSolT As Double
Dim sNroCredito As String
Dim sCtaCod As String
Dim nSumaCreditoVerde As Double
Dim nMontoSubDestinoEditar As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim lvCreditoVerdeFilaEliminada As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM



Private Sub btnSalir_Click()
cmdsalir_Click
Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    cmdQuitar.Enabled = False
    cmdModificar.Enabled = False
    
    If UBound(fvListaCreditoVerde) > 0 Then
        For i = 1 To UBound(fvListaCreditoVerde)
            Call AdicionaFila(fvListaCreditoVerde(i))
        Next

        cmdQuitar.Enabled = True
        cmdModificar.Enabled = True
    End If
End Sub
Public Sub Inicio(ByRef pvListaCreditoVerde() As TCreditoVerde, ByVal psnTpoProducto As Integer, ByVal pDestCred As String, ByRef pMontoSol As Double, ByVal psCtaCod As String, ByRef pnSumaCreditoVerde As Double)
    fvListaCreditoVerde = pvListaCreditoVerde
    fnTpoProducto = psnTpoProducto
    sDestCred = pDestCred
    dMontoSol = pMontoSol
    dMontoSolT = dMontoSol
    sCtaCod = psCtaCod
    Show 1
    pvListaCreditoVerde = fvListaCreditoVerde
    pMontoSol = dMontoSol
    pnSumaCreditoVerde = nSumaCreditoVerde
End Sub
Private Sub cmdAgregar_Click()
    Dim frm As New frmCredVerdeDet
    Dim lvCreditoVerde As TCreditoVerde
    Dim lvTemp() As TCreditoVerde
    Dim bOK As Boolean
    Dim Index As Integer
    Dim ixCD As Integer
    Dim nSuma As Double 'joep
    Dim nbeneficiarios As Integer
    lvTemp = fvListaCreditoVerde 'Temporal para no modificar el actual array
    
    If (UBound(fvListaCreditoVerde)) Then
        For ixCD = 1 To 1
            nbeneficiarios = fvListaCreditoVerde(ixCD).nBeneficia
        Next
    End If
    
    nSuma = SumarCampo(feAguaSaneamiento, 4)
    dMontoSol = dMontoSolT
    bOK = frm.Registrar(lvCreditoVerde, lvTemp, fnTpoProducto, sDestCred, dMontoSol, sCtaCod, nbeneficiarios)
    dMontoSolT = dMontoSol
    If bOK Then
        Index = UBound(fvListaCreditoVerde) + 1
        ReDim Preserve fvListaCreditoVerde(Index)
        fvListaCreditoVerde(Index) = lvCreditoVerde
        
        AdicionaFila lvCreditoVerde
        
        cmdQuitar.Enabled = True
        cmdModificar.Enabled = True
    End If
    Set frm = Nothing
End Sub
Private Sub CmdModificar_Click()
    Dim frm As frmCredVerdeDet
    Dim lvCreditoVerde As TCreditoVerde
    Dim lvTemp() As TCreditoVerde
    Dim bOK As Boolean
    Dim Index As Integer
    Dim nSuma As Double
    If feAguaSaneamiento.TextMatrix(1, 0) = "" Then Exit Sub
    Index = feAguaSaneamiento.row
    lvTemp = fvListaCreditoVerde 'Temporal para no modificar el actual array
    lvCreditoVerde = fvListaCreditoVerde(Index)
    nMontoSubDestinoEditar = lvCreditoVerde.nMontoS 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    Set frm = New frmCredVerdeDet
    nSuma = SumarCampo(feAguaSaneamiento, 4)
    dMontoSol = dMontoSolT
    bOK = frm.Modificar(lvCreditoVerde, Index, lvTemp, fnTpoProducto, dMontoSol, sCtaCod, nSuma, nMontoSubDestinoEditar) 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    fvListaCreditoVerde = lvTemp
    dMontoSolT = dMontoSol 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    If bOK Then
        fvListaCreditoVerde(Index) = lvCreditoVerde
        ModificaFila Index, lvCreditoVerde
    End If
    Set frm = Nothing
End Sub
Private Sub cmdQuitar_Click()
    Dim lvListaTemp() As TCreditoVerde
    Dim Index As Integer
    Dim i As Integer
    Dim j As Integer
        
    If feAguaSaneamiento.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = feAguaSaneamiento.row
    
    If MsgBox("Se va a eliminar el Sub Destino. ¿Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    ReDim lvListaTemp(0)
    
    For i = 1 To UBound(fvListaCreditoVerde)
        If i <> Index Then
            j = UBound(lvListaTemp) + 1
            ReDim Preserve lvListaTemp(j)
            lvListaTemp(j) = fvListaCreditoVerde(i)
        End If
    Next
    lvCreditoVerdeFilaEliminada = fvListaCreditoVerde(Index).nMontoS 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    fvListaCreditoVerde = lvListaTemp 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    dMontoSol = dMontoSol + lvCreditoVerdeFilaEliminada 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    dMontoSolT = dMontoSol 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    feAguaSaneamiento.EliminaFila Index
    
    If UBound(fvListaCreditoVerde) = 0 Then
        cmdQuitar.Enabled = False
        cmdModificar.Enabled = False
    End If
End Sub

Private Sub AdicionaFila(ByRef pvListaCreditoVerde As TCreditoVerde)

    Dim i As Integer
    feAguaSaneamiento.AdicionaFila
    i = feAguaSaneamiento.row
    feAguaSaneamiento.TextMatrix(i, 1) = pvListaCreditoVerde.sDestinoDescripcion
    feAguaSaneamiento.TextMatrix(i, 2) = sCtaCod 'pvListaAguaSaneamiento.sNroCredito
    feAguaSaneamiento.TextMatrix(i, 3) = pvListaCreditoVerde.sSubDestinoDescripcion
    feAguaSaneamiento.TextMatrix(i, 4) = Format(pvListaCreditoVerde.nMontoS, "#,##0.00")
    feAguaSaneamiento.TextMatrix(i, 5) = pvListaCreditoVerde.nBeneficia
    feAguaSaneamiento.BackColorRow vbWhite
    


End Sub
Private Sub ModificaFila(ByVal pnIndex As Integer, ByRef pvCreditoVerde As TCreditoVerde)
    Dim i As Integer
    feAguaSaneamiento.TextMatrix(pnIndex, 1) = pvCreditoVerde.sDestinoDescripcion
    feAguaSaneamiento.TextMatrix(pnIndex, 2) = pvCreditoVerde.sNroCredito
    feAguaSaneamiento.TextMatrix(pnIndex, 3) = pvCreditoVerde.sSubDestinoDescripcion
    'feAguaSaneamiento.TextMatrix(pnIndex, 4) = pvCreditoVerde.nMontoP
    feAguaSaneamiento.TextMatrix(pnIndex, 4) = Format(pvCreditoVerde.nMontoS, "#,##0.00")
    For i = 1 To feAguaSaneamiento.rows - 1
    feAguaSaneamiento.TextMatrix(i, 5) = pvCreditoVerde.nBeneficia
    Next
End Sub
Private Sub cmdsalir_Click()
    Dim nSuma As Double
    nSuma = SumarCampo(feAguaSaneamiento, 4)
'    If (feAguaSaneamiento.TextMatrix(1, 1) <> "") Then
'        If (UBound(fvListaCreditoVerde) = 0 And Right(sDestCred, 2) = 26) Then
'            MsgBox "Debe ingresar el detalle de Agua y Saneamiento, no podrá grabar la solicitud", vbInformation, "Aviso"
'        End If
'
'        If (nSuma <> dMontoSolT And Right(sDestCred, 2) = 26) Then
'            MsgBox "La suma de los subdestinos de EcoAhorro debe ser igual al monto solicitado", vbInformation, "Aviso"
'            Exit Sub
'        End If
'
'        If nSuma > 8000 Then
'        MsgBox "El monto total de Agua y Saneamiento no debe superar más de S/ 8000.00", vbInformation, "Aviso"
'        Else
'            Unload Me
'        End If
    
    'Else
    nSumaCreditoVerde = nSuma
    Unload Me
    'End If
End Sub

