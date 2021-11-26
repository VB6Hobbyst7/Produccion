VERSION 5.00
Begin VB.Form frmCredEndeuCuotaSistFinanc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuotas en el Sistema Financiero"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13020
   Icon            =   "frmCredEndeuCuotaSistFinanc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   13020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Datos Sistema Financiero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   12900
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   4800
         TabIndex        =   14
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   495
         Left            =   6600
         TabIndex        =   13
         Top             =   3120
         Width           =   1815
      End
      Begin SICMACT.FlexEdit FEDeudaSF 
         Height          =   2775
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   4895
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "cod_IFI-IFI-Dia de Pago-Cta.Cont.IFI-Deuda IFI-Cuota.Est.-interesDeve-Observación"
         EncabezadosAnchos=   "0-4200-1200-0-2000-1200-0-4000"
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-L-R-R-C-L"
         FormatosEdit    =   "0-0-3-0-4-4-0-0"
         TextArray0      =   "cod_IFI"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12900
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9240
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Credito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   12
         Top             =   1000
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Crédito :"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1050
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   315
         Width           =   570
      End
      Begin VB.Label lblDocNatural 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Identidad :"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   660
         Width           =   1140
      End
      Begin VB.Label lblDocJuridico 
         AutoSize        =   -1  'True
         Caption         =   "RUC :"
         Height          =   195
         Left            =   2760
         TabIndex        =   6
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   5700
      End
      Begin VB.Label lblDocNat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lblDocJur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label LblPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Top             =   255
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmCredEndeuCuotaSistFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ALPA 20160516
'Obsrevación de la SBS por Sobreendeudamiento
Option Explicit
Dim lsCtaCod As String
Dim lsPersCod As String
Dim fbAceptar As Boolean
Dim fbSalir As Boolean
Dim fbFocoGrilla As Boolean
Dim fbCheckGrilla As Boolean
Public Sub Inicio(ByVal psCtaCod As String, ByVal psPersCod As String)
    lsCtaCod = psCtaCod
    lsPersCod = psPersCod
    Me.Show 1
End Sub

Private Sub cmdBuscar_Click()

Dim oPersona As COMDPersona.UCOMPersona
Dim sPersCod As String
Dim oCliPre As COMNCredito.NCOMCredito
Set oCliPre = New COMNCredito.NCOMCredito
Dim lsPersTDoc As String
    Set oPersona = frmBuscaPersona.Inicio
    If Not oPersona Is Nothing Then
        LblPersCod.Caption = oPersona.sPersCod
        lblNomPers.Caption = oPersona.sPersNombre
        lblDocnat.Caption = Trim(oPersona.sPersIdnroDNI)
        lblDocJur.Caption = Trim(oPersona.sPersIdnroRUC)
        lsPersTDoc = "1"
        If oPersona.sPersPersoneria = "1" Then
            'fbPersNatural = True
            If Trim(oPersona.sPersIdnroDNI) = "" Then
                If Not Trim(oPersona.sPersIdnroOtro) = "" Then
                    lblDocnat.Caption = Trim(oPersona.sPersIdnroOtro)
                    lsPersTDoc = Trim(oPersona.sPersTipoDoc)
                End If
            End If
        Else
            'fbPersNatural = False
            lsPersTDoc = "3"
        End If
    Else
        Exit Sub
    End If
    
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdGrabar_Click()
Dim i As Integer
Dim Prev As previo.clsprevio
Set Prev = New previo.clsprevio
Dim oCredDoc As COMNCredito.NCOMCredDoc
Set oCredDoc = New COMNCredito.NCOMCredDoc
Dim objCredito As COMDCredito.DCOMCredito
Set objCredito = New COMDCredito.DCOMCredito
Dim bLogicoSinFecha As Boolean
    bLogicoSinFecha = True
    If FEDeudaSF.TextMatrix(1, 0) = "" Then
        MsgBox "No existen registros de deuda en el sistema financiero", vbInformation, "Aviso!"
        Exit Sub
    End If
    For i = 1 To FEDeudaSF.Rows - 1
        If FEDeudaSF.TextMatrix(i, 2) <= 0 Or Not IsNumeric(FEDeudaSF.TextMatrix(i, 2)) Then
            bLogicoSinFecha = False
            MsgBox "Favor ingresar el día de vencimiento de la cuota, no debe ser cero", vbInformation, "Aviso!"
            Exit Sub
        End If
    Next i
    For i = 1 To FEDeudaSF.Rows - 1
        Call objCredito.ActualizaCredEndeuCuotaSistFinanc(gdFecSis, lsCtaCod, FEDeudaSF.TextMatrix(i, 0), FEDeudaSF.TextMatrix(i, 2), "", FEDeudaSF.TextMatrix(i, 4), FEDeudaSF.TextMatrix(i, 5), FEDeudaSF.TextMatrix(i, 6))
    Next i
    Prev.Show oCredDoc.ImprimeCuotasEndeudamiento(lsCtaCod, lblNomPers.Caption, "", gdFecSis, 0, "CMAC Maynas", 0, "109", gsCodUser), "", True
    Set oCredDoc = Nothing
    Set Prev = Nothing
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim oPersona As UPersona_Cli
    Set oPersona = New UPersona_Cli
    Dim uPersona As COMDPersona.UCOMPersona
    
    Set uPersona = New COMDPersona.UCOMPersona
    
    
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Set ClsPersona = New COMDPersona.DCOMPersonas
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Set R = ClsPersona.BuscaCliente(lsPersCod, BusquedaCodigo)
    LblPersCod.Caption = R!cPersCod
    lblNomPers.Caption = R!cPersNombre
    lblDocnat.Caption = IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI)
    lblDocJur.Caption = IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC)
    Credito.Caption = lsCtaCod
    Call CargarDatos
    Frame1.Enabled = False
    Set oPersona = Nothing
End Sub
Private Sub CargarDatos()
Dim objCredito As COMDCredito.DCOMCredito
Set objCredito = New COMDCredito.DCOMCredito
Dim i As Integer
Dim oRS As ADODB.Recordset
Set oRS = New ADODB.Recordset
Set oRS = objCredito.RecuperaCredEndeuCuotaSistFinanc(lsCtaCod)
LimpiaFlex FEDeudaSF
    Do While Not oRS.EOF
        FEDeudaSF.AdicionaFila
        FEDeudaSF.TextMatrix(oRS.Bookmark, 0) = IIf(IsNull(oRS!Cod_Emp), 0, oRS!Cod_Emp)
        FEDeudaSF.TextMatrix(oRS.Bookmark, 1) = IIf(IsNull(oRS!Nombre), 0, oRS!Nombre)
        FEDeudaSF.TextMatrix(oRS.Bookmark, 2) = IIf(IsNull(oRS!nDiasPago), 0, oRS!nDiasPago)
        FEDeudaSF.TextMatrix(oRS.Bookmark, 4) = Format(IIf(IsNull(oRS!nDeudaIFI), 0, oRS!nDeudaIFI), gsFormatoNumeroView)
        FEDeudaSF.TextMatrix(oRS.Bookmark, 5) = Format(IIf(IsNull(oRS!nCuotaEstimada), 0, oRS!nCuotaEstimada), gsFormatoNumeroView)
        FEDeudaSF.TextMatrix(oRS.Bookmark, 6) = Format(IIf(IsNull(oRS!nInteresDevengado), 0, oRS!nInteresDevengado), gsFormatoNumeroView)
        If IIf(IsNull(oRS!nDeudaIFI), 0, oRS!nDeudaIFI) <= 0 Then
            FEDeudaSF.TextMatrix(oRS.Bookmark, 7) = "Solicita cronograma de pago..."
        End If
        oRS.MoveNext
    Loop
End Sub
Private Sub FEDeudaSF_GotFocus()
    fbFocoGrilla = True
End Sub
Private Sub FEDeudaSF_LostFocus()
    fbFocoGrilla = False
End Sub
Private Sub FEDeudaSF_OnCellChange(pnRow As Long, pnCol As Long)
    Dim iGP As Integer
    Dim nValDiario As Currency
    Dim ldFechaFinDeMes  As Date
    If FEDeudaSF.TextMatrix(pnRow, 0) <> "" Then
        If pnCol = 2 Then
             ldFechaFinDeMes = DateSerial(Year(gdFecSis), Month(gdFecSis) + 1, 1 - 1)
             
             If CDbl(FEDeudaSF.TextMatrix(pnRow, 2)) > 31 Then
                MsgBox "La cuota no puede ser mas de 31...", vbInformation, "Aviso!"
                FEDeudaSF.TextMatrix(pnRow, 2) = 0
                FEDeudaSF.TextMatrix(pnRow, 5) = 0
                Exit Sub
             End If
             
             If Day(ldFechaFinDeMes) < CLng(FEDeudaSF.TextMatrix(pnRow, 2)) Then
                MsgBox "El día de vencimmiento de la cuota no puede ser mayor al último día de este mes...", vbInformation, "Aviso!"
                FEDeudaSF.TextMatrix(pnRow, 2) = 0
                FEDeudaSF.TextMatrix(pnRow, 5) = 0
                Exit Sub
             End If
             If CInt(FEDeudaSF.TextMatrix(pnRow, 2)) > 240 Then
                MsgBox "La cuota no puede ser mas de 240...", vbInformation, "Aviso!"
                FEDeudaSF.TextMatrix(pnRow, 2) = 0
                FEDeudaSF.TextMatrix(pnRow, 5) = 0
                Exit Sub
             End If
             nValDiario = 1 + (CDbl(Day(ldFechaFinDeMes)) - CDbl(FEDeudaSF.TextMatrix(pnRow, 2)))
             nValDiario = (CDbl(FEDeudaSF.TextMatrix(pnRow, 6)) / nValDiario) * Day(ldFechaFinDeMes)
             FEDeudaSF.TextMatrix(pnRow, 5) = Round(CDbl(FEDeudaSF.TextMatrix(pnRow, 4)) + nValDiario, 2)
        End If
    End If
End Sub
