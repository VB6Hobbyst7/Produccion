VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCredLineaCreditoPriorizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lineas de Crédito - Prioridad de Linea"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13260
   Icon            =   "FrmCredLineaCreditoPriorizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Estado de Líneas"
      TabPicture(0)   =   "FrmCredLineaCreditoPriorizar.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Moneda Nacional"
         Height          =   4215
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   12735
         Begin VB.CommandButton CmdCerrar 
            Caption         =   "Cerrar"
            Height          =   375
            Left            =   11400
            TabIndex        =   13
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   11400
            TabIndex        =   12
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton CmdExportar 
            Caption         =   "Exportar"
            Height          =   375
            Left            =   11400
            TabIndex        =   11
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CommandButton CmdBajar 
            Caption         =   "Bajar"
            Height          =   375
            Left            =   11400
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   "Subir"
            Height          =   375
            Left            =   11400
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin Sicmact.FlexEdit FEPrioridad 
            Height          =   3855
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6800
            Cols0           =   12
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Prio.-Demoninación-Desem.Adeud.-Saldo Adeud.-Saldo Colocado-% Colocado-Diferencia-% max-Estado-Agencia-cMoneda"
            EncabezadosAnchos=   "0-500-2000-1200-1200-1200-1200-1200-1200-1200-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-8-9-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-4-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-R-R-R-R-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   10095
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   8400
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   5880
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox cboAgencia 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label2 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   5040
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Agencia:"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "FrmCredLineaCreditoPriorizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nNumero As Integer
Public Sub MostrarLineasCreditos()
    Dim objLinea As DLineaCreditoV2
    Set objLinea = New DLineaCreditoV2
    Dim objRs As ADODB.Recordset
    Set objRs = New ADODB.Recordset
    Set objRs = objLinea.ObtenerLineaCreditoPriorizar(Trim(Right(cboAgencia.Text, 8)), Trim(Right(cboMoneda.Text, 8)), gdFecSis)
    Dim nPrioridad As Integer
    nPrioridad = 1
    nNumero = 0
    FormateaFlex FEPrioridad
    If Not (objRs.BOF Or objRs.EOF) Then
    Do While Not objRs.EOF
        FEPrioridad.AdicionaFila
        nNumero = nNumero + 1
        FEPrioridad.TextMatrix(nNumero, 0) = IIf(objRs!cLineaCreditoCod = "", 1, objRs!cLineaCreditoCod)
        FEPrioridad.TextMatrix(nNumero, 1) = nPrioridad
        FEPrioridad.TextMatrix(nNumero, 2) = objRs!cLineaCreditoDes
        FEPrioridad.TextMatrix(nNumero, 3) = Format(objRs!nMontoCol, gsFormatoNumeroView)
        FEPrioridad.TextMatrix(nNumero, 4) = Format(objRs!nSaldoLinea, gsFormatoNumeroView)
        FEPrioridad.TextMatrix(nNumero, 5) = Format(objRs!nSaldoCol, gsFormatoNumeroView)
        FEPrioridad.TextMatrix(nNumero, 6) = Format(objRs!nPorceCol * 100, gsFormatoNumeroView)
        FEPrioridad.TextMatrix(nNumero, 7) = Format(objRs!nDiferenc, gsFormatoNumeroView)
        FEPrioridad.TextMatrix(nNumero, 8) = Format(objRs!nPormax, gsFormatoNumeroView)
        FEPrioridad.TextMatrix(nNumero, 9) = objRs!nEstado
        FEPrioridad.TextMatrix(nNumero, 10) = objRs!cAgecod
        FEPrioridad.TextMatrix(nNumero, 11) = objRs!cMoneda
        nPrioridad = nPrioridad + 1
        objRs.MoveNext
    Loop
    End If
End Sub


Private Sub cmdBajar_Click()
    Dim sNumeroA As String
    Dim nPrioridadA As Currency
    Dim sDemoninacionA As String
    Dim nTotalLineaA As Currency
    Dim nSaldoLineaA As Currency
    Dim nSaldoColocadoA As Currency
    Dim nPorColocadoA As Currency
    Dim nDiferenciaA As Currency
    Dim nPormaxA As Currency
    Dim bEstadoA As String
    Dim psAgeCodA As String
    Dim psMonedaA As String
    
    Dim sNumeroB As String
    Dim nPrioridadB As Currency
    Dim sDemoninacionB As String
    Dim nTotalLineaB As Currency
    Dim nSaldoLineaB As Currency
    Dim nSaldoColocadoB As Currency
    Dim nPorColocadoB As Currency
    Dim nDiferenciaB As Currency
    Dim nPormaxB As Currency
    Dim bEstadoB As String
    Dim nPosicion As Integer
    Dim psAgeCodB As String
    Dim psMonedaB As String
    
    nPosicion = FEPrioridad.Row
    If validarRegistroDatosPriorizacion = False Then Exit Sub

    If nPosicion >= 1 And FEPrioridad.Rows - 1 > nPosicion Then
        sNumeroA = FEPrioridad.TextMatrix(nPosicion + 1, 0)
        nPrioridadA = FEPrioridad.TextMatrix(nPosicion + 1, 1)
        sDemoninacionA = FEPrioridad.TextMatrix(nPosicion + 1, 2)
        nTotalLineaA = FEPrioridad.TextMatrix(nPosicion + 1, 3)
        nSaldoLineaA = FEPrioridad.TextMatrix(nPosicion + 1, 4)
        nSaldoColocadoA = FEPrioridad.TextMatrix(nPosicion + 1, 5)
        nPorColocadoA = FEPrioridad.TextMatrix(nPosicion + 1, 6)
        nDiferenciaA = FEPrioridad.TextMatrix(nPosicion + 1, 7)
        nPormaxA = FEPrioridad.TextMatrix(nPosicion + 1, 8)
        bEstadoA = FEPrioridad.TextMatrix(nPosicion + 1, 9)
        psAgeCodA = FEPrioridad.TextMatrix(nPosicion + 1, 10)
        psMonedaA = FEPrioridad.TextMatrix(nPosicion + 1, 11)
        
        sNumeroB = FEPrioridad.TextMatrix(nPosicion, 0)
        nPrioridadB = FEPrioridad.TextMatrix(nPosicion, 1)
        sDemoninacionB = FEPrioridad.TextMatrix(nPosicion, 2)
        nTotalLineaB = FEPrioridad.TextMatrix(nPosicion, 3)
        nSaldoLineaB = FEPrioridad.TextMatrix(nPosicion, 4)
        nSaldoColocadoB = FEPrioridad.TextMatrix(nPosicion, 5)
        nPorColocadoB = FEPrioridad.TextMatrix(nPosicion, 6)
        nDiferenciaB = FEPrioridad.TextMatrix(nPosicion, 7)
        nPormaxB = FEPrioridad.TextMatrix(nPosicion, 8)
        bEstadoB = FEPrioridad.TextMatrix(nPosicion, 9)
        psAgeCodB = FEPrioridad.TextMatrix(nPosicion, 10)
        psMonedaB = FEPrioridad.TextMatrix(nPosicion, 11)
        
        FEPrioridad.TextMatrix(nPosicion + 1, 0) = sNumeroB
        FEPrioridad.TextMatrix(nPosicion + 1, 1) = nPrioridadA
        FEPrioridad.TextMatrix(nPosicion + 1, 2) = sDemoninacionB
        FEPrioridad.TextMatrix(nPosicion + 1, 3) = nTotalLineaB
        FEPrioridad.TextMatrix(nPosicion + 1, 4) = nSaldoLineaB
        FEPrioridad.TextMatrix(nPosicion + 1, 5) = nSaldoColocadoB
        FEPrioridad.TextMatrix(nPosicion + 1, 6) = nPorColocadoB
        FEPrioridad.TextMatrix(nPosicion + 1, 7) = nDiferenciaB
        FEPrioridad.TextMatrix(nPosicion + 1, 8) = nPormaxB
        FEPrioridad.TextMatrix(nPosicion + 1, 9) = IIf(bEstadoB = ".", 1, 0)
        FEPrioridad.TextMatrix(nPosicion + 1, 10) = psAgeCodB
        FEPrioridad.TextMatrix(nPosicion + 1, 11) = psMonedaB
        
        
        FEPrioridad.TextMatrix(nPosicion, 0) = sNumeroA
        FEPrioridad.TextMatrix(nPosicion, 1) = nPrioridadB
        FEPrioridad.TextMatrix(nPosicion, 2) = sDemoninacionA
        FEPrioridad.TextMatrix(nPosicion, 3) = nTotalLineaA
        FEPrioridad.TextMatrix(nPosicion, 4) = nSaldoLineaA
        FEPrioridad.TextMatrix(nPosicion, 5) = nSaldoColocadoA
        FEPrioridad.TextMatrix(nPosicion, 6) = nPorColocadoA
        FEPrioridad.TextMatrix(nPosicion, 7) = nDiferenciaA
        FEPrioridad.TextMatrix(nPosicion, 8) = nPormaxA
        FEPrioridad.TextMatrix(nPosicion, 9) = IIf(bEstadoA = ".", 1, 0)
        FEPrioridad.TextMatrix(nPosicion, 10) = psAgeCodA
        FEPrioridad.TextMatrix(nPosicion, 11) = psMonedaA
               
        FEPrioridad.Row = nPosicion + 1
        FEPrioridad.SetFocus
        CmdExportar.Enabled = False
    End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim objLinea As DLineaCreditoV2
Set objLinea = New DLineaCreditoV2
Dim sNumeroA As String
Dim nPrioridadA As Currency
Dim sDemoninacionA As String
Dim nTotalLineaA As Currency
Dim nSaldoLineaA As Currency
Dim nSaldoColocadoA As Currency
Dim nPorColocadoA As Currency
Dim nDiferenciaA As Currency
Dim nPormaxA As Currency
Dim bEstadoA As String
Dim sAgeCodA As String
Dim sMonedaA As String
Dim I As Integer
Call objLinea.EliminarLineaCreditoPrioridad(sAgeCodA, sMonedaA)
If nNumero > 0 Then
    For I = 1 To FEPrioridad.Rows - 1
        sNumeroA = FEPrioridad.TextMatrix(I, 0)
        nPrioridadA = FEPrioridad.TextMatrix(I, 1)
        sDemoninacionA = FEPrioridad.TextMatrix(I, 2)
        nTotalLineaA = FEPrioridad.TextMatrix(I, 3)
        nSaldoLineaA = FEPrioridad.TextMatrix(I, 4)
        nSaldoColocadoA = FEPrioridad.TextMatrix(I, 5)
        nPorColocadoA = FEPrioridad.TextMatrix(I, 6)
        nDiferenciaA = FEPrioridad.TextMatrix(I, 7)
        nPormaxA = FEPrioridad.TextMatrix(I, 8)
        bEstadoA = FEPrioridad.TextMatrix(I, 9)
        sAgeCodA = FEPrioridad.TextMatrix(I, 10)
        sMonedaA = FEPrioridad.TextMatrix(I, 11)
        Call objLinea.InsertaLineaCreditoPrioridad(nPrioridadA, sNumeroA, sDemoninacionA, nTotalLineaA, nSaldoLineaA, nSaldoColocadoA, nPorColocadoA, nDiferenciaA, nPormaxA, IIf(bEstadoA = ".", 1, 0), sAgeCodA, sMonedaA)
    Next I
    MsgBox "La priorización de línea de crédito se guardó correctamente", vbInformation, "Aviso"
End If
 CmdExportar.Enabled = True

End Sub

Private Sub cmdMostrar_Click()
If Trim(Right(cboAgencia.Text, 8)) = "" Then
    MsgBox "Seleccionar Agencia", vbCritical, "Aviso"
    Exit Sub
End If
If Trim(Right(cboMoneda.Text, 8)) = "" Then
    MsgBox "Seleccionar Moneda", vbCritical, "Aviso"
    Exit Sub
End If
Call MostrarLineasCreditos
CmdExportar.Enabled = True
End Sub
Private Function validarRegistroDatosPriorizacion() As Boolean
    validarRegistroDatosPriorizacion = True
    Dim I As Long, j As Long
    For I = 1 To FEPrioridad.Rows - 1 'valida fila x fila
        For j = 1 To FEPrioridad.Cols - 2 '2 xq el ultimo es aux
'            If j = 1 Or j = 3 Or j = 4 Then 'xq las formulas y la desc son opcionales
                If Trim(FEPrioridad.TextMatrix(I, j)) = "" And j <> 9 Then
                    validarRegistroDatosPriorizacion = False
                    MsgBox "Ud. debe de ingresar el dato '" & UCase(FEPrioridad.TextMatrix(0, j)) & "'", vbInformation, "Aviso"
                    FEPrioridad.Row = I
                    FEPrioridad.Col = j
                    FEPrioridad.SetFocus
                    Exit Function
                End If
'            End If
        Next
    Next
End Function

'Private Sub cmdQuitar_Click()
'    FEPrioridad.EliminaFila FEPrioridad.Row
'    CmdExportar.Enabled = False
'End Sub

Private Sub cmdSubir_Click()
    Dim sNumeroA As String
    Dim nPrioridadA As Currency
    Dim sDemoninacionA As String
    Dim nTotalLineaA As Currency
    Dim nSaldoLineaA As Currency
    Dim nSaldoColocadoA As Currency
    Dim nPorColocadoA As Currency
    Dim nDiferenciaA As Currency
    Dim nPormaxA As Currency
    Dim bEstadoA As String
    
    Dim sNumeroB As String
    Dim nPrioridadB As Currency
    Dim sDemoninacionB As String
    Dim nTotalLineaB As Currency
    Dim nSaldoLineaB As Currency
    Dim nSaldoColocadoB As Currency
    Dim nPorColocadoB As Currency
    Dim nDiferenciaB As Currency
    Dim nPormaxB As Currency
    Dim bEstadoB As String
    Dim nPosicion As Integer
    nPosicion = FEPrioridad.Row
    If validarRegistroDatosPriorizacion = False Then Exit Sub

    If nPosicion > 1 Then
        sNumeroA = FEPrioridad.TextMatrix(nPosicion - 1, 0)
        nPrioridadA = FEPrioridad.TextMatrix(nPosicion - 1, 1)
        sDemoninacionA = FEPrioridad.TextMatrix(nPosicion - 1, 2)
        nTotalLineaA = FEPrioridad.TextMatrix(nPosicion - 1, 3)
        nSaldoLineaA = FEPrioridad.TextMatrix(nPosicion - 1, 4)
        nSaldoColocadoA = FEPrioridad.TextMatrix(nPosicion - 1, 5)
        nPorColocadoA = FEPrioridad.TextMatrix(nPosicion - 1, 6)
        nDiferenciaA = FEPrioridad.TextMatrix(nPosicion - 1, 7)
        nPormaxA = FEPrioridad.TextMatrix(nPosicion - 1, 8)
        bEstadoA = FEPrioridad.TextMatrix(nPosicion - 1, 9)
        
        sNumeroB = FEPrioridad.TextMatrix(nPosicion, 0)
        nPrioridadB = FEPrioridad.TextMatrix(nPosicion, 1)
        sDemoninacionB = FEPrioridad.TextMatrix(nPosicion, 2)
        nTotalLineaB = FEPrioridad.TextMatrix(nPosicion, 3)
        nSaldoLineaB = FEPrioridad.TextMatrix(nPosicion, 4)
        nSaldoColocadoB = FEPrioridad.TextMatrix(nPosicion, 5)
        nPorColocadoB = FEPrioridad.TextMatrix(nPosicion, 6)
        nDiferenciaB = FEPrioridad.TextMatrix(nPosicion, 7)
        nPormaxB = FEPrioridad.TextMatrix(nPosicion, 8)
        bEstadoB = FEPrioridad.TextMatrix(nPosicion, 9)
        
        
        FEPrioridad.TextMatrix(nPosicion - 1, 0) = sNumeroB
        FEPrioridad.TextMatrix(nPosicion - 1, 1) = nPrioridadA
        FEPrioridad.TextMatrix(nPosicion - 1, 2) = sDemoninacionB
        FEPrioridad.TextMatrix(nPosicion - 1, 3) = nTotalLineaB
        FEPrioridad.TextMatrix(nPosicion - 1, 4) = nSaldoLineaB
        FEPrioridad.TextMatrix(nPosicion - 1, 5) = nSaldoColocadoB
        FEPrioridad.TextMatrix(nPosicion - 1, 6) = nPorColocadoB
        FEPrioridad.TextMatrix(nPosicion - 1, 7) = nDiferenciaB
        FEPrioridad.TextMatrix(nPosicion - 1, 8) = nPormaxB
        FEPrioridad.TextMatrix(nPosicion - 1, 9) = IIf(bEstadoB = ".", 1, 0)
        
        FEPrioridad.TextMatrix(nPosicion, 0) = sNumeroA
        FEPrioridad.TextMatrix(nPosicion, 1) = nPrioridadB
        FEPrioridad.TextMatrix(nPosicion, 2) = sDemoninacionA
        FEPrioridad.TextMatrix(nPosicion, 3) = nTotalLineaA
        FEPrioridad.TextMatrix(nPosicion, 4) = nSaldoLineaA
        FEPrioridad.TextMatrix(nPosicion, 5) = nSaldoColocadoA
        FEPrioridad.TextMatrix(nPosicion, 6) = nPorColocadoA
        FEPrioridad.TextMatrix(nPosicion, 7) = nDiferenciaA
        FEPrioridad.TextMatrix(nPosicion, 8) = nPormaxA
        FEPrioridad.TextMatrix(nPosicion, 9) = IIf(bEstadoA = ".", 1, 0)
               
        FEPrioridad.Row = nPosicion - 1
        FEPrioridad.SetFocus
        CmdExportar.Enabled = False
    End If
End Sub

Private Sub FEPrioridad_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 8 Then
        FEPrioridad.TextMatrix(pnRow, pnCol) = Format(FEPrioridad.TextMatrix(pnRow, pnCol), gsFormatoNumeroView)
    End If
End Sub

Private Sub FEPrioridad_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Mat() As String
    Mat = Split(FEPrioridad.ColumnasAEditar, "-")
    If Mat(pnCol) = "X" Then
        MsgBox "Esta columna es no editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    If pnCol = 8 Then
        If Not IsNumeric(FEPrioridad.TextMatrix(pnRow, pnCol)) Then
            MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        Else
            If CCur(FEPrioridad.TextMatrix(pnRow, pnCol)) <= 0 Then
                MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
        End If
    End If
    CmdExportar.Enabled = False
End Sub

Private Sub Form_Load()
    Dim objLinea As DLineaCreditoV2
    Set objLinea = New DLineaCreditoV2
    Dim lvItem As ListItem
    Dim objRs As ADODB.Recordset
    
    Call MostrarLineasCreditos
    
    Set objRs = New ADODB.Recordset
    Set objRs = objLinea.ObtenerLineaCreditoAgencia("")
    Call LlenarCombo(cboAgencia, objRs)
    
    Set objRs = New ADODB.Recordset
    Set objRs = objLinea.ObtenerLineaCreditoMoneda
    Call LlenarCombo(cboMoneda, objRs)
    nNumero = 0
End Sub

Private Sub LlenarCombo(ByRef pCombo As ComboBox, ByRef prs As ADODB.Recordset)
'    pRs.MoveFirst
    If (prs.BOF Or prs.EOF) Then
    Exit Sub
    End If
    pCombo.Clear
    Do While Not prs.EOF
        pCombo.AddItem prs!CDESCRI & Space(300) & prs!cCodigo
        prs.MoveNext
    Loop
End Sub

Private Sub cmdExportar_Click()
Dim oTipCambio As nTipoCambio
Dim lnTipCambio As Currency
Set oTipCambio = New nTipoCambio
    lnTipCambio = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#0.000")
Set oTipCambio = Nothing

Call MostrarReportePrioridad(gdFecSis, lnTipCambio)
End Sub
Private Sub MostrarReportePrioridad(ByVal pdFechaProc As Date, ByVal pnTipCamb As Double)
   Dim objLinea As DLineaCreditoV2
    
Dim R As ADODB.Recordset
Dim lMatCabecera As Variant
Dim lsMensaje As String
Dim lsNombreArchivo As String

    lsNombreArchivo = "ReportePrioridad"
    
    ReDim lMatCabecera(12, 0)
    
    Set objLinea = New DLineaCreditoV2
    Set R = objLinea.ObtenerLineaCreditoPriorizarImprimir(Trim(Right(cboAgencia.Text, 8)), Trim(Right(cboMoneda.Text, 8)), gdFecSis)
    Set objLinea = Nothing
    If Not R Is Nothing Then
        Call GeneraReporteEnArchivoExcelInicio(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Priododidad", "", lsNombreArchivo, lMatCabecera, R, 2, , , True, True)
    Else
        MsgBox lsMensaje, vbInformation, "AVISO"
    End If
End Sub
