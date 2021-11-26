VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmColAsignacionGastoLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de gastos en lote"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Height          =   390
      Left            =   4320
      TabIndex        =   16
      Top             =   6600
      Width           =   1440
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7440
      TabIndex        =   14
      Top             =   6600
      Width           =   1350
   End
   Begin VB.CommandButton CmdAsignarGasto 
      Caption         =   "Asignar &Gasto"
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
      Height          =   390
      Left            =   5880
      TabIndex        =   13
      Top             =   6600
      Width           =   1440
   End
   Begin VB.Frame fraCreditos 
      Height          =   3210
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   8745
      Begin SICMACT.FlexEdit FECredito 
         Height          =   2985
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   5265
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "--Credito-Cliente-Saldo Cap-Cuota-Monto. Pend.-"
         EncabezadosAnchos=   "300-300-1800-3000-1200-800-1200-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X-X"
         ListaControles  =   "0-4-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-R-R-R-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         SelectionMode   =   1
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7440
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ComboBox CmbTipoCred 
      Height          =   315
      ItemData        =   "frmColAsignacionGastoLote.frx":0000
      Left            =   1680
      List            =   "frmColAsignacionGastoLote.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   2160
   End
   Begin MSComCtl2.Animation AnmBuscar 
      Height          =   675
      Left            =   6600
      TabIndex        =   4
      Top             =   2400
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1191
      _Version        =   393216
      FullWidth       =   53
      FullHeight      =   45
   End
   Begin VB.Frame fraGastos 
      Caption         =   "Gastos"
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
      Height          =   2150
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8730
      Begin SICMACT.FlexEdit FEGastos 
         Height          =   1845
         Left            =   165
         TabIndex        =   1
         Top             =   240
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   3545
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "-Codigo-Gasto-Tipo Valor-Valor-Moneda---Apl. Tip. Cambio-Valor a Tip. Camb."
         EncabezadosAnchos=   "300-800-3000-1000-1000-1000-0-0-1500-1500"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-C-C-C-C-C-L-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         SelectionMode   =   1
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
      End
   End
   Begin MSMask.MaskEdBox mskDe 
      Height          =   315
      Left            =   1680
      TabIndex        =   18
      Top             =   2640
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskHasta 
      Height          =   330
      Left            =   3360
      TabIndex        =   20
      Top             =   2640
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame3 
      Height          =   570
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Width           =   2175
      Begin VB.OptionButton OptGasto 
         Caption         =   "Ninguno"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptGasto 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   11
         Top             =   255
         Width           =   840
      End
   End
   Begin VB.CheckBox chkProdTodos 
      Caption         =   "Todos los Productos"
      Height          =   195
      Left            =   3960
      TabIndex        =   7
      Top             =   2280
      Width           =   1995
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   " Al"
      Height          =   195
      Left            =   3000
      TabIndex        =   19
      Top             =   2640
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Desembolsados del:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Label lblMensaje 
      BackColor       =   &H80000005&
      Caption         =   "  Buscando ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4920
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Producto :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1320
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Nro. Total de Registros: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2400
      TabIndex        =   15
      Top             =   6360
      Width           =   4515
   End
End
Attribute VB_Name = "frmColAsignacionGastoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private R As ADODB.Recordset
Private nFila As Integer

Private Sub chkProdTodos_Click()
    If chkProdTodos.value = 1 Then
        cmbTipoCred.Enabled = False
        cmbTipoCred.ListIndex = -1
    Else
        cmbTipoCred.Enabled = True
    End If
End Sub

Private Sub cmdAplicar_Click()
Dim oDCred As COMDCredito.DCOMCredito
Dim i As Integer
Dim lsMoneda As String
Dim lsProducto As String
    
    If Trim(FEGastos.TextMatrix(0, 1)) = "" Then
        MsgBox "No existen Gastos Aplicables a las Cuotas", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If cmbTipoCred.ListIndex = -1 And chkProdTodos.value = 2 Then
        MsgBox "No a seleccionado ningún producto", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If IsDate(mskDe.Text) = False Then
        MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
        mskDe.SetFocus
        Exit Sub
    End If
    If IsDate(mskHasta.Text) = False Then
        MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
        mskHasta.SetFocus
        Exit Sub
    End If
                
    LimpiaFlex FECredito
    Set R = Nothing
    Set oDCred = New COMDCredito.DCOMCredito
    
    lblMensaje.Visible = True
    Call AbrirControlAnimation(AnmBuscar, 0)
    
    lsProducto = IIf(chkProdTodos.value = 1, "", Mid(Trim(Right(cmbTipoCred.Text, 10)), 1, 1))
    lsMoneda = IIf(Trim(FEGastos.TextMatrix(FEGastos.Row, 8)) = "SI", "", Trim(FEGastos.TextMatrix(FEGastos.Row, 6)))
    
    Set R = oDCred.RecuperaCreditosParaAsignarGastoLote(0, 0, gColocCalendAplCuota, _
            Trim(FEGastos.TextMatrix(FEGastos.Row, 1)), mskDe.Text, mskHasta.Text, lsProducto, lsMoneda)
    Set oDCred = Nothing
    lblMensaje.Visible = False
    Call CerrarControlAnimation(AnmBuscar)
    
    Set FECredito.Recordset = R
    For i = 1 To FECredito.Rows - 1
        FECredito.TextMatrix(i, 0) = i
    Next i
    nFila = FEGastos.Row
    If FECredito.Rows <= 1 Then
        FECredito.Enabled = False
    Else
        FECredito.Enabled = True
    End If
    lblTotal.Caption = "Nro. Total de Registros: " & R.RecordCount
    
    If R.RecordCount > 0 Then
        fraGastos.Enabled = False
        CmdAplicar.Enabled = False
        CmdAsignarGasto.Enabled = True
        cmdCancelar.Enabled = True
    End If
    
End Sub

Private Sub CmdAsignarGasto_Click()
Dim i As Integer
Dim oNCred As COMNCredito.NCOMCredito
Dim MatCuentas() As Variant
Dim bAplTipCamb As Boolean
Dim TamMat As Long, Posi As Long

    If FECredito.Rows <= 1 Then
        MsgBox "No Existen Creditos para Asignar Gastos", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Trim(FECredito.TextMatrix(1, 2)) = "" Then
        MsgBox "No Existen Creditos para Asignar Gastos", vbInformation, "Aviso"
        Exit Sub
    End If
    
    For i = 1 To FECredito.Rows - 1
        If FECredito.TextMatrix(i, 1) = "." Then
            TamMat = TamMat + 1
        End If
    Next i
    
    If TamMat = 0 Then
        MsgBox "No a seleccionado ningún crédito, es necesario que por lo menos un crédito se encuentre seleccionado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Se va a asignar el gasto: " & FEGastos.TextMatrix(nFila, 1) & " - " & FEGastos.TextMatrix(nFila, 2) & _
            " a los " & CStr(TamMat) & " créditos vigentes seleccionados, desembolsados entre las fechas de " & mskDe.Text & " y " & mskHasta.Text & "  , tener en cuenta que una vez asignado el gasto no podrá revertirse, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
     
    ReDim MatCuentas(TamMat, 3)
    Posi = 0
    For i = 1 To FECredito.Rows - 1
        If FECredito.TextMatrix(i, 1) = "." Then
            'ReDim Preserve MatCuentas(UBound(MatCuentas, 1) + 1, 3)
            MatCuentas(Posi, 0) = FECredito.TextMatrix(i, 2) 'cuenta
            MatCuentas(Posi, 1) = FECredito.TextMatrix(i, 7) 'Nro calen
            MatCuentas(Posi, 2) = FECredito.TextMatrix(i, 5) 'cuota
            Posi = Posi + 1
        End If
    Next i
    
    bAplTipCamb = IIf(Trim(FEGastos.TextMatrix(nFila, 8)) = "SI", True, False)
    
    Set oNCred = New COMNCredito.NCOMCredito
    Call oNCred.AsignarGastoLoteCreditoACuotaPendienteTotal(0, 0, MatCuentas, Trim(FEGastos.TextMatrix(nFila, 1)), CDbl(FEGastos.TextMatrix(nFila, 4)), CInt(FEGastos.TextMatrix(nFila, 7)), bAplTipCamb, CDbl(FEGastos.TextMatrix(nFila, 9)))
    
    MsgBox "La asignación de gasto en lote finalizó satisfactoriamente", vbInformation, "Mensaje"
    Set oNCred = Nothing
    CmdAsignarGasto.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    fraGastos.Enabled = True
    LimpiaFlex FECredito
    CmdAplicar.Enabled = True
    CmdAsignarGasto.Enabled = False
    cmdCancelar.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraSdi Me
    mskDe.Text = Format(gdFecSis, "dd/MM/YYYY")
    mskHasta.Text = Format(gdFecSis, "dd/MM/YYYY")
    Call CargaControles
    Call CargaGastos
End Sub

Private Sub CargaControles()
Dim oDCred As COMDCredito.DCOMCredito
Dim RTemp As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oDCred = New COMDCredito.DCOMCredito
    Set RTemp = oDCred.RecuperaTiposCredito
    cmbTipoCred.Clear
    Do While Not RTemp.EOF
        cmbTipoCred.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Set oDCred = Nothing
    Exit Sub

ERRORCargaControles:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CargaGastos()
Dim oDGasto As COMDCredito.DCOMGasto
Dim R As ADODB.Recordset

    On Error GoTo ErrorCargaGastos
    LimpiaFlex FEGastos
    Set oDGasto = New COMDCredito.DCOMGasto
    Set R = oDGasto.RecuperaGastosAplicablesCuotas(, "'MA'")
    Set oDGasto = Nothing
    Do While Not R.EOF
        If R!nTpoValor = 1 Then 'Por el momento considerar solo gastos tipo valor
            FEGastos.AdicionaFila
            FEGastos.TextMatrix(R.Bookmark, 1) = Trim(Str(R!nPrdConceptoCod))
            FEGastos.TextMatrix(R.Bookmark, 2) = Trim(R!cdescripcion)
            FEGastos.TextMatrix(R.Bookmark, 3) = IIf(R!nTpoValor = 1, "VALOR", "PORCENTAJE")
            FEGastos.TextMatrix(R.Bookmark, 4) = Format(R!nValor, "#0.00")
            FEGastos.TextMatrix(R.Bookmark, 5) = IIf(R!nmoneda = gMonedaNacional, "SOLES", "DOLARES")
            FEGastos.TextMatrix(R.Bookmark, 6) = Trim(Str(R!nmoneda))
            FEGastos.TextMatrix(R.Bookmark, 7) = Trim(Str(R!nTpoValor))
            If R!bAplTipCamb = True Then
                FEGastos.TextMatrix(R.Bookmark, 8) = "SI"
                FEGastos.TextMatrix(R.Bookmark, 9) = Format(R!nValorATipCamb, "#0.00")
            End If
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    Exit Sub

ErrorCargaGastos:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub OptGasto_Click(Index As Integer)
Dim i As Integer
    If FECredito.Rows <= 1 Then
        Exit Sub
    End If
    If FECredito.TextMatrix(1, 2) = "" Then
        Exit Sub
    End If
    If Index = 0 Then
        For i = 1 To FECredito.Rows - 1
            FECredito.TextMatrix(i, 1) = "1"
        Next i
    Else
        For i = 1 To FECredito.Rows - 1
            FECredito.TextMatrix(i, 1) = "0"
        Next i
    End If

End Sub
