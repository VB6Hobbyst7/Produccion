VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredValorizaCheque 
   Caption         =   "Valorizar Cheque de Creditos"
   ClientHeight    =   4830
   ClientLeft      =   2175
   ClientTop       =   3570
   ClientWidth     =   10530
   Icon            =   "frmCredValorizaCheque.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   10530
   Begin VB.CheckBox chkenValorrizacion 
      Caption         =   "En Valorizacion"
      Height          =   210
      Left            =   1770
      TabIndex        =   13
      Top             =   4275
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdNuevaBusq 
      Caption         =   "Nueva Busqueda"
      Height          =   435
      Left            =   7245
      TabIndex        =   6
      Top             =   4275
      Width           =   1665
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   8925
      TabIndex        =   7
      Top             =   4275
      Width           =   1470
   End
   Begin VB.CommandButton CmdValorizar 
      Caption         =   "&En Valorización"
      Enabled         =   0   'False
      Height          =   435
      Left            =   105
      TabIndex        =   5
      Top             =   4275
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   105
      TabIndex        =   8
      Top             =   75
      Width           =   10320
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "Aplicar"
         Height          =   405
         Left            =   8535
         TabIndex        =   3
         Top             =   405
         Width           =   1605
      End
      Begin VB.ComboBox CboMoneda 
         Height          =   315
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   465
         Width           =   1905
      End
      Begin VB.Frame fraValor 
         Caption         =   "Valorizacion"
         Height          =   795
         Left            =   90
         TabIndex        =   9
         Top             =   150
         Width           =   3570
         Begin MSMask.MaskEdBox TxtFecIni 
            Height          =   300
            Left            =   510
            TabIndex        =   0
            Top             =   330
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtFecFin 
            Height          =   300
            Left            =   2175
            TabIndex        =   1
            Top             =   330
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Caption         =   "Al"
            Height          =   270
            Left            =   1860
            TabIndex        =   11
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   270
            Left            =   105
            TabIndex        =   10
            Top             =   345
            Width           =   450
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda :"
         Height          =   285
         Left            =   3750
         TabIndex        =   12
         Top             =   495
         Width           =   735
      End
   End
   Begin SICMACT.FlexEdit FECheques 
      Height          =   2895
      Left            =   105
      TabIndex        =   4
      Top             =   1290
      Width           =   10335
      _ExtentX        =   18203
      _ExtentY        =   5027
      Cols0           =   11
      FixedCols       =   0
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-OPT-Banco-Cta Bco-Nro Docum-Monto-Producto-Estado-Valorizacion-nEstado-cPersCod"
      EncabezadosAnchos=   "400-450-2800-1200-1200-1200-2300-1200-1800-0-0"
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
      ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   1
      ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-R-L-L-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0-0"
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483646
   End
End
Attribute VB_Name = "frmCredValorizaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset

Private Sub CboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdAplicar.SetFocus
    End If
End Sub

Private Sub CmdAplicar_Click()
Dim odCred As COMDCredito.DCOMCredito
Dim sCad As String
Dim i As Integer
Dim j As Integer

    If CboMoneda.ListIndex = -1 Then
        MsgBox "Seleccione una Moneda", vbInformation, "Aviso"
        CboMoneda.SetFocus
        Exit Sub
    End If

    sCad = ValidaFecha(TxtFecIni.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        TxtFecIni.SetFocus
        Exit Sub
    End If

    sCad = ValidaFecha(TxtFecFin.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        TxtFecFin.SetFocus
        Exit Sub
    End If

    Set odCred = New COMDCredito.DCOMCredito
    Set R = odCred.RecuperaChequeCreditos(CDate(TxtFecIni.Text), CDate(TxtFecFin.Text), CInt(Right(CboMoneda.Text, 2)))
    Set odCred = Nothing
    
    Set FECheques.Recordset = R
    
    If FECheques.Rows > 1 Then
        For i = 1 To FECheques.Rows - 1
            If FECheques.TextMatrix(i, 9) = "0" Then
                FECheques.Row = i
                For j = 1 To 9
                    FECheques.Col = j
                    FECheques.CellBackColor = vbGreen
                Next j
            End If
        Next i
    End If
    
    fraValor.Enabled = False
    CboMoneda.Enabled = False
    If FECheques.Rows > 1 Then
        CmdValorizar.Enabled = True
    Else
        CmdValorizar.Enabled = False
    End If
    FECheques.SetFocus
End Sub

Private Sub cmdNuevaBusq_Click()
    LimpiaFlex FECheques
    TxtFecIni.Text = "__/__/____"
    TxtFecFin.Text = "__/__/____"
    CboMoneda.ListIndex = -1
    CboMoneda.Enabled = True
    fraValor.Enabled = True
    TxtFecIni.SetFocus
    CmdValorizar.Enabled = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdValorizar_Click()
Dim oCred As COMDCredito.DCOMCredActBD
Dim pnEstChequeNew As ChequeEstado
Dim i As Long
Dim MatCheques() As String
Dim MatIndex() As Integer

    If MsgBox("Se va a Valorizar el Cheque, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    If chkenValorrizacion = 1 Then
        pnEstChequeNew = gChqEstEnValorizacion
    Else
        pnEstChequeNew = gChqEstValorizado
    End If
'Set oCred = New COMDCredito.DCOMCredActBD
    
    ReDim MatIndex(0)
    For i = 1 To Me.FECheques.Rows - 1
        If FECheques.TextMatrix(i, 1) <> "" Then
        'Call oCred.dUpdateDocRec(FECheques.TextMatrix(FECheques.Row, 3), FECheques.TextMatrix(FECheques.Row, 9), False, pnEstChequeNew)
        'Call oCred.dUpdateDocRecEst(FECheques.TextMatrix(FECheques.Row, 3), FECheques.TextMatrix(FECheques.Row, 9), False, pnEstChequeNew)
    '    Call oCred.dUpdateDocRec(FECheques.TextMatrix(i, 4), FECheques.TextMatrix(i, 10), False, pnEstChequeNew)
    '    Call oCred.dUpdateDocRecEst(FECheques.TextMatrix(i, 4), FECheques.TextMatrix(i, 10), False, pnEstChequeNew)
            ReDim Preserve MatIndex(UBound(MatIndex) + 1)
            MatIndex(UBound(MatIndex) - 1) = i
        End If
    Next i

    ReDim MatCheques(UBound(MatIndex), 2)
    For i = 0 To UBound(MatCheques) - 1
        MatCheques(i, 0) = FECheques.TextMatrix(MatIndex(i), 4)
        MatCheques(i, 1) = FECheques.TextMatrix(MatIndex(i), 10)
    Next i

Set oCred = New COMDCredito.DCOMCredActBD
Call oCred.ValorizarCheques(MatCheques, pnEstChequeNew)
Set oCred = Nothing
Call CmdAplicar_Click
End Sub

Private Sub FECheques_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If Trim(FECheques.TextMatrix(pnRow, 9)) = "2" Or Trim(FECheques.TextMatrix(pnRow, 9)) = "1" Then
    MsgBox "No se puede Valorizar el Cheque, porque el estado ya es valorizado o en Valorizacion ", vbInformation, "Aviso"
    FECheques.TextMatrix(pnRow, 1) = ""
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim oCons As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset

CentraForm Me
Set oCons = New COMDConstantes.DCOMConstantes
Set rs = oCons.RecuperaConstantes(gMoneda)
Call Llenar_Combo_con_Recordset(rs, CboMoneda)
Set oCons = Nothing
'Call CargaComboConstante(gMoneda, CboMoneda)
End Sub

Private Sub TxtFecFin_GotFocus()
    fEnfoque TxtFecFin
End Sub

Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboMoneda.SetFocus
    End If
End Sub

Private Sub TxtFecIni_GotFocus()
    fEnfoque TxtFecIni
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtFecFin.SetFocus
    End If
End Sub
