VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBanyOtrasInstSisFinan 
   Caption         =   "Bancos y otras Instituciones del Sistema Financiero"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   Icon            =   "frmBanyOtrasInstSisFinan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   9
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   12120
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Frame frDatos 
      Caption         =   "A. Consolidado de depositos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   11160
         TabIndex        =   15
         Top             =   6840
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parametros de Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   13095
         Begin VB.CommandButton cmdVisualizar 
            Caption         =   "Mostar Detalle"
            Height          =   375
            Left            =   10680
            TabIndex        =   10
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox CboMoneda 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   3015
         End
         Begin VB.ComboBox cboTpoCta 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1080
            Width           =   5775
         End
         Begin VB.ComboBox cboCuadro 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   3015
         End
         Begin MSMask.MaskEdBox txtFechaProceso 
            Height          =   300
            Left            =   10440
            TabIndex        =   8
            Top             =   360
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   13
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   12
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Cuadros (Bancos o Cajas)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   11
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de proceso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8640
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
      End
      Begin Sicmact.FlexEdit FEDetalle 
         Height          =   4335
         Left            =   0
         TabIndex        =   1
         Top             =   2400
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   7646
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "##-CodBanCaja-cMoneda-cTpoCta-NroCuadro-cAgeCod-Agencia-Moneda-TipoCta-Banco-Nro Cta-Saldo"
         EncabezadosAnchos=   "100-0-0-0-0-0-2200-1200-2000-2500-2000-2500"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-11"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483628
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-L-L-L-L-L-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-4"
         AvanceCeldas    =   1
         TextArray0      =   "##"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   105
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   16
         Top             =   6880
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmBanyOtrasInstSisFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sMatMatrix() As String
Dim nPost As Integer
Dim ldFecha As Date
Dim lsOpeCod As String
Dim lsMoneda As String
Public Sub inicio(psOpeCod As String)
    ldFecha = gdFecSis
    lsOpeCod = psOpeCod
    lsMoneda = Mid(psOpeCod, 3, 1)
    Me.Show 1
End Sub

Private Sub CmdCancelar_Click()

End Sub

Private Sub cmdGrabar_Click()
    Dim i As Integer
    Dim oDBalance As DbalanceCont
    For i = 1 To nPost
        Set oDBalance = New DbalanceCont
        Call oDBalance.InsertaSaldoCtas(FEDetalle.TextMatrix(i, 5), FEDetalle.TextMatrix(i, 1), lsMoneda, FEDetalle.TextMatrix(i, 3), FEDetalle.TextMatrix(i, 4), txtFechaProceso.Text, FEDetalle.TextMatrix(i, 11))
    Next i
'FEDetalle.TextMatrix(j, 1) = rs!cCodigoBancoAgeCod
'FEDetalle.TextMatrix(j, 2) = rs!cmoneda
'FEDetalle.TextMatrix(j, 3) = rs!cTpoCta
'FEDetalle.TextMatrix(j, 4) = rs!NroCuadro
'FEDetalle.TextMatrix(j, 5) = rs!cAgeCod
'FEDetalle.TextMatrix(j, 6) = rs!cAgeDescripcion
'FEDetalle.TextMatrix(j, 7) = rs!cMonedaDesc
'FEDetalle.TextMatrix(j, 8) = rs!cTpoCtaDesc
'FEDetalle.TextMatrix(j, 9) = rs!cDescripcionBancoAge
'FEDetalle.TextMatrix(j, 10) = rs!cNroCtaCte
'FEDetalle.TextMatrix(j, 11) = rs!nSaldo
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVisualizar_Click()
    Call Actualizar_FlexEdit
    FEDetalle.lbEditarFlex = True
End Sub

Private Sub FEDetalle_OnCellChange(pnRow As Long, pnCol As Long)
    Dim nTotal As Currency
    Dim i As Integer
    For i = 1 To nPost
        nTotal = nTotal + FEDetalle.TextMatrix(i, 11)
    Next i
    txtTotal.Text = CStr(nTotal)
End Sub
'Private Sub FEDetalle_KeyPress(KeyAscii As Integer)
'    FEDetalle.lbEditarFlex = True
'End Sub




Private Sub Form_Load()
    Call obtener_Cuadros
    Call obtener_Monedas
    Call obtener_TpoCta
    txtFechaProceso.Text = Format(ldFecha, "YYYY/MM/DD")
    Call Actualizar_FlexEdit
End Sub

Public Sub obtener_Cuadros()
    cboCuadro.AddItem "BANCOS           " & Space(200) & "1"
    cboCuadro.AddItem "CAJAS            " & Space(200) & "3"
End Sub

Public Sub obtener_Monedas()
    If lsMoneda = "1" Then
        cboMoneda.AddItem "SOLES           " & Space(200) & "1"
    Else
        cboMoneda.AddItem "DOLARES         " & Space(200) & "2"
    End If
End Sub

Public Sub obtener_TpoCta()
    cboTpoCta.AddItem "CTAS CORRIENTES           " & Space(200) & "1"
    cboTpoCta.AddItem "CTAS AHORROS              " & Space(200) & "2"
    cboTpoCta.AddItem "OVERNIGH                  " & Space(200) & "3"
    cboTpoCta.AddItem "PLAZOS FIJOS              " & Space(200) & "4"
    cboTpoCta.AddItem "FONDOS MUTUOS             " & Space(200) & "5"
    cboTpoCta.AddItem "CERTIF.BANK.              " & Space(200) & "6"
End Sub

Private Sub Actualizar_FlexEdit()
 Dim i, j As Integer
 Dim nTotal As Currency
        Dim oDBalan As DbalanceCont
        Set oDBalan = New DbalanceCont
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        nTotal = 0
        Set rs = oDBalan.ListarSaldoCtas(Right(cboMoneda.Text, 1), Right(cboTpoCta.Text, 1), Right(cboCuadro.Text, 1), txtFechaProceso.Text)
        If nPost > 0 Then
            For i = 1 To nPost
                FEDetalle.EliminaFila (1)
            Next i
        End If
        nPost = 0
        j = 0
        If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
        Do Until rs.EOF
            j = j + 1
            FEDetalle.AdicionaFila
            FEDetalle.TextMatrix(j, 0) = j
            FEDetalle.TextMatrix(j, 1) = rs!cCodigoBancoAgeCod
            FEDetalle.TextMatrix(j, 2) = rs!cmoneda
            FEDetalle.TextMatrix(j, 3) = rs!cTpoCta
            FEDetalle.TextMatrix(j, 4) = rs!NroCuadro
            FEDetalle.TextMatrix(j, 5) = rs!cAgeCod
            FEDetalle.TextMatrix(j, 6) = rs!cAgeDescripcion
            FEDetalle.TextMatrix(j, 7) = rs!cMonedaDesc
            FEDetalle.TextMatrix(j, 8) = rs!cTpoCtaDesc
            FEDetalle.TextMatrix(j, 9) = rs!cDescripcionBancoAge
            FEDetalle.TextMatrix(j, 10) = rs!cNroCtaCte
            FEDetalle.TextMatrix(j, 11) = rs!nSaldo
            nTotal = nTotal + rs!nSaldo
            ReDim Preserve sMatMatrix(0 To 11, 1 To j)
            nPost = j
            sMatMatrix(0, j) = j
            sMatMatrix(1, j) = rs!cCodigoBancoAgeCod
            sMatMatrix(2, j) = rs!cmoneda
            sMatMatrix(3, j) = rs!cTpoCta
            sMatMatrix(4, j) = rs!NroCuadro
            sMatMatrix(5, j) = rs!cAgeCod
            sMatMatrix(6, j) = rs!cAgeDescripcion
            sMatMatrix(7, j) = rs!cMonedaDesc
            sMatMatrix(8, j) = rs!cTpoCtaDesc
            sMatMatrix(9, j) = rs!cDescripcionBancoAge
            sMatMatrix(10, j) = rs!cNroCtaCte
            sMatMatrix(11, j) = rs!nSaldo
           rs.MoveNext
        Loop
        End If
        txtTotal.Text = CStr(nTotal)
End Sub
