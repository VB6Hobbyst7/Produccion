VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredTransfGarantiaAdjudiRemate 
   Caption         =   "Remate de Garantias"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8850
   Begin VB.Frame frmBotones 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   6960
      Width           =   8655
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   2950
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1600
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmDatos 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   8655
      Begin VB.TextBox txtInte 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4920
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CheckBox chkVendido2 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox chkVendido 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   375
      End
      Begin VB.OptionButton chkMoneda 
         Caption         =   "Dolar"
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
         Left            =   7200
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtMonto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4920
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdAAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtLugar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   5295
      End
      Begin MSMask.MaskEdBox TxFecRemate 
         Height          =   345
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.TxtBuscar txtCodPersona 
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1800
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Enabled         =   0   'False
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin SICMACT.TxtBuscar txtCodPersona2 
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   2280
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Enabled         =   0   'False
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Interes"
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
         Left            =   4080
         TabIndex        =   28
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblNombreComprador2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   2280
         Width           =   4575
      End
      Begin VB.Label lblMonto 
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
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblNombreComprador 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label lblComprador 
         Caption         =   "Adjudicatario (s) :"
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
         Left            =   600
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblLugar 
         Caption         =   "Lugar Remate :"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha    :"
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame frmFEdit 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   8655
      Begin VB.CheckBox chAdjudicado 
         Caption         =   "Adjudicado"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   1335
      End
      Begin SICMACT.FlexEdit FERemates 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Remate-Fecha-Lugar-Vendido-Comprador-Monto-cPersCod-nEstado-nMoneda"
         EncabezadosAnchos=   "400-800-1200-3200-800-4000-1200-1200-0-1200"
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
         EncabezadosAlineacion=   "C-C-C-L-C-L-C-C-C-C"
         FormatosEdit    =   "0-5-5-5-0-0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame frmCreditoGarantia 
      Caption         =   "Credito"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.TextBox txtcGarantia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4920
         TabIndex        =   3
         Top             =   280
         Width           =   1935
      End
      Begin SICMACT.ActXCodCta ActXcCtaCod 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Credito"
      End
      Begin VB.Label lblGarantia 
         Caption         =   "Garantia"
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
         Left            =   3960
         TabIndex        =   2
         Top             =   345
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCredTransfGarantiaAdjudiRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCtaCod As String
Dim cNumGaran As String
Dim nEstado As String
Dim MatrixRemate() As String
Dim nPos As Integer
Dim nDat As Integer
Public Sub Iniciar(ByVal pcCtaCod As String, ByVal pcNumGaran, ByRef pnEstado As Integer)
ActXcCtaCod.NroCuenta = pcCtaCod
txtcGarantia.Text = pcNumGaran
nEstado = pnEstado
    If nEstado = 8 Or nEstado = 9 Then
        chAdjudicado.value = 1
        chAdjudicado.Enabled = False
        cmdGrabar.Enabled = False
    End If
Call ObtenerArreglo(txtcGarantia.Text, ActXcCtaCod.NroCuenta)
Me.Show 1
End Sub

Private Sub chkVendido_Click()

If chkVendido.value = 0 Then
    txtCodPersona.Enabled = False
    txtMonto.Enabled = False
Else
    txtCodPersona.Enabled = True
    txtMonto.Enabled = True
End If

End Sub

Private Sub chkVendido2_Click()
If chkVendido2.value = 1 Then
    txtCodPersona2.Enabled = True
Else
    txtCodPersona2.Enabled = False
End If
End Sub

Private Sub cmdAAgregar_Click()
  Dim J As Integer
    Dim i As Integer
    Dim NCaDAr As Integer
    NCaDAr = 0
    Dim nMoneda As Integer
    nMoneda = 0


    If chkMoneda.value = True Then
        nMoneda = 1
    End If
    
If Val(txtMonto.Text) = 0 Or txtLugar.Text = "" Or TxFecRemate.Text = "__/__/____" Then
        MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
        Exit Sub
End If
 If chkVendido.value = 1 Then
        If txtCodPersona.Text = "" Then
            MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
 If chkVendido2.value = 1 Then
        If txtCodPersona2.Text = "" Or chkVendido.value = 0 Then
            MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
            Exit Sub
        End If
 End If
If nDat = 1 Then
For i = 0 To nPos
    If Format(Trim(MatrixRemate(2, i)), "YYYY/MM/DD") = Format(TxFecRemate.Text, "YYYY/MM/DD") Then
        MsgBox "Este dato ya fue registrado...", vbInformation, "Aviso"
        TxFecRemate.Text = "__/__/____"
        Exit Sub
    End If
Next i
End If
'20080811***********/*/
    nDat = 1
    FERemates.AdicionaFila
     If FERemates.Row = 1 Then
        ReDim MatrixRemate(1 To 12, 0 To 0)
     End If
     nPos = FERemates.Row - 1
     MatrixRemate(1, nPos) = FERemates.Row
     ReDim Preserve MatrixRemate(1 To 12, 0 To UBound(MatrixRemate, 2) + 1)
     'If nPos >= 1 Then
        MatrixRemate(1, nPos) = FERemates.Row
        MatrixRemate(2, nPos) = Format((TxFecRemate.Text), "YYYY/MM/DD")
        MatrixRemate(3, nPos) = Trim(txtLugar.Text)
        If chkVendido.value Then
            MatrixRemate(8, nPos) = gPersGarantEstadoRematado
            MatrixRemate(4, nPos) = "Rematado"
            chkVendido.value = 1
            chkVendido.Enabled = False
            cmdAAgregar.Enabled = False
            MatrixRemate(5, nPos) = Trim(Me.lblNombreComprador.Caption)
            If chkVendido2.value Then
            MatrixRemate(11, nPos) = Trim(Me.lblNombreComprador2.Caption)
            End If
        Else
        If nPos = 2 Then
            MatrixRemate(8, nPos) = gPersGarantEstadoAdjudicado
            MatrixRemate(4, nPos) = "Adjudicado"
            'chkVendido.value = 1
            cmdAAgregar.Enabled = False
            MatrixRemate(5, nPos) = ""
            chkVendido.Enabled = False
        Else
            MatrixRemate(8, nPos) = nEstado
            MatrixRemate(4, nPos) = "NO"
            chkVendido.value = 0
            chkVendido.Enabled = True
            cmdAAgregar.Enabled = True
            MatrixRemate(5, nPos) = ""
        End If
        End If
        MatrixRemate(6, nPos) = Val(txtMonto.Text)
        MatrixRemate(7, nPos) = txtCodPersona.Text
        MatrixRemate(10, nPos) = txtCodPersona2.Text
        MatrixRemate(9, nPos) = nMoneda
        MatrixRemate(12, nPos) = Val(txtInte.Text)
    'End If

    For i = 0 To nPos
        FERemates.EliminaFila (1)
    Next i
    For i = 0 To nPos
        FERemates.AdicionaFila
        FERemates.TextMatrix(FERemates.Row, 1) = MatrixRemate(1, i)
        FERemates.TextMatrix(FERemates.Row, 2) = MatrixRemate(2, i)
        FERemates.TextMatrix(FERemates.Row, 3) = MatrixRemate(3, i)
        FERemates.TextMatrix(FERemates.Row, 4) = MatrixRemate(4, i)
        'FERemates.TextMatrix(FERemates.Row, 5) = MatrixRemate(5, i)
        If MatrixRemate(11, i) <> "" Then
            FERemates.TextMatrix(FERemates.Row, 5) = MatrixRemate(5, i) & "/" & MatrixRemate(11, i)
        Else
            FERemates.TextMatrix(FERemates.Row, 5) = MatrixRemate(5, i)
        End If
        FERemates.TextMatrix(FERemates.Row, 6) = CDbl(MatrixRemate(6, i)) + CDbl(MatrixRemate(12, i))
        'FERemates.TextMatrix(FERemates.Row, 7) = MatrixRemate(7, i)
        If MatrixRemate(10, i) <> "" Then
            FERemates.TextMatrix(FERemates.Row, 7) = MatrixRemate(7, i) & "/" & MatrixRemate(10, i)
        Else
            FERemates.TextMatrix(FERemates.Row, 7) = MatrixRemate(7, i)
        End If
        FERemates.TextMatrix(FERemates.Row, 8) = MatrixRemate(8, i)
        NCaDAr = 1
    Next
    txtMonto.Text = "0.00"
    cmdEliminar.Enabled = True
End Sub
Private Sub cmdGrabar_Click()
Dim i As Integer
Dim J As Integer
Dim nCont As Integer

Dim bVendido As Boolean
bVendido = False
If chkVendido.value = 1 Then
    bVendido = True
End If
Dim oGaran As COMNCredito.NCOMGarantia
Set oGaran = New COMNCredito.NCOMGarantia
For i = 0 To nPos
    nCont = nCont + 1
    Call oGaran.InsertarGarantiaRemate(txtcGarantia.Text, ActXcCtaCod.NroCuenta, CInt(MatrixRemate(1, i)), MatrixRemate(3, i), MatrixRemate(2, i), gdFecSis, CInt(MatrixRemate(8, i)), MatrixRemate(7, i), CDbl(MatrixRemate(6, i)), i, gsCodUser, bVendido, gsCodAge, CInt(MatrixRemate(9, i)), MatrixRemate(10, i), MatrixRemate(12, i))
Next i
If nCont > 0 Then
     MsgBox "Datos se registraron correctamente...", vbInformation, "Aviso"
     cmdGrabar.Enabled = False
Else
    MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Call ObtenerArreglo(cCtaCod, cNumGaran)
End Sub


Private Sub txtCodPersona_EmiteDatos()
Me.lblNombreComprador.Caption = Trim(txtCodPersona.psDescripcion)
End Sub

Public Sub ObtenerArreglo(ByVal sNumGarantia As String, ByVal sCodCta As String)
    Dim oGaran As COMDCredito.DCOMGarantia
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Set oGaran = New COMDCredito.DCOMGarantia
    
    Dim i As Integer
    For i = 0 To nPos
        FERemates.EliminaFila (1)
    Next i
    nPos = 0
    Set R = oGaran.RecuperaDatosGarantiaRemate(sNumGarantia)
    If R.RecordCount > 0 Then
            If Not R.EOF And Not R.BOF Then
                R.MoveFirst
            End If
    Do Until R.EOF
        FERemates.AdicionaFila
        nPos = FERemates.Row - 1
        ReDim Preserve MatrixRemate(1 To 12, 0 To nPos + 1)
        FERemates.AdicionaFila
        MatrixRemate(1, nPos) = R!nNumRemate
        MatrixRemate(2, nPos) = Format((R!dFechaRemate), "YYYY/MM/DD")
        MatrixRemate(3, nPos) = RTrim((R!cLugarRemate))
        MatrixRemate(4, nPos) = R!cDesEstado
        MatrixRemate(5, nPos) = R!cPersNombre
        MatrixRemate(6, nPos) = R!nMonto
        MatrixRemate(7, nPos) = R!cPersCod
        MatrixRemate(8, nPos) = R!nEstadoAdju
        MatrixRemate(9, nPos) = R!nMoneda
        MatrixRemate(10, nPos) = R!cPersCod2
        MatrixRemate(11, nPos) = R!cPersNombre2
        MatrixRemate(12, nPos) = R!nInteres
        If R!nEstadoAdju = gPersGarantEstadoAdjudicado Or R!nEstadoAdju = gPersGarantEstadoRematado Then
            chAdjudicado.value = 1
            chAdjudicado.Enabled = False
            cmdAAgregar.Enabled = False
            cmdGrabar.Enabled = False
            cmdEliminar.Enabled = False
        End If
        FERemates.TextMatrix(FERemates.Row, 1) = MatrixRemate(1, nPos)
        FERemates.TextMatrix(FERemates.Row, 2) = MatrixRemate(2, nPos)
        FERemates.TextMatrix(FERemates.Row, 3) = MatrixRemate(3, nPos)
        FERemates.TextMatrix(FERemates.Row, 4) = MatrixRemate(4, nPos)
        If MatrixRemate(11, nPos) <> "" Then
            FERemates.TextMatrix(FERemates.Row, 5) = Trim(MatrixRemate(5, nPos)) & "/" & Trim(MatrixRemate(11, nPos))
        Else
            FERemates.TextMatrix(FERemates.Row, 5) = MatrixRemate(5, nPos)
        End If
        FERemates.TextMatrix(FERemates.Row, 6) = MatrixRemate(6, nPos) + MatrixRemate(12, nPos)
        If MatrixRemate(10, nPos) <> "" Then
            FERemates.TextMatrix(FERemates.Row, 7) = MatrixRemate(7, nPos) & "/" & MatrixRemate(10, nPos)
        Else
            FERemates.TextMatrix(FERemates.Row, 7) = MatrixRemate(7, nPos)
        End If
        FERemates.TextMatrix(FERemates.Row, 8) = MatrixRemate(8, nPos)
        FERemates.TextMatrix(FERemates.Row, 9) = MatrixRemate(9, nPos)
        R.MoveNext
        cmdEliminar.Enabled = True
    Loop
    End If
End Sub

Private Sub txtCodPersona2_EmiteDatos()
Me.lblNombreComprador2.Caption = Trim(txtCodPersona2.psDescripcion)
End Sub
