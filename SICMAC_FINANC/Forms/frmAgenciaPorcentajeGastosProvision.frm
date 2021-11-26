VERSION 5.00
Begin VB.Form frmAgenciaPorcentajeGastosProvision 
   Caption         =   "Distribución de Gastos - Provisión"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   Icon            =   "frmAgenciaPorcentajeGastosProvision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Monto Inafecto a IGV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   21
      Top             =   4680
      Width           =   2055
      Begin VB.TextBox txtInafecto 
         Height          =   405
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmGasto 
      Caption         =   "Porcentaje Gasto"
      Height          =   1095
      Left            =   9240
      TabIndex        =   15
      Top             =   3480
      Width           =   1935
      Begin VB.OptionButton optCartCred 
         Caption         =   "% Cartera Cred."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optCartAho 
         Caption         =   "% Cartera Aho."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optIngFin 
         Caption         =   "% Ingreso Financ."
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Monto total (Valor Opcional)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   10
      Top             =   4680
      Width           =   3375
      Begin VB.TextBox txtAgeCod 
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtTotal 
         Height          =   405
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox ckOpcionActivar 
         Caption         =   "Check1"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Monto a Distribuir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   8
      Top             =   4680
      Width           =   1935
      Begin VB.TextBox txtMonto 
         Height          =   405
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Elegir opción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   3135
      Begin VB.OptionButton optCantitra 
         Caption         =   "Nro. de trabajadores"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optPorcentaje 
         Caption         =   "Porcentaje"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Porcentaje de gastos  por agencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin Sicmact.FlexEdit FEGasAge 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7011
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-CodAge-DescAge-Porcentaje-Activar"
         EncabezadosAnchos=   "400-1200-4000-1400-1200"
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
         ColumnasAEditar =   "X-X-X-3-4"
         ListaControles  =   "0-0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-C"
         FormatosEdit    =   "0-0-0-2-0"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin Sicmact.TxtBuscar txtCta 
      Height          =   330
      Left            =   2400
      TabIndex        =   19
      Top             =   5520
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   582
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sTitulo         =   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contrapartida Asiento Distrib.:"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   5520
      Width           =   2085
   End
End
Attribute VB_Name = "frmAgenciaPorcentajeGastosProvision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsMatrizDatos() As Variant
Dim lnTipoMatriz As Integer
Dim nPost As Integer
Dim j As Integer
Dim FENoMoverdeFila As Integer
Dim sMatrizTemp() As String
Dim nContador As Integer
Dim nPorcentaje As Currency
Dim lnSeleAge As Integer
Dim lnMontoGasto As Currency
Dim lcCtaContrapDistriAstoManual As String '*** PEAC 20100712
Dim lsCtaContrapDistriAstoManual As String
Dim lnMontoInafecto As Currency
Dim lbCtaCtrapDistMan As Boolean


Public Sub Inicio(ByRef sMatrizDatos As Variant, ByRef nTipoMatriz As Integer, ByRef nSeleAge As Integer, ByRef nMontogasto As Currency, Optional ByRef sCtaContrapDistriAstoManual As String = "", Optional ByVal bCtaCtrapDistMan As Boolean = False, Optional ByRef nMontoInafecto As Currency = 0)
'*** PEAC 20100712 - S AGREGO EL PARAMETRO "sCtaContrapDistriAstoManual" y "bCtaCtrapDistMan"

    '*** PEAC 20100713
    If bCtaCtrapDistMan Then
        Me.txtCta.Visible = True
        lbCtaCtrapDistMan = True
    Else
        Me.txtCta.Visible = False
        lbCtaCtrapDistMan = False
    End If
    '*** FIN PEAC
    lsMatrizDatos = sMatrizDatos
    lnTipoMatriz = nTipoMatriz
    lnMontoGasto = nMontogasto
    lnSeleAge = nSeleAge
    lcCtaContrapDistriAstoManual = sCtaContrapDistriAstoManual
    lnMontoInafecto = nMontoInafecto
    
    Show 1
    sMatrizDatos = lsMatrizDatos
    nTipoMatriz = lnTipoMatriz
    nSeleAge = lnSeleAge
    nMontogasto = lnMontoGasto
    sCtaContrapDistriAstoManual = lsCtaContrapDistriAstoManual '*** PEAC 20100713
    nMontoInafecto = lnMontoInafecto
    
    
    
End Sub

Private Sub OptionDatos()
    If optPorcentaje.value = True Then
        Call LLamarCargarDatosRs
        lnTipoMatriz = 1
    Else
        Call CargarDatosCantidad
        lnTipoMatriz = 2
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Integer
    Dim nContadorTemp  As Integer
    Dim nValorTotalCanti As Currency
    
'    '*********** BEGIN valida importes porcentuales al 100%
'    nValorTotalCanti = 0
'    nPost = FEGasAge.Rows - 1
'
'    For i = 1 To nPost
'        If FEGasAge.TextMatrix(i, 6) = "." Then
'            If CInt(FEGasAge.TextMatrix(i, 5)) <= 0# Then
'                MsgBox "Agencia " & Trim(FEGasAge.TextMatrix(i, 1)) & " no tiene valor, revise." ''& 1
'                Exit Sub
'            End If
'            nValorTotalCanti = nValorTotalCanti + FEGasAge.TextMatrix(i, 5)
'        End If
'    Next i
'
'    If nValorTotalCanti <> 100 Then
'        MsgBox "Los porcentajes ingresados tienen que dar a 100%, los ingresados suman : " & Trim(CStr(nValorTotalCanti)) & "%" & IIf(nValorTotalCanti > 100#, " demas " & Trim(CStr(nValorTotalCanti - 100#)), " falta " & Trim(CStr(100# - nValorTotalCanti)))
'    Else
'        MsgBox "Porcentajes al 100% OK.", vbOKOnly, "Atención"
'    End If
'    '*********** END valida importes porcentuales al 100%
    
    
    
    nContador = 0
    nContadorTemp = 0
    nValorTotalCanti = 0
    For i = 1 To nPost
        If FEGasAge.TextMatrix(i, 4) = "." Then
            If CInt(FEGasAge.TextMatrix(i, 3)) < 0 Then
                MsgBox "Agencia seleccionada no tiene valor porcentual."
                Exit Sub
            End If
            nContador = nContador + 1
            nValorTotalCanti = nValorTotalCanti + FEGasAge.TextMatrix(i, 3)
        End If
    Next i
    
    If nValorTotalCanti <> 100 Then
        MsgBox "Los porcentajes ingresados tienen que dar a 100%, los ingresados suman : " & Trim(CStr(nValorTotalCanti)) & "%" & IIf(nValorTotalCanti > 100#, " demas " & Trim(CStr(nValorTotalCanti - 100#)), " falta " & Trim(CStr(100# - nValorTotalCanti)))
'    Else
'        MsgBox "Porcentajes al 100% OK.", vbOKOnly, "Atención"
    End If
    
    
    If IIf(Trim(txtMonto.Text) = "", 0, txtMonto.Text) = 0 Then
        MsgBox "Ingresar monto del gasto"
        Exit Sub
    End If
    lnMontoGasto = CDbl(txtMonto.Text)
    If nContador <= 0 Then
            MsgBox "Seleccionar al menos una agencia "
            Exit Sub
    End If
    If nContador <> nPost Then
    nPorcentaje = 100 / nContador
'    If nValorTotalCanti > 0 Then
    For i = 1 To nPost
        If FEGasAge.TextMatrix(i, 4) = "." Then
            nContadorTemp = nContadorTemp + 1
            ReDim Preserve lsMatrizDatos(1 To 4, 1 To i)
            lsMatrizDatos(1, nContadorTemp) = FEGasAge.TextMatrix(i, 1)
            lsMatrizDatos(2, nContadorTemp) = FEGasAge.TextMatrix(i, 2)
            If nValorTotalCanti > 0 Then
             If lnTipoMatriz = 1 Then
            
                 lsMatrizDatos(3, nContadorTemp) = nPorcentaje
             Else
                 lsMatrizDatos(3, nContadorTemp) = (FEGasAge.TextMatrix(i, 3) / nValorTotalCanti) * 100
             End If
            lsMatrizDatos(4, nContadorTemp) = 0
            Else
                MsgBox "Ingresar los valores en los casilleros"
                Exit Sub
            End If
        End If
    Next i
'    End If
    Else
     If lnTipoMatriz <> 1 Then
        For i = 1 To nPost
            If FEGasAge.TextMatrix(i, 4) = "." Then
                nContadorTemp = nContadorTemp + 1
                ReDim Preserve lsMatrizDatos(1 To 4, 1 To i)
                lsMatrizDatos(1, nContadorTemp) = FEGasAge.TextMatrix(i, 1)
                lsMatrizDatos(2, nContadorTemp) = FEGasAge.TextMatrix(i, 2)
                lsMatrizDatos(3, nContadorTemp) = FEGasAge.TextMatrix(i, 3)
                lsMatrizDatos(4, nContadorTemp) = 0
            End If
        Next i
    Else
        If ckOpcionActivar.value = 0 Then
            For i = 1 To nPost
                If FEGasAge.TextMatrix(i, 4) = "." Then
                    nContadorTemp = nContadorTemp + 1
                    ReDim Preserve lsMatrizDatos(1 To 4, 1 To i)
                    lsMatrizDatos(1, nContadorTemp) = FEGasAge.TextMatrix(i, 1)
                    lsMatrizDatos(2, nContadorTemp) = FEGasAge.TextMatrix(i, 2)
                    lsMatrizDatos(3, nContadorTemp) = FEGasAge.TextMatrix(i, 3)
                    lsMatrizDatos(4, nContadorTemp) = 0
                End If
            Next i
        Else
            If txtMonto.Text = "" Or txtTotal.Text = "" Then
                 If txtMonto.Text = "" Then
                    MsgBox "Ingresar monto de distribución"
                    Exit Sub
                 End If
                 If txtTotal.Text = "" Then
                    MsgBox "Ingresar monto total"
                    Exit Sub
                 End If
            End If
            Dim nContadorPorcentaje As Integer
            Dim nPorcentajeSubTotal As Currency
            For i = 1 To nPost
                If FEGasAge.TextMatrix(i, 4) = "." Then
                    nContadorPorcentaje = nContadorPorcentaje + FEGasAge.TextMatrix(i, 3)
                End If
               ' nPorcentajeSubTotal = CDbl(txtMonto.Text / txtTotal.Text)
                
            Next i
             nPorcentajeSubTotal = CDbl(txtMonto.Text / txtTotal.Text)
            For i = 1 To nPost
            If FEGasAge.TextMatrix(i, 4) = "." Then
                    nContadorTemp = nContadorTemp + 1
                    ReDim Preserve lsMatrizDatos(1 To 4, 1 To i)
                    lsMatrizDatos(1, nContadorTemp) = FEGasAge.TextMatrix(i, 1)
                    lsMatrizDatos(2, nContadorTemp) = FEGasAge.TextMatrix(i, 2)
                    lsMatrizDatos(3, nContadorTemp) = FEGasAge.TextMatrix(i, 3) * nPorcentajeSubTotal
                    lsMatrizDatos(4, nContadorTemp) = 0
                End If
                Next i
                    Dim oDAgencia As DAgencia
                    Set oDAgencia = New DAgencia
                    nContadorTemp = nContadorTemp + 1
                    ReDim Preserve lsMatrizDatos(1 To 4, 1 To i)
                    lsMatrizDatos(1, nContadorTemp) = txtAgeCod.Text
                    lsMatrizDatos(2, nContadorTemp) = oDAgencia.GetAgencias(txtAgeCod.Text)
                    lsMatrizDatos(3, nContadorTemp) = ((1 - nPorcentajeSubTotal) * 100)
                    lsMatrizDatos(4, nContadorTemp) = 0
                    Set oDAgencia = Nothing
                    lnMontoGasto = CDbl(txtTotal.Text)
            End If
    End If
    End If
    lnMontoInafecto = CDbl(IIf(Me.txtInafecto.Text = "", 0, Me.txtInafecto.Text))
    lnSeleAge = nContadorTemp
    '*** PEAC 20100713
    If lbCtaCtrapDistMan Then
        'Me.txtCta.Visible = True
        lsCtaContrapDistriAstoManual = Trim(txtCta.Text)
    Else
        'Me.txtCta.Visible = False
        lsCtaContrapDistriAstoManual = ""
    End If
    
    '*** FIN PEAC
    
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
txtMonto.Text = ""
ckOpcionActivar.value = 0
txtTotal.Text = ""
txtAgeCod.Text = ""

lnSeleAge = 0

End Sub

Private Sub cmdSalir_Click()
    lnSeleAge = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Call LLamarCargarDatosRs
    lnTipoMatriz = 1
    optPorcentaje.value = True
    
    '*** PEAC 20100713
    Dim clsCta As DCtaCont
    Set clsCta = New DCtaCont
    'txtCta.rs = clsCta.CargaCtaCont(" cCtaContCod LIKE '[12]_[^0]%' ", "CtaCont")
    txtCta.rs = clsCta.CargaCtaCont(" cCtaContCod LIKE '%' ", "CtaCont")
    txtCta.lbUltimaInstancia = False
    txtCta.EditFlex = False
    txtCta.TipoBusqueda = BuscaGrid
    '*** FIN PEAC
    
End Sub

Private Sub cmdCargarDatos(rs As ADODB.Recordset)
    Dim i As Integer
    If nPost > 0 Then
        For i = 1 To nPost
            FEGasAge.EliminaFila (1)
        Next i
    End If
    nPost = 0
    rs.MoveFirst
    If (rs.EOF Or rs.BOF) Then
        MsgBox "No existen porcenctajes de gastos de Agencias"
        Exit Sub
    End If
    nPost = 0
    
    Do While Not (rs.EOF Or rs.BOF)
        nPost = nPost + 1
        FEGasAge.AdicionaFila
        
        ReDim Preserve sMatrizTemp(1 To 3, 1 To nPost)
        
        '*** PEAC 20100712
'        FEGasAge.TextMatrix(nPost, 0) = "1"
'        FEGasAge.TextMatrix(nPost, 1) = rs!cAgecod
'        FEGasAge.TextMatrix(nPost, 2) = rs!cAgeDescripcion
'        FEGasAge.TextMatrix(nPost, 3) = rs!nAgePorcentaje
'        FEGasAge.TextMatrix(nPost, 4) = 1
        
        FEGasAge.TextMatrix(nPost, 0) = "1"
        FEGasAge.TextMatrix(nPost, 1) = rs!cAgecod
        FEGasAge.TextMatrix(nPost, 2) = rs!cAgeDescripcion
        If Me.optCartCred.value = True Then
            FEGasAge.TextMatrix(nPost, 3) = rs!nAgePorcentaje
        ElseIf Me.optCartAho.value = True Then
            FEGasAge.TextMatrix(nPost, 3) = rs!nPorcenCarteraAho
        Else
            FEGasAge.TextMatrix(nPost, 3) = rs!nPorcenIngFinan
        End If
            FEGasAge.TextMatrix(nPost, 4) = 1
        '*** FIN PEAC
        
        sMatrizTemp(3, nPost) = rs!nAgePorcentaje
        
        rs.MoveNext
    Loop
    
End Sub
Private Sub LLamarCargarDatosRs()
Dim obDAgencia As DAgencia
    Set obDAgencia = New DAgencia
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = obDAgencia.GetAgenciaPorcentajeGastosxValor
    Call cmdCargarDatos(rs)
End Sub

Private Sub optCantitra_Click()
    Call OptionDatos
End Sub

Private Sub Option1_Click()
    
End Sub

Private Sub optCartAho_Click()
    If optPorcentaje.value = True Then
        Call LLamarCargarDatosRs
        lnTipoMatriz = 1
    End If
End Sub

Private Sub optCartCred_Click()
    If optPorcentaje.value = True Then
        Call LLamarCargarDatosRs
        lnTipoMatriz = 1
    End If
End Sub

Private Sub optIngFin_Click()
    If optPorcentaje.value = True Then
        Call LLamarCargarDatosRs
        lnTipoMatriz = 1
    End If
End Sub

Private Sub optPorcentaje_Click()
    Call OptionDatos
End Sub
Private Sub CargarDatosCantidad()
        Dim i As Integer
        If nPost > 0 Then
           For i = 1 To nPost
               FEGasAge.TextMatrix(i, 3) = 0
           Next i
        End If
End Sub
Private Sub CargarDatosVaciaPorcentaje()
        Dim i As Integer
        If nPost > 0 Then
           For i = 1 To nPost
               FEGasAge.TextMatrix(i, 3) = sMatrizTemp(3, i)
           Next i
        End If
End Sub
