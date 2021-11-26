VERSION 5.00
Begin VB.Form frmPoliSeguPatriDistriPorcen 
   Caption         =   "Distribución de Gastos"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   Icon            =   "frmPoliSeguPatriDistriPorcen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Caption         =   "Marcar Todos"
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Ver suma %"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
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
      Width           =   11055
      Begin Sicmact.FlexEdit FEGasAge 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7011
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-CodAge-DescAge-Cod. Seguro-Tipo Seguro-Porcentaje-Activar"
         EncabezadosAnchos=   "400-1000-3000-1000-3000-1200-1000"
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
         ColumnasAEditar =   "X-X-X-X-X-5-6"
         ListaControles  =   "0-0-0-0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-L-R-C"
         FormatosEdit    =   "0-0-0-0-0-2-0"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmPoliSeguPatriDistriPorcen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsMatrizDatos() As Variant
Dim lnTipoMatriz As Integer
Dim nPost As Integer
Dim J As Integer
Dim FENoMoverdeFila As Integer
Dim sMatrizTemp() As String
Dim nContador As Integer
Dim nPorcentaje As Currency
Dim lnSeleAge As Integer
Dim lnMontoGasto As Currency
Dim lnTipoPoli As Integer

Dim sMatrizDatos() As Variant
Dim nTipoMatriz As Integer
Dim nCantAgeSel As Integer
Dim I As Integer
Dim nValorTotalCanti As Currency


Public Sub Inicio(ByVal nTipoPoli As Integer, ByRef sMatrizDatos As Variant, ByRef nTipoMatriz As Integer, ByRef nSeleAge As Integer)
    
    lnTipoPoli = nTipoPoli
    
    
    lsMatrizDatos = sMatrizDatos
    lnTipoMatriz = nTipoMatriz
    lnSeleAge = nSeleAge
    Show 1
    
    sMatrizDatos = lsMatrizDatos
    nTipoMatriz = lnTipoMatriz
    nSeleAge = lnSeleAge
    
End Sub

Private Sub chktodos_Click()
    If Me.chkTodos.value = 1 Then
        nPost = FEGasAge.Rows - 1
        For I = 1 To nPost
            If FEGasAge.TextMatrix(I, 1) <> "" Then
                FEGasAge.TextMatrix(I, 6) = "1"
            End If
        Next I
    Else
        nPost = FEGasAge.Rows - 1
        For I = 1 To nPost
            If FEGasAge.TextMatrix(I, 1) <> "" Then
                FEGasAge.TextMatrix(I, 6) = "0"
            End If
        Next I
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim nContadorTemp  As Integer

    nContador = 0
    nContadorTemp = 0
    nValorTotalCanti = 0
    
    nPost = FEGasAge.Rows - 1
    
    For I = 1 To nPost
        If FEGasAge.TextMatrix(I, 6) = "." Then
            If CInt(FEGasAge.TextMatrix(I, 5)) <= 0# Then
                MsgBox "Agencia " & Trim(FEGasAge.TextMatrix(I, 1)) & " no tiene valor, revise." ''& 1
                Exit Sub
            End If
            nContador = nContador + 1
            nValorTotalCanti = nValorTotalCanti + FEGasAge.TextMatrix(I, 5)
        End If
    Next I
    
    If nValorTotalCanti <> 100 Then
        MsgBox "Los porcentajes ingresados tienen que dar a 100%, los ingresados suman : " & Trim(CStr(nValorTotalCanti)) & "%" & IIf(nValorTotalCanti > 100#, " demas " & Trim(CStr(nValorTotalCanti - 100#)), " falta " & Trim(CStr(100# - nValorTotalCanti)))
        Exit Sub
    End If
        
    If nContador <> nPost Then
        For I = 1 To nPost
        If FEGasAge.TextMatrix(I, 1) <> "" Then
            If FEGasAge.TextMatrix(I, 6) = "." Then
                nContadorTemp = nContadorTemp + 1
                ReDim Preserve lsMatrizDatos(1 To 6, 1 To I)
                lsMatrizDatos(1, nContadorTemp) = FEGasAge.TextMatrix(I, 1)
                lsMatrizDatos(2, nContadorTemp) = FEGasAge.TextMatrix(I, 2)
                If nValorTotalCanti > 0 Then
                         lsMatrizDatos(3, nContadorTemp) = (FEGasAge.TextMatrix(I, 5)) ''/ nValorTotalCanti) * 100
                    lsMatrizDatos(4, nContadorTemp) = 0
                Else
                    MsgBox "Ingresar los valores en los casilleros"
                    Exit Sub
                End If
            End If
        End If
        Next I
    
    Else
        If lnTipoMatriz <> 1 Then
           For I = 1 To nPost
               If FEGasAge.TextMatrix(I, 6) = "." Then
                   nContadorTemp = nContadorTemp + 1
                   ReDim Preserve lsMatrizDatos(1 To 6, 1 To I)
                   lsMatrizDatos(1, nContadorTemp) = FEGasAge.TextMatrix(I, 1)
                   lsMatrizDatos(2, nContadorTemp) = FEGasAge.TextMatrix(I, 2)
                   lsMatrizDatos(3, nContadorTemp) = FEGasAge.TextMatrix(I, 3)
                   lsMatrizDatos(4, nContadorTemp) = FEGasAge.TextMatrix(I, 4)
                   lsMatrizDatos(5, nContadorTemp) = FEGasAge.TextMatrix(I, 5)
                   lsMatrizDatos(6, nContadorTemp) = 0
               End If
           Next I
        Else
        End If
    End If
    lnSeleAge = nContadorTemp
    Unload Me
End Sub

Private Sub cmdCancelar_Click()

    nValorTotalCanti = 0
    nPost = FEGasAge.Rows - 1
    
    For I = 1 To nPost
        If FEGasAge.TextMatrix(I, 6) = "." Then
            If CInt(FEGasAge.TextMatrix(I, 5)) <= 0# Then
                MsgBox "Agencia " & Trim(FEGasAge.TextMatrix(I, 1)) & " no tiene valor, revise." ''& 1
                Exit Sub
            End If
            nValorTotalCanti = nValorTotalCanti + FEGasAge.TextMatrix(I, 5)
        End If
    Next I
    
    If nValorTotalCanti <> 100 Then
        MsgBox "Los porcentajes ingresados tienen que dar a 100%, los ingresados suman : " & Trim(CStr(nValorTotalCanti)) & "%" & IIf(nValorTotalCanti > 100#, " demas " & Trim(CStr(nValorTotalCanti - 100#)), " falta " & Trim(CStr(100# - nValorTotalCanti)))
    Else
        MsgBox "Porcentajes al 100% OK.", vbOKOnly, "Atención"
    End If

End Sub

Private Sub cmdSalir_Click()
    lnSeleAge = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Call LLamarCargarDatosRs
    
End Sub

Private Sub LLamarCargarDatosRs()
    
    Dim obDAgencia As DAgencia
    Set obDAgencia = New DAgencia
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = obDAgencia.CargaPorcenPoliSeguro(lnTipoPoli)
    Call cmdCargarArch(rs)

End Sub


Private Sub cmdCargarArch(rs As ADODB.Recordset)
        
        Dim I As Integer
        If nPost > 0 Then
            For I = 1 To nPost
                FEGasAge.EliminaFila (1)
            Next I
        End If
        nPost = 0
        If (rs.EOF Or rs.BOF) Then
            MsgBox "No existen porcenctajes de gastos de Agencias"
            Exit Sub
        End If
        rs.MoveFirst
        nPost = 0
        
        Me.FEGasAge.Clear
        Me.FEGasAge.rsFlex = rs

End Sub

Private Sub CargarDatosCantidad()
        Dim I As Integer
        If nPost > 0 Then
           For I = 1 To nPost
               FEGasAge.TextMatrix(I, 3) = 0
           Next I
        End If
End Sub
Private Sub CargarDatosVaciaPorcentaje()
        Dim I As Integer
        If nPost > 0 Then
           For I = 1 To nPost
               FEGasAge.TextMatrix(I, 3) = sMatrizTemp(3, I)
           Next I
        End If
End Sub
