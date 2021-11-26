VERSION 5.00
Begin VB.Form frmViaticoVisitaAgencias 
   Caption         =   "Selección de Agencias"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   Icon            =   "frmViaticoVisitaAgencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Caption         =   "Marcar Todos"
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6600
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
      Width           =   7695
      Begin Sicmact.FlexEdit FEGasAge 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7435
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-CodAge-DescAge-Activar-Importe"
         EncabezadosAnchos=   "400-1000-3000-1000-0"
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
         ListaControles  =   "0-0-0-4-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-R"
         FormatosEdit    =   "0-0-0-0-0"
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
Attribute VB_Name = "frmViaticoVisitaAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''********************************************''
''Formulario:   frmViaticoVisitaAgencias    **''
''Creación:     22/03/2011                  **''
''Programador:  Pedro Acuña                 **''
''********************************************''
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

Dim sMatrizDatos() As Variant
Dim nTipoMatriz As Integer
Dim nCantAgeSel As Integer
Dim I As Integer
Dim lnMontoDistri As Currency

Public Sub Inicio(ByRef sMatrizDatos As Variant, ByRef nTipoMatriz As Integer, ByRef nSeleAge As Integer, Optional ByVal nMontoDistri As Currency = 0)

    lsMatrizDatos = sMatrizDatos
    lnTipoMatriz = nTipoMatriz
    lnSeleAge = nSeleAge
    lnMontoDistri = nMontoDistri
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
                FEGasAge.TextMatrix(I, 3) = "1"
            End If
        Next I
    Else
        nPost = FEGasAge.Rows - 1
        For I = 1 To nPost
            If FEGasAge.TextMatrix(I, 1) <> "" Then
                FEGasAge.TextMatrix(I, 3) = "0"
            End If
        Next I
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim nContadorTemp  As Integer
    Dim lnTotDis As Currency

    nContador = 0
    nContadorTemp = 0
    nPost = FEGasAge.Rows - 1
        
    For I = 1 To nPost
        If FEGasAge.TextMatrix(I, 3) = "." Then
            nContador = nContador + 1
        End If
    Next I
    
    If nContador = 0 Then
        MsgBox "Por favor seleccione al menos una Agencia o cancele la selección.", vbOKOnly + vbExclamation, "Atención"
        Exit Sub
    End If

    If nContador <> nPost Then
        For I = 1 To nPost
        If FEGasAge.TextMatrix(I, 1) <> "" Then
            If FEGasAge.TextMatrix(I, 3) = "." Then
                nContadorTemp = nContadorTemp + 1
                ReDim Preserve lsMatrizDatos(1 To 3, 1 To I)
                lsMatrizDatos(1, nContadorTemp) = FEGasAge.TextMatrix(I, 1)
                lsMatrizDatos(2, nContadorTemp) = FEGasAge.TextMatrix(I, 2)
                lsMatrizDatos(3, nContadorTemp) = FEGasAge.TextMatrix(I, 4)
            End If
        End If
        Next I
    Else
        If lnTipoMatriz <> 1 Then
           For I = 1 To nPost
               If FEGasAge.TextMatrix(I, 3) = "." Then
                   nContadorTemp = nContadorTemp + 1
                   ReDim Preserve lsMatrizDatos(1 To 3, 1 To I)
                   lsMatrizDatos(1, nContadorTemp) = FEGasAge.TextMatrix(I, 1)
                   lsMatrizDatos(2, nContadorTemp) = FEGasAge.TextMatrix(I, 2)
                   lsMatrizDatos(3, nContadorTemp) = FEGasAge.TextMatrix(I, 4)
               End If
           Next I
        Else
        End If
    End If
    
    '***Modificado por ELRO el 20120927, según OYP-RFC111-2012
    'If lnMontoDistri > 0 And (gsOpeCod = "401140" Or gsOpeCod = "402140") Then
    If lnMontoDistri > 0 And (gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401141" Or gsOpeCod = "402141") Then
    '***Fin Modificado por ELRO el 20120927*******************
        lnTotDis = 0
        For I = 1 To nPost
            If FEGasAge.TextMatrix(I, 3) = "." Then
                If IIf(FEGasAge.TextMatrix(I, 4) = "", 0, FEGasAge.TextMatrix(I, 4)) = 0 Then
                    MsgBox "Por favor ingrese un monto en el código de agencia ''" + Trim(FEGasAge.TextMatrix(I, 1)) + "'' que está seleccionado.", vbOKOnly + vbExclamation, "Atención"
                    Exit Sub
                End If
                lnTotDis = lnTotDis + FEGasAge.TextMatrix(I, 4)
            End If
        Next I
            
        If lnMontoDistri <> lnTotDis Then
            MsgBox "Los importes ingresados no cuadran con lo registrado en el Gasto.", vbOKOnly + vbExclamation, "Atención"
            Exit Sub
        End If
    End If
    
    lnSeleAge = nContadorTemp
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    lnSeleAge = 0
    Unload Me
End Sub

Private Sub Form_Load()
    
    Call LLamarCargarDatosRs
    
    For I = 1 To FEGasAge.Rows - 1
          FEGasAge.TextMatrix(I, 4) = 0
    Next I
    
    '***Modificado por ELRO el 20120927, según OYP-RFC111-2012
    'If gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401371" Or gsOpeCod = "402371" Then
    If gsOpeCod = "401140" Or gsOpeCod = "402140" Or gsOpeCod = "401371" Or gsOpeCod = "402371" Or gsOpeCod = "401141" Or gsOpeCod = "402141" Or gsOpeCod = "401344" Then
    '***Fin Modificado por ELRO el 20120927*******************
        Me.FEGasAge.ColWidth(4) = 1200
    Else
        Me.FEGasAge.ColWidth(4) = 0
    End If
    
End Sub

Private Sub LLamarCargarDatosRs()

    Dim obDAgencia As DAgencia
    Set obDAgencia = New DAgencia
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = obDAgencia.CargaViaticoVisitaAgencias
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
            MsgBox "No existen Agencias."
            Exit Sub
        End If
        rs.MoveFirst
        nPost = 0
        
        Me.FEGasAge.Clear
        Me.FEGasAge.rsFlex = rs
End Sub
