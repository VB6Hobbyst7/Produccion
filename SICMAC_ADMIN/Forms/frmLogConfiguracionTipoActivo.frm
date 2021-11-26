VERSION 5.00
Begin VB.Form frmLogConfiguracionTipoActivo 
   Caption         =   "Configuración Tipo Activo"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmLogConfiguracionTipoActivo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Cambiar Montos"
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   4695
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtMonto 
         Height          =   375
         Left            =   3000
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optTCnt 
         Caption         =   "Tasa Cnt."
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optVCnt 
         Caption         =   "Vida Util Cnt."
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optTT 
         Caption         =   "Tasa Tribut."
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optVT 
         Caption         =   "Vida Util Tribut."
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tributario"
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   8760
      Width           =   5295
      Begin VB.TextBox txtMesT 
         Height          =   315
         Left            =   3600
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtTasaT 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Vida Util"
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Tasa"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contable"
      Height          =   855
      Left            =   5760
      TabIndex        =   7
      Top             =   8760
      Width           =   5295
      Begin VB.TextBox txtTasaC 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtMesC 
         Height          =   315
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Tasa"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Vida Util"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.ComboBox cboTipoActivo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo activo"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   795
      End
   End
   Begin Sicmact.FlexEdit FEGasAge 
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7858
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-CodAge-DescAge-Tasa Cnt.%-Vida Util Cnt.-Tasa Tribut.%-Vida Util Tribut.-AF"
      EncabezadosAnchos=   "400-1200-4000-1200-1200-1200-1200-400"
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
      ColumnasAEditar =   "X-X-X-3-4-5-6-X"
      ListaControles  =   "0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-R-R-R-R-C"
      FormatosEdit    =   "0-0-0-2-2-2-2-0"
      lbEditarFlex    =   -1  'True
      lbFlexDuplicados=   0   'False
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmLogConfiguracionTipoActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim nListIndex As Integer
Dim nPost As Integer
Dim j As Integer

Private Sub cboTipoActivo_Click()
    'rs!cConsDescripcion & Space(150 - Len(rs!cConsDescripcion)) & rs!nDepreTipoActivo & Space(5 - Len(rs!nDepreTipoActivo)) & rs!nDeprePorcC & Space(10 - Len(rs!nDeprePorcC)) & rs!nDepreMesC & Space(10 - Len(rs!nDepreMesC)) & rs!nDeprePorcT & Space(10 - Len(rs!nDeprePorcT)) & rs!nDepreMesT & Space(10 - Len(rs!nDepreMesT))

'*** PEAC 20120530
'    txtTasaC.Text = Trim(Mid(cboTipoActivo.Text, 156, 5))
'    txtMesC.Text = Trim(Mid(cboTipoActivo.Text, 166, 3))
'    txtTasaT.Text = Trim(Mid(cboTipoActivo.Text, 176, 5))
'    txtMesT.Text = Trim(Mid(cboTipoActivo.Text, 186, 3))
        
    Dim dLog As DLogDeprecia
    Set dLog = New DLogDeprecia
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = dLog.ObtienePorcentVidaUtlAF(Trim(Right(cboTipoActivo.Text, 4)))
    Call llenarGridRs(rs)
    Set dLog = Nothing
    
'*** FIN PEAC
        
End Sub
'*** PEAC 20120530
Private Sub llenarGridRs(rs As ADODB.Recordset)
        
        Dim I As Integer
        
        If nPost > 0 Then
            For I = 1 To nPost
                FEGasAge.EliminaFila (1)
            Next I
        End If

        nPost = 0
        If (rs.EOF Or rs.BOF) Then
            MsgBox "No existen datos de Agencias"
            Exit Sub
        End If
        rs.MoveFirst
        
        nPost = 0
        Do While Not (rs.EOF Or rs.BOF)
            nPost = nPost + 1
            FEGasAge.AdicionaFila
            FEGasAge.TextMatrix(nPost, 0) = nPost
            FEGasAge.TextMatrix(nPost, 1) = rs!cAgecod
            FEGasAge.TextMatrix(nPost, 2) = rs!cAgeDescripcion
            FEGasAge.TextMatrix(nPost, 3) = IIf(IsNull(rs!nDeprePorcC), 0, rs!nDeprePorcC)
            FEGasAge.TextMatrix(nPost, 4) = IIf(IsNull(rs!nDepreMesC), 0, rs!nDepreMesC)
            FEGasAge.TextMatrix(nPost, 5) = IIf(IsNull(rs!nDeprePorcT), 0, rs!nDeprePorcT)
            FEGasAge.TextMatrix(nPost, 6) = IIf(IsNull(rs!nDepreMesT), 0, rs!nDepreMesT)
            FEGasAge.TextMatrix(nPost, 7) = rs!nDepreTipoActivo
            rs.MoveNext
        Loop
        
End Sub


Private Sub cmdCambiar_Click()

    Dim nNumCol As Integer
    Dim I As Integer

    If Me.FEGasAge.Rows > 2 Then
        If Me.optTCnt.value = True Then
            nNumCol = 3
        ElseIf Me.optVCnt.value = True Then
            nNumCol = 4
        ElseIf Me.optTT.value = True Then
            nNumCol = 5
        ElseIf Me.optVT.value = True Then
            nNumCol = 6
        Else
            MsgBox "Seleccione un tipo de Depreciación", vbInformation + vbOKOnly, "Atención"
            Exit Sub
        End If
        
        For I = 1 To nPost
            FEGasAge.TextMatrix(I, nNumCol) = Round(CDbl(Me.txtMonto.Text), 2)
        Next I
        
    End If
End Sub

Private Sub cmdCancelar_Click()
'    txtTasaC.Text = Trim(Mid(cboTipoActivo.Text, 156, 5))
'    txtMesC.Text = Trim(Mid(cboTipoActivo.Text, 166, 3))
'    txtTasaT.Text = Trim(Mid(cboTipoActivo.Text, 176, 5))
'    txtMesT.Text = Trim(Mid(cboTipoActivo.Text, 186, 3))
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
 
' Dim oLog As DLogDeprecia
' Set oLog = New DLogDeprecia
'  nListIndex = cboTipoActivo.ListIndex
'  'Call oLog.InsertarTipoDepreciacionActivo(Trim(Mid(cboTipoActivo.Text, 151, 4)), txtTasaC.Text, txtMesC.Text, txtTasaT.Text, txtMesT.Text)
' Set oLog = Nothing
' Call ActualizarCombo
' Call cmdCancelar_Click
' cboTipoActivo.ListIndex = nListIndex
' MsgBox "Tipo de Activo fue actualizado correctamente", vbCritical

'-----------------
'    Dim obDAgencia As DAgencia
'    Set obDAgencia = New DAgencia
    
    Dim oLog As DLogDeprecia
    Set oLog = New DLogDeprecia
    
    Dim nDescuadreEncontrado As Integer, I As Integer, j As Integer, K As Integer
    Dim nSumPorcen As Currency
    
    
    If FEGasAge.Rows <= 2 Then
        MsgBox "No existe datos cargados.", vbApplicationModal + vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If
    
    'If nPost > 0 Then
    If nPost > 0 Then
    
    '*********************** BEGIN valida los porcentajes al 100% - PEAC 20100708
    '-- 1 creditos 2 ahorros 3 financiero
    nDescuadreEncontrado = 0
    For I = 3 To 6
        nSumPorcen = 0
        
        For K = 1 To nPost
            'nSumPorcen = nSumPorcen + FEGasAge.TextMatrix(K, i)
            If FEGasAge.TextMatrix(K, I) = 0 Then
                nSumPorcen = nSumPorcen + 1
            End If
        Next K
        
        'If nSumPorcen <> 100 Then
        If nSumPorcen > 0 Then
            MsgBox "Los montos de la columna ''" + FEGasAge.TextMatrix(0, I) + "'' tienen montos en cero, por favor revise. ", vbOKOnly + vbExclamation, "Atencion"
            nDescuadreEncontrado = 1
        End If

    Next I
    If nDescuadreEncontrado = 1 Then Exit Sub
    '*********************** END valida los porcentajes al 100%
        
    For j = 1 To nPost
        
        '*** PEAC 20100708
        'obDAgencia.ActualizarAgenciaPorcentajeGastos FEGasAge.TextMatrix(J, 1), FEGasAge.TextMatrix(J, 3)
        'obDAgencia.ActualizarAgenciaPorcentajeGastos FEGasAge.TextMatrix(j, 1), FEGasAge.TextMatrix(j, 3), FEGasAge.TextMatrix(j, 4), FEGasAge.TextMatrix(j, 5)
        Call oLog.ActualizaInsertaPorcenDepreAF(CInt(Trim(Right(cboTipoActivo.Text, 4))), FEGasAge.TextMatrix(j, 1), FEGasAge.TextMatrix(j, 3), FEGasAge.TextMatrix(j, 4), FEGasAge.TextMatrix(j, 5), FEGasAge.TextMatrix(j, 6))
        
        
    Next j
    MsgBox "Datos se registraron correctamente", vbApplicationModal
    End If
    
    
    FEGasAge.Clear
    FEGasAge.FormaCabecera
    FEGasAge.Rows = 2

'---------------------------------


End Sub
Private Sub ActualizarCombo()
    Dim dLog As DLogDeprecia
    Set dLog = New DLogDeprecia
    Set rs = New ADODB.Recordset
    Set rs = dLog.GetDepreciacion
    Call llenarRs(rs)
    Set dLog = Nothing
End Sub
Private Sub Form_Load()
    Call ActualizarCombo
End Sub

Private Sub llenarRs(rs As ADODB.Recordset)
    cboTipoActivo.Clear
    If Not (rs.BOF Or rs.EOF) Then
    Do While Not rs.EOF
        'cboTipoActivo.AddItem rs!cConsDescripcion & Space(150 - Len(rs!cConsDescripcion)) & rs!nDepreTipoActivo & Space(5 - Len(rs!nDepreTipoActivo)) & rs!nDeprePorcC & Space(10 - Len(rs!nDeprePorcC)) & IIf(IsNull(rs!nDepreMesC), 0, rs!nDepreMesC) & Space(10 - Len(IIf(IsNull(rs!nDepreMesC), 0, rs!nDepreMesC))) & rs!nDeprePorcT & Space(10 - Len(rs!nDeprePorcT)) & IIf(IsNull(rs!nDepreMesT), 0, rs!nDepreMesT)
        cboTipoActivo.AddItem rs!cConsDescripcion & Space(150 - Len(rs!cConsDescripcion)) & rs!nDepreTipoActivo '& Space(5 - Len(rs!nDepreTipoActivo)) & rs!nDeprePorcC & Space(10 - Len(rs!nDeprePorcC)) & IIf(IsNull(rs!nDepreMesC), 0, rs!nDepreMesC) & Space(10 - Len(IIf(IsNull(rs!nDepreMesC), 0, rs!nDepreMesC))) & rs!nDeprePorcT & Space(10 - Len(rs!nDeprePorcT)) & IIf(IsNull(rs!nDepreMesT), 0, rs!nDepreMesT)
        rs.MoveNext
    Loop
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then 'KeyAscii <> Asc(".") Then
        If KeyAscii <> Asc(".") Then
            'KeyAscii = 8 es el retroceso o BackSpace
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
       End If
    End If

'If KeyAscii = 13 Then 'cuando se oprime enter ocurre lo de abajo
'    If IsNumeric(Me.txtMonto.Text) Then 'si el text es sólo numérico entonces
'        cmdCambiar.SetFocus  'pase al text o objeto que se desee
'    Else 'sino
'        MsgBox "SÓLO SE PERMITE CARACTERES NUMÉRICOS", 0 + 48, "ERROR" 'mensaje que se desee
'    End If 'fin si
'End If 'fin si

End Sub
