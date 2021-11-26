VERSION 5.00
Begin VB.Form frmCredGarantiasCob 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobertura de Garantias"
   ClientHeight    =   4140
   ClientLeft      =   3840
   ClientTop       =   3660
   ClientWidth     =   7710
   Icon            =   "frmCredGarantiasCob.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   7440
      Begin VB.ComboBox CboGarantia 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   285
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "Monto Disp."
         Height          =   240
         Left            =   5340
         TabIndex        =   11
         Top             =   720
         Width           =   900
      End
      Begin VB.Label LblDisp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6300
         TabIndex        =   10
         Top             =   705
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "Moneda :"
         Height          =   240
         Left            =   150
         TabIndex        =   9
         Top             =   720
         Width           =   750
      End
      Begin VB.Label LblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   975
         TabIndex        =   8
         Top             =   698
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Monto Garantía"
         Height          =   240
         Left            =   2640
         TabIndex        =   7
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblMontoReal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3915
         TabIndex        =   6
         Top             =   705
         Width           =   990
      End
      Begin VB.Label LblGarantia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2490
         TabIndex        =   5
         Top             =   285
         Width           =   4830
      End
      Begin VB.Label Label1 
         Caption         =   "Garantia :"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2910
      TabIndex        =   1
      Top             =   3615
      Width           =   1545
   End
   Begin SICMACT.FlexEdit FECreditos 
      Height          =   2145
      Left            =   135
      TabIndex        =   0
      Top             =   1410
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   3784
      Cols0           =   7
      EncabezadosNombres=   "-Credito-Estado-Saldo-Cobertura-Ratio-Cobertura Total"
      EncabezadosAnchos=   "400-2000-1800-1200-1200-600-1500"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-C-C-C-R-R"
      FormatosEdit    =   "0-0-0-0-0-2-2"
      SelectionMode   =   1
      ColWidth0       =   405
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmCredGarantiasCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim MatGarantias() As String
Dim MatCredito() As String

Private Sub CboGarantia_Click()
Dim oGar As COMDCredito.DCOMGarantia
Dim i As Integer
Dim R As ADODB.Recordset
    
    On Error GoTo errcombo
    
    Screen.MousePointer = 11
    For i = 0 To UBound(MatGarantias) - 1
        If MatGarantias(i, 0) = CboGarantia.Text Then
            lblGarantia.Caption = MatGarantias(i, 1)
            Me.lblmoneda.Caption = MatGarantias(i, 2)
            lblMontoReal.Caption = Format(MatGarantias(i, 3), gsFormatoNumeroView)
            Me.lblDisp.Caption = Format(MatGarantias(i, 4), gsFormatoNumeroView)
            Exit For
        End If
    Next i

    LimpiaFlex FECreditos
    Set oGar = New COMDCredito.DCOMGarantia
    Set R = oGar.RecuperaGarantiaCreditoDatosVigente(MatGarantias(i, 0))
    i = 0
    Do While Not R.EOF
        If i >= 1 Then
            FECreditos.AdicionaFila
        End If
        FECreditos.TextMatrix(FECreditos.Rows - 1, 1) = R!cCtaCod
        FECreditos.TextMatrix(FECreditos.Rows - 1, 2) = R!cEstado
        FECreditos.TextMatrix(FECreditos.Rows - 1, 3) = Format(R!nSaldo, gsFormatoNumeroView)
        FECreditos.TextMatrix(FECreditos.Rows - 1, 4) = Format(R!nGravado, gsFormatoNumeroView)
        FECreditos.TextMatrix(FECreditos.Rows - 1, 5) = Format(R!nRatio, gsFormatoNumeroView)
        FECreditos.TextMatrix(FECreditos.Rows - 1, 6) = Format(Round(R!nGravado * R!nRatio, 2), gsFormatoNumeroView)
        i = i + 1
        R.MoveNext
    Loop
    R.Close
    Set oGar = Nothing
    Screen.MousePointer = 0
    Exit Sub
errcombo:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub

Public Sub Inicio(ByVal psCtaCod As String)
Dim R As ADODB.Recordset
Dim oCred As COMDCredito.DCOMCredito
Dim i As Integer
    
    On Error GoTo errInicio
    
    Screen.MousePointer = 11
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaGarantiasCredito(psCtaCod)
    Set oCred = Nothing
    ReDim MatGarantias(R.RecordCount, 5)
    i = 0
    Do While Not R.EOF
        MatGarantias(i, 0) = R!cNumGarant
        MatGarantias(i, 1) = R!cTpoGarDescripcion
        MatGarantias(i, 2) = R!nmoneda 'Str(R!nMoneda)
        'MatGarantias(i, 3) = R!nRealizacion
        MatGarantias(i, 3) = R!nValorGarantia
        'MatGarantias(i, 4) = R!nPorGravar
        MatGarantias(i, 4) = R!nDisponible
        i = i + 1
        R.MoveNext
    Loop
    R.Close
    
    'Carga Combo
    Me.CboGarantia.Clear
    For i = 0 To UBound(MatGarantias) - 1
        CboGarantia.AddItem MatGarantias(i, 0)
    Next i
    CboGarantia.ListIndex = -1
    Screen.MousePointer = 0
    
    Me.Show 1
    Exit Sub
errInicio:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

