VERSION 5.00
Begin VB.Form frmCredCodModular 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizacion de Codigos Modulares"
   ClientHeight    =   5760
   ClientLeft      =   915
   ClientTop       =   2010
   ClientWidth     =   10515
   Icon            =   "frmCredCodModular.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   9165
      TabIndex        =   8
      Top             =   5040
      Width           =   1260
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   420
      Left            =   7770
      TabIndex        =   7
      Top             =   5040
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Caption         =   "Institucion"
      Height          =   660
      Left            =   105
      TabIndex        =   3
      Top             =   45
      Width           =   7140
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   360
         Left            =   5985
         TabIndex        =   5
         Top             =   210
         Width           =   1020
      End
      Begin VB.ComboBox CboInst 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   5730
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   645
      Left            =   7275
      TabIndex        =   0
      Top             =   45
      Width           =   2580
      Begin VB.OptionButton OptBusq 
         Caption         =   "Agencias"
         Height          =   240
         Index           =   1
         Left            =   1095
         TabIndex        =   2
         Top             =   285
         Width           =   1425
      End
      Begin VB.OptionButton OptBusq 
         Caption         =   "Local"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   285
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin SICMACT.FlexEdit FEPagoLote 
      Height          =   4080
      Left            =   165
      TabIndex        =   6
      Top             =   855
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   7197
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-OK-Credito-Cod. Modular-Titular-Cargo-Carben-Tipo Planilla"
      EncabezadosAnchos=   "400-400-2000-2000-4000-1200-1200-1200"
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
      ColumnasAEditar =   "X-1-X-3-X-5-6-7"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0"
      AvanceCeldas    =   1
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483635
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmCredCodModular.frx":030A
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   405
      TabIndex        =   9
      Top             =   5025
      Width           =   5550
   End
End
Attribute VB_Name = "frmCredCodModular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset

Private Sub CargaControles()
    
Dim oPersonas As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset

    'Carga Instituciones
    'Call CargaComboPersonasTipo(gPersTipoConvenio, CboInst)
    
    Set oPersonas = New COMDPersona.DCOMPersonas
    Set rs = oPersonas.RecuperaPersonasTipo(gPersTipoConvenio)
    Set oPersonas = Nothing
    
    Do While Not rs.EOF
        CboInst.AddItem PstaNombre(rs!cPersNombre) & Space(250) & rs!cPersCod
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub cmdAplicar_Click()
Dim sCadAge As String
Dim oCred As COMDCredito.DCOMCreditos

    If Me.OptBusq(1).value Then
        sCadAge = frmSelectAgencias.RecupAgencias
    Else
        sCadAge = "('" & gsCodAge & "')"
    End If
    Set oCred = New COMDCredito.DCOMCreditos
    Set R = oCred.RecuperaCreditosCodigoModular(Trim(Right(CboInst.Text, 15)), sCadAge)
    Set oCred = Nothing
    Call LimpiaFlex(FEPagoLote)
    FEPagoLote.FormaCabecera
    Set FEPagoLote.Recordset = R
End Sub

Private Sub CmdGrabar_Click()
Dim oCred As COMDCredito.DCOMCredActBD
Dim I As Integer
Dim MatDatos() As String
Dim MatIndex() As Integer

    If MsgBox("Se va a Grabar las Modificaciones, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    'Set oCred = New COMDCredito.DCOMCredActBD
    'oCred.dBeginTrans
    ReDim MatIndex(0)
    For I = 1 To Me.FEPagoLote.Rows - 1
        If Me.FEPagoLote.TextMatrix(I, 1) = "." Then
            ReDim Preserve MatIndex(UBound(MatIndex) + 1)
            MatIndex(UBound(MatIndex) - 1) = I
        '    Call oCred.dUpdateColocacConvenio(FEPagoLote.TextMatrix(I, 2), , False, FEPagoLote.TextMatrix(I, 3))
        End If
    Next I
    ReDim MatDatos(UBound(MatIndex), 5)
    For I = 0 To UBound(MatDatos) - 1
        MatDatos(I, 0) = FEPagoLote.TextMatrix(MatIndex(I), 2)
        MatDatos(I, 1) = FEPagoLote.TextMatrix(MatIndex(I), 3)
        MatDatos(I, 2) = FEPagoLote.TextMatrix(MatIndex(I), 5)
        MatDatos(I, 3) = FEPagoLote.TextMatrix(MatIndex(I), 6)
        MatDatos(I, 4) = FEPagoLote.TextMatrix(MatIndex(I), 7)
    Next I
    'oCred.dCommitTrans
    Set oCred = New COMDCredito.DCOMCredActBD
    Call oCred.ActualizaCodigoModular(MatDatos)
    Set oCred = Nothing
    Call cmdAplicar_Click
End Sub

Private Sub cmdSalir_Click()
    Set frmSelectAgencias = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Call CargaControles
End Sub

Private Sub OptBusq_Click(Index As Integer)
    If OptBusq(1).value Then
        frmSelectAgencias.Show 1
    End If
End Sub
