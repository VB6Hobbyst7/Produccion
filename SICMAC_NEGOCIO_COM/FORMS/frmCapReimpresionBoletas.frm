VERSION 5.00
Begin VB.Form frmCapReimpresionBoletas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REIMPRESION PARA CONTROL DE CALIDAD OPERACIONES"
   ClientHeight    =   7485
   ClientLeft      =   285
   ClientTop       =   2295
   ClientWidth     =   13080
   Icon            =   "frmCapReimpresionBoletas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   13080
   Begin VB.CheckBox chkTodaOperacion 
      Caption         =   "Imprimir Toda la Operacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   165
      TabIndex        =   8
      Top             =   75
      Width           =   3015
   End
   Begin VB.TextBox txtNota 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   705
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmCapReimpresionBoletas.frx":030A
      Top             =   6600
      Width           =   10260
   End
   Begin VB.Frame Frame1 
      Height          =   6210
      Left            =   105
      TabIndex        =   2
      Top             =   315
      Width           =   12885
      Begin SICMACT.FlexEdit grdCargaDatos 
         Height          =   5895
         Left            =   75
         TabIndex        =   3
         Top             =   180
         Width           =   12750
         _ExtentX        =   22490
         _ExtentY        =   10398
         Cols0           =   10
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nro. Cuenta-Operacion-Monto-MOV:Sald.Cnt.-MOV:Sald.Dis.-Estado-HORA-NroMov-COPECOD"
         EncabezadosAnchos=   "800-2000-3500-1400-1400-1400-700-1200-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
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
         TextStyleFixed  =   1
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColor       =   -2147483639
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-R-R-L-C-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-6-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   0
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   795
         RowHeight0      =   300
         ForeColor       =   -2147483630
         ForeColorFixed  =   -2147483630
         CellForeColor   =   -2147483630
         CellBackColor   =   -2147483639
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   11955
      TabIndex        =   1
      Top             =   6840
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   1050
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Captaciones"
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Prendario"
      Height          =   285
      Left            =   1905
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1560
   End
   Begin SICMACT.Usuario ctlUsuario 
      Left            =   225
      Top             =   5925
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario: <NOMBRE DE USUARIO>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6540
      TabIndex        =   6
      Top             =   75
      Width           =   6390
   End
End
Attribute VB_Name = "frmCapReimpresionBoletas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImprimir_Click()
Dim bTodaop As Boolean
bTodaop = IIf(chkTodaOperacion.value = vbChecked, True, False)
Call ImprimeBoletas(grdCargaDatos.TextMatrix(grdCargaDatos.Row, 2), grdCargaDatos.TextMatrix(grdCargaDatos.Row, 8), grdCargaDatos.TextMatrix(grdCargaDatos.Row, 9), bTodaop)

End Sub
Private Sub ImprimeBoletas(ByVal cOpedesc As String, ByVal nmovnro As Integer, ByVal cOpecod As String, ByVal bTodaop As Boolean)
     
     
End Sub


Private Sub CmdSalir_Click()
 Unload Me
End Sub

Private Sub Form_Load()
Call ctlUsuario.Inicio(gsCodUser)
lblUsuario.Caption = gsCodUser & ": " & UCase(ctlUsuario.UserNom)
MsgBox "Se esta procesando informacion....", vbOKOnly + vbInformation, "AVISO"
Screen.MousePointer = vbHourglass
    CargaOperacUser
Screen.MousePointer = vbDefault

End Sub

Private Sub CargaOperacUser()
 Dim ocon As DConecta
 Set ocon = New DConecta
 Dim rs As ADODB.Recordset, sSql As String
 Set rs = New ADODB.Recordset
 sSql = " declare @fecha datetime "
 sSql = sSql & " set @fecha=cast('" & Format(gdFecSis, "yyyy-mm-dd") & "' as datetime) "
 sSql = sSql & "exec CapGetInfoUser @fecha,'" & gsCodUser & "'"
 
 ocon.AbreConexion
' ocon.ConexionActiva.CommandTimeout = 99999999
 If rs.State = 1 Then rs.Close
 rs.Open sSql, ocon.ConexionActiva, adOpenForwardOnly, adLockOptimistic, adCmdText
     
     grdCargaDatos.Clear
     If rs.State = 1 Then
        If Not (rs.EOF Or rs.BOF) Then
            Set grdCargaDatos.Recordset = rs
        End If
      End If
     grdCargaDatos.FormaCabecera
     
 ocon.CierraConexion
 
 Set ocon = Nothing
 Set rs = Nothing
 
End Sub

Private Sub Label1_Click()

End Sub
