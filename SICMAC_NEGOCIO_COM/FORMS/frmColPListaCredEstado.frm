VERSION 5.00
Begin VB.Form frmColPListaCredEstado 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7230
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6975
      Begin SICMACT.FlexEdit gdLista 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4260
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-CUENTA-NOMBRE"
         EncabezadosAnchos=   "300-2100-4200"
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
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4785
      TabIndex        =   1
      Top             =   3075
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5940
      TabIndex        =   0
      Top             =   3075
      Width           =   1140
   End
End
Attribute VB_Name = "frmColPListaCredEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RGrid As ADODB.Recordset
Dim vsSelecPers As String
Dim vsSelecDacion As Long
Dim vTipoBusqueda As Long
Dim vAgenciaSelect As String
Public vcodper As String


Public Function Inicio(ByVal psCaption As String, ByVal psEstadoCad As String, ByVal psAgecod As String, ByVal pnTpoBusca As Integer, Optional psPersCod As String = "") As String
'RECO20141016 ERS032-2014, SE AGREGO PARAMETRO: psPersCod
    Dim oNCOMColPContrato As COMNColoCPig.NCOMColPContrato
    Set oNCOMColPContrato = New COMNColoCPig.NCOMColPContrato
    Dim i As Integer
    'Set RGrid = oNCOMColPContrato.ListaCredPignoEstado(psEstadoCad, psAgeCod, pnTpoCreAmp)
    'Set DGPersonas.DataSource = RGrid
    Call CargarGrilla(psEstadoCad, psAgecod, pnTpoBusca, psPersCod) 'RECO20141016 ERS.032-2014
    vTipoBusqueda = 1
    Me.Caption = psCaption
    vsSelecPers = ""
    Screen.MousePointer = 0
    Me.Show 1
    
    'Set RGrid = Nothing
    
    Inicio = vsSelecPers
   
End Function

Private Sub CmdAceptar_Click()
    
'    If RGrid Is Nothing Then
'        Unload Me
'        Exit Sub
'    End If
'    If RGrid.RecordCount > 0 Then
'        Select Case vTipoBusqueda
'            Case 1
'                vsSelecPers = RGrid.Fields(0)
'            Case 2
'                vsSelecDacion = RGrid.Fields(0)
'        End Select
'         vcodper = RGrid.Fields(2)
'    Else
'        Select Case vTipoBusqueda
'            Case 1
'                vsSelecPers = ""
'            Case 2
'                vsSelecDacion = -1
'        End Select
'
'    End If
    If gdLista.TextMatrix(1, 1) <> "" Then
        vsSelecPers = gdLista.TextMatrix(gdLista.row, 1)
    Else
        MsgBox "No se encontraron datos.", vbCritical, "Aviso"
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Select Case vTipoBusqueda
        Case 1
            vsSelecPers = ""
        Case 2
            vsSelecDacion = -1
    End Select
    Unload Me
End Sub



Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    cmdAceptar.Default = True
End Sub

Public Sub CargarGrilla(ByVal psEstadoCad As String, ByVal psAgecod As String, ByVal pnTpoBusca As Integer, ByVal psPersCor As String)
    Dim oNCOMColPContrato As COMNColoCPig.NCOMColPContrato
    Dim oDR As ADODB.Recordset
    
    Set oNCOMColPContrato = New COMNColoCPig.NCOMColPContrato
    Set oDR = New ADODB.Recordset
    'cCtaCod , cPersNombre, cPersCod
    Dim i As Integer
    Set oDR = oNCOMColPContrato.ListaCredPignoEstado(psEstadoCad, psAgecod, pnTpoBusca, psPersCor)
    If Not (oDR.EOF And oDR.BOF) Then
        For i = 1 To oDR.RecordCount
            gdLista.AdicionaFila
            gdLista.TextMatrix(i, 1) = oDR!cCtaCod
            gdLista.TextMatrix(i, 2) = oDR!cPersNombre
            oDR.MoveNext
        Next
    End If
    Set oDR = Nothing
End Sub

Private Sub gdLista_DblClick()
    Call CmdAceptar_Click
End Sub

Private Sub gdLista_OnChangeCombo()
    Call CmdAceptar_Click
End Sub

