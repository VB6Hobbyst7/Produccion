VERSION 5.00
Begin VB.Form frmCredPreAprobacionListar 
   Caption         =   "Historial de Niveles de Aprobacion"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   Icon            =   "frmCredPreAprobacionListar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Titular:"
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
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTitular 
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cargos de Aprobacion"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7575
      Begin SICMACT.FlexEdit grdCargo 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   5318
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-CargoCod-Cargo-Estado-Comentario"
         EncabezadosAnchos=   "300-1000-3850-800-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0"
         CantEntero      =   12
         CantDecimales   =   4
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCredPreAprobacionListar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CargaDatos (ActxCta.NroCuenta)
        CargarHistorialNivApr (ActxCta.NroCuenta)
    End If
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim objNivApr As COMDCredito.DCOMNivelAprobacion
    Set objNivApr = New COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = objNivApr.CargarDatosPreApr(psCtaCod, gsCodCargo)
    
    If rs.RecordCount <> "0" Then
        lblTitular = Trim(rs!cPersNombre)
    Else
        MsgBox "No se Encontro el Credito", vbCritical, "Aviso"
    End If

    Set objNivApr = Nothing
    Set rs = Nothing
End Function

Private Sub CargarHistorialNivApr(ByVal cCtaCod As String)
    Dim objNivApr As COMDCredito.DCOMNivelAprobacion
    Set objNivApr = New COMDCredito.DCOMNivelAprobacion
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset

    Dim L As ListItem

    Set rs1 = objNivApr.ListarHistorialNivApr(cCtaCod)
    Set objNivApr = Nothing
    nNroReg = 0
    If Not (rs1.EOF And rs1.BOF) Then
        Set grdCargo.Recordset = rs1
        nNroReg = grdCargo.Rows
        If Not bConsulta Then
            grdCargo.lbEditarFlex = True
        End If
    Else
        If Not bConsulta Then
            grdCargo.lbEditarFlex = True
        End If
    End If
    
    Set objNivApr = Nothing
    Set rs1 = Nothing
    
End Sub

Private Sub Form_Load()
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub
