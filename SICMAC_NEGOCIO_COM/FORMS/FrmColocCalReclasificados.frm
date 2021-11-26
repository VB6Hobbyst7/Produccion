VERSION 5.00
Begin VB.Form FrmColocCalReclasificados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reclasificacion de Creditos Mes a Comercial"
   ClientHeight    =   5070
   ClientLeft      =   2265
   ClientTop       =   2865
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmColocCalReclasificados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7890
   Begin SICMACT.TxtBuscar txtBuscarPers 
      Height          =   330
      Left            =   5730
      TabIndex        =   10
      Top             =   540
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
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
      TipoBusqueda    =   3
      sTitulo         =   ""
   End
   Begin VB.CheckBox chkCOM_MES 
      Caption         =   "Reclasificar Comercial a Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4965
      TabIndex        =   9
      Top             =   180
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   7680
      Begin VB.CommandButton CmdReclasifica 
         Caption         =   "&Reclasifica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2175
         TabIndex        =   5
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   200
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         Top             =   200
         Width           =   1095
      End
   End
   Begin SICMACT.FlexEdit Flex 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   5953
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "--CodPers-Nombre-Saldo Cap-Cred Rec"
      EncabezadosAnchos=   "150-250-1200-4000-1200-800"
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
      ColumnasAEditar =   "X-1-X-X-X-X"
      ListaControles  =   "0-4-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-R-L"
      FormatosEdit    =   "0-0-0-0-4-0"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   150
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   510
      Width           =   975
   End
   Begin VB.Label LblTipoCambio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3.458"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cambio"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label LblMestitulo 
      AutoSize        =   -1  'True
      Caption         =   "Creditos Reclasificados al Mes de"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3075
   End
End
Attribute VB_Name = "FrmColocCalReclasificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub chkCOM_MES_Click()
If Me.chkCOM_MES.value = 1 Then
    Me.txtBuscarPers.Visible = True
    Me.cmdBuscar.Visible = False
Else
    Me.txtBuscarPers.Visible = False
    Me.cmdBuscar.Visible = True
End If
End Sub

Private Sub cmdBuscar_Click()
Dim Riesgo As COMDCredito.DCOMColocEval
Dim Rcc As COMDCredito.DCOMColocEval
Dim rs As ADODB.Recordset
Dim sql As String
Dim DataConsol As String
Dim DataRCC As String
Set Rcc = New COMDCredito.DCOMColocEval
Set Riesgo = New COMDCredito.DCOMColocEval

DataConsol = Rcc.ServConsol(43)
DataRCC = Rcc.ServConsol(144) '& ".."
Set rs = Riesgo.GetCreditodMesReclasificados(DataConsol, CDbl(Me.LblTipoCambio), DataRCC, IIf(chkCOM_MES = 1, True, False))
If Not (rs.EOF And rs.BOF) Then
    While Not rs.EOF
        With Flex
            .AdicionaFila
            .TextMatrix(.Rows - 1, 1) = 1
            .TextMatrix(.Rows - 1, 2) = rs!cPersCod
            .TextMatrix(.Rows - 1, 3) = rs!cPersNombre
            .TextMatrix(.Rows - 1, 4) = rs!nEndeudaDol
        End With
        rs.MoveNext
    Wend
Else
    MsgBox "No se encontraron Creditos Reclasificados", vbInformation, "AVISO"
End If
Set rs = Nothing
Set Rcc = Nothing
Set Riesgo = Nothing
End Sub

Private Sub cmdCancelar_Click()
Flex.Rows = 2
Flex.Clear
Flex.FormaCabecera
End Sub

Private Sub CmdReclasifica_Click()
Dim I As Integer
Dim Riesgo As COMDCredito.DCOMColocEval
Dim Mo As COMDMov.DCOMMov
Dim cMovNro As String
Dim nValida As Integer
Set Riesgo = New COMDCredito.DCOMColocEval
Set Mo = New COMDMov.DCOMMov

For I = 1 To Flex.Rows - 1
    If Flex.TextMatrix(I, 1) = "." Then
        nValida = Riesgo.NuevoCredReclasificado(gdFecSis, Flex.TextMatrix(I, 2), gsCodCMAC, IIf(chkCOM_MES = 1, gColPYMEEmp, gColComercEmp), cMovNro, gsCodAge, gsCodUser)
        If nValida = -1 Then
            Flex.TextMatrix(I, 5) = "Error"
        Else
            Flex.TextMatrix(I, 5) = nValida
        End If
    End If
Next I
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oGen As COMDConstSistema.DCOMGeneral
Me.LblMestitulo.Caption = "Creditos Reclasificados al Mes de " & UCase(Format(gdFecData, "MMMM"))

Set oGen = New COMDConstSistema.DCOMGeneral
Me.LblTipoCambio.Caption = oGen.EmiteTipoCambio(gdFecSis, TCFijoMes)
Set oGen = Nothing
End Sub

Private Sub txtBuscarPers_EmiteDatos()
'Dim oCon As DConecta
'Dim Sql As String
Dim rs As ADODB.Recordset
Dim oEval As COMDCredito.DCOMColocEval

If txtBuscarPers <> "" Then
    'Set oCon = New DConecta
    'Sql = "SELECT CPERSCOD, CPERSNOMBRE, 0 AS nEndeudaDol  FROM PERSONA  WHERE CPERSCOD ='" & txtBuscarPers & "'"
    'oCon.AbreConexion
    'Set Rs = oCon.CargaRecordSet(Sql)
    Set oEval = New COMDCredito.DCOMColocEval
    Set rs = oEval.ObtenerPersonas(txtBuscarPers.Text)
    Set oEval = Nothing
    If Not (rs.EOF And rs.BOF) Then
        While Not rs.EOF
            With Flex
                .AdicionaFila
                .TextMatrix(.Rows - 1, 1) = 1
                .TextMatrix(.Rows - 1, 2) = rs!cPersCod
                .TextMatrix(.Rows - 1, 3) = rs!cPersNombre
                .TextMatrix(.Rows - 1, 4) = rs!nEndeudaDol
            End With
            rs.MoveNext
        Wend
    Else
        MsgBox "No se encontraron Creditos Reclasificados", vbInformation, "AVISO"
    End If
    rs.Close
    Set rs = Nothing
    
    'oCon.CierraConexion
End If
End Sub
