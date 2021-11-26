VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgosAutorizacionListado 
   Caption         =   "Listado del Control de Autorizaciones"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14490
   Icon            =   "frmCredRiesgosAutorizacionListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab stabAutorizaciones 
      DragIcon        =   "frmCredRiesgosAutorizacionListado.frx":030A
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Listado de Control de Autorizaciones"
      TabPicture(0)   =   "frmCredRiesgosAutorizacionListado.frx":0614
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdMostrar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameCabecera"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameAutorizaciones"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdEditar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdGuardar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   12720
         TabIndex        =   17
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   360
         Left            =   11160
         TabIndex        =   16
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   360
         Left            =   9600
         TabIndex        =   15
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Frame frameAutorizaciones 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3855
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   14160
         Begin SICMACT.FlexEdit feAutorizacion 
            Height          =   3255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   13935
            _extentx        =   24580
            _extenty        =   5741
            cols0           =   14
            highlight       =   1
            encabezadosnombres=   "#-Autorizacion / Tpo.Cred-Util-Saldo-Total-Util-Saldo-Total-Util-Saldo-Total-Util-Saldo-Total"
            encabezadosanchos=   "250-5000-700-700-700-700-700-700-700-700-700-700-700-700"
            font            =   "frmCredRiesgosAutorizacionListado.frx":0630
            fontfixed       =   "frmCredRiesgosAutorizacionListado.frx":0658
            columnasaeditar =   "X-X-X-X-4-X-X-7-X-X-10-X-X-13"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            encabezadosalineacion=   "R-L-R-R-R-R-C-R-C-C-R-C-C-R"
            formatosedit    =   "3-0-2-2-3-2-0-3-0-0-3-0-0-3"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            lbultimainstancia=   -1  'True
            tipobusqueda    =   6
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   255
            rowheight0      =   300
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CONSUMO"
            Height          =   255
            Left            =   5400
            TabIndex        =   12
            Top             =   240
            Width           =   2150
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HIPOTECARIO"
            Height          =   255
            Left            =   7520
            TabIndex        =   11
            Top             =   240
            Width           =   2145
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MINORISTA"
            Height          =   255
            Left            =   9645
            TabIndex        =   10
            Top             =   240
            Width           =   2105
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NO MINORISTA"
            Height          =   255
            Left            =   11750
            TabIndex        =   9
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.Frame frameCabecera 
         Caption         =   "Buscar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   9135
         Begin VB.ComboBox cmbAgencia 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   195
            Width           =   2745
         End
         Begin VB.ComboBox cmbMes 
            Height          =   315
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   195
            Width           =   2025
         End
         Begin VB.ComboBox cboAnio 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   195
            Width           =   1545
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agencia:"
            Height          =   195
            Left            =   480
            TabIndex        =   14
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año:"
            Height          =   195
            Left            =   4200
            TabIndex        =   6
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes:"
            Height          =   195
            Left            =   6360
            TabIndex        =   5
            Top             =   240
            Width           =   345
         End
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   360
         Left            =   9720
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCredRiesgosAutorizacionListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre      : frmCredRiesgosAutorizacionMensual                                             *
'** Descripción : Formulario para el Listado de las Autorizaciones                              *
'** Referencia  : ERS065-2016                                                                    *
'** Creación    : ARLO, 20170601 09:00:00 AM                                                    *
'************************************************************************************************
Option Explicit
Dim i, j As Integer
Dim cols As Integer
Dim rows As Integer

Public Function Inicio()
    Call CargarAño
    Call CargaMes
    Call CargaAgencia
    Call CargaAutorizacion
    Me.cmdEditar.Enabled = False
    Me.cmdGuardar.Enabled = False
    Me.feAutorizacion.Enabled = False
    frmCredRiesgosAutorizacionListado.Show 1
End Function
Public Function InicioLectura()
    Call CargarAño
    Call CargaMes
    Call CargaAgencia
    Call CargaAutorizacion
    Me.cmdEditar.Visible = False
    Me.cmdGuardar.Visible = False
    Me.feAutorizacion.Enabled = False
    frmCredRiesgosAutorizacionListado.Show 1
End Function

Private Sub CargarAño()
    
    Dim nAño As String
    Dim nAño2 As Integer
    Dim nAñoIncio As Integer
    
    nAñoIncio = 2016
    nAño = Year(gdFecSis)
    nAño2 = CInt(nAño)
        Do
        cboAnio.AddItem "" & nAñoIncio
        cboAnio.ItemData(cboAnio.NewIndex) = "" & nAñoIncio
        nAñoIncio = nAñoIncio + 1
        Loop While (nAñoIncio <= nAño2 + 1)
        cboAnio.ListIndex = 0
End Sub
Public Function CargaMes() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = DevuelveMes()
    
    CargarComboBox rs, cmbMes
    
    rs.Close
    Set rs = Nothing
    End Function
Public Function DevuelveMes() As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaMes"
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveMes = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function CargaAgencia() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = DevuelveAgencia()
    
    CargarComboBox rs, Me.cmbAgencia
    
    rs.Close
    Set rs = Nothing
    End Function
    Public Function DevuelveAgencia() As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaAgencia"
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveAgencia = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function CargaAutorizacion() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim lnAgeCodAct As Integer
    Set rs = New ADODB.Recordset
    Dim nNumFila As Integer
    
    Set rs = DevuelveAutorizacion()
    
    Do While Not rs.EOF
            
            feAutorizacion.AdicionaFila
            nNumFila = feAutorizacion.rows - 1
            feAutorizacion.TextMatrix(nNumFila, 0) = nNumFila
            feAutorizacion.TextMatrix(nNumFila, 1) = rs!cDescripcion
             rs.MoveNext
      Loop
    rs.Close
    Set rs = Nothing
    End Function
    Public Function DevuelveAutorizacion() As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaAutorizacionAnual"
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveAutorizacion = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdEditar_Click()

    Me.feAutorizacion.Enabled = True
    Me.cmdGuardar.Enabled = True
    Me.cmdEditar.Enabled = False
End Sub

Private Sub Cmdguardar_Click()
Dim nMes As Integer
Dim nAño As Integer
Dim cAgeCod As String
Dim nConsValor As String
Dim nTotal As Integer
Dim nTipoCredito As Integer

nAño = cboAnio.ItemData(cboAnio.ListIndex)
nMes = cmbMes.ItemData(cmbMes.ListIndex)
cAgeCod = cmbAgencia.ItemData(cmbAgencia.ListIndex)

If (Len(cAgeCod) = 1) Then
cAgeCod = "0" + cAgeCod
Else
cAgeCod = cAgeCod
End If

If cboAnio.ListIndex = -1 Then
        MsgBox "Eliga un Año, Por Favor", vbInformation, "Aviso"
Exit Sub
End If

If cmbMes.ListIndex = -1 Then
        MsgBox "Eliga una Mes, Por Favor", vbInformation, "Aviso"
Exit Sub
End If

'Primera Matriz
cols = feAutorizacion.cols - 2
rows = feAutorizacion.rows - 1
nTipoCredito = 0
nConsValor = 1
For j = 1 To rows
        For i = 1 To cols
        nTotal = feAutorizacion.TextMatrix(j, i + 3)
        i = i + 2
        nTipoCredito = nTipoCredito + 1
        If (nTipoCredito = 5) Then
        nTipoCredito = 1
        End If
        Call AutulizaDatosMensual(cAgeCod, nAño, nMes, nConsValor, nTipoCredito, nTotal)
    Next i
    nConsValor = nConsValor + 1
Next j

MsgBox "Se Registraron los Datos Correctamente", vbInformation, "Aviso"

    Me.feAutorizacion.Clear
    Me.feAutorizacion.rows = 2
    Me.feAutorizacion.FormaCabecera
    Me.feAutorizacion.Enabled = True
    Call CargaAutorizacion
    Me.cmdEditar.Enabled = False
    Me.cmdGuardar.Enabled = False
    
End Sub

Private Sub cmdMostrar_Click()
Dim rs, rs1 As ADODB.Recordset
Dim nMes As Integer
Dim nAño As Integer
Dim cAgeCod As String

nAño = cboAnio.ItemData(cboAnio.ListIndex)
nMes = cmbMes.ItemData(cmbMes.ListIndex)
cAgeCod = cmbAgencia.ItemData(cmbAgencia.ListIndex)
If (Len(cAgeCod) = 1) Then
cAgeCod = "0" + cAgeCod
Else
cAgeCod = cAgeCod
End If

Call CargaListaMes(nAño, nMes, cAgeCod)

Set rs = DevuelveListaLimites(nAño, nMes, cAgeCod)

     If (rs.RecordCount > 0) Then
            If (DateDiff("m", gdFecSis, rs!dFechaReg) = 0) Then
            Me.cmdEditar.Enabled = True
            Else
            Me.cmdEditar.Enabled = False
            Me.cmdGuardar.Enabled = False
            End If
    Else
    End If


End Sub
Public Function CargaListaMes(ByVal nAño As Integer, ByVal nMes As String, ByVal cAgeCod As String) As ADODB.Recordset
    Dim rs, rs1 As ADODB.Recordset
    Dim lnAgeCodAct As Integer
    Set rs = New ADODB.Recordset
    Dim nNumFila As Integer
    
    Set rs = DevuelveListaMes(nAño, nMes, cAgeCod)

    cols = feAutorizacion.cols - 2
    rows = feAutorizacion.rows - 1
    
    If (rs.RecordCount > 0) Then

    
            For j = 1 To rows
            For i = 1 To cols
            
            feAutorizacion.TextMatrix(j, i + 1) = rs!nUtil
            feAutorizacion.TextMatrix(j, i + 2) = rs!nSaldo
            feAutorizacion.TextMatrix(j, i + 3) = rs!nTotal
            i = i + 2
            rs.MoveNext
            Next i
            Next j
            Me.cmdEditar.Enabled = True
    Else
    MsgBox "No se encontraron Registros", vbInformation, "Aviso"
    Me.feAutorizacion.Clear
    Me.feAutorizacion.rows = 2
    Me.feAutorizacion.FormaCabecera
    Me.feAutorizacion.Enabled = False
    Me.cmdEditar.Enabled = False
    Me.cmdGuardar.Enabled = False
    Call CargaAutorizacion
    End If
    rs.Close
    Set rs = Nothing
    End Function
 Public Function DevuelveListaMes(ByVal nAño As Integer, ByVal nMes As Integer, ByVal cAgeCod As String) As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_Sel_ERS0652016_ListaAutorizacionXMes '" & cAgeCod & " '," & nAño & "," & nMes
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveListaMes = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub AutulizaDatosMensual(ByVal cAgeCod As String, ByVal nAño As Integer, ByVal nMes As Integer, _
                                ByVal nConsValor As Integer, ByVal nTipoCredito As Integer, ByVal nTotal As Integer)
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim pbTran As Boolean
    
    lsSQL = "exec stp_Upd_ERS0652016_LimiteAutorizacionXMes '" & cAgeCod & "'," & nAño & "," & nMes & "," & nConsValor & "," & nTipoCredito & "," & nTotal
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.Ejecutar (lsSQL)
    loReg.CierraConexion
End Sub

 Public Function DevuelveListaLimites(ByVal nAño As Integer, ByVal nMes As Integer, ByVal cAgeCod As String) As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_Sel_ERS0652016_validaLimites '" & cAgeCod & " '," & nAño & "," & nMes
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveListaLimites = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Private Sub feAutorizacion_OnCellChange(pnRow As Long, pnCol As Long)
            
        If IsNumeric(feAutorizacion.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feAutorizacion.TextMatrix(pnRow, pnCol) < 0 Or Len(feAutorizacion.TextMatrix(pnRow, pnCol)) > 5 Then
            feAutorizacion.TextMatrix(pnRow, pnCol) = 0
            MsgBox "Por favor, ingrese un numero mayor que cero (0) o menor de 5 digitos.", vbInformation, "Aviso"
        End If
        Else
        feAutorizacion.TextMatrix(pnRow, pnCol) = 0
        End If
            
        feAutorizacion.TextMatrix(pnRow, pnCol - 1) = feAutorizacion.TextMatrix(pnRow, pnCol)

End Sub
