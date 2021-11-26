VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgosAutorizacionMensual 
   Caption         =   "Registro de Categoría de Agencia (Mensual)"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmAutorizacionRiesgoMensual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab stabAutorizaciones 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Registro Categoría de Agencia (Mensual)"
      TabPicture(0)   =   "frmAutorizacionRiesgoMensual.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameCabecera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdMostrar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameAutorizaciones"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   6720
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   360
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   1200
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "Guardar"
            Height          =   360
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1200
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "Nuevo"
            Height          =   360
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   1200
         End
      End
      Begin VB.Frame frameAutorizaciones 
         Caption         =   "Agencia - Categoría"
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
         Height          =   6405
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   6480
         Begin SICMACT.FlexEdit feAutorizacion 
            Height          =   5895
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   6255
            _extentx        =   11033
            _extenty        =   10398
            cols0           =   5
            highlight       =   1
            encabezadosnombres=   "-Código-Agencia-Categoria-Estado"
            encabezadosanchos=   "0-1000-2500-1200-1000"
            font            =   "frmAutorizacionRiesgoMensual.frx":0326
            fontfixed       =   "frmAutorizacionRiesgoMensual.frx":034E
            columnasaeditar =   "X-X-X-3-4"
            listacontroles  =   "0-0-0-3-4"
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            encabezadosalineacion=   "C-C-L-C-C"
            formatosedit    =   "0-0-0-1-0"
            lbeditarflex    =   -1  'True
            lbultimainstancia=   -1  'True
            tipobusqueda    =   6
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Noviembre"
            Height          =   255
            Left            =   11110
            TabIndex        =   12
            Top             =   2640
            Width           =   2130
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diciembre"
            Height          =   255
            Left            =   13225
            TabIndex        =   11
            Top             =   2640
            Width           =   2130
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Junio"
            Height          =   255
            Left            =   13225
            TabIndex        =   10
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mayo"
            Height          =   255
            Left            =   11110
            TabIndex        =   9
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abril"
            Height          =   255
            Left            =   8995
            TabIndex        =   8
            Top             =   240
            Width           =   2130
         End
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   360
         Left            =   5400
         TabIndex        =   6
         Top             =   600
         Width           =   1215
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5175
         Begin VB.ComboBox cboAnio 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   195
            Width           =   1545
         End
         Begin VB.ComboBox cmbMes 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   195
            Width           =   2265
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes:"
            Height          =   195
            Left            =   2280
            TabIndex        =   5
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   345
         End
      End
   End
End
Attribute VB_Name = "frmCredRiesgosAutorizacionMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre      : frmCredRiesgosAutorizacionMensual                                             *
'** Descripción : Formulario para Registro/ Mantenimiento de Autorizaciones x Producto Mensual  *
'** Referencia  : ERS065-2016                                                                   *
'** Creación    : ARLO, 20170601 09:00:00 AM                                                    *
'************************************************************************************************

Option Explicit
Dim nValorChek As Integer
Dim cols As Integer
Dim rows As Integer
Dim i, J As Integer

Public Function Inicio()
    Call CargarAño
    Call CargaMes
    Call CargaAgencia
    Me.cmdNuevo.Enabled = False
    Me.feAutorizacion.Enabled = False
    Me.cmdGuardar.Enabled = False
    frmCredRiesgosAutorizacionMensual.Show 1
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
Public Function DevuelveCategoria() As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaTpoCreditoxCategoria"
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveCategoria = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function CargaAgencia() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim rsC As ADODB.Recordset
    Dim i As Long
    Dim lnAgeCodAct As Integer
    Set rs = New ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim nNumFila As Integer
    Dim aa As String
    Dim sValor As String
    
    Set rs = DevuelveAgencia()
    'Set rsC = DevuelveCategoria()
     
   'feAutorizacion.CargaCombo rsC
    
    
    Do While Not rs.EOF
            
            feAutorizacion.AdicionaFila
            nNumFila = feAutorizacion.rows - 1
            feAutorizacion.TextMatrix(nNumFila, 1) = rs!cValor
            feAutorizacion.TextMatrix(nNumFila, 2) = rs!cDescripcion
            rs.MoveNext
      Loop

    rs.Close
    Set rs = Nothing
    End Function
    
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Cmdguardar_Click()
Dim lsMes As String
Dim lsAnio As String
Dim lsMovNro As String
Dim dFechaReg As Date
Dim cAgeCod As String
Dim nConsValor As String
Dim nEstado As Integer

lsAnio = cboAnio.ItemData(cboAnio.ListIndex)
lsMes = cmbMes.ItemData(cmbMes.ListIndex)
If (Len(lsMes) = 1) Then
lsMes = "0" + lsMes
Else
lsMes = lsMes
End If
lsMovNro = GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser)
dFechaReg = gdFecSis

If FlexVacio(feAutorizacion) Then
        MsgBox "Por favor, llene todos los datos.", vbInformation, "Aviso"
Exit Sub
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
cols = feAutorizacion.cols - 1
rows = feAutorizacion.rows - 1
For J = 1 To rows
    For i = 1 To cols
        cAgeCod = feAutorizacion.TextMatrix(J, i)
        If (feAutorizacion.TextMatrix(J, i + 2) = "A") Then
        nConsValor = 1
        ElseIf (feAutorizacion.TextMatrix(J, i + 2) = "B") Then
        nConsValor = 2
        ElseIf (feAutorizacion.TextMatrix(J, i + 2) = "C") Then
        nConsValor = 3
        ElseIf (feAutorizacion.TextMatrix(J, i + 2) = "D") Then
        nConsValor = 4
        End If
        If (feAutorizacion.TextMatrix(J, i + 3) = "") Then
        nEstado = 0
        Else
        nEstado = 1
        End If
        i = i + 3
        Call GrabaDatosMensuales(cAgeCod, lsAnio, lsMes, lsMovNro, nConsValor, nEstado)
    Next i
Next J

Call GrabaLimiteAutorizacion(lsAnio, lsMes, lsMovNro)

MsgBox "Se Registraron los Datos Correctamente", vbInformation, "Aviso"

Me.cmdGuardar.Enabled = False

End Sub

Private Sub cmdMostrar_Click()
    Dim nMes As Integer
    Dim nAño As Integer
    
    nAño = cboAnio.ItemData(cboAnio.ListIndex)
    nMes = cmbMes.ItemData(cmbMes.ListIndex)
    Call CargaConfMes(nAño, nMes)
End Sub

Private Sub cmdNuevo_Click()
    Me.feAutorizacion.Enabled = True
    Me.cmdGuardar.Enabled = True
    Me.cmdNuevo.Enabled = False
End Sub

Private Sub feAutorizacion_Click()
    Dim rsC As ADODB.Recordset
    
    Select Case feAutorizacion.Col
    Case 3
    Set rsC = DevuelveCategoria()
    feAutorizacion.CargaCombo rsC
    Set rsC = Nothing
    End Select
    'nValorChek = feAutorizacion.TextMatrix(feAutorizacion.row, 4) = "" 'COMENTADO POR ARLO20170718
End Sub

Private Sub feAutorizacion_OnCellChange(pnRow As Long, pnCol As Long)
Select Case pnCol
    Case 3
        feAutorizacion.TextMatrix(feAutorizacion.row, 3) = Left(UCase(feAutorizacion.TextMatrix(feAutorizacion.row, 3)), 1)
    Case 4
        If IsNumeric(feAutorizacion.TextMatrix(feAutorizacion.row, 4)) Then
            nValorChek = UCase(feAutorizacion.TextMatrix(feAutorizacion.row, 4))
        End If
       feAutorizacion.TextMatrix(feAutorizacion.row, 3) = Left(UCase(feAutorizacion.TextMatrix(feAutorizacion.row, 3)), 1)
    End Select
End Sub

'COMENTADO POR ARLO20170718

'Private Sub feAutorizacion_OnRowChange(pnRow As Long, pnCol As Long)
'
'    feAutorizacion.TextMatrix(feAutorizacion.row - 1, pnCol) = Left(UCase(feAutorizacion.TextMatrix(feAutorizacion.row, pnCol)), 1)
'
'End Sub
'COMENTADO POR ARLO20170718

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
Public Sub GrabaDatosMensuales(ByVal cAgeCod As String, ByVal lsAño As String, ByVal lsMes As String, ByVal lsMovNro As String, _
                                ByVal nConsValor As Integer, ByVal nEstado As Integer)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    
    lsSQL = "exec stp_ins_ERS0652016_AutorizacionXAgeXCategoria '" & cAgeCod & "','" & lsAño & "','" & lsMes & "','" & lsMovNro & "'," & nConsValor & "," & nEstado
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.Ejecutar (lsSQL)
    loReg.CierraConexion
End Sub
    Public Function DevuelveDatosMensual(ByVal nAño As Integer, ByVal nMes As Integer) As ADODB.Recordset
    Dim rsVar As New ADODB.Recordset
    Dim ssql As String, sSqlaux As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    ssql = "stp_sel_ERS0652016_CargaConfTpoXMes " & nAño & "," & nMes
    
    Set rsVar = oCon.CargaRecordSet(ssql)
    Set DevuelveDatosMensual = rsVar
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function CargaConfMes(ByVal nAño As Integer, ByVal nMes As String) As ADODB.Recordset
    Dim rs, rs1 As ADODB.Recordset
    Dim lnAgeCodAct As Integer
    Set rs = New ADODB.Recordset
    Dim nNumFila As Integer
    
    Set rs = DevuelveDatosMensual(nAño, nMes)

    cols = feAutorizacion.cols - 1
    rows = feAutorizacion.rows - 1
    
    If (rs.RecordCount > 0) Then

    
            For J = 1 To rows
                For i = 1 To cols
                    feAutorizacion.TextMatrix(J, i + 2) = rs!nConsValor
                    If (rs!nEstado = 1) Then
                    feAutorizacion.TextMatrix(J, i + 3) = vbChecked
                    Else
                    feAutorizacion.TextMatrix(J, i + 3) = vbUnchecked
                    End If
                    i = i + 3
                    rs.MoveNext
                Next i
            Next J
            Me.cmdNuevo.Enabled = False
            Me.cmdGuardar.Enabled = False
            Me.feAutorizacion.Enabled = False
    Else
    MsgBox "No se encontraron Registros", vbInformation, "Aviso"
    Me.feAutorizacion.Clear
    Me.feAutorizacion.rows = 2
    Me.feAutorizacion.FormaCabecera
    Call CargaAgencia
    Me.cmdNuevo.Enabled = True
    Me.feAutorizacion.Enabled = False
    End If
    rs.Close
    Set rs = Nothing
    End Function

Public Function FlexVacio(ByVal pflex As FlexEdit) As Boolean
    Dim cols As Integer
    Dim rows As Integer
    Dim i, J As Integer
    
    cols = feAutorizacion.cols - 1
    rows = feAutorizacion.rows - 1
    
    For J = 1 To rows
        For i = 1 To cols
            If (feAutorizacion.TextMatrix(J, i + 2) = "") Then
            FlexVacio = True
            Else
            FlexVacio = False
            End If
            i = i + 3
        Next i
    Next J
End Function

Public Sub GrabaLimiteAutorizacion(ByVal lsAnio As String, ByVal lsMes As String, ByVal lsMovNro As String)
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim pbTran As Boolean
    
    lsSQL = "exec stp_Sel_ERS0652016_ConfLimitesAutorizaciones '" & lsAnio & "','" & lsMes & "','" & lsMovNro & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.Ejecutar (lsSQL)
    loReg.CierraConexion
End Sub
