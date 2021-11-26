VERSION 5.00
Begin VB.Form frmEnvioEstadoCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de envío de Estado de Cuenta"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "frmEnvioEstadoCta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSiNo 
      Height          =   315
      ItemData        =   "frmEnvioEstadoCta.frx":030A
      Left            =   4080
      List            =   "frmEnvioEstadoCta.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feLista 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2143
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Codigo-Cliente-Domicilio-Correo-Envio"
      EncabezadosAnchos=   "0-0-1900-1900-1800-540"
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
      ColumnasAEditar =   "X-X-X-3-4-5"
      ListaControles  =   "0-0-0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CheckBox chkDebitoCta 
      Caption         =   "Débito misma cuenta"
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cboModoRecep 
      Height          =   315
      ItemData        =   "frmEnvioEstadoCta.frx":030E
      Left            =   2040
      List            =   "frmEnvioEstadoCta.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1155
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "¿Recibir información mensual de estado de cuenta?"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Medio de Recepción:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmEnvioEstadoCta.frx":0312
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmEnvioEstadoCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmEnvioEstadoCta
'** Descripción : Formulario para aprobar elegir el tipo de envio de estado de cuenta TI-ERS057-2013
'** Creación : JUEZ, 20130520 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim lnOperacion As Integer '1 Ahorros Registro, 2 Ahorros Mantenimiento, 3 Creditos
Dim rsCli As ADODB.Recordset
Dim lnProd As COMDConstantes.Producto
Private lnValida As Boolean
Private rsLista As ADODB.Recordset
Private lnModoEnvio As Integer
Private lnValidaSiNo As Boolean 'APRI20180215 ERS026-2017
Private fnTipo As Boolean 'APRI20180215 ERS026-2017
Private lnDebitoMismaCta As Integer
Dim fsCtaCod As String

Property Let RegistraEnvio(pRegEnvio As String)
   lnValida = pRegEnvio
End Property
Property Get RegistraEnvio() As String
    RegistraEnvio = lnValida
End Property
Property Let RecordSetDatos(rsDatos As ADODB.Recordset)
   Set rsLista = rsDatos
End Property
Property Get RecordSetDatos() As ADODB.Recordset
    Set RecordSetDatos = rsLista
End Property
Property Let ModoEnvioEstCta(pnModoEnvio As Integer)
   lnModoEnvio = pnModoEnvio
End Property
Property Get ModoEnvioEstCta() As Integer
    ModoEnvioEstCta = lnModoEnvio
End Property
Property Let RegistraEnvioSiNo(pRegEnvioSiNo As String)
   lnValidaSiNo = pRegEnvioSiNo
End Property
Property Get RegistraEnvioSiNo() As String
    RegistraEnvioSiNo = lnValidaSiNo
End Property
Property Let DebitoMismaCta(pnDebitoMismaCta As Integer)
   lnDebitoMismaCta = pnDebitoMismaCta
End Property
Property Get DebitoMismaCta() As Integer
    DebitoMismaCta = lnDebitoMismaCta
End Property
Public Sub InicioCap(ByVal pnProducto As COMDConstantes.Producto, ByVal RSClientes As ADODB.Recordset)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(9110)
    Set clsGen = Nothing

    cboModoRecep.Clear
    While Not rsConst.EOF
        cboModoRecep.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
    
    'APRI20180309  ERS036-2017
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(4046)
    cboSiNo.Clear
    While Not rsConst.EOF
        cboSiNo.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
    cboSiNo.ListIndex = 1
    cboModoRecep.ListIndex = 1
    cboModoRecep.Enabled = False
    feLista.Enabled = False
    'END APRI
    Set rsCli = RSClientes
    rsCli.MoveFirst
    lnProd = pnProducto
    lnOperacion = 1
    lnValida = False
    Set rsLista = Nothing
    lnModoEnvio = 0
    lnDebitoMismaCta = 0
    fsCtaCod = ""
'COMENTADO POR APRI20180406 ERS036-2017
'    If lnProd = gCapAhorros Then
'        chkDebitoCta.Visible = True
'        chkDebitoCta.value = 1
'    Else
'        chkDebitoCta.Visible = False
'    End If
    chkDebitoCta.Visible = False
    chkDebitoCta.Enabled = False
'END APRI
    Me.Show 1
End Sub

Public Sub MantenimientoCap(ByVal psCtaCod As String, ByVal RSClientes As ADODB.Recordset)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim rsConst As New ADODB.Recordset
    Dim rsDatos As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(9110)
    Set clsGen = Nothing
    
    cboModoRecep.Clear
    While Not rsConst.EOF
        cboModoRecep.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
    Set rsCli = RSClientes
    rsCli.MoveFirst
    lnProd = Mid(psCtaCod, 6, 3)
    lnOperacion = 2
    lnValida = False
    fsCtaCod = psCtaCod
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsDatos = oCap.RecuperaDatosEnvioEstadoCta(fsCtaCod)
    Set oCap = Nothing
    If Not rsDatos.EOF Then
        Set rsLista = rsDatos
        lnModoEnvio = IIf(IsNull(rsDatos!nModoEnvio), 0, rsDatos!nModoEnvio)
        cboModoRecep.ListIndex = IndiceListaCombo(cboModoRecep, Space(30) & CStr(lnModoEnvio))
        lnDebitoMismaCta = IIf(rsDatos!cCtaCod = rsDatos!Cuenta, 1, 0)
        If Mid(psCtaCod, 6, 3) = gCapAhorros Then
            chkDebitoCta.Visible = True
            chkDebitoCta.value = IIf(rsDatos!cCtaCod = rsDatos!Cuenta, 1, 0)
        Else
            chkDebitoCta.Visible = False
        End If
    Else
        Set rsLista = Nothing
        lnModoEnvio = 1
        cboModoRecep.ListIndex = IndiceListaCombo(cboModoRecep, Space(30) & CStr(lnModoEnvio))
        lnDebitoMismaCta = 0
        If Mid(psCtaCod, 6, 3) = gCapAhorros Then
            chkDebitoCta.Visible = True
            chkDebitoCta.value = 1
        Else
            chkDebitoCta.Visible = False
        End If
    End If
    Set rsDatos = Nothing
    Me.Show 1
End Sub

Public Sub InicioCol(ByVal psCtaCod As String, ByVal RSRelaciones As ADODB.Recordset, ByVal nTipo As Boolean)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim rsConst As New ADODB.Recordset
    Dim rsDatos As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(9110)
    Set clsGen = Nothing

    cboModoRecep.Clear
    While Not rsConst.EOF
        cboModoRecep.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
    
    'APRI20180309  ERS036-2017
    fnTipo = nTipo
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(4046)
    Set clsGen = Nothing
    cboSiNo.Clear
    While Not rsConst.EOF
        cboSiNo.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
    
    If Len(psCtaCod) < 18 Then
    cboSiNo.ListIndex = 1
    cboModoRecep.ListIndex = 1
    cboModoRecep.Enabled = False
    feLista.Enabled = False
    End If
    'END APRI
    
    fsCtaCod = psCtaCod
    Set rsCli = RSRelaciones
    rsCli.MoveFirst
    lnOperacion = 3
    lnValida = False
    'feLista.EncabezadosAnchos = "0-0-2800-2800-540-0"
    feLista.EncabezadosAnchos = "0-0-2100-2100-2100-0" 'APRI20180309  ERS036-2017
    If Len(fsCtaCod) = 18 Then
        Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsDatos = oCap.RecuperaDatosEnvioEstadoCta(fsCtaCod)
        Set oCap = Nothing
        If Not rsDatos.EOF Then
            Set rsLista = rsDatos
            cboSiNo.ListIndex = 0 'APRI20180309  ERS036-2017
            lnModoEnvio = IIf(IsNull(rsDatos!nModoEnvio), 0, rsDatos!nModoEnvio)
            cboModoRecep.ListIndex = IndiceListaCombo(cboModoRecep, Space(30) & CStr(lnModoEnvio))
        Else
            Set rsLista = Nothing
            lnModoEnvio = 0
            'APRI20180309  ERS036-2017
            cboSiNo.ListIndex = 1
           ' cboModoRecep.ListIndex = 1
            'cboModoRecep.Enabled = False
            feLista.Enabled = False
            'END APRI
        End If
    Else
        Set rsLista = Nothing
        lnModoEnvio = 0
    End If
    lnDebitoMismaCta = 0
    chkDebitoCta.Visible = False
    chkDebitoCta.Enabled = False
    Me.Show 1
End Sub

Private Sub cboModoRecep_Click()
    If Trim(Right(cboModoRecep.Text, 2)) = "1" Then
        
        feLista.EncabezadosAnchos = "0-0-3000-0-2800-540"
        'chkDebitoCta.value = 0
        'chkDebitoCta.Enabled = False
        Call LimpiaFlex(feLista)
        'feLista.lbEditarFlex = False
        
        
    Else
    feLista.EncabezadosAnchos = "0-0-3000-2800-0-540"
'        If lnProd = gCapAhorros Then
'            chkDebitoCta.value = 1
'            chkDebitoCta.Enabled = True
'        End If
        Call LimpiaFlex(feLista)
        'feLista.lbEditarFlex = True
    End If
    ListarClientes
End Sub

Private Sub ListarClientes()
    Select Case lnOperacion
        Case 1
            ListaClientesDefault
        Case 2, 3
            Dim rs As ADODB.Recordset
            Dim i As Integer
            Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
            If Len(fsCtaCod) = 18 Then
                Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
                Set rs = oCap.RecuperaDatosEnvioEstadoCta(fsCtaCod)
                Set oCap = Nothing
                Call LimpiaFlex(feLista)
                i = 1
                If rs.EOF Then ListaClientesDefault
                While Not rs.EOF
                    feLista.AdicionaFila
                    feLista.TextMatrix(i, 1) = rs!codigo
                    feLista.TextMatrix(i, 2) = rs!Cliente
                    feLista.TextMatrix(i, 3) = Trim(rs!Domicilio)
                    feLista.TextMatrix(i, 4) = Trim(rs!Email)
                    feLista.TextMatrix(i, 5) = 1
                    'feLista.TextMatrix(i, 4) = IIf(Trim(Right(cboModoRecep.Text, 2)) = "2", rs!Envio, 0)
                    'feLista.TextMatrix(i, 5) = IIf(rs!cCtaCod = rs!Cuenta, "", rs!Cuenta)
                    i = i + 1
                    rs.MoveNext
                Wend
            Else
                ListaClientesDefault
            End If
    End Select
End Sub

Private Sub ListaClientesDefault()
    Dim oPers As COMDPersona.DCOMPersona
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim nRelacion As Integer
    rsCli.MoveFirst
    Call LimpiaFlex(feLista)
    i = 1
    While Not rsCli.EOF

    nRelacion = Trim(Right(rsCli!Relacion, 2))
 
     If nRelacion <> 12 And nRelacion <> 13 And nRelacion <> 23 And nRelacion <> 25 Then 'APRI20180309 ERS036-2017
        Set oPers = New COMDPersona.DCOMPersona
        Set rs = oPers.RecuperaPersonaEnvioEstadoCta(rsCli!codigo)
        Set oPers = Nothing
        feLista.AdicionaFila
        feLista.TextMatrix(i, 1) = rs!cPersCod
        feLista.TextMatrix(i, 2) = rs!cPersNombre
        feLista.TextMatrix(i, 3) = Trim(rs!cPersDireccDomicilio)
        feLista.TextMatrix(i, 4) = Trim(rs!cPersEmail) '0 'APRI20180309 ERS036-2017
        feLista.TextMatrix(i, 5) = ""
        i = i + 1
      End If
        rsCli.MoveNext
    Wend
End Sub
Private Sub cboSiNo_Click()
 If Trim(Right(cboSiNo.Text, 2)) = "1" Then
        cboModoRecep.Enabled = True
        feLista.Enabled = True
        feLista.lbEditarFlex = True
        If feLista.TextMatrix(1, 1) = "" Then
            Call cboModoRecep_Click
        End If
    Else
        cboModoRecep.ListIndex = 1
        cboModoRecep.Enabled = False
        feLista.Enabled = False
        feLista.EncabezadosAnchos = "0-0-1900-1900-1900-540"
        Call LimpiaFlex(feLista)
    End If
End Sub

Private Sub chkDebitoCta_Click()
    If chkDebitoCta.value = 1 Then
        Dim i As Integer
        For i = 1 To feLista.rows - 1
            If feLista.TextMatrix(feLista.row, 4) = "." Then
                feLista.TextMatrix(feLista.row, 5) = ""
            End If
        Next i
    End If
End Sub

Private Sub CmdAceptar_Click()
Dim Msg As String
If fnTipo Then
    If Trim(Right(cboSiNo.Text, 2)) = 1 Then
         If Trim(Right(cboModoRecep.Text, 2)) <> "" Then
            Dim j As Integer
            Dim nCantEnv0 As Integer
            For j = 1 To feLista.rows - 1
                If feLista.TextMatrix(j, 5) = "." Then
                    nCantEnv0 = nCantEnv0 + 1
                    If Trim(Right(cboModoRecep.Text, 2)) = "2" And feLista.TextMatrix(j, 3) = "" Then
                        MsgBox "Debe ingresar la dirección del cliente", vbInformation, "Aviso"
                        Exit Sub
                    End If
                    If Trim(Right(cboModoRecep.Text, 2)) = "1" And feLista.TextMatrix(j, 4) = "" Then
                        MsgBox "Debe ingresar el correo electrónico del cliente", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            Next j
            If nCantEnv0 = 0 Then
                Msg = "Debe elegir como mínimo un cliente para recepcionar la información " & IIf(Trim(Right(cboModoRecep.Text, 2)) = "1", "electrónica", "física")
                MsgBox Msg, vbInformation, "Aviso"
                Exit Sub
            End If
     
    
            Set rsLista = feLista.GetRsNew
            lnModoEnvio = CInt(Trim(Right(cboModoRecep.Text, 2)))
            Call GuardarRegistroEnvioEstadoCta(1, fsCtaCod, rsLista, lnModoEnvio, 0, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
           
        Else
            MsgBox "Debe escoger el modo de recepción del estado de cuenta", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
            Dim oEnvEstCta As COMDCaptaGenerales.DCOMCaptaGenerales
            Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
            Call oEnvEstCta.EliminarRegistroEnvioEstadoCta(fsCtaCod)
            Set oEnvEstCta = Nothing
        
    End If
    lnValida = True
    Unload Me
Else
    lnValidaSiNo = IIf(Trim(Right(cboSiNo.Text, 2)) = "1", True, False) ' APRI20180310 ERS036-2017
    If Trim(Right(cboSiNo.Text, 2)) = 1 Then
        If Trim(Right(cboModoRecep.Text, 2)) <> "" Then
    'COMENTADO POR APRI20180310 ERS036-2017
    '        If Trim(Right(cboModoRecep.Text, 2)) = "2" Then
    '            Dim i As Integer
    '            Dim nCantEnv As Integer
    '            For i = 1 To feLista.rows - 1
    '                If feLista.TextMatrix(i, 4) = "." Then
    '                    nCantEnv = nCantEnv + 1
    '                    If feLista.TextMatrix(i, 3) = "" Then
    '                        MsgBox "Debe ingresar la dirección del cliente", vbInformation, "Aviso"
    '                        Exit Sub
    '                    End If
    '                    If lnOperacion <> 3 Then
    '                        If chkDebitoCta.value = 0 Then
    '                            If feLista.TextMatrix(i, 5) = "" Then
    '                                MsgBox "Debe elegir la cuenta del cliente " & feLista.TextMatrix(i, 2) & " desde la que se realizará el débito", vbInformation, "Aviso"
    '                                Exit Sub
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            Next i
    '            If nCantEnv = 0 Then
    '                MsgBox "Debe elegir como mínimo un cliente para recepcionar la información física", vbInformation, "Aviso"
    '                Exit Sub
    '            End If
    '        End If
    
           
            Dim i As Integer
            Dim nCantEnv As Integer
            For i = 1 To feLista.rows - 1
                If feLista.TextMatrix(i, 5) = "." Then
                    nCantEnv = nCantEnv + 1
                    If Trim(Right(cboModoRecep.Text, 2)) = "2" And feLista.TextMatrix(i, 3) = "" Then
                        MsgBox "Debe ingresar la dirección del cliente", vbInformation, "Aviso"
                        Exit Sub
                    End If
                    If Trim(Right(cboModoRecep.Text, 2)) = "1" And feLista.TextMatrix(i, 4) = "" Then
                        MsgBox "Debe ingresar el correo electrónico del cliente", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            Next i
            If nCantEnv = 0 Then
                Msg = "Debe elegir como mínimo un cliente para recepcionar la información " & IIf(Trim(Right(cboModoRecep.Text, 2)) = "1", "electrónica", "física")
                MsgBox Msg, vbInformation, "Aviso"
                Exit Sub
            End If
     
    
            Set rsLista = feLista.GetRsNew
            lnValida = True
            lnModoEnvio = CInt(Trim(Right(cboModoRecep.Text, 2)))
            lnDebitoMismaCta = IIf(lnOperacion <> 3, chkDebitoCta.value, 0)
            Unload Me
        Else
            MsgBox "Debe escoger el modo de recepción del estado de cuenta", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        lnValida = False
        Unload Me
    End If
End If
End Sub

Private Sub cmdCancelar_Click()
    lnValida = False
    lnValidaSiNo = False
    Unload Me
End Sub
'COMENTADO POR APRI20180406 ERS036-2017
'Private Sub feLista_Click()
'    If feLista.Col = 5 Then
'        RecuperaCtas
'    End If
'End Sub
'
'Private Sub feLista_EnterCell()
'    If feLista.Col = 5 Then
'        RecuperaCtas
'    End If
'End Sub
'END APRI

Private Sub feLista_OnCellChange(pnRow As Long, pnCol As Long)
    'If feLista.TextMatrix(feLista.row, 4) = "" Then
    If feLista.TextMatrix(feLista.row, 5) = "" Then
        Dim oPers As COMDPersona.DCOMPersona
        Dim rs As ADODB.Recordset
        Set oPers = New COMDPersona.DCOMPersona
        Set rs = oPers.RecuperaPersonaEnvioEstadoCta(feLista.TextMatrix(feLista.row, 1))
        Set oPers = Nothing
        feLista.TextMatrix(feLista.row, 3) = rs!cPersDireccDomicilio
        feLista.TextMatrix(feLista.row, 4) = rs!cPersEmail 'APRI20180406 ERS036-2017
        feLista.TextMatrix(feLista.row, 5) = ""
    Else
        If Me.chkDebitoCta.value = 1 Then
            feLista.TextMatrix(feLista.row, 5) = ""
        End If
    End If
End Sub

Private Sub feLista_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    'If feLista.TextMatrix(pnRow, 4) = "." Then
    If feLista.TextMatrix(pnRow, 5) = "." Then
        RecuperaCtas
    Else
        Dim oPers As COMDPersona.DCOMPersona
        Dim rs As ADODB.Recordset
        Set oPers = New COMDPersona.DCOMPersona
        Set rs = oPers.RecuperaPersonaEnvioEstadoCta(feLista.TextMatrix(feLista.row, 1))
        Set oPers = Nothing
        feLista.TextMatrix(feLista.row, 3) = rs!cPersDireccDomicilio
        feLista.TextMatrix(feLista.row, 5) = ""
    End If
End Sub

Private Sub feLista_OnChangeCombo()
    If chkDebitoCta.value = 1 Then
        feLista.TextMatrix(feLista.row, 5) = ""
    End If
End Sub

Private Sub RecuperaCtas()
    Dim rsCtasPermitidas As ADODB.Recordset
    Dim oCaptac As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCaptac = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsCtasPermitidas = oCaptac.RecuperaCtasParaDebitoEnvioEstadoCta(feLista.TextMatrix(feLista.row, 1))
    Set oCaptac = Nothing
    feLista.CargaCombo rsCtasPermitidas
    Set rsCtasPermitidas = Nothing
End Sub

Public Sub GuardarRegistroEnvioEstadoCta(ByVal pnTipoProd As Integer, ByVal psCtaCod As String, ByVal prsEnvEstCta As ADODB.Recordset, _
                                         ByVal pnModoEnvioEstCta As Integer, ByVal pnDebitoMismaCta As Integer, ByVal psMovNro As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Call clsMant.GuardarRegistroEnvioEstadoCta(pnTipoProd, psCtaCod, prsEnvEstCta, pnModoEnvioEstCta, pnDebitoMismaCta, psMovNro)
    Set clsMant = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'lnValida = False
End Sub
