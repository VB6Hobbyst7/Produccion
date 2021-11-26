VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DB786848-D4E8-474E-A2C2-DCBC1D43DA22}#2.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmCCEATMCargaCuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Asociadas a la Tarjeta"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   Icon            =   "frmCCEATMCargaCuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2910
      Left            =   83
      TabIndex        =   13
      Top             =   900
      Width           =   12600
      Begin MSComctlLib.ListView LstCta 
         Height          =   2670
         Left            =   75
         TabIndex        =   14
         Top             =   135
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   4710
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "DNI"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Tipo Programa"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Fecha Apertura"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tipo Cuenta"
            Object.Width           =   3246
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   720
      Left            =   83
      TabIndex        =   10
      Top             =   4560
      Width           =   12600
      Begin VB.CommandButton cmdVBSinTarj 
         Caption         =   "VB Sin Tarjeta"
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   225
         Width           =   1530
      End
      Begin VB.CommandButton CmdSelec 
         Caption         =   "Seleccionar"
         Enabled         =   0   'False
         Height          =   390
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   1395
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   390
         Left            =   11205
         TabIndex        =   11
         Top             =   225
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   83
      TabIndex        =   5
      Top             =   -15
      Width           =   12585
      Begin VB.TextBox TxtNumTarj 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2385
         MaxLength       =   16
         TabIndex        =   7
         Top             =   225
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   390
         Left            =   11205
         TabIndex        =   6
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
         Height          =   195
         Left            =   1650
         TabIndex        =   9
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblNumTarjeta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2625
         TabIndex        =   8
         Top             =   240
         Width           =   3225
      End
   End
   Begin VB.Frame fraClave 
      Height          =   735
      Left            =   83
      TabIndex        =   1
      Top             =   3810
      Width           =   12600
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver Detalle"
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdPedClaveAnt 
         Caption         =   "Pedir Clave"
         Height          =   360
         Left            =   8745
         TabIndex        =   2
         Top             =   255
         Width           =   1305
      End
      Begin VB.Label lblClave 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NO INGRESADO"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6465
         TabIndex        =   4
         Top             =   300
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave  :"
         Height          =   195
         Left            =   5625
         TabIndex        =   3
         Top             =   285
         Width           =   540
      End
   End
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   375
      Left            =   1815
      TabIndex        =   0
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmCCEATMCargaCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmCCEATMCargaCuentas
'** Descripción : Para la Carga de Cuentas Proyecto: Implementacion del Servicio de Compensaciòn Electrónica Diferido de Instrumentos Compensables CCE
'** Creación : PASI, 20160613
'**********************************************************************
Option Explicit

Private sCodCta As String
Private sNumTarj As String
Private nProd As Integer
Private c As ADODB.Connection

Dim sPin As String
Dim sPVV As String
Dim sOpeCod As String
Dim nValidaPIN As Integer
Dim lsTarjeta As String
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
Dim sPVVOrig As String


Public Function RecuperaCuenta(Optional ByVal psOpecod As String = "", _
                               Optional ByRef psNumTarj As String = "", _
                               Optional ByVal pnProd As Integer) As String
    nValidaPIN = 0
    sCodCta = ""
    sOpeCod = psOpecod
    sNumTarj = ""
    nProd = pnProd
    Me.Show 1
    
    psNumTarj = sNumTarj
    RecuperaCuenta = sCodCta
End Function
Private Sub cmdPedClaveAnt_Click()
    nValidaPIN = 0
    nValidaPIN = Tarjeta.PedirPinYValida(lsTarjeta, gNMKPOS, gWKPOS, gIpPuertoPinVerifyPOS, sPVV, gCanalIdPOS, gnTipoPinPad, gnPinPadPuerto)

    If nValidaPIN <> 0 Then
        Me.lblClave.Caption = "CLAVE INGRESADA"
    Else
        Me.lblClave.Caption = "NO INGRESADO"
    End If
    
    If sPVV = sPVVOrig Then
        nValidaPIN = 2
    End If
End Sub
Private Sub CmdSelec_Click()
Dim sResp As String
Dim i As Integer
Dim oNSegTar As COMNCaptaGenerales.NCOMSeguros
    If sOpeCod = "" Then
        sCodCta = Me.LstCta.SelectedItem.Text
        Unload Me
    Else
        Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
        If Left(sOpeCod, 4) = "9300" Then
            i = nValidaPIN
            If i = 1 Then
                sResp = "00"
                sCodCta = Me.LstCta.SelectedItem.Text
                Unload Me
            ElseIf nValidaPIN = 2 Then
                sResp = "99"
                Call MsgBox("Error en la clave, esta clave no es segura por favor cambiar de clave", vbInformation, "Aviso")
                Me.lblClave.Caption = "NO INGRESADO"
            Else
                sResp = "99"
                Call MsgBox("Error en la clave, reintente por favor, y si persiste el error comunicar al Area de Sistemas", vbInformation, "Aviso")
                Me.lblClave.Caption = "NO INGRESADO"
            End If
        Else
            sCodCta = Me.LstCta.SelectedItem.Text
            Unload Me
        End If
        Set oNSegTar = Nothing
    End If
    
End Sub
Private Sub cmdVBSinTarj_Click()
    If MsgBox("Se cobrará una comision desea continuar con la operacion", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then
        Set loVistoElectronico = New frmVistoElectronico
        
        lbVistoVal = loVistoElectronico.Inicio(5, sOpeCod)
            
        If Not lbVistoVal Then
            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones, se cobrara comision por esta operacion", vbInformation, "Mensaje del Sistema"
            Exit Sub
        End If
        
        sCodCta = "123456789"
        loVistoElectronico.RegistraVistoElectronico (0)
        
        Unload Me
    End If
End Sub
Private Sub cmdVer_Click()
    If Me.LstCta.ListItems.Count > 0 Then
        sCodCta = Me.LstCta.SelectedItem.Text
        Call frmCapMantenimiento.MuestraPosicionCliente(sCodCta)
    Else
        MsgBox "No se seleccionó ninguna cuenta", vbInformation, "Alerta"
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
    End If
End Sub

Private Sub SelecCtas_RecuperaDatosDETarjetas(ByVal psNumtarjeta As String, ByRef pnCond As Integer, _
    ByRef pnRetenTar As Integer, ByRef pnNOOperMonExt As Integer, ByRef nSuspOper As Integer, _
    ByRef dFecVenc As Date, ByRef psDescEstado As String)
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@PAN", adVarChar, adParamInput, 20)
    Prm.value = psNumtarjeta
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCondicion", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nRetenerTarjeta", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nNOOperMonExt", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSuspOper", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
     Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dfecVenc", adDBDate, adParamOutput)
    Cmd.Parameters.Append Prm
    
     Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psDescEstado", adVarChar, adParamOutput, 100)
    Cmd.Parameters.Append Prm
            
        
    Cmd.ActiveConnection = AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaDatosTarjeta"
    Cmd.Execute
    
    pnCond = Cmd.Parameters(1).value
    pnRetenTar = Cmd.Parameters(2).value
    pnNOOperMonExt = Cmd.Parameters(3).value
    nSuspOper = Cmd.Parameters(4).value
    dFecVenc = Cmd.Parameters(5).value
    psDescEstado = Cmd.Parameters(6).value
    

        Call CerrarConexion

    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim L As ListItem
Dim lrs As New ADODB.Recordset
Dim loConec As New COMConecta.DCOMConecta

    loConec.AbreConexion
    Set lrs = loConec.ConexionActiva.Execute(" exec stp_sel_CCE_ATM_RecuperaCtasSeleccion '" & sNumTarj & "'")
    If lrs.RecordCount = 0 Then
        Set lrs = loConec.ConexionActiva.Execute(" exec stp_sel_VISARecuperaCtasPersona '" & sNumTarj & "','" & CStr(nProd) & "'")
    End If
    
    LstCta.ListItems.Clear
    Do While Not lrs.EOF
            Set L = LstCta.ListItems.Add(, , lrs!cCtaCod)
            Call L.ListSubItems.Add(, , lrs!cPersNombre)
            Call L.ListSubItems.Add(, , lrs!cPersIDnro)
            Call L.ListSubItems.Add(, , IIf(Mid(lrs!cCtaCod, 9, 1) = "1", "SOLES", "DOLARES"))
            Call L.ListSubItems.Add(, , lrs!cTipoPrograma)
            Call L.ListSubItems.Add(, , Format(lrs!dApertura, "DD/MM/YYYY"))
            Call L.ListSubItems.Add(, , lrs!cTpoCta)
        lrs.MoveNext
    Loop
    
    If lrs.RecordCount > 0 Then
        CmdSelec.Enabled = True
        cmdVBSinTarj.Enabled = False
    End If
    
    lrs.Close
    loConec.CierraConexion
    Set lrs = Nothing
    Set loConec = Nothing
    
    Me.cmdVer.Visible = True
End Sub

Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    End If
End Sub
Public Function AbrirConexion() As ADODB.Connection
Dim sCadCadConex As String
End Function
Public Sub CerrarConexion()
    c.Close
    Set c = Nothing
End Sub
Public Function ExisteTarjeta_Selec(ByVal psNumtarjeta As String) As Boolean
Dim lrs As ADODB.Recordset
Dim loConec As New COMConecta.DCOMConecta

    Set lrs = New ADODB.Recordset
    loConec.AbreConexion
    Set lrs = loConec.ConexionActiva.Execute(" exec ATM_VerificaExisteTarjeta '" & psNumtarjeta & "'")
    
    If Not (lrs.EOF And lrs.BOF) Then
        If lrs("nExiste") > 0 Then
            ExisteTarjeta_Selec = True
        Else
            ExisteTarjeta_Selec = False
        End If
    End If
    Set lrs = Nothing
    loConec.CierraConexion
    Set loConec = Nothing
End Function
Private Sub CmdLecTarj_Click()

Me.Caption = "Activación de Tarjeta - PASE LA TARJETA"

sNumTarj = Mid(Tarjeta.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto, gnTimeOutAg), 2, 16)
lblNumTarjeta.Caption = Left(sNumTarj, 6) & "- - - - - -" & Right(sNumTarj, 4)
lsTarjeta = sNumTarj
Me.Caption = "Cuentas Asociadas a la Tarjeta - F12 para Digitar Tarjeta"

If sNumTarj = "" Then
    lblNumTarjeta.Caption = ""
    Me.LstCta.ListItems.Clear
    MsgBox "No hay conexion con el PINPAD o no paso la tarjeta, Intente otra vez", vbInformation, "MENSAJE DEL SISTEMA"
    Exit Sub
End If

If Not ExisteTarjeta_Selec(sNumTarj) Then
    lblNumTarjeta.Caption = ""
    Me.LstCta.ListItems.Clear
    MsgBox "La Tarjeta N° " & sNumTarj & " no Existe, Intente otra vez", vbInformation, "MENSAJE DEL SISTEMA"
    Exit Sub
End If

If Not ValidaEstadoTarjeta_Selec(sNumTarj) Then
    lblNumTarjeta.Caption = ""
    Me.LstCta.ListItems.Clear
    MsgBox "La Tarjeta no esta activa", vbInformation, "MENSAJE DEL SISTEMA"
    Exit Sub
End If

If Left(sNumTarj, 3) <> "ERR" Then
    sPVV = RecuperaPVV(sNumTarj)
    sPVVOrig = RecuperaPVVOrig(sNumTarj)
    Call CargaDatos
End If
End Sub
Private Sub cmdsalir_Click()
    sCodCta = ""
    sNumTarj = ""
    Unload Me
End Sub
Public Function RecuperaPVV(ByVal pPAN As String) As String
Dim lrs As ADODB.Recordset
Dim loCn As COMConecta.DCOMConecta
 
    Set lrs = New ADODB.Recordset
    Set loCn = New COMConecta.DCOMConecta
    loCn.AbreConexion
    Set lrs = loCn.ConexionActiva.Execute(" exec ATM_RecuperaPVV '" & pPAN & "'")
    If Not (lrs.EOF And lrs.BOF) Then
        RecuperaPVV = lrs("cPVV")
    Else
        RecuperaPVV = ""
    End If
    loCn.CierraConexion
    Set lrs = Nothing
    Set loCn = Nothing
 End Function
Public Function ValidaEstadoTarjeta_Selec(ByVal psNumtarjeta As String) As Boolean
Dim lrs As ADODB.Recordset
Dim loConec As New COMConecta.DCOMConecta

    Set lrs = New ADODB.Recordset
    loConec.AbreConexion
    Set lrs = loConec.ConexionActiva.Execute(" exec stp_sel_VISAValidaTarjeta '" & psNumtarjeta & "'")
    If Not (lrs.EOF And lrs.BOF) Then
        If lrs("nCondicion") <> 1 Then
            ValidaEstadoTarjeta_Selec = False
        Else
            ValidaEstadoTarjeta_Selec = True
        End If
    End If
    Set lrs = Nothing
    loConec.CierraConexion
    Set loConec = Nothing
End Function
Public Function RecuperaPVVOrig(ByVal pPAN As String) As String
Dim lrs As ADODB.Recordset
Dim loCn As COMConecta.DCOMConecta
 
    Set lrs = New ADODB.Recordset
    Set loCn = New COMConecta.DCOMConecta
    loCn.AbreConexion
    Set lrs = loCn.ConexionActiva.Execute(" exec ATM_RecuperaPVVOrig '" & pPAN & "'")
    If Not (lrs.EOF And lrs.BOF) Then
        RecuperaPVVOrig = lrs("cPVVOrig")
    Else
        RecuperaPVVOrig = ""
    End If
    loCn.CierraConexion
    Set lrs = Nothing
    Set loCn = Nothing
 End Function
