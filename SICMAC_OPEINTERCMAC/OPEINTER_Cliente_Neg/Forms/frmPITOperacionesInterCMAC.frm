VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPITOperacionesInterCMAC 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7500
   ClientLeft      =   2430
   ClientTop       =   750
   ClientWidth     =   7875
   Icon            =   "frmPITOperacionesInterCMAC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOperacion 
      Caption         =   "Digite Operación"
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
      Height          =   765
      Left            =   6060
      TabIndex        =   6
      Top             =   60
      Width           =   1635
      Begin VB.TextBox txtOperacion 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame fraCMAC 
      Caption         =   "CMAC"
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
      Height          =   765
      Left            =   180
      TabIndex        =   4
      Top             =   60
      Width           =   5355
      Begin VB.ComboBox cmbCMACS 
         Height          =   315
         ItemData        =   "frmPITOperacionesInterCMAC.frx":030A
         Left            =   120
         List            =   "frmPITOperacionesInterCMAC.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   270
         Width           =   5055
      End
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Seleccione Operación"
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
      Height          =   5895
      Left            =   180
      TabIndex        =   3
      Top             =   960
      Width           =   7515
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   960
         Top             =   5520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPITOperacionesInterCMAC.frx":030E
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPITOperacionesInterCMAC.frx":0660
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPITOperacionesInterCMAC.frx":09B2
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPITOperacionesInterCMAC.frx":0D04
               Key             =   "Bebe"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwOperacion 
         Height          =   5475
         Left            =   150
         TabIndex        =   0
         Top             =   250
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9657
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   5520
      TabIndex        =   1
      Top             =   6960
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   6660
      TabIndex        =   2
      Top             =   6960
      Width           =   1020
   End
End
Attribute VB_Name = "frmPITOperacionesInterCMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsOpeLLama As ADODB.Recordset
Dim rsOpeRecep As ADODB.Recordset
Dim rsCMACs As ADODB.Recordset
Dim rsAgencia As ADODB.Recordset
Dim rsParametros As ADODB.Recordset


Private Function GetValorComision(ByVal nOperacion As Long) As Double
Dim rsPar As ADODB.Recordset
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion

Select Case nOperacion
    Case "261001" 'Retiro
        Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2115)
    Case "261002" 'Deposito
        Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2124)
    Case "261003" 'Consulta Cuentas de Ahorro
        Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2116)
    Case "261004" 'Consulta de Movimientos Ahorros
        Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2127)
    Case "104001" 'Pago de Credito
        Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2125)
    Case "104002" 'Consulta Cuentas de Credito
        Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2126)
End Select

'If (nOperacion = "261001" Or nOperacion = "261002" Or nOperacion = "104001") Then
'    Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2115)
'Else
'    Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2116)
'End If
    
Set oCap = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetValorComision = 0
Else
    GetValorComision = rsPar("nParValor")
End If
rsPar.Close
Set rsPar = Nothing
End Function

Private Sub LlenaArbol(ByVal rsOpe As ADODB.Recordset) '(ByVal sfiltro As String)
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node
    Do While Not rsOpe.EOF
        sOpeCod = rsOpe("cOpeCod")
        sOperacion = sOpeCod & " - " & UCase(rsOpe("cOpeDesc"))
        Select Case rsOpe("nOpeNiv")
            Case "1"
                sOpePadre = "P" & sOpeCod
                Set nodOpe = tvwOperacion.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
                nodOpe.Tag = sOpeCod
            Case "2"
                sOpeHijo = "H" & sOpeCod
                Set nodOpe = tvwOperacion.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
                nodOpe.Tag = sOpeCod
            Case "3"
                sOpeHijito = "J" & sOpeCod
                Set nodOpe = tvwOperacion.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
                nodOpe.Tag = sOpeCod
            Case "4"
                Set nodOpe = tvwOperacion.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
                nodOpe.Tag = sOpeCod
        End Select
        rsOpe.MoveNext
    Loop
End Sub

Private Sub EjecutaOperacion(ByVal nOperacion As CaptacOperacion, ByVal sDescOperacion As String)
    'Operaciones InterCMACs
    Select Case nOperacion
        Case "261001", "261002" 'Retiro, depósitos
            frmPITCapOpeInterCajaEnvio.inicia nOperacion, sDescOperacion, Right(Trim(cmbCMACS.Text), 13), Left(Trim(cmbCMACS.Text), Len(cmbCMACS.Text) - 73), GetValorComision(nOperacion)
        Case "261003", "261004" 'Cosulta de cuenta de ahorro, consulta de movimientos
            frmPITCapOpeInterCajaConsultasEnvio.inicia nOperacion, sDescOperacion, Right(Trim(cmbCMACS.Text), 13), Left(Trim(cmbCMACS.Text), Len(cmbCMACS.Text) - 73), GetValorComision(nOperacion)
        Case "104001" 'Pago de créditos
            frmPITColOpeInterCajaEnvio.Inicio nOperacion, sDescOperacion, Right(Trim(cmbCMACS.Text), 13), Left(Trim(cmbCMACS.Text), Len(cmbCMACS.Text) - 73), GetValorComision(nOperacion)
        Case "104002" 'Consulta de cuentas de crédito
            frmPITColOpeInterCajaConsultasEnvio.inicia nOperacion, sDescOperacion, Right(Trim(cmbCMACS.Text), 13), Left(Trim(cmbCMACS.Text), Len(cmbCMACS.Text) - 73), GetValorComision(nOperacion)
    End Select
End Sub



Private Sub cmdAceptar_Click()
Dim nodOpe As Node
Dim sDesc As String

    If Trim(cmbCMACS.Text) <> "" Then
        Set nodOpe = tvwOperacion.SelectedItem
        If Not nodOpe Is Nothing Then
            sDesc = Mid(nodOpe.Text, 10, Len(nodOpe.Text) - 7)
            EjecutaOperacion CLng(nodOpe.Tag), sDesc
        End If
        Set nodOpe = Nothing
    Else
        MsgBox "Seleccione CMAC antes de realizar la operacion.", vbInformation, "Aviso"
        cmbCMACS.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Public Sub inicia(ByVal sCaption As String, Optional ByVal prsUsu As ADODB.Recordset, Optional prsCMACs As ADODB.Recordset)
    Me.Caption = sCaption
    'gsCodAge = "01"
    'gsCodUser = "GITU"
    Me.Show 1
End Sub

Private Sub Form_Load()
    
    Call Inicializa
    Call LlenaArbol(rsOpeLLama)
    Call PIT_Llenar_Combo_con_Recordset(rsCMACs, cmbCMACS)
    
    fraCMAC.Enabled = True
        
End Sub

Private Sub tvwOperacion_DblClick()
    If tvwOperacion.SelectedItem.Child Is Nothing Then
        If Trim(cmbCMACS.Text) <> "" Then
            Dim nodOpe As Node
            Dim sDesc As String
            Set nodOpe = tvwOperacion.SelectedItem
            sDesc = Mid(nodOpe.Text, 10, Len(nodOpe.Text) - 7)
            EjecutaOperacion CLng(nodOpe.Tag), sDesc
            Set nodOpe = Nothing
        Else
            MsgBox "Seleccione CMAC antes de realizar la operacion.", vbInformation, "Aviso"
            'txtCMAC.SetFocus
            cmbCMACS.SetFocus
        End If
    End If
End Sub

Private Sub tvwOperacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If tvwOperacion.SelectedItem.Child Is Nothing Then
            If Trim(cmbCMACS.Text) <> "" Then
                Dim nodOpe As Node
                Dim sDesc As String
                Set nodOpe = tvwOperacion.SelectedItem
                sDesc = Mid(nodOpe.Text, 10, Len(nodOpe.Text) - 7)
                EjecutaOperacion CLng(nodOpe.Tag), sDesc
                Set nodOpe = Nothing
            Else
                MsgBox "Seleccione CMAC antes de realizar la operacion.", vbInformation, "Aviso"
                cmbCMACS.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmbCMACS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtOperacion.SetFocus
End If
End Sub


Private Sub txtOperacion_KeyUp(KeyCode As Integer, Shift As Integer)
Dim nodOpe As Node
For Each nodOpe In tvwOperacion.Nodes
    If nodOpe.Tag = Trim(txtOperacion) Then
        tvwOperacion.SelectedItem = nodOpe
        Exit For
    End If
Next
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub Inicializa()
Dim loPITNeg As COMOpeInterCMAC.dFuncionesNeg
    Set loPITNeg = New COMOpeInterCMAC.dFuncionesNeg
    
    Set rsOpeLLama = loPITNeg.obtenerTipoOperacionesInterCMAC
    
    Set rsCMACs = loPITNeg.obtenerCMACS
    
    Set rsAgencia = loPITNeg.obtenerDatosAgencia(gsCodAge)
    
    If Not (rsAgencia.EOF And rsAgencia.BOF) Then
        gsAgeCiudad = Trim(Replace(rsAgencia!cUbiGeoDescripcion, "(CIUDAD)", ""))
    End If
    
    Set rsParametros = loPITNeg.obtenerParametros()
    
    If (gsCodAge <> "02" And gsCodAge <> "07" And gsCodAge <> "25") Then
        While Not (rsParametros.EOF Or rsParametros.BOF)
            Select Case rsParametros!nParametroId
                Case 1000
                    gnMontoMinRetMN = rsParametros!nValor
                Case 1001
                    gnMontoMaxRetMN = rsParametros!nValor
                Case 1002
                    gnMontoMinRetME = rsParametros!nValor
                Case 1003
                    gnMontoMaxRetME = rsParametros!nValor
                Case 1004
                    gnMontoMinRetMNReqDNI = rsParametros!nValor
                Case 1005
                    gnMontoMinRetMEReqDNI = rsParametros!nValor
                Case 1006
                    gnMontoMaxOpeMNxDia = rsParametros!nValor
                Case 1007
                    gnMontoMaxOpeMExDia = rsParametros!nValor
                Case 1008
                    gnMontoMaxOpeMNxMes = rsParametros!nValor
                Case 1009
                    gnMontoMaxOpeMExMes = rsParametros!nValor
                Case 1010
                    gnNumeroMaxOpeXDia = rsParametros!nValor
                Case 1011
                    gnNumeroMaxOpeXMes = rsParametros!nValor
                Case 1012
                    gnMontoMinDepMN = rsParametros!nValor
                Case 1013
                    gnMontoMaxDepMN = rsParametros!nValor
                Case 1014
                    gnMontoMinDepME = rsParametros!nValor
                Case 1015
                    gnMontoMaxDepME = rsParametros!nValor
            End Select
            rsParametros.MoveNext
        Wend
    ElseIf gsCodAge = "02" Then     'Parametros para la Agencia Huanuco
        While Not (rsParametros.EOF Or rsParametros.BOF)
            Select Case rsParametros!nParametroId
                Case 1000
                    gnMontoMinRetMN = rsParametros!nValor
                Case 1020
                    gnMontoMaxRetMN = rsParametros!nValor
                Case 1002
                    gnMontoMinRetME = rsParametros!nValor
                Case 1021
                    gnMontoMaxRetME = rsParametros!nValor
                Case 1004
                    gnMontoMinRetMNReqDNI = rsParametros!nValor
                Case 1005
                    gnMontoMinRetMEReqDNI = rsParametros!nValor
                Case 1028
                    gnMontoMaxOpeMNxDia = rsParametros!nValor
                Case 1029
                    gnMontoMaxOpeMExDia = rsParametros!nValor
                Case 1008
                    gnMontoMaxOpeMNxMes = rsParametros!nValor
                Case 1009
                    gnMontoMaxOpeMExMes = rsParametros!nValor
                Case 1010
                    gnNumeroMaxOpeXDia = rsParametros!nValor
                Case 1011
                    gnNumeroMaxOpeXMes = rsParametros!nValor
                Case 1012
                    gnMontoMinDepMN = rsParametros!nValor
                Case 1024
                    gnMontoMaxDepMN = rsParametros!nValor
                Case 1014
                    gnMontoMinDepME = rsParametros!nValor
                Case 1025
                    gnMontoMaxDepME = rsParametros!nValor
            End Select
            rsParametros.MoveNext
        Wend
    Else    'Parametros para la agencia Tingo Maria y Cerro de Pasco
        While Not (rsParametros.EOF Or rsParametros.BOF)
            Select Case rsParametros!nParametroId
                Case 1000
                    gnMontoMinRetMN = rsParametros!nValor
                Case 1018
                    gnMontoMaxRetMN = rsParametros!nValor
                Case 1002
                    gnMontoMinRetME = rsParametros!nValor
                Case 1019
                    gnMontoMaxRetME = rsParametros!nValor
                Case 1004
                    gnMontoMinRetMNReqDNI = rsParametros!nValor
                Case 1005
                    gnMontoMinRetMEReqDNI = rsParametros!nValor
                Case 1026
                    gnMontoMaxOpeMNxDia = rsParametros!nValor
                Case 1027
                    gnMontoMaxOpeMExDia = rsParametros!nValor
                Case 1008
                    gnMontoMaxOpeMNxMes = rsParametros!nValor
                Case 1009
                    gnMontoMaxOpeMExMes = rsParametros!nValor
                Case 1010
                    gnNumeroMaxOpeXDia = rsParametros!nValor
                Case 1011
                    gnNumeroMaxOpeXMes = rsParametros!nValor
                Case 1012
                    gnMontoMinDepMN = rsParametros!nValor
                Case 1022
                    gnMontoMaxDepMN = rsParametros!nValor
                Case 1014
                    gnMontoMinDepME = rsParametros!nValor
                Case 1023
                    gnMontoMaxDepME = rsParametros!nValor
            End Select
            rsParametros.MoveNext
        Wend
    End If
    Set loPITNeg = Nothing
    
End Sub



