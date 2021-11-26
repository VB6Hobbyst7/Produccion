VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperacionesInterCMAC 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7500
   ClientLeft      =   2430
   ClientTop       =   750
   ClientWidth     =   7875
   Icon            =   "frmOperacionesInterCMAC.frx":0000
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
         ItemData        =   "frmOperacionesInterCMAC.frx":030A
         Left            =   120
         List            =   "frmOperacionesInterCMAC.frx":030C
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
               Picture         =   "frmOperacionesInterCMAC.frx":030E
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOperacionesInterCMAC.frx":0660
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOperacionesInterCMAC.frx":09B2
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOperacionesInterCMAC.frx":0D04
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
Attribute VB_Name = "frmOperacionesInterCMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsOpeLLama As ADODB.Recordset
Dim rsOpeRecep As ADODB.Recordset
Dim rsCMACs As ADODB.Recordset
Option Explicit

Private Function GetValorComision(ByVal nOperacion As Long) As Double
'Dim rsPar As ADODB.Recordset
'Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
'Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
'Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, gCostoOperacionCMACLlam)
'Set oCap = Nothing
'If rsPar.EOF And rsPar.BOF Then
    GetValorComision = 0
'Else
'    GetValorComision = rsPar("nParValor")
'End If
'rsPar.Close
'Set rsPar = Nothing
End Function

Private Sub LlenaArbol(ByVal rsUsu As ADODB.Recordset) '(ByVal sfiltro As String)
'Dim clsGen As COMDConstSistema.DCOMGeneral
'Dim rsUsu As ADODB.Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node

'Set clsGen = New COMDConstSistema.DCOMGeneral
'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, sfiltro, MatOperac, NroRegOpe)
'Set clsGen = Nothing

Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
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
    rsUsu.MoveNext
Loop
'rsUsu.Close
'Set rsUsu = Nothing
End Sub

Private Sub EjecutaOperacion(ByVal nOperacion As CaptacOperacion, ByVal sDescOperacion As String)
Select Case nOperacion
    'OPERACIONES DE LLAMADA
    'Abonos y Retiros
    Case gCMACOTAhoDepEfec, gCMACOTAhoDepChq, gCMACOTAhoRetEfec, gCMACOTAhoRetOP
        'frmCapOpeCMACLlam.Inicia nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC, GetValorComision(nOperacion)
        frmCapOpeCMACLlam.Inicia nOperacion, sDescOperacion, Right(Trim(cmbCMACS.Text), 13), Left(Trim(cmbCMACS.Text), Len(cmbCMACS.Text) - 73), GetValorComision(nOperacion)
    'Pago de Creditos
    Case "107001"
        frmColPOpeCMACLlam.Inicio nOperacion, sDescOperacion, Right(Trim(cmbCMACS.Text), 13), Left(Trim(cmbCMACS.Text), Len(cmbCMACS.Text) - 73), GetValorComision(nOperacion)
    'Cosulta de Saldo
    Case "260505", "260506"
        frmCapConsultaSaldos.Inicia nOperacion, sDescOperacion, Right(Trim(cmbCMACS.Text), 13), Left(Trim(cmbCMACS.Text), Len(cmbCMACS.Text) - 73)
    Case "107004"
        FrmColConsultaCtaCred.Inicia nOperacion, sDescOperacion, Right(Trim(cmbCMACS.Text), 13), Left(Trim(cmbCMACS.Text), Len(cmbCMACS.Text) - 73)
End Select
End Sub



Private Sub cmdAceptar_Click()
If Trim(cmbCMACS.Text) <> "" Then
    Dim nodOpe As Node
    Dim sDesc As String
    Set nodOpe = tvwOperacion.SelectedItem
    If Not nodOpe Is Nothing Then
        sDesc = Mid(nodOpe.Text, 10, Len(nodOpe.Text) - 7)
        EjecutaOperacion CLng(nodOpe.Tag), sDesc
    End If
    Set nodOpe = Nothing
Else
    MsgBox "Seleccione CMAC antes de realizar la operacion.", vbInformation, "Aviso"
    'txtCMAC.SetFocus
    cmbCMACS.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Public Sub Inicia(ByVal sCaption As String, ByVal prsUsu As ADODB.Recordset, prsCMACs As ADODB.Recordset)
    'Me.Caption = sCaption
    'LlenaArbol rsUsu 'sfiltro(I)
    'txtCMAC.rs = prsCMACs
    'fraCMAC.Enabled = True
    'txtCMAC.Enabled = True
    gsCodAge = "01"
    gsCodUser = "GITU"
    Me.Show 1
End Sub

Private Sub Form_Load()
'Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    Dim sImpresora As String
    Dim lnPos As Long
    
    Call FuncionesIni
    
    Me.Caption = "Cajero - Operaciones InterCMAC's"
    LlenaArbol rsOpeLLama
    Call Llenar_Combo_con_Recordset(rsCMACs, cmbCMACS)
    'txtCMAC.rs = prsCMACs
    fraCMAC.Enabled = True
    'txtCMAC.Enabled = True
    gsCodAge = "01"
    gsCodUser = "GITU"
    
    sImpresora = Printer.DeviceName
    If Left(sImpresora, 2) <> "\\" Then
        lnPos = InStr(1, Printer.Port, ":", vbTextCompare)
        If lnPos > 0 Then
            sLpt = Mid(Printer.Port, 1, lnPos - 1)
        Else
            sLpt = "LPT1"
        End If
    Else
        sLpt = frmImpresora.EliminaEspacios(sImpresora)
    End If
    
    '    DeshabilitaOpeacionesPendientes

    MsgBox "Por favor Configure su Impresora antes de Empezar sus operaciones", vbInformation, "Aviso"
    frmImpresora.Show 1
End Sub

Private Sub tvwOperacion_DblClick()
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
End Sub

Private Sub tvwOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(cmbCMACS.Text) <> "" Then
        Dim nodOpe As Node
        Dim sDesc As String
        Set nodOpe = tvwOperacion.SelectedItem
        sDesc = Mid(nodOpe.Text, 10, Len(nodOpe.Text) - 7)
        EjecutaOperacion CLng(nodOpe.Tag), sDesc
        Set nodOpe = Nothing
    Else
        MsgBox "Seleccione CMAC antes de realizar la operacion.", vbInformation, "Aviso"
    '    txtCMAC.SetFocus
        cmbCMACS.SetFocus
    End If
End If
End Sub

'Private Sub txtCMAC_EmiteDatos()
'If txtCMAC.Text <> "" Then
'    lblCMAC = Trim(txtCMAC.psDescripcion)
'    If txtOperacion.Enabled = False Then
'        txtOperacion.SetFocus
'    End If
'End If
'End Sub

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

Private Sub FuncionesIni()
    Call CargaVarSistema
    Call fgITFParametros
    Call obtenerRSOperaciones
End Sub
Private Sub obtenerRSOperaciones()
    Dim clsFun As DFunciones.dFuncionesNeg
    Set clsFun = New DFunciones.dFuncionesNeg
        
    clsFun.GetOperaciones rsOpeLLama, rsCMACs
End Sub
