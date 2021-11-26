VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajeroOpeCMAC 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6600
   ClientLeft      =   2430
   ClientTop       =   750
   ClientWidth     =   6660
   Icon            =   "frmCajeroOpeCMAC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
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
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   4395
      Begin SICMACT.TxtBuscar txtCMAC 
         Height          =   390
         Left            =   105
         TabIndex        =   8
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   688
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
         sTitulo         =   ""
      End
      Begin VB.Label lblCMAC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1995
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
      Height          =   5055
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   6435
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   1260
         Top             =   5700
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
               Picture         =   "frmCajeroOpeCMAC.frx":030A
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOpeCMAC.frx":065C
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOpeCMAC.frx":09AE
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOpeCMAC.frx":0D00
               Key             =   "Bebe"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwOperacion 
         Height          =   4635
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   8176
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
      End
   End
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
      Left            =   4500
      TabIndex        =   4
      Top             =   60
      Width           =   1995
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
         Left            =   300
         TabIndex        =   0
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   2340
      TabIndex        =   2
      Top             =   6120
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   3480
      TabIndex        =   3
      Top             =   6120
      Width           =   1020
   End
End
Attribute VB_Name = "frmCajeroOpeCMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetValorComision(ByVal nOperacion As Long) As Double
Dim rsPar As ADODB.Recordset
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, gCostoOperacionCMACLlam)
Set oCap = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetValorComision = 0
Else
    GetValorComision = rsPar("nParValor")
End If
rsPar.Close
Set rsPar = Nothing
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
    'OPERACIONES DE RECEPCION
    'Abonos
    Case gCMACOAAhoDepEfec, gCMACOAAhoDepChq
        frmCapAbonos.Inicia gCapAhorros, nOperacion, Right(Trim(txtCMAC), 13), lblCMAC
    Case gCMACOACTSDepEfec
        frmCapAbonos.Inicia gCapCTS, nOperacion, Right(Trim(txtCMAC), 13), lblCMAC
    'Cargos
    Case gCMACOAAhoRetEfec, gCMACOAAhoRetOP
        frmCapCargos.Inicia gCapAhorros, nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC
    Case gCMACOACTSRetEfec
        frmCapCargos.Inicia gCapCTS, nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC
    Case gCMACOAPFRetInt
        frmCapOpePlazoFijo.Inicia nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC
    'OPERACIONES DE LLAMADA
    'Abonos
    Case gCMACOTAhoDepEfec, gCMACOTAhoDepChq, gCMACOTAhoRetEfec, gCMACOTAhoRetOP
        frmCapOpeCMACLlam.Inicia nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC, GetValorComision(nOperacion)
    
    ' **** OPERACIONES DE RECEPCION - PIGNORATICIO
    Case geColPRenEnOtCj
        Call frmColPRenovacion.Inicio(nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC)
    Case geColPCanceEnOtCj
        Call frmColPCancelacion.Inicio(nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC)
        'Call frmColPOpeCMACLlam.Inicio(nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC)
    Case geColPAmortEnOtCj
        Call frmColPAmortizacion.Inicio(nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC)

'a
    Case "106001"
        frmCredPagoCuotas.RecepcionCmac (Right(Trim(txtCMAC), 13))
 'a
    
    ' **** OPERACIONES DE LLAMADA - PIGNORATICIO
    Case geColPRenDEOtCj
        Call frmColPOpeCMACLlam.Inicio(nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC, GetValorComision(nOperacion))
    Case geColPCanDEOtCj
        Call frmColPOpeCMACLlam.Inicio(nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC, GetValorComision(nOperacion))
    
'a
    Case "107001"
        Call frmColPOpeCMACLlam.Inicio(nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC, GetValorComision(nOperacion))
'a
    Case gColRecOpePagJudSDEnOtCjEfe
        Call frmColRecPagoCredRecup.Inicio(nOperacion, sDescOperacion, Right(Trim(txtCMAC), 13), lblCMAC, True)
    Case "106002"
        'FrmCreOpeRFA.RecepcionCmac (Right(Trim(txtCMAC), 13))
End Select
End Sub

Private Sub CmdAceptar_Click()
If Trim(lblCMAC) <> "" Then
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
    txtCMAC.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Public Sub Inicia(ByVal sCaption As String, _
                    ByVal prsUsu As ADODB.Recordset) 'ByVal sfiltro As Variant

Dim rsUsu As ADODB.Recordset
Dim clsCtaIF As COMDPersona.DCOMInstFinac

Me.Caption = sCaption

LlenaArbol prsUsu 'sfiltro(I)
Set clsCtaIF = New COMDPersona.DCOMInstFinac
Set rsUsu = clsCtaIF.CargaCmacs()
txtCMAC.psRaiz = "Cajas Municipales"
txtCMAC.rs = rsUsu
Set rsUsu = Nothing
fraCMAC.Enabled = True
txtCMAC.Enabled = True
Me.Show 1
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub tvwOperacion_DblClick()
If Trim(lblCMAC) <> "" Then
    Dim nodOpe As Node
    Dim sDesc As String
    Set nodOpe = tvwOperacion.SelectedItem
    sDesc = Mid(nodOpe.Text, 10, Len(nodOpe.Text) - 7)
    EjecutaOperacion CLng(nodOpe.Tag), sDesc
    Set nodOpe = Nothing
Else
    MsgBox "Seleccione CMAC antes de realizar la operacion.", vbInformation, "Aviso"
    txtCMAC.SetFocus
End If
End Sub

Private Sub tvwOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(lblCMAC) <> "" Then
        Dim nodOpe As Node
        Dim sDesc As String
        Set nodOpe = tvwOperacion.SelectedItem
        sDesc = Mid(nodOpe.Text, 10, Len(nodOpe.Text) - 7)
        EjecutaOperacion CLng(nodOpe.Tag), sDesc
        Set nodOpe = Nothing
    Else
        MsgBox "Seleccione CMAC antes de realizar la operacion.", vbInformation, "Aviso"
        txtCMAC.SetFocus
    End If
End If
End Sub

Private Sub txtCMAC_EmiteDatos()
If txtCMAC.Text <> "" Then
    lblCMAC = Trim(txtCMAC.psDescripcion)
    If txtOperacion.Enabled = False Then
        txtOperacion.SetFocus
    End If
End If
End Sub

Private Sub txtCMAC_KeyPress(KeyAscii As Integer)
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

