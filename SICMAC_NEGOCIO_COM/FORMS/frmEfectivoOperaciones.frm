VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEfectivoOperaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Contabilidad: Selección de Operaciones"
   ClientHeight    =   5970
   ClientLeft      =   1920
   ClientTop       =   1740
   ClientWidth     =   7125
   HelpContextID   =   210
   Icon            =   "frmEfectivoOperaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tvOpe 
      Height          =   5625
      Left            =   1575
      TabIndex        =   0
      Top             =   120
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   9922
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imglstFiguras"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.Animation Logo 
      Height          =   645
      Left            =   413
      TabIndex        =   6
      Top             =   210
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1138
      _Version        =   393216
      FullWidth       =   45
      FullHeight      =   43
   End
   Begin VB.Frame frmMoneda 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   113
      TabIndex        =   3
      Top             =   1020
      Width           =   1275
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &E."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   540
         Width           =   795
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &N."
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   150
      TabIndex        =   2
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   150
      TabIndex        =   1
      Top             =   5370
      Width           =   1200
   End
   Begin RichTextLib.RichTextBox rtxt 
      Height          =   315
      Left            =   300
      TabIndex        =   7
      Top             =   4350
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmEfectivoOperaciones.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   420
      Top             =   2280
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
            Picture         =   "frmEfectivoOperaciones.frx":038A
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEfectivoOperaciones.frx":06DC
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEfectivoOperaciones.frx":0A2E
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEfectivoOperaciones.frx":0D80
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEfectivoOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lExpand  As Boolean
Dim lExpandO As Boolean
Dim sArea  As String
Dim dFecha As Date, dFecha2 As Date

Public Sub Inicio(sObj As String, Optional plExpandO As Boolean = False)
    sArea = sObj
    lExpandO = plExpandO
    Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo ErrorAceptar
    If tvOpe.Nodes.Count = 0 Then
        MsgBox "Lista de operaciones se encuentra vacia", vbInformation, "Aviso"
        Exit Sub
    End If
    If tvOpe.SelectedItem.Tag = "1" Then
        MsgBox "Operación seleccionada no valida...!", vbInformation, "Aviso"
        tvOpe.SetFocus
        Exit Sub
    End If
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    If Left(tvOpe.SelectedItem.Key, 1) <> "P" Then
        gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescHijo = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescPadre = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60))
    Else
      gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    End If
    
    Select Case Mid(gsOpeCod, 1, 6)
        'Boveda Agencia
        Case gOpeBoveAgeConfHabCG
            frmCajaGenLista.Show 1
        Case gOpeBoveAgeHabAgeACG
            frmCajaGenHabilitacion.Show 1
        Case gOpeBoveAgeHabEntreAge
            frmCajaGenHabilitacion.Show 1
        Case gOpeBoveAgeHabCajero
            frmCajeroHab.Show 1
        Case gOpeBoveAgeExtConfHabCG, gOpeBoveAgeExtHabAgeACG, gOpeBoveAgeExtHabEntreAge
            frmCajaGenLista.Show 1
        Case gOpeBoveAgeExtHabCajero
            frmCajeroExtornos.Show 1
    
        'Cajero Moneda Extranjera
        Case gOpeCajeroMETipoCambio
            frmMantTipoCambio.Show 1
        Case gOpeCajeroMECompra
            frmCajeroCompraVenta.Show 1
        Case gOpeCajeroMEVenta
            frmCajeroCompraVenta.Show 1
        Case gOpeCajeroMEExtCompra, gOpeCajeroMEExtVenta
            frmCajeroExtornos.Show 1
        
        
        'Operaciones Cajero
        Case gOpeHabCajRegEfect
            frmCajaGenEfectivo.RegistroEfectivo True, gOpeHabCajRegEfect
        Case gOpeHabCajDevABove
            frmCajeroHab.Show 1
        Case gOpeHabCajTransfEfectCajeros
            frmCajeroHab.Show 1
        Case gOpeHabCajConfHabBovAge
            frmCajeroExtornos.Show 1
        Case gOpeHabCajRegSobFalt
            frmCajeroIngEgre.Show 1
        
        Case gOpeHabCajIngEfectRegulaFalt
            frmCajeroRegFaltSob.Show 1
        Case gOpeHabCajExtTransfEfectCajeros
            frmCajeroExtornos.Show 1
        Case gOpeHabCajExtConfHabBovAge
            frmCajeroExtornos.Show 1
        Case gOpeHabCajExtIngEfectRegulaFalt
            frmCajeroExtornos.Show 1
            
        Case gOpeHabCajExtDevABove
            frmCajeroExtornos.Show 1
        
    End Select
    Exit Sub
ErrorAceptar:
    MsgBox Err.Description, vbInformation, "Aviso Error"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
tvOpe.SetFocus
End Sub

Private Sub Form_Load()
    Dim RegOpe As ADODB.Recordset
    Dim nodX As Node
    Dim sCod As String
    On Error GoTo Error
    Logo.AutoPlay = True
    Logo.Open App.path & "\videos\LogoA.avi"
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    If Not lExpandO Then
       Dim oConst As New NConstSistemas
       sCod = oConst.LeeConstSistema(gConstSistContraerListaOpe)
       If sCod <> "" Then
         lExpand = IIf(UCase(Trim(sCod)) = "FALSE", False, True)
       End If
       Set oConst = Nothing
    Else
       lExpand = lExpandO
    End If
    LoadOpeUsu "2"
    Exit Sub
Error:
    MsgBox Err.Description, vbExclamation, Me.Caption
End Sub

Private Sub LoadOpeUsu(ByVal nmoneda As Moneda)
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node

Set clsGen = New DGeneral
'ARCV 20-07-2006
'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, sArea, MatOperac, NroRegOpe, nMoneda)
Set rsUsu = clsGen.GetOperacionesUsuario_NEW(sArea, nmoneda, gRsOpeRepo)

Set clsGen = Nothing
tvOpe.Nodes.Clear
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
            nodOpe.Tag = sOpeCod
    End Select
    nodOpe.Expanded = lExpand
    rsUsu.MoveNext
Loop
rsUsu.Close
Set rsUsu = Nothing
End Sub


Private Sub optMoneda_Click(Index As Integer)
Dim nDig As Moneda
If optMoneda(0) Then
    nDig = gMonedaExtranjera
Else
    nDig = gMonedaNacional
End If
LoadOpeUsu nDig
tvOpe.SetFocus
End Sub

Private Sub tvOpe_Collapse(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H80000008"
End Sub

Private Sub tvOpe_DblClick()
    If tvOpe.Nodes.Count > 0 Then
       cmdAceptar_Click
    End If
End Sub

Private Sub tvOpe_Expand(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H8000000D"
End Sub

Private Sub tvOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdAceptar_Click
    End If
End Sub
