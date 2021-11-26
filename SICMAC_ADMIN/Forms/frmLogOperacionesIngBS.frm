VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogOperacionesIngBS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Contabilidad: Selección de Operaciones"
   ClientHeight    =   5970
   ClientLeft      =   1920
   ClientTop       =   1740
   ClientWidth     =   7125
   HelpContextID   =   210
   Icon            =   "frmLogOperacionesIngBS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tvOpe 
      Height          =   5745
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   10134
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imglstFiguras"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.Animation Logo 
      Height          =   645
      Left            =   435
      TabIndex        =   6
      Top             =   195
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1138
      _Version        =   393216
      FullWidth       =   45
      FullHeight      =   43
   End
   Begin VB.Frame frmMoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   150
      TabIndex        =   3
      Top             =   1020
      Width           =   1275
      Begin VB.OptionButton optMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "M. &N."
         ForeColor       =   &H80000008&
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
      Left            =   120
      TabIndex        =   2
      Top             =   4950
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   5370
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtxt 
      Height          =   315
      Left            =   270
      TabIndex        =   7
      Top             =   3810
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmLogOperacionesIngBS.frx":08CA
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
            Picture         =   "frmLogOperacionesIngBS.frx":094B
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogOperacionesIngBS.frx":0C9D
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogOperacionesIngBS.frx":0FEF
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogOperacionesIngBS.frx":1341
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLogOperacionesIngBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lExpand  As Boolean
Dim lExpandO As Boolean
Dim sArea  As String
Dim dfecha As Date, dFecha2 As Date

Public Sub Inicio(sObj As String, Optional plExpandO As Boolean = False)
    sArea = sObj
    lExpandO = plExpandO
    Me.Show 0, MDISicmact
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
        gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescHijo = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescPadre = Trim(Mid(tvOpe.SelectedItem.parent.Text, 9, 60))
    Else
      gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    End If

    'gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))

    If Right(gsOpeCod, 2) = "00" Then Exit Sub

    Select Case Mid(gsOpeCod, 1, 6)
        Case gnSubIniProceso
            frmLogSubastaIni.Ini gsOpeCod, gsOpeDesc, False

        Case gnSubListados
            frmLogSubastaIni.Ini gsOpeCod, gsOpeDesc, True

        Case gnSubVentaBol
            'frmLogSubastaVenta.Ini gsOpeCod, gsOpeDesc, True

        Case gnSubVentaFac
            'frmLogSubastaVenta.Ini gsOpeCod, gsOpeDesc, False

        Case gnSubRegBilletaje
            frmCajaGenEfectivo.RegistroEfectivo True

        Case gnSubCuadreCaja
            frmLogSubastaCuadreCaja.Ini gsCodUser, gsOpeCod, gsOpeDesc

        Case gnSubCierreProceso
            frmLogSubastaIni.Ini gsOpeCod, gsOpeDesc, False, True

        '/*ALMACEN*/
        Case gnAlmaReqAreaReg
            frmLogSalAlmacen.Ini gsOpeCod, gsOpeDesc, True, , True

        Case gnAlmaReqAreaMant
            frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, , True

        Case gnAlmaReqAreaExt
            frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True, True

        Case gnAlmaReqAreaRechPar
            'frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True
            frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, , , , True

        Case gnAlmaSalXAtencion 'SALIDA POR ATENCION
            frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True, True

        Case gnAlmaIngXCompras  'INGRESO POR COMPRAS
            frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc

        Case gnAlmaIngXComprasConfirma
            frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True  'Confirma

        Case gnAlmarReporteMovReq
            frmLogDocMovimiento.Inicio "70", gsOpeDesc

        Case gnAlmarReporteMovNotIng
            frmLogDocMovimiento.Inicio "42", gsOpeDesc

        Case gnAlmarReporteMovGuiaSal
            frmLogDocMovimiento.Inicio "71", gsOpeDesc

        Case gnAlmaMantXIngreso
            frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True, True

        Case gnAlmaMantXSalida
            frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True, True, , , True

        Case gnAlmaExtornoXIngreso
            frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True, False, True

        Case gnAlmaExtornoXSalida
            frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True, True, True

        Case gnAlmarReporteListaIng
            frmLogReportes.Ini True, gsOpeDesc

        Case gnAlmarReporteListaSal
            frmLogReportes.Ini False, gsOpeDesc

        Case gnAlmarReporteResumenMovMes
            frmLogRepMensual.Show 1
        
        Case gnAlmarReporteIngAFBND
            frmLogAfTrans.Ini True, False, "Ingresos de Mes"
        
        Case gnAlmarReporteSalAFBND
            frmLogAfTrans.Ini False, True, "Salidas de Mes"
        
        Case gnAlmarReporteActivoFijo
            frmLogAfTrans.Ini False, False, "Transferencia de Activo Fijo"
            
        Case "571001", "572001", "571002", "572002", "571003", "572003", _
             "571003", "572003", "571004", "572004", "571005", "572005", _
             "571006", "572006", "571007", "572007", "571008", "572008", _
             "571009", "572009"
            frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc
        Case "571101", "572101"
            frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False
        Case Else
            If Left(gnAlmaSalXAtencion, 4) = Left(gsOpeCod, 4) Then   'OTRAS SALIDAS
                frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True, True
            ElseIf gnAlmaIngXDacionPago = gsOpeCod Or gnAlmaIngXAdjudicacion = gsOpeCod Then  'DACION EN AGO Y ADJUDICACION
                frmLogIngAlmacen2.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, True
            ElseIf Left(gnAlmaIngXComprasConfirma, 4) = Left(gsOpeCod, 4) Then
                'frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, True
                frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, True

            'Para el Control Vehicular
            ElseIf Left(gsOpeCod, 2) = "55" Then
                Select Case gsOpeCod
                    'Case "551011"
                        ' frmLogVehiculoMnt.Show 1
                    'Case "551012"
                        ' frmLogVehiculoCond.Show 1
                    Case "551013"
                    
                    Case "551021"
                    
                    Case "551022"
                    Case "551023"
                    Case "551024"
                    Case "551025"
                    Case "551026"
                    Case "551031"
                    Case "551032"
                    Case "551041"
                End Select

            'FIN de Control Vehicular
            ElseIf Left(gsOpeCod, 2) = "56" Then
                If Mid(gsOpeCod, 4, 1) = "4" Then
                    frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, True
                ElseIf Mid(gsOpeCod, 4, 1) = "6" Then
                    frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False
                Else
                    frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True
                End If
            ElseIf Left(gsOpeCod, 2) = "57" Then
                If Mid(gsOpeCod, 4, 1) = "4" Then
                    frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, True
                ElseIf Mid(gsOpeCod, 4, 1) = "6" Then
                    frmLogIngAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False
                ElseIf Mid(gsOpeCod, 4, 1) = "3" Then
                    frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, True, True
                ElseIf Mid(gsOpeCod, 4, 1) = "2" Then
                    frmLogSalAlmacen.Ini Mid(gsOpeCod, 1, 6), gsOpeDesc, False, True
                End If
    
            End If
    End Select
    Exit Sub
ErrorAceptar:
    MsgBox Err.Description, vbInformation, "Aviso Error"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
'tvOpe.SetFocus
End Sub

Private Sub Form_Load()
    Dim sCod As String
    On Error GoTo ERROR
    MDISicmact.Enabled = False

    Logo.AutoPlay = True
    Logo.Open App.path & "\videos\LogoA.avi"

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
ERROR:
    MsgBox Err.Description, vbExclamation, Me.Caption
End Sub

Sub LoadOpeUsu(psMoneda As String)
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node

Set clsGen = New DGeneral
Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, sArea, MatOperac, NroRegOpe, psMoneda)

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
RSClose rsUsu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDISicmact.Enabled = True
End Sub

Private Sub optMoneda_Click(Index As Integer)
    Dim sDig As String
    Dim sCod As String
    On Error GoTo ERROR
    If optMoneda(0) Then
        sDig = "2"
    Else
        sDig = "1"
    End If
    AbreConexion
    LoadOpeUsu sDig
    CierraConexion
    tvOpe.SetFocus
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, Me.Caption
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
