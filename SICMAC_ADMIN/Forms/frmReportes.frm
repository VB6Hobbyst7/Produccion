VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   5970
   ClientLeft      =   1920
   ClientTop       =   1740
   ClientWidth     =   8415
   HelpContextID   =   210
   Icon            =   "frmReportes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTCambio 
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
      Height          =   750
      Left            =   4980
      TabIndex        =   16
      Top             =   4710
      Width           =   2835
      Begin VB.OptionButton optMoneda 
         Caption         =   "A&justado"
         Height          =   255
         Index           =   3
         Left            =   4500
         TabIndex        =   21
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtTipCambio 
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
         Height          =   345
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Cambio"
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   330
         Width           =   1380
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha"
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
      Height          =   675
      Left            =   4980
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   390
         TabIndex        =   12
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   300
         Width           =   135
      End
   End
   Begin MSComctlLib.TreeView tvOpe 
      Height          =   5355
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   9446
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imglstFiguras"
      BorderStyle     =   1
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
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   4980
      TabIndex        =   3
      Top             =   0
      Width           =   2835
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &Extranjera"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &Nacional"
         Height          =   285
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   330
      Left            =   5670
      TabIndex        =   2
      Top             =   5580
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   330
      Left            =   6960
      TabIndex        =   1
      Top             =   5580
      Width           =   1245
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   180
      Top             =   4860
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
            Picture         =   "frmReportes.frx":030A
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":065C
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":09AE
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":0D00
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraFechaRango 
      Caption         =   "Rango de Fechas"
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
      Height          =   675
      Left            =   4980
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   3315
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   300
         Left            =   510
         TabIndex        =   7
         Top             =   255
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   315
         Left            =   2010
         TabIndex        =   8
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   1770
         TabIndex        =   10
         Top             =   315
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Periodo"
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
      Height          =   1230
      Left            =   4980
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   2835
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmReportes.frx":1052
         Left            =   870
         List            =   "frmReportes.frx":107A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   270
         Width           =   1815
      End
      Begin VB.TextBox txtAnio 
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
         Height          =   315
         Left            =   1575
         MaxLength       =   4
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   285
         TabIndex        =   20
         Top             =   330
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   780
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lExpand  As Boolean
Dim lExpandO As Boolean
Dim sArea  As String

Public Sub Inicio(sObj As String, Optional plExpandO As Boolean = False)
    sArea = sObj
    lExpandO = plExpandO
    Me.Show 0, frmMdiMain
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
   If fraFechaRango.Visible Then
      If Not ValFecha(txtFechaDel) Then
         txtFechaDel.SetFocus: Exit Function
      End If
      If Not ValFecha(txtFechaAl) Then
         txtFechaAl.SetFocus: Exit Function
      End If
   End If
   If fraFecha.Visible Then
      If Not ValFecha(txtfecha) Then
         txtfecha.SetFocus: Exit Function
      End If
   End If
   If fraPeriodo.Visible Then
      If nVal(txtAnio) = 0 Then
         MsgBox "Ingrese Año para generar Reporte...", vbInformation, "¡Aviso!"
         txtAnio.SetFocus
         Exit Function
      End If
      If cboMes.ListIndex = -1 Then
        MsgBox "Selecciones Mes para generar Reporte...", vbInformation, "¡Aviso!"
        cboMes.SetFocus
        Exit Function
      End If
   End If
   If Not tvOpe.SelectedItem.Child Is Nothing Then
        MsgBox "Seleccione Reporte de último Nivel", vbInformation, "¡Aviso!"
        tvOpe.SetFocus
        Exit Function
   End If
ValidaDatos = True
End Function

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAnio.SetFocus
    End If
End Sub

Private Sub cmdGenerar_Click()
Dim lsMoneda As String
On Error GoTo cmdGenerarErr
    If tvOpe.Nodes.Count = 0 Then
        MsgBox "Lista de operaciones se encuentra vacia", vbInformation, "Aviso"
        Exit Sub
    End If
    If tvOpe.SelectedItem.Tag = "1" Then
        MsgBox "Operación seleccionada no valida...!", vbInformation, "Aviso"
        tvOpe.SetFocus
        Exit Sub
    End If
    If Not ValidaDatos Then Exit Sub
    
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    lsMoneda = IIf(optMoneda(0).value, "1", "2")
    If Left(tvOpe.SelectedItem.Key, 1) <> "P" Then
        gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescHijo = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescPadre = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60))
    Else
      gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    End If
    
    'gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    
    Select Case Mid(gsOpeCod, 1, 5)
        Case Mid(gContRepBaseFormula, 1, 5)
            frmRepBaseFormula.Inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo
            
    End Select
    
    Select Case Mid(gsOpeCod, 1, 6)
            
        '***************** CARTAS FIANZA ********************************
        Case OpeCGCartaFianzaRepIngreso, OpeCGCartaFianzaRepIngresoME
            ImprimeCartasFianza txtFechaDel, txtFechaAl, True
        Case OpeCGCartaFianzaRepSalida, OpeCGCartaFianzaRepSalidaME
            ImprimeCartasFianza txtFechaDel, txtFechaAl, False
        
        '********************* REPORTES DE CAJA GENERAL **************************
        
        Case OpeCGRepRepBancosFlujoMensMN, OpeCGRepRepBancosFlujoMensME
            frmCajaGenRepFlujos.Show , Me
        Case OpeCGRepRepBancosFlujoPFMN, OpeCGRepRepBancosFlujoPFME
            frmCajaGenRepFlujos.Show , Me
        Case OpeCGRepRepBancosSaldosCtasMN, OpeCGRepRepBancosSaldosCtasME
        
        Case OpeCGRepRepCMACSFlujoMensMN, OpeCGRepRepCMACSFlujoMensME
            frmCajaGenRepFlujos.Show , Me
        Case OpeCGRepRepCMACSFlujoPFMN, OpeCGRepRepCMACSFlujoPFME
            frmCajaGenRepFlujos.Show , Me
            
        Case OpeCGRepRepCMACSSaldosCtasMN, OpeCGRepRepCMACSSaldosCtasME
        
        Case OpeCGRepRepOPGirMN, OpeCGRepRepOPGirME
        
        Case OpeCGRepRepChqRecDetMN, OpeCGRepRepChqRecDetME
        Case OpeCGRepRepChqRecResMN, OpeCGRepRepChqRecResME
        
        Case OpeCGRepRepChqValDetMN, OpeCGRepRepChqValDetME
        Case OpeCGRepRepChqValResMN, OpeCGRepRepChqValResME
        
        Case OpeCGRepRepChqValorizadosDetMN, OpeCGRepRepChqValorizadosDetME
        Case OpeCGRepRepChqValorizadosResMN, OpeCGRepRepChqValorizadosResME
        
        Case OpeCGRepRepChqAnulDetMN, OpeCGRepRepChqAnulDetME
        Case OpeCGRepRepChqAnulResMN, OpeCGRepRepChqAnulResME

        Case OpeCGRepRepChqObsDetMN, OpeCGRepRepChqObsDetME
        Case OpeCGRepRepChqObsResMN, OpeCGRepRepChqObsResME
        
        'ENCAJE
        Case OpeCGRepEncajeConsolSdoEnc, OpeCGRepEncajeConsolSdoEncME
            frmCajaGenReportes.ConsolidaSdoEnc gsOpeCod, txtFechaDel, txtFechaAl
        Case OpeCGRepEncajeAgencia, OpeCGRepEncajeAgenciaME
        
        Case OpeCGRepEncajeConsolPosLiq, OpeCGRepEncajeConsolPosLiqME
            frmCajaGenReportes.ConsolidaSdoEnc gsOpeCod, txtFechaDel, txtFechaAl, 2
        
        'Informe de Encaje al BCR
         Case RepCGEncBCRObligacion, RepCGEncBCRObligacionME, _
              RepCGEncBCRCredDeposi, RepCGEncBCRCredDeposiME, _
              RepCGEncBCRCredRecibi, RepCGEncBCRCredRecibiME, _
              RepCGEncBCRObligaExon, RepCGEncBCRObligaExonME, _
              RepCGEncBCRLinCredExt, RepCGEncBCRLinCredExtME
            frmRepEncajeBCR.ImprimeEncajeBCR gsOpeCod, txtFechaDel, txtFechaAl
        
        '************************* CNTABILIDAD **********************************
        Case gContLibroDiario
               frmContabDiario.Show 0, Me
        Case gContLibroMayor
               frmContabMayor.Show 0, Me
        Case gContLibroMayCta
               frmContabMayorDet.Show 0, Me
        Case gContRegCompraGastos
               frmRegCompraGastos.Show 0, Me
        Case gContRegVentas
               frmRegVenta.Show 0, Me
        Case gContRepCompraVenta
            frmRepResCVenta.Show 0, Me

        'Otros Ajustes
        Case gContAjReclasiCartera
            frmAjusteReCartera.Show , Me
        Case gContAjReclasiGaranti
            frmAjusteGarantias.Show , Me
        Case gContAjInteresDevenga
            frmAjusteIntDevengado.Inicio True
        Case gContAjInteresSuspens
            frmAjusteIntDevengado.Inicio False
            
        'ANEXOS
        Case gContAnx02CredTpoGarantia 'Creditos Directos por Tipo de Garantia
            frmAnx02CreDirGarantia.GeneraAnx02CreditosTipoGarantia txtAnio, cboMes.ListIndex + 1, nVal(txtTipCambio), cboMes.Text
        Case gContAnx03FujoCrediticio
            frmAnx02CreDirGarantia.GeneraAnx03FlujoCrediticioPorTipoCred txtAnio, cboMes.ListIndex + 1, nVal(txtTipCambio), cboMes.Text
        
        Case gContAnx07
            frmAnexo7RiesgoInteres.Inicio True
        Case gContAnx15A_Estad      'Informe Estadístico
            frmAnx15AEstadDia.ImprimeEstadisticaDiaria gsOpeCod, lsMoneda, txtfecha
        Case gContAnx15A_Efect      'Descomposición de Efectivo
            frmAnx15AEfectivoCaja.ImprimeEfectivoCaja gsOpeCod, lsMoneda, txtfecha
        Case gContAnx15A_Banco      'Consolidado Bancos
            frmAnx15AConsolBancos.ImprimeConsolidaBancos gsOpeCod, lsMoneda, txtfecha
        Case gContAnx15A_Repor      'Anexo 15A
            frmAnx15AReporte.ImprimeAnexo15A gsOpeCod, lsMoneda, txtfecha
        Case gContAnx17A_FSD
            frmFondoSeguroDep.Inicio txtFechaDel, txtFechaAl
    End Select
Exit Sub
cmdGenerarErr:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso Error"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
tvOpe.SetFocus
End Sub

Private Sub Form_Load()
    Dim sCod As String
    On Error GoTo ERROR
    CentraForm Me
    frmMdiMain.Enabled = False
    
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
    
    txtAnio = Year(gdFecSis)
    cboMes.ListIndex = Month(gdFecSis) - 1
    
    Exit Sub
ERROR:
    MsgBox TextErr(Err.Description), vbExclamation, Me.Caption
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
    frmMdiMain.Enabled = True
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
    MsgBox TextErr(Err.Description), vbExclamation, Me.Caption
End Sub

Private Sub ActivaControles(Optional plFechaRango As Boolean = True, _
                           Optional plFechaAl As Boolean = False, _
                           Optional plFechaPeriodo As Boolean = False, _
                           Optional plTpoCambio As Boolean = False _
                           )
fraFechaRango.Visible = plFechaRango
fraFecha.Visible = plFechaAl
fraPeriodo.Visible = plFechaPeriodo
fraTCambio.Visible = plTpoCambio
End Sub

Private Sub tvOpe_NodeClick(ByVal Node As MSComctlLib.Node)
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    
    Select Case Mid(gsOpeCod, 1, 5)
        Case Mid(gContRepBaseFormula, 1, 5)
            ActivaControles False, False, True
    End Select
    
    Select Case Mid(gsOpeCod, 1, 6)
        '***************** CARTAS FIANZA ********************************
        Case OpeCGCartaFianzaRepIngreso, OpeCGCartaFianzaRepIngresoME
            ActivaControles True
            txtFechaDel.SetFocus
            
        Case OpeCGCartaFianzaRepSalida, OpeCGCartaFianzaRepSalidaME
            ActivaControles True
            txtFechaDel.SetFocus
            
        '********************* REPORTES DE CAJA GENERAL **************************
        Case OpeCGRepRepBancosFlujoMensMN, OpeCGRepRepBancosFlujoMensME
            ActivaControles False
        Case OpeCGRepRepBancosFlujoPFMN, OpeCGRepRepBancosFlujoPFME
            ActivaControles False
        Case OpeCGRepRepBancosSaldosCtasMN, OpeCGRepRepBancosSaldosCtasME
            ActivaControles False
        Case OpeCGRepRepCMACSFlujoMensMN, OpeCGRepRepCMACSFlujoMensME
            ActivaControles False
        Case OpeCGRepRepCMACSFlujoPFMN, OpeCGRepRepCMACSFlujoPFME
            ActivaControles False
        Case OpeCGRepRepCMACSSaldosCtasMN, OpeCGRepRepCMACSSaldosCtasME
            ActivaControles False
        Case OpeCGRepRepOPGirMN, OpeCGRepRepOPGirME
            ActivaControles False
        
        Case OpeCGRepRepChqRecDetMN, OpeCGRepRepChqRecDetME
            ActivaControles False
        Case OpeCGRepRepChqRecResMN, OpeCGRepRepChqRecResME
            ActivaControles False
        
        Case OpeCGRepRepChqValDetMN, OpeCGRepRepChqValDetME
            ActivaControles False
        Case OpeCGRepRepChqValResMN, OpeCGRepRepChqValResME
            ActivaControles False
        
        Case OpeCGRepRepChqValorizadosDetMN, OpeCGRepRepChqValorizadosDetME
            ActivaControles False
        Case OpeCGRepRepChqValorizadosResMN, OpeCGRepRepChqValorizadosResME
            ActivaControles False
        
        Case OpeCGRepRepChqAnulDetMN, OpeCGRepRepChqAnulDetME
            ActivaControles False
        Case OpeCGRepRepChqAnulResMN, OpeCGRepRepChqAnulResME
            ActivaControles False

        Case OpeCGRepRepChqObsDetMN, OpeCGRepRepChqObsDetME
            ActivaControles False
        Case OpeCGRepRepChqObsResMN, OpeCGRepRepChqObsResME
            ActivaControles False
        
         'ENCAJE
        Case OpeCGRepEncajeConsolSdoEnc, OpeCGRepEncajeConsolSdoEncME, OpeCGRepEncajeAgencia, OpeCGRepEncajeAgenciaME, OpeCGRepEncajeConsolPosLiq, OpeCGRepEncajeConsolPosLiqME
            ActivaControles True
            txtFechaDel.SetFocus
            
        'Informe de Encaje al BCR
         Case RepCGEncBCRObligacion, RepCGEncBCRObligacionME, RepCGEncBCRCredDeposi, RepCGEncBCRCredDeposiME, RepCGEncBCRCredRecibi, RepCGEncBCRCredRecibiME, RepCGEncBCRObligaExon, RepCGEncBCRObligaExonME, RepCGEncBCRLinCredExt, RepCGEncBCRLinCredExtME
            ActivaControles True
            txtFechaDel.SetFocus
            
        '************************* CONTABILIDAD **********************************
        Case gContLibroDiario
            ActivaControles False
        Case gContLibroMayor
            ActivaControles False
        Case gContLibroMayCta
            ActivaControles False
        Case gContRegCompraGastos
            ActivaControles False
        Case gContRegVentas
            ActivaControles False
        Case gContRepCompraVenta
            ActivaControles False

        'Otros Ajustes
        Case gContAjReclasiCartera
            ActivaControles False
        Case gContAjReclasiGaranti
            ActivaControles False
        Case gContAjInteresDevenga
            ActivaControles False
        Case gContAjInteresSuspens
            ActivaControles False
            
        'ANEXOS
        Case gContAnx02CredTpoGarantia, gContAnx03FujoCrediticio
            ActivaControles False, False, True, True
            
        Case gContAnx07
        Case gContAnx15A_Estad, gContAnx15A_Efect, gContAnx15A_Banco, gContAnx15A_Repor
            ActivaControles False, True
            txtfecha.SetFocus
        Case gContAnx17A_FSD
            ActivaControles True, , , False
            txtFechaDel.SetFocus
    End Select
End Sub

Private Sub tvOpe_Collapse(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H80000008"
End Sub

Private Sub tvOpe_Expand(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H8000000D"
End Sub

Private Sub tvOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdGenerar_Click
    End If
End Sub

Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If fraTCambio.Visible Then
        txtTipCambio.SetFocus
    Else
        cmdGenerar.SetFocus
    End If
End If
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtfecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtfecha) = True Then
       cmdGenerar.SetFocus
    End If
End If
End Sub

Private Sub txtFechaDel_GotFocus()
fEnfoque txtFechaDel
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaDel) = True Then
       txtFechaAl.SetFocus
    End If
End If
End Sub

Private Sub txtFechaAl_GotFocus()
fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaAl) = True Then
       cmdGenerar.SetFocus
    End If
End If
End Sub

Private Sub txtTipCambio_GotFocus()
fEnfoque txtTipCambio
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub
