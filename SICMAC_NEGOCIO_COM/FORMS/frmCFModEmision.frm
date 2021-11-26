VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCFModEmision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Modificación"
   ClientHeight    =   7455
   ClientLeft      =   2370
   ClientTop       =   405
   ClientWidth     =   8055
   Icon            =   "frmCFModEmision.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraAvalado 
      Caption         =   "Avalado "
      Enabled         =   0   'False
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
      Left            =   180
      TabIndex        =   22
      Top             =   1800
      Width           =   7800
      Begin VB.CheckBox chkConsorcio 
         Caption         =   "Consorcio"
         Height          =   195
         Left            =   1080
         TabIndex        =   29
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscarAvalado 
         BackColor       =   &H80000004&
         Caption         =   "..."
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
         Left            =   7185
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Buscar al Acreedor"
         Top             =   225
         Width           =   420
      End
      Begin VB.TextBox txtConsorcio 
         Height          =   300
         Left            =   1020
         MaxLength       =   700
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label lblNomAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2460
         TabIndex        =   27
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   4725
      End
      Begin VB.Label lblCodAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1020
         TabIndex        =   26
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Avalado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame fraAcreedor 
      Caption         =   "Acreedor"
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
      Height          =   555
      Left            =   180
      TabIndex        =   13
      Top             =   1200
      Width           =   7800
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2455
         TabIndex        =   16
         Tag             =   "txtnombre"
         Top             =   180
         Width           =   5100
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1020
         TabIndex        =   15
         Tag             =   "txtcodigo"
         Top             =   180
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "E&xaminar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5820
      TabIndex        =   12
      ToolTipText     =   "Buscar Credito"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Height          =   690
      Left            =   180
      TabIndex        =   8
      Top             =   6720
      Width           =   7800
      Begin VB.CommandButton cmdGenerarPDF 
         Caption         =   "Vista Previa"
         Height          =   390
         Left            =   4200
         TabIndex        =   21
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   5520
         TabIndex        =   11
         Top             =   195
         Width           =   1185
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   390
         Left            =   225
         TabIndex        =   10
         ToolTipText     =   "Grabar Datos de Aprobacion de Credito"
         Top             =   195
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   1620
         TabIndex        =   9
         ToolTipText     =   "Ir al Menu Principal"
         Top             =   195
         Width           =   1185
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Carta Fianza"
      Enabled         =   0   'False
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
      Height          =   4035
      Left            =   180
      TabIndex        =   4
      Top             =   2640
      Width           =   7800
      Begin VB.Frame frFinaApo 
         Height          =   2295
         Left            =   40
         TabIndex        =   34
         Top             =   1700
         Width           =   7700
         Begin VB.TextBox TxtFinalidad 
            Height          =   1455
            Left            =   40
            MaxLength       =   700
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   780
            Width           =   7575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Analista"
            Height          =   195
            Index           =   1
            Left            =   3960
            TabIndex        =   40
            Top             =   240
            Width           =   555
         End
         Begin VB.Label lblAnalista 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   4800
            TabIndex        =   39
            Top             =   180
            Width           =   2775
         End
         Begin VB.Label lblApo 
            AutoSize        =   -1  'True
            Caption         =   "Apoderado"
            Height          =   195
            Index           =   3
            Left            =   40
            TabIndex        =   38
            Top             =   250
            Width           =   780
         End
         Begin VB.Label lblApoderado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   960
            TabIndex        =   37
            Top             =   180
            Width           =   2775
         End
         Begin VB.Label lblFina 
            AutoSize        =   -1  'True
            Caption         =   "Finalidad"
            Height          =   195
            Index           =   8
            Left            =   40
            TabIndex        =   36
            Top             =   540
            Width           =   630
         End
      End
      Begin VB.Frame frModOtrsMod 
         Height          =   810
         Left            =   50
         TabIndex        =   31
         Top             =   900
         Width           =   3820
         Begin VB.TextBox txtModOtrsMod 
            Height          =   525
            Left            =   940
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label lblModOtrs 
            Caption         =   "Modalidad Otros"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.ComboBox cboModalidad 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   600
         Width           =   2760
      End
      Begin MSMask.MaskEdBox txtfechaAsig 
         Height          =   315
         Left            =   4920
         TabIndex        =   28
         Top             =   600
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         Height          =   195
         Index           =   6
         Left            =   3960
         TabIndex        =   20
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Emisión"
         Height          =   195
         Index           =   5
         Left            =   3960
         TabIndex        =   19
         Top             =   660
         Width           =   540
      End
      Begin VB.Label lblMontoApr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   4920
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   735
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Afianzado"
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
      Height          =   600
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   7800
      Begin VB.Label lblNomcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2455
         TabIndex        =   3
         Top             =   180
         Width           =   5100
      End
      Begin VB.Label lblCodcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1020
         TabIndex        =   2
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   225
         Width           =   480
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   180
      TabIndex        =   17
      Top             =   120
      Width           =   3645
      _extentx        =   6429
      _extenty        =   688
      texto           =   "Cta. Fianza"
   End
End
Attribute VB_Name = "frmCFModEmision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodCta As String
Dim fpComision As Double
Dim fsEstado As String
Dim fnRenovacion As Integer
Dim fsCodEnvio As String
Dim fbRemesado As Boolean
Dim objPista As COMManejador.Pista
Dim bConsorcio As Boolean
Dim fbAvalado As Boolean
Dim fsNomAvalado As String
Private oPersBuscada As COMDPersona.UCOMPersona
Dim fnPlazoCF As Integer
Dim fdFecVenc As Date
Dim fdFechaAsi As Date 'WIOR 20130330

Function ValidaDatos() As Boolean


If Me.txtfechaAsig.Text = "__/__/____" Then
    MsgBox "Ingrese fecha de Emisión.", vbInformation, "Aviso"
    ValidaDatos = False
    txtfechaAsig.SetFocus
    Exit Function
End If
'WIOR 20130330 *************************
If CDate(fdFechaAsi) > CDate(txtfechaAsig.Text) Then
    MsgBox "Fecha no puede ser menor a la Actual", vbInformation, "Aviso"
    txtfechaAsig.Text = Format(fdFechaAsi, "dd/mm/yyyy")
    ValidaDatos = False
    txtfechaAsig.SetFocus
    Exit Function
End If
'WIOR FIN  ******************************
Dim nPosicion As Integer
nPosicion = InStr(Me.txtfechaAsig.Text, "_")

If nPosicion > 0 Then
    MsgBox "Error en fecha de Emisión.", vbInformation, "Aviso"
    ValidaDatos = False
    txtfechaAsig.SetFocus
    Exit Function
End If
        
If CInt(Mid(Me.txtfechaAsig.Text, 7, 4)) > CInt(Mid(gdFecSis, 7, 4)) Or CInt(Mid(Me.txtfechaAsig.Text, 7, 4)) < (CInt(Mid(gdFecSis, 7, 4)) - 100) Then
    MsgBox "Año fuera del rango.", vbInformation, "Aviso"
    ValidaDatos = False
    txtfechaAsig.SetFocus
    Exit Function
End If
        
If CInt(Mid(Me.txtfechaAsig.Text, 4, 2)) > 12 Or CInt(Mid(Me.txtfechaAsig.Text, 4, 2)) < 1 Then
    MsgBox "Error en Mes.", vbInformation, "Aviso"
    ValidaDatos = False
    txtfechaAsig.SetFocus
    Exit Function
End If
            
Dim nDiasEnMes As Integer
nDiasEnMes = CInt(DateDiff("d", CDate("01" & Mid(Me.txtfechaAsig.Text, 3, 8)), DateAdd("M", 1, CDate("01" & Mid(Me.txtfechaAsig.Text, 3, 8)))))
        
If CInt(Mid(Me.txtfechaAsig.Text, 1, 2)) > nDiasEnMes Or CInt(Mid(Me.txtfechaAsig.Text, 1, 2)) < 1 Then
    MsgBox "Dia fuera del rango.", vbInformation, "Aviso"
    ValidaDatos = False
    txtfechaAsig.SetFocus
    Exit Function
End If

'JOEP20181224 CP
If frModOtrsMod.Visible = True And txtModOtrsMod.Text = "" Then
    MsgBox "Registre Otras Modalidades.", vbInformation, "Aviso"
    ValidaDatos = False
    txtModOtrsMod.SetFocus
    Exit Function
End If
'JOEP20181224 CP
ValidaDatos = True
End Function

Sub LimpiaDatos()
    ActXCodCta.Enabled = True
    ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
    lblNomcli.Caption = ""
    lblCodcli.Caption = ""
    lblNomcli.Caption = ""
    lblCodAcreedor.Caption = ""
    lblNomAcreedor.Caption = ""
    lblTipoCF.Caption = ""
    TxtFinalidad.Text = ""
    'lblModalidad.Caption = ""'WIOR 20130311 COMENTO
    lblMontoApr.Caption = ""
    Me.txtfechaAsig.Text = "__/__/____"
    lblAnalista.Caption = ""
    lblApoderado.Caption = ""
    CmdGrabar.Enabled = False
    fraDatos.Enabled = False
    fraAvalado.Enabled = False
    Me.chkConsorcio.value = 0
    lblCodAvalado.Caption = ""
    cmdGenerarPDF.Enabled = False
    fbRemesado = False
    Me.lblNomAvalado.Caption = ""
    fbAvalado = False
    fsNomAvalado = ""
    cboModalidad.Clear 'WIOR 20130311
End Sub
Sub LimpiaDatosG()
    fraDatos.Enabled = False
    fraAvalado.Enabled = False
    ActXCodCta.Enabled = False
    CmdGrabar.Enabled = False
    cmdGenerarPDF.Enabled = True
    Me.CmdCancelar.Enabled = True
    Me.CmdSalir.Enabled = True
End Sub

Private Sub CargaDatos(ByVal psCta As String)
Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim R As New ADODB.Recordset
Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos
'Dim loConstante As COMDConstantes.DCOMConstantes'WIOR 20130311 COMENTO
Dim loCFValida As COMNCartaFianza.NCOMCartaFianzaValida
Dim lbTienePermiso As Boolean
Dim lnComisionPagada As Double
Dim lnComisionCalculada As Double
Dim ldFechaAsi As Date
Dim rsCartaFianza As ADODB.Recordset
Dim nNumEnvios As Long
ActXCodCta.Enabled = False

fbAvalado = False
fsNomAvalado = ""

    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaAModificar(psCta)
    Set oCF = Nothing
    
    If Not R.BOF And Not R.EOF Then
        Call CP_CargaComboxMod(49000) 'JOEP20181224 CP
        lblCodcli.Caption = R!cPersCod
        lblNomcli.Caption = PstaNombre(R!cPersNombre)
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        ldFechaAsi = R!dAsignacion
        fdFechaAsi = ldFechaAsi 'WIOR 20130330
        lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion)
        lblAnalista.Caption = IIf(IsNull(R!cAnalista), "", R!cAnalista)
        lblApoderado.Caption = IIf(IsNull(R!cApoderado), "", R!cApoderado)
        
        
        If Trim(IIf(IsNull(R!cPersCod), "", R!cPersCod)) = Trim(IIf(IsNull(R!cAvalCod), "", R!cAvalCod)) Then
            If Trim(IIf(IsNull(R!cPersNombre), "", R!cPersNombre)) = Trim(IIf(IsNull(R!cAvalNombre), "", R!cAvalNombre)) Then
                Me.chkConsorcio.value = 0
                lblCodAvalado.Caption = Trim(IIf(IsNull(R!cAvalCod), "", R!cAvalCod))
                lblNomAvalado.Caption = Trim(IIf(IsNull(R!cAvalNombre), "", R!cAvalNombre))
                fbAvalado = True
                fsNomAvalado = lblNomAvalado.Caption
            Else
                chkConsorcio.value = 1
                txtConsorcio.Text = PstaNombre(Trim(IIf(IsNull(R!cAvalNombre), "", R!cAvalNombre)))
                fbAvalado = True
                fsNomAvalado = txtConsorcio.Text
            End If
        Else
            Me.chkConsorcio.value = 0
            If Trim(IIf(IsNull(R!cAvalCod), "", R!cAvalCod)) <> "" Then
                lblCodAvalado.Caption = Trim(IIf(IsNull(R!cAvalCod), "", R!cAvalCod))
                lblNomAvalado.Caption = Trim(IIf(IsNull(R!cAvalNombre), "", R!cAvalNombre))
                fbAvalado = True
                fsNomAvalado = lblNomAvalado.Caption
            Else
                fbAvalado = False
                fsNomAvalado = ""
            End If
        End If

       
        TxtFinalidad.Text = IIf(IsNull(R!cfinalidad), "", R!cfinalidad)
        lblMontoApr = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00"))
        Me.txtfechaAsig.Text = IIf(IsNull(R!dAsignacion), "", Format(R!dAsignacion, "dd/mm/yyyy"))
        fsEstado = R!nPrdEstado
        fnRenovacion = IIf(IsNull(R!nRenovacion), 0, R!nRenovacion)
        fdFecVenc = IIf(IsNull(R!dVencimiento), "", Format(R!dVencimiento, "dd/mm/yyyy"))
        
        fnPlazoCF = CInt(Abs(DateDiff("d", CDate(Me.txtfechaAsig.Text), CDate(fdFecVenc))))
        
        'WIOR 20130311 COMENTO
        'Set loConstante = New COMDConstantes.DCOMConstantes
        '    lblModalidad = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)
        'Set loConstante = Nothing
        
        'WIOR 20130311 **********************************************************
        'Call CargaControles 'Comento JEOP20181224 CP
        cboModalidad.ListIndex = IndiceListaCombo(cboModalidad, R!nModalidad)
        'WIOR FIN ***************************************************************
        
        Me.fraAvalado.Enabled = True
        fraDatos.Enabled = True
        CmdGrabar.Enabled = True
        cmdGenerarPDF.Enabled = True
    End If

End Sub


Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(ActXCodCta.NroCuenta)) > 0 Then
            Call CargaDatos(ActXCodCta.NroCuenta)
        Else
            Call LimpiaDatos
        End If
    End If
End Sub

'JOEP20181224 CP
Private Sub cboModalidad_Click()
    If Trim(Right(cboModalidad.Text, 9)) = "13" Then
         frModOtrsMod.Visible = True
        txtModOtrsMod.Enabled = True
        frmCFModEmision.Height = 7875
        Frame6.top = 6700
        fraDatos.Height = 4035
        frFinaApo.BorderStyle = 0
        frFinaApo.top = 1700
        If txtModOtrsMod.Enabled = True And frFinaApo.Visible = True Then
            txtModOtrsMod.SetFocus
        End If
    Else
        frModOtrsMod.Visible = False
        txtModOtrsMod.Enabled = False
        txtModOtrsMod.Text = ""
        frmCFModEmision.Height = 7125
        Frame6.top = 5903
        fraDatos.Height = 3195
        frFinaApo.BorderStyle = 0
        frFinaApo.top = 880
    End If
End Sub
'JOEP20181224 CP

Private Sub chkConsorcio_Click()
If Me.chkConsorcio.value = 1 Then
        Me.txtConsorcio.Visible = True
        Me.cmdBuscarAvalado.Visible = False
        Me.lblCodAvalado.Visible = False
        Me.lblNomAvalado.Visible = False
        Me.lblCodAvalado.Caption = ""
        Me.lblNomAvalado.Caption = ""
        bConsorcio = True
      
    Else
        Me.txtConsorcio.Visible = False
        Me.txtConsorcio.Text = ""
        Me.cmdBuscarAvalado.Visible = True
        Me.lblCodAvalado.Visible = True
        Me.lblNomAvalado.Visible = True
        bConsorcio = False
    End If
End Sub

Private Sub cmdBuscarAvalado_Click()
Set oPersBuscada = New COMDPersona.UCOMPersona
Set oPersBuscada = frmBuscaPersona.Inicio
If oPersBuscada Is Nothing Then Exit Sub
lblCodAvalado.Caption = oPersBuscada.sPersCod
lblNomAvalado.Caption = oPersBuscada.sPersNombre
Set oPersBuscada = Nothing
End Sub

Private Sub cmdCancelar_Click()
    LimpiaDatos
End Sub

Private Sub cmdExaminar_Click()
Dim lsCta As String

    lsCta = frmCFPersEstado.Inicio(Array(gColocEstModificada), "Cartas Fianza a Modificar", Array(gColCFComercial, gColCFPYME, gColCFTpoProducto), 3)
    If Len(Trim(lsCta)) > 0 Then
        ActXCodCta.NroCuenta = lsCta
        Call CargaDatos(lsCta)
    Else
        Call LimpiaDatos
    End If
End Sub



Private Sub cmdGrabar_Click()
Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza
Dim loImprime As COMNCartaFianza.NCOMCartaFianzaImpre
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnMontoEmi As Currency
Dim ldNeoEmi As Date



vCodCta = ActXCodCta.NroCuenta
lnMontoEmi = Format(lblMontoApr, "#0.00")
ldNeoEmi = Format(Me.txtfechaAsig.Text, "dd/mm/yyyy")
    
If ValidaDatos = False Then
    Exit Sub
End If

If MsgBox("Desea Guardar Modificacion de la Carta Fianza", vbInformation + vbYesNo, "Aviso") = vbYes Then

    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    Dim nPoliza As Long
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida


    nPoliza = CLng(loRs.GetCF_Poliza(vCodCta))
    Set loNCartaFianza = New COMNCartaFianza.NCOMCartaFianza
        Call loNCartaFianza.nCFEmision(vCodCta, lsFechaHoraGrab, lsMovNro, fdFecVenc, lnMontoEmi, 1)
        
        Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
        Dim rsCartaFianza As ADODB.Recordset
        Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
        
        Dim oBase As COMDCredito.DCOMCredActBD
        Set oBase = New COMDCredito.DCOMCredActBD
        
        Call oBase.dDeleteConsorcio(vCodCta)
        Call oBase.dDeleteProductoPersonaRelac(vCodCta, gColRelPersAval)
        If Me.chkConsorcio.value = 1 Then
            Call oBase.dInsertProductoPersona2(vCodCta, Trim(lblCodcli.Caption), gColRelPersAval, Trim(UCase(Me.txtConsorcio.Text)))
        Else
            If Trim(lblCodAvalado.Caption) <> "" Then 'WIOR 20130330
                Call oBase.dInsertProductoPersona(vCodCta, Trim(lblCodAvalado.Caption), gColRelPersAval)
            End If 'WIOR 20130330
        End If
        
        'Call oBase.dUpdateColocCartaFianza(vCodCta, , Trim(Right(cboModalidad.Text, 5)), Trim(Me.txtfechaAsig.Text), fdFecVenc, Trim(TxtFinalidad.Text), , False, , "01/01/1950") 'WIOR 20130311 AGREGO Trim(Right(cboModalidad.Text, 5))'Comento JOEP20181224 CP
        Call oBase.dUpdateColocCartaFianza(vCodCta, , Trim(Right(cboModalidad.Text, 5)), Trim(Me.txtfechaAsig.Text), fdFecVenc, Trim(TxtFinalidad.Text), , False, , "01/01/1950", , Trim(txtModOtrsMod.Text)) 'JOEP20181224 AGREGO Trim(txtModOtrsMod.Text)
        
        
            Call oCartaFianza.ActualizarFolio(CLng(nPoliza), 2, , gdFecSis)
            Call oCartaFianza.QuitarEmisionFolio(vCodCta)
            Set rsCartaFianza = oCartaFianza.UltimoRegistroEnvio(, , CLng(nPoliza))
            If rsCartaFianza.RecordCount > 0 Then
                Call oCartaFianza.ActualizarEnvioFolios(Trim(rsCartaFianza!nCodEnvio), 2)
            End If

            Set oCartaFianza = Nothing
            Set rsCartaFianza = Nothing

        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Modificacion de CF", vCodCta, gCodigoCuenta
        Set objPista = Nothing
    Set loNCartaFianza = Nothing
    
    Dim loImp As COMNCartaFianza.NCOMCartaFianzaReporte
    
    
    CmdGrabar.Enabled = False
    cmdGenerarPDF.Enabled = False
    Call CargaDatosG(vCodCta)
    MsgBox "Carta Fianza Modificada con exito", vbInformation, "Aviso"
    LimpiaDatosG
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ActxCodCta_keypressEnter()
    vCodCta = ActXCodCta.NroCuenta
    If Len(vCodCta) > 0 Then
        Call CargaDatos(vCodCta)
        ActXCodCta.Enabled = False
    Else
        Call LimpiaDatos
    End If
End Sub

Private Sub Form_Load()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fpComision = loParam.dObtieneColocParametro(4001)
Set loParam = Nothing
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
LimpiaDatos
gsOpeCod = gCredEmisionCF
fbRemesado = False
Me.chkConsorcio.value = 0

'JOEP20181224 CP
    frModOtrsMod.Visible = False
    txtModOtrsMod.Enabled = False
    txtModOtrsMod.Text = ""
    frmCFModEmision.Height = 7125
    Frame6.top = 5903
    fraDatos.Height = 3195
    frFinaApo.BorderStyle = 0
    frFinaApo.top = 880
'JOEP20181224 CP

End Sub

Private Function UbiAgencia() As String

    Dim lszona As String
    Dim lscUbiGeoCod As String
    
    Dim lRstZona As ADODB.Recordset
    Dim OlZona  As COMDConstantes.DCOMZonas
    Set OlZona = New COMDConstantes.DCOMZonas

    Dim lRstAgencia As ADODB.Recordset
    Dim OlAgencia  As COMDConstantes.DCOMAgencias
    Set OlAgencia = New COMDConstantes.DCOMAgencias
    
    
    Set lRstAgencia = OlAgencia.RecuperaAgencias(gsCodAge)
        lscUbiGeoCod = lRstAgencia("cUbiGeoCod")
    Set lRstAgencia = Nothing
    Set lRstZona = OlZona.DameUnaZona(lscUbiGeoCod)
        lszona = Trim(lRstZona("cUbiGeoDescripcion"))
        UbiAgencia = lszona
    Set lRstZona = Nothing
    
End Function


Private Function DirAgencia() As String
    Dim lscAgeDireccion As String
    Dim lRstAgencia As ADODB.Recordset
    Dim OlAgencia  As COMDConstantes.DCOMAgencias
    Set OlAgencia = New COMDConstantes.DCOMAgencias
    
    
    Set lRstAgencia = OlAgencia.RecuperaAgencias(gsCodAge)
        lscAgeDireccion = lRstAgencia("cAgeDireccion")
        DirAgencia = lscAgeDireccion
    Set lRstAgencia = Nothing
    
End Function

Sub ImpreDoc(ByVal psCtaCod As String)
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    
    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    Dim rsCartaFianza As ADODB.Recordset
    Dim nDias As Integer
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim lsFechas As String
    
    Dim nPoliza As Long
    Dim cDirecAgencia As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim nCFPoliza As Long
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCtaCod)
    Set lrDataT = loRs.RecuperaDatosT(psCtaCod)
    Set lrDataCR = loRs.RecuperaDatosAcreedor(psCtaCod)
    
    cDirecAgencia = loRs.Get_Agencia_CF(psCtaCod)
End Sub


Private Sub txtNumPoliza_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
Private Sub cmdGenerarPDF_Click()
Dim oCF As COMNCartaFianza.NCOMCartaFianzaValida
Dim nPoliza As Long
On Error GoTo ErrorGenerarPdf
vCodCta = ActXCodCta.NroCuenta

If ValidaDatos = False Then
    Exit Sub
End If


Call ImprimirPDF(vCodCta, fbAvalado, "1", 1)

MsgBox "Archivo Previo Generado Satisfacoriamente.", vbInformation, "Aviso"
Exit Sub
ErrorGenerarPdf:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub ImprimirPDF(ByVal psCodCta As String, ByVal pbAvalado As Boolean, ByVal psNumFolio As String, ByVal nTipo As Integer)
    On Error GoTo ErrorImprimirPDF
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    Dim dfechaini As Date
    Dim nCFPoliza As Long
    Dim sParrafo1 As String
    Dim sParrafo2 As String
    Dim sParrafo3 As String
    Dim sParrafo4 As String
    Dim nTamano As Integer
    Dim nValidar As Double
    Dim nTop As Integer
    Dim sFechaActual As String
    Dim sSenores As String
    Dim sAval As String
    Dim sSolicitante As String
    Dim sMonto As String
    Dim sModalidad As String
    Dim sFinalidad As String
    Dim dfechafin As Date
    Dim sVencimiento As String
    Dim sDireccion As String
    Dim lnPosicion As Integer
    Dim oDoc  As cPDF
    
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCodCta)
    Set lrDataT = loRs.RecuperaDatosT(psCodCta)
    Set oDoc = New cPDF
    
    nCFPoliza = psNumFolio
    
    'Creacion de Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Modificación de Carta Fianza Nº " & psCodCta
    oDoc.Title = "Modificación de Carta Fianza Nº " & psCodCta
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & IIf(nTipo = 1, "PrevioModificación", "") & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Bold, WinAnsiEncoding
    
    oDoc.NewPage A4_Vertical
 
    'sFechaActual = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")'WIOR 20130330
    sFechaActual = Format(Trim(txtfechaAsig.Text), "dd") & " de " & Format(Trim(txtfechaAsig.Text), "mmmm") & " del " & Format(Trim(txtfechaAsig.Text), "yyyy")
    sSenores = PstaNombre(lblNomAcreedor, True)
    If pbAvalado Then
        sAval = PstaNombre(UCase(fsNomAvalado), False)
    End If
    sSolicitante = PstaNombre(lblNomcli, True)

    Dim sSaldo As String
    sSaldo = Format(lrDataT!nSaldo, "#,###0.00")
    sMonto = IIf(Mid(psCodCta, 9, 1) = "1", "S/. ", "$. ") & sSaldo & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(psCodCta, 9, 1) = "2", "", " Y " & IIf(InStr(1, sSaldo, ".") = 0, "00", Mid(sSaldo, InStr(1, sSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(psCodCta, 9, 1) = "1", "NUEVOS SOLES)", "US DOLARES)")
    'sModalidad = Trim(lblModalidad.Caption)'COMENTADO POR WIOR 20130311
    sModalidad = Trim(Left(cboModalidad.Text, Len(cboModalidad.Text) - 5)) 'WIOR 20130311
    sFinalidad = Trim(TxtFinalidad.Text)
    

    dfechafin = CDate(fdFecVenc)
    sVencimiento = Format(dfechafin, "dd") & " de " & Format(dfechafin, "mmmm") & " del " & Format(dfechafin, "yyyy")
    sDireccion = loRs.Get_Agencia_CF(psCodCta)
    lnPosicion = InStr(sDireccion, "(")
    sDireccion = Left(sDireccion, lnPosicion - 2)
    
    oDoc.WTextBox 70, 50, 10, 450, Left(psCodCta, 3) & "-" & Mid(psCodCta, 4, 2) & "-" & Mid(psCodCta, 6, 3) & "-" & Mid(psCodCta, 9, 10), "F1", 12, hRight
    oDoc.WTextBox 120, 50, 10, 450, "CARTA FIANZA N° " & Format(nCFPoliza, "0000000"), "F1", 12, hCenter
    oDoc.WTextBox 170, 50, 10, 450, sFechaActual, "F1", 12, hRight
    oDoc.WTextBox 220, 50, 10, 450, "Señores:", "F1", 12, hLeft
    oDoc.WTextBox 232, 50, 10, 450, sSenores, "F1", 12, hLeft
    oDoc.WTextBox 260, 50, 10, 450, "Ciudad.-", "F1", 12, hLeft
    oDoc.WTextBox 280, 50, 10, 450, "Muy Señores Nuestros:", "F1", 12, hLeft
    sAval = " garantizando a " & sAval
    sParrafo1 = "A solicitud de " & sSolicitante & ", otorgamos por el presente " & _
                "documento una fianza solidaria, irrevocable, incondicional, de " & _
                "ejecución inmediata, con renuncia expresa al beneficio de " & _
                "excusión e indivisible, a favor de ustedes" & IIf(pbAvalado = True, sAval, "") & _
                ", hasta por la suma de " & sMonto & ", a fin de garantizar " & _
                "la Carta Fianza por " & IIf(Trim(Right(cboModalidad.Text, 9)) = 13, Trim(txtModOtrsMod.Text), sModalidad) & ", objeto del proceso: " & sFinalidad & "."
    'Agrego JOEP20181224 CP IIf(Trim(Right(cboModalidad.Text, 9)) = 13, Trim(txtModOtrsMod.Text), sModalidad)
    nTamano = Len(sParrafo1)
    nValidar = nTamano / 72
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    nTop = 270
    
    oDoc.WTextBox nTop, 0, nTamano * 20, 580, String(20, "-") & " " & sParrafo1, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True

    nTop = nTop + (nTamano * 10) + 12
      
    sParrafo2 = "Dejamos claramente establecido que la presente " & String(1, vbTab) & "Carta " & String(1, vbTab) & "Fianza no " & _
                "podrá ser usada " & String(1, vbTab) & "para operaciones comprendidas en la prohibición " & _
                "indicada en el inciso ''5'' del Articulo 217 de la " & String(1, vbTab) & "Ley  26702, Ley " & _
                "General del " & String(1, vbTab) & "Sistema " & String(1, vbTab) & "Financiero y del Sistema de Seguros y Orgánica " & _
                "de la Superintendencia de --- Banca y Seguros."
                
                
    nTamano = Len(sParrafo2)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo2, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    oDoc.WTextBox nTop + 75, 520, 10, 20, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12) + 12
                       
    sParrafo3 = "Por efecto de este compromiso la CAJA MUNICIPAL DE MAYNAS S.A " & _
                        "asume con su fiado las responsabilidades en que éste llegara a " & _
                        "incurrir siempre que el " & String(1, vbTab) & "monto de las  mismas  no " & String(1, vbTab) & "exceda por ningún " & _
                        "motivo de la suma antes mencionada y que estén estrictamente " & _
                        "vinculadas al cumplimiento de lo arriba indicado."

    nTamano = Len(sParrafo3)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo3, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12) + 13
    
    sParrafo4 = "" & _
                        "La presente garantía rige a partir de la fecha y vencerá " & _
                        "el " & sVencimiento & ". Cualquier  reclamo en virtud de esta " & _
                        "garantía deberá ceñirse estrictamente a lo estipulado por " & _
                        "el Art. 1898 del Código Civil y deberá ser formulado por vía " & _
                        "notarial y en nuestra oficina ubicada en " & sDireccion & "."
    nTamano = Len(sParrafo4)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo4, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 13) + 80

    oDoc.WTextBox nTop, 50, 10, 450, "Atentamente,", "F1", 12, hCenter, vMiddle, , , , False
    oDoc.WTextBox nTop + 12, 50, 10, 450, "CAJA MUNICIPAL DE MAYNAS S.A", "F1", 12, hCenter, vMiddle, , , , False

    oDoc.PDFClose
    oDoc.Show
    
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub lblNomAvalado_Change()
fsNomAvalado = Me.lblNomAvalado
If Trim(fsNomAvalado) <> "" Then
    fbAvalado = True
Else
    fbAvalado = False
End If
End Sub

Private Sub txtConsorcio_Change()
fsNomAvalado = Me.txtConsorcio.Text
If Trim(fsNomAvalado) <> "" Then
    fbAvalado = True
Else
    fbAvalado = False
End If
End Sub
Private Sub CargaDatosG(ByVal psCta As String)

Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim R As New ADODB.Recordset
Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos
'Dim loConstante As COMDConstantes.DCOMConstantes 'WIOR 20130311 COMENTO
Dim loCFValida As COMNCartaFianza.NCOMCartaFianzaValida
Dim lbTienePermiso As Boolean
Dim lnComisionPagada As Double
Dim lnComisionCalculada As Double
Dim ldFechaAsi As Date
Dim rsCartaFianza As ADODB.Recordset
Dim nNumEnvios As Long
ActXCodCta.Enabled = False

fbAvalado = False
fsNomAvalado = ""

    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaEmision(psCta)
    Set oCF = Nothing
    
    If Not R.BOF And Not R.EOF Then
        Call CP_CargaComboxMod(49000) 'JOEP20181224 CP
        lblCodcli.Caption = R!cPersCod
        lblNomcli.Caption = PstaNombre(R!cPersNombre)
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        ldFechaAsi = R!dAsignacion
        lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion)
        lblAnalista.Caption = IIf(IsNull(R!cAnalista), "", R!cAnalista)
        lblApoderado.Caption = IIf(IsNull(R!cApoderado), "", R!cApoderado)
        
        If Trim(IIf(IsNull(R!cPersCod), "", R!cPersCod)) = Trim(IIf(IsNull(R!cAvalCod), "", R!cAvalCod)) Then
            If Trim(IIf(IsNull(R!cPersNombre), "", R!cPersNombre)) = Trim(IIf(IsNull(R!cAvalNombre), "", R!cAvalNombre)) Then
                Me.chkConsorcio.value = 0
                lblCodAvalado.Caption = Trim(IIf(IsNull(R!cAvalCod), "", R!cAvalCod))
                lblNomAvalado.Caption = Trim(IIf(IsNull(R!cAvalNombre), "", R!cAvalNombre))
                fbAvalado = True
                fsNomAvalado = lblNomAvalado.Caption
            Else
                chkConsorcio.value = 1
                txtConsorcio.Text = PstaNombre(Trim(IIf(IsNull(R!cAvalNombre), "", R!cAvalNombre)))
                fbAvalado = True
                fsNomAvalado = txtConsorcio.Text
            End If
        Else
            Me.chkConsorcio.value = 0
            If Trim(IIf(IsNull(R!cAvalCod), "", R!cAvalCod)) <> "" Then
                lblCodAvalado.Caption = Trim(IIf(IsNull(R!cAvalCod), "", R!cAvalCod))
                lblNomAvalado.Caption = Trim(IIf(IsNull(R!cAvalNombre), "", R!cAvalNombre))
                fbAvalado = True
                fsNomAvalado = lblNomAvalado.Caption
            Else
                fbAvalado = False
                fsNomAvalado = ""
            End If
        End If

       
        TxtFinalidad.Text = IIf(IsNull(R!cfinalidad), "", R!cfinalidad)
        lblMontoApr = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00"))
        Me.txtfechaAsig.Text = IIf(IsNull(R!dAsignacion), "", Format(R!dAsignacion, "dd/mm/yyyy"))
        fsEstado = R!nPrdEstado
        fnRenovacion = IIf(IsNull(R!nRenovacion), 0, R!nRenovacion)
        fdFecVenc = IIf(IsNull(R!dVencimiento), "", Format(R!dVencimiento, "dd/mm/yyyy"))

        'WIOR 20130311 COMENTO
        'Set loConstante = New COMDConstantes.DCOMConstantes
        '    lblModalidad = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)
        'Set loConstante = Nothing
        'WIOR 20130311 **********************************************************
        'Call CargaControles 'comento JOEP20181224 CP
        cboModalidad.ListIndex = IndiceListaCombo(cboModalidad, R!nModalidad)
        
        'WIOR FIN ***************************************************************
                
        Me.fraAvalado.Enabled = True
        fraDatos.Enabled = True
        CmdGrabar.Enabled = True
        cmdGenerarPDF.Enabled = True
    End If

End Sub


Private Sub txtfechaAsig_LostFocus()
If ValidaDatos Then
    fdFecVenc = DateAdd("D", fnPlazoCF, CDate(Me.txtfechaAsig.Text))
End If
End Sub

'WIOR 20130311 ***************************************************
Private Sub CargaControles()
    'Carga Modalidad de Carta Fianza
    Call CargaComboConstante(gColCFModalidad, cboModalidad)
    Call CambiaTamañoCombo(cboModalidad, 300)
End Sub
'WIOR FIN ********************************************************

'JOEP20181218 CP
Private Sub CP_CargaComboxMod(ByVal nParCod As Long)
Dim objCatalogoLlenaCombox As COMDCredito.DCOMCredito
Dim rsCatalogoCombox As ADODB.Recordset
Set objCatalogoLlenaCombox = New COMDCredito.DCOMCredito
Set rsCatalogoCombox = objCatalogoLlenaCombox.getCatalogoCombo("514", nParCod)

If Not (rsCatalogoCombox.BOF And rsCatalogoCombox.EOF) Then
    If nParCod = 49000 Then
        cboModalidad.Clear
        Call Llenar_Combo_con_Recordset(rsCatalogoCombox, cboModalidad)
        Call CambiaTamañoCombo(cboModalidad, 300)
    End If
End If

End Sub
'JOEP20181218 CP
