VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingLiberar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contingencias: Liberar"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "frmContingLiberar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabConting 
      Height          =   2940
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5186
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Liberar / Ejecutar Contingencia"
      TabPicture(0)   =   "frmContingLiberar.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdLiberar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdLiberar 
         Caption         =   "Liberar"
         Height          =   345
         Left            =   3360
         TabIndex        =   13
         Top             =   2400
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   345
         Left            =   4560
         TabIndex        =   12
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         Caption         =   "Banco Interviniente"
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5535
         Begin Sicmact.TxtBuscar txtBancos 
            Height          =   315
            Left            =   960
            TabIndex        =   9
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Enabled         =   0   'False
            EnabledText     =   0   'False
         End
         Begin VB.Label lblCta 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            TabIndex        =   14
            Top             =   1515
            Width           =   4430
         End
         Begin VB.Label Label3 
            Caption         =   "Banco : "
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1130
            Width           =   1200
         End
         Begin VB.Label lblBanco 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2010
            TabIndex        =   10
            Top             =   1080
            Width           =   3380
         End
         Begin VB.Label lblDemanda 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   4200
            TabIndex        =   8
            Top             =   645
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label lblMoneda 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1560
            TabIndex        =   7
            Top             =   645
            Width           =   390
         End
         Begin VB.Label lblProvision 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1995
            TabIndex        =   6
            Top             =   645
            Width           =   1155
         End
         Begin VB.Label lblLabelDem 
            Caption         =   "Demanda : "
            Height          =   255
            Left            =   3360
            TabIndex        =   5
            Top             =   675
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Contingencia :"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Contingencia : "
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   690
            Width           =   1335
         End
         Begin VB.Label lblTipoConting 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1560
            TabIndex        =   2
            Top             =   270
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frmContingLiberar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingLiberar
'** Descripción : Liberar Contingencias creado segun RFC056-2012
'** Creación : JUEZ, 20120622 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim oCtaIf As NCajaCtaIF
Dim psOpeCod As String
Dim sNumRegistro As String
Dim sSubCtaContBC As String
Dim sSubCtaContCta As String
Dim nCalif As Integer
Dim nTipo As Integer
Dim nItem As Integer

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdLiberar_Click()
    If txtBancos.Text <> "" Then
    
    If MsgBox("Está seguro de Liberar la Contingencia? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set oConting = New DContingencia
        Dim sMensaje As String
        Dim oMov As DMov
        Dim oImpr As NContImprimir
        Dim oPrevio As clsPrevioFinan
        Set oPrevio = New clsPrevioFinan
        Set oMov = New DMov
        Set oImpr = New NContImprimir
        gdFecha = gdFecSis
        gsMovNro = oMov.GeneraMovNro(gdFecha, gsCodAge, gsCodUser)

        Call oConting.LiberarContingencia(nItem, sNumRegistro, sSubCtaContBC, sSubCtaContCta, nCalif, nTipo, IIf(Me.lblMoneda.Caption = gcPEN_SIMBOLO, 1, 2), _
                                        lblProvision.Caption, lblDemanda.Caption, psOpeCod, gsMovNro, sMensaje) 'marg ers044-2016
        If sMensaje = "" Then
            MsgBox "Se ha liberado con exito la Contingencia", vbInformation, "Aviso"
            oPrevio.Show oImpr.ImprimeAsientoContable(gsMovNro, 66, 79), gsOpeDesc, False, 66, gImpresora
            'ImprimeAsientoContable gsMovNro
            Unload Me
        Else
            MsgBox sMensaje, vbInformation, "Aviso!"
        End If
    Else
        MsgBox "Es necesario que ingrese el banco", vbInformation, "Aviso"
    End If
End Sub

Public Sub Liberar(ByVal psNumRegistro As String)
    sNumRegistro = psNumRegistro
    psOpeCod = gLiberarContigencia
    Set oConting = New DContingencia
    
    Set rs = oConting.BuscaContingenciaParaLiberar(sNumRegistro)
    
    nItem = rs!nIdInfTec
    lblTipoConting.Caption = IIf(Left(rs!cNumRegistro, 1) = 1, "Activo", "Pasivo") & " Contingente"
    lblMoneda.Caption = rs!cMoneda
    lblProvision.Caption = Format(rs!nProvision, "#,##0.00")
    nCalif = rs!nCalif
    nTipo = rs!nTipo
    If nTipo = 1 Then
        lblLabelDem.Visible = True
        lblDemanda.Visible = True
    Else
        lblLabelDem.Visible = False
        lblDemanda.Visible = False
    End If
    lblDemanda.Caption = Format(rs!nDemandaLab, "#,##0.00")
    
    Dim oBanco As DInstFinanc
    Set oBanco = New DInstFinanc
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    txtBancos.lbUltimaInstancia = False
    txtBancos.psRaiz = "INSTITUCIONES FINANCIERAS PARA LIBERAR CONTINGENCIAS"
    txtBancos.rs = oOpe.GetOpeObj(psOpeCod, "2") 'oBanco.RecuperaBancos(1)
    txtBancos.Enabled = True
    Me.Show 1
End Sub

Private Sub txtBancos_EmiteDatos()
    If txtBancos = "" Then Exit Sub
        Set oCtaIf = New NCajaCtaIF
        lblBanco.Caption = oCtaIf.NombreIF(Mid(txtBancos.Text, 4, 13))
        lblCta.Caption = oCtaIf.EmiteTipoCuentaIF(Mid(txtBancos.Text, 18, Len(txtBancos.Text))) & " " & txtBancos.psDescripcion
        sSubCtaContBC = Mid(txtBancos.Text, 2, 1) & oCtaIf.SubCuentaIF(Mid(txtBancos.Text, 4, 13))
        sSubCtaContCta = Mid(txtBancos.Text, 18, 2)
End Sub
