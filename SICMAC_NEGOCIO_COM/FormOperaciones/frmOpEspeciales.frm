VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpEspeciales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones Especiales"
   ClientHeight    =   4305
   ClientLeft      =   3000
   ClientTop       =   2520
   ClientWidth     =   7305
   Icon            =   "frmOpEspeciales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6060
      TabIndex        =   10
      Top             =   3855
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3855
      Width           =   1095
   End
   Begin VB.Frame fraDetalle 
      Appearance      =   0  'Flat
      Caption         =   "Detalle"
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
      Height          =   3705
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtCredMIVIVIENDA 
         Enabled         =   0   'False
         Height          =   350
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   39
         Top             =   2880
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame fratipoHip 
         Caption         =   "Tipo Hipotecas y Prendas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   660
         Left            =   360
         TabIndex        =   36
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
         Begin VB.OptionButton optElaboracion 
            Caption         =   "Elaboracion"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   38
            Top             =   300
            Value           =   -1  'True
            Width           =   1350
         End
         Begin VB.OptionButton optLevantamiento 
            Caption         =   "Levantamiento"
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   37
            Top             =   300
            Visible         =   0   'False
            Width           =   1470
         End
      End
      Begin VB.TextBox txtNumPag 
         Height          =   330
         Left            =   3240
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Frame frabusqueda 
         Caption         =   "Tiempo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   660
         Left            =   240
         TabIndex        =   27
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
         Begin VB.OptionButton optmenor 
            Caption         =   "Menor a 1 año"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   29
            Top             =   300
            Value           =   -1  'True
            Width           =   1470
         End
         Begin VB.OptionButton optmayor 
            Caption         =   "Mayor a  un año"
            Height          =   195
            Index           =   0
            Left            =   1560
            TabIndex        =   28
            Top             =   300
            Width           =   1470
         End
      End
      Begin VB.ComboBox cbotipoopc 
         Height          =   315
         ItemData        =   "frmOpEspeciales.frx":030A
         Left            =   3240
         List            =   "frmOpEspeciales.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   3810
      End
      Begin VB.TextBox txtIdAut 
         Height          =   330
         Left            =   5340
         TabIndex        =   19
         Top             =   435
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   465
         Width           =   2850
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   350
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   3
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtGlosa 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   105
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1590
         Width           =   6915
      End
      Begin SICMACT.TxtBuscar txtPers 
         Height          =   315
         Left            =   105
         TabIndex        =   1
         Top             =   1035
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         Appearance      =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   345
         Left            =   4485
         TabIndex        =   4
         Top             =   2505
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCredMIVIVIENDA 
         Caption         =   "Crédito :"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   200
         TabIndex        =   40
         Top             =   2950
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblPag 
         Caption         =   "Paginas :"
         Height          =   210
         Left            =   3240
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo :"
         Height          =   210
         Left            =   3960
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   3795
         TabIndex        =   24
         Top             =   2985
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   3765
         TabIndex        =   23
         Top             =   3375
         Width           =   450
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   300
         Left            =   4485
         TabIndex        =   22
         Top             =   2955
         Width           =   1755
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   300
         Left            =   4485
         TabIndex        =   21
         Top             =   3330
         Width           =   1755
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Id Autorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4005
         TabIndex        =   20
         Top             =   495
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblSimbolo 
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6345
         TabIndex        =   18
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblMoneda 
         Caption         =   "Moneda :"
         Height          =   210
         Left            =   165
         TabIndex        =   17
         Top             =   255
         Width           =   885
      End
      Begin VB.Label lblPers 
         Caption         =   "Persona"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   825
         Width           =   915
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto :"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   3750
         TabIndex        =   15
         Top             =   2595
         Width           =   735
      End
      Begin VB.Label lblDoc 
         Caption         =   "Nro.Doc :"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   195
         TabIndex        =   14
         Top             =   2580
         Width           =   735
      End
      Begin VB.Label lblGlosa 
         Caption         =   "Glosa"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label lblPersona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   1050
         Width           =   5340
      End
   End
   Begin VB.Frame fraRemitente 
      Appearance      =   0  'Flat
      Caption         =   "Datos del Remitente"
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
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   3720
      Width           =   7095
      Begin VB.TextBox txtRemitCiudad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtRemitNombre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   5655
      End
      Begin MSMask.MaskEdBox txtFecEnvio 
         Height          =   300
         Left            =   5640
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Remitente :"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "País o Ciudad :"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Envío :"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   4440
         TabIndex        =   33
         Top             =   600
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmOpEspeciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsOpeCod As COMDConstantes.CaptacOperacion
Dim lsCaption As String

'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String

Dim bOpeAfecta As Boolean
Dim ela, lev, Men, May, dup, his, com, comx, comS, comD As String 'madm 20091211  comv
Dim eSunarp As String 'venb, venl 'madm 20102402 ' MADM 20101112
Dim nTC As Double ' madm 2001224
Dim nRedondeoITF As Double ' BRGO 20110914
'FRHU 20140505 ERS063-2014
Dim sMovNroAut As String
Dim nMovVistoElec As Long
Dim nMovAutoriza As Long
'FIN FRHU 20140505 ERS063-2014
Public Sub Ini(psOpeCod As COMDConstantes.CaptacOperacion, psCaption As String)

'/*Verificar cantidad de operaciones disponibles ANDE 20171218*/
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim bProsigue As Boolean
    Dim cMsgValid As String
    bProsigue = oCaptaLN.OperacionPermitida(gsCodUser, gdFecSis, psOpeCod, cMsgValid)
    If bProsigue = False Then
        MsgBox cMsgValid, vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
'/*end ande*/

    Dim oConst As COMDConstSistema.FCOMITF
    Set oConst = New COMDConstSistema.FCOMITF
        lsOpeCod = psOpeCod
         bOpeAfecta = oConst.VerifOpeVariasAfectaITF(Str(lsOpeCod))
    Set oConst = Nothing
    lsCaption = Mid(psCaption, 3, Len(psCaption) - 2)
    'FRHU 20140625 ERS063-2014 - OBSERVACION
    If gOtrOpeAhoOtrosEgresos = lsOpeCod Then
        'FRHU 20140501 ERS063-2014
        sMovNroAut = ""
        Dim loVistoElectronico As New frmVistoElectronico
        Dim lbVistoVal As Boolean
        lbVistoVal = loVistoElectronico.inicio(11, lsOpeCod)
                        
        If Not lbVistoVal Then
            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones.", vbInformation, "Mensaje del Sistema"
            Exit Sub
        End If
                            
        Call loVistoElectronico.RegistraVistoElectronico(0, nMovVistoElec)
        'FIN FRHU 20140501 ERS063-2014
    End If
    'FIN FRHU 20140625
    
    'WIOR 20160108 ***
    lblCredMIVIVIENDA.Visible = False
    txtCredMIVIVIENDA.Visible = False
    cboMoneda.Enabled = True
    txtCredMIVIVIENDA.Text = ""
    If lsOpeCod = "300443" Or lsOpeCod = "300543" Then
        lblCredMIVIVIENDA.Visible = True
        txtCredMIVIVIENDA.Visible = True
        cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, 1) 'solo para pruebas
        cboMoneda.Enabled = False
    End If
    'WIOR FIN ********
    Me.Show 1
End Sub

Private Sub cboMoneda_Click()
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        
    If Left(gOtrOpeAhoOtrosEgresos, 4) = Left(lsOpeCod, 4) Then
        txtMonto.psTipOpe False
    Else
        txtMonto.psTipOpe True
    End If
    
    If Right(cboMoneda.Text, 3) = COMDConstantes.gMonedaNacional Then
        Me.txtMonto.psSoles True
        Me.lblSimbolo.Caption = "S/."
        'madm 29122009
        If lsOpeCod = "300409" Then
           If comS <> "" Then
           'Me.txtMonto.value = IIf(comS = "", Format(0, "#0.00"), Format(comS, "#0.00"))
            Me.txtMonto.value = Format(comS, "#0.00")
           End If
        End If
       
    Else
        Me.txtMonto.psSoles False
        Me.lblSimbolo.Caption = "$."
        'madm 29122009
        If lsOpeCod = "300409" Then
           If comD <> "" Then
            'Me.txtMonto.value = IIf(comD = "", Format(0, "#0.00"), Format(comD, "#0.00"))
             Me.txtMonto.value = Format(comD, "#0.00")
           End If
        End If
    End If

    Me.lblITF.BackColor = txtMonto.BackColor
    Me.lblTotal.BackColor = txtMonto.BackColor

    If Left(gOtrOpeDuplicadoTarjeta, 4) = Left(lsOpeCod, 4) Then
        txtMonto.Text = Format(clsCapMov.GetTarifa(lsOpeCod, Right(Me.cboMoneda.Text, 3)), "#0.00")
    End If
    Set clsCapMov = Nothing

End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cbotipoopc.Visible = True Then
            cbotipoopc.SetFocus
        Else
            txtPers.SetFocus
        End If
    End If
End Sub


'madm 20091224 ------------------------------
Private Sub cbotipoopc_Click()
'Dim clsTC As COMDConstSistema.NCOMTipoCambio
'Set clsTC = New COMDConstSistema.NCOMTipoCambio
'nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'Set clsTC = Nothing
'modificado x MADM 20100225
        If lsOpeCod = "300406" Or lsOpeCod = "300410" Then
            If Right(cbotipoopc.Text, 2) = "-1" Then
                Me.txtMonto.value = Format(0, "#0.00")
            ElseIf Right(cbotipoopc.Text, 4) = "3122" Or Right(cbotipoopc.Text, 4) = "3125" Then
                Me.txtMonto.value = Format(Men, "#0.00")
            ElseIf Right(cbotipoopc.Text, 4) = "3123" Or Right(cbotipoopc.Text, 4) = "3130" Then
                Me.txtMonto.value = Format(May, "#0.00")
            ElseIf Right(cbotipoopc.Text, 4) = "3126" Or Right(cbotipoopc.Text, 4) = "3131" Then
                Me.txtMonto.value = Format(dup, "#0.00")
            ElseIf his <> "" Then
                Me.txtMonto.value = Format(his, "#0.00")
            End If
                lblTotal.Caption = Format(0, "#0.00")
                Me.lblITF.Caption = Format(0, "#0.00")
        End If
'MADM vuelto a comentar 03032010
'''        If lsOpeCod = 300403 Then
'''            If Right(cbotipoopc.Text, 2) = "-1" Then
'''                Me.txtMonto.value = Format(0, "#0.00")
'''            ElseIf Right(cbotipoopc.Text, 4) = "3124" Then
'''                Me.txtMonto.value = Format(venb, "#0.00")
'''            Else
'''                Me.txtMonto.value = Format(venl, "#0.00")
'''            End If
'''            lbltotal.Caption = Format(0, "#0.00")
'''            Me.LblItf.Caption = Format(0, "#0.00")
'''        End If
    
End Sub

Private Sub cbotipoopc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
       txtPers.SetFocus
    End If
End Sub

'--------------------------------------------
Private Sub cmdCancelar_Click()
    Me.txtPers.Text = ""
    Me.txtGlosa.Text = ""
    Me.txtNroDoc.Text = ""
    Me.lblPersona.Caption = ""
    
    If Not Left(gOtrOpeDuplicadoTarjeta, 4) = Left(lsOpeCod, 4) Then
      Me.txtMonto.value = Format(0, "#0.00")
      lblTotal.Caption = Format(0, "#0.00")
    End If
    
    If Me.cboMoneda.Enabled And cboMoneda.Visible Then
        cboMoneda.SetFocus
    End If
    Me.txtMonto.value = Format(0, "#0.00")
    lblTotal.Caption = Format(0, "#0.00")
    Me.lblITF.Caption = Format(0, "#0.00")
    'madm 20091211 ---------------------------------------------------
    If lsOpeCod = "300418" Then
        If optLevantamiento.item(1).value = True Then
           optLevantamiento.item(1).value = False
           optElaboracion.item(0).value = True
           Me.txtMonto.value = Format(ela, "#0.00")
           
        Else
            Me.txtMonto.value = Format(ela, "#0.00")
        End If
        Me.lblITF.Caption = Format(0, "#0.00")
    End If
    
    'madm 2001228 ------------------------------------
    If lsOpeCod = "300411" Then
        Me.txtMonto.value = com
        Me.lblITF.Caption = Format(0, "#0.00")
    End If
 
 'madm 2001228 ------------------------------------
    'If lsOpeCod = "300450" Then
    If lsOpeCod = gComiOtrServBusqRegSUNARP Then 'JUEZ 20150928
        Me.txtMonto.value = eSunarp
        Me.txtNumPag.Text = 1
        Me.lblITF.Caption = Format(0, "#0.00")
    End If
'end madm
'comentado x MADM 20102024
'    If lsOpeCod = 300403 Then
'        Me.txtMonto.value = comv
'        Me.lblITF.Caption = Format(0, "#0.00")
'    End If
      
    'modificado x MADM 20100224 vuelto a comentar
'''    If (lsOpeCod = 300403) Then
'''        If Right(cbotipoopc.Text, 2) = "-1" Then
'''            Me.txtMonto.value = Format(0, "#0.00")
'''        ElseIf Right(cbotipoopc.Text, 4) = "3124" Then
'''            Me.txtMonto.value = Format(venb, "#0.00")
'''        Else
'''            Me.txtMonto.value = Format(venl, "#0.00")
'''        End If
'''    End If
      
    If lsOpeCod = "300409" Then
        If Right(cboMoneda.Text, 3) = COMDConstantes.gMonedaNacional Then
            Me.txtMonto.value = Format(comS, "#0.00")
            Me.lblITF.Caption = Format(0, "#0.00")
        Else
            Me.txtMonto.value = Format(comD, "#0.00")
            Me.lblITF.Caption = Format(0, "#0.00")
        End If
    End If
    
    '-------------------------------------------------
    'madm 2001224
    If lsOpeCod = "300406" Or lsOpeCod = "300410" Then
        If Right(cbotipoopc.Text, 2) = "-1" Then
            Me.txtMonto.value = Format(0, "#0.00")
        ElseIf Right(cbotipoopc.Text, 4) = "3122" Or Right(cbotipoopc.Text, 4) = "3125" Then
            Me.txtMonto.value = Format(Men, "#0.00")
        ElseIf Right(cbotipoopc.Text, 4) = "3123" Or Right(cbotipoopc.Text, 4) = "3130" Then
            Me.txtMonto.value = Format(May, "#0.00")
        ElseIf Right(cbotipoopc.Text, 4) = "3126" Or Right(cbotipoopc.Text, 4) = "3131" Then
            Me.txtMonto.value = Format(dup, "#0.00")
        ElseIf his <> "" Then
            Me.txtMonto.value = Format(his, "#0.00")
        End If
        Me.lblITF.Caption = 0
    End If
    '-----------------------------------------------------------------
    '*** BRGO 20110309 *************************************************
    If lsOpeCod = "300517" Then
        Me.txtFecEnvio.Text = "__/__/____"
        Me.txtRemitCiudad.Text = ""
        Me.txtRemitNombre.Text = ""
        Me.cmdGrabar.Enabled = False
    End If
    '*******************************************************************
    nRedondeoITF = 0
    sMovNroAut = "" 'FRHU 20140505 ERS063-2014
    
    'WIOR 20160108 ***
    If lsOpeCod = "300443" Or lsOpeCod = "300543" Then
       txtCredMIVIVIENDA.Text = ""
    End If
    'WIOR FIN ********
End Sub

Private Sub cmdGrabar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII

    
    Dim CodOpe As String
    Dim lnMonto As Currency
    Dim Moneda As String
    Dim lsMov As String
    Dim lsMovITF As String
    Dim lsDocumento As String
        
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    
    Dim clsCapM As COMDCaptaGenerales.DCOMCaptaMovimiento
    Set clsCapM = New COMDCaptaGenerales.DCOMCaptaMovimiento
    
    Dim ClsMov As COMDMov.DCOMMov
    Set ClsMov = New COMDMov.DCOMMov
    
    On Error GoTo Error
    
    lnMonto = txtMonto.value
    lsMov = FechaHora(gdFecSis)
    lsDocumento = Me.txtNroDoc.Text
    Dim lnMovNro As Long
    Dim lnMovNroITF As Long
    Dim lbBan As Boolean
    
    'madm 20091215 --------------------------------------
    'If (lsOpeCod = "300450" Or lsOpeCod = "300418" Or lsOpeCod = "300406" Or lsOpeCod = "300411" Or lsOpeCod = "300403" Or lsOpeCod = "300410") And (Right(cboMoneda.Text, 3) = COMDConstantes.gMonedaExtranjera) Then
    If (lsOpeCod = gComiOtrServBusqRegSUNARP Or lsOpeCod = "300418" Or lsOpeCod = "300406" Or lsOpeCod = "300411" Or lsOpeCod = "300403" Or lsOpeCod = "300410") And (Right(cboMoneda.Text, 3) = COMDConstantes.gMonedaExtranjera) Then
            MsgBox "La Operación no es permita en Dólares", vbInformation, "Aviso"
            cboMoneda.SetFocus
            Exit Sub
    End If
     
    'madm 20101112 --------------------------------------
    'If lsOpeCod = "300450" Then
    If lsOpeCod = gComiOtrServBusqRegSUNARP Then 'JUEZ 20150928
            
            If (Me.txtMonto.Text <> eSunarp * CInt(txtNumPag)) Then
                MsgBox "La Monto de la Operación no es Correcta", vbInformation, "Aviso"
                Exit Sub
            End If
    End If
    
    If lsOpeCod = "300406" Or lsOpeCod = "300410" Then
            If cbotipoopc.ListIndex = -1 Then
                    MsgBox "Falta Seleccionar Tipo de Comisión", vbInformation, "Aviso"
                    cbotipoopc.SetFocus
                    Exit Sub
            Else
                    lsCaption = Left(Me.cbotipoopc.Text, 36)
            End If
    End If
    
'''   'MADM 20100224 vuelto a comentar
'''   If lsOpeCod = 300403 Then
'''            If cbotipoopc.ListIndex = -1 Then
'''                    MsgBox "Falta Seleccionar Valor de Venta", vbInformation, "Aviso"
'''                    cbotipoopc.SetFocus
'''                    Exit Sub
'''             Else
'''                    lsCaption = Left(Me.cbotipoopc.Text, 15)
'''             End If
'''    End If
'''   'END MADM

    If Me.lblTotal.Caption = 0 Then
        MsgBox "Cantidad Total incorrecta", vbInformation, "Aviso"
        Exit Sub
    End If
    '----------------------------------------------------
    
    If Len(Me.lblPersona.Caption) = 0 Then
        MsgBox "Ingrese un Nombre", vbInformation, "Aviso"
        txtPers.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ingrese la glosa o comentario correspondiente", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
    
    '******** BRGO 20110308 *********************************************
    If lsOpeCod = "300517" Then
        If Len(Trim(Me.txtRemitNombre.Text)) = 0 Then
            MsgBox "Debe ingresar el nombre del remitente", vbInformation, "Aviso"
            txtRemitNombre.SetFocus
            Exit Sub
        End If
        If Len(Trim(Me.txtRemitCiudad.Text)) = 0 Then
            MsgBox "Debe ingresar la ciudad y/o país del remitente", vbInformation, "Aviso"
            txtRemitCiudad.SetFocus
            Exit Sub
        End If
        If IsDate(Me.txtFecEnvio.Text) = False Then
            MsgBox "Debe ingresar una fecha correcta", vbInformation, "Aviso"
            txtFecEnvio.SetFocus
            Exit Sub
        End If
        If DateDiff("d", CDate(Me.txtFecEnvio.Text), gdFecSis) < 0 Then
            MsgBox "La fecha de envío no puede ser mayor a la fecha del sistema", vbInformation, "Aviso"
            txtFecEnvio.SetFocus
            Exit Sub
        End If
    End If
    '** End BRGO *********************************************************
   
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    'FRHU 20140626 ERS063-2014 - OBSERVACION
    If gOtrOpeAhoOtrosEgresos = lsOpeCod Then
    '*** FRHU 20140505 ERS063-2014
    Dim bRechazado As Boolean
    bRechazado = False
    If VerificarAutorizacion(bRechazado) = False Then
        If bRechazado = True Then
            Unload Me
        End If
        Exit Sub
    End If
    '*** FRHU 20140505 ERS063-2014
    End If
    'FRHU 20140626 ERS063-2014 - OBSERVACION
    
    'WIOR 20160108 ***
    If lsOpeCod = "300443" Or lsOpeCod = "300543" Then
        If Trim(txtCredMIVIVIENDA.Text) = "" Then
            MsgBox "Seleccione el Crédito MIVIVIENDA correspondiente al BBP - FMV", vbInformation, "Aviso"
            txtPers.SetFocus
            Exit Sub
        End If
        
        lsDocumento = IIf(Trim(lsDocumento) = "", Trim(txtCredMIVIVIENDA.Text), Trim(txtCredMIVIVIENDA.Text) & "." & Trim(lsDocumento))
    End If
    'WIOR FIN ********
    
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    
    If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
        'By Capi 28022008
        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
        Dim loLavDinero As frmMovLavDinero
        Dim sPersLavDinero As String
        Dim nMontoLavDinero As Double
        Dim lnMoneda As String
        Dim nTC As Double
        
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set loLavDinero = New frmMovLavDinero
        
        lnMoneda = CInt(Right(Me.cboMoneda.Text, 3))
        
        If lsOpeCod = "300517" Then
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            If lnMoneda = gMonedaNacional Then
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If lnMonto >= Round(nMontoLavDinero * nTC, 2) Then
                'By Capi 1402208
                Call IniciaLavDinero(loLavDinero)
                'ALPA 20081009********************************************************
                'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, CDbl(lnMonto), "109" & gsCodAge & "XXX" & CStr(lnMoneda), Mid(Me.Caption, 15), True, "", , , , , lnMoneda, True)
                sPersLavDinero = loLavDinero.Inicia(, , , , False, True, CDbl(lnMonto), "109" & gsCodAge & "XXX" & CStr(lnMoneda), Mid(Me.Caption, 15), True, "", , , , , lnMoneda, True, gnTipoREU, gnMontoAcumulado, gsOrigen)
                '*********************************************************************
                If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            End If
        End If
        
             
        'lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lsOpeCod, lnMonto, lsDocumento, Me.txtGlosa.Text, Right(Me.cboMoneda.Text, 3), Me.txtPers.Text)
        'ALPA 20081009*********************************************************************************
        'lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lsOpeCod, lnMonto, lsDocumento, Me.txtGlosa.Text, Right(Me.cboMoneda.Text, 3), Me.txtPers.Text, , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero)
        'FRHU 20140505 ERS063-2014
        'lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lsOpeCod, lnMonto, lsDocumento, Me.txtGlosa.Text, Right(Me.cboMoneda.Text, 3), Me.txtPers.Text, , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
        'FRHU 20140626 ERS063-2014 - OBSERVACION
        If gOtrOpeAhoOtrosEgresos = lsOpeCod Then
            lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lsOpeCod, lnMonto, lsDocumento, Me.txtGlosa.Text, Right(Me.cboMoneda.Text, 3), Me.txtPers.Text, nMovVistoElec, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
        Else
            lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lsOpeCod, lnMonto, lsDocumento, Me.txtGlosa.Text, Right(Me.cboMoneda.Text, 3), Me.txtPers.Text, , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
        End If
        'FRHU 20140626 ERS063-2014 - OBSERVACION
        'FIN FRHU 20140505 ERS063-2014
        'ALPA20130930********************************
        If gnMovNro = 0 Then
            MsgBox "La operación no se realizó, favor intentar nuevamente", vbInformation, "Aviso"
            Exit Sub
        End If
        '*** BRGO 20110309 ******************************************************************************************************
        If lsOpeCod = "300517" Then
            Call clsCapM.AgregaMovWesterUnionRemitente(gnMovNro, Me.txtRemitNombre.Text, Me.txtRemitCiudad.Text, Me.txtFecEnvio.Text)
        End If
        '************************************************************************************************************************
        
        '*********************************************
        
        'ALPA 20081010****************************************************
        If gnMovNro > 0 Then
            'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
             Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
        End If
        '**********************************************************************************************
        '***********************************************************************
        Set loLavDinero = Nothing
        If gbITFAplica And CCur(Me.lblITF.Caption) <> 0 Then
            lsMovITF = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser, lsMov)
            lnMovNroITF = clsCapMov.OtrasOperaciones(lsMovITF, COMDConstantes.gITFCobroEfectivo, Abs(Me.lblITF.Caption), lsDocumento, Me.txtGlosa.Text, Right(Me.cboMoneda.Text, 3), Me.txtPers.Text)
            '***BRGO 20110914 Redondeo ITF ****************************
            Call ClsMov.InsertaMovRedondeoITF(lsMovITF, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption))
        End If
        
        Set clsCont = Nothing
        Set clsCapMov = Nothing
        Set ClsMov = Nothing
        
        Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
        Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
            lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", Left(lsCaption, 36), "", Str(lnMonto), lblPersona.Caption, "________" & Trim(Right(Me.cboMoneda.Text, 3)), lsDocumento, 0, "0", IIf(Len(lsDocumento) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, lblITF)
        Set oBol = Nothing
        
        Dim oBITF As COMNCaptaGenerales.NCOMCaptaMovimiento
'        Set oBITF = New COMNCaptaGenerales.NCOMCaptaMovimiento
'            If gbITFAplica And CCur(Me.lblITF.Caption) > 0 Then
'                lsBoletaITF = oBITF.fgITFImprimeBoleta(Me.lblPersona.Caption, CCur(Me.lblITF.Caption), Me.Caption, lnMovNroITF, sLpt, , , , , , , False, , , , 0, 0, , "")
'            End If
'        Set oBITF = Nothing
        Do
           If Trim(lsBoleta) <> "" Then
                lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Print #nFicSal, ""
                Close #nFicSal
          End If
          
          If Trim(lsBoletaITF) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoletaITF
                Print #nFicSal, ""
            Close #nFicSal
          End If
            
        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        Set oBol = Nothing

        cmdCancelar_Click
        'FRHU 20140626 ERS063-2014 - OBSERVACION
        If gOtrOpeAhoOtrosEgresos = lsOpeCod Then
            Unload Me 'FRHU 20140505 ERS063-2014
        End If
        'FIN FRHU 20140626 - OBSERVACION
        
    
    '************ Registrar actividad de opertaciones especiales - ANDE 2017-12-18
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim RVerOpe As ADODB.Recordset
    Dim nEstadoActividad As Integer
    nEstadoActividad = oCaptaLN.RegistrarActividad(gsOpeCod, gsCodUser, gdFecSis)
   
    If nEstadoActividad = 1 Then
        MsgBox "He detectado un problema; su operación no fue afectada, pero por favor comunciar a TI-Desarrollo.", vbError, "Error"
    ElseIf nEstadoActividad = 2 Then
        MsgBox "Ha usado el total de operaciones permitidas para el día de hoy. Si desea realizar más operaciones, comuníquese con el área de Operaciones.", vbInformation + vbOKOnly, "Aviso"
        Unload Me
    End If
    ' END ANDE ******************************************************************
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
    'FIN
           
    End If
  
    Exit Sub
Error:
      MsgBox Str(err.Number) & err.Description
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim nMonto As Double
Dim oPersona As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsPers As New ADODB.Recordset
poLavDinero.TitPersLavDinero = txtPers.psCodigoPersona
poLavDinero.TitPersLavDineroNom = txtPers.psDescripcion
End Sub

Private Sub Form_Activate()
    If txtPers.Text = "" Then
        If Me.cboMoneda.Enabled Then
            Me.cboMoneda.SetFocus
        Else
            '**Modificado por DAOR 20080314 ********************
            If Me.txtPers.Enabled Then
                Me.txtPers.SetFocus
            End If
            'Me.txtPers.SetFocus
            '****************************************************
        End If
    Else
        '**Modificado por DAOR 20080314 ********************
        If txtGlosa.Enabled Then
            txtGlosa.SetFocus
        End If
        'txtGlosa.SetFocus
        '****************************************************
    End If
End Sub

Private Sub Form_Load()
  Dim clsCon As COMDConstantes.DCOMConstantes
  Set clsCon = New COMDConstantes.DCOMConstantes
  Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
  Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
  Dim lnMoneda As COMDConstantes.Moneda
  'madm 20091211
  Dim loParam As COMDColocPig.DCOMColPCalculos
  Set loParam = New COMDColocPig.DCOMColPCalculos
  Dim clsTC As COMDConstSistema.NCOMTipoCambio
  Set clsTC = New COMDConstSistema.NCOMTipoCambio
  nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
  '----------------------------------------------------------------
  '**DAOR 20080314, se tarjo esta línea del eventoç Activate *******
  Me.Icon = LoadPicture(App.Path & gsRutaIcono)
  '****************************************************************
  
  CargaCombo Me.cboMoneda, clsCon.RecuperaConstantes(gMoneda)
  
  Me.Caption = "Otras Operaciones - " & lsCaption
  
  Call ImpreSensa
  
  'cboMoneda.ListIndex = 1
  cboMoneda.ListIndex = 0 'madm 20090104
  
  If Left(COMDConstantes.gOtrOpeDuplicadoTarjeta, 4) = Left(lsOpeCod, 4) Then
    Me.txtMonto.Text = Format(clsCapMov.GetTarifa(lsOpeCod, Right(Me.cboMoneda.Text, 3)), "#0.00")
    If Val(txtMonto.Text) = 0 Then
        Me.txtMonto.Enabled = True
    Else
        Me.txtMonto.Enabled = False
    End If
  End If
    
     'madm 20091224----------------------------------------------
  If lsOpeCod = "300406" Then
        Dim s As String
        Dim p1, p2, p3, p4 As Integer
        Men = ""
        May = ""
        dup = ""
        his = ""
        p1 = 3122
        p2 = 3123
        p3 = 3126
        p4 = 3127
        Men = loParam.dObtieneColocParametro(p1)
        May = loParam.dObtieneColocParametro(p2)
        dup = loParam.dObtieneColocParametro(p3)
        his = loParam.dObtieneColocParametro(p4)
        
        s = CStr(p1 & "," & p2 & "," & p3 & "," & p4)
        
        Label1.Visible = True
        Call Llenar_Combo_con_Recordset(loParam.dObtieneListaDescripcParametros(s), cbotipoopc)
        cbotipoopc.ListIndex = IndiceListaCombo(cbotipoopc, -1)
        cbotipoopc.Visible = True
        
  End If
  'END MADM---------------------------------------------------
    
    'madm 20091211 ----------------------------
    If lsOpeCod = "300418" Then
        Dim p6, p7 As Integer
        p6 = 3120
        p7 = 3121
        ela = ""
        lev = ""
        Me.fratipoHip.Visible = True
        ela = loParam.dObtieneColocParametro(p6)
        lev = loParam.dObtieneColocParametro(p7)
        txtMonto.value = Format(ela, "#,##0.00")
    End If
    '------------------------------------------
    'madm 20091228 ----------------------------
    If lsOpeCod = "300411" Then
        Dim p8 As Integer
        p8 = 3129
        com = ""
        com = loParam.dObtieneColocParametro(p8)
        txtMonto.value = Format(com, "#,##0.00")
    End If
    
      'madm 20101112 ----------------------------
    'If lsOpeCod = "300450" Then
    If lsOpeCod = gComiOtrServBusqRegSUNARP Then 'JUEZ 20150928
        Dim pSunarp As Integer
        pSunarp = 3119
        eSunarp = ""
        eSunarp = loParam.dObtieneColocParametro(pSunarp)
        txtMonto.value = Format(eSunarp, "#,##0.00")
        lblPag.Visible = True
        txtNumPag.Visible = True
        txtNumPag.Text = "1"
    End If
    '------------------------------------------
    'END MADM
    
     'MODIFICADO 20102401 - vuelto a comentar
'''    If lsOpeCod = 300403 Then
''''        Dim p9 As Integer
''''        p9 = 3124
''''        comv = ""
''''        comv = loParam.dObtieneColocParametro(p9)
''''        txtMonto.value = Format(comv, "#,##0.00")
'''        Dim s3 As String
'''        Dim p313, p323 As Integer
'''        venb = ""
'''        venl = ""
'''        p313 = 3124
'''        p323 = 3133
'''        venb = loParam.dObtieneColocParametro(p313)
'''        venl = loParam.dObtieneColocParametro(p323)
'''
'''        s3 = CStr(p313 & "," & p323)
'''
'''        Label1.Visible = True
'''        Call Llenar_Combo_con_Recordset(loParam.dObtieneListaDescripcParametros(s3), cbotipoopc)
'''        cbotipoopc.ListIndex = IndiceListaCombo(cbotipoopc, -1)
'''        cbotipoopc.Visible = True
'''
'''    End If
'''     'END MODIFICADO
'''
    If lsOpeCod = "300409" Then
        Dim p11 As Integer
        p11 = 3128
        comS = loParam.dObtieneColocParametro(p11)
        comD = loParam.dObtieneColocParametro(p11) / nTC
        txtMonto.value = Format(comS, "#,##0.00")
    End If
    '------------------------------------------
    
     'madm 20100104----------------------------------------------
  If lsOpeCod = "300410" Then
        Dim sv As String
        Dim p21, p22, p23, P24 As Integer
        Men = ""
        May = ""
        dup = ""
        his = ""
        p21 = 3125
        p22 = 3130
        p23 = 3131
        P24 = 3132
        Men = loParam.dObtieneColocParametro(p21)
        May = loParam.dObtieneColocParametro(p22)
        dup = loParam.dObtieneColocParametro(p23)
        his = loParam.dObtieneColocParametro(P24)
        
        sv = CStr(p21 & "," & p22 & "," & p23 & "," & P24)
        
        Label1.Visible = True
        Call Llenar_Combo_con_Recordset(loParam.dObtieneListaDescripcParametros(sv), cbotipoopc)
        cbotipoopc.ListIndex = IndiceListaCombo(cbotipoopc, -1)
        cbotipoopc.Visible = True
        
  End If
  'END MADM---------------------------------------------------
  '*** BRGO 20110308 ******************************************
  If lsOpeCod = "300517" Then
     Me.CmdCancelar.Top = Me.CmdCancelar.Top + 1100
     Me.cmdGrabar.Top = Me.cmdGrabar.Top + 1100
     Me.cmdSalir.Top = Me.cmdSalir.Top + 1100
     Me.Height = Me.Height + 1150
     Me.fraRemitente.Visible = True
  Else
     Me.fraRemitente.Visible = False
  End If
  '**End BRGO *****************************************************
  Set clsTC = Nothing
  Set clsCapMov = Nothing
  Set clsCon = Nothing
End Sub
'madm20091211
Private Sub optElaboracion_Click(Index As Integer)
    txtMonto.value = Format(ela, "#,##0.00")
    lblTotal.Caption = 0
    Me.lblITF.Caption = 0
End Sub
Private Sub optLevantamiento_Click(Index As Integer)
    txtMonto.value = Format(lev, "#,##0.00")
    lblTotal.Caption = 0
    Me.lblITF.Caption = 0
End Sub

'-------------------------
Private Sub txtGlosa_GotFocus()
    txtGlosa.SelStart = 0
    txtGlosa.SelLength = Len(txtGlosa.Text)
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        txtNroDoc.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtIdAut_KeyPress(KeyAscii As Integer)
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim oCap As COMDCaptaGenerales.COMDCaptAutorizacion
   Dim nmoneda As Integer
    nmoneda = CInt(Right(Me.cboMoneda.Text, 3))
    Gtitular = CStr(txtPers.psCodigoPersona)
        
   If KeyAscii = 13 And Trim(txtIdAut.Text) <> "" Then
      Set oCap = New COMDCaptaGenerales.COMDCaptAutorizacion
        Set rs = oCap.SAA(Left(CStr(lsOpeCod), 4) & "00", Vusuario, "", Gtitular, CInt(nmoneda), CLng(txtIdAut.Text))
      Set oCap = Nothing
      If rs.State = 1 Then
         If rs.RecordCount > 0 Then
            txtMonto.Text = rs!nMontoAprobado
         Else
            MsgBox "No Existe este Id de Autorización para esta cuenta." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
            txtIdAut.Text = ""
         End If
       End If
   End If
   
 If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 13 Or KeyAscii = 8) Then
      KeyAscii = 0
 End If
End Sub

Private Sub txtMonto_GotFocus()
'If lsOpeCod = "300450" Then
If lsOpeCod = gComiOtrServBusqRegSUNARP Then 'JUEZ 20150928
    'MADM 20101115
    If eSunarp <> "" Then
        Me.txtMonto.Text = eSunarp * CInt(txtNumPag)
    End If
    'END MADM
End If

With txtMonto
        .SelStart = 0
        .SelLength = Len(.Text)
End With

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Dim oITF As New COMDConstSistema.FCOMITF
       
       If lsOpeCod = "300517" Then '**BRGO 20110308
            txtRemitNombre.SetFocus
       Else
             If txtMonto.value = 0 Then
                 cmdGrabar.Enabled = False
             Else
                 cmdGrabar.Enabled = True
             End If
             
             oITF.fgITFParametros
             If oITF.gbITFAplica And bOpeAfecta Then
                 If Left(lsOpeCod, 4) <> "3005" Then
                     Me.lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                 Else
                     Me.lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(txtMonto.value), "#,##0.00") * -1
                 End If
             End If
             Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + txtMonto.value, "#,##0.00")
             Set oITF = Nothing
             If cmdGrabar.Enabled Then Me.cmdGrabar.SetFocus
       End If
       '*** BRGO 20110908 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
        If nRedondeoITF > 0 Then
            Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
            Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + txtMonto.value, "#,##0.00")
        End If
      '*** END BRGO
    End If
    'madm 20100104 ---------------- vuelto a comentar
    'If lsOpeCod = 300418 Or lsOpeCod = 300411 Or lsOpeCod = 300406 Or lsOpeCod = 300403 Or lsOpeCod = 300409 Or lsOpeCod = 300410 Then
     If lsOpeCod = "300418" Or lsOpeCod = "300411" Or lsOpeCod = "300406" Or lsOpeCod = "300409" Or lsOpeCod = "300410" Then
        If Not (lsOpeCod = "300410" And Right(Me.cbotipoopc.Text, 4) = "3132") Then
            KeyAscii = 0
        End If
    End If
    '------------------------------

End Sub

Private Sub txtNroDoc_GotFocus()
    txtNroDoc.SelStart = 0
    txtNroDoc.SelLength = 50
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    Dim lnMonto, lnITF As Currency
    Dim lnValor As Double
    If KeyAscii = 13 Then
        If lsOpeCod = "300436" And Len(Trim(txtNroDoc)) = 18 And (Mid(txtNroDoc, 6, 3) = "515" Or Mid(txtNroDoc, 6, 3) = "516") Then
            Dim obLeasing As COMNCredito.NCOMLeasing
            Set obLeasing = New COMNCredito.NCOMLeasing
            
            If Trim(obLeasing.ValidaCreditoPersona(txtNroDoc.Text, txtPers.Text)) <> "" Then
                MsgBox obLeasing.ValidaCreditoPersona(txtNroDoc.Text, txtPers.Text) & " " & Me.lblPersona.Caption
                Exit Sub
            End If
            
            If Mid(txtNroDoc, 9, 1) <> Trim(Right(cboMoneda.Text, 3)) Then
                MsgBox "Moneda de Operacion no coincide con moneda seleccionada ", vbCritical
                Exit Sub
            End If
            
            Dim oITF As COMDConstSistema.FCOMITF
            Set oITF = New COMDConstSistema.FCOMITF
            
            lnMonto = obLeasing.ObtenerComisionLeasingOtrasOperaciones(txtNroDoc)
            Set obLeasing = Nothing
    
            oITF.fgITFParametros
            lnValor = lnMonto * oITF.gnITFPorcent
            lnValor = oITF.CortaDosITF(lnValor)
            lnITF = lnValor
    
            nRedondeoITF = fgDiferenciaRedondeoITF(lnITF)
            If nRedondeoITF > 0 Then
                lnITF = Format(lnITF - nRedondeoITF, "#,##0.00")
            End If
            lnMonto = lnMonto + lnITF
            cmdGrabar.Enabled = True
        
            'Juez 20120809
            txtMonto.Text = lnMonto
            lblTotal.Caption = lnMonto
        End If
       If txtMonto.Enabled = True Then
          txtMonto.SetFocus
       Else
          If cmdGrabar.Enabled Then cmdGrabar.SetFocus
       End If
       '**Se colocó dentro del if del Leasing
       'txtMonto.Text = lnMonto
       'lblTotal.Caption = lnMonto
    End If
End Sub

'MADM 20101112
Private Sub txtNumPag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtMonto.Text = Format(eSunarp * CInt(txtNumPag), "#,###,##0.00")
    
   If txtMonto.Enabled = True Then
          txtMonto.SetFocus
   Else
          txtPers.SetFocus
   End If
End If

End Sub
'END MADM

Private Sub txtPers_EmiteDatos()
    Me.lblPersona.Caption = txtPers.psDescripcion
    If Me.lblPersona.Caption <> "" Then
        'WIOR 20160108 ***
        If lsOpeCod = "300443" Or lsOpeCod = "300543" Then
            Dim oDCredito As COMDCredito.DCOMCredito
            Dim oCuentas As COMDPersona.UCOMProdPersona
            
            Dim rsCredito As ADODB.Recordset
            
            Set oDCredito = New COMDCredito.DCOMCredito
            
            Set rsCredito = oDCredito.CreditosMIVIVIENDAPersona(txtPers.psCodigoPersona)
            
            txtCredMIVIVIENDA.Text = ""
            If Not (rsCredito.EOF And rsCredito.BOF) Then
                Set oCuentas = New COMDPersona.UCOMProdPersona
                Set oCuentas = frmProdPersona.inicio(lblPersona.Caption, rsCredito)
                If oCuentas.sCtaCod <> "" Then
                    txtCredMIVIVIENDA.Text = Mid(oCuentas.sCtaCod, 1, 18)
                    cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, CInt(Mid(txtCredMIVIVIENDA.Text, 9, 1)))
                End If
            Else
                MsgBox "Persona No cuenta con Créditos MIVIVIENDA", vbInformation, "Aviso"
                txtCredMIVIVIENDA.Text = ""
            End If
        End If
        'WIOR FIN ********
        
        Me.txtGlosa.SetFocus
    End If
End Sub

Private Function Cargousu(ByVal NomUser As String) As String
 Dim rs As New ADODB.Recordset
 Dim oCons As COMDConstSistema.DCOMUAcceso
 Set oCons = New COMDConstSistema.DCOMUAcceso
 
 Set rs = oCons.Cargousu(NomUser)
  If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
  End If
 Set rs = Nothing
 'rs.Close
 Set oCons = Nothing
End Function
'****** BRGO 20110308 ************************************************
Private Sub txtRemitCiudad_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        txtFecEnvio.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtRemitNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        txtRemitCiudad.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
Private Sub txtFecEnvio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Dim oITF As New COMDConstSistema.FCOMITF
            If txtMonto.value = 0 Then
                cmdGrabar.Enabled = False
            Else
                cmdGrabar.Enabled = True
            End If
            
            oITF.fgITFParametros
            If oITF.gbITFAplica And bOpeAfecta Then
                If Left(lsOpeCod, 4) <> "3005" Then
                    Me.lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                Else
                    Me.lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(txtMonto.value), "#,##0.00") * -1
                End If
            End If
            Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + txtMonto.value, "#,##0.00")
       Set oITF = Nothing
       If cmdGrabar.Enabled Then Me.cmdGrabar.SetFocus
    End If
End Sub
'****** END BRGO ************************************************
'******** FRHU 20140505 ERS063-2014
Private Function VerificarAutorizacion(ByRef pbRechazado As Boolean) As Boolean

Dim ocapaut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim oCapAutN  As COMNCaptaGenerales.NCOMCaptAutorizacion
Dim rs As New ADODB.Recordset

Dim lsmensaje As String
Dim nMonto As Double
Dim cmoneda As String
'Dim lbRechazado As Boolean

nMonto = txtMonto.value
cmoneda = Trim(Right(cboMoneda.Text, 3))
   
Set oCapAutN = New COMNCaptaGenerales.NCOMCaptAutorizacion
If sMovNroAut = "" Then 'Si es nueva, registra nueva solicitud
    
    oCapAutN.NuevaSolicitudOtrasOperaciones Trim(Me.txtPers.Text), "1", gdFecSis, nMonto, cmoneda, Trim(txtGlosa.Text), gsCodUser, gOpeAutorizacionOtrosEgresosEfectivo, gsCodAge, sMovNroAut, nMovVistoElec
    
    MsgBox "Solicitud Registrada, comunique a su Admnistrador para la Aprobación..." & Chr$(10) & _
        " No salir de esta operación mientras se realice el proceso..." & Chr$(10) & _
        " Porque sino se procedera a grabar otra Solicitud...", vbInformation, "Aviso"
    VerificarAutorizacion = False
Else
    'Valida el estado de la Solicitud
    If Not oCapAutN.VerificarAutorizacionOtrasOperaciones("1", nMonto, sMovNroAut, lsmensaje, pbRechazado) Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        VerificarAutorizacion = False
    Else
        VerificarAutorizacion = True
    End If
End If
Set oCapAutN = Nothing
End Function
'FIN FRHU 20140505 ERS063-2014
