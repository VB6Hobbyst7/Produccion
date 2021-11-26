VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmTransferencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencia de Activo Fijo"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmTransferencia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEstadistico 
      Caption         =   "Solo estadístico / No genera asiento contable"
      Height          =   435
      Left            =   120
      TabIndex        =   27
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   6660
      TabIndex        =   6
      Top             =   3420
      Width           =   1125
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   7860
      TabIndex        =   5
      Top             =   3420
      Width           =   1125
   End
   Begin VB.Frame fraOpe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Transferencia de Activo Fijo"
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
      Height          =   3330
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8940
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1005
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2340
         Width           =   7845
      End
      Begin Sicmact.TxtBuscar txtAgeO 
         Height          =   345
         Left            =   990
         TabIndex        =   4
         Top             =   855
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Enabled         =   0   'False
         Appearance      =   0
         EnabledText     =   0   'False
      End
      Begin Sicmact.TxtBuscar txtSerie 
         Height          =   315
         Left            =   6825
         TabIndex        =   2
         Top             =   495
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         Appearance      =   0
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
         TipoBusqueda    =   2
      End
      Begin Sicmact.TxtBuscar txtBS 
         Height          =   345
         Left            =   990
         TabIndex        =   1
         Top             =   480
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
         Appearance      =   0
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
      End
      Begin Sicmact.TxtBuscar txtAgeD 
         Height          =   345
         Left            =   1005
         TabIndex        =   11
         Top             =   1935
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         Appearance      =   0
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
         sTitulo         =   ""
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Depre.Ejerc.:"
         Height          =   195
         Left            =   5400
         TabIndex        =   24
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Depre.Ejerc.Ant.:"
         Height          =   195
         Left            =   2520
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor Ini."
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblEjer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6480
         TabIndex        =   21
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label lblEjercAnt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3840
         TabIndex        =   20
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label lblValorIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   18
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie :"
         Height          =   225
         Left            =   6825
         TabIndex        =   17
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblComentario 
         Caption         =   "Coment."
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label lblDestino 
         Caption         =   "Destino"
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
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   1710
         Width           =   2205
      End
      Begin VB.Label lblOrigen 
         Caption         =   "Origen"
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
         Height          =   210
         Left            =   135
         TabIndex        =   14
         Top             =   225
         Width           =   2205
      End
      Begin VB.Label lblAgeDG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   13
         Top             =   1950
         Width           =   6225
      End
      Begin VB.Label lblAgeD 
         Caption         =   "Agencia D:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2010
         Width           =   840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   8940
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Line Line1 
         X1              =   45
         X2              =   8910
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Label lblAgeOG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   9
         Top             =   870
         Width           =   6225
      End
      Begin VB.Label lblAgeO 
         Caption         =   "Agencia O"
         Height          =   180
         Left            =   105
         TabIndex        =   8
         Top             =   930
         Width           =   840
      End
      Begin VB.Label lblBien 
         Caption         =   "Bien :"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   510
         Width           =   840
      End
      Begin VB.Label lblBienG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   3
         Top             =   495
         Width           =   4200
      End
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   315
      Left            =   4095
      TabIndex        =   25
      Top             =   3480
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha"
      Height          =   225
      Left            =   3240
      TabIndex        =   26
      Top             =   3525
      Width           =   810
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1425
   End
End
Attribute VB_Name = "frmTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnMovNroIni As Long
Dim lnAnio As Long
Dim ldFechaGraba As Date '*** PEAC 20130121

Private Sub chkEstadistico_Click()
    If chkEstadistico.value = 0 Then
         mskFecha.Enabled = False
         mskFecha.Text = "__/__/____"
    Else
        mskFecha.Enabled = True
        mskFecha.Text = "__/__/____"
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsMovNro As String
    Dim lnMovNro As Long
    
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir
    Dim RsAF As ADODB.Recordset
    Set RsAF = New ADODB.Recordset
    Dim lnItem  As Integer
    Dim lsCtaDH As String
    Dim lsCtaCont, lsCtaOtro, lsCtaDebeOtro, lsCtaHaberOtro, lsCtaDebeOtro2, lsCtaHaberOtro2 As String
    Dim lcSubCtaOri, lcSubCtaDes As String '*** PEAC 20120725
        
    If Me.txtBS.Text = "" Then
        MsgBox "Debe ingresar un codigo de Bien.", vbInformation, "Aviso"
        Me.txtBS.SetFocus
        Exit Sub
    ElseIf Me.txtSerie.Text = "" Then
        MsgBox "Debe ingresar una serie valida.", vbInformation, "Aviso"
        Me.txtSerie.SetFocus
        Exit Sub
    ElseIf Me.txtAgeD.Text = "" Then
        MsgBox "Debe ingresar el Area  Agencia de destino.", vbInformation, "Aviso"
        Me.txtAgeD.SetFocus
        Exit Sub
    ElseIf Me.txtAgeO.Text = "" Then
        MsgBox "Este bien debe tener una Agencia Origen, sinó consulte a TI para corregirlo.", vbInformation, "Aviso"
        Me.txtAgeD.SetFocus
        Exit Sub
    ElseIf Me.chkEstadistico.value <> 0 And Not IsDate(mskFecha.Text) Then '*** PEAC 20130121
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        Me.mskFecha.SetFocus
        Exit Sub
    End If

    '*** PEAC 20130121
    If Me.chkEstadistico.value = 0 Then
        ldFechaGraba = gdFecSis
    Else
        ldFechaGraba = Me.mskFecha.Text
    End If
    '*** FIN PEAC
    
    If MsgBox("Desea grabar la Transferencia del Activo Fijo", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oMov.BeginTrans
    
        lsMovNro = oMov.GeneraMovNro(ldFechaGraba, Right(gsCodAge, 2), gsCodUser)
        
        oMov.InsertaMov lsMovNro, gnTransAF, Me.txtComentario.Text, 10
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        oMov.InsertaMovBSAF lnAnio, lnMovNroIni, 1, Me.txtBS.Text, Me.txtSerie.Text, lnMovNro
        
        If Me.txtAgeO.Text <> "" Then oMov.InsertaMovObj lnMovNro, 1, 1, ObjCMACAgenciaArea
        oMov.InsertaMovObj lnMovNro, 2, 1, ObjCMACAgenciaArea
        If Me.txtAgeO.Text <> "" Then oMov.InsertaMovObjAgenciaArea lnMovNro, 1, 1, IIf(Len(Trim(Me.txtAgeO.Text)) > 3, Mid(Me.txtAgeO.Text, 4, 2), "01"), Left(Me.txtAgeO.Text, 3)
        oMov.InsertaMovObjAgenciaArea lnMovNro, 2, 1, IIf(Len(Trim(Me.txtAgeD.Text)) > 3, Mid(Me.txtAgeD.Text, 4, 2), "01"), Left(Me.txtAgeD.Text, 3)
        oALmacen.AFActualizaAreaAge Left(Me.txtAgeD.Text, 3), IIf(Len(Trim(Me.txtAgeD.Text)) > 3, Mid(Me.txtAgeD.Text, 4, 2), "01"), lnAnio, Me.txtBS.Text, Me.txtSerie.Text
                
        
        If Me.chkEstadistico.value = 0 Then
        
            '*** PEAC 20120424
            lnItem = 1
            lsCtaDH = "H"
    
            '*** PEAC 20120725
            lcSubCtaOri = IIf(Len(Trim(Me.txtAgeO.Text)) > 3, Right(Me.txtAgeO.Text, 2), "01")
            lcSubCtaDes = IIf(Len(Trim(Me.txtAgeD.Text)) > 3, Right(Me.txtAgeD.Text, 2), "01")
    
            Set RsAF = oMov.BuscaCtaPlantillaAF(gnTransAF, Left(Me.txtSerie.Text, 6), lsCtaDH)
    
            lsCtaCont = IIf(lsCtaDH = "H", RsAF!cCtaContCodH, RsAF!cCtaContCodD)
            lsCtaCont = Replace(lsCtaCont, "AG", lcSubCtaOri)
    
            lsCtaOtro = IIf(lsCtaDH = "D", RsAF!cCtaContCodH, RsAF!cCtaContCodD)
            lsCtaOtro = Replace(lsCtaOtro, "AG", lcSubCtaDes)
    
            lsCtaDebeOtro = IIf(IsNull(RsAF!cCtaContCodOtroD), "", RsAF!cCtaContCodOtroD)
            lsCtaDebeOtro = Replace(lsCtaDebeOtro, "AG", lcSubCtaOri)
    
            lsCtaHaberOtro = IIf(IsNull(RsAF!cCtaContCodOtroH), "", RsAF!cCtaContCodOtroH)
            lsCtaHaberOtro = Replace(lsCtaHaberOtro, "AG", lcSubCtaDes)
    
            lsCtaDebeOtro2 = IIf(IsNull(RsAF!cCtaContCodOtro2D), "", RsAF!cCtaContCodOtro2D)
            lsCtaDebeOtro2 = Replace(lsCtaDebeOtro2, "AG", lcSubCtaDes)
    
            lsCtaHaberOtro2 = IIf(IsNull(RsAF!cCtaContCodOtro2H), "", RsAF!cCtaContCodOtro2H)
            lsCtaHaberOtro2 = Replace(lsCtaHaberOtro2, "AG", lcSubCtaOri)
    
            oMov.InsertaMovCta lnMovNro, lnItem, lsCtaCont, Round(CDbl(Me.lblValorIni.Caption) * -1, 2)
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, lsCtaOtro, Round(CDbl(Me.lblValorIni.Caption), 2)
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, lsCtaDebeOtro, Round((CDbl(Me.lblEjercAnt.Caption) + CDbl(Me.lblEjer.Caption)), 2)
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, lsCtaHaberOtro, Round((CDbl(Me.lblEjercAnt.Caption) + CDbl(Me.lblEjer.Caption)) * -1, 2)
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, lsCtaDebeOtro2, Round(CDbl(Me.lblEjer.Caption), 2)
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, lsCtaHaberOtro2, Round(CDbl(Me.lblEjer.Caption) * -1, 2)
            
            '*** FIN PEAC
        End If
        
    oMov.CommitTrans

    '*** PEAC 20130121
    If Me.chkEstadistico.value = 0 Then
        oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80, Caption), Caption, True
    Else
        MsgBox "EL Activo Fijo " & Me.txtBS.Text & "-" & Me.txtSerie.Text & " ha sido transferido a : " & Me.txtAgeO.Text & " " & Me.lblAgeDG.Caption + Chr(10) + _
        "Sin generar Asiento Contable.", vbOKOnly + vbExclamation, "Atención."
    End If
    'MsgBox "EL Activo Fijo " & Me.txtBS.Text & "-" & Me.txtSerie.Text & " ha sido transferido a : " & Me.txtAgeO.Text & " " & Me.lblAgeDG.Caption
    
    'Unload Me
    Limpia
    
    '*** FIN EPAC
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    
    Me.txtAgeO.rs = oArea.GetAgenciasAreas
    Me.txtAgeD.rs = oArea.GetAgenciasAreas
    Me.txtBS.rs = oALmacen.GetAFBienes
    
    ldFechaGraba = gdFecSis
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtAgeD_EmiteDatos()
    Me.lblAgeDG.Caption = txtAgeD.psDescripcion
    
    If txtAgeD.psDescripcion <> "" Then
        Me.txtComentario.SetFocus
    End If
End Sub

Private Sub txtAgeO_EmiteDatos()
    Me.lblAgeOG.Caption = txtAgeO.psDescripcion
End Sub

Private Sub txtBS_EmiteDatos()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    If txtBS.Text <> "" Then
        Me.lblBienG.Caption = txtBS.psDescripcion
        '2=transfencias
        Me.txtSerie.rs = oALmacen.GetAFBSSerie(txtBS.Text, Year(ldFechaGraba), "2")
    End If
    
    Set oALmacen = Nothing
End Sub

Private Sub txtSerie_EmiteDatos()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    
    If txtBS.Text <> "" And txtSerie.Text <> "" Then
        Set rs = oALmacen.GetAFBSDetalle(txtBS.Text, txtSerie.Text)
        lnMovNroIni = rs.Fields(2)
        lnAnio = rs.Fields(3)
        Me.txtAgeO.Text = rs.Fields(0) & rs.Fields(1)
        txtAgeO_EmiteDatos
        
        '*** PEAC 20120425
        Set rs = oInventario.ObtenerReporteAF("", Left(Me.txtBS.Text, 5), "", CDate("01/" & Format(Trim(Month(ldFechaGraba)), "00") & "/" & Format(Trim(Year(ldFechaGraba)), "0000")), "1", "19850101", Format(ldFechaGraba, "yyyymmdd"), txtSerie.Text)
        
        Me.lblValorIni.Caption = IIf(IsNull(rs!nBSValor), "0.00", rs!nBSValor)
        Me.lblEjercAnt.Caption = IIf(IsNull(rs!nDepreDelEjerAnt), "0.00", rs!nDepreDelEjerAnt)
        Me.lblEjer.Caption = IIf(IsNull(rs!nDepreDelEjer), "0.00", rs!nDepreDelEjer)
        
        '*** FIN PEAC
        
        Me.txtAgeD.SetFocus
    End If
    
    Set oALmacen = Nothing
    Set rs = Nothing
End Sub

'*** PEAC 20130122
Private Sub Limpia()
    Me.txtBS.Text = ""
    Me.lblBienG.Caption = ""
    Me.lblAgeOG.Caption = ""
    Me.lblEjercAnt.Caption = ""
    Me.lblEjer.Caption = ""
    Me.lblAgeDG.Caption = ""
    Me.lblValorIni.Caption = ""
    
    Me.txtAgeD.Text = ""
    Me.txtAgeO.Text = ""
    Me.txtSerie.Text = ""
    Me.txtComentario.Text = ""
    Me.mskFecha.Text = "__/__/____"
    Me.chkEstadistico.value = 0
    
End Sub
