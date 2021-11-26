VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIGVReversion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion del IGV No considerado como Credito  Fiscal"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmIGVReversion.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   4755
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   8387
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Proceso"
      TabPicture(0)   =   "frmIGVReversion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAceptar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAplicar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Sustento"
      TabPicture(1)   =   "frmIGVReversion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flex"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   4365
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   5340
         TabIndex        =   8
         Top             =   4350
         Width           =   1110
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   4155
         TabIndex        =   7
         Top             =   4350
         Width           =   1110
      End
      Begin Sicmact.FlexEdit flex 
         Height          =   4260
         Left            =   -74910
         TabIndex        =   6
         Top             =   405
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   7514
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Periodo-IngGrabado-IngNoGrabado-IngTotal"
         EncabezadosAnchos=   "500-1100-1500-1500-1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R"
         FormatosEdit    =   "0-0-2-2-2"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Reversion de Impuesto General a las Ventas"
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
         Height          =   3915
         Left            =   105
         TabIndex        =   1
         Top             =   375
         Width           =   6345
         Begin VB.TextBox txtGlosa 
            Appearance      =   0  'Flat
            Height          =   1035
            Left            =   1020
            MaxLength       =   250
            TabIndex        =   3
            Top             =   1155
            Width           =   5220
         End
         Begin MSMask.MaskEdBox mskFecha 
            Height          =   300
            Left            =   1005
            TabIndex        =   2
            Top             =   555
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblTotCredNoUsadG 
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
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   4350
            TabIndex        =   13
            Top             =   3180
            Width           =   1770
         End
         Begin VB.Label lblTotCredUsadG 
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
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   4350
            TabIndex        =   12
            Top             =   2850
            Width           =   1785
         End
         Begin VB.Label lblTotCredNoUsad 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Porcentaje de Cred Fiscal No Usado"
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
            Height          =   240
            Left            =   225
            TabIndex        =   11
            Top             =   3180
            Width           =   5880
         End
         Begin VB.Label lblTotCredUsad 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Porcentaje de Cred Fiscal Usado"
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
            Height          =   240
            Left            =   225
            TabIndex        =   10
            Top             =   2850
            Width           =   5895
         End
         Begin VB.Label lblGlosa 
            Caption         =   "Glosa :"
            Height          =   180
            Left            =   180
            TabIndex        =   5
            Top             =   1185
            Width           =   600
         End
         Begin VB.Label lblFecha 
            Caption         =   "&Fecha"
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   615
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmIGVReversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub cmdAceptar_Click()
    Dim dFecha As Date
    Dim lnCorr As Long
    Dim lsAsiento As String
    Dim lsCtaIGVRev As String
    
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    
    Dim oMov As DMov
    Set oMov = New DMov
    
    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim rsAF As ADODB.Recordset
    Set rsAF = New ADODB.Recordset
    
    Dim lnAcum As Currency
    Dim lnItem As Currency
    
    Dim oAsi As NContImprimir
    Set oAsi = New NContImprimir
    
    Dim lnFactor As Double
    
    lsCtaIGVRev = oConst.LeeConstSistema(gConstSistCtaImpGralVentas)
    
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        Me.SSTab.Tab = 0
        mskFecha.SetFocus
        Exit Sub
    ElseIf Me.txtGlosa.Text = "" Then
        MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
    
    
    
    dFecha = CDate("01/" & Format(Month(DateAdd("m", 1, CDate(Me.mskFecha.Text))), "00") & "/" & Right(Me.mskFecha.Text, 4))
    
    lnFactor = oMov.GetFactorCreditoFiscal(DateAdd("m", -12, dFecha), dFecha)
    
    If oMov.VerfRevCredFiscal(gnAlmacenAsignaIGVNoUsado, DateAdd("d", -1, dFecha)) Then
        MsgBox "Ya se registro la reversion del credito fiscal.", vbInformation, "Aviso"
        Me.mskFecha.SetFocus
        Exit Sub
    End If
    
    If MsgBox("El factor de uso del credito fiscal es :" & Round(lnFactor, 6) & " Desea continuar ? ", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    
    Set rs = oMov.GetMovCredFiscal(DateAdd("m", -1, dFecha), 1 - lnFactor, gnAlmacenAsignaIGVNoUsado)
    Set rsAF = oMov.GetMovCredFiscalAF(DateAdd("m", -1, dFecha), DateAdd("d", -1, dFecha), 1 - lnFactor)
    lnCorr = 0
    lnAcum = 0
    
    While Not rs.EOF
        lnItem = Round(rs.Fields(1), 2)
        lnAcum = lnAcum + lnItem
        rs.MoveNext
    Wend
    
    If lnAcum = 0 Then
        MsgBox "No existe información para el calculo del Crédito Fiscal.", vbInformation, "Aviso"
    End If
    
    oMov.BeginTrans
        lsMovNro = oMov.GeneraMovNro(CDate(Me.mskFecha), gsCodAge, gsCodUser)
        oMov.InsertaMov lsMovNro, gnAlmacenAsignaIGVNoUsado, Me.txtGlosa.Text
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        lnCorr = lnCorr + 1
        oMov.InsertaMovCta lnMovNro, lnCorr, lsCtaIGVRev, lnAcum
    
        While Not rsAF.EOF
            lnCorr = lnCorr + 1
            lnAcum = lnAcum + rsAF.Fields(0)
            oMov.InsertaMovCta lnMovNro, lnCorr, rsAF!Cta, rsAF.Fields(0)
            rsAF.MoveNext
        Wend
    
        lnCorr = lnCorr + 1
        oMov.InsertaMovCta lnMovNro, lnCorr, gcCtaIGV, lnAcum * -1
    
        Call oMov.SetCredFiscalAF(DateAdd("m", -1, dFecha), DateAdd("d", -1, dFecha), 1 - lnFactor)
    oMov.CommitTrans
    
    lsAsiento = oAsi.ImprimeAsientoContable(lsMovNro, 60, 80, , , , , False)
    
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaAsignaIGVCreditoNoFiscal
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Asignacion IGV no Considerado Como Credito Fiscal |Fecha " & mskFecha.Text & " Glosa : " & txtGlosa.Text
            Set objPista = Nothing
            '*******

    oPrevio.Show lsAsiento, Caption, True, , gImpresora

End Sub

Private Sub cmdAplicar_Click()
    Dim dFecha As Date
    
    Dim oMov As DMov
    Set oMov = New DMov
    
    Dim lnFactor As Double
    
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        Me.SSTab.Tab = 0
        mskFecha.SetFocus
        Exit Sub
    End If
    
    dFecha = CDate("01/" & Format(Month(DateAdd("m", 1, CDate(Me.mskFecha.Text))), "00") & "/" & Right(Me.mskFecha.Text, 4))
    
    lnFactor = oMov.GetFactorCreditoFiscal(DateAdd("m", -12, dFecha), dFecha)
    
    Me.lblTotCredUsadG.Caption = Format(lnFactor, "0.000000000")
    Me.lblTotCredNoUsadG.Caption = Format(1 - lnFactor, "0.000000000")
    
    flex.rsFlex = oMov.GetFactorCreditoFiscalCuadro(DateAdd("m", -12, dFecha), dFecha)
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 50
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtGlosa.SetFocus
    End If
End Sub

Private Sub txtGlosa_GotFocus()
    txtGlosa.SelStart = 0
    txtGlosa.SelLength = 300
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
