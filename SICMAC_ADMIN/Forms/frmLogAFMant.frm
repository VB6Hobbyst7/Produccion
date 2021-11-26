VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogAFMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AGREGAR ACTIVO FIJO"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "frmLogAFMant.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDepHist 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3885
      MaxLength       =   30
      TabIndex        =   30
      Text            =   "0.00"
      Top             =   6000
      Width           =   1305
   End
   Begin VB.TextBox txtDepreAcum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6375
      MaxLength       =   30
      TabIndex        =   29
      Text            =   "0.00"
      Top             =   6000
      Width           =   1305
   End
   Begin VB.TextBox txtMesAcum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6390
      MaxLength       =   30
      TabIndex        =   28
      Text            =   "0.00"
      Top             =   6375
      Width           =   1290
   End
   Begin VB.TextBox txtTasaDepre 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3885
      MaxLength       =   30
      TabIndex        =   27
      Text            =   "0.00"
      Top             =   6390
      Width           =   1305
   End
   Begin VB.CommandButton cmdMigra 
      Caption         =   "Migracion"
      Height          =   315
      Left            =   2580
      TabIndex        =   23
      Top             =   4950
      Width           =   1860
   End
   Begin VB.TextBox txtTabla 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      TabIndex        =   22
      Top             =   4950
      Width           =   2460
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   210
      Left            =   75
      TabIndex        =   21
      Top             =   5355
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6870
      TabIndex        =   9
      Top             =   4935
      Width           =   960
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   5910
      TabIndex        =   8
      Top             =   4935
      Width           =   960
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      Caption         =   "Activo Fijo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4815
      Left            =   45
      TabIndex        =   10
      Top             =   0
      Width           =   7800
      Begin VB.Frame Frame1 
         Caption         =   "Tipo  de Bien"
         Height          =   495
         Left            =   1680
         TabIndex        =   41
         Top             =   2760
         Width           =   3135
         Begin VB.OptionButton optTransf 
            Caption         =   "No Depreciable"
            Height          =   195
            Left            =   1560
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optBaja 
            Caption         =   "Depreciable"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Left            =   3840
         TabIndex        =   40
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtSeriePlaca 
         Height          =   285
         Left            =   3840
         TabIndex        =   39
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtMarca 
         Height          =   285
         Left            =   3840
         TabIndex        =   38
         Top             =   1320
         Width           =   3855
      End
      Begin VB.ComboBox cboTpo 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2385
         Width           =   3180
      End
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   105
         MaxLength       =   300
         TabIndex        =   7
         Top             =   4215
         Width           =   7575
      End
      Begin VB.TextBox txtMesTDeprecia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   2025
         Width           =   1140
      End
      Begin VB.TextBox txtMontoIni 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1650
         Width           =   1125
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5655
         MaxLength       =   30
         TabIndex        =   1
         Top             =   2985
         Width           =   2010
      End
      Begin Sicmact.TxtBuscar txtAF 
         Height          =   315
         Left            =   135
         TabIndex        =   0
         Top             =   255
         Width           =   1530
         _ExtentX        =   2699
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
      End
      Begin MSMask.MaskEdBox mskFecIng 
         Height          =   285
         Left            =   1695
         TabIndex        =   2
         Top             =   1320
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar txtAge 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   825
         Width           =   1530
         _ExtentX        =   2699
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
      End
      Begin Sicmact.TxtBuscar txtPersona 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   3645
         Width           =   1530
         _ExtentX        =   2699
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
         TipoBusqueda    =   3
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie/Placa :"
         Height          =   195
         Left            =   2880
         TabIndex        =   37
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Modelo :"
         Height          =   195
         Left            =   3120
         TabIndex        =   36
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Marca :"
         Height          =   195
         Left            =   3120
         TabIndex        =   35
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblCategoriaBien 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   26
         Top             =   3330
         Width           =   270
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Act. Fijo:"
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   2430
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Comentario"
         Height          =   210
         Left            =   165
         TabIndex        =   20
         Top             =   3990
         Width           =   810
      End
      Begin VB.Label lblPersonaG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1665
         TabIndex        =   19
         Top             =   3675
         Width           =   6000
      End
      Begin VB.Label lblPersona 
         Caption         =   "Persona"
         Height          =   210
         Left            =   135
         TabIndex        =   18
         Top             =   3435
         Width           =   810
      End
      Begin VB.Label lblAgencia 
         Caption         =   "Agencia"
         Height          =   210
         Left            =   165
         TabIndex        =   17
         Top             =   615
         Width           =   810
      End
      Begin VB.Label lblAgeG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1680
         TabIndex        =   16
         Top             =   855
         Width           =   6000
      End
      Begin VB.Label lblMesTDeprecia 
         AutoSize        =   -1  'True
         Caption         =   "Meses Tot depre. :"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lnlFecIng 
         AutoSize        =   -1  'True
         Caption         =   "F. Ingreso :"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1365
         Width           =   795
      End
      Begin VB.Label lblMontoIni 
         AutoSize        =   -1  'True
         Caption         =   "Monto Ini. Anual :"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label lblSerie 
         AutoSize        =   -1  'True
         Caption         =   "Serie :"
         Height          =   195
         Left            =   5085
         TabIndex        =   12
         Top             =   3015
         Width           =   450
      End
      Begin VB.Label lblBSG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1695
         TabIndex        =   11
         Top             =   270
         Width           =   6000
      End
   End
   Begin VB.Label lblDepAcum 
      AutoSize        =   -1  'True
      Caption         =   "Dep.Ajustada :"
      Height          =   195
      Left            =   5280
      TabIndex        =   34
      Top             =   6030
      Width           =   1050
   End
   Begin VB.Label lblMesAcumDep 
      AutoSize        =   -1  'True
      Caption         =   "Mes Acum :"
      Height          =   195
      Left            =   5295
      TabIndex        =   33
      Top             =   6420
      Width           =   840
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "Dep. Hist :"
      Height          =   195
      Left            =   3135
      TabIndex        =   32
      Top             =   6030
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tasa Deprec.:"
      Height          =   195
      Left            =   2880
      TabIndex        =   31
      Top             =   6390
      Width           =   1020
   End
End
Attribute VB_Name = "frmLogAFMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lcBSCod As String
Dim lcSerie As String
Dim ldFecha As Date
Dim lnMIni As Currency
Dim lnMHist As Currency
Dim lnMAjus As Currency
Dim lnVida_Util As Integer
Dim lnPerDep As Integer
Dim lsAreaAge As String
Dim lsPersCod As String
Dim lsComentario As String

Dim lbIngreso As Boolean
Dim lbModifica As Boolean
Dim lbElimina As Boolean

Dim lnOpeTpo As Integer
'*** PEAC 20100511
Dim lnPorcenDepre As Double, lnMontoDepre As Double, lnMontoUIT As Double
Dim lcTextoCategoriaBien As String, lcCodCategoriaBien As String
Dim lsAgencia As String
'*** FIN PEAC


Public Sub Ini(pnOpeTpo As Integer, pbIngreso As Boolean, pbModifica As Boolean, pbElimina As Boolean, _
    pcBSCod As String, pcSerie As String, pdFecha As Date, pnMIni As Currency, pnMHist As Currency, pnMAjus As Currency, _
    pnVida_Util As Integer, pnPerDep As Integer, psAreaAge As String, psPersCod As String, psComentario As String)
    
    lcBSCod = pcBSCod
    lcSerie = pcSerie
    ldFecha = pdFecha
    lnMIni = pnMIni
    lnMHist = pnMHist
    lnMAjus = pnMAjus
    lnVida_Util = pnVida_Util
    lnPerDep = pnPerDep
    lsAreaAge = psAreaAge
    lsPersCod = psPersCod
    lsComentario = psComentario
    
    lbIngreso = pbIngreso
    lbModifica = pbModifica
    lbElimina = pbElimina
    
    lnOpeTpo = pnOpeTpo
    Me.Show 1
End Sub

Private Sub cboTpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'txtSerie.SetFocus
        Me.optBaja.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    Dim oAF As DMov
    Set oAF = New DMov
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim lnMovNroAjus As Long
    Dim lsMovNroAjus As String
    Dim lsCtaCont As String
    Dim lsAgencia As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim lsCtaOpeBS As String
    Dim lsTipo As String
    
    lsTipo = Right(cboTpo.Text, 1)
    If Not Valida Then Exit Sub
    
    If MsgBox("Desea guardar los cambios ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If Mid(Me.txtAge.Text, 4, 2) = "" Then
        lsAgencia = "01"
    Else
        lsAgencia = Mid(Me.txtAge.Text, 4, 2)
    End If

    If True Then
        oAF.BeginTrans
            lsMovNro = oAF.GeneraMovNro(gdFecSis, , gsCodUser) 'NAGL Cambió de Me.mskFecIng.Text a gdFecSis Según RFC1910190001
            oAF.InsertaMov lsMovNro, gnDepAF, "Depre. de ActivoFijo" & txtComentario.Text
            lnMovNro = oAF.GetnMovNro(lsMovNro)
            lsMovNroAjus = oAF.GeneraMovNro(gdFecSis, , gsCodUser, lsMovNro) 'NAGL Cambió de Me.mskFecIng.Text a gdFecSis Según RFC1910190001
            oAF.InsertaMov lsMovNroAjus, gnDepAjusteAF, "Deprr. Ajustada de ActivoFijo" & txtComentario.Text
            lnMovNroAjus = oAF.GetnMovNro(lsMovNroAjus)
            
            If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaOpeBS = oAF.GetCtaDep(Me.txtAF.Text)
            'oAF.InsertaMovBSActivoFijoUnico Year(Me.mskFecIng.Text), lnMovNro, Me.txtAF.Text, Me.txtSerie.Text, Me.txtMontoIni.Text, Me.txtDepreAcum.Text, "0", CDate(Me.mskFecIng.Text), Left(Me.txtAge.Text, 3), lsAgencia, Me.txtMesTDeprecia.Text, Me.txtMesAcum.Text, Me.txtComentario.Text, "1", "0", "1", Me.txtSerie.Text, Me.txtSerie.Text, CDate(Me.mskFecIng.Text), CDate(Me.mskFecIng.Text), Right(Me.cboTpo.Text, 2), "0", txtPersona.Text, lcCodCategoriaBien, Me.txtTasaDepre.Text '*** PEAC 20100511 - SE AGREGO EL PRAMETRO (lcCodCategoriaBien,Me.txtTasaDepre.Text)
            '*** PEAC 20120326
            oAF.InsertaMovBSActivoFijoUnico Year(Me.mskFecIng.Text), lnMovNro, Me.txtAF.Text, Me.txtSerie.Text, Me.txtMontoIni.Text, Me.txtDepreAcum.Text, "0", CDate(Me.mskFecIng.Text), Left(Me.txtAge.Text, 3), lsAgencia, Me.txtMesTDeprecia.Text, Me.txtMesAcum.Text, Me.txtComentario.Text, "1", "0", "1", Me.txtSerie.Text, Me.txtSerie.Text, CDate(Me.mskFecIng.Text), CDate(Me.mskFecIng.Text), Right(Me.cboTpo.Text, 2), "0", txtPersona.Text, lcCodCategoriaBien, Me.txtTasaDepre.Text, fgReemplazaCaracterEspecial(Me.txtMarca.Text), fgReemplazaCaracterEspecial(Me.txtModelo.Text), fgReemplazaCaracterEspecial(Me.txtSeriePlaca.Text)

            'Depre. Historica
             oAF.InsertaMovBSAF Year(gdFecSis), lnMovNro, 1, Me.txtAF.Text, Me.txtSerie.Text, lnMovNro, lsTipo
            lsCtaCont = oAF.GetOpeCtaCtaOtro(gnDepAF, lsCtaOpeBS, "", False)
            If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNro, 1, lsCtaCont, CCur(Me.txtDepHist.Text) * -1
            
            If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaCont = oAF.GetOpeCtaCta(gnDepAF, lsCtaOpeBS, "")
            If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNro, 2, lsCtaCont, CCur(Me.txtDepHist.Text)
            
            'Depre. Ajsutada
            If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaCont = oOpe.EmiteOpeCta(gnDepAF, "D")
            oAF.InsertaMovBSAF Year(gdFecSis), lnMovNro, 1, Me.txtAF.Text, Me.txtSerie.Text, lnMovNroAjus, lsTipo
            If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNroAjus, 1, lsCtaCont, CCur(Me.txtDepreAcum.Text) * -1
            
            lsCtaCont = oAF.GetOpeCtaCta(gnDepAF, lsCtaOpeBS, "")
            lsCtaCont = Left(lsCtaCont, 2) & "6" & Mid(lsCtaCont, 4, 100)
            If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNroAjus, 2, lsCtaCont, CCur(Me.txtDepreAcum.Text)
            
            If Me.txtPersona.Text <> "" Then
                oAF.InsertaMovGasto lnMovNro, Me.txtPersona.Text, ""
                oAF.InsertaMovGasto lnMovNroAjus, Me.txtPersona.Text, ""
            End If
            
        oAF.CommitTrans
    
    ElseIf lbElimina Then
        oAF.EliminaActivoFijo Me.txtAF.Text, Me.txtSerie.Text, Year(gdFecSis)
    End If
    
    Set oAF = Nothing
    Unload Me
End Sub

Private Sub cmdMigra_Click()
    Dim oAF As DMov
    Set oAF = New DMov
        
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim lnMovNroAjus As Long
    Dim lsMovNroAjus As String
    Dim lsCtaCont As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim lsCtaOpeBS As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lnCorr As Long
    Dim ldFecha As Date
    
    oCon.AbreConexion
    
    If Me.txtTabla.Text = "" Then
        MsgBox "Debe ingresar una tabla.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'sql = "Select * from " & Me.txtTabla.Text & " where  VINI is Not null AND VHIS iS nOT nULL"
    sql = "Select * from " & Me.txtTabla.Text & " where  Codigo iS nOT nULL And Ban = " & Right(Me.cboTpo.Text, 6)
    
    If Me.txtComentario = "" Then
        MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
        
        Exit Sub
    End If
    If MsgBox("Desea Migrar la Data de la Tabla " & Me.txtTabla.Text & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set rs = oCon.CargaRecordSet(sql)
    
    Me.prg.Max = (rs.RecordCount * 4) + 10
    Me.prg.value = 0
    
    ldFecha = CDate(Me.mskFecIng.Text)
    
    If True Then
        oAF.BeginTrans
            lsMovNro = oAF.GeneraMovNro(ldFecha, , gsCodUser)
            oAF.InsertaMov lsMovNro, gnDepAF, "Depre. de ActivoFijo - " & Me.cboTpo.Text & " - " & txtComentario.Text, 25
            lnMovNro = oAF.GetnMovNro(lsMovNro)
            lsMovNroAjus = oAF.GeneraMovNro(ldFecha, , gsCodUser, lsMovNro)
            oAF.InsertaMov lsMovNroAjus, gnDepAjusteAF, "Deprr. Ajustada de ActivoFijo - " & Me.cboTpo.Text & " - " & txtComentario.Text, 25
            lnMovNroAjus = oAF.GetnMovNro(lsMovNroAjus)
            
            
'****************************************
            lnCorr = 0
            While Not rs.EOF
                'lsCtaOpeBS = oAF.GetCtaDep(rs.Fields("Cod"), True) 'Para intangible
                lsCtaOpeBS = oAF.GetCtaDep(rs.Fields("Cod"))  'Para Bienes
                lsCtaOpeBS = Replace(lsCtaOpeBS, "AG", Format(rs.Fields("cAgeCod"), "00"))
                
                
                'Depre. Historica
                lnCorr = lnCorr + 1
                Me.txtSerie.Text = Format(rs!CorrID, "00000") & "-" & rs!nItem & "-" & IIf(IsNull(rs!codigo), "07", rs!codigo)
                Me.txtMontoIni.Text = Round(IIf(IsNull(rs!VHIS), 0, rs!VHIS), 2)
                Me.txtDepreAcum.Text = Round(IIf(IsNull(rs!VAJUS), rs!VHIS, rs!VAJUS), 2)
                'Me.mskFecIng.Text = Format(rs!FCompra, "dd/mm/yyyy")
                Me.txtAge.Text = IIf(IsNull(rs!cAreaCod), "026", Format(rs!cAreaCod, "000")) & IIf(IsNull(rs!cAgeCod), "07", Format(rs!cAgeCod, "00"))
                Me.txtMesTDeprecia.Text = Round(rs!VUTIL, 0)
                If rs!MDEP = 0 Or rs!MDEP = rs!VUTIL Then
                    Me.txtMesAcum.Text = Round(rs!MDEP, 0)
                Else
                    Me.txtMesAcum.Text = Round(rs!MDEP, 0)
                End If
                lsComentario = Replace(rs.Fields("DESCRIPCION") & " - " & rs.Fields("Ubicacion"), Chr(39), "")
                
                If Year(rs!FCompra) <> 2007 Then
                    oAF.InsertaMovBSActivoFijo Year(ldFecha), lnMovNro, rs!cod, Me.txtSerie.Text, Me.txtMontoIni.Text, Me.txtDepreAcum.Text, CDate("31/12/2006"), Left(Me.txtAge.Text, 3), Mid(Me.txtAge.Text, 4, 2), Me.txtMesTDeprecia.Text, Me.txtMesAcum.Text, lsComentario, rs!FCompra, CDate(Me.mskFecIng.Text), Right(Me.cboTpo.Text, 8)
                Else
                    oAF.InsertaMovBSActivoFijo Year(ldFecha), lnMovNro, rs!cod, Me.txtSerie.Text, Me.txtMontoIni.Text, Me.txtDepreAcum.Text, rs!FCompra, Left(Me.txtAge.Text, 3), Mid(Me.txtAge.Text, 4, 2), Me.txtMesTDeprecia.Text, Me.txtMesAcum.Text, lsComentario, rs!FCompra, CDate(Me.mskFecIng.Text), Right(Me.cboTpo.Text, 8)
                End If
                oAF.InsertaMovBSAF Year(ldFecha), lnMovNro, lnCorr, rs!cod, Me.txtSerie.Text, lnMovNro
                
                'Me.txtDepHist.Text = Round(rs!DHIS, 2)
                Me.txtDepHist.Text = Round(rs!DHIS, 2)
                lsCtaCont = oAF.GetOpeCtaCtaOtro(gnDepAF, lsCtaOpeBS, "", False, rs!cAgeCod)
                lsCtaCont = Replace(lsCtaCont, "AG", rs!cAgeCod)
                oAF.InsertaMovCta lnMovNro, lnCorr, lsCtaCont, CCur(Me.txtDepHist.Text)
                
                rs.MoveNext
                Me.prg.value = Me.prg.value + 1
            Wend
            
'            rs.MoveFirst
'            While Not rs.EOF
'                'Depre. Historica
'                lnCorr = lnCorr + 1
'
'                Me.txtDepHist.Text = Round(rs!DHIS, 2)
'
'                lsCtaCont = oAF.GetOpeCtaCta(gnDepAF, lsCtaOpeBS, "")
'                oAF.InsertaMovCta lnMovNro, lnCorr, lsCtaCont, CCur(Me.txtDepHist.Text)
'                rs.MoveNext
'                Me.prg.value = Me.prg.value + 1
'            Wend

'**************************************
''            rs.MoveFirst
''            lnCorr = 0
''            lsCtaCont = oOpe.EmiteOpeCta(gnDepAF, "D")
''            While Not rs.EOF
''                'Depre. Ajsutada
''                'Me.txtDepreAcum.Text = Round(rs!DAJUS, 2)
''                Me.txtDepreAcum.Text = Round(rs!DAJUS, 2)
''                lnCorr = lnCorr + 1
''                Me.txtSerie.Text = Format(rs!CorrID, "00000") & "-" & rs!Item & "-" & IIf(IsNull(rs!codigo), "07", rs!codigo)
''                oAF.InsertaMovBSAF Year(ldFecha), lnMovNro, lnCorr, rs!Cod, Me.txtSerie.Text, lnMovNroAjus
''                oAF.InsertaMovCta lnMovNroAjus, lnCorr, lsCtaCont, CCur(Me.txtDepreAcum.Text)
''                rs.MoveNext
''                Me.prg.value = Me.prg.value + 1
''                Caption = prg.value
''            Wend
            
'            While Not rs.EOF
'                'Depre. Ajsutada
'                lnCorr = lnCorr + 1
'
'                Me.txtDepreAcum.Text = Round(rs!DAJUS, 2)
'
'                lsCtaCont = oAF.GetOpeCtaCta(gnDepAF, lsCtaOpeBS, "")
'                lsCtaCont = Left(lsCtaCont, 2) & "6" & Mid(lsCtaCont, 4, 100)
'                oAF.InsertaMovCta lnMovNroAjus, lnCorr, lsCtaCont, CCur(Me.txtDepreAcum.Text)
'                rs.MoveNext
'                Me.prg.value = Me.prg.value + 1
'            Wend
        oAF.CommitTrans
    
    ElseIf lbElimina Then
        oAF.EliminaActivoFijo Me.txtAF.Text, Me.txtSerie.Text, Year(ldFecha)
    End If
    
    Set oAF = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    
    '*** PEAC 20100511
    Dim sSQL As String
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    '*** FIN PEAC
    
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
        
    Me.txtAF.rs = oAlmacen.GetBienesAlmacen(, gnLogBSTpoBienFijo, True)
    Me.txtAge.rs = oArea.GetAgenciasAreas
    
    '*** PEAC 20100511
    oCon.AbreConexion
    sSQL = " exec stp_sel_RecuperaParamDepreAF"
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    '***FIN PEAC
    
    Set oAlmacen = Nothing
    Set oArea = Nothing
    Me.mskFecIng = "01/01/2002"
    
    lnPorcenDepre = rs!nPorcenDepre: lnMontoDepre = rs!nMontoDepre: lnMontoUIT = rs!nMontoUIT '*** PEAC 20100511
    
    If lbModifica Or lbElimina Then
        Me.txtAF.Text = lcBSCod
        txtAF_EmiteDatos
        Me.txtSerie.Text = lcSerie
        Me.mskFecIng.Text = Format(ldFecha, gcFormatoFechaView)
        Me.txtMontoIni.Text = Format(lnMIni, "#,##0.0")
        Me.txtDepHist.Text = lnMHist
        Me.txtDepreAcum.Text = lnMAjus
        Me.txtMesTDeprecia.Text = lnVida_Util
        Me.txtMesAcum.Text = lnPerDep
        Me.txtAge.Text = lsAreaAge
        txtAge_EmiteDatos
        Me.txtPersona.Text = lsPersCod
        txtPersona_EmiteDatos
        Me.txtComentario.Text = lsComentario
        
        Me.txtDepHist.Enabled = False
        Me.txtDepreAcum.Enabled = False
        
    End If
    
    If lbElimina Then
        Me.fra.Enabled = False
    End If
    
    Set rs = oGen.GetConstante(5062, False)
    Me.cboTpo.Clear
    While Not rs.EOF
        cboTpo.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    
    
End Sub

Private Sub mskFecIng_GotFocus()
    Me.mskFecIng.SelStart = 0
    Me.mskFecIng.SelLength = 50
End Sub

Private Sub mskFecIng_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMontoIni.SetFocus
    End If
End Sub

Private Sub optBaja_Click()
    '*** PEAC 20100506
    
'    If MsgBox("¿Es un Bien Depreciable?", vbQuestion + vbYesNo, "Atención") = vbYes Then
        lcTextoCategoriaBien = "BIEN DEPRECIABLE"
        lcCodCategoriaBien = "1"
'    Else
'        lcTextoCategoriaBien = "BIEN NO DEPRECIABLE"
'        lcCodCategoriaBien = "0"
'    End If
    
    Me.lblCategoriaBien.Caption = lcTextoCategoriaBien
    
    If Mid(Me.txtAge.Text, 4, 2) = "" Then
        lsAgencia = "01"
    Else
        lsAgencia = Mid(Me.txtAge.Text, 4, 2)
    End If

    GenerarCodInventario lsAgencia, Me.txtAF
       
    '*** FIN PEAC

End Sub

Private Sub optTransf_Click()
    '*** PEAC 20100506
    
'    If MsgBox("¿Es un Bien Depreciable?", vbQuestion + vbYesNo, "Atención") = vbYes Then
'        lcTextoCategoriaBien = "BIEN DEPRECIABLE"
'        lcCodCategoriaBien = "1"
'    Else
        lcTextoCategoriaBien = "BIEN NO DEPRECIABLE"
        lcCodCategoriaBien = "0"
'    End If
    
    Me.lblCategoriaBien.Caption = lcTextoCategoriaBien
    
    If Mid(Me.txtAge.Text, 4, 2) = "" Then
        lsAgencia = "01"
    Else
        lsAgencia = Mid(Me.txtAge.Text, 4, 2)
    End If

    GenerarCodInventario lsAgencia, Me.txtAF
       
    '*** FIN PEAC

End Sub

Private Sub txtAF_EmiteDatos()
    Me.lblBSG.Caption = txtAF.psDescripcion
End Sub

Private Sub txtAF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAge.SetFocus
    End If
End Sub

Private Sub txtAge_EmiteDatos()
    Me.lblAgeG.Caption = txtAge.psDescripcion
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFecIng.SetFocus
    End If
End Sub

Private Sub txtComentario_GotFocus()
    txtComentario.SelStart = 0
    txtComentario.SelLength = 300
End Sub

Private Sub txtDepHist_GotFocus()
    txtDepHist.SelStart = 0
    txtDepHist.SelLength = 50
End Sub

Private Sub txtDepHist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDepreAcum.SetFocus
    Else
        KeyAscii = NumerosDecimales(Me.txtDepHist, KeyAscii, 15)
    End If
End Sub

Private Sub txtDepreAcum_GotFocus()
    txtDepreAcum.SelStart = 0
    txtDepreAcum.SelLength = 50
End Sub

Private Sub txtDepreAcum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMesTDeprecia.SetFocus
    Else
        KeyAscii = NumerosDecimales(Me.txtDepreAcum, KeyAscii, 15)
    End If
End Sub

Private Sub txtMarca_GotFocus()
    txtMarca.SelStart = 0
    txtMarca.SelLength = 50
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.txtTasaDepre.SetFocus
        Me.txtModelo.SetFocus
    End If

End Sub


Private Sub txtMesAcum_GotFocus()
    txtMesAcum.SelStart = 0
    txtMesAcum.SelLength = 50
End Sub

Private Sub txtMesAcum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cboTpo.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub txtMesTDeprecia_GotFocus()
    txtMesTDeprecia.SelStart = 0
    txtMesTDeprecia.SelLength = Len(txtMesTDeprecia)
End Sub

Private Sub txtMesTDeprecia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.txtTasaDepre.SetFocus
        Me.txtMarca.SetFocus
    Else
        KeyAscii = NumerosDecimales(Me.txtMesTDeprecia, KeyAscii, 15)
    End If
    
End Sub

Private Sub txtModelo_GotFocus()
    txtModelo.SelStart = 0
    txtModelo.SelLength = 50
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.txtTasaDepre.SetFocus
        Me.txtSeriePlaca.SetFocus
    End If

End Sub


Private Sub txtMontoIni_GotFocus()
    txtMontoIni.SelStart = 0
    txtMontoIni.SelLength = 50
End Sub

Private Sub txtMontoIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMesTDeprecia.SetFocus
    Else
        KeyAscii = NumerosDecimales(Me.txtMontoIni, KeyAscii, 15)
    End If
End Sub

Private Sub txtPersona_EmiteDatos()
    Me.lblPersonaG.Caption = Me.txtPersona.psDescripcion
End Sub

Private Sub txtPersona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentario.SetFocus
    End If
End Sub

Private Sub txtSerie_GotFocus()
    txtSerie.SelStart = 0
    txtSerie.SelLength = 50
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.mskFecIng.SetFocus
        Me.txtPersona.SetFocus
    End If
End Sub

Private Function Valida() As Boolean
    If Me.txtAF.Text = "" Then
        MsgBox "Debe ingresar un bien o un intangble.", vbInformation, "Aviso"
        Me.txtAF.SetFocus
        Valida = False
    ElseIf Me.txtSerie.Text = "" Then
        MsgBox "Debe ingresar un numero de serie.", vbInformation, "Aviso"
        Me.txtSerie.SetFocus
        Valida = False
    ElseIf Not IsDate(Me.mskFecIng.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        Me.mskFecIng.SetFocus
        Valida = False
    ElseIf Not IsNumeric(Me.txtMontoIni.Text) Then
        MsgBox "Debe ingresar un monto valido.", vbInformation, "Aviso"
        Me.txtMontoIni.SetFocus
        Valida = False
    ElseIf Not IsNumeric(Me.txtMontoIni.Text) Then
        MsgBox "Debe ingresar un monto valido.", vbInformation, "Aviso"
        Me.txtMontoIni.SetFocus
        Valida = False
    ElseIf Not IsNumeric(Me.txtDepHist.Text) Then
        MsgBox "Debe ingresar un monto valido.", vbInformation, "Aviso"
        Me.txtDepHist.SetFocus
        Valida = False
    ElseIf Not IsNumeric(Me.txtDepreAcum.Text) Then
        MsgBox "Debe ingresar un monto valido.", vbInformation, "Aviso"
        Me.txtDepreAcum.SetFocus
        Valida = False
    ElseIf Not IsNumeric(Me.txtMesTDeprecia.Text) Then
        MsgBox "Debe ingresar un numero valido.", vbInformation, "Aviso"
        Me.txtMesTDeprecia.SetFocus
        Valida = False
    ElseIf CCur(Me.txtMesTDeprecia.Text) <= 0 Then '*** PEAC 20110819
        MsgBox "Debe ingresar el número de meses a depreciar.", vbInformation, "Aviso"
        Me.txtMesTDeprecia.SetFocus
        Valida = False
    ElseIf Not IsNumeric(Me.txtMesAcum.Text) Then
        MsgBox "Debe ingresar un numero valido.", vbInformation, "Aviso"
        Me.txtMesAcum.SetFocus
        Valida = False
    ElseIf Me.txtAge.Text = "" Then
        MsgBox "Debe ingresar un area.", vbInformation, "Aviso"
        Me.txtAge.SetFocus
        Valida = False
    ElseIf Me.txtComentario.Text = "" Then
        MsgBox "Debe ingresar un comentario.", vbInformation, "Aviso"
        Me.txtComentario.SetFocus
        Valida = False
    ElseIf Me.cboTpo.Text = "" Then
        MsgBox "Debe ingresar un Tipo.", vbInformation, "Aviso"
        Me.cboTpo.SetFocus
        Valida = False
    ElseIf Me.optBaja.value = False And Me.optTransf.value = False Then
        MsgBox "Debe seleccionar un Tipo de Bien.", vbInformation, "Aviso"
        Me.optBaja.SetFocus
        Valida = False
    Else
        Valida = True
    End If
End Function

'*** PEAC 20100511

Private Sub GenerarCodInventario(ByVal lsCodAge As String, ByVal lsBSCod As String)
    Dim sCodInventario As String
    sCodInventario = DevolverCorrelativo(Mid(lsBSCod, 1, 5), lsCodAge, lcCodCategoriaBien)   'JIPR20200313
    'If lcCodCategoriaBien = "1" Then '' bien depreciable
        'sCodInventario = DevolverCtaCont(Mid(lsBSCod, 1, 5)) & "0" & lsCodAge & Format(DevolverCorrelativo(Mid(lsBSCod, 1, 5)), "00000")
    'Else ''0=bien no depreciable
      '  sCodInventario = "45130111" & "0" & lsCodAge & Format(DevolverCorrelativoNoDepre(Mid(lsBSCod, 1, 5)), "00000")
   ' End If
    Me.txtSerie.Text = sCodInventario
End Sub

'JIPR20200313 COMENTÓ
'*** PEAC 20100511
'Private Function DevolverCtaCont(ByVal lsBSCod As String) As String
'    Dim sCodInventario As String
'    sCodInventario = ""
'    'Mobiliario
'    If lsBSCod = "11200" Then
'        sCodInventario = "181301"
'    End If
'
'    'Equipo de Computo
'    If lsBSCod = "11201" Then
'        sCodInventario = "181302"
'    End If
'
'    'Vehiculos
'    If lsBSCod = "11202" Then
'        sCodInventario = "181401"
'    End If
'
'    'Terrenos
'    If lsBSCod = "11203" Then
'        sCodInventario = "181101"
'    End If
'
'    'Edificios
'    If lsBSCod = "11204" Then
'        sCodInventario = "181201"
'    End If
'
'    'Instalaciones
'    If lsBSCod = "11205" Then
'        sCodInventario = "181202"
'    End If
'
'    'Mejoras Locales Propios (Edificios)
'    If lsBSCod = "11206" Then
'        sCodInventario = "181201"
'    End If
'
'    'If lsBSCod = "11207" Then
'    '    sCodInventario = "181201"
'    'End If
'
'    'Instalaciones en Locales Alquiladas
'    If lsBSCod = "11208" Then
'        sCodInventario = "181701"
'    End If
'
'    'mejoras Locales Alquilados
'    If lsBSCod = "11209" Then
'        sCodInventario = "181702"
'    End If
'
'    'If lsBSCod = "11210" Then
'    '    sCodInventario = "181702"
'    'End If
'
'    'Maquinarias
'    If lsBSCod = "11211" Then
'        sCodInventario = "181402"
'    End If
'
'    DevolverCtaCont = sCodInventario
'End Function



'*** PEAC 20100511
Private Function DevolverCorrelativo(ByVal sBSCod As String, Optional sAgencia As String, Optional nTipo As String) As String 'JIPR20200313 AGREGÓ Optional sAgencia As String, Optional nTipo As String
    Dim sCorrelativo As String
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Dim rs As ADODB.Recordset
    sCorrelativo = ""
    Set rs = oInventario.DarCorrelativoAF(sBSCod, sAgencia, nTipo) 'JIPR20200313
    sCorrelativo = rs!cSerieCod '+ 1 JIPR20200313
    DevolverCorrelativo = sCorrelativo
    Set rs = Nothing
End Function

'JIPR20200313
'*** PEAC 20100511
'Private Function DevolverCorrelativoNoDepre(ByVal lsBSCod As String) As String
'    Dim sCorrelativo As String
'    Dim oInventario As NInvActivoFijo
'    Set oInventario = New NInvActivoFijo
'    Dim rs As ADODB.Recordset
'    sCorrelativo = ""
'    Set rs = oInventario.DarCorrelativoNoDepre(lsBSCod)
'    sCorrelativo = rs!Maximo + 1
'    DevolverCorrelativoNoDepre = sCorrelativo
'    Set rs = Nothing
'End Function

Private Sub txtSeriePlaca_GotFocus()
    txtSeriePlaca.SelStart = 0
    txtSeriePlaca.SelLength = 50
End Sub

Private Sub txtSeriePlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.txtTasaDepre.SetFocus
        Me.cboTpo.SetFocus
    End If

End Sub


Private Sub txtSeriePlaca_LostFocus()
'    '*** PEAC 20100506
'
'    If MsgBox("¿Es un Bien Depreciable?", vbQuestion + vbYesNo, "Atención") = vbYes Then
'        lcTextoCategoriaBien = "BIEN DEPRECIABLE"
'        lcCodCategoriaBien = "1"
'    Else
'        lcTextoCategoriaBien = "BIEN NO DEPRECIABLE"
'        lcCodCategoriaBien = "0"
'    End If
'
'    Me.lblCategoriaBien.Caption = lcTextoCategoriaBien
'
'    If Mid(Me.txtAge.Text, 4, 2) = "" Then
'        lsAgencia = "01"
'    Else
'        lsAgencia = Mid(Me.txtAge.Text, 4, 2)
'    End If
'
'    GenerarCodInventario lsAgencia, Me.txtAF
'
'    '*** FIN PEAC

End Sub

Private Sub txtTasaDepre_GotFocus()
    txtTasaDepre.SelStart = 0
    txtTasaDepre.SelLength = 50
End Sub

Private Sub txtTasaDepre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMesAcum.SetFocus
    Else
        KeyAscii = NumerosDecimales(Me.txtTasaDepre, KeyAscii, 15)
    End If
End Sub

'*** PEAC 20120427
Private Function fgReemplazaCaracterEspecial(ByVal psNom As String) As String
Dim lsNombrePers As String
            
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, "-", "", , , vbTextCompare)), 80)
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, ".", " ", , , vbTextCompare)), 80)
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, "Ñ", "#", , , vbTextCompare)), 80)
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, "ñ", "#", , , vbTextCompare)), 80)
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, "'", "", , , vbTextCompare)), 80)
            
lsNombrePers = Replace(psNom, "-", "")
lsNombrePers = Replace(psNom, ".", " ")
lsNombrePers = Replace(psNom, "Ñ", "N")
lsNombrePers = Replace(psNom, "ñ", "n")
lsNombrePers = Replace(psNom, "'", "")
            
fgReemplazaCaracterEspecial = lsNombrePers
End Function

