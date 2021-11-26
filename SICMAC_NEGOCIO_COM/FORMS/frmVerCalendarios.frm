VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmVerCalendarios 
   Caption         =   "Verificacion de Calendarios"
   ClientHeight    =   6405
   ClientLeft      =   1575
   ClientTop       =   2550
   ClientWidth     =   10770
   Icon            =   "frmVerCalendarios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   10770
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   300
      TabIndex        =   26
      Top             =   5865
      Width           =   1500
   End
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9195
      TabIndex        =   19
      Top             =   5895
      Width           =   1410
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5640
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   9948
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Plan Pagos SIAFC"
      TabPicture(0)   =   "frmVerCalendarios.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "lblNomCliSIAFC"
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(3)=   "lblVigenciaSIAFC"
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(5)=   "lblDesembSiafc"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "lblCodAna"
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(9)=   "lblTasaSIAFC"
      Tab(0).Control(10)=   "lblLineaSIAFC"
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(12)=   "dgSiafc"
      Tab(0).Control(13)=   "txtCodCta"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Plan Pagos SICMAC"
      TabPicture(1)   =   "frmVerCalendarios.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblNomCliSICMAC"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblVigenciaSICMAC"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lbldesembSICMAC"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblCtaCod"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label9"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblTasIntSICMAC"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblLineaSICMAC"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label12"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "dgSicmac"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtCodCli"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lstCtasSICMAC"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.ListBox lstCtasSICMAC 
         Height          =   645
         Left            =   7440
         TabIndex        =   12
         Top             =   525
         Width           =   2805
      End
      Begin SICMACT.TxtBuscar txtCodCli 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Top             =   495
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin MSMask.MaskEdBox txtCodCta 
         Height          =   375
         Left            =   -73335
         TabIndex        =   4
         Top             =   555
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   23
         Mask            =   "###-###-##-#-#######.##"
         PromptChar      =   "_"
      End
      Begin MSDataGridLib.DataGrid dgSiafc 
         Height          =   3675
         Left            =   -74880
         TabIndex        =   1
         Top             =   1785
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   6482
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgSicmac 
         Height          =   3525
         Left            =   135
         TabIndex        =   2
         Top             =   1860
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   6218
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "Linea Credito:"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   1455
         Width           =   990
      End
      Begin VB.Label lblLineaSICMAC 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1365
         TabIndex        =   29
         Top             =   1455
         Width           =   5610
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Linea de Credito:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   28
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lblLineaSIAFC 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -73335
         TabIndex        =   27
         Top             =   1290
         Width           =   4110
      End
      Begin VB.Label lblTasIntSICMAC 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5805
         TabIndex        =   25
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Int.:"
         Height          =   195
         Left            =   5085
         TabIndex        =   24
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label lblTasaSIAFC 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -66135
         TabIndex        =   23
         Top             =   1005
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Int.:"
         Height          =   195
         Left            =   -66885
         TabIndex        =   22
         Top             =   1035
         Width           =   675
      End
      Begin VB.Label lblCodAna 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -67845
         TabIndex        =   21
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista:"
         Height          =   195
         Left            =   -68565
         TabIndex        =   20
         Top             =   1005
         Width           =   600
      End
      Begin VB.Label lblCtaCod 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3195
         TabIndex        =   18
         Top             =   495
         Width           =   2850
      End
      Begin VB.Label lbldesembSICMAC 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3690
         TabIndex        =   17
         Top             =   1170
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desembolso:"
         Height          =   195
         Left            =   2730
         TabIndex        =   16
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label lblVigenciaSICMAC 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1365
         TabIndex        =   15
         Top             =   1140
         Width           =   1305
      End
      Begin VB.Label Label5 
         Caption         =   "Vigencia :"
         Height          =   255
         Left            =   195
         TabIndex        =   14
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label lblNomCliSICMAC 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1380
         TabIndex        =   13
         Top             =   810
         Width           =   5670
      End
      Begin VB.Label lblDesembSiafc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -70245
         TabIndex        =   10
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Capital Desembolsado:"
         Height          =   195
         Left            =   -71925
         TabIndex        =   9
         Top             =   1020
         Width           =   1620
      End
      Begin VB.Label lblVigenciaSIAFC 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -73335
         TabIndex        =   8
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Vigencia :"
         Height          =   255
         Left            =   -74730
         TabIndex        =   7
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lblNomCliSIAFC 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   -70935
         TabIndex        =   6
         Top             =   540
         Width           =   6165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Cliente:"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Credito SIAFC:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   3
         Top             =   615
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmVerCalendarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oConexFox As ADODB.Connection
Dim lsRutaCli As String
Dim lsRutaCred As String
Public Sub Inicio(ByVal psRutaCli As String, ByVal psRutaCred As String, oConFox As ADODB.Connection)
lsRutaCli = psRutaCli
lsRutaCred = psRutaCred
Set oConexFox = oConFox
Me.Show 1
End Sub
Sub GetCalendarioSiCMAC(ByVal lsCodCta As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta

oCon.AbreConexion

sql = "SELECT   C.CCTACOD AS [N° CREDITO], C.NCUOTA AS CUOTA , C.DVENC AS VENCIMIENTO,C.nColocCalendEstado AS ESTADO,  " _
    & "         SUM(CASE WHEN nPrdConceptoCod IN (1000,1100,1102) THEN NMONTO ELSE 0 END) AS CUOTA, " _
    & "         SUM(CASE WHEN nPrdConceptoCod=1000 THEN NMONTO ELSE 0 END) AS CAPITAL, " _
    & "         SUM(CASE WHEN nPrdConceptoCod=1100 THEN NMONTO ELSE 0 END) AS INTERES, " _
    & "         SUM(CASE WHEN nPrdConceptoCod=1102 THEN NMONTO ELSE 0 END) AS GRACIA, " _
    & "         SUM(CASE WHEN nPrdConceptoCod NOT IN (1000,1100,1102) THEN NMONTO ELSE 0 END) AS OTROS, " _
    & "         SUM(CASE WHEN nPrdConceptoCod=1000 THEN nMontoPagado ELSE 0 END) AS CAPPAG, " _
    & "         SUM(CASE WHEN nPrdConceptoCod=1100 THEN nMontoPagado ELSE 0 END) AS INTPAG, " _
    & "         C.nNroCalen , C.nColocCalendApl " _
    & " FROM    COLOCCALENDARIO C " _
    & "         JOIN COLOCCALENDDET CD ON CD.cCtaCod =C.cCtaCod AND CD.nNroCalen =C.nNroCalen AND CD.nColocCalendApl = C.nColocCalendApl AND CD.nCuota= C.nCuota " _
    & " WHERE   C.nColocCalendApl = 1 AND C.CCTACOD ='" & lsCodCta & "' " _
    & "         AND C.nNroCalen = (SELECT CR.nNroCalen FROM COLOCACCRED CR WHERE CR.CCTACOD = C.CCTACOD ) " _
    & " GROUP BY C.CCTACOD, C.NCUOTA , C.DVENC ,C.nColocCalendEstado, C.NNROCALEN,  C.nColocCalendApl "

Set rs = oCon.CargaRecordSet(sql)
Set dgSicmac.DataSource = rs
dgSicmac.Refresh
oCon.CierraConexion
End Sub
Sub GetCalendarioSIAFC(ByVal lsCodCta As String)
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT CCODCTA AS NRO_CREDITO, CNROCUO AS CUOTA, DFECVEN AS VENCIMIENTO,cEstado AS ESTADO,    " _
     & " NCAPITA + NINTERE AS CUOTA ,  NCAPITA AS CAPITAL, NINTERE AS INTERES, NCAPPAG AS CAPPAG, NINTPAG AS NINTPAG   " _
    & " FROM " & lsRutaCred & "KPYDPPG WHERE CCODCTA ='" & lsCodCta & "' AND CTIPOPE <>'D'"
Set rs = New ADODB.Recordset

rs.CursorLocation = adUseClient
rs.Open sql, oConexFox, adOpenStatic, adLockOptimistic, adCmdText
rs.ActiveConnection = Nothing
Set dgSiafc.DataSource = rs
dgSiafc.Refresh
End Sub
Sub GetDatosSIAFC(ByVal lsCodCta As String)
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT K.CCODCTA, K.CCODCLI, C.CNOMCLI, K.NCAPDES, K.DFECVIG, K.NTASINT, K.cCodAna, L.CDESLIN, L.CCODLIN " _
    & " FROM    " & lsRutaCred & "KPYMCRE K  " _
    & "         JOIN " & lsRutaCli & "CLIMIDE C ON C.CCODCLI = K.CCODCLI " _
    & "         join " & lsRutaCred & "KPYTLIN L ON L.CNROLIN = K.CNROLIN " _
    & " WHERE   K.CCODCTA ='" & lsCodCta & "' "

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sql, oConexFox, adOpenStatic, adLockOptimistic, adCmdText
rs.ActiveConnection = Nothing

If Not rs.EOF And Not rs.BOF Then
    Me.lblNomCliSIAFC = rs!cNomCli
    Me.lblVigenciaSIAFC = rs!dFecVig
    Me.lblDesembSiafc = Format(rs!nCapDes, "#0.00")
    Me.lblCodAna = rs!cCodAna
    Me.lblTasaSIAFC = Format(rs!nTasInt, "#0.00")
    Me.lblLineaSIAFC = "[" & Trim(rs!CCODLIN) & "]     " & Trim(rs!CDESLIN)
    GetCalendarioSIAFC rs!cCodCta
End If
rs.Close
Set rs = Nothing

End Sub

Private Sub cmdcancelar_Click()
Me.lblCodAna = ""
Me.lblCtaCod = ""
Me.lblDesembSiafc = "0.00"
Me.lblNomCliSIAFC = ""
Me.lbldesembSICMAC = "0.00"
Me.lblNomCliSICMAC = ""
Me.lblTasaSIAFC = "0.00"
Me.lblTasIntSICMAC = "0.00"
Me.lblVigenciaSIAFC = ""
Me.lblVigenciaSICMAC = ""
Me.txtCodCta.Mask = ""
Me.txtCodCta = ""
Me.txtCodCta.Mask = "###-###-##-#-#######.##"
Me.txtCodCli = ""
Me.lblLineaSIAFC = ""
Me.lblLineaSICMAC = ""
Me.lstCtasSICMAC.Clear

GetCalendarioSiCMAC ""
GetCalendarioSIAFC ""
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub dgSiafc_GotFocus()
Me.dgSiafc.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub dgSiafc_LostFocus()
Me.dgSiafc.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub dgSicmac_GotFocus()
Me.dgSicmac.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub dgSicmac_LostFocus()
Me.dgSicmac.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub lstCtasSICMAC_Click()
GetDatosCreditoSICMAC lstCtasSICMAC
End Sub

Private Sub lstCtasSICMAC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    GetCalendarioSiCMAC lstCtasSICMAC
    lstCtasSICMAC.Visible = False
End If
End Sub

Private Sub SSTab1_DblClick()
txtCodCta_KeyPress 13
End Sub

Private Sub txtCodCli_EmiteDatos()
If txtCodCli = "" Then Exit Sub
lstCtasSICMAC.Visible = True
GetDatosSICMAC txtCodCli
Me.lblNomCliSICMAC = txtCodCli.psDescripcion
lstCtasSICMAC.SetFocus
End Sub

Private Sub txtCodCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    GetDatosSIAFC Replace(Replace(txtCodCta, "-", ""), ".", "")
End If

End Sub
Sub GetDatosSICMAC(ByVal psPerscod As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta


oCon.AbreConexion

sql = "SELECT  PP.CCTACOD, PP.NTASAINTERES, PP.NSALDO, P.CPERSNOMBRE " _
    & " FROM    PERSONA P " _
    & "         JOIN PRODUCTOPERSONA R ON R.CPERSCOD = P.CPERSCOD AND nPrdPersRelac= 20 " _
    & "         JOIN PRODUCTO PP ON PP.CCTACOD = R.CCTACOD " _
    & " WHERE   PP.nPrdEstado IN (2020, 2021,2022, 2030, 2031,2032,2002) AND P.CPERSCOD = '" & psPerscod & "'"

Set rs = oCon.CargaRecordSet(sql)
lstCtasSICMAC.Clear
Do While Not rs.EOF
    lstCtasSICMAC.AddItem rs!cCtaCod
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
oCon.CierraConexion
End Sub
Sub GetDatosCreditoSICMAC(ByVal lsCodCta As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta

oCon.AbreConexion

sql = "SELECT   C.CCTACOD, C.NCUOTA AS CUOTA , C.DVENC AS VENCIMIENTO,C.nColocCalendEstado AS ESTADO,  " _
    & "         SUM(CASE WHEN nPrdConceptoCod=1000 THEN NMONTO ELSE 0 END) AS CAPITAL, P.nTasaInteres, L.cLineaCred, L.cDescripcion " _
    & " FROM    COLOCCALENDARIO C " _
    & "         JOIN COLOCCALENDDET CD ON CD.cCtaCod =C.cCtaCod AND CD.nNroCalen =C.nNroCalen AND CD.nColocCalendApl = C.nColocCalendApl AND CD.nCuota= C.nCuota " _
    & "         join producto p ON P.CCTACOD = C.CCTACOD " _
    & "         JOIN COLOCACIONES CL ON CL.CCTACOD = P.CCTACOD " _
    & "         JOIN COLOCLINEACREDITO L ON L.cLineaCred = CL.cLineaCred " _
    & " WHERE   C.nColocCalendApl = 0 AND C.CCTACOD ='" & lsCodCta & "' AND C.nColocCalendEstado= 1 " _
    & "         AND C.nNroCalen = (SELECT CR.nNroCalen FROM COLOCACCRED CR WHERE CR.CCTACOD = C.CCTACOD ) " _
    & " GROUP BY C.CCTACOD, C.NCUOTA , C.DVENC ,C.nColocCalendEstado, C.NNROCALEN,  C.nColocCalendApl,P.nTasaInteres,L.cLineaCred, L.cDescripcion"

Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    lbldesembSICMAC = Format(rs!capital, "#0.00")
    lblVigenciaSICMAC = rs!Vencimiento
    Me.lblTasIntSICMAC = Format(rs!nTasaInteres, "#0.00")
    Me.lblCtaCod = rs!cCtaCod
    Me.lblLineaSICMAC = "[" & Trim(rs!cLineaCred) & "]   " & Trim(rs!cDescripcion)
    rs.MoveNext
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion

End Sub
