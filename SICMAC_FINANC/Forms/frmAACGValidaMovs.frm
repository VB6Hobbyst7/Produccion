VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAACGValidaMovs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VALIDACION DE ASIENTOS DIA A DIA"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   90
      TabIndex        =   4
      Top             =   60
      Width           =   5385
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4065
         TabIndex        =   14
         Top             =   1050
         Width           =   1140
      End
      Begin VB.CommandButton cmdValidar 
         Caption         =   "&Validar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2850
         TabIndex        =   7
         Top             =   1050
         Width           =   1140
      End
      Begin VB.ComboBox cboCuenta 
         Height          =   315
         ItemData        =   "frmAACGValidaMovs.frx":0000
         Left            =   2930
         List            =   "frmAACGValidaMovs.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   420
         Width           =   1290
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmAACGValidaMovs.frx":003E
         Left            =   4320
         List            =   "frmAACGValidaMovs.frx":0048
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   420
         Width           =   870
      End
      Begin MSMask.MaskEdBox mskFecha1 
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   420
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecha2 
         Height          =   315
         Left            =   1510
         TabIndex        =   9
         Top             =   420
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   90
         X2              =   5160
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4335
         TabIndex        =   13
         Top             =   195
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2925
         TabIndex        =   12
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1515
         TabIndex        =   11
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   165
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   330
      Left            =   1650
      ScaleHeight     =   270
      ScaleWidth      =   3780
      TabIndex        =   0
      Top             =   1785
      Width           =   3840
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   165
         Left            =   15
         TabIndex        =   1
         Top             =   60
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Avance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label lblAvance 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   690
      TabIndex        =   2
      Top             =   1845
      Width           =   945
   End
End
Attribute VB_Name = "frmAACGValidaMovs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdValidar.SetFocus
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdValidar_Click()

Dim oPrevio As New PrevioFinan.clsPrevioFinan

Dim oCon As New DConecta
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim sSql As String

Dim dFecha1 As Date
Dim dFecha2 As Date

Dim cCuentaTemp As String
Dim nMontoAnt As Double

Dim nMontoDebe As Double
Dim nMontoHaber As Double

Dim lsCadena As String

Dim nMoneda As Integer

If Len(Trim(cboMoneda.Text)) = 0 Then
    MsgBox "Seleccione Moneda", vbInformation, "Aviso"
    Exit Sub
End If


If Len(Trim(cboCuenta.Text)) = 0 Then
    MsgBox "Seleccione Nivel", vbInformation, "Aviso"
    Exit Sub
End If

If IsDate(mskFecha1.Text) = False Then
    MsgBox "Fecha no válida", vbInformation, "Aviso"
    Exit Sub
End If

dFecha1 = mskFecha1.Text
dFecha2 = mskFecha2.Text

lsCadena = ""

If Month(dFecha1) <> Month(dFecha2) Or Year(dFecha1) <> Year(dFecha2) Then
    MsgBox "Fechas deben corresponder al mismo mes", vbInformation, "Aviso"
    Exit Sub
End If

nMoneda = Val(cboMoneda.Text)
PB1.Visible = True
'Sacamos todas las cuentas
sSql = "Select C1.cCtaContCod, C1.cCtaContDesc, C1.dFecha, "
If nMoneda = 1 Then
    sSql = sSql & " C1.nSaldoMN as nSaldo "
ElseIf nMoneda = 2 Then
    sSql = sSql & " C1.nSaldoME as nSaldo "
End If

sSql = sSql & " From "
sSql = sSql & " ( "
sSql = sSql & " Select  F.dFecha, CC.cCtacontcod, CC.cCtaContDesc, "

If nMoneda = 1 Then
    sSql = sSql & "   dbo.getsaldocta(F.dFecha, CC.cctacontcod,1) nSaldoMN "
ElseIf nMoneda = 2 Then
    sSql = sSql & "   dbo.getsaldocta(F.dFecha, CC.cctacontcod,2) nSaldoME  "
End If

sSql = sSql & " From    dbo.FechaTmp('" & Format(dFecha1, "MM/dd/YYYY") & "') F, ctacont CC "

sSql = sSql & " Where   CC.cctacontcod like '" & Trim(Str(cboCuenta.Text)) & "_" & Trim(Str(nMoneda)) & "%' "

sSql = sSql & "     and dbo.getsaldocta(F.dFecha,CC.cctacontcod,1) <> 0 "
sSql = sSql & "     and F.dFecha>='" & Format(dFecha1, "MM/dd/YYYY") & "' "
sSql = sSql & "     and F.dFecha < DateAdd(dd,1,'" & Format(dFecha2, "MM/dd/YYYY") & "') "
sSql = sSql & " ) C1 "
sSql = sSql & " Order By C1.cCtaContCod, C1.dFecha asc "

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
If rs.BOF Then
Else
    PB1.Min = 0
    PB1.Max = rs.RecordCount
    PB1.value = 0
    
    lblAvance.Caption = "0.00%"
    
    Do While Not rs.EOF
        PB1.value = PB1.value + 1
        lblAvance.Caption = Format(rs.AbsolutePosition / rs.RecordCount * 100, "0.00") & "%"
        DoEvents
        
        If cCuentaTemp <> rs!cCtaContCod Then
            
            'Saco Monto Anterior
            If nMoneda = 1 Then
                sSql = "Select ISNULL(dbo.getSaldoCta('" & Format(DateAdd("d", -1, dFecha1), "MM/dd/YYYY") & "', '" & rs!cCtaContCod & "', 1),0) as nMonto "
            ElseIf nMoneda = 2 Then
                sSql = "Select ISNULL(dbo.getSaldoCta('" & Format(DateAdd("d", -1, dFecha1), "MM/dd/YYYY") & "', '" & rs!cCtaContCod & "', 2),0) as nMonto "
            End If
            
            Set rs1 = oCon.CargaRecordSet(sSql)
            If rs1.BOF Then
                nMontoAnt = 0
            Else
                nMontoAnt = rs1!nMonto
            End If
            rs1.Close
            
            cCuentaTemp = rs!cCtaContCod
            
        End If
        
        sSql = "SELECT "
        
        If nMoneda = 1 Then
            sSql = sSql & " ISNULL(SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END),0) as nDebe, "
            sSql = sSql & " ISNULL(SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END),0) as nHaber "
        ElseIf nMoneda = 2 Then
            sSql = sSql & " ISNULL(SUM(CASE WHEN me.nMovMEImporte > 0 THEN me.nMovMEImporte END),0) as nDebe, "
            sSql = sSql & " ISNULL(SUM(CASE WHEN me.nMovMEImporte < 0 THEN me.nMovMEImporte * -1 END),0) as nHaber "
        End If
        
        sSql = sSql & " FROM   Mov M "
        sSql = sSql & "     JOIN MovCta a "
        sSql = sSql & "         ON a.nMovNro = M.nMovNro "
        sSql = sSql & "     LEFT JOIN MovME me "
        sSql = sSql & "         ON me.nMovNro = a.nMovNro and me.nMovItem = a.nMovItem "
        sSql = sSql & "     JOIN CtaCont c "
        sSql = sSql & "         ON c.cCtaContCod = a.cCtaContCod "
        sSql = sSql & " WHERE   M.nMovEstado = '10' and not M.nMovFlag in ('1', '5') and a.cCtaContCod Like '" & rs!cCtaContCod & "%' "
        sSql = sSql & "  and substring(M.cmovnro,1,8) Like '" & Format(rs!dFecha, "YYYYMMdd") & "%' "
        sSql = sSql & " GROUP BY a.cCtaContCod, c.cCtaContDesc "
        
        Set rs1 = oCon.CargaRecordSet(sSql)
        If rs1.BOF Then
            nMontoDebe = 0
            nMontoHaber = 0
        Else
            nMontoDebe = rs1!nDebe
            nMontoHaber = rs1!nHaber
        End If
        rs1.Close
        
        If nMontoAnt + nMontoDebe - nMontoHaber <> rs!nSaldo Then
            lsCadena = lsCadena & "Cuenta: " & rs!cCtaContCod & " Fecha: " & rs!dFecha & " Saldo Ant: " & Format(nMontoAnt, "0.00")
            lsCadena = lsCadena & " Debe: " & Format(nMontoDebe, "0.00") & " Haber: " & Format(nMontoHaber, "0.00")
            lsCadena = lsCadena & " Saldo New: " & Format(rs!nSaldo, "0.00") & Chr(10)
        End If
                
        nMontoAnt = rs!nSaldo
        
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

If Len(Trim(lsCadena)) > 0 Then
    
    lsCadena = Chr(10) & Chr(10) & "DIFERENCIAS ENCONTRADAS" & Chr(10) & "=======================" & Chr(10) & Chr(10) & lsCadena

    oPrevio.Show lsCadena, "Difencias Encontradas", True, , gImpresora
End If


End Sub


Private Sub Form_Load()
CentraForm Me
End Sub

Private Sub mskFecha1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFecha2.SetFocus
    End If
End Sub
 

Private Sub mskFecha2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboCuenta.SetFocus
    End If
End Sub
