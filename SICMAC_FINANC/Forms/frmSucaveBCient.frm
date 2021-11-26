VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSucaveBCient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sucave: BCIENT"
   ClientHeight    =   1935
   ClientLeft      =   2070
   ClientTop       =   2835
   ClientWidth     =   4920
   Icon            =   "frmSucaveBCient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   1935
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   2535
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   405
      Left            =   1245
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
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
      Height          =   1305
      Left            =   105
      TabIndex        =   4
      Top             =   75
      Width           =   4725
      Begin VB.CheckBox chkCalculoUtilidad 
         Caption         =   "Calculo Utilidad"
         Height          =   240
         Left            =   2925
         TabIndex        =   9
         Top             =   847
         Width           =   1560
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmSucaveBCient.frx":030A
         Left            =   2910
         List            =   "frmSucaveBCient.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   1620
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   330
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNroDigitos 
         Height          =   315
         Left            =   2250
         TabIndex        =   8
         Top             =   810
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número de Dígitos :"
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
         Left            =   240
         TabIndex        =   7
         Top             =   870
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
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
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmSucaveBCient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql      As String
Dim rs        As New ADODB.Recordset
Dim nTipoBala As Integer
Dim nMoneda   As Integer
Dim oBarra    As clsProgressBar

Private Sub GeneraBCient()
Dim sCondBala As String, sCondBala2 As String
Dim sCondBala6 As String
Dim sCond1 As String, sCond2 As String
Dim sCta   As String
Dim sCta2   As String
Dim N      As Integer
Dim nPos   As Variant
Dim FecIni As Date
Dim FecFin As Date
Dim R As New ADODB.Recordset
Dim sCodigo As String
Dim sSalIni As String
Dim sDebe As String
Dim sHaber As String
Dim sSalFin As String
Dim CadRep As String
Dim nTotal As Integer
Dim cPar2(6) As Integer
Dim I As Integer
Dim ContBarra As Long
Dim Total As Integer
Dim CadTemp As String

Dim dBalance As New DbalanceCont
Dim lnBalanceCate As Integer
On Error GoTo SucaveERR


lnBalanceCate = 5
   cPar2(0) = 0
   cPar2(1) = 1
   cPar2(2) = 2
   cPar2(3) = 3
   cPar2(4) = 4
   cPar2(5) = 6
   FecIni = CDate("01" + "/" + Right("0" + Trim(Right(Trim(cboMes.Text), 2)), 2) + "/" + txtAnio.Text)
   FecFin = DateAdd("m", 1, FecIni) - 1
       
   DoEvents
   MousePointer = 11
   Set oBarra = New clsProgressBar
   oBarra.ShowForm Me
   oBarra.Max = 7
   oBarra.Progress 0, "BCIENT", "", "Eliminando BCient"
   dBalance.EliminaBCient FecFin
   For I = 0 To 5
      nMoneda = cPar2(I)
      DoEvents
      MousePointer = 11
          
      dBalance.EliminaBalanceTemp lnBalanceCate, nMoneda
      oBarra.Progress I + 1, "BCIENT", "", "Generando BCient"
      
      'Saldos Iniciales
      dBalance.InsertaBalanceTmpSaldos lnBalanceCate, nMoneda, Format(FecIni - 1, gsFormatoFecha), True
      DoEvents
      
      'Movimientos del Mes
      dBalance.InsertaMovimientosMes lnBalanceCate, nMoneda, Format(FecIni, "yyyymmdd"), Format(FecFin, "yyyymmdd"), True
      DoEvents
      sCodigo = txtAnio.Text + Right("0" + Trim(Right(Trim(cboMes.Text), 2)), 2) + gsCodCMAC
      dBalance.InsertaBCient FecFin, sCodigo, lnBalanceCate, nMoneda
      dBalance.ActualizaBCient FecFin, sCodigo, lnBalanceCate, nMoneda
   Next I
   
   DoEvents
   CadRep = ""
   CadTemp = ""
   Set R = dBalance.CargaBCient(Format(FecFin, gsFormatoFecha), txtNroDigitos, IIf(chkCalculoUtilidad.value = 1, True, False), gsCodCMAC)
   
   If Not R.BOF And Not R.EOF Then
      oBarra.Max = R.RecordCount
      Do While Not R.EOF
        DoEvents
         'Creacion de Reporte
         sCodigo = R!cCtaContCod
         sSalIni = IIf(R!nSaldoIniImporte >= 0, "+", "") + Left(Format(R!nSaldoIniImporte, "#0.00"), Len(Format(R!nSaldoIniImporte, "#0.00")) - 3) + Right(Format(R!nSaldoIniImporte, "#0.00"), 2)
         sSalFin = IIf(R!nSaldoFinImporte >= 0, "+", "") + Left(Format(R!nSaldoFinImporte, "#0.00"), Len(Format(R!nSaldoFinImporte, "#0.00")) - 3) + Right(Format(R!nSaldoFinImporte, "#0.00"), 2)
         sDebe = IIf(R!nDebe >= 0, "+", "") + Left(Format(R!nDebe, "#0.00"), Len(Format(R!nDebe, "#0.00")) - 3) + Right(Format(R!nDebe, "#0.00"), 2)
         sHaber = IIf(R!nHaber >= 0, "+", "") + Left(Format(R!nHaber, "#0.00"), Len(Format(R!nHaber, "#0.00")) - 3) + Right(Format(R!nHaber, "#0.00"), 2)
         
         CadTemp = CadTemp + sCodigo + space(18 - Len(sSalIni)) + sSalIni + space(18 - Len(sDebe)) + sDebe + space(18 - Len(sHaber)) + sHaber + space(18 - Len(sSalFin)) + sSalFin + oImpresora.gPrnSaltoLinea
         
         If Len(CadTemp) >= 1000 Then
            CadRep = CadRep & CadTemp
            CadTemp = ""
         End If
         
         ContBarra = ContBarra + 1
         oBarra.Progress R.Bookmark, "BCIENT", "Generando Formato TXT", "Procesando...", vbBlue
         R.MoveNext
      Loop
   End If
   RSClose R
   If Len(Trim(CadTemp)) > 0 Then
     CadRep = CadRep & CadTemp
     CadTemp = ""
   End If
   oBarra.CloseForm Me
   Set dBalance = Nothing
   MousePointer = 0
   
   EnviaPrevio CadRep, "SUCAVE: BCIENT", gnLinPage, False
   CadRep = ""
   
Exit Sub
SucaveERR:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdProcesar.SetFocus
    End If
End Sub

Private Sub cmdProcesar_Click()

    'PEAC 20210514
    If Len(Trim$(txtNroDigitos.Text)) = 0 Then
        If MsgBox("Por favor ingrese el número de dígitos.", vbOKOnly, "Atención") = vbOk Then
            Exit Sub
        End If
    ElseIf Len(Trim$(txtAnio.Text)) = 0 Then
        If MsgBox("Por favor ingrese el año.", vbOKOnly, "Atención") = vbOk Then
            Exit Sub
        End If
    End If
    'FIN PEAC

   If ValidaAnio(txtAnio) Then
       Call GeneraBCient
   End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
   CentraForm Me
   frmMdiMain.Enabled = False
   txtAnio = Year(gdFecSis)
   cboMes.ListIndex = Month(gdFecSis) - 1
   txtNroDigitos.Text = "14"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMdiMain.Enabled = True
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
   If KeyAscii = 13 Then
       cboMes.SetFocus
   End If
End Sub

Private Sub txtNroDigitos_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
   If KeyAscii = 13 Then
       CmdProcesar.SetFocus
   End If

End Sub
