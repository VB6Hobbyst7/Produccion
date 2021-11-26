VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFTPBSI_Reportxt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sectorial Institucional"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmFTPBSI_Reportxt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4725
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmFTPBSI_Reportxt.frx":030A
         Left            =   2880
         List            =   "frmFTPBSI_Reportxt.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   1620
      End
      Begin VB.CheckBox chkCalculoUtilidad 
         Caption         =   "Calculo Utilidad"
         Height          =   240
         Left            =   2925
         TabIndex        =   7
         Top             =   847
         Width           =   1560
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   900
         TabIndex        =   2
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
         TabIndex        =   6
         Top             =   810
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         TabIndex        =   3
         Top             =   360
         Width           =   480
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
         TabIndex        =   1
         Top             =   360
         Width           =   465
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
         TabIndex        =   5
         Top             =   870
         Width           =   1725
      End
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   405
      Left            =   1140
      TabIndex        =   8
      Top             =   1365
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   2430
      TabIndex        =   9
      Top             =   1365
      Width           =   1215
   End
End
Attribute VB_Name = "frmFTPBSI_Reportxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'Nombre : frmFTPBSI_Reportxt
'Descripcion:Formulario para Generar el Archivo de Balance Sectorial Institucional(BSI)
'Creacion: PASI TI-ERS002-2017
'********************************************************************
Option Explicit
Dim sSql      As String
Dim rs        As New ADODB.Recordset
Dim nTipoBala As Integer
Dim nMoneda   As Integer
Dim oBarra    As clsProgressBar
Private Sub GeneraFTPBSI()
Dim sCondBala As String, sCondBala2 As String
Dim sCondBala6 As String
Dim sCond1 As String, sCond2 As String
Dim sCta   As String
Dim sCta2   As String
Dim n      As Integer
Dim nPos   As Variant
Dim FecIni As Date
Dim FecFin As Date
Dim R, R2 As New ADODB.Recordset
Dim sCodigo As String
Dim sSalIni As String
Dim sDebe As String
Dim sHaber As String
Dim sSalFin As String
Dim CadRep As String
Dim nTotal As Integer
Dim cPar2(6) As Integer
Dim i As Integer
Dim ContBarra As Long
Dim Total As Integer
Dim CadTemp As String
Dim lsCuentaTemporal As String
Dim dBalance As New DbalanceCont
Dim lnBalanceCate As Integer
Dim lsBSI_Temp As String
On Error GoTo SucaveERR
lnBalanceCate = 7
   cPar2(0) = 0
   cPar2(1) = 1
   cPar2(2) = 2
'   cPar2(3) = 3
'   cPar2(4) = 4
'   cPar2(5) = 6
   FecIni = CDate("01" + "/" + Right("0" + Trim(Right(Trim(cboMes.Text), 2)), 2) + "/" + txtAnio.Text)
   FecFin = DateAdd("m", 1, FecIni) - 1
       
   DoEvents
   MousePointer = 11
   Set oBarra = New clsProgressBar
   oBarra.ShowForm Me
   oBarra.Max = 7
   oBarra.Progress 0, "BSI", "", "Eliminando BSI"
   dBalance.EliminaBSI FecFin
   For i = 0 To 2
      nMoneda = cPar2(i)
      DoEvents
      MousePointer = 11
          
      dBalance.EliminaBalanceTemp lnBalanceCate, nMoneda
      oBarra.Progress i + 1, "BSI", "", "Generando BSI"
      
      'Saldos Iniciales
      dBalance.InsertaBalanceTmpSaldos lnBalanceCate, nMoneda, Format(FecIni - 1, gsFormatoFecha), True
      DoEvents
      
      'Movimientos del Mes
      dBalance.InsertaMovimientosMes lnBalanceCate, nMoneda, Format(FecIni, "yyyymmdd"), Format(FecFin, "yyyymmdd"), True
      DoEvents
      sCodigo = txtAnio.Text + Right("0" + Trim(Right(Trim(cboMes.Text), 2)), 2) + gsCodCMAC
      dBalance.InsertaFTPBSI FecFin, sCodigo, lnBalanceCate, nMoneda
      dBalance.ActualizaFTPBSI FecFin, sCodigo, lnBalanceCate, nMoneda
   Next i
   
   DoEvents
   CadRep = ""
   CadTemp = ""
   lsCuentaTemporal = ""
   
   dBalance.CargarSectorCreditos FecFin
   Set R = dBalance.CargaFTPBSI(Format(FecFin, gsFormatoFecha), txtNroDigitos, IIf(chkCalculoUtilidad.value = 1, True, False), gsCodCMAC)
   'CadTemp = "BSI109201606P" + oImpresora.gPrnSaltoLinea /*Comments PASI20170731*/
   CadTemp = "BSI109" + Trim(CStr(txtAnio.Text)) + Right("0" + Trim(Right(Trim(cboMes.Text), 2)), 2) + IIf(DatePart("D", gdFecSis) <= 9, "P", "D") + oImpresora.gPrnSaltoLinea
   If Not R.BOF And Not R.EOF Then
      oBarra.Max = R.RecordCount
      Do While Not R.EOF
        DoEvents
         
         'Creacion de Reporte
         sCodigo = R!cCtaContCod
         sSalIni = IIf(R!nSaldoIniImporte >= 0, "", "") + Left(Format(R!nSaldoIniImporte, "#0.00"), Len(Format(R!nSaldoIniImporte, "#0.00")) - 3) + Right(Format(R!nSaldoIniImporte, "#0.00"), 2)
         sSalFin = IIf(R!nSaldoFinImporte >= 0, "", "") + Left(Format(R!nSaldoFinImporte, "#0.00"), Len(Format(R!nSaldoFinImporte, "#0.00")) - 3) + Right(Format(R!nSaldoFinImporte, "#0.00"), 2)
         If Left(sSalIni, 1) = "-" Then
            sSalIni = "-" + String(17 - Len((CStr(Abs(sSalIni)))), "0") + CStr(Abs(sSalIni))
         Else
            sSalIni = String(18 - Len(sSalIni), "0") + CStr(sSalIni)
         End If
         
         If Left(sSalFin, 1) = "-" Then
            sSalFin = "-" + String(17 - Len((CStr(Abs(sSalFin)))), "0") + CStr(Abs(sSalFin))
         Else
            sSalFin = String(18 - Len(sSalFin), "0") + CStr(sSalFin)
         End If
         'CadTemp = CadTemp + sCodigo + String(12, "X") + String(18 - Len(sSalIni), "0") + sSalIni + String(18 - Len(sSalFin), "0") + sSalFin + oImpresora.gPrnSaltoLinea
         If Mid(sCodigo, 1, 19) = "2016071091425120602" Then
            MsgBox ""
         End If
         If R!nSector = 1 Then
            If Mid(sCodigo, 10, 6) = "310101" Or Mid(sCodigo, 10, 6) = "311101" Then
                lsBSI_Temp = "SP3000000000"
                CadTemp = CadTemp + sCodigo + String(12, "X") + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea
                CadTemp = CadTemp + sCodigo + lsBSI_Temp + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea
            ElseIf Mid(sCodigo, 10, 4) = "2406" Or Mid(sCodigo, 10, 4) = "2606" Or Mid(sCodigo, 10, 4) = "2416" Or Mid(sCodigo, 10, 4) = "2616" Or Mid(sCodigo, 10, 4) = "2426" Or Mid(sCodigo, 10, 4) = "2626" Then
                lsBSI_Temp = "SF2OIFFOM000"
                 CadTemp = CadTemp + sCodigo + String(12, "X") + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea
                CadTemp = CadTemp + sCodigo + lsBSI_Temp + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea
            Else
                lsBSI_Temp = "EH1OTS000000"
                Set R2 = dBalance.CargarBSI_Sector(Mid(sCodigo, 1, 19))
                If Not (R2.BOF Or R2.EOF) Then
                CadTemp = CadTemp + sCodigo + String(12, "X") + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea
                Do While Not R2.EOF
                    lsBSI_Temp = R2!cPersCodBSI
                    sSalIni = IIf(R2!nSaldoInicial >= 0, "", "") + Left(Format(R2!nSaldoInicial, "#0.00"), Len(Format(R2!nSaldoInicial, "#0.00")) - 3) + Right(Format(R2!nSaldoInicial, "#0.00"), 2)
                    sSalFin = IIf(R2!nSaldofinal >= 0, "", "") + Left(Format(R2!nSaldofinal, "#0.00"), Len(Format(R2!nSaldofinal, "#0.00")) - 3) + Right(Format(R2!nSaldofinal, "#0.00"), 2)
                    If Left(sSalIni, 1) = "-" Then
                        sSalIni = "-" + String(17 - Len((CStr(Abs(sSalIni)))), "0") + CStr(Abs(sSalIni))
                    Else
                        sSalIni = String(18 - Len(sSalIni), "0") + CStr(sSalIni)
                    End If
         
                    If Left(sSalFin, 1) = "-" Then
                        sSalFin = "-" + String(17 - Len((CStr(Abs(sSalFin)))), "0") + CStr(Abs(sSalFin))
                    Else
                        sSalFin = String(18 - Len(sSalFin), "0") + CStr(sSalFin)
                    End If
                    CadTemp = CadTemp + sCodigo + lsBSI_Temp + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea
                    R2.MoveNext
                Loop
                Else
                 'CadTemp = CadTemp + sCodigo + String(12, "X") + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea   'VAPA20170929 COMENTADO
                 CadTemp = CadTemp + sCodigo + String(12, "X") + sSalIni + sSalFin + oImpresora.gPrnSaltoLineaDef
                 'CadTemp = CadTemp + sCodigo + lsBSI_Temp + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea        'VAPA20170929 COMENTADO
                 CadTemp = CadTemp + sCodigo + lsBSI_Temp + sSalIni + sSalFin + oImpresora.gPrnSaltoLineaDef
                End If
            End If
           
         Else
            CadTemp = CadTemp + sCodigo + String(12, "X") + sSalIni + sSalFin + oImpresora.gPrnSaltoLinea
         End If
         
         If Len(CadTemp) >= 1000 Then
            CadRep = CadRep & CadTemp
            CadTemp = ""
         End If
         
         ContBarra = ContBarra + 1
         oBarra.Progress R.Bookmark, "BSI", "Generando Formato TXT", "Procesando...", vbBlue
         R.MoveNext
         lsCuentaTemporal = ""
      Loop
   End If
   'RSClose R
   If Len(Trim(CadTemp)) > 0 Then
     CadRep = CadRep & CadTemp
     CadTemp = ""
   End If
   oBarra.CloseForm Me
   Set dBalance = Nothing
   MousePointer = 0
   
   EnviaPrevio CadRep, "FTP: BSI", gnLinPage, False
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
    If Not ValidaAnio(txtAnio) Then Exit Sub
    If cboMes.ListIndex = -1 Then MsgBox "No se ha seleccionado el periodo. Verifique", vbInformation: cboMes.SetFocus: Exit Sub
    Call GeneraFTPBSI
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
   CentraForm Me
   frmMdiMain.Enabled = False
   txtAnio = Year(gdFecSis)
   cboMes.ListIndex = Month(gdFecSis) - 1
   txtNroDigitos.Text = "20"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMdiMain.Enabled = True
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
   If KeyAscii = 13 Then cboMes.SetFocus
End Sub
Private Sub txtNroDigitos_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
   If KeyAscii = 13 Then CmdProcesar.SetFocus
End Sub

