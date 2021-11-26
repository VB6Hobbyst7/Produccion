VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBCREnvio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " BCR: Generacion de archivo txt"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   Icon            =   "frmBCREnvio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4950
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
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4725
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmBCREnvio.frx":030A
         Left            =   2910
         List            =   "frmBCREnvio.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Width           =   1620
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   900
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   870
         Width           =   1725
      End
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   405
      Left            =   1140
      TabIndex        =   1
      Top             =   1365
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   2430
      TabIndex        =   0
      Top             =   1365
      Width           =   1215
   End
End
Attribute VB_Name = "frmBCREnvio"
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

Private Sub GeneraArcBCC()
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




Dim oBalance As New NBalanceCont
Dim lnBalanceCate As Integer
Set oBarra = New clsProgressBar
Dim psArchivoAGrabar As String
Dim fs As New Scripting.FileSystemObject
Dim CadImp As String

oBarra.ShowForm Me
oBarra.Max = 7
On Error GoTo ErrGeneraArcBCC
      'Movimientos del Mes
      Set rs = oBalance.LeeBalanceHisto(nTipoBala, nMoneda, Right("0" + Trim(Right(Trim(cboMes.Text), 2)), 2), CInt(txtAnio.Text), , , CInt(txtNroDigitos.Text), 0)
      'DoEvents
   psArchivoAGrabar = App.path & "\SPOOLER\109D0001YY200904.dat"
    Dim ArcSal As Integer
    ArcSal = FreeFile
    Open psArchivoAGrabar For Output As ArcSal
    Print #ArcSal, "109D0001YY200904"
   If Not rs.BOF And Not rs.EOF Then
      oBarra.Max = rs.RecordCount
      Do While Not rs.EOF
        DoEvents
         'Creacion de Reporte
        CadImp = Format(txtAnio.Text, "0000") & Format(Right("0" + Trim(Right(Trim(cboMes.Text), 2)), 2), "00") & "109" & rs!cCtaContCod & String(20 - Len(rs!cCtaContCod), "0")
        CadImp = CadImp & IIf(rs!nSaldoIniImporte >= 0, "+", "") & Replace(Format(rs!nSaldoIniImporte, "0.00"), ".", "") & Space(18 - (" " & Len(CStr(Format(rs!nSaldoIniImporte, "0.00")))))
        CadImp = CadImp & IIf(rs!nDebe >= 0, "+", "") & Replace(Format(rs!nDebe, "0.00"), ".", "") & Space(18 - (" " & Len(CStr(Format(rs!nDebe, "0.00")))))
        CadImp = CadImp & IIf(rs!nHaber >= 0, "+", "") & Replace(Format(rs!nHaber, "0.00"), ".", "") & Space(18 - (" " & Len(CStr(Format(rs!nHaber, "0.00")))))
        CadImp = CadImp & IIf(rs!nSaldoFinImporte >= 0, "+", "") & Replace(Format(rs!nSaldoFinImporte, "0.00"), ".", "") & Space(18 - (" " & Len(CStr(Format(rs!nSaldoFinImporte, "0.00")))))

        CadImp = CadImp & String(27, "0")
        If CadImp <> "" Then
           Print #ArcSal, CadImp
        End If
         ContBarra = ContBarra + 1
         oBarra.Progress rs.Bookmark, "BCR:", "Generando Formato DAT", "Procesando...", vbBlue
         rs.MoveNext
      Loop
   End If
   Close ArcSal
   MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
   RSClose rs
   
   oBarra.CloseForm Me
   Set oBalance = Nothing
   MousePointer = 0
   
   
Exit Sub
ErrGeneraArcBCC:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdProcesar.SetFocus
    End If
End Sub

Private Sub cmdProcesar_Click()
   If ValidaAnio(txtAnio) Then
       Call GeneraArcBCC
   End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
   CentraForm Me
   'frmMdiMain.Enabled = False
   txtAnio = Year(gdFecSis)
   cboMes.ListIndex = Month(gdFecSis) - 1
   txtNroDigitos.Text = "24"
   nTipoBala = 1
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

