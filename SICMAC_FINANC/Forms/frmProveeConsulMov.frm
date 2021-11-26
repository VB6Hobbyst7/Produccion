VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProveeConsulMov 
   Caption         =   "Proveedores: Consulta de Movimientos"
   ClientHeight    =   6690
   ClientLeft      =   1140
   ClientTop       =   1695
   ClientWidth     =   10260
   Icon            =   "frmProveeConsulMov.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10260
   Begin VB.CommandButton cmdPlantilla 
      Caption         =   "&Plantilla"
      Height          =   345
      Left            =   1350
      TabIndex        =   28
      Top             =   5880
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdImpuestos 
      Caption         =   "&Impuestos"
      Height          =   345
      Left            =   120
      TabIndex        =   26
      Top             =   5880
      Width           =   1185
   End
   Begin VB.CommandButton cmdCertificado 
      Caption         =   "&Certificado"
      Height          =   345
      Left            =   2580
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CheckBox chkImp 
      Caption         =   "Detallar Impuestos/Reteneciones  de Documento"
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   5430
      Width           =   4185
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   5790
      ScaleHeight     =   315
      ScaleWidth      =   4335
      TabIndex        =   17
      Top             =   5430
      Width           =   4395
      Begin VB.Label lblTotD 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2820
         TabIndex        =   21
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label lblTotS 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   750
         TabIndex        =   20
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label Label8 
         Caption         =   "Total $"
         Height          =   195
         Left            =   2220
         TabIndex        =   19
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Total S/."
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fechas"
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
      Height          =   1155
      Left            =   7395
      TabIndex        =   12
      Top             =   75
      Width           =   2805
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         TabIndex        =   27
         Top             =   420
         Width           =   1035
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   390
         TabIndex        =   13
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   390
         TabIndex        =   15
         Top             =   690
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Caption         =   "Al"
         Height          =   225
         Left            =   90
         TabIndex        =   16
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Del"
         Height          =   225
         Left            =   90
         TabIndex        =   14
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   7590
      TabIndex        =   4
      Top             =   5910
      Width           =   1185
   End
   Begin MSAdodcLib.Adodc adoProv 
      Height          =   330
      Left            =   240
      Top             =   4890
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   8820
      TabIndex        =   5
      Top             =   5910
      Width           =   1185
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgProv 
      Height          =   4065
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   7170
      _Version        =   393216
      Cols            =   11
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proveedor"
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
      Height          =   1155
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   7215
      Begin Sicmact.TxtBuscar txtCodProv 
         Height          =   345
         Left            =   750
         TabIndex        =   1
         Top             =   240
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
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
         sTitulo         =   ""
         TipoBusPers     =   2
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   690
         Width           =   1935
      End
      Begin VB.TextBox txtProvDNI 
         Height          =   315
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   0
         Tag             =   "txtDocumento"
         Top             =   690
         Width           =   1575
      End
      Begin VB.TextBox txtNomProv 
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "txtNombre"
         Top             =   240
         Width           =   4725
      End
      Begin VB.Label Label3 
         Caption         =   "Documento"
         Height          =   285
         Left            =   4260
         TabIndex        =   22
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "L.E."
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   825
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   285
      Left            =   150
      TabIndex        =   9
      Top             =   5130
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmProveeConsulMov.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   6390
      Visible         =   0   'False
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6315
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14463
            MinWidth        =   14463
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProveeConsulMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rs As New ADODB.Recordset
Dim lsFile  As String
Dim sCtaCod As String
Dim lSalir As Boolean
Dim sImpCta As String
Dim oCon As DConecta

Private Sub cmdCertificado_Click()
Dim K As Integer
Dim nCol As Integer
Dim lsImpre As String
Dim lsImpuesto As String
Dim lsPlantilla As String
Dim nImp As Integer
Dim ldFecha As Date
Dim lsRepo  As String
Dim oPla As NPlantilla
Set oPla = New NPlantilla
lsPlantilla = oPla.GetPlantillaDoc(lsFile)
If lsPlantilla = "" Then
    MsgBox "Es necesario definir planilla de Certificado", vbInformation, "¡Aviso!"
End If
K = 1
If fgProv.TextMatrix(1, 1) = "" Then
    MsgBox "No existen datos", vbInformation, "¡Aviso!"
    Exit Sub
End If

Dim aTotImpProv() As Currency
Dim lsProvNom As String
Dim lsProvRuc As String
Dim oPlan As New clsDocPago

lsRepo = ""
prg.Visible = True
Me.prg.Max = fgProv.Rows - 1
Do While K < fgProv.Rows
    prg.value = K
    ReDim aTotImpProv(fgProv.Cols - 11)
    gnImporte = 0
    gsMovNro = fgProv.TextMatrix(K, 1)
    gdFecha = CDate(fgProv.TextMatrix(K, 4))
    lsProvNom = PstaNombre(fgProv.TextMatrix(K, 8))
    lsProvRuc = fgProv.TextMatrix(K, 10)
    
    
    gsDocNro = ""
    Do While lsProvNom = PstaNombre(fgProv.TextMatrix(K, 8))
        For nCol = 11 To fgProv.Cols - 1
            aTotImpProv(nCol - 10) = aTotImpProv(nCol - 10) + fgProv.TextMatrix(K, nCol)
        Next
        gnImporte = gnImporte + fgProv.TextMatrix(K, 5)
        K = K + 1
        If K >= fgProv.Rows Then
            Exit Do
        End If
    Loop
    Dim lsImpuestoDesc As String
    Dim lnTotImp As Currency
    lsImpuesto = ""
    lnTotImp = 0
    For nCol = 1 To UBound(aTotImpProv, 1)
        For nImp = 1 To frmProveeConsulMovImp.lvImp.ListItems.Count
            If fgProv.TextMatrix(0, nCol + 10) = frmProveeConsulMovImp.lvImp.ListItems(nImp).SubItems(2) Then
                lsImpuestoDesc = frmProveeConsulMovImp.lvImp.ListItems(nImp).SubItems(1)
                Exit For
            End If
        Next
        Linea lsImpuesto, oImpresora.gPrnBoldON & Space(1) & Format(nCol, "0") & " - " & Justifica(lsImpuestoDesc, 38) & ":  " & gcMN & " " & PrnVal(aTotImpProv(nCol), 15, 2) & oImpresora.gPrnBoldOFF
        lnTotImp = lnTotImp + aTotImpProv(nCol)
    Next
    Linea lsImpuesto, oImpresora.gPrnSaltoLinea & oImpresora.gPrnBoldON & Space(1) & "" & Justifica("Importe Total Retenido", 42) & ":  " & gcMN & " " & PrnVal(lnTotImp, 15, 2) & oImpresora.gPrnBoldOFF
    ldFecha = gdFecSis
    lsImpre = Replace(lsPlantilla, "<<BON>>", oImpresora.gPrnBoldON)
    lsImpre = Replace(lsImpre, "<<BOFF>>", oImpresora.gPrnBoldOFF)
    lsImpre = Replace(lsImpre, "<<FechaLarga>>", Format(ldFecha, "dd") & " de " & Format(ldFecha, "mmmm") & " de " & Format(ldFecha, "yyyy"))
    lsImpre = Replace(lsImpre, "<<Impuesto>>", lsImpuesto)
    lsImpre = Replace(lsImpre, "<<gcAnio>>", Str(Year(gdFecha)))
    lsImpre = oPlan.ProcesaPlantilla(lsImpre, True, gsMovNro, ldFecha, gsNomCmac, lsProvNom, gnImporte, gsSimbolo, gsNomCmacRUC, lsProvRuc, gsDocNro, gnColPage * 1.1, gnMgIzq + 4, gnMgDer, "1", 7) & oImpresora.gPrnSaltoPagina
    lsRepo = lsRepo & lsImpre
Loop
Set oPlan = Nothing
prg.Visible = False
EnviaPrevio lsRepo, "Certificados de Proveedores", gnLinPage, False
End Sub

Private Sub cmdImprimir_Click()
Dim N As Integer
Dim nLin As Integer, P As Integer
Dim nTot As Currency
Dim nCol As Integer
Dim lOk As Boolean
If fgProv.TextMatrix(fgProv.row, 1) = "" Then
   MsgBox "No existen elementos que Imprimir...!", vbInformation, "Error"
   Exit Sub
End If
nLin = gnLinPage
rtf = ""
nTot = 0
prg.Min = 0
prg.Max = fgProv.Rows - 1
prg.Visible = True
Me.Enabled = False
Dim sProv As String
Dim sImpre As String, sTexto As String
Dim nTotS As Currency, nTotD As Currency

sProv = "": sImpre = "": sTexto = ""
nTotS = 0: nTotD = 0
Dim aTot() As Currency
ReDim aTot(fgProv.Cols - 11)

For N = 1 To fgProv.Rows - 1
   DoEvents
   prg.value = N
   Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
   Linea sImpre, CabeceraRepo(nLin, P), 0, nLin
   If fgProv.TextMatrix(N, 8) <> sProv Then
      sProv = fgProv.TextMatrix(N, 8)
      If N > 1 Then
         Linea sImpre, oImpresora.gPrnBoldON & oImpresora.gPrnCondensadaON & Space(40) & "TOTAL PROVEEDOR...." & Right(Space(14) & gcMN & " " & Format(nTotS, gsFormatoNumeroView), 14) & "   " & Right(Space(14) & gcME & " " & Format(nTotD, gsFormatoNumeroView), 14) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF, 2, nLin
         nTotS = 0: nTotD = 0
      End If
      Linea sImpre, CabeceraRepo(nLin, P), 0, nLin
      Linea sImpre, oImpresora.gPrnBoldON & " PROVEEDOR : " & sProv & oImpresora.gPrnBoldOFF, , nLin
   End If
   nTotS = nTotS + nVal(fgProv.TextMatrix(N, 5))
   nTotD = nTotD + nVal(fgProv.TextMatrix(N, 6))
   Linea sImpre, CabeceraRepo(nLin, P), 0, nLin
   Linea sImpre, oImpresora.gPrnCondensadaON & " " & Format(N, "000") & " " & Mid(fgProv.TextMatrix(N, 1), 1, 8) & "-" & Mid(fgProv.TextMatrix(N, 1), 9, 6) & " " & Justifica(fgProv.TextMatrix(N, 2), 3) & " " & Justifica(fgProv.TextMatrix(N, 3), 18) & " " & " " & Mid(fgProv.TextMatrix(N, 4), 1, 10) & " " & Right(Space(14) & fgProv.TextMatrix(N, 5), 14) & " " & Right(Space(14) & fgProv.TextMatrix(N, 6), 14) & " " & Mid(Replace(fgProv.TextMatrix(N, 7), Chr(13) & oImpresora.gPrnSaltoLinea, " ") & Space(52), 1, 52) & " " & Justifica(fgProv.TextMatrix(N, 9), 16), 0, nLin
   For nCol = 11 To fgProv.Cols - 1
      Linea sImpre, PrnVal(nVal(fgProv.TextMatrix(N, nCol)), 12, 2) & " ", 0, nLin
      aTot(nCol - 11) = aTot(nCol - 11) + nVal(fgProv.TextMatrix(N, nCol))
   Next
   Linea sImpre, "" & oImpresora.gPrnCondensadaOFF, , nLin
   nTot = nTot + nVal(fgProv.TextMatrix(N, 5))
   If P Mod 10 = 0 Then
      sTexto = sTexto & sImpre
      sImpre = ""
   End If
Next
sImpre = sTexto & sImpre
Status.Panels(1).Text = "Proceso Terminado"
Linea sImpre, oImpresora.gPrnBoldON & oImpresora.gPrnCondensadaON & Space(37) & "TOTAL PROVEEDOR...." & Right(Space(14) & gcMN & " " & Format(nTotS, gsFormatoNumeroView), 14) & " " & Right(Space(14) & gcME & " " & Format(nTotD, gsFormatoNumeroView), 14) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF & oImpresora.gPrnSaltoLinea, , nLin
Linea sImpre, oImpresora.gPrnCondensadaON & " =======================================================================================================================================================", 2, nLin
Linea sImpre, oImpresora.gPrnBoldON & Space(37) & "TOTAL GASTOS......." & Right(Space(14) & gcMN & " " & lblTotS, 14) & " " & Right(Space(14) & gcME & " " & lblTotD, 14) & Space(60), 0, nLin
For nCol = 11 To fgProv.Cols - 1
   Linea sImpre, PrnVal(aTot(nCol - 11), 12, 2) & " ", 0, nLin
Next
Linea sImpre, oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF & oImpresora.gPrnSaltoLinea, , nLin
EnviaPrevio sImpre, "Planilla de Provisiones", gnLinPage, False
prg.Visible = False
Me.Enabled = True
fgProv.SetFocus
End Sub

Private Sub cmdImpuestos_Click()
frmProveeConsulMovImp.Show 1
End Sub

Private Sub cmdPlantilla_Click()
Dim oDPago As clsDocPago
Set oDPago = New clsDocPago
oDPago.InicioCarta "", "", Mid(gsOpeCod, 1, 2) + "1" + Right(gsOpeCod, 3), gsOpeDesc, "", lsFile, 0, gdFecSis, gsNomCmac, gsNomCmacRUC, Me.txtNomProv, Me.txtCodProv, ""
End Sub

Private Sub cmdProcesar_Click()
Dim K As Integer
Dim lvItem As ListItem
If Me.chkImp.value = vbChecked Then
   If frmProveeConsulMovImp Is Nothing Then
      MsgBox "Definir Impuestos/Retenciones para Reporte", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   sImpCta = ""
   For K = 1 To frmProveeConsulMovImp.lvImp.ListItems.Count
      If frmProveeConsulMovImp.lvImp.ListItems.Item(K).Checked Then
         sImpCta = sImpCta & "'" & frmProveeConsulMovImp.lvImp.ListItems.Item(K).Text & "',"
      End If
   Next
   If sImpCta <> "" Then
      sImpCta = Mid(sImpCta, 1, Len(sImpCta) - 1)
   End If
End If
KardexProveedor
FormatoConsulta
fgProv.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
If Not frmProveeConsulMovImp Is Nothing Then
   Unload frmProveeConsulMovImp
End If
End Sub

Private Sub fgProv_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = 2 Then   '   Copiar  [Ctrl+C]
   'KeyUp_Flex fgProv, KeyCode, Shift
End If
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
CentraForm Me
Set oCon = New DConecta
oCon.AbreConexion
lSalir = False
Dim oOpe As New DOperacion
Set rs = oOpe.CargaOpeCtaArbol(gsOpeCod)
CentraForm Me
If rs.EOF Then
   MsgBox "No se asignó Cuenta de Análisis a Consulta", vbCritical, "Error"
   lSalir = True
   Exit Sub
End If
sCtaCod = "'" & rs!cCtaContCod & "'"
Do While Not rs.EOF
   sCtaCod = sCtaCod & ",'" & rs!cCtaContCod & "'"
   rs.MoveNext
Loop
Set oOpe = Nothing
Dim oDoc As New DDocumento
Set rs = oDoc.CargaDocumento()
Do While Not rs.EOF
   cboDoc.AddItem Justifica(rs!nDocTpo, 3) & " " & rs!cDocDesc
   rs.MoveNext
Loop
cboDoc.AddItem "XX " & "Todos"
cboDoc.ListIndex = cboDoc.ListCount - 1
FormatoConsulta
lsFile = gsOpeCod & "0001"
txtCodProv.TipoBusqueda = BuscaPersona
txtCodProv.TipoBusPers = BusPersDocumentoRuc

End Sub

Private Sub FormatoConsulta()
Dim K As Integer
Dim nCol As Integer
fgProv.TextMatrix(0, 1) = "Nro.Movimiento"
fgProv.TextMatrix(0, 2) = "Tpo"
fgProv.TextMatrix(0, 3) = "Doc.Número"
fgProv.TextMatrix(0, 4) = "Doc.Fecha"
fgProv.TextMatrix(0, 5) = "Importe " & gcMN
fgProv.TextMatrix(0, 6) = "Importe " & gcME
fgProv.TextMatrix(0, 7) = "Glosa"
nCol = 10
If Me.chkImp.value = vbChecked Then
   If Not frmProveeConsulMovImp Is Nothing Then
      For K = 1 To frmProveeConsulMovImp.lvImp.ListItems.Count
         If frmProveeConsulMovImp.lvImp.ListItems.Item(K).Checked Then
            nCol = nCol + 1
            fgProv.TextMatrix(0, nCol) = frmProveeConsulMovImp.lvImp.ListItems.Item(K).SubItems(2)
            fgProv.ColWidth(nCol) = 1200
         End If
      Next
   End If
End If

fgProv.ColWidth(0) = 280
fgProv.ColWidth(1) = 2450
fgProv.ColWidth(2) = 500
fgProv.ColWidth(3) = 1200
fgProv.ColWidth(4) = 1100
fgProv.ColWidth(5) = 1200
fgProv.ColWidth(6) = 1200
fgProv.ColWidth(7) = 4500
fgProv.ColWidth(9) = 1500
fgProv.ColAlignment(5) = 7
fgProv.ColAlignment(6) = 7
If txtCodProv <> "" Then
   fgProv.ColWidth(8) = 0
   fgProv.ColWidth(10) = 0
Else
   fgProv.TextMatrix(0, 8) = "Proveedor"
   fgProv.TextMatrix(0, 10) = "RUC"
   fgProv.ColWidth(8) = 3000
   fgProv.ColWidth(10) = 1500
End If
fgProv.TextMatrix(0, 9) = "Cuenta"
fgProv.RowHeight(-1) = 285
End Sub

Private Sub KardexProveedor()
Dim sFecCond As String
Dim K    As Integer
Dim nCol As Integer
Dim sFiltro   As String
Dim sImpuesto As String
sFiltro = ""
If chkImp.value = vbChecked Then
   If Not frmProveeConsulMovImp Is Nothing Then
      For K = 1 To frmProveeConsulMovImp.lvImp.ListItems.Count
         If frmProveeConsulMovImp.lvImp.ListItems.Item(K).Checked Then
            sFiltro = sFiltro & ", SUM(ISNULL(CASE WHEN imp.cCtaContCod = '" & frmProveeConsulMovImp.lvImp.ListItems.Item(K).Text & "' THEN nImporte END,0)) Imp" & K & " "
         End If
      Next
      If sFiltro <> "" Then
         sImpuesto = "LEFT JOIN (SELECT nMovNro, cmovotrovariable cCtaContCod, Abs(Sum(nMovOtroImporte)) nImporte " _
                   & "FROM  movotrositem where cmovotrovariable in (" & sImpCta & ") " _
                   & "GROUP BY nMovNro, cmovotrovariable " _
                   & "Union " _
                   & "SELECT nMovNro, cCtaContCod, Abs(Sum(nMovImporte)) nImporte " _
                   & "FROM  movcta where cCtaContCod in (" & sImpCta & ") " _
                   & "GROUP BY nMovNro, cCtaContCod " _
                   & ") Imp ON Imp.nMovNro = a.nMovNro"
      End If
   End If
End If

If Trim(txtFecha) <> "/  /" And Trim(txtFecha2) <> "/  /" Then
   sFecCond = " and substring(b.cMovNro,1,8) Between '" & Format(txtFecha, "yyyymmdd") & "' and '" & Format(txtFecha2, "yyyymmdd") & "' "
ElseIf Trim(txtFecha) <> "/  /" Then
    sFecCond = " and substring(b.cMovNro,1,8) >= '" & Format(txtFecha, "yyyymmdd") & "' "
ElseIf Trim(txtFecha2) <> "/  /" Then
    sFecCond = " and substring(b.cMovNro,1,8) <= '" & Format(txtFecha2, "yyyymmdd") & "' "
End If
sSql = "SELECT a.nMovNro, Doc.cDocAbrev, e.cDocNro, Convert(varchar(10),e.dDocFecha,103) dDocFecha , " _
     & "       STR(c.nMovImporte,16,2) nMovImporteD, " _
     & "       ISNULL(STR(me.nMovMEImporte,16,2),'') nMovImporteME, " _
     & "       b.cMovDesc,P.cPersNombre cNomPers, c.cCtaContCod, ISNULL(pid.cPersIDnro,'') cDocPers " & sFiltro _
     & "FROM Mov b JOIN MovGasto a ON b.nMovNro = a.nMovNro " _
     & "           JOIN MovCta c ON c.nMovNro= a.nMovNro " _
     & "           JOIN Persona P ON P.cPersCod = a.cPersCod LEFT JOIN PersID pid ON pid.cPersCod = p.cPersCod and cPersIDTpo = " & gPersIdRUC & " " _
     & "      LEFT JOIN MovMe me ON me.nMovNro = c.nMovNro and me.nMovItem = c.nMovItem " & sImpuesto & ", " _
     & "     MovDoc e JOIN Documento Doc ON Doc.nDocTpo = e.nDocTpo " _
     & "WHERE b.nMovEstado = '10' and not c.cctacontcod in ('" & gcConvMED & "') and not b.nMovFlag IN ('1','5') " & IIf(txtCodProv.psCodigoPersona = "", "", " and a.cPersCod = '" & "" & Trim(txtCodProv.psCodigoPersona) & "'") & IIf(Left(cboDoc, 2) = "XX", "", " and e.nDocTpo = " & Left(cboDoc, 2) & " ") _
     & "      and e.nMovNro = a.nMovNro " _
     & "      and c.nMovImporte > 0 and NOT cOpeCod like '50_30%' " & sFecCond _
     & IIf(sFiltro = "", "", "GROUP BY a.nMovNro, Doc.cDocAbrev, e.cDocNro, Convert(varchar(10),e.dDocFecha,103), STR(c.nMovImporte,16,2), ISNULL(STR(me.nMovMEImporte,16,2),''), b.cMovDesc , P.cPersNombre, c.cCtaContCod, ISNULL(pid.cPersIDnro,'') ") _
     & " ORDER BY cNomPers, a.nMovNro"

Status.Panels(1).Text = "Procesando Datos..."

'VAPI COMENTADO POR VAPI SEGUN ERS004-2015
'adoProv.ConnectionString = oCon.ConexionActiva.ConnectionString
'adoProv.RecordSource = sSql
'adoProv.Refresh
'FIN VAPI ***********************************

'VAPI AGREGADO POR POR VAPI SEGUN ERS004-2015
Dim oDConecta As DConecta
Set oDConecta = New DConecta

Dim RsSource  As ADODB.Recordset
Set RsSource = New ADODB.Recordset

If oDConecta.AbreConexion = False Then Exit Sub

Set RsSource = oDConecta.CargaRecordSet(sSql)
oDConecta.CierraConexion
Set oDConecta = Nothing

'FIN VAPI ***********************************

'Set fgProv.DataSource = adoProv  VAPI COMENTADO POR VAPI SEGUN ERS004-2015

Set fgProv.DataSource = RsSource 'VAPI AGREGADO POR POR VAPI SEGUN ERS004-2015


If fgProv.Rows > 10 Then
   fgProv.row = fgProv.Rows - 1
   fgProv.TopRow = fgProv.Rows - 1
End If
Dim nTotalS As Currency, nTotalD As Currency, nTotalImp As Currency

nTotalS = 0: nTotalD = 0
For K = 1 To fgProv.Rows - 1
   nTotalS = nTotalS + Val(fgProv.TextMatrix(K, 5))
   nTotalD = nTotalD + Val(fgProv.TextMatrix(K, 6))

   fgProv.TextMatrix(K, 5) = Format(fgProv.TextMatrix(K, 5), gsFormatoNumeroView)
   fgProv.TextMatrix(K, 6) = Format(fgProv.TextMatrix(K, 6), gsFormatoNumeroView)
   If fgProv.Cols > 11 Then
      For nCol = 11 To fgProv.Cols - 1
         fgProv.TextMatrix(K, nCol) = Format(fgProv.TextMatrix(K, nCol), gsFormatoNumeroView)
      Next
   End If
Next
lblTotS.Caption = Format(nTotalS, gsFormatoNumeroView)
lblTotD.Caption = Format(nTotalD, gsFormatoNumeroView)

Status.Panels(1).Text = "Proceso Terminado..."
End Sub

Private Function CabeceraRepo(ByRef nLin As Integer, ByRef P As Integer) As String
Dim sTit As String
Dim sImpre As String
Dim sCab1  As String
Dim nCol   As Integer
Dim K As Integer
sTit = " CONSULTA DE MOVIMIENTOS POR PROVEEDOR "
If nLin > gnLinPage - 3 Then
   If P > 0 Then sImpre = sImpre & oImpresora.gPrnSaltoPagina
   P = P + 1
   nLin = 1
   Linea sImpre, Mid(gsInstCmac & Space(20), 1, 20) & Space(42) & gdFecSis & " - " & Format(Time, "hh:mm:ss"), , nLin
   Linea sImpre, Space(72) & "Pag. " & Format(P, "000"), , nLin
   Linea sImpre, oImpresora.gPrnBoldON & Centra(sTit, gnColPage), , nLin
   Linea sImpre, oImpresora.gPrnCondensadaON, , nLin
   sCab1 = "  "
   For K = 11 To fgProv.Cols - 1
      sCab1 = " " & sCab1 & Centra(fgProv.TextMatrix(0, K), 12)
   Next
   Linea sImpre, " =========================================================================================================================================================" & String((fgProv.Cols - 11) * 13, "="), , nLin
   Linea sImpre, " Item Nro de           Comprobante            Fecha de           IMPORTE      IMPORTE                 C O N C E P T O                             Cuenta " & sCab1, , nLin
   Linea sImpre, "      Movimiento       Tpo   Número           Emisión              M.N.         M.E.                                                             Contable", , nLin
   Linea sImpre, " ---------------------------------------------------------------------------------------------------------------------------------------------------------" & String((fgProv.Cols - 11) * 13, "-") & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF, , nLin
End If
CabeceraRepo = sImpre
End Function

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub txtCodProv_EmiteDatos()
txtNomProv = txtCodProv.psDescripcion
txtCodProv.Tag = txtCodProv.psCodigoPersona
txtProvDNI = Trim(txtCodProv.sPersNroDoc)
If txtNomProv <> "" Then
   txtFecha.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(txtFecha) = "/  /" Then
   Else
      If ValidaFecha(txtFecha) <> "" Then
         MsgBox "Fecha no Valida...!"
         Exit Sub
      End If
   End If
   txtFecha2.SetFocus
End If
End Sub

Private Sub txtFecha2_GotFocus()
fEnfoque txtFecha2
End Sub

Private Sub txtFecha2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(txtFecha2) = "/  /" Then
   Else
      If ValidaFecha(txtFecha2) <> "" Then
         MsgBox "Fecha no Valida...!"
         Exit Sub
      End If
   End If
   cmdProcesar.SetFocus
End If
End Sub

