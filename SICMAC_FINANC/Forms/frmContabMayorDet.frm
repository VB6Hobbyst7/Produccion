VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmContabMayorDet 
   Caption         =   "Análisis de Cuentas Contables"
   ClientHeight    =   7920
   ClientLeft      =   930
   ClientTop       =   645
   ClientWidth     =   10230
   Icon            =   "frmContabMayorDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEExcel 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   6000
      TabIndex        =   26
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CheckBox chkPersona 
      Caption         =   "Persona"
      Height          =   210
      Left            =   165
      TabIndex        =   21
      Top             =   6360
      Width           =   2625
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   4785
      Left            =   90
      TabIndex        =   6
      Top             =   1410
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   8440
      Cols0           =   13
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Movimiento-DEBE-HABER-SALDO-Concepto-Tipo-Número-Fecha-DEBE ME-HABER ME-Persona Gasto-Persona A Rendir"
      EncabezadosAnchos=   "300-2400-1200-1200-1600-4500-470-1200-1000-1200-1200-3000-3000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-R-R-R-L-C-L-C-C-C-L-L"
      FormatosEdit    =   "0-0-2-2-2-0-0-0-0-2-2-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8700
      TabIndex        =   11
      Top             =   7440
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Frame Frame1 
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
      Height          =   1245
      Left            =   90
      TabIndex        =   12
      Top             =   60
      Width           =   10005
      Begin VB.CheckBox chkHistorica 
         Caption         =   "Histórica"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin Sicmact.TxtBuscar txtCtaCod 
         Height          =   360
         Left            =   1740
         TabIndex        =   0
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   635
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
      End
      Begin VB.ComboBox cboFiltro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmContabMayorDet.frx":030A
         Left            =   7380
         List            =   "frmContabMayorDet.frx":0320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   765
         Width           =   690
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5700
         MaxLength       =   16
         TabIndex        =   3
         Top             =   765
         Width           =   1665
      End
      Begin VB.CommandButton cmdProcesar 
         BackColor       =   &H80000016&
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
         Height          =   330
         Left            =   8310
         TabIndex        =   5
         Top             =   750
         Width           =   1470
      End
      Begin VB.TextBox txtCtaDesc 
         Enabled         =   0   'False
         Height          =   345
         Left            =   3750
         TabIndex        =   10
         Top             =   270
         Width           =   6015
      End
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   345
         Left            =   1740
         TabIndex        =   1
         Top             =   735
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   345
         Left            =   3360
         TabIndex        =   2
         Top             =   735
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         Left            =   4920
         TabIndex        =   20
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "AL"
         Height          =   195
         Left            =   3030
         TabIndex        =   15
         Top             =   810
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DEL"
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   810
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
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
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   2970
      TabIndex        =   17
      Top             =   6180
      Width           =   7125
      Begin VB.TextBox txtSaldoIniME 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtSaldoFinalME 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtSaldoFin 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   150
         Width           =   1875
      End
      Begin VB.TextBox txtSaldoIni 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   150
         Width           =   1875
      End
      Begin VB.Label Label8 
         Caption         =   "SALDO INICIAL          ME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   25
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label9 
         Caption         =   "SALDO FINAL             ME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3690
         TabIndex        =   24
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label Label6 
         Caption         =   "SALDO FINAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3690
         TabIndex        =   19
         Top             =   210
         Width           =   1425
      End
      Begin VB.Label Label5 
         Caption         =   "SALDO INICIAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   210
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmContabMayorDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim sCtaCod As String, sCtaDesc As String
Dim nItem As Integer, nLin As Integer, P As Integer
Dim sSql As String
Dim sTipoCta As String
Dim fsNombreCtaCont As String
Dim oBarra   As New clsProgressBar
Dim WithEvents oImp     As NContImprimir
Attribute oImp.VB_VarHelpID = -1

Private Sub cboFiltro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdProcesar.SetFocus
End If
End Sub
'EJVG20130304 ***
Private Sub chkHistorica_Click()
    Dim oCont As New DCtaCont
    
    txtCtaCod.Text = ""
    txtCtaDesc.Text = ""
    If Me.chkHistorica.value = 1 Then
        fsNombreCtaCont = "CtaContBackup20121231"
    Else
        fsNombreCtaCont = "CtaCont"
    End If
    txtCtaCod.rs = oCont.CargaCtaCont("", fsNombreCtaCont)
    txtCtaCod.TipoBusqueda = BuscaGrid
    txtCtaCod.lbUltimaInstancia = False
    Set oCont = Nothing
End Sub
'END EJVG *******
Private Sub cmdEExcel_Click()

    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim rs      As ADODB.Recordset
    Dim rsCta   As ADODB.Recordset
    Dim oCont As New NContAsientos
    Dim oSdo  As New NCtasaldo
    Set oSdo = Nothing
    Set oImp = New NContImprimir
    Dim nDebe  As Currency, nHaber  As Currency
    Dim nDebeD As Currency, nHaberH As Currency
    Dim nHaberD As Currency
    Dim nSaldo  As Currency, nSaldoIni As Currency
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Integer
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AnalisisDeCuenta"
    lsNomHoja = "AnalisisCuenta"
    nLin = 0
    lsArchivo1 = "\spooler\AnalisisDeCuenta" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    xlHoja1.Cells(4, 1) = "Del " & Format(txtFechaDel.Text, "DD/MM/YYYY") & " Al " & Format(txtFechaAl.Text, "DD/MM/YYYY")
    xlHoja1.Cells(1, 4) = Format(gdFecSis, "DD/MM/YYYY") & " " & Format$(Time(), "HH:MM:SS")
    xlHoja1.Cells(11, 2) = txtCtaCod.Text
    xlHoja1.Cells(11, 3) = txtCtaDesc.Text
    
    nLin = gnLinPage
    nSaltoContador = 12
    P = 0
    
    Set rsCta = oCont.GetMayorCuenta(txtCtaCod, Format(txtFechaDel, "yyyymmdd"), Format(txtFechaAl, "yyyymmdd"), nVal(txtImporte), cboFiltro, , , IIf(Me.chkPersona.value = 1, True, False))

    Set oCont = Nothing
    If rsCta.EOF Then

            xlHoja1.Cells(nSaltoContador, 4) = "SALDO INICIAL : "
            xlHoja1.Cells(nSaltoContador, 5) = PrnVal(nSaldo, 16, 2, False)
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
            nSaltoContador = nSaltoContador + 1
            xlHoja1.Cells(nSaltoContador, 4) = "SALDO FINAL : "
            xlHoja1.Cells(nSaltoContador, 5) = PrnVal(nSaldo, 16, 2, False)
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
            nSaltoContador = nSaltoContador + 1
        Set oCont = Nothing
        Exit Sub
    End If
    Dim oCta  As New DCtaCont

    Set rs = oCta.CargaCtaContClase(txtCtaCod)
    Set oCta = Nothing
    If Not rs.EOF Then
        sTipoCta = Trim(rs!cCtaCaracter)
    Else
        sTipoCta = "D"
    End If

    nLin = nLin + 2
    nDebe = 0: nHaber = 0
   ' N = 1

    sFecha = Mid(rsCta!cMovNro, 1, 8)
    nSaldo = oSdo.GetCtaSaldo(rsCta!cCtaContCod, Format(CDate(txtFechaDel) - 1, gsFormatoFecha))
    nSaldoIni = nSaldo
    Do While Not rsCta.EOF
        DoEvents

        sMov = rsCta!cMovNro
        If sFecha <> Mid(sMov, 1, 8) Then
            xlHoja1.Cells(nSaltoContador, 4) = "TOTAL DIARIO  : "
            xlHoja1.Cells(nSaltoContador, 5) = PrnVal(nDebeD, 16, 2, False)
            xlHoja1.Cells(nSaltoContador, 6) = PrnVal(nHaberD, 16, 2, False)
            xlHoja1.Cells(nSaltoContador, 7) = PrnVal(nSaldo, 16, 2)
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
            nSaltoContador = nSaltoContador + 1
            nDebeD = 0: nHaberD = 0
            sFecha = Mid(rsCta!cMovNro, 1, 8)
                        
        End If
        If sTipoCta = "D" Then
            nSaldo = nSaldo + rsCta!nDebe - rsCta!nHaber
        Else
            nSaldo = nSaldo + rsCta!nHaber - rsCta!nDebe
        End If
        sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " "
        sDocFecha = Justifica(rsCta!dDocFecha, 10)
        sMov = rsCta!cMovNro
   
        xlHoja1.Cells(nSaltoContador, 1) = Mid(sMov, 1, 8) & "-" & Mid(sMov, 9, 6) & " " & Right(RTrim(sMov), 4)
        xlHoja1.Cells(nSaltoContador, 2) = sDoc
        xlHoja1.Cells(nSaltoContador, 3) = sDocFecha
        xlHoja1.Cells(nSaltoContador, 4) = rsCta!cMovDesc
        xlHoja1.Cells(nSaltoContador, 5) = PrnVal(rsCta!nDebe, 16, 2, False)
        xlHoja1.Cells(nSaltoContador, 6) = PrnVal(rsCta!nHaber, 16, 2, False)
        xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
        nLin = nLin + 1
        nSaltoContador = nSaltoContador + 1
        nDebeD = nDebeD + rsCta!nDebe
        nDebe = nDebe + rsCta!nDebe
        nHaberD = nHaberD + rsCta!nHaber
        nHaber = nHaber + rsCta!nHaber
        rsCta.MoveNext
        If rsCta.EOF Then
           Exit Do
        End If
        Do While sMov = rsCta!cMovNro
            sDoc = Justifica(rsCta!cDocAbrev, 3) & " " & Justifica(rsCta!cDocNro, 20) & " "
            sDocFecha = rsCta!dDocFecha
      
            xlHoja1.Cells(nSaltoContador, 2) = IIf(Trim(sDoc) = "", "", sDoc)
            xlHoja1.Cells(nSaltoContador, 3) = sDocFecha
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
            nSaltoContador = nSaltoContador + 1
            rsCta.MoveNext
            If rsCta.EOF Then
                Exit Do
            End If
        Loop
        If rsCta.EOF Then
            Exit Do
        End If
    Loop
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Font.Bold = True
    xlHoja1.Cells(nSaltoContador, 4) = "TOTAL DIARIO  : "
    xlHoja1.Cells(nSaltoContador, 5) = PrnVal(nDebeD, 16, 2, False)
    xlHoja1.Cells(nSaltoContador, 6) = PrnVal(nHaberD, 16, 2, False)
    xlHoja1.Cells(nSaltoContador, 7) = PrnVal(nSaldo, 16, 2)
    nSaltoContador = nSaltoContador + 1

    nLin = nLin + 2
    N = N + 1
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Font.Bold = True
    xlHoja1.Cells(nSaltoContador, 4) = "TOTAL CUENTA  : "
    xlHoja1.Cells(nSaltoContador, 5) = PrnVal(nDebe, 16, 2, False)
    xlHoja1.Cells(nSaltoContador, 6) = PrnVal(nHaber, 16, 2, False)
    nSaltoContador = nSaltoContador + 1
    
  

    nLin = nLin + 2
    
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Font.Bold = True
    xlHoja1.Cells(nSaltoContador, 4) = "SALDO INICIAL : "
    xlHoja1.Cells(nSaltoContador, 5) = PrnVal(nSaldoIni, 16, 2, False)
    xlHoja1.Cells(nSaltoContador, 6) = "SALDO INICIAL ME :"
    xlHoja1.Cells(nSaltoContador, 7) = PrnVal(IIf(txtSaldoIniME = "", 0, txtSaldoIniME), 16, 2)
    nSaltoContador = nSaltoContador + 1

    xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Font.Bold = True
    xlHoja1.Cells(nSaltoContador, 4) = "SALDO FINAL : "
    xlHoja1.Cells(nSaltoContador, 5) = PrnVal(nSaldo, 16, 2, False)
    xlHoja1.Cells(nSaltoContador, 6) = "SALDO FINAL ME :"
    xlHoja1.Cells(nSaltoContador, 7) = PrnVal(IIf(txtSaldoFinalME = "", 0, txtSaldoFinalME), 16, 2)
    nSaltoContador = nSaltoContador + 1
        
    nLin = pnLinPage + 1
    
    rs.Close: Set rs = Nothing
    rsCta.Close: Set rsCta = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    Set oImp = Nothing
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Sub cmdProcesar_Click()
Dim rsCta As New ADODB.Recordset
Dim sCond As String
Dim sDoc As String
Dim nSaldo As Currency
Dim sMov As String
Dim sFecha As String
Dim nRow As Integer
Dim nDebe  As Currency, nHaber  As Currency
Dim nDebeD As Currency, nHaberD As Currency
Dim nSaldoInicialME As Currency 'ALPA 20130507
'John ****
Dim nDebeME  As Currency, nHaberME  As Currency
Dim nDebeDME As Currency, nHaberDME As Currency
'****
Dim oDBalanceContA As DbalanceCont
Set oDBalanceContA = New DbalanceCont

On Error GoTo ErrImprime
txtSaldoIni.Text = ""
txtSaldoFin.Text = ""
txtSaldoIniME.Text = ""
txtSaldoFinalME.Text = ""

If Not Me.Enabled Then
   Exit Sub
End If

nItem = 0
If txtCtaCod = "" Then
   MsgBox "Falta indicar Cuenta Contable...", vbInformation, "Aviso"
   txtCtaCod.SetFocus
   Exit Sub
End If
If CDate(txtFechaDel) > CDate(txtFechaAl) Then
   MsgBox "Fecha Inicial debe ser menor o igual que fecha final.", vbInformation, "Aviso"
   txtFechaDel.SetFocus
   Exit Sub
End If

Me.Enabled = False
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = 1
oBarra.Progress 0, "Mayor de Cuenta Contable : " & txtCtaCod, "Cargando datos...", , vbBlue
DoEvents
Dim oSdo  As New NCtasaldo
Dim oCta  As New DCtaCont
Dim oCont As New NContAsientos

'Cambio en dolares
Dim nSaldoME As Currency

'*********** John  JEOM  *************
Dim bME As Boolean
If Mid(txtCtaCod.Text, 3, 1) = 1 Then
   bME = False
Else
   bME = True
End If
'Set rsCta = oCont.GetMayorCuenta(txtCtaCod.Text, Format(txtFechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), nVal(txtImporte), cboFiltro)
'Set rsCta = oCont.GetMayorCuenta(txtCtaCod.Text, Format(txtFechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), nVal(txtImporte), cboFiltro, bME)
Set rsCta = oCont.GetMayorCuenta(txtCtaCod.Text, Format(txtFechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), nVal(txtImporte), cboFiltro, bME, , , fsNombreCtaCont)
'*****************Fin Cambio************

Set oCont = Nothing
fg.Rows = 2
fg.Clear
fg.FormaCabecera
If rsCta.EOF Then
   fg.Rows = 2
   fg.Row = 1
   nSaldo = oSdo.GetCtaSaldo(txtCtaCod, Format(CDate(txtFechaDel) - 1, gsFormatoFecha))
   txtSaldoIni = PrnVal(nSaldo, 16, 2)
   txtSaldoFin = PrnVal(nSaldo, 16, 2)
   If txtCtaCod = "2918090101" Then
    Call oDBalanceContA.InsertaCtaContSaldoDiario("Saldo_1_1", CDate(txtFechaAl), "763406", CDbl(txtSaldoFin))
   End If


   'John********
   If bME = True Then
      'ALPA 20130507***************************************************************************************
      'nSaldoME = oSdo.GetCtaSaldo(txtCtaCod & "%", Format(CDate(txtFechaDel) - 1, gsFormatoFecha), False)
      nSaldoME = oSdo.GetCtaSaldoBalanceME(txtCtaCod & "%", nSaldoInicialME, Format(CDate(txtFechaDel), "YYYY/MM/DD"), Format(CDate(txtFechaAl), "YYYY/MM/DD"), "2", nVal(txtImporte), cboFiltro, bME, , , fsNombreCtaCont)
      'MIOL 20130802 Se cambio gsFormatoFecha por "YYYY/MM/DD"
      '****************************************************************************************************
      If txtSaldoIni.Text = 0 Then
        txtSaldoIniME = PrnVal(0, 16, 2)
        txtSaldoFinalME = PrnVal(0, 16, 2)
      Else
        txtSaldoIniME = PrnVal(nSaldoME, 16, 2)
        txtSaldoFinalME = PrnVal(nSaldoME, 16, 2)
      End If
      If txtCtaCod = "2928090101" Then
      Call oDBalanceContA.InsertaCtaContSaldoDiario("Saldo_2_1", CDate(txtFechaAl), "763406", CDbl(txtSaldoFinalME))
      End If
   End If
   '***********
   
   oBarra.CloseForm Me
   MousePointer = 0
   Me.Enabled = True
   Exit Sub
End If
Set rs = oCta.CargaCtaContClase(txtCtaCod)
Set oCta = Nothing
If Not rs.EOF Then
   sTipoCta = Trim(rs!cCtaCaracter)
Else
   sTipoCta = "D"
End If

fg.BackColorControl = vbBlue
sCtaCod = txtCtaCod

nSaldo = oSdo.GetCtaSaldo(sCtaCod & "%", Format(CDate(txtFechaDel) - 1, gsFormatoFecha))
'John********
If bME = True Then
'ALPA 20120419***************************
   If nSaldo = 0 Then
        nSaldoME = 0
   Else
        'ALPA 20130507**********************************************************************************************************************
        'nSaldoME = oSdo.GetCtaSaldo(sCtaCod & "%", Format(CDate(txtFechaDel) - 1, gsFormatoFecha), False)
         nSaldoME = oSdo.GetCtaSaldoBalanceME(txtCtaCod & "%", nSaldoInicialME, Format(CDate(txtFechaDel), "YYYY/MM/DD"), Format(CDate(txtFechaAl), "YYYY/MM/DD"), "2", nVal(txtImporte), cboFiltro, bME, , , fsNombreCtaCont, sTipoCta)
        'MIOL 20130802 Se cambio gsFormatoFecha por "YYYY/MM/DD"
        '***********************************************************************************************************************************
   End If
End If
'****************************************
'---Fin
txtSaldoIni = PrnVal(nSaldo, 16, 2)
'John*****
If bME = True Then
   txtSaldoIniME = PrnVal(nSaldoME, 16, 2)
End If
'**Fin
Set oSdo = Nothing
oBarra.Max = rsCta.RecordCount

Dim rsMovViat As ADODB.Recordset
Dim rsMovViat2 As ADODB.Recordset
Dim lnNombreGasto As String
Dim lnNroRef As Long
Dim lnNombreGasto1 As String

sFecha = Mid(rsCta!cMovNro, 1, 8)
Do While Not rsCta.EOF
   DoEvents
   oBarra.Progress rsCta.Bookmark, "Mayor de Cuenta Contable : " & sCtaCod, "", "Procesando... ", vbBlue
   sMov = rsCta!cMovNro
   
   'JEOM
   Set rsMovViat = New ADODB.Recordset
   Set rsMovViat = oCont.GetMovPersonaARendir(rsCta!nMovNro)
   If Not rsMovViat.EOF And Not rsMovViat.BOF Then
        lnNombreGasto = rsMovViat!cPersNombre
        lnNroRef = rsMovViat!nMovNroRef
    Else
        lnNombreGasto = ""
        lnNroRef = 0
   End If
   RSClose rsMovViat
   If lnNroRef = 0 Then
      lnNombreGasto1 = ""
   Else
      Set rsMovViat2 = New ADODB.Recordset
      Set rsMovViat2 = oCont.GetMovPersonaARendirRef(lnNroRef)
      If Not rsMovViat2.EOF And Not rsMovViat2.BOF Then
        lnNombreGasto1 = rsMovViat2!cPersNombre
      Else
        lnNombreGasto1 = ""
      End If
      RSClose rsMovViat2
    End If
   'FIN
   
   If sFecha <> Mid(sMov, 1, 8) Then
      fg.AdicionaFila
      nRow = fg.Row
      fg.TextMatrix(nRow, 0) = ""
      fg.TextMatrix(nRow, 1) = "TOTAL DIA " & sFecha
      fg.TextMatrix(nRow, 2) = PrnVal(nDebeD, 16, 2)
      fg.TextMatrix(nRow, 3) = PrnVal(nHaberD, 16, 2)
      fg.TextMatrix(nRow, 4) = PrnVal(nSaldo, 16, 2)
      nDebeD = 0: nHaberD = 0
      sFecha = Mid(rsCta!cMovNro, 1, 8)
      fg.BackColorRow "&H00E0E0E0", True
   End If
   If sTipoCta = "D" Then
      nSaldo = nSaldo + rsCta!nDebe - rsCta!nHaber
      If bME = True Then
        nSaldoME = nSaldoME + rsCta!nDebeME - rsCta!nHaberME
      End If
   Else
      nSaldo = nSaldo + rsCta!nHaber - rsCta!nDebe
      If bME = True Then
        nSaldoME = nSaldoME + rsCta!nHaberME - rsCta!nDebeME
      End If
   End If
   fg.AdicionaFila
   nRow = fg.Row
   fg.TextMatrix(nRow, 0) = ""
   fg.TextMatrix(nRow, 1) = rsCta!cMovNro
   fg.TextMatrix(nRow, 2) = PrnVal(rsCta!nDebe, 16, 2)
   fg.TextMatrix(nRow, 3) = PrnVal(rsCta!nHaber, 16, 2)
   fg.TextMatrix(nRow, 4) = PrnVal(nSaldo, 16, 2)
   fg.TextMatrix(nRow, 5) = rsCta!cMovDesc
   fg.TextMatrix(nRow, 6) = rsCta!cDocAbrev
   fg.TextMatrix(nRow, 7) = rsCta!cDocNro
   fg.TextMatrix(nRow, 8) = rsCta!dDocFecha
   'JOHN
   fg.TextMatrix(nRow, 11) = lnNombreGasto
   fg.TextMatrix(nRow, 12) = lnNombreGasto1
   'FIN
   'John
   If bME = True Then
      fg.TextMatrix(nRow, 9) = PrnVal(rsCta!nDebeME, 16, 2)
      fg.TextMatrix(nRow, 10) = PrnVal(rsCta!nHaberME, 16, 2)
   End If
   'fg.TextMatrix(nRow, 11) = rsCta!cPersNombreGasto Modificado mientras se coordina GITU
   'fg.TextMatrix(nRow, 12) = rsCta!cPersNombreRendir Modificado mientras se coordina GITU
   '******
   nDebeD = nDebeD + rsCta!nDebe
   nDebe = nDebe + rsCta!nDebe
   nHaberD = nHaberD + rsCta!nHaber
   nHaber = nHaber + rsCta!nHaber
   
   'John*******
   If bME = True Then
      nDebeDME = nDebeDME + rsCta!nDebeME
      nDebeME = nDebeME + rsCta!nDebeME
      nHaberDME = nHaberDME + rsCta!nHaberME
      nHaberME = nHaberME + rsCta!nHaberME
   End If
   '***********
   rsCta.MoveNext
   If rsCta.EOF Then
      Exit Do
   End If
   Do While sMov = rsCta!cMovNro
      fg.AdicionaFila
      nRow = fg.Row
      fg.TextMatrix(nRow, 0) = ""
      fg.TextMatrix(nRow, 1) = rsCta!cMovNro
      fg.TextMatrix(nRow, 6) = rsCta!cDocAbrev
      fg.TextMatrix(nRow, 7) = rsCta!cDocNro
      fg.TextMatrix(nRow, 8) = rsCta!dDocFecha
      'John ********
      If bME = True Then
         fg.TextMatrix(nRow, 11) = PrnVal(rsCta!nDebeME, 16, 2)
         fg.TextMatrix(nRow, 12) = PrnVal(rsCta!nHaberME, 16, 2)
      End If
      '*************
      rsCta.MoveNext
      If rsCta.EOF Then
         Exit Do
      End If
   Loop
   If rsCta.EOF Then
      Exit Do
   End If
Loop
fg.AdicionaFila
nRow = fg.Row
fg.TextMatrix(nRow, 0) = ""
fg.TextMatrix(nRow, 1) = "TOTAL DIA " & sFecha
fg.TextMatrix(nRow, 2) = PrnVal(nDebeD, 16, 2)
fg.TextMatrix(nRow, 3) = PrnVal(nHaberD, 16, 2)
fg.TextMatrix(nRow, 4) = PrnVal(nSaldo, 16, 2)
'John **********
If bME = True Then
   fg.TextMatrix(nRow, 9) = PrnVal(nDebeDME, 16, 2)
   fg.TextMatrix(nRow, 10) = PrnVal(nHaberDME, 16, 2)
End If
'***************
nDebeD = 0: nHaberD = 0
'John******
nDebeDME = 0: nHaberDME = 0
'**********
fg.BackColorRow "&H00E0E0E0", True

fg.AdicionaFila
nRow = fg.Row
fg.TextMatrix(nRow, 0) = ""
fg.TextMatrix(nRow, 1) = "TOTAL MAYOR  "
fg.TextMatrix(nRow, 2) = PrnVal(nDebe, 16, 2)
fg.TextMatrix(nRow, 3) = PrnVal(nHaber, 16, 2)
fg.BackColorRow "&H00FFFFC0", True
'John ********
If bME = True Then
   fg.TextMatrix(nRow, 9) = PrnVal(nDebeME, 16, 2)
   fg.TextMatrix(nRow, 10) = PrnVal(nHaberME, 16, 2)
End If
'*************

oBarra.CloseForm Me
MousePointer = 0
txtSaldoFin = PrnVal(nSaldo, 16, 2)
If txtCtaCod = "2918090101" Then
    Call oDBalanceContA.InsertaCtaContSaldoDiario("Saldo_1_1", CDate(txtFechaAl), "763406", CDbl(txtSaldoFin))
End If
'John *****
If bME = True Then
   txtSaldoFinalME = PrnVal(nSaldoME, 16, 2)
   If txtCtaCod = "2928090101" Then
        Call oDBalanceContA.InsertaCtaContSaldoDiario("Saldo_2_1", CDate(txtFechaAl), "763406", CDbl(txtSaldoFin))
   End If
End If
'********************
Me.Enabled = True
RSClose rs
RSClose rsCta
fg.SetFocus
Exit Sub
ErrImprime:
 MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdImprimir_Click()
Dim sTexto As String
Set oImp = New NContImprimir
Me.Enabled = False
sTexto = oImp.ImprimeMayorCta(txtCtaCod, txtCtaDesc, CDate(txtFechaDel), CDate(txtFechaAl), nVal(txtImporte), nVal(txtSaldoIniME), nVal(txtSaldoFinalME), cboFiltro, gnLinPage, IIf(Me.chkPersona.value = 1, True, False))
EnviaPrevio sTexto, "Mayor de Cuentas", gnLinPage, False
Me.Enabled = True
Set oImp = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmReportes.Enabled = False
CentraForm Me
txtFechaDel = Format(gdFecSis, "dd/mm/yyyy")
txtFechaAl = Format(gdFecSis, "dd/mm/yyyy")
fsNombreCtaCont = "CtaCont" 'EJVG20130304
Dim oCont As New DCtaCont
'txtCtaCod.rs = oCont.CargaCtaCont("", "CtaCont")
txtCtaCod.rs = oCont.CargaCtaCont("", fsNombreCtaCont) 'EJVG20130304
txtCtaCod.TipoBusqueda = BuscaGrid
txtCtaCod.lbUltimaInstancia = False
Set oCont = Nothing

cboFiltro.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oBarra = Nothing
frmReportes.Enabled = True
End Sub

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub

Private Sub txtCtaCod_EmiteDatos()
txtCtaDesc = txtCtaCod.psDescripcion
If txtCtaDesc <> "" Then
   txtFechaDel.SetFocus
End If
End Sub

Private Sub txtFechaDel_GotFocus()
txtFechaDel.SelStart = 0
txtFechaDel.SelLength = Len(txtFechaDel)
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaDel) <> "" Then
      MsgBox "Fecha no válida...", vbInformation, "Aviso"
   Else
      txtFechaAl.SetFocus
   End If
End If
End Sub

Private Sub txtFechaAl_GotFocus()
txtFechaAl.SelStart = 0
txtFechaAl.SelLength = Len(txtFechaAl)
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaAl) <> "" Then
      MsgBox "Fecha no Válida...", vbInformation, "Aviso"
      Exit Sub
   End If
   txtImporte.SetFocus
End If
End Sub

Private Sub txtFechaDel_Validate(Cancel As Boolean)
If txtFechaDel = "" Then
   Exit Sub
End If
   If ValidaFecha(txtFechaAl) <> "" Then
      MsgBox "Fecha no válida...", vbInformation, "Aviso"
      Cancel = True
   End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 16, 2)
If KeyAscii = 13 Then
   If txtImporte <> "" Then
      txtImporte = Format(txtImporte, gsFormatoNumeroDato)
   End If
   cmdProcesar.SetFocus
End If
End Sub

Private Sub txtImporte_Validate(Cancel As Boolean)
If txtImporte <> "" Then
   txtImporte = Format(txtImporte, gsFormatoNumeroDato)
End If
End Sub

