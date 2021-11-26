VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajeroOpeDevSob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Castigo de Sobrante y Faltanes "
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "frmCajeroOpeDevSob.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lblMonto 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   4840
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   3480
      Width           =   1760
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Faltantes"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2475
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   6555
      Begin VB.CheckBox Check1 
         Caption         =   "Todos"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
      Begin SICMACT.FlexEdit fgDev 
         Height          =   1815
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3201
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-OK-Tipo-User-Monto-Movim-Fecha-Age"
         EncabezadosAnchos=   "350-400-1200-1200-1200-0-1200-0"
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
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-4-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3900
      TabIndex        =   13
      Top             =   4620
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5250
      TabIndex        =   12
      Top             =   4620
      Width           =   1365
   End
   Begin VB.TextBox txtGlosa 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   210
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3840
      Width           =   3555
   End
   Begin VB.Frame Frame3 
      Height          =   900
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6570
      Begin VB.CheckBox Check2 
         Caption         =   "Sobrantes"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   120
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Dolares"
         Height          =   240
         Index           =   2
         Left            =   1020
         TabIndex        =   1
         Top             =   165
         Width           =   1050
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Soles"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   0
         Top             =   150
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.CommandButton btn_buscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txNotaIni 
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txNotaFin 
         Height          =   300
         Left            =   3120
         TabIndex        =   5
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2280
         TabIndex        =   11
         Top             =   525
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   510
         Width           =   555
      End
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ITF :"
      Height          =   195
      Left            =   4095
      TabIndex        =   20
      Top             =   3975
      Width           =   330
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total :"
      Height          =   195
      Left            =   4095
      TabIndex        =   19
      Top             =   4305
      Width           =   450
   End
   Begin VB.Label lblITF 
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
      ForeColor       =   &H80000001&
      Height          =   300
      Left            =   4845
      TabIndex        =   18
      Top             =   3870
      Width           =   1755
   End
   Begin VB.Label lblTotal 
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
      ForeColor       =   &H80000001&
      Height          =   300
      Left            =   4845
      TabIndex        =   17
      Top             =   4230
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Monto :"
      Height          =   195
      Left            =   4095
      TabIndex        =   16
      Top             =   3540
      Width           =   540
   End
   Begin VB.Label lblGlosa 
      Caption         =   "Glosa"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   915
   End
End
Attribute VB_Name = "frmCajeroOpeDevSob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCajero As nCajero
Dim lnOpeCod As Long
Dim lsOperacion As String
Dim bOpeAfecta As Boolean
Dim lsCaption As String
Dim objPista As COMManejador.Pista 'MAVM 05102010
Dim lnCastDev As Integer 'MADM 20110930
'**MADM 20101006 *************************************
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
'*****************************************************
Dim nRedondeoITF As Double
Dim nMontoAPagar As Currency
Dim lnMonto As Currency
Dim lsOpeCodFal As String 'PASIERS1552014
Dim lsOpeCodCom As String 'PASIERS1552014
Public Sub inicia(ByVal pnOpeCod As Long, ByVal PsOperacion As String)
lnOpeCod = pnOpeCod
lsOperacion = PsOperacion
bOpeAfecta = VerifOpeVariasAfectaITF(Str(lnOpeCod))
lsCaption = Mid(PsOperacion, 3, Len(PsOperacion) - 2)
lblTotal.ForeColor = &H80000012
lblITF.ForeColor = &H80000012
lblMonto.ForeColor = &H80000012
Me.txNotaFin.Visible = True
Me.txNotaIni.Visible = True
Me.btn_buscar.Visible = True
Me.txNotaFin.Text = gdFecSis
Me.txNotaIni.Text = gdFecSis
Me.Caption = lsCaption
lsOpeCodFal = "300530" 'PASIERS1552014
lsOpeCodCom = "300531" 'PASIERS1552014
Me.Show 1
End Sub

Private Sub CargaFaltantes(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal pnMoneda As COMDConstantes.Moneda, Optional ByVal num As Integer = 0)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oITF As New COMDConstSistema.FCOMITF
Dim oCaj As COMNCajaGeneral.NCOMCajero
Dim lnTotal As Currency
Dim ind As Integer

lblTotal = "0.00"
lblITF = "0.00"
lblMonto.Text = "0.00"
lnTotal = "0.00"

fgDev.Clear
fgDev.FormaCabecera
fgDev.Rows = 2

oITF.fgITFParametros
oITF.gbITFAplica = False

ind = 0
If Me.Check2.value And Me.Check3.value Then
    ind = 0
ElseIf Me.Check2.value And Me.Check3.value = 0 Then
    ind = 1
Else
    ind = 2
End If

    Set oCaj = New COMNCajaGeneral.NCOMCajero
    Set rs = oCaj.ObtenerRegularizacionSobranteCastigo(pdFecIni, pdFecFin, pnMoneda, ind)
    Set oCaj = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        Do While Not rs.EOF
            fgDev.AdicionaFila , , True
                
                If rs!nMontran > 0 Then
                    fgDev.TextMatrix(fgDev.row, 2) = "SOBRANTE"
                Else
                    fgDev.TextMatrix(fgDev.row, 2) = "FALTANTE"
                End If
                
                fgDev.TextMatrix(fgDev.row, 3) = rs!cCodusu
                fgDev.TextMatrix(fgDev.row, 4) = rs!nMontran
                fgDev.TextMatrix(fgDev.row, 5) = rs!cNrotran
                fgDev.TextMatrix(fgDev.row, 6) = rs!dFecTran
                fgDev.TextMatrix(fgDev.row, 7) = rs!Age 'PASIERS1552014
            rs.MoveNext
        Loop
    fgDev.lbEditarFlex = True
    lblMonto.Text = Format(lnTotal, "#0.00")
    oITF.gbITFAplica = False 'EJVG20121029
    If oITF.gbITFAplica Then
        lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(CDbl(lblMonto.Text)), "#,##0.00")
        '*** BRGO 20110908 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
        If nRedondeoITF > 0 Then
            Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
        End If
        '*** END BRGO
    End If
    lblTotal.Caption = Format(CCur(CCur(lblMonto.Text) - Me.lblITF.Caption), "#,##0.00")
Else
    'MsgBox "No se encontraron datos para el Usuario Seleccionado", vbInformation, "Aviso"
    MsgBox "No se encontraron datos", vbInformation, "Aviso"
End If
rs.Close: Set rs = Nothing
Set oITF = Nothing
End Sub

Private Sub btn_buscar_Click()
'EJVG20121029 ***
'   If Me.txNotaFin.Text <> "" And Me.txNotaIni <> "" Then
'    If IsDate(txNotaIni) And IsDate(txNotaFin) Then
'        CargaFaltantes CDate(Me.txNotaIni.Text), CDate(Me.txNotaFin.Text), IIf(Me.optMon(3).value = True, gMonedaNacional, IIf(Me.optMon(2).value = True, gMonedaExtranjera, gMonedaNacional))
'        Me.txtGlosa.SetFocus
'    Else
'        MsgBox "Los Valores indicados en los textos no son Correctos"
'    End If
'Else
'    MsgBox "Complete los Datos para Generar el Reporte"
'End If
    If Not ValidaBuscar Then Exit Sub
    CargaFaltantes CDate(Me.txNotaIni.Text), CDate(Me.txNotaFin.Text), IIf(Me.optMon(3).value = True, gMonedaNacional, IIf(Me.optMon(2).value = True, gMonedaExtranjera, gMonedaNacional))
    Me.txtGlosa.SetFocus
'END EJVG *******
End Sub

Private Sub Check1_Click()
    Dim i As Long
    nMontoAPagar = 0
    'EJVG20121029 ***
    'lblTotal.Caption = 0
    lblTotal.Caption = "0.00"
    If fgDev.Rows = 2 And fgDev.TextMatrix(1, 0) = "" Then 'Esta Vacío
        Exit Sub
    End If
    'END EJVG *******
    If Me.Check1.value Then
        For i = 1 To fgDev.Rows - 1
            fgDev.TextMatrix(i, 1) = 1
            nMontoAPagar = CDbl(nMontoAPagar) + CDbl(fgDev.TextMatrix(i, 4))
        Next i
    Else
         For i = 1 To fgDev.Rows - 1
            fgDev.TextMatrix(i, 1) = 0
            nMontoAPagar = 0
            lblITF.Caption = Format(nMontoAPagar, "#0.00")
        Next i
    End If
    'EJVG20121029 ***
    'lblMonto.Text = Format(nMontoAPagar, "#0.00")
    lblMonto.Text = Format(nMontoAPagar, gsFormatoNumeroView)
    lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + lblMonto.Text, gsFormatoNumeroView)
    cmdGrabar.Enabled = False
    'END EJVG *******
End Sub

Private Sub cmdCancelar_Click()
lblTotal = "0.00"
fgDev.Clear
fgDev.Rows = 2
fgDev.FormaCabecera
txtGlosa = ""
lblMonto.Text = "0.00"
lblTotal.Caption = "0.00"
lblITF = "0.00"
nRedondeoITF = 0
'EJVG20121029 ***
Check1.value = 0
cmdGrabar.Enabled = False
'END EJVG *******
End Sub

Private Sub cmdGrabar_Click()
Dim CodOpe As String
Dim Moneda As String
Dim lsMov As String
Dim lsMovITF As String
Dim lsDocumento As String
Dim i As Long

Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim clsCont As COMNContabilidad.NCOMContFunciones
Set clsCont = New COMNContabilidad.NCOMContFunciones
Dim oFunFecha As New COMDConstSistema.DCOMGeneral
Dim oITF As New COMDConstSistema.FCOMITF
Dim lnMovNroBR As Long 'ALPA20131002

Dim k As Integer 'PASIERS1552014
Dim J, bestad As Boolean 'PASIERS1552014
Dim TMatSobFal() As TCastigoSobFal 'PASIERS1552014


lnMonto = CCur(lblMonto.Text)
lsMov = oFunFecha.FechaHora(gdFecSis)
Set oFunFecha = Nothing
lsDocumento = ""

Dim lnMovNro As Long
Dim lnMovNroITF As Long
Dim lbBan As Boolean

Dim loMov As COMDMov.DCOMMov
Set loMov = New COMDMov.DCOMMov

'EJVG20121029 ***
'If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbYes Then
If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If
'END EJVG *******

'Modificado PASIERS1552014
    '    For i = 1 To fgDev.Rows - 1
    '        'MADM 20110930 lnCastDev - lnMonto
    '        lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    '        'lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lnOpeCod, Trim(fgDev.TextMatrix(i, 3)), lsdocumento, txtGlosa.Text, IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera), txtCodPers.Text, Trim(fgDev.TextMatrix(i, 5)))
    '         If fgDev.TextMatrix(i, 1) = "." Then
    '            If fgDev.TextMatrix(i, 2) = "SOBRANTE" Then
    '               'loMov.InsertaMov lsMov, COMDConstSistema.gOpeHabCajIngRegulaSob, "Castigo Sobrante"
    '               'lnMovNro = clsCapMov.OtrasOperaciones(lsMov, COMDConstSistema.gOpeHabCajIngRegulaSob, Trim(fgDev.TextMatrix(i, 4)), lsdocumento, txtGlosa.Text, IIf(Me.optMon(3).value = True, gMonedaNacional, gMonedaExtranjera), "", Trim(fgDev.TextMatrix(i, 5)))
    '               lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lnOpeCod, Trim(fgDev.TextMatrix(i, 4)), lsDocumento, "Castigo Sobrante", IIf(Me.optMon(3).value = True, gMonedaNacional, gMonedaExtranjera), "", Trim(fgDev.TextMatrix(i, 5)), , , , , , lnMovNroBR)
    '            Else
    '               'loMov.InsertaMov lsMov, gsOpeCod, "Castigo Faltante"
    '               'lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lnOpeCod, Trim(fgDev.TextMatrix(i, 4)), lsdocumento, txtGlosa.Text, IIf(Me.optMon(3).value = True, gMonedaNacional, gMonedaExtranjera), "", Trim(fgDev.TextMatrix(i, 5)))
    '               lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lnOpeCod, Trim(fgDev.TextMatrix(i, 4)), lsDocumento, "Castigo Faltante", IIf(Me.optMon(3).value = True, gMonedaNacional, gMonedaExtranjera), "", Trim(fgDev.TextMatrix(i, 5)), , , , , , lnMovNroBR)
    '            End If
    '            'ALPA20131001*********************************
    '            If lnMovNroBR = 0 Then
    '                MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
    '                Exit Sub
    '            End If
    '            '*********************************************
    '        End If
    '        'loMov.InsertaMovOpeVarias lnMovNro, lsdocumento, Trim(txtGlosa.Text), Trim(fgDev.TextMatrix(i, 3)), IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera)
    '        'loMov.InsertaMovRef lnMovNro, Trim(fgDev.TextMatrix(i, 5))
    '    Next
    ''End If

    ReDim Preserve TMatSobFal(0)
    'bestad = False /**Comentado PASI20170213**/
    J = True
    For i = 1 To fgDev.Rows - 1
        bestad = False 'PASI20170213
        If fgDev.TextMatrix(i, 1) = "." Then
            If J Then
                ReDim Preserve TMatSobFal(1)
                TMatSobFal(1).nOpeCod = IIf(fgDev.TextMatrix(i, 2) = "SOBRANTE", CLng(300509), CLng(300530))
                TMatSobFal(1).sUser = fgDev.TextMatrix(i, 3)
                TMatSobFal(1).nMonto = fgDev.TextMatrix(i, 4)
                TMatSobFal(1).nMovNro = fgDev.TextMatrix(i, 5)
                TMatSobFal(1).sAge = fgDev.TextMatrix(i, 7)
                TMatSobFal(1).bEstado = 0
                J = False
            Else
                For k = 1 To UBound(TMatSobFal, 1)
                    If TMatSobFal(k).sUser = fgDev.TextMatrix(i, 3) And _
                       TMatSobFal(k).sAge = fgDev.TextMatrix(i, 7) And _
                       Abs(TMatSobFal(k).nMonto) = Abs(fgDev.TextMatrix(i, 4)) And _
                       TMatSobFal(k).bEstado = 0 And _
                       (TMatSobFal(k).nOpeCod <> IIf(fgDev.TextMatrix(i, 2) = "SOBRANTE", CLng(300509), CLng(300530))) Then
                       TMatSobFal(k).nOpeCod = 300531
                       TMatSobFal(k).bEstado = 1
                       TMatSobFal(k).nMovNro2 = fgDev.TextMatrix(i, 5) 'PASI20150414
                       bestad = True
                       Exit For
                    End If
                Next
                If Not bestad Then
                    ReDim Preserve TMatSobFal(UBound(TMatSobFal, 1) + 1)
                    TMatSobFal(UBound(TMatSobFal, 1)).nOpeCod = IIf(fgDev.TextMatrix(i, 2) = "SOBRANTE", CLng(300509), CLng(300530))
                    TMatSobFal(UBound(TMatSobFal, 1)).sUser = fgDev.TextMatrix(i, 3)
                    TMatSobFal(UBound(TMatSobFal, 1)).nMonto = fgDev.TextMatrix(i, 4)
                    TMatSobFal(UBound(TMatSobFal, 1)).nMovNro = fgDev.TextMatrix(i, 5)
                    TMatSobFal(UBound(TMatSobFal, 1)).sAge = fgDev.TextMatrix(i, 7)
                    TMatSobFal(UBound(TMatSobFal, 1)).bEstado = 0
                End If
            End If
        End If
    Next
    If UBound(TMatSobFal, 1) > 0 Then
        For i = 1 To UBound(TMatSobFal, 1)
            Sleep 1000
            lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            lnMovNro = clsCapMov.OtrasOperaciones(lsMov, TMatSobFal(i).nOpeCod, TMatSobFal(i).nMonto, lsDocumento, IIf(TMatSobFal(i).nOpeCod = 300509, "Castigo Sobrante", IIf(TMatSobFal(i).nOpeCod = 300530, "Castigo Faltante", "Castigo Sobrante Faltante")), IIf(Me.optMon(3).value = True, gMonedaNacional, gMonedaExtranjera), "", TMatSobFal(i).nMovNro, , , , , , lnMovNroBR, TMatSobFal(i).nMovNro2)
            If lnMovNroBR = 0 Then
                MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
                Exit Sub
            End If
        Next
    End If
'END PASI
  
   oITF.gbITFAplica = False
    
    If oITF.gbITFAplica And CCur(Me.lblITF.Caption) <> 0 Then
        lsMovITF = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser, lsMov)
        lnMovNroITF = clsCapMov.OtrasOperaciones(lsMovITF, COMDConstantes.gITFCobroEfectivo, Abs(Me.lblITF.Caption), lsDocumento, Me.txtGlosa.Text, IIf(Me.optMon(0).value = True, COMDConstantes.gMonedaNacional, COMDConstantes.gMonedaExtranjera), "")
        Call loMov.InsertaMovRedondeoITF(lsMovITF, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption)) 'BRGO 20110914
        Set loMov = Nothing
    End If
    
    'MAVM 05102010 ***
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, lsMov, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Castigo Sobrante Faltante ", , gCodigoPersona
    Set objPista = Nothing
    '***
    
    Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim oBolITF As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    Dim nFicSal As Integer
    
    Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
        lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", Left(lsCaption, 15), "", Str(lblMonto.Text), "", "________" & IIf(optMon(3).value = True, gMonedaNacional, gMonedaExtranjera), lsDocumento, 0, "0", IIf(Len(lsDocumento) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, (lblITF.Caption * -1))
    Set oBol = Nothing
    
    Do
         If Trim(lsBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                Print #nFicSal, ""
            Close #nFicSal
          End If
                   
    Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
  
    cmdCancelar_Click

Set oITF = Nothing
Set clsCapMov = Nothing
Set clsCont = Nothing
End Sub

Private Sub fgDev_OnCellChange(pnRow As Long, pnCol As Long)
    Call fgDev_OnCellCheck(1, 1)
End Sub


Private Sub fgDev_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim i As Long
    nMontoAPagar = 0
    
    lblMonto.Text = ""
    lblTotal.Caption = ""
    For i = 0 To fgDev.Rows - 2
        If fgDev.TextMatrix(i + 1, 1) = "." Then
            nMontoAPagar = CDbl(nMontoAPagar) + CDbl(fgDev.TextMatrix(i + 1, 4))
        End If
    Next i

    'EJVG20121029 ***
    'lblMonto.Text = Format(nMontoAPagar, "#0.00")
    lblMonto.Text = Format(nMontoAPagar, gsFormatoNumeroView)
    lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + CCur(lblMonto.Text), gsFormatoNumeroView)
    cmdGrabar.Enabled = False
    'END EJVG *******
End Sub
Private Sub lblMonto_GotFocus()
    With lblMonto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub lblMonto_KeyPress(KeyAscii As Integer)
    Dim i As Integer 'PASI20150414
    Dim bestad As Boolean 'PASI20150414
    
      If KeyAscii = 13 Then
       Dim oITF As New COMDConstSistema.FCOMITF

            'If lblMonto.value <> 0 Then
            If lblMonto.Text <> 0 Then 'EJVG20121029
                cmdGrabar.Enabled = True
            Else
                'cmdGrabar.Enabled = false 'Comentado PASI20150414
                
                'PASI20150414
                For i = 1 To fgDev.Rows - 1
                    If fgDev.TextMatrix(i, 1) = "." Then
                        bestad = True
                    End If
                Next
                If bestad Then
                     cmdGrabar.Enabled = True
                Else
                    cmdGrabar.Enabled = False
                End If
                'end PASI
                
            End If

            oITF.fgITFParametros
            oITF.gbITFAplica = False
            If oITF.gbITFAplica Then
                Me.lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(lblMonto.Text), "#,##0.00")
                '*** BRGO 20110908 ************************************************
                    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
                    If nRedondeoITF > 0 Then
                       Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
                    End If
                '*** END BRGO
            End If
            Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + lblMonto.Text, "#,##0.00")
       Set oITF = Nothing
       If cmdGrabar.Enabled Then Me.cmdGrabar.SetFocus
    End If
End Sub

'EJVG20121029 ***
Private Function ValidaBuscar() As Boolean
    ValidaBuscar = True
    If Check2.value = 0 And Check3.value = 0 Then
        MsgBox "Ud. debe seleccionar entre Sobrantes y/o Faltantes", vbInformation, "Aviso"
        ValidaBuscar = False
        Check2.SetFocus
        Exit Function
    End If
    If Not IsDate(txNotaIni.Text) Then
        MsgBox "La fecha inicio de búsqueda es incorrecta", vbInformation, "Aviso"
        ValidaBuscar = False
        txNotaIni.SetFocus
        Exit Function
    End If
    If Not IsDate(txNotaFin.Text) Then
        MsgBox "La fecha fin de búsqueda es incorrecta", vbInformation, "Aviso"
        ValidaBuscar = False
        txNotaFin.SetFocus
        Exit Function
    End If
    If CDate(txNotaIni.Text) > CDate(txNotaFin.Text) Then
        MsgBox "La fecha fin de búsqueda no puede ser mayor que la de inicio", vbInformation, "Aviso"
        ValidaBuscar = False
        txNotaFin.SetFocus
        Exit Function
    End If
End Function
Private Sub Check2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Check3.SetFocus
    End If
End Sub
Private Sub Check3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txNotaIni.SetFocus
    End If
End Sub
Private Sub txNotaIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txNotaFin.SetFocus
    End If
End Sub
Private Sub txNotaFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btn_buscar.SetFocus
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
'END EJVG *******
