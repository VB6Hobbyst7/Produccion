VERSION 5.00
Begin VB.Form frmCapServConvAbonoInst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEPOSITO ENT. CONVENIO"
   ClientHeight    =   7500
   ClientLeft      =   4140
   ClientTop       =   1860
   ClientWidth     =   8715
   Icon            =   "frmCapServConvAbonoInst.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCobranza 
      Caption         =   "Cobranza"
      Enabled         =   0   'False
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
      Height          =   3750
      Left            =   45
      TabIndex        =   8
      Top             =   3225
      Width           =   8565
      Begin VB.TextBox txtGlosa 
         Height          =   1080
         Left            =   180
         MaxLength       =   255
         TabIndex        =   34
         Top             =   2340
         Width           =   3705
      End
      Begin VB.Frame fraMonto 
         Caption         =   "Monto"
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
         Height          =   1530
         Left            =   4905
         TabIndex        =   25
         Top             =   1995
         Width           =   3465
         Begin VB.CheckBox chkITFEfectivo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Efect"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   900
            TabIndex        =   26
            Top             =   825
            Width           =   705
         End
         Begin SICMACT.EditMoney txtMonto 
            Height          =   375
            Left            =   900
            TabIndex        =   27
            Top             =   345
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12648447
            ForeColor       =   12582912
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblMon 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2895
            TabIndex        =   33
            Top             =   1215
            Width           =   315
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   195
            TabIndex        =   32
            Top             =   435
            Width           =   540
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   900
            TabIndex        =   31
            Top             =   1155
            Width           =   1905
         End
         Begin VB.Label lblITF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1710
            TabIndex        =   30
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Total :"
            Height          =   195
            Left            =   255
            TabIndex        =   29
            Top             =   1245
            Width           =   450
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "ITF :"
            Height          =   195
            Left            =   255
            TabIndex        =   28
            Top             =   825
            Width           =   330
         End
      End
      Begin VB.OptionButton Opt 
         Caption         =   " Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2685
         TabIndex        =   19
         Top             =   420
         Width           =   1800
      End
      Begin VB.OptionButton Opt 
         Caption         =   " Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   450
         Value           =   -1  'True
         Width           =   1800
      End
      Begin VB.Frame fraClente 
         Height          =   1125
         Left            =   105
         TabIndex        =   9
         Top             =   825
         Width           =   8280
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2235
            TabIndex        =   13
            Top             =   210
            Width           =   390
         End
         Begin VB.Label lblDICli 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   4815
            TabIndex        =   17
            Top             =   225
            Width           =   2250
         End
         Begin VB.Label lbllCodigoCli 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   795
            TabIndex        =   16
            Top             =   195
            Width           =   1380
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Documento Identidad:"
            Height          =   195
            Left            =   3150
            TabIndex        =   14
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label lblNomCli 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   780
            TabIndex        =   12
            Top             =   600
            Width           =   7350
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   60
            TabIndex        =   11
            Top             =   660
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Left            =   60
            TabIndex        =   10
            Top             =   255
            Width           =   585
         End
      End
      Begin VB.Frame FraReferencia 
         Height          =   1125
         Left            =   105
         TabIndex        =   20
         Top             =   825
         Width           =   7305
         Begin VB.TextBox TxtNombre 
            Height          =   315
            Left            =   1695
            MaxLength       =   100
            TabIndex        =   24
            Top             =   630
            Width           =   5490
         End
         Begin VB.TextBox txtDI 
            Height          =   315
            Left            =   1710
            MaxLength       =   20
            TabIndex        =   23
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   105
            TabIndex        =   22
            Top             =   690
            Width           =   600
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Documento Identidad:"
            Height          =   195
            Left            =   105
            TabIndex        =   21
            Top             =   315
            Width           =   1575
         End
      End
      Begin VB.Label lblTransferGlosa 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   2010
         Width           =   495
      End
   End
   Begin VB.Frame fraInstitucion 
      Caption         =   "Institución"
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
      Height          =   3225
      Left            =   60
      TabIndex        =   3
      Top             =   -15
      Width           =   8580
      Begin SICMACT.FlexEdit grdCuentas 
         Height          =   1500
         Left            =   90
         TabIndex        =   7
         Top             =   1140
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   2646
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   2
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nro Cuenta-Moneda-Fecha Apertura-(*)"
         EncabezadosAnchos=   "400-3000-1600-1600-800"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnasAEditar =   "X-X-X-X-4"
         ListaControles  =   "0-0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-L"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label LblDII 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4200
         TabIndex        =   37
         Top             =   300
         Width           =   3060
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DOC NRO.  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3255
         TabIndex        =   36
         Top             =   375
         Width           =   960
      End
      Begin VB.Label lblCodI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1050
         TabIndex        =   15
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO"
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
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE"
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
         Top             =   720
         Width           =   810
      End
      Begin VB.Label lblNombreI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1050
         TabIndex        =   4
         Top             =   660
         Width           =   6225
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   7035
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6555
      TabIndex        =   0
      Top             =   7050
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7620
      TabIndex        =   1
      Top             =   7050
      Width           =   1000
   End
End
Attribute VB_Name = "frmCapServConvAbonoInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCuentaPension As String
Dim sCuentaMora As String
Dim sCuentaGasto As String
Dim sCodigoPers As String
Dim nMoraDia As Double

Dim nPersoneria As Integer
Dim pbOrdPag As Boolean

Dim nInstConvenio As COMDConstantes.CaptacConvenios
Dim nPos As Integer, vsCuenta As String
Dim lbITFCtaExonerada As Boolean
Dim nRedondeoITF As Double 'BRGO 20110908

'Private Sub GetTotalPagar()
'Dim i As Long
'Dim nMontoMN As Double, nMontoME As Double
'nMontoMN = 0
'nMontoME = 0
'For i = 1 To grdPago.Rows - 1
'    If CLng(Mid(grdPago.TextMatrix(i, 1), 9, 1)) = gMonedaNacional Then
'        nMontoMN = nMontoMN + CDbl(grdPago.TextMatrix(i, 3))
'    Else
'        nMontoME = nMontoME + CDbl(grdPago.TextMatrix(i, 3))
'    End If
'Next i
'lblTotalMN = "S/. " & Format$(nMontoMN, "#,##0.00")
'lblTotalME = "US$ " & Format$(nMontoME, "#,##0.00")
'End Sub

'Private Sub ClearScreen()
'fraAlumno.Enabled = True
'Select Case nInstConvenio
'    Case gCapConvJuanPabloInst
'        txtCodigo.Mask = "####"
'        txtCodigo.Text = "____"
'    Case gCapConvJuanPabloII
'        txtCodigo.Mask = "#####"
'        txtCodigo.Text = "_____"
'    Case gCapConvNarvaez
'        txtCodigo.Mask = "C###-##"
'        txtCodigo.Text = "____-__"
'    Case gCapConvMarianoSantos, gCapConvSantaRosa
'        txtCodigo.Mask = "C##"
'        txtCodigo.Text = "___"
'End Select
'lblNombre = ""
'lblNivel = ""
'lblSeccion = ""
'lblGrado = ""
'LblCondicion = ""
'lblTotalMN = "S/. 0.00"
'lblTotalME = "US$ 0.00"
'txtFecha = "__/__/____"
'GetMesesAño
'cmdCancelar.Enabled = False
'cmdGrabar.Enabled = False
'FraPago.Enabled = False
'End Sub

'Private Sub CalculaMora()
'Dim nDias As Integer
'Dim nFeriados As Long
'Dim clsGen As DGeneral
'If aPlanPago(nPos).bMora Then
'    If aPlanPago(nPos).bFeriado Then
'        Set clsGen = New DGeneral
'        nFeriados = clsGen.GetNumDiasFeriado(CDate(txtFecha.Text), gdFecSis)
'        Set clsGen = Nothing
'    Else
'        nFeriados = 0
'    End If
'    nDias = DateDiff("d", CDate(txtFecha), gdFecSis) - nFeriados
'    If nDias > 0 Then
'        Select Case nInstConvenio
'            Case gCapConvNarvaez, gCapConvSantaRosa
'                grdPago.TextMatrix(3, 3) = Format$(nMoraDia * nDias, "#,##0.00")
'            Case gCapConvJuanPabloII, gCapConvJuanPabloInst, gCapConvMarianoSantos
'                grdPago.TextMatrix(2, 3) = Format$(nMoraDia * nDias, "#,##0.00")
'        End Select
'        LblDiasAtraso = Format$(nDias, "#0")
'    Else
'        Select Case nInstConvenio
'            Case gCapConvNarvaez, gCapConvSantaRosa
'                grdPago.TextMatrix(3, 3) = "0.00"
'            Case gCapConvJuanPabloII, gCapConvJuanPabloInst, gCapConvMarianoSantos
'                grdPago.TextMatrix(2, 3) = "0.00"
'        End Select
'        LblDiasAtraso = "0"
'    End If
'End If
'End Sub

'Private Function GetPosicionMes(ByVal nMes As Integer) As Integer
'Dim i As Integer
'Dim dReferencia As Date
'dReferencia = CDate("01/" & nMes & "/" & Year(gdFecSis))
'For i = 1 To UBound(aPlanPago, 1)
'    If DateDiff("m", aPlanPago(i).dVencimiento, dReferencia) = 0 Then
'        GetPosicionMes = i
'        Exit For
'    End If
'Next i
'End Function

'Private Sub GetMesesAño()
'cboMes.Clear
'cboMes.AddItem "Enero" & Space(50) & "01"
'cboMes.AddItem "Febrero" & Space(50) & "02"
'cboMes.AddItem "Marzo" & Space(50) & "03"
'cboMes.AddItem "Abril" & Space(50) & "04"
'cboMes.AddItem "Mayo" & Space(50) & "05"
'cboMes.AddItem "Junio" & Space(50) & "06"
'cboMes.AddItem "Julio" & Space(50) & "07"
'cboMes.AddItem "Agosto" & Space(50) & "08"
'cboMes.AddItem "Septiembre" & Space(50) & "09"
'cboMes.AddItem "Octubre" & Space(50) & "10"
'cboMes.AddItem "Noviembre" & Space(50) & "11"
'cboMes.AddItem "Diciembre" & Space(50) & "12"
'End Sub



Public Sub Inicia(ByVal sPersona As String, ByVal sNombre As String, ByVal sDNI As String)
 
 Me.lblCodI = sPersona
 Me.lblNombreI = sNombre
 Me.LblDII.Caption = sDNI
 Call GetCuentasAbono(sPersona)
 
End Sub

Private Sub GetCuentasAbono(ByVal sPersCod As String)
Dim rsCta As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
Set rsCta = clsServ.GetServConvCuentas(sPersCod, nInstConvenio, "S")
Set clsServ = Nothing
If rsCta.EOF And rsCta.BOF Then
    MsgBox "No están registradas las cuentas de abono para esta institución.", vbExclamation, "Aviso"
    fraCobranza.Enabled = False
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    Unload Me
    Exit Sub
Else
    'Obtienes los datos de las cuentas a las cuales se realizará el abono
    Do While Not rsCta.EOF
        grdCuentas.AdicionaFila
        If CLng(Mid(rsCta("Nro Cuenta"), 9, 1)) = gMonedaExtranjera Then
            grdCuentas.BackColorRow &HC0FFC0
        Else
        End If
        grdCuentas.TextMatrix(grdCuentas.Row, 1) = rsCta("Nro Cuenta")
        grdCuentas.TextMatrix(grdCuentas.Row, 2) = rsCta("Moneda")
        grdCuentas.TextMatrix(grdCuentas.Row, 3) = rsCta("Fecha Apertura")
        rsCta.MoveNext
    Loop
    
    If rsCta.RecordCount = 1 Then
        grdCuentas.TextMatrix(grdCuentas.Row, 4) = "1"
        Call grdCuentas_OnCellCheck(grdCuentas.Row, 4)
    End If
End If
rsCta.Close
Set rsCta = Nothing
Me.Show 1

End Sub

Private Sub chkITFEfectivo_Click()
  If chkITFEfectivo.value = 1 Then
        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
        Me.lblTotal.Caption = Format(Me.txtMonto.value + CCur(Me.LblITF.Caption), "#,##0.00")
    Else
        If gbITFAsumidoAho Then
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
        
        Else
                    Me.lblTotal.Caption = Format(txtMonto.value - CCur(Me.LblITF.Caption), "#,##0.00")
        End If
    
        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String, lsDNI As String
Dim lsEstados As String


On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    lsDNI = loPers.sPersIdnroDNI


If lsPersCod <> "" Then
    Me.lbllCodigoCli.Caption = lsPersCod
    Me.LblNomCli.Caption = lsPersNombre
    Me.lblDICli.Caption = lsDNI
End If

Set loPers = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdCancelar_Click()
'ClearScreen
Limpiar

End Sub

Private Sub Limpiar()
Dim i As Long
 Opt(0).value = True
 FraReferencia.Visible = False
 Me.fraClente.Visible = True
 txtDI.Text = ""
 txtNombre.Text = ""
 lblDICli.Caption = ""
 Me.lbllCodigoCli.Caption = ""
 Me.LblNomCli.Caption = ""
 chkITFEfectivo.value = vbUnchecked
 
        txtMonto.Text = "0.00"
        Me.LblITF.Caption = "0.00"
        Me.lblTotal.Caption = "0.00"
        
        fraInstitucion.Enabled = True
        fraCobranza.Enabled = False
        i = 1
        For i = 1 To grdCuentas.Rows - 1
               grdCuentas.TextMatrix(i, 4) = ""
        Next i
        vsCuenta = ""
nRedondeoITF = 0
        
End Sub


Private Sub cmdGrabar_Click()
'Dim nMontoMN As Double, nMontoME As Double
Dim nMonto As Double, nmoneda As Integer
Dim sDIPers As String, sNombrePers As String, sCuenta As String, nOperacion As CaptacOperacion

nOperacion = gAhoDepEntConv

sCuenta = vsCuenta


''Leo los datos para impresion
If Opt(0).value = True Then
    sNombrePers = ImpreCarEsp(Trim(Me.LblNomCli.Caption))
    sDIPers = ImpreCarEsp(Trim(Me.lblDICli.Caption))
Else
    sNombrePers = ImpreCarEsp(Trim(Me.txtNombre.Text))
    sDIPers = ImpreCarEsp(Trim(Me.txtDI.Text))
End If


nMonto = CDbl(txtMonto.Text)
'
'nMontoMN = CDbl(Replace(lblTotal.Caption, "S/.", "", 1, , vbTextCompare))
'nMontoME = CDbl(Replace(lblTotal.Caption, "US$", "", 1, , vbTextCompare))

If Trim(sNombrePers) = "" Or Trim(sDIPers) = "" Then
    MsgBox "Los datos de la persona que realiza la transacción estan incompletos!!!.", vbOKOnly + vbExclamation, "AVISO"
    Exit Sub
End If

''Valida de que el monto por Abono sea mayor que cero
If nMonto = 0 Then
    MsgBox "Monto de Abono deber ser mayo a cero", vbInformation, "Aviso"
    Exit Sub
End If

Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
Dim nMontoMinDep As Double
Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'txtCuenta.NroCuenta
nMontoMinDep = clsDef.GetMontoMinimoDepPersoneria(gCapAhorros, Mid(sCuenta, 9, 1), nPersoneria, pbOrdPag)
If nMontoMinDep > txtMonto.value Then
       MsgBox "El Monto del Abono es menor al mínimo permitido de " & IIf(Mid(sCuenta, 9, 1) = 1, "S/. ", "US$. ") & CStr(nMontoMinDep), vbOKOnly + vbInformation, "Aviso"
       Exit Sub
End If
Set clsDef = Nothing


''Inicia el proceso de grabación
If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Dim sMovNro As String, sPersLavDinero As String ', sReaPersLavDinero As String, sBenPersLavDinero As String 'JACA 20110224
    
    Dim loLavDinero As SICMACT.frmMovLavDinero 'JACA 20110224
    Set loLavDinero = New SICMACT.frmMovLavDinero 'JACA 20110224
    
    Dim loMov As COMDMov.DCOMMov
    
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim nSaldo As Double, nPorcDisp As Double
    Dim nMontoLavDinero As Double, nTC As Double, pnExtracto As Long
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    Dim clsAge As COMDConstantes.DCOMAgencias, sNomAge As String
    Set clsAge = New COMDConstantes.DCOMAgencias
    
    Dim lsmensaje As String
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    
    Dim nFicSal As Integer
    
    sNomAge = clsAge.NombreAgencia(gsCodAge)
    Set clsAge = Nothing
    
    On Error GoTo ErrGraba:
    
    'Realiza la Validación para el Lavado de Dinero
    sCuenta = vsCuenta
     
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
        Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            If nmoneda = gMonedaNacional Then
                Dim clsTC As nTipoCambio
                Set clsTC = New nTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                'JACA 20110225
                sPersLavDinero = IniciaLavDinero(loLavDinero)
                sPersLavDinero = loLavDinero.OrdPersLavDinero
                
'               sPersLavDinero = gVarPublicas.gReaPersLavDinero
'               sReaPersLavDinero = gVarPublicas.gReaPersLavDinero
'               sBenPersLavDinero = gVarPublicas.gBenPersLavDinero
                               
                'JACA END
                If sPersLavDinero = "" Then Exit Sub
            End If
        Else
            Set clsExo = Nothing
        End If
        
        
  
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
     
       
    Set clsMov = Nothing
    On Error GoTo ErrGraba
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    
    'ALPA 20081010***************************************************************************************************************************************************************************************************************************************************************************************
    'nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, Trim(txtGlosa.Text), , , , , , , , , gsNomAge, sLpt, sPersLavDinero, True, , , , gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), , pnExtracto, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF)
    'nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, Trim(txtGlosa.Text), , , , , , , , , gsNomAge, sLpt, sPersLavDinero, True, , , , gsCodCMAC, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), , pnExtracto, sBenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, , , , , , , gnMovNro) JACA 20110225
     nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, Trim(txtGlosa.Text), , , , , , , , , gsNomAge, sLpt, sPersLavDinero, True, , , , gsCodCMAC, , gbITFAplica, Me.LblITF.Caption, gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), , pnExtracto, loLavDinero.BenPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, , , , , , , gnMovNro) 'JACA 20110225
    'ALPA 20081010***********************
    If gnMovNro > 0 Then
        'Call frmMovLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sBenPersLavDinero, , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen)
        Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110225
        '*** BRGO 20110906 ***************************
        If gITF.gbITFAplica Then
           Set loMov = New COMDMov.DCOMMov
           Call loMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(Me.LblITF) + nRedondeoITF, CCur(Me.LblITF))
           Set loMov = Nothing
        End If
        '*** BRGO
    End If
    '********************************************************************************************************************************************************************************************************************************************************************************************
    Set clsCap = Nothing
    
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation
     End If
    
'    Open sLpt For Output As nFicSal
'        Print #nFicSal, lsboleta & Chr$(12)
'        Print #nFicSal, ""
'    Close #nFicSal
'
'    Open sLpt For Output As nFicSal
'        Print #nFicSal, lsboletaitf & Chr$(12)
'        Print #nFicSal, ""
'    Close #nFicSal
    
                '   vNmovnro = dclscap.GetnMovNro(sMovNro)
                '   Call clsServ.AgregaMovConvCobranza(nMovNro, lblCodI.Caption, "", 2, nMonto, sNombrePers + sDIPers)
    
    Dim dclscap As COMDCaptaGenerales.DCOMCaptaMovimiento, clserv As COMNCaptaServicios.NCOMCaptaServicios, vNmovnro As Long
    Set dclscap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        vNmovnro = dclscap.GetnMovNro(sMovNro)
    Set dclscap = Nothing
    
    If vNmovnro = 0 Then
        MsgBox "NO se grabó operación." & vbCrLf & "Consulte con el área de T.I.", vbOKOnly + vbExclamation, "AVISO"
        Exit Sub
    End If
    
    Set clserv = New COMNCaptaServicios.NCOMCaptaServicios
        Call clserv.GrabaMovCobranzas(vNmovnro, Trim(lblCodI.Caption), "", 1, nMonto, Trim(sNombrePers) & Space(100 - Len(Trim(sNombrePers))) & sDIPers)
    Set clserv = Nothing

    
    
    Dim sTipDep As String, sCodOpe As String
    Dim sModDep As String, sTipApe As String
    Dim sNomTit As String
    sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, "SOLES", "DOLARES")
    sCodOpe = Trim(nOperacion)
    sModDep = "DEPOSITO EFECTIVO"
    sTipApe = "DEPOSITO AHORROS CUENTAS DE CONVENIO"
    Dim clsTit As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Set clsTit = New COMNCaptaGenerales.NCOMCaptaGenerales
        sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
    Set clsTit = Nothing
    Set clsMant = Nothing
    Dim lsCadImp As String
Do
'    If sNroDoc <> "" Then
'        Select Case nTipoDoc
'            Case TpoDocCheque
'                If sCodCMAC <> "" Then
'                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos
'                Else
'                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos
'                End If
'            Case TpoDocNotaAbono
'                ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos
'        End Select
'    Else
'        If sCodCMAC <> "" Then
'            sTipApe = "DEPOSITO CMAC AHORROS"
'            ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos
'        Else
          
           
           clsServ.ImprimeBoletaConvenio "DEPOSITO AHORROS CTA CONVENIO", ImpreCarEsp(sModDep), sCodOpe, nMonto, sNomTit, sCuenta, CStr(pnExtracto), gdFecSis, sNomAge, gsCodUser, sLpt, Trim(sNombrePers), sDIPers, gbITFAplica, Val(Me.LblITF.Caption), gbITFAsumidoAho, IIf(Me.chkITFEfectivo.value = 0, gITFCobroCargo, gITFCobroEfectivo), lsBoleta, lsBoletaITF
           Set clsServ = Nothing
           
           If Trim(lsBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                     Print #nFicSal, lsBoleta
                     Print #nFicSal, ""
                Close #nFicSal
           End If
           
           If Trim(lsBoletaITF) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                     Print #nFicSal, lsBoletaITF
                     Print #nFicSal, ""
                Close #nFicSal
           End If
'        End If
'    End If
        
'    If pbITFAplica And pnITFValor > 0 Then
'        If pbITFAsumido Then
'            If psITFOperacion = gITFCobroEfectivo Then
'                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont
'            Else
'                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, "ASUMIDO POR LA CMAC-ICA", False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont
'            End If
'        Else
'            If psITFOperacion = gITFCobroEfectivo Then
'                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 1, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont
'
'            Else
'                fgITFImprimeBoleta sNomTit, pnITFValor, sGlosa, dFecSis, sLpt, nExtracto, , 2, Mid(sCuenta, 9, 1), sCuenta, , False, sNomAge, sCodUser, bImpreSaldos, lnSaldoDisp, lnSaldoCont
'            End If
'        End If
'    End If
'
Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo

'*************PARA IMPRESION DE BOLETAS
    
Set clsLav = Nothing
cmdCancelar_Click

End If

Exit Sub
ErrGraba:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub

'Private Function IniciaLavDinero() As String 'comentado JACA 20110225
Private Function IniciaLavDinero(ByVal loLavDinero As SICMACT.frmMovLavDinero)    'JACA 20110225
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nMonto As Double
Dim sCuenta As String
'For i = 1 To grdCliente.Rows - 1
 '   nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
  '  If nPersoneria = gPersonaNat Then
  '      If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            sPersCod = Me.lblCodI
            sNombre = Me.lblNombreI
            sDireccion = ""
            sDocId = Me.lblDICli
'            Exit For
     '   End If
  '  Else
   '     If nRelacion = gCapRelPersTitular Then
'            sPersCod = grdCliente.TextMatrix(i, 1)
'            sNombre = grdCliente.TextMatrix(i, 2)
'            sDireccion = grdCliente.TextMatrix(i, 4)
'            sDocId = grdCliente.TextMatrix(i, 5)
'            Exit For
    '    End If
   ' End If
'Next i
nMonto = txtMonto.value
sCuenta = vsCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
    'ALPA 20081009***************************************************************************
    'IniciaLavDinero = frmMovLavDinero.Inicia(sPerscod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "200204")
    'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "200204", , , , , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen) 'comenato x JACA 20110225
     IniciaLavDinero = loLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, "200204", , , , , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen) 'JACA 20110225
    '****************************************************************************************
'End If
End Function



Private Sub cmdsalir_Click()
Unload Me
End Sub



Private Sub Form_Load()
    Me.Caption = "Convenio - Instituciones Otros Ingresos "
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub grdCuentas_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
 Dim nmoneda As Long, clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento, clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
 Dim sMsg As String, rsCta As ADODB.Recordset
 Set rsCta = New ADODB.Recordset
 
 
 Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(grdCuentas.TextMatrix(pnRow, 1), True)
 Set clsCap = Nothing
 
 Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
  
  If grdCuentas.TextMatrix(pnRow, 4) = "." Then
            vsCuenta = grdCuentas.TextMatrix(pnRow, 1)
            Set rsCta = clsMant.GetDatosCuenta(vsCuenta)
        Set clsMant = Nothing
                
        If sMsg <> "" Then
            grdCuentas.TextMatrix(pnRow, 4) = "0"
            MsgBox sMsg, vbInformation, "Operacion"
            Exit Sub
        End If
        lbITFCtaExonerada = fgITFVerificaExoneracion(vsCuenta)
        fgITFParamAsume Mid(vsCuenta, 4, 2), Mid(vsCuenta, 6, 3)
                
        If Not rsCta.EOF Then
            nPersoneria = rsCta("npersoneria").value
            pbOrdPag = rsCta("bOrdPag").value
        End If
                
        nmoneda = CLng(Mid(grdCuentas.TextMatrix(pnRow, 1), 9, 1))
        If nmoneda = gMonedaNacional Then
'            sMoneda = "MONEDA NACIONAL"
            txtMonto.BackColor = &HC0FFFF
            lblMon.Caption = "S/."
            LblITF.BackColor = &HC0FFFF
            lblTotal.BackColor = &HC0FFFF
        Else
'            sMoneda = "MONEDA EXTRANJERA"
            txtMonto.BackColor = &HC0FFC0
            LblITF.BackColor = &HC0FFC0
            lblTotal.BackColor = &HC0FFC0
            lblMon.Caption = "$"
        End If
        
        fraInstitucion.Enabled = False
        fraCobranza.Enabled = True
 End If

End Sub

Private Sub grdCuentas_OnRowChange(pnRow As Long, pnCol As Long)
Dim nmoneda As Long
 lbITFCtaExonerada = fgITFVerificaExoneracion(grdCuentas.TextMatrix(pnRow, 1))
        fgITFParamAsume Mid(grdCuentas.TextMatrix(pnRow, 1), 4, 2), Mid(grdCuentas.TextMatrix(pnRow, 1), 6, 3)
            
        Me.chkITFEfectivo.value = 0
        If gbITFAsumidoAho Then
                Me.chkITFEfectivo.Visible = False
                
        Else
                Me.chkITFEfectivo.Visible = True
        End If
        

        nmoneda = CLng(Mid(grdCuentas.TextMatrix(pnRow, 1), 9, 1))
        If nmoneda = gMonedaNacional Then
'            sMoneda = "MONEDA NACIONAL"
            txtMonto.BackColor = &HC0FFFF
            lblMon.Caption = "S/."
            LblITF.BackColor = &HC0FFFF
            lblTotal.BackColor = &HC0FFFF
        Else
'            sMoneda = "MONEDA EXTRANJERA"
            txtMonto.BackColor = &HC0FFC0
            LblITF.BackColor = &HC0FFC0
            lblTotal.BackColor = &HC0FFC0
            lblMon.Caption = "$"
        End If

End Sub

Private Sub opt_Click(Index As Integer)
 Select Case Index
    Case 0
       txtDI.Text = ""
       txtNombre.Text = ""
       Me.FraReferencia.Visible = False
       Me.fraClente.Visible = True
           
           
    Case 1
       lblDICli.Caption = ""
       Me.lbllCodigoCli.Caption = ""
       Me.LblNomCli.Caption = ""
       Me.fraClente.Visible = False
       Me.FraReferencia.Visible = True
 End Select
End Sub

Private Sub txtDI_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And Trim(txtDI.Text) <> "" Then
     Me.txtNombre.SetFocus
 ElseIf KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
    KeyAscii = 0
 End If
End Sub

Private Sub txtMonto_Change()
    If gbITFAplica Then       'Filtra para CTS
        If txtMonto.value > gnITFMontoMin Then
            If Not lbITFCtaExonerada Then
                Me.LblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                '*** BRGO 20110908 ************************************************
                nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblITF.Caption))
                If nRedondeoITF > 0 Then
                   Me.LblITF.Caption = Format(CCur(Me.LblITF.Caption) - nRedondeoITF, "#,##0.00")
                End If
                '*** END BRGO
            Else
                Me.LblITF.Caption = "0.00"
            End If
            
                If gbITFAsumidoAho Then
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                    Exit Sub
                ElseIf chkITFEfectivo.value = vbChecked Then
                    Me.lblTotal.Caption = Format(CCur(txtMonto.Text) + CCur(Me.LblITF.Caption), "#,##0.00")
                    Exit Sub
                Else
                    Me.lblTotal.Caption = Format(CCur(txtMonto.Text) - CCur(LblITF.Caption), "#,##0.00")
                    Exit Sub
                End If
        End If
    
    End If
    
    
    If txtMonto.value = 0 Then
        Me.LblITF.Caption = "0.00"
        Me.lblTotal.Caption = "0.00"
    End If
    chkITFEfectivo_Click
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
 KeyAscii = fgIntfMayusculas(KeyAscii)
 If KeyAscii = 13 Then
        txtMonto.SetFocus
 End If
End Sub
