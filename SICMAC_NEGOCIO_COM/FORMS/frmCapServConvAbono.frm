VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCapServConvAbono 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5430
   ClientLeft      =   4140
   ClientTop       =   1860
   ClientWidth     =   6135
   Icon            =   "frmCapServConvAbono.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAlumno 
      Caption         =   "Datos Alumno"
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
      Height          =   1455
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Width           =   5955
      Begin MSMask.MaskEdBox txtCodigo 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Código :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblEtqGrado 
         AutoSize        =   -1  'True
         Caption         =   "Grado :"
         Height          =   195
         Left            =   2700
         TabIndex        =   20
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label lblEtqSeccion 
         AutoSize        =   -1  'True
         Caption         =   "Sección :"
         Height          =   195
         Left            =   4080
         TabIndex        =   19
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblEtqNivel 
         AutoSize        =   -1  'True
         Caption         =   "Nivel :"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label lblNombre 
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
         Left            =   960
         TabIndex        =   17
         Top             =   660
         Width           =   4575
      End
      Begin VB.Label lblNivel 
         Alignment       =   2  'Center
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
         Left            =   960
         TabIndex        =   16
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label lblGrado 
         Alignment       =   2  'Center
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
         Left            =   3300
         TabIndex        =   15
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblSeccion 
         Alignment       =   2  'Center
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
         Left            =   4860
         TabIndex        =   14
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblCondicion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3360
         TabIndex        =   13
         Top             =   300
         Width           =   2115
      End
   End
   Begin VB.Frame fraPago 
      Caption         =   "Pago"
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
      Height          =   3375
      Left            =   60
      TabIndex        =   7
      Top             =   1500
      Width           =   5955
      Begin SICMACT.FlexEdit grdPago 
         Height          =   1695
         Left            =   180
         TabIndex        =   3
         Top             =   1200
         Width           =   5655
         _extentx        =   9975
         _extenty        =   2990
         cols0           =   5
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Cuenta-Tipo-Monto-nTipoCuenta"
         encabezadosanchos=   "350-1800-1600-1500-0"
         font            =   "frmCapServConvAbono.frx":030A
         font            =   "frmCapServConvAbono.frx":0332
         font            =   "frmCapServConvAbono.frx":035A
         font            =   "frmCapServConvAbono.frx":0382
         font            =   "frmCapServConvAbono.frx":03AA
         fontfixed       =   "frmCapServConvAbono.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-3-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-R-C"
         formatosedit    =   "0-0-0-2-0"
         textarray0      =   "#"
         lbeditarflex    =   -1  'True
         colwidth0       =   345
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   375
         Left            =   3900
         TabIndex        =   2
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTotalME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   2460
         TabIndex        =   27
         Top             =   2940
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Días Atraso :"
         Height          =   195
         Left            =   2820
         TabIndex        =   26
         Top             =   810
         Width           =   930
      End
      Begin VB.Label lblDiasAtraso 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   3900
         TabIndex        =   25
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label lblCuota 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   720
         TabIndex        =   24
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cuota :"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   810
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   390
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento :"
         Height          =   195
         Left            =   2880
         TabIndex        =   10
         Top             =   390
         Width           =   960
      End
      Begin VB.Label lblTotalMN 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   9
         Top             =   2940
         Width           =   1500
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL :"
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
         Left            =   960
         TabIndex        =   8
         Top             =   2940
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   4980
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   4980
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   4980
      Width           =   1000
   End
End
Attribute VB_Name = "frmCapServConvAbono"
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
Dim aPlanPago() As PlanPago
Dim nInstConvenio As CaptacConvenios
Dim nPos As Integer

Private Type PlanPago
    dVencimiento As Date
    nMontoPago As Double
    nMontoGasto As Double
    bMora As Boolean
    bFeriado As Boolean
    nCuota As Long
End Type

Private Sub GetTotalPagar()
Dim i As Long
Dim nMontoMN As Double, nMontoME As Double
nMontoMN = 0
nMontoME = 0
For i = 1 To grdPago.Rows - 1
    If CLng(Mid(grdPago.TextMatrix(i, 1), 9, 1)) = gMonedaNacional Then
        nMontoMN = nMontoMN + CDbl(grdPago.TextMatrix(i, 3))
    Else
        nMontoME = nMontoME + CDbl(grdPago.TextMatrix(i, 3))
    End If
Next i
lblTotalMN = "S/. " & Format$(nMontoMN, "#,##0.00")
lblTotalME = "US$ " & Format$(nMontoME, "#,##0.00")
End Sub

Private Sub ClearScreen()
fraAlumno.Enabled = True
Select Case nInstConvenio
    Case gCapConvJuanPabloInst
        txtCodigo.Mask = "####"
        txtCodigo.Text = "____"
    Case gCapConvJuanPabloII
        txtCodigo.Mask = "#####"
        txtCodigo.Text = "_____"
    Case gCapConvNarvaez
        txtCodigo.Mask = "C###-##"
        txtCodigo.Text = "____-__"
    Case gCapConvMarianoSantos, gCapConvSantaRosa
        txtCodigo.Mask = "C##"
        txtCodigo.Text = "___"
End Select
lblNombre = ""
lblNivel = ""
lblSeccion = ""
lblGrado = ""
lblCondicion = ""
lblTotalMN = "S/. 0.00"
lblTotalME = "US$ 0.00"
txtFecha = "__/__/____"
GetMesesAño
cmdCancelar.Enabled = False
cmdGrabar.Enabled = False
fraPago.Enabled = False
End Sub

Private Sub CalculaMora()
Dim nDias As Integer
Dim nFeriados As Long
Dim clsGen As COMDConstSistema.DCOMGeneral
If aPlanPago(nPos).bMora Then
    If aPlanPago(nPos).bFeriado Then
        Set clsGen = New COMDConstSistema.DCOMGeneral
        nFeriados = clsGen.GetNumDiasFeriado(CDate(txtFecha.Text), gdFecSis)
        Set clsGen = Nothing
    Else
        nFeriados = 0
    End If
    nDias = DateDiff("d", CDate(txtFecha), gdFecSis) - nFeriados
    If nDias > 0 Then
        Select Case nInstConvenio
            Case gCapConvNarvaez, gCapConvSantaRosa
                grdPago.TextMatrix(3, 3) = Format$(nMoraDia * nDias, "#,##0.00")
            Case gCapConvJuanPabloII, gCapConvJuanPabloInst, gCapConvMarianoSantos
                grdPago.TextMatrix(2, 3) = Format$(nMoraDia * nDias, "#,##0.00")
        End Select
        lblDiasAtraso = Format$(nDias, "#0")
    Else
        Select Case nInstConvenio
            Case gCapConvNarvaez, gCapConvSantaRosa
                grdPago.TextMatrix(3, 3) = "0.00"
            Case gCapConvJuanPabloII, gCapConvJuanPabloInst, gCapConvMarianoSantos
                grdPago.TextMatrix(2, 3) = "0.00"
        End Select
        lblDiasAtraso = "0"
    End If
End If
End Sub

Private Function GetPosicionMes(ByVal nMes As Integer) As Integer
Dim i As Integer
Dim dReferencia As Date
dReferencia = CDate("01/" & nMes & "/" & Year(gdFecSis))
For i = 1 To UBound(aPlanPago, 1)
    If DateDiff("m", aPlanPago(i).dVencimiento, dReferencia) = 0 Then
        GetPosicionMes = i
        Exit For
    End If
Next i
End Function

Private Sub GetMesesAño()
cboMes.Clear
cboMes.AddItem "Enero" & Space(50) & "01"
cboMes.AddItem "Febrero" & Space(50) & "02"
cboMes.AddItem "Marzo" & Space(50) & "03"
cboMes.AddItem "Abril" & Space(50) & "04"
cboMes.AddItem "Mayo" & Space(50) & "05"
cboMes.AddItem "Junio" & Space(50) & "06"
cboMes.AddItem "Julio" & Space(50) & "07"
cboMes.AddItem "Agosto" & Space(50) & "08"
cboMes.AddItem "Septiembre" & Space(50) & "09"
cboMes.AddItem "Octubre" & Space(50) & "10"
cboMes.AddItem "Noviembre" & Space(50) & "11"
cboMes.AddItem "Diciembre" & Space(50) & "12"
End Sub

Private Sub ObtieneDatosAlumno(ByVal sCodigo As String)
Dim sValor() As String
Dim i As Integer, nPos As Integer
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Dim rsAlu As Recordset

Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
Set rsAlu = clsServ.GetConvenioAlumno(nInstConvenio, sCodigo)
Set clsServ = Nothing
If rsAlu.EOF And rsAlu.BOF Then
    MsgBox "Alumno NO Encontrado", vbInformation, "Aviso"
Else
    Dim sCad As String, sNivel As String, sCondicion As String
    sCad = Trim(rsAlu("cReferencia"))
    lblNombre = rsAlu("cNombre")
    If nInstConvenio <> gCapConvSantaRosa Then
        i = 0
        nPos = 1
        Do While sCad <> "" And nPos > 0
            nPos = InStr(1, sCad, "-", vbTextCompare)
            i = i + 1
            ReDim Preserve sValor(i)
            If nPos > 1 Then
                sValor(i) = Left(sCad, nPos - 1)
                sCad = Mid(sCad, nPos + 1, Len(sCad) - nPos)
            ElseIf nPos = 1 Then
                sValor(i) = ""
                sCad = Mid(sCad, 2, Len(sCad) - 1)
            Else
                sValor(i) = sCad
                sCad = ""
            End If
        Loop
        sNivel = sValor(3)
        If nInstConvenio = gCapConvJuanPabloInst Then
            lblEtqGrado = "Ciclo :"
            If sNivel = "10" Then
                lblNivel = "INICIAL"
            ElseIf sNivel = "20" Then
                lblNivel = "PRIMARIA"
            ElseIf sNivel = "3L" Then
                lblNivel = "LENGUA Y LITERATURA"
            ElseIf sNivel = "3H" Then
                lblNivel = "HISTORIA Y RELIGION"
            ElseIf sNivel = "35" Then
                lblNivel = "EDUCACION FISICA"
            ElseIf sNivel = "36" Then
                lblNivel = "MATEMATICA"
            End If
        Else
            lblEtqGrado = "Grado :"
            If sNivel = "P" Then
                lblNivel = "PRIMARIA"
            ElseIf sNivel = "S" Then
                lblNivel = "SECUNDARIA"
            ElseIf sNivel = "I" Then
                lblNivel = "INICIAL"
            End If
        End If
        
        lblGrado = sValor(1) & "°"
        lblSeccion = IIf(sValor(2) = "U", "UNICA", sValor(2))
        sCondicion = ""
        If UBound(sValor(), 1) = 4 Then sCondicion = sValor(4)
        If sCondicion = "B" Then
            lblCondicion = "BECADO"
        ElseIf sCondicion = "M" Then
            lblCondicion = "MEDIA BECA"
        ElseIf sCondicion = "C" Then
            lblCondicion = "CUARTO BECA"
        Else
            lblCondicion = ""
        End If
    Else
        lblNivel = sCad
    End If
    fraAlumno.Enabled = False
    fraPago.Enabled = True
    cboMes.SetFocus
    cmdCancelar.Enabled = True
    cmdGrabar.Enabled = True
    cboMes.ListIndex = Month(gdFecSis) - 1
End If
rsAlu.Close
Set rsAlu = Nothing
End Sub

Public Sub Inicia(ByVal nConvenio As CaptacConvenios)
nInstConvenio = nConvenio
lblNivel.Visible = True
If nInstConvenio = gCapConvSantaRosa Then
    lblEtqNivel = "Doc Id :"
    lblEtqGrado.Visible = False
    lblGrado.Visible = False
    lblEtqSeccion.Visible = False
    lblSeccion.Visible = False
Else
    lblEtqNivel = "Nivel :"
    lblEtqGrado.Visible = True
    lblGrado.Visible = True
    lblEtqSeccion.Visible = True
    lblSeccion.Visible = True
End If
ClearScreen
GetCuentasAbono
GetMesesAño
cmdCancelar.Enabled = False
cmdGrabar.Enabled = False
Me.Show 1
End Sub

Private Sub GetCuentasAbono()
Dim rsCta As Recordset
Dim i As Integer, J As Integer
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
Set rsCta = clsServ.GetServConvCuentas(, nInstConvenio)
If rsCta.EOF And rsCta.BOF Then
    MsgBox "No están registradas las cuentas de abono para esta institución.", vbExclamation, "Aviso"
    fraAlumno.Enabled = False
    fraPago.Enabled = False
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
Else
    'Obtienes los datos de las cuentas a las cuales se realizará el abono
    Do While Not rsCta.EOF
        If rsCta("nTpoCuenta") = gCapConvTpoCtaMora Then
            nMoraDia = rsCta("nMontoMoraDia")
        End If
        grdPago.AdicionaFila
        If CLng(Mid(rsCta("cCtaCod"), 9, 1)) = gMonedaExtranjera Then
            grdPago.BackColorRow &HC0FFC0
        Else
        End If
        grdPago.TextMatrix(grdPago.Row, 1) = rsCta("cCtaCod")
        grdPago.TextMatrix(grdPago.Row, 2) = rsCta("cConsDescripcion")
        grdPago.TextMatrix(grdPago.Row, 3) = "0.00"
        grdPago.TextMatrix(grdPago.Row, 4) = rsCta("nTpoCuenta")
        rsCta.MoveNext
    Loop
    Set rsCta = clsServ.GetServPlanPagos(, nInstConvenio)
    i = 0
    Do While Not rsCta.EOF
        i = i + 1
        ReDim Preserve aPlanPago(i)
        aPlanPago(i).dVencimiento = rsCta("dFecVenc")
        aPlanPago(i).nMontoPago = rsCta("nMontoCuota")
        aPlanPago(i).nMontoGasto = rsCta("nMontoGasto")
        aPlanPago(i).bMora = rsCta("bAfectoMora")
        aPlanPago(i).bFeriado = rsCta("bAfectoFeriado")
        aPlanPago(i).nCuota = rsCta("nNroCuota")
        rsCta.MoveNext
    Loop
End If
rsCta.Close
Set rsCta = Nothing
End Sub

Private Sub cboMes_Click()
Dim dFecVenc As Date

nPos = GetPosicionMes(CInt(Right(cboMes.Text, 2)))
dFecVenc = aPlanPago(nPos).dVencimiento
dFecVenc = DateAdd("yyyy", DateDiff("yyyy", dFecVenc, gdFecSis), dFecVenc)
txtFecha = Format$(dFecVenc, "dd/mm/yyyy")
lblCuota = Format$(aPlanPago(nPos).nCuota, "00")
If nInstConvenio = gCapConvJuanPabloII Or nInstConvenio = gCapConvJuanPabloInst Then
    If lblCondicion = "MEDIA BECA" Then
        grdPago.TextMatrix(1, 3) = Format$(aPlanPago(nPos).nMontoGasto, "#,##0.00")
    ElseIf lblCondicion = "CUARTO BECA" Then
        grdPago.TextMatrix(1, 3) = Format$(aPlanPago(nPos).nMontoPago * 0.752, "#,##0.00")
    ElseIf lblCondicion = "" Then
        grdPago.TextMatrix(1, 3) = Format$(aPlanPago(nPos).nMontoPago, "#,##0.00")
    Else
        grdPago.TextMatrix(1, 3) = "0.00"
    End If
ElseIf nInstConvenio = gCapConvMarianoSantos Then
    If lblNivel = "PRIMARIA" Or lblNivel = "INICIAL" Then
        grdPago.TextMatrix(1, 3) = Format$(aPlanPago(nPos).nMontoPago, "#,##0.00")
    ElseIf lblNivel = "SECUNDARIA" Then
        grdPago.TextMatrix(1, 3) = Format$(aPlanPago(nPos).nMontoGasto, "#,##0.00")
    Else
        grdPago.TextMatrix(1, 3) = "0.00"
    End If
ElseIf nInstConvenio = gCapConvSantaRosa Then
    grdPago.TextMatrix(1, 3) = Format$(aPlanPago(nPos).nMontoPago, "#,##0.00")
    grdPago.TextMatrix(2, 3) = Format$(aPlanPago(nPos).nMontoGasto, "#,##0.00")
End If
CalculaMora
GetTotalPagar
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtFecha.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
ClearScreen
txtCodigo.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim nMontoMN As Double, nMontoME As Double
Dim sCodigo As String, sNombreAlumno As String * 30
Dim sMes As String
Dim nCuota As Integer
Dim bReImp As Boolean

'Leo los datos para impresion
nCuota = CInt(lblCuota)
sMes = Format$(CDate(txtFecha), "mmm-yyyy")
sNombreAlumno = ImpreCarEsp(Trim(lblNombre))
sCodigo = Trim(txtCodigo)
sCodigo = Replace(sCodigo, "_", "", 1, , vbTextCompare)

nMontoMN = CDbl(Replace(lblTotalMN, "S/.", "", 1, , vbTextCompare))
nMontoME = CDbl(Replace(lblTotalME, "US$", "", 1, , vbTextCompare))
'Valida de que el monto por Abono sea mayor que cero
If nMontoMN = 0 And nMontoME = 0 Then
    MsgBox "Monto de Abono deber ser mayo a cero", vbInformation, "Aviso"
    Exit Sub
End If

'Inicia el proceso de grabación
If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim rsPago As Recordset
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    
    Set rsPago = grdPago.GetRsNew()
    
'clsCap.CapAbonoServConvenio sMovNro, sCodigoPers, rsPago, gsNomAge
'                ImprimeBoleta "DEPOSITO AHORRO PENSION", "Depósito Efectivo", gsACDepEfe, Trim(nMontoPension), sTitularPension, sCuentaPension, sMes, nSaldo, DameInteresAcumulado(sCuentaPension), "Mes Pago", lnSNumExt, gnSaldCnt, False, False, , , , True, , sCodigo & " " & sNombreAlumno
'                If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'                    bReImp = True
'                Else
'                    bReImp = False
'                End If
'            Loop Until Not bReImp
'        End If
'    End If
'    If nMontoDonacion > 0 Then
'        nSaldo = ACDepositoEfectivo(sCuentaDonacion, Trim(nMontoDonacion), sCodDonacion & sCodigo, gsACDepEfe, gsACOADepEfe, gsACOTDepEfe, gsCodAge, False, , , , , Trim(nMes))
'        If gbOpeOk Then
'            bReImp = False
'            Do
'                ImprimeBoleta "DEPOSITO AHORRO DONACION", "Depósito Efectivo", gsACDepEfe, Trim(nMontoDonacion), sTitularDonacion, sCuentaDonacion, sMes, nSaldo, DameInteresAcumulado(sCuentaDonacion), "Mes Pago", lnSNumExt, gnSaldCnt, False, False, , , , True, , sCodigo & " " & sNombreAlumno
'                If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'                    bReImp = True
'                Else
'                    bReImp = False
'                End If
'            Loop Until Not bReImp
'        End If
'    End If
'    If nMontoMora > 0 Then
'        nSaldo = ACDepositoEfectivo(sCuentaPension, Trim(nMontoMora), sCodMora & sCodigo, gsACDepEfe, gsACOADepEfe, gsACOTDepEfe, gsCodAge, False, , , , , Trim(nMes))
'        If gbOpeOk Then
'            bReImp = False
'            Do
'                ImprimeBoleta "DEPOSITO AHORRO MORA", "Depósito Efectivo", gsACDepEfe, Trim(nMontoMora), sTitularPension, sCuentaPension, sMes, nSaldo, DameInteresAcumulado(sCuentaPension), "Mes Pago", lnSNumExt, gnSaldCnt, False, False, , , , True, , sCodigo & " " & sNombreAlumno
'                If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'                    bReImp = True
'                Else
'                    bReImp = False
'                End If
'            Loop Until Not bReImp
'        End If
'    End If
'    ClearScreen
'    txtCodigo.SetFocus
'    If Left(sCuentaPension, 2) <> Right(gsCodAge, 2) Then
'        CierraConeccion
'    End If
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Convenio - Pagos "
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub grdPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grdPago.Row >= grdPago.Rows - 1 Then
        cmdGrabar.SetFocus
    Else
        grdPago.Row = grdPago.Row + 1
    End If
    Exit Sub
End If
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCodigo As String
    sCodigo = Trim(txtCodigo)
    sCodigo = Replace(sCodigo, "_", "", 1, , vbTextCompare)
    ObtieneDatosAlumno sCodigo
End If
End Sub

Private Sub txtFecha_Change()
If IsDate(txtFecha) Then
    CalculaMora
End If
End Sub

Private Sub txtFecha_GotFocus()
With txtFecha
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdPago.Row = 1
    grdPago.Col = 3
    grdPago.SetFocus
End If
End Sub




