VERSION 5.00
Begin VB.Form frmAjusteGramosDeOroEnCustodia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajustes para Gramos de Oro en Custodia"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmAjusteGramosDeOroEnCustodia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "&Grabar Asiento de Ajuste"
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
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3540
      Width           =   4155
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   4155
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Mostrar Comparación Estadística - Contable"
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
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   4155
   End
   Begin VB.Frame frmMoneda 
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
      Height          =   2190
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4170
      Begin VB.OptionButton optMoneda 
         Caption         =   "A&justado"
         Height          =   255
         Index           =   3
         Left            =   4500
         TabIndex        =   9
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2220
         MaxLength       =   16
         TabIndex        =   8
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox txtTipCambioVenta 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2220
         MaxLength       =   16
         TabIndex        =   7
         Top             =   840
         Width           =   1425
      End
      Begin VB.TextBox txtTipCambioCompra 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2220
         MaxLength       =   16
         TabIndex        =   6
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cambio Fijo"
         Height          =   315
         Left            =   345
         TabIndex        =   12
         Top             =   300
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cambio Compra"
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   1920
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Cambio Venta"
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   1680
      End
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
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4170
      Begin VB.ComboBox CboMes 
         Height          =   315
         ItemData        =   "frmAjusteGramosDeOroEnCustodia.frx":030A
         Left            =   690
         List            =   "frmAjusteGramosDeOroEnCustodia.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   1830
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   1
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   390
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmAjusteGramosDeOroEnCustodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmAjusteGramosDeOroEnCustodia
'Descripcion:Formulario para Generar el Ajuste de Gramos de Oro en Custodia
'Creacion: PASI TI-ERS076-2015
'*****************************
Option Explicit
Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra As New clsProgressBar
Dim rsAjuste As ADODB.Recordset
Public Sub inicio(ByVal psOpeCod As String)
    Me.Show 0, frmMdiMain
End Sub
Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CboMes.ListIndex > -1 And txtAnio <> "" Then
        txtTipCambio = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 1)
        txtTipCambioVenta = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 2)
        txtTipCambioCompra = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 3)
        txtAnio.SetFocus
    End If
 End If
End Sub
Private Sub CboMes_Click()
    If CboMes.ListIndex > -1 And txtAnio <> "" Then
        txtTipCambio = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 1)
        txtTipCambioVenta = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 2)
        txtTipCambioCompra = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 3)
        txtTipCambioCompra.SetFocus
    End If
End Sub
Private Sub cmdAsiento_Click()
    Dim rs As ADODB.Recordset
    Dim nTotal As Currency
    Dim nItem As Integer
    Dim lTransActiva As Boolean
    Dim nMes As Integer, nAnio As Integer, dFecha As Date
    
    On Error GoTo AsientoError
    If MsgBox("¿ Seguro desea Grabar Asiento ? ", vbQuestion + vbYesNo + vbDefaultButton2, "¡Confirmación!") = vbNo Then
        Exit Sub
    End If
    nMes = CboMes.ListIndex + 1
    nAnio = txtAnio.Text
    dFecha = DateAdd("M", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio.Text, "0000")) - 1
    
    Dim oCont As New NContFunciones
    Dim oMov As New DMov
    Dim oAju As New DAjusteCont
    
    If Not oCont.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
        MsgBox "Mes ya cerrado. Imposible generar Asiento de Ajuste para Gramos de Oro en Custodia.", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    gsMovNro = Format(dFecha, "yyyymmdd")
    If oCont.ExisteMovimiento(Left(gsMovNro, 6), gsOpeCod) Then
        MsgBox "El Asiento para el periodo ya fue generado. ", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    Me.Enabled = False
    If rsAjuste.EOF And rsAjuste.BOF Then
        MsgBox "No ha sido posible generar los asientos contables.", vbInformation, "¡Aviso!"
    Else
        oImp_BarraShow rsAjuste.RecordCount
        oMov.BeginTrans
        lTransActiva = True
        
        gsGlosa = "ASIENTO DE AJUSTE PARA GRAMOS DE ORO EN CUSTODIA AL " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
        gsMovNro = oMov.GeneraMovNro(dFecha, gsCodAge, gsCodUser)
        oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
        gnMovNro = oMov.GetnMovNro(gsMovNro)
        nItem = 0
        Do While Not rsAjuste.EOF
            nItem = nItem + 1
            gnImporte = IIf(rsAjuste!nSaldoDebe > 0, rsAjuste!nSaldoDebe, rsAjuste!nSaldoHaber)
            oMov.InsertaMovCta gnMovNro, nItem, rsAjuste!cCtaContCod, gnImporte, Mid(gsOpeCod, 3, 1)
            oImp_BarraProgress rsAjuste.Bookmark, "ASIENTO DE AJUSTE PARA GRAMOS DE ORO EN CUSTODIA", "", "Grabando...", vbBlue
            rsAjuste.MoveNext
        Loop
        If dFecha < gdFecSis Then
            oMov.ActualizaSaldoMovimiento gsMovNro, "+"
        End If
        oMov.CommitTrans
        lTransActiva = False
        oImp_BarraClose
    End If
    ImprimeAsientoContable gsMovNro, , , , , , , , , , , , 1, "ASIENTO PARA GRAMOS DE ORO EN CUSTODIA"
Set oMov = Nothing
Set oCont = Nothing
Set oAju = Nothing
Me.Enabled = True
Exit Sub
AsientoError:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   Me.Enabled = True
   If lTransActiva Then
        oMov.RollbackTrans
   End If
End Sub
Private Sub cmdGenerar_Click()
    If Not ValidaDatos Then
        Exit Sub
    End If
    If MsgBox("¿ Seguro que desea generar Cuadro de Depreciación ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
        Exit Sub
    End If
    GeneraReporte
End Sub
Private Sub GeneraReporte()
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
Dim sImpre As String
Dim oCont As New NContFunciones
On Error GoTo ReDeprecAdjud
nMes = CboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1
Me.Enabled = False
sImpre = oImp.ImprimeCuadroAjusteParaGramosDeOroEnCustodia(dFecha, gnLinPage, rsAjuste)
EnviaPrevio sImpre, "CUADRO DE AJUSTE PARA GRAMOS DE ORO EN CUSTODIA", gnLinPage, False
Set oCont = Nothing
Me.Enabled = True
cmdAsiento.Enabled = True
cmdAsiento.SetFocus
Exit Sub
ReDeprecAdjud:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso¡!"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub txtAnio_Change()
cmdAsiento.Enabled = False
End Sub
Private Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub
Private Sub CboMes_Validate(Cancel As Boolean)
If CboMes.ListIndex <> Val(CboMes.Tag) Then
   cmdAsiento.Enabled = False
   CboMes.Tag = CboMes.ListIndex
End If
End Sub
Private Sub txtTipCambio_GotFocus()
    fEnfoque txtTipCambio
End Sub
Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub
Private Sub txtTipCambioVenta_GotFocus()
    fEnfoque txtTipCambioVenta
End Sub
Private Sub txtTipCambioVenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambioVenta, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub
Private Sub txtTipCambioCompra_GotFocus()
    fEnfoque txtTipCambioCompra
End Sub
Private Sub txtTipCambioCompra_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTipCambioCompra, KeyAscii, 14, 5)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Function ValidaDatos() As Boolean
ValidaDatos = False
If CboMes.ListIndex = -1 Then
   MsgBox "Debe seleccionarse mes de proceso", vbInformation, "!Aviso!"
   CboMes.SetFocus
   Exit Function
End If
If Not ValidaAnio(txtAnio) Then
   txtAnio.SetFocus
   Exit Function
End If
If nVal(txtTipCambio) = 0 Then
   MsgBox "Debe indicar Tipo de Cambio", vbInformation, "!Aviso!"
   txtTipCambio.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = TextBox_SoloNumeros(KeyAscii)
If KeyAscii = 13 Then
    If CboMes.ListIndex > -1 And txtAnio <> "" Then
        txtTipCambio = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 1)
        txtTipCambioVenta = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 2)
        txtTipCambioCompra = TiposCambiosCierreMensual(txtAnio, CboMes.ListIndex + 1, False, 3)
        txtTipCambioCompra.SetFocus
    End If
    txtTipCambioCompra.SetFocus
End If
End Sub
Private Sub Form_Load()
    CboMes.ListIndex = Month(gdFecSis) - 1
    txtAnio = Year(gdFecSis)
    Set oImp = New NContImprimir
    Set rsAjuste = New ADODB.Recordset
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
