VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFondoSeguroDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro y Distribución de gastos de Fondo Seguro de Depósito"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   Icon            =   "frmFondoSeguroDeposito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   7215
         Begin VB.TextBox txtTC 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   325
            Left            =   4620
            TabIndex        =   4
            Text            =   "0"
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txtTasa 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   325
            Left            =   1380
            TabIndex        =   3
            Text            =   "0"
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txtPrimaD 
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
            ForeColor       =   &H0000C000&
            Height          =   345
            Left            =   4620
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtPrimaS 
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
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1425
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   720
            Width           =   1515
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio :"
            Height          =   195
            Left            =   3540
            TabIndex        =   21
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tasa % :"
            Height          =   195
            Left            =   660
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tot. Dep. US$. :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   3180
            TabIndex        =   19
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tot. Dep. S/. :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   825
            Width           =   1275
         End
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "frmFondoSeguroDeposito.frx":030A
         Left            =   1935
         List            =   "frmFondoSeguroDeposito.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2340
      End
      Begin VB.CommandButton cmdexaminar 
         Caption         =   "E&xaminar"
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
         Left            =   6105
         TabIndex        =   14
         Top             =   240
         Width           =   1230
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   600
         TabIndex        =   1
         Top             =   270
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Trimestre:"
         Height          =   195
         Left            =   4440
         TabIndex        =   23
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lblNumTrimestre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   1440
         TabIndex        =   16
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame fracontrol 
      Height          =   585
      Left            =   75
      TabIndex        =   0
      Top             =   2040
      Width           =   7500
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
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
         Left            =   1005
         TabIndex        =   12
         Top             =   165
         Width           =   900
      End
      Begin VB.CommandButton cmdsalir 
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
         Height          =   345
         Left            =   6285
         TabIndex        =   9
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eli&minar"
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
         Left            =   1920
         TabIndex        =   10
         Top             =   165
         Width           =   915
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
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
         Left            =   5310
         TabIndex        =   8
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
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
         Left            =   60
         TabIndex        =   11
         Top             =   165
         Width           =   930
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
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
         Left            =   75
         TabIndex        =   7
         Top             =   165
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmFondoSeguroDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTipoOperacion As Integer '0 Nuevo...1 Modificar
Dim lnTotal As Double
Dim lsMovNro As String

Dim sMatrizDatos() As Variant
Dim nTipoMatriz As Integer
Dim nCantAgeSel As Integer

Sub Limpiar_Controles()
    lnTotal = 0
    
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
    Me.cmbMes.ListIndex = -1
    Me.lblNumTrimestre.Caption = ""
    Me.txtTasa.Text = "0.0000"
    Me.txtTC.Text = "0.0000"
    Me.txtPrimaS.Text = "0.00"
    Me.txtPrimaD.Text = "0.00"


'    txtNumCertif.Text = ""
'    LblAsegPersCod.Caption = ""
'    LblAsegPersNombre.Caption = ""
'    cboTipo.ListIndex = -1
'    cboMoneda.ListIndex = -1
'    txtSumaA.Text = "0.00"
'    Me.txtDel.Text = "__/__/____"
'    Me.txtAl.Text = "__/__/____"
'    Me.txtDescripcion.Text = ""
    
End Sub

Private Sub cmbMes_GotFocus()
 'fEnfoque cmbMes
End Sub

Private Sub cmbMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub cmbMes_LostFocus()
    Select Case Val(Trim(Right(cmbMes.Text, 2)))
        Case 1, 2, 3
            Me.lblNumTrimestre = "1"
        Case 4, 5, 6
            Me.lblNumTrimestre = "2"
        Case 7, 8, 9
            Me.lblNumTrimestre = "3"
        Case 10, 11, 12
            Me.lblNumTrimestre = "4"
        Case Else
            Me.lblNumTrimestre = ""
    End Select
End Sub

'Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub

'Private Sub cboTipo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub

'Private Sub cmdAsignaGasto_Click()
'
'Dim i As Integer, j As Integer
'
'    If cboTipo.ListIndex = -1 Then
'        MsgBox "Debe indicar el tipo de poliza", vbInformation, "Mensaje"
'        cboTipo.SetFocus
'        Exit Sub
'    End If
'    Call frmPoliSeguPatriDistriPorcen.Inicio(CInt(Trim(Right(cboTipo.Text, 20))), sMatrizDatos, nTipoMatriz, nCantAgeSel)
'
''    Dim matriz1() As Variant
''    Dim matriz2(1, 3) As Variant
''
''    For i = 1 To 3
''        For j = 1 To nCantAgeSel
''            matriz2(j, i) = sMatrizDatos(i, j)
''        Next j
''    Next i
'
'End Sub

'Private Sub cmdBuscaAseg_Click()
'Dim oPersona As UPersona
'Dim sPersCod As String
'
'    Set oPersona = frmBuscaPersona.Inicio
'    If Not oPersona Is Nothing Then
'        LblAsegPersCod.Caption = oPersona.sPersCod
'        LblAsegPersNombre.Caption = oPersona.sPersNombre
'    End If
'    Set oPersona = Nothing
'End Sub



'Private Sub cmdBuscaAseg_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub

Private Sub cmdCancelar_Click()
    
    Call Limpiar_Controles
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
    Limpiar_Controles

    Me.cmdNuevo.Enabled = True
    Me.cmdeditar.Enabled = False
    Me.cmdEliminar.Enabled = False
    
End Sub

Private Sub cmdEditar_Click()
nTipoOperacion = 1
Call Habilita_Grabar(True)
Call Habilita_Datos(True)

cmdGrabar.Enabled = True
cmdeditar.Enabled = False
cmdEliminar.Enabled = False
cmdCancelar.Enabled = True
lnTotal = 0
End Sub

Private Sub cmdEliminar_Click()

    Dim oContFunc As NContFunciones

    Dim oAge As DOperacion
    Dim overi As DOperacion
    Dim rs As ADODB.Recordset

    Dim lcFecPeriodo As String

    Dim lcMovNroExtorno As Integer
    Dim lnMovNroExtorno As Integer
    Dim lnMovNroExt As Long
    Dim dFecha As Date
    Dim nMES As Integer
    Dim nAnio As Integer
    
    lcFecPeriodo = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lcFecPeriodo = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lcFecPeriodo = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
    
    Set overi = New DOperacion
    Set rs = overi.VerificaAsientoCont(gsOpeCod, lcFecPeriodo)
    Set overi = Nothing

    If (rs.EOF And rs.BOF) Then
        MsgBox "Este periodo no tiene asiento contable.", vbCritical, "Aviso!"
        Exit Sub
    Else
        lnMovNroExt = rs!nMovNro
    End If

    If MsgBox("¿Esta seguro que desea eliminar este registro y su asiento contable? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub

    nMES = Val(Trim(Right(cmbMes.Text, 2)))
    nAnio = Val(mskAnio.Text)
    dFecha = DateAdd("m", 1, "01/" & Format(nMES, "00") & "/" & Format(nAnio, "0000")) - 1
    Set oContFunc = New NContFunciones
    If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
       Set oContFunc = Nothing
       MsgBox "Imposible realizar este proceso ya que la fecha ingresada pertenece a un mes cerrado.", vbInformation, "Aviso"
       Exit Sub
    End If

    Set oAge = New DOperacion
        Call oAge.AnulaRegistroFSD(nAnio, nMES)
        Call oAge.AnulaAsientoContFSD(lcFecPeriodo)
    Set oAge = Nothing
    
    MsgBox "El registro y el asiento contable del FSD fue eliminado satisfactoriamente.", vbInformation, "Aviso"

    Call Limpiar_Controles
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
    Limpiar_Controles

    Me.cmdNuevo.Enabled = True
    Me.cmdeditar.Enabled = False
    Me.cmdEliminar.Enabled = False

End Sub

Private Sub cmdexaminar_Click()

Dim oPol As DOperacion
Dim rs As ADODB.Recordset
Dim lcDatos As String

Set oPol = New DOperacion

lcDatos = "0"

Call frmFondoSeguDepositoListado.Inicio(0)

Set rs = oPol.CargaDatosFondoSeguDepo(frmFondoSeguDepositoListado.nAnio, frmFondoSeguDepositoListado.nMES)
If Not rs.EOF Then
    
    Me.mskAnio = rs!nAnio
    Me.cmbMes.ListIndex = IndiceListaCombo(cmbMes, Trim(Str(rs!nMES)))
    Me.lblNumTrimestre.Caption = rs!nTrimestre
    Me.txtTasa.Text = rs!ntasa
    Me.txtTC.Text = rs!nTC
    Me.txtPrimaS.Text = rs!nTotDepMN
    Me.txtPrimaD.Text = rs!nTotDepME
    
    lcDatos = "1"
    
    Call Habilita_Datos(False)
    Me.mskAnio.Enabled = False
    Me.cmbMes.Enabled = False
    
End If
Set oPol = Nothing

If lcDatos = "1" Then
    Me.cmdeditar.Enabled = True
    Me.cmdEliminar.Enabled = True
    Me.cmdNuevo.Enabled = False
    
End If

End Sub

Private Sub CmdGrabar_Click()

Dim oCont As NContFunciones
Dim oPol As DOperacion
Dim i As Integer, j As Integer
Dim nMES As Integer, nAnio As Integer
Dim lsFecha As String
Dim rs As ADODB.Recordset
Dim lnTCPondVentaDebe As Double, lnTCFijoHaber As Double
Dim lnPrimaSol As Double, lnPrimaDol As Double
Dim dFecha As Date

If Valida_Datos = False Then Exit Sub

nMES = Val(Trim(Right(cmbMes.Text, 2)))
nAnio = Val(mskAnio.Text)
dFecha = DateAdd("m", 1, "01/" & Format(nMES, "00") & "/" & Format(nAnio, "0000")) - 1

If nTipoOperacion = 0 Then ''SI NO ES MODIFICACION
    Set oPol = New DOperacion
    Set rs = oPol.CargaDatosFondoSeguDepo(nAnio, nMES)
    If Not rs.EOF Then
        MsgBox "La prima de FSD de este periodo ya fue registrado.", vbOKOnly, "Atención"
        Exit Sub
    End If
End If

If MsgBox("¿Esta seguro de registrar los datos ingresados?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Set oCont = New NContFunciones
    If Not oCont.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
       Set oCont = Nothing
       MsgBox "Imposible realizar este proceso ya que la fecha ingresada pertenece a un mes cerrado.", vbInformation, "Aviso"
       Exit Sub
    End If

    Set oPol = New DOperacion
    Set oCont = New NContFunciones
    
    lsMovNro = oCont.GeneraMovNro(dFecha, gsCodAge, gsCodUser)
    
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If

'    lnPrimaSol = oPol.CargaPrima(nAnio, Val(lblNumTrimestre), nMES, 1)
'    lnPrimaDol = oPol.CargaPrima(nAnio, Val(lblNumTrimestre), nMES, 2)

    lnTCPondVentaDebe = oPol.CargaTipoCambioDelDia(Format(dFecha, "yyyymmdd"), 1)
    lnTCFijoHaber = oPol.CargaTipoCambioDelDia(Format(dFecha, "yyyymmdd"), 2)

    If nTipoOperacion = 0 Then
        Call oPol.InsertaFondoSeguDepo(nAnio, nMES, Val(lblNumTrimestre), CDbl(Me.txtTasa), CDbl(Me.txtTC), CDbl(Me.txtPrimaS), CDbl(Me.txtPrimaD), gsOpeCod, lsFecha, dFecha, gsCodAge, gsCodUser, lnTCPondVentaDebe, lnTCFijoHaber, lsMovNro)
        Call oPol.GeneraAsientoContableFSD(nAnio, nMES, Val(lblNumTrimestre), CDbl(Me.txtTasa), CDbl(Me.txtTC), CDbl(Me.txtPrimaS), CDbl(Me.txtPrimaD), gsOpeCod, lsFecha, dFecha, gsCodAge, gsCodUser, lnTCPondVentaDebe, lnTCFijoHaber, lsMovNro)
    Else
        Call oPol.ActualizaFondoSeguDepo(nAnio, nMES, Val(lblNumTrimestre), CDbl(Me.txtTasa), CDbl(Me.txtTC), CDbl(Me.txtPrimaS), CDbl(Me.txtPrimaD), lsMovNro)
        Call oPol.AnulaAsientoContFSD(lsFecha)
        Call oPol.GeneraAsientoContableFSD(nAnio, nMES, Val(lblNumTrimestre), CDbl(Me.txtTasa), CDbl(Me.txtTC), CDbl(Me.txtPrimaS), CDbl(Me.txtPrimaD), gsOpeCod, lsFecha, dFecha, gsCodAge, gsCodUser, lnTCPondVentaDebe, lnTCFijoHaber, lsMovNro)
    End If
    Set oPol = Nothing
    Set oCont = Nothing

    cmdGrabar.Enabled = False
    cmdeditar.Enabled = True
    cmdEliminar.Enabled = True
    cmdCancelar.Enabled = False
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
    Limpiar_Controles


'''*************----------------------------------
'    Dim oMov As DMov
'    Set oMov = New DMov
'
'    Dim oDep As DOperacion
'    Set oDep = New DOperacion
'
'    Dim oConect As DConecta
'    Set oConect = New DConecta
'
'    Dim lnMovNro As Long
'    Dim lsMovNro As String
'
'    Dim lsTipo As String
'
'    Dim i As Integer
'    Dim lnI As Long
'    Dim lnContador As Long
'    Dim lsCtaCont As String
'    Dim oPrevio As clsPrevioFinan
'    Dim oAsiento As NContImprimir
'    Dim nConta As Integer, lcCtaDif As String
'    Dim overi As DOperacion
'    Dim lnDebe As Double, lnHaber As Double, lnTotHaber As Double, lnTotDebe As Double
'    Dim lnDebeME As Double, lnTotDebeME As Double
'    Dim oContFunc As NContFunciones
'
'    Set oPrevio = New clsPrevioFinan
'    Set oAsiento = New NContImprimir
'
'    Dim rs As ADODB.Recordset, rs1 As ADODB.Recordset
'
'    Dim ldFechaDepre As Date
'    Dim ldFechaRegistro As Date
'
'
'
'
'    Dim lsFecha As String
'
'    gsOpeCod = "300440"
'
'    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
'
'    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
'        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
'    Else
'        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
'    End If
'
'    Set rs = New ADODB.Recordset
'
'    Set overi = New DOperacion
'    Set rs = overi.VerificaAsientoCont(gsOpeCod, lsFecha)
'    Set overi = Nothing
'
'    If Not rs.EOF Then
'        MsgBox "Asiento ya fue generado.", vbCritical, "Aviso!"
'        Exit Sub
'    End If
'
'    If MsgBox("¿Esta seguro de registrar los datos ingresados? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
'
'    nMES = Val(Trim(Right(cmbMes.Text, 2)))
'    nAnio = Val(mskAnio.Text)
'    dFecha = DateAdd("m", 1, "01/" & Format(nMES, "00") & "/" & Format(nAnio, "0000")) - 1
'    Set oContFunc = New NContFunciones
'    If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
'       Set oContFunc = Nothing
'       MsgBox "Imposible grabar ya que el registro pertenece a un mes cerrado.", vbInformation, "Aviso"
'       Exit Sub
'    End If
'
'    ldFechaRegistro = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
'
'    oMov.BeginTrans
'        lsMovNro = oMov.GeneraMovNro(ldFechaRegistro, Right(gsCodAge, 2), gsCodUser)
'
'        oMov.InsertaMov lsMovNro, gsOpeCod, "REG. " & Trim(Mid(Me.cboTpo.Text, 1, Len(Me.cboTpo.Text) - 2))
'
'        lnMovNro = oMov.GetnMovNro(lsMovNro)
'
'        nConta = 0: lnTotHaber = 0: lnHaber = 0: lnDebe = 0
'
'        For lnI = 1 To Me.FlexEdit1.Rows - 1
'            If Len(Trim(Me.FlexEdit1.TextMatrix(lnI, 1))) > 0 And Val(Me.FlexEdit1.TextMatrix(lnI, 4)) > 0 Then
'                nConta = nConta + 1
'                    lsCtaCont = oDep.ObtieneCtasPoliSeguPatri(gsOpeCod, "D", 0)
'                    lsCtaCont = Replace(lsCtaCont, "AG", Right(Me.FlexEdit1.TextMatrix(lnI, 2), 2))
'                    'lsCtaCont = Replace(lsCtaCont, "M", Trim(Me.FlexEdit1.TextMatrix(lnI, 24)))
'                    lsCtaCont = Replace(lsCtaCont, "M", "2")
'
'                    lnDebe = Round(Me.FlexEdit1.TextMatrix(lnI, 17), 2)
'                    lnHaber = Round(Me.FlexEdit1.TextMatrix(lnI, 16), 2)
'
'                    lnTotDebe = lnTotDebe + lnDebe
'                    lnTotHaber = lnTotHaber + lnHaber
'
'                    lnDebeME = Round(Me.FlexEdit1.TextMatrix(lnI, 10), 2)
'                    lnTotDebeME = lnTotDebeME + lnDebeME
'
'                    oMov.InsertaMovCta lnMovNro, nConta, lsCtaCont, lnDebe
'                    oMov.InsertaMovMe lnMovNro, nConta, lnDebeME
'            End If
'        Next lnI
'
'        If Me.FlexEdit1.Rows - 1 > 0 Then
'
'            nConta = nConta + 1
'
'            lsCtaCont = oDep.ObtieneCtasPoliSeguPatri(gsOpeCod, "H", 0)
'            'lsCtaCont = Replace(lsCtaCont, "AG", Right(Me.FlexEdit1.TextMatrix(lnI, 2), 2))
'            lsCtaCont = Replace(lsCtaCont, "M", "2")
'
'            oMov.InsertaMovCta lnMovNro, nConta, lsCtaCont, lnTotHaber * -1
'            oMov.InsertaMovMe lnMovNro, nConta, lnTotDebeME * -1
'
'            nConta = nConta + 1
'
'            If lnTotHaber > lnTotDebe Then
'                lcCtaDif = rs1!cCtaGasto
'                oMov.InsertaMovCta lnMovNro, nConta, lcCtaDif, (lnTotHaber - lnTotDebe)
'            ElseIf lnTotHaber < lnTotDebe Then
'                lcCtaDif = rs1!cCtaIngreso
'                oMov.InsertaMovCta lnMovNro, nConta, lcCtaDif, (lnTotDebe - lnTotHaber) * -1
'            End If
'
'        End If
'
'    oMov.CommitTrans
'
'    oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80, Caption), Caption, True

End Sub

Private Sub cmdImprimir_Click()

End Sub

Private Sub cmdNuevo_Click()
    nTipoOperacion = 0
    Call Limpiar_Controles
    Call Habilita_Grabar(True)
    Call Habilita_Datos(True)
    cmdGrabar.Enabled = True
    cmdeditar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    
'    Me.txtNumCertif.SetFocus
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
        
    Call Limpiar_Controles
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
    Limpiar_Controles

    Me.cmdNuevo.Enabled = True
    Me.cmdeditar.Enabled = False
    Me.cmdEliminar.Enabled = False
    
    Set rs = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
        
    Me.txtTasa.Text = "0.0000"
    Me.txtTC.Text = "0.0000"
   
    Me.txtTasa.Text = oConst.LeeConstSistema("59")
   
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
    
    gsOpeCod = "300440"

End Sub

Sub Habilita_Grabar(ByVal pbHabilita As Boolean)
    cmdGrabar.Visible = pbHabilita
    cmdNuevo.Visible = Not pbHabilita
End Sub

Private Sub Option1_Click()
'        LblNroCredito.Visible = False
'        txtNroCredito.Visible = False
'        cmdValidar.Visible = False
End Sub

Private Sub Option2_Click()
'    LblNroCredito.Visible = True
'    txtNroCredito.Visible = True
'    cmdValidar.Visible = True
End Sub

'Private Sub txtMonto_KeyPress(KeyAscii As Integer)
'     KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
'     If KeyAscii = 13 Then
'        cboTipo.SetFocus
'     End If
'End Sub

'Private Sub txtMonto_LostFocus()
'If txtMonto.Text = "" Then
'    txtMonto.Text = "0.00"
'Else
'    txtMonto.Text = Format(txtMonto.Text, "#0.00")
'End If
'End Sub

Function Valida_Datos() As Boolean
Valida_Datos = True

If Not IsNumeric(Me.mskAnio.Text) Then
    MsgBox "Debe Ingresar un año Valido.", vbInformation, "Aviso"
    Me.mskAnio.SetFocus
    Exit Function
ElseIf Me.cmbMes.Text = "" Then
    MsgBox "Debe Ingresar un mes Valido.", vbInformation, "Aviso"
    Me.cmbMes.SetFocus
    Exit Function
ElseIf CDbl(Me.txtTasa.Text) = 0 Then
    MsgBox "Debe indicar la tasa.", vbInformation, "Mensaje"
    txtTasa.SetFocus
    Valida_Datos = False
    Exit Function
ElseIf CDbl(Me.txtTC.Text) = 0 Then
    MsgBox "Debe indicar el Tipo de cambio.", vbInformation, "Mensaje"
    txtTC.SetFocus
    Valida_Datos = False
    Exit Function
ElseIf CDbl(Me.txtPrimaS.Text) = 0 Then
    MsgBox "Debe ingresar el Total depósito en MN.", vbInformation, "Mensaje"
    txtPrimaS.SetFocus
    Valida_Datos = False
    Exit Function
ElseIf CDbl(Me.txtPrimaD.Text) = 0 Then
    MsgBox "Debe ingresar el Total depósito en ME.", vbInformation, "Mensaje"
    txtPrimaD.SetFocus
    Valida_Datos = False
    Exit Function
End If


'If CDbl(txtMonto.Text) = 0 Then
'    MsgBox "Debe indicar el valor de la prima", vbInformation, "Mensaje"
'    txtMonto.SetFocus
'    Valida_Datos = False
'    Exit Function
'End If
'If LblContPersCod.Caption = "" Then
'    MsgBox "Debe indicar el contratante", vbInformation, "Mensaje"
'    Valida_Datos = False
'    cmdBuscaCont.SetFocus
'    Exit Function
'End If
'If LblAsegPersCod.Caption = "" Then
'    MsgBox "Debe indicar la aseguradora", vbInformation, "Mensaje"
'    Valida_Datos = False
'    cmdBuscaAseg.SetFocus
'    Exit Function
'End If

'If cboTipo.ListIndex = -1 Then
'    MsgBox "Debe indicar el tipo de poliza", vbInformation, "Mensaje"
'    Valida_Datos = False
'    cboTipo.SetFocus
'    Exit Function
'End If

'If cboMoneda.ListIndex = -1 Then
'    MsgBox "Debe indicar la moneda de la poliza", vbInformation, "Mensaje"
'    Valida_Datos = False
'    cboMoneda.SetFocus
'    Exit Function
'End If

'If CDbl(txtSumaA.Text) = 0 Then
'    MsgBox "Debe indicar la suma asegurada", vbInformation, "Mensaje"
'    Valida_Datos = False
'    txtSumaA.SetFocus
'End If

End Function

Private Sub Habilita_Datos(ByVal pbHabilita As Boolean)

If nTipoOperacion = 1 Then
    Me.txtTasa.Enabled = pbHabilita
    Me.txtTC.Enabled = pbHabilita
    Me.txtPrimaS.Enabled = pbHabilita
    Me.txtPrimaD.Enabled = pbHabilita
Else
    Me.mskAnio.Enabled = pbHabilita
    Me.cmbMes.Enabled = pbHabilita
    Me.txtTasa.Enabled = pbHabilita
    Me.txtTC.Enabled = pbHabilita
    Me.txtPrimaS.Enabled = pbHabilita
    Me.txtPrimaD.Enabled = pbHabilita
End If
    
End Sub

Private Sub mskAnio_GotFocus()
  fEnfoque mskAnio
End Sub

Private Sub mskAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtPrimaD_GotFocus()
    fEnfoque txtPrimaD
End Sub

Private Sub txtPrimaD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtPrimaS_GotFocus()
    fEnfoque txtPrimaS
End Sub

Private Sub txtPrimaS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

'Private Sub txtAl_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub

'Private Sub txtDel_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub

'Private Sub txtNumCertif_GotFocus()
'    fEnfoque txtNumCertif
'End Sub

'peac 20071128 convierte a mayuscula miestras escribe
'Private Sub txtNumCertif_Change()
''Dim i As Integer
''    txtNumCertif.Text = UCase(txtNumCertif.Text)
''    i = Len(txtNumCertif.Text)
''    txtNumCertif.SelStart = i
'End Sub

'
'Private Sub txtNumCertif_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub

'Private Sub txtSumaA_KeyPress(KeyAscii As Integer)
''     KeyAscii = NumerosDecimales(txtSumaA, KeyAscii)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub

'Private Sub txtSumaA_LostFocus()
'If txtSumaA.Text = "" Then
'    txtSumaA.Text = "0.00"
'Else
'    txtSumaA.Text = Format(txtSumaA.Text, "#0.00")
'End If
'End Sub


Private Sub txtTasa_GotFocus()
    fEnfoque txtTasa
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtTC_GotFocus()
    fEnfoque txtTC
End Sub

Private Sub txtTC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub
