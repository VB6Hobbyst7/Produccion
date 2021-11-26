VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMntSaldosIni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldos Iniciales: "
   ClientHeight    =   6450
   ClientLeft      =   1440
   ClientTop       =   2625
   ClientWidth     =   9030
   Icon            =   "frmMntSaldosIni.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7530
      TabIndex        =   25
      Top             =   4635
      Width           =   1290
   End
   Begin VB.Frame Frame4 
      Height          =   1275
      Left            =   6165
      TabIndex        =   19
      Top             =   0
      Width           =   2820
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Puesta en Producción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   60
         TabIndex        =   21
         Top             =   720
         Width           =   2715
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Saldos Iniciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   330
         TabIndex        =   20
         Top             =   330
         Width           =   2190
      End
   End
   Begin VB.Frame fraBuscar 
      Enabled         =   0   'False
      Height          =   630
      Left            =   105
      TabIndex        =   17
      Top             =   630
      Width           =   6000
      Begin Sicmact.TxtBuscar TxtBuscar 
         Height          =   315
         Left            =   2100
         TabIndex        =   24
         Top             =   210
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
      Begin VB.Label Label8 
         Caption         =   "Buscar Cuenta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   285
         TabIndex        =   18
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3930
      Left            =   105
      TabIndex        =   13
      Top             =   1230
      Width           =   8880
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   7425
         TabIndex        =   5
         Top             =   1290
         Width           =   1290
      End
      Begin VB.CommandButton cmdTransferir 
         Caption         =   "&Transferir"
         Height          =   345
         Left            =   7425
         TabIndex        =   6
         Top             =   2970
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   345
         Left            =   7425
         TabIndex        =   4
         Top             =   875
         Width           =   1290
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   345
         Left            =   7425
         TabIndex        =   3
         Top             =   450
         Width           =   1290
      End
      Begin MSDataGridLib.DataGrid DBGCtas 
         Height          =   3495
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   19
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cCtaContCod"
            Caption         =   "Cuentas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "nCtaSaldoImporte"
            Caption         =   "Importe MN"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "nCtaSaldoImporteME"
            Caption         =   "Importe ME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "dCtaSaldofecha"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            ScrollBars      =   2
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               DividerStyle    =   6
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               DividerStyle    =   6
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               DividerStyle    =   6
               ColumnWidth     =   1620.284
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               DividerStyle    =   6
               ColumnWidth     =   1425.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraFecha 
      Height          =   630
      Left            =   105
      TabIndex        =   11
      Top             =   0
      Width           =   6000
      Begin MSMask.MaskEdBox TxtFecha 
         Height          =   300
         Left            =   2100
         TabIndex        =   0
         Top             =   225
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
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
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
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
         Left            =   4110
         TabIndex        =   1
         Top             =   195
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Saldos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   285
         TabIndex        =   12
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.Frame FraDatos 
      Height          =   1140
      Left            =   105
      TabIndex        =   14
      Top             =   5175
      Visible         =   0   'False
      Width           =   7275
      Begin VB.TextBox txtImpS2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   3990
         TabIndex        =   22
         Top             =   660
         Width           =   1515
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   5760
         TabIndex        =   10
         Top             =   690
         Width           =   1320
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   5760
         TabIndex        =   9
         Top             =   285
         Width           =   1320
      End
      Begin VB.TextBox TxtImpS 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3990
         TabIndex        =   8
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox TxtCta 
         Height          =   330
         Left            =   1005
         TabIndex        =   7
         Top             =   255
         Width           =   1710
      End
      Begin VB.Label Label12 
         Caption         =   "Importe S/ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2880
         TabIndex        =   23
         Top             =   675
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Importe :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2895
         TabIndex        =   16
         Top             =   285
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   150
         TabIndex        =   15
         Top             =   285
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMntSaldosIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nAccion As Integer
Dim dFechaCta As Date
Dim clsSdo As DCtaSaldo
Dim rsSdo  As ADODB.Recordset
Dim lTransferir As Boolean
Dim lSdoInicial As Boolean
'ARLO20170208****
Dim objPista As COMManejador.Pista
Dim lsPalabra, lsAccion As String
'************

Public Sub Inicio(plTransferir As Boolean)
lTransferir = plTransferir
Me.Show 1
End Sub

Private Sub CargaDatos(psFecha As String, Optional psCtaCod As String = "")
   Set rsSdo = clsSdo.CargaCtaSaldo(, Format(psFecha, "yyyymmdd"), adLockOptimistic)
   Set DBGCtas.DataSource = rsSdo
   If rsSdo.RecordCount > 0 Then
      fraBuscar.Enabled = True
   Else
      fraBuscar.Enabled = False
   End If
   If psCtaCod <> "" Then
      rsSdo.Find "cCtaContCod = '" & psCtaCod & "'"
   End If
End Sub

Private Sub ActivaBotones(lActiva As Boolean)
   If lActiva Then
      Height = 5670
      cmdSalir.Top = 4635
   Else
      Height = 6825
      cmdSalir.Top = 5865
   End If
   
   If Not lTransferir And lSdoInicial Then
      fraBuscar.Enabled = lActiva
      cmdNuevo.Enabled = lActiva
      cmdModificar.Enabled = lActiva
      cmdEliminar.Enabled = lActiva
   End If
   cmdtransferir.Enabled = lActiva
End Sub

Private Function SaldoIni(ByVal psCtaCod As String) As Double
Dim R As New ADODB.Recordset
   Set R = clsSdo.CargaCtaSaldo(psCtaCod, Format(txtFecha.Text, gsFormatoFecha))
   If Not R.BOF And Not R.EOF Then
      If Mid(psCtaCod, 3, 1) = "1" Or Mid(psCtaCod, 3, 1) = "6" Then
          SaldoIni = R!nCtaSaldoImporte
      Else
          SaldoIni = R!nCtaSaldoImporteme
      End If
   End If
End Function

Private Function ValidaDatos() As Boolean
ValidaDatos = False
On Error GoTo ValidaDatosErr
If fraDatos.Visible Then
   If txtCta.Text = "" Then
      MsgBox "Falta ingresar Cuenta Contable", vbInformation, "¡Aviso!"
      txtCta.SetFocus
      Exit Function
   End If
   Dim clsCta As New DCtaCont
   If Not clsCta.ExisteCuenta(txtCta.Text, True) Then
      txtCta.SetFocus
      Set clsCta = Nothing
      Exit Function
   End If
   Set clsCta = Nothing
Else
'   If Mid(txtOrigen.Text, 3, 1) <> Mid(txtDestino.Text, 3, 1) Then
'       MsgBox "Cuentas no deben ser de Diferente Moneda", vbInformation, "¡Aviso!"
'       txtOrigen.SetFocus
'       Exit Function
'   End If
'   If Trim(txtOrigen.Text) = Trim(txtDestino.Text) Then
'       MsgBox "No se PuedeTransferir a la misma Cuenta", vbInformation, "¡Aviso!"
'       txtOrigen.SetFocus
'       Exit Function
'   End If
'   If Trim(txtOrigen.Text) = "" Then
'       MsgBox "Falta Ingresar Cuenta Origen", vbInformation, "¡Aviso!"
'       txtOrigen.SetFocus
'       Exit Function
'   End If
'   If Trim(txtDestino.Text) = "" Then
'       MsgBox "Falta Ingresar Cuenta Destino", vbInformation, "¡Aviso!"
'       txtDestino.SetFocus
'       Exit Function
'   End If
End If
ValidaDatos = True
Exit Function
ValidaDatosErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Sub cmdAceptar_Click()
Dim sFecha As String
Dim nMontoD As Currency
Dim nMontoS As Currency
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro que desea guardar Datos ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
sFecha = Format(dFechaCta, gsFormatoFecha)
nMontoS = Format(txtImpS2, gsFormatoNumeroDato)
nMontoD = Format(TxtImpS, gsFormatoNumeroDato)
If nAccion = 1 Then
   Dim clsNSdo As New NCtasaldo
   If clsNSdo.ExisteCuentaSaldo(txtCta.Text, sFecha) Then
      Set clsNSdo = Nothing
      rsSdo.Find "cCtaContCod = '" & txtCta.Text & "'"
      If rsSdo!dCtaSaldofecha = dFechaCta Then
         MsgBox "Cuenta Contable ya tiene saldo a la Fecha", vbInformation, "¡Aviso!"
         txtCta.SetFocus
         Exit Sub
      Else
         If MsgBox("Cuenta ya tiene Saldo al " & rsSdo!dCtaSaldofecha & ". ¿Desea Ingresar Nuevo Saldo?", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            txtCta.SetFocus
            Exit Sub
         End If
      End If
   End If
   Set clsNSdo = Nothing
   If Mid(txtCta.Text, 3, 1) = "2" Then
      clsSdo.InsertaCtaSaldo txtCta.Text, sFecha, nMontoS, nMontoD
   Else
      clsSdo.InsertaCtaSaldo txtCta.Text, sFecha, nMontoD, 0
   End If
Else
   If Mid(txtCta.Text, 3, 1) = "2" Then
      clsSdo.ActualizaCtaSaldo txtCta.Text, sFecha, nMontoS, nMontoD
   Else
      clsSdo.ActualizaCtaSaldo txtCta.Text, sFecha, nMontoD, 0
   End If
End If
CargaDatos Format(txtFecha, gsFormatoFecha), txtCta.Text
            
'ARLO20170208
If (nAccion = 1) Then
lsPalabra = "Agrego"
lsAccion = "1"
Else: lsPalabra = "Modifico"
lsAccion = "2"
End If
Set objPista = New COMManejador.Pista
gsOpeCod = LogMantSaldoProducto
objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, "Se " & lsPalabra & " la Cuenta Puesta en Producción |Cod : " & txtCta.Text & "|Monto Soles " & nMontoS & " |Monto Dolares " & nMontoD
Set objPista = Nothing
'*******

fraDatos.Visible = False
fraFecha.Enabled = True
ActivaBotones True
DBGCtas.SetFocus
End Sub

Private Sub cmdAplicar_Click()
    If Len(ValidaFecha(txtFecha)) > 0 Then
        MsgBox "Fecha no válida...!", vbInformation, "¡Aviso!"
        txtFecha.SetFocus
        Exit Sub
    End If
    CargaDatos Format(txtFecha, gsFormatoFecha)
    TxtBuscar.rs = rsSdo
    ActivaBotones True
End Sub

Private Sub cmdCancelar_Click()
   ActivaBotones True
   fraDatos.Visible = False
End Sub

Private Sub cmdEliminar_Click()
If rsSdo Is Nothing Then
    Exit Sub
End If
If rsSdo.EOF Or rsSdo.RecordCount = 0 Then
    Exit Sub
End If
   If MsgBox(" ¿ Desea Eliminar El registro ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
      clsSdo.EliminaCtaSaldo rsSdo!cCtaContCod, Format(rsSdo!dCtaSaldofecha, gsFormatoFecha)
      clsSdo.EliminaCtaObjSaldo rsSdo!cCtaContCod, Format(rsSdo!dCtaSaldofecha, gsFormatoFecha)
      'ARLO20170208
        Set objPista = New COMManejador.Pista
        gsOpeCod = LogMantSaldoProducto
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se elimino la Cuenta Puesta en Producción |Cod : " & rsSdo!cCtaContCod
        Set objPista = Nothing
      '********
      rsSdo.Delete adAffectCurrent
   End If
   DBGCtas.SetFocus
End Sub

Private Sub cmdModificar_Click()
If rsSdo Is Nothing Then
    Exit Sub
End If
    If rsSdo.EOF Or rsSdo.RecordCount = 0 Then
        MsgBox "No se Seleccionó una Cuenta de la Lista", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    ActivaBotones False
    fraDatos.Visible = True
    fraFecha.Enabled = False
    
    txtCta.Enabled = False
    txtCta.Text = rsSdo!cCtaContCod
    If Mid(txtCta.Text, 3, 1) = "1" Or Mid(txtCta.Text, 3, 1) = "6" Then
        TxtImpS.Text = rsSdo!nCtaSaldoImporte
        txtImpS2.Text = "0.00"
        txtImpS2.Enabled = False
    Else
        TxtImpS.Text = IIf(IsNull(rsSdo!nCtaSaldoImporteme), 0, rsSdo!nCtaSaldoImporteme)
        txtImpS2.Text = IIf(IsNull(rsSdo!nCtaSaldoImporte), 0, rsSdo!nCtaSaldoImporte)
        txtImpS2.Enabled = True
    End If
    dFechaCta = CDate(Format(rsSdo!dCtaSaldofecha, gsFormatoFechaView))
    nAccion = 2
    TxtImpS.SetFocus
End Sub


Private Sub cmdNuevo_Click()
   ActivaBotones False
   fraDatos.Visible = True
   fraFecha.Enabled = False
   txtCta.Enabled = True
   txtCta.Text = ""
   TxtImpS.Text = "0.00"
   txtImpS2.Text = "0.00"
   dFechaCta = txtFecha.Text
   nAccion = 1
   txtCta.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTransferir_Click()
glAceptar = False
If rsSdo Is Nothing Then
    Exit Sub
End If
If rsSdo.EOF Then
    Exit Sub
End If
   frmMntSaldosIniTransf.Inicio rsSdo!cCtaContCod, Format(rsSdo!dCtaSaldofecha, gsFormatoFecha), rsSdo!nCtaSaldoImporte, rsSdo!nCtaSaldoImporteme, Format(txtFecha, gsFormatoFecha)
   If glAceptar Then
      CargaDatos Format(txtFecha, gsFormatoFecha), frmMntSaldosIniTransf.psDestino
   End If
   DBGCtas.SetFocus
End Sub

Private Sub Form_Load()
lSdoInicial = False
   Set clsSdo = New DCtaSaldo
   
   CentraForm Me
   If lTransferir Then
      Me.Caption = Me.Caption & "Transferencia entre Cuentas"
      Label10 = "Transferencia"
      Label11 = "entre Cuentas"
      cmdtransferir.Visible = True
      cmdNuevo.Visible = False
      cmdModificar.Visible = False
      cmdEliminar.Visible = False
   Else
      Me.Caption = Me.Caption & "Puesta en Producción: Mantenimiento"
      fraFecha.Visible = False
      fraBuscar.Top = fraFecha.Top
      Dim oSdo As New NCtasaldo
      txtFecha.Text = Replace(oSdo.GetFechaSdoInicial(), " ", "_")
      If txtFecha.Text = "__/__/____" Then
         txtFecha.Text = gdFecSis
         lSdoInicial = True
         ActivaBotones True
      Else
         lSdoInicial = oSdo.PermiteMntSdoInicial()
         If Not lSdoInicial Then
            MsgBox "Se detectaron Movimientos en el Sistema. Sólo puede Consultar Saldos Iniciales", vbInformation, "¡Aviso!"
         End If
         cmdAplicar_Click
         fraBuscar.Enabled = True
         If Not lSdoInicial Then
            cmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            cmdEliminar.Enabled = False
         End If
      End If
      Set oSdo = Nothing
   End If
   
   TxtBuscar.TipoBusqueda = BuscaDatoEnGrid
   TxtBuscar.EditFlex = False
   TxtBuscar.lbUltimaInstancia = False
      
   Height = 5670
   cmdSalir.Top = 4635
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsSdo
Set clsSdo = Nothing
End Sub

Private Sub TxtBuscar_EmiteDatos()
If TxtBuscar.psDescripcion <> "" Then
    If DBGCtas.Visible Then
        DBGCtas.SetFocus
    End If
End If
End Sub

Private Sub TxtBuscar_GotFocus()
    fEnfoque TxtBuscar
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtCta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If Mid(Trim(txtCta.Text), 3, 1) = "2" Then
            txtImpS2.Enabled = True
        Else
            txtImpS2.Enabled = False
        End If
        TxtImpS.SetFocus
    End If
End Sub


Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAplicar.SetFocus
    End If
End Sub

Private Sub TxtImpS_GotFocus()
fEnfoque TxtImpS
End Sub

Private Sub TxtImpS_KeyPress(KeyAscii As Integer)
Dim nTpoCam As Currency
Dim R As New ADODB.Recordset
Dim sSql As String
   KeyAscii = NumerosDecimales(TxtImpS, KeyAscii, 15, 2)
   If KeyAscii = 13 Then
      If Mid(Trim(txtCta.Text), 3, 1) = "2" Then
         nTpoCam = LeeTpoCambio(Me.txtFecha, TCFijoDia)
         txtImpS2.Text = Format(Round(CDbl(TxtImpS.Text) * nTpoCam, 2), gsFormatoNumeroView)
      End If
      If txtImpS2.Enabled Then
         txtImpS2.SetFocus
      Else
         cmdAceptar.SetFocus
      End If
   End If
End Sub

Private Sub TxtImpS_LostFocus()
    TxtImpS.Text = Format(CDbl(TxtImpS.Text), gsFormatoNumeroView)
End Sub

Private Sub txtImpS2_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImpS2, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub txtImpS2_LostFocus()
    If Len(Trim(txtImpS2.Text)) = 0 Then
        txtImpS2.Text = "0.00"
    End If
End Sub

