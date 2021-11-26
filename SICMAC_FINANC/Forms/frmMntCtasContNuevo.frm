VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMntCtasContNuevo 
   Caption         =   "Cuentas Contables"
   ClientHeight    =   3285
   ClientLeft      =   2700
   ClientTop       =   2580
   ClientWidth     =   7065
   Icon            =   "frmMntCtasContNuevo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgencia 
      Caption         =   "&Generar Agencias >>>"
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
      TabIndex        =   16
      Top             =   2850
      Width           =   2355
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   2850
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2850
      Width           =   1215
   End
   Begin VB.Frame fraMoneda 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   300
      TabIndex        =   13
      Top             =   1380
      Width           =   6435
      Begin VB.CheckBox chkMoneda 
         Caption         =   "6 - Partidas no Monetarias Ajustadas"
         Height          =   195
         Index           =   5
         Left            =   3180
         TabIndex        =   10
         Top             =   840
         Width           =   3015
      End
      Begin VB.CheckBox chkMoneda 
         Caption         =   "4 - Operaciones de Capital Reajustables"
         Height          =   195
         Index           =   4
         Left            =   3180
         TabIndex        =   9
         Top             =   570
         Width           =   3165
      End
      Begin VB.CheckBox chkMoneda 
         Caption         =   "3 - De actualización Constante"
         Height          =   195
         Index           =   3
         Left            =   3180
         TabIndex        =   8
         Top             =   270
         Width           =   2835
      End
      Begin VB.CheckBox chkMoneda 
         Caption         =   "2 - Moneda Extranjera"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chkMoneda 
         Caption         =   " 0 - Integrador"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   5
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox chkMoneda 
         Caption         =   "1 - Moneda Nacional"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   570
         Width           =   2085
      End
   End
   Begin VB.TextBox txtDigRest 
      BackColor       =   &H00F0FFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2100
      MaxLength       =   23
      TabIndex        =   2
      Top             =   375
      Width           =   2115
   End
   Begin VB.TextBox txtDigInt 
      Alignment       =   2  'Center
      BackColor       =   &H00F0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   1
      ToolTipText     =   "Indica la moneda para la cuenta contable"
      Top             =   375
      Width           =   255
   End
   Begin VB.TextBox txtDigDos 
      BackColor       =   &H00F0FFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   1500
      MaxLength       =   2
      TabIndex        =   0
      Top             =   375
      Width           =   375
   End
   Begin VB.TextBox txtCtaContDescrip 
      BackColor       =   &H00F0FFFF&
      Height          =   375
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   4
      Top             =   900
      Width           =   5235
   End
   Begin MSComctlLib.ListView lvAgencia 
      Height          =   2265
      Left            =   120
      TabIndex        =   17
      Top             =   3330
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   3995
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Agencia"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción"
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
      Left            =   300
      TabIndex        =   11
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
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
      Left            =   300
      TabIndex        =   3
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   2595
      Left            =   120
      TabIndex        =   12
      Top             =   180
      Width           =   6795
   End
End
Attribute VB_Name = "frmMntCtasContNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lNuevo As Boolean
Dim sTabla As String
Dim sCod As String, sDesc As String
Dim lOk  As Boolean
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim lConsolidado As Boolean

Dim clsCtaCont As DCtaCont

Public Sub Inicia(plNuevo As Boolean, psCod As String, psDesc As String, Optional psTabla As String = "CtaCont", Optional plConsolidado As Boolean = True)
lNuevo = plNuevo
If sTabla = "CtaCont" Then
   sTabla = gsCentralCom & psTabla
Else
   sTabla = psTabla
End If
sCod = psCod
sDesc = psDesc
lConsolidado = plConsolidado
frmMdiMain.staMain.Panels(2).Text = "Mantenimiento de Cuentas Contables"
Me.Show 1
End Sub

Private Sub chkMoneda_Click(Index As Integer)
If txtDigInt = "" Then
   chkMoneda(Index).value = 0
   Exit Sub
End If
Select Case Index
   Case 1, 2, 5
      If chkMoneda(1).value = 0 And chkMoneda(2).value = 0 And chkMoneda(5).value = 0 Then
         chkMoneda(0).value = 0
      Else
         chkMoneda(0).value = 1
      End If
End Select
End Sub

Private Sub cmdAceptar_Click()
Dim SQLctas As String
Dim sCta      As String
Dim lAgencias As Boolean
Dim m As Integer
Dim n As Integer
On Error GoTo errAcepta
If Len(Trim(txtCtaContDescrip)) = 0 Then
   MsgBox " Descripción de Cuenta Contable vacia...! ", vbCritical, "Error de datos"
   txtCtaContDescrip.SetFocus
   Exit Sub
End If
If txtDigDos = "" And txtDigInt = "" And txtDigRest = "" Then
   MsgBox " No se definió Cuenta Contable...! ", vbInformation, "Aviso"
   txtDigDos.SetFocus
   Exit Sub
End If

sCta = txtDigDos & "0" & txtDigRest
If sTabla = gsCentralCom & "CtaCont" Then
   glAceptar = True
   Set rs = clsCtaCont.CargaCtaCont("substring('" & sCta & "',1,LEN(cCtaContCod)) = cCtaContCod", "CtaContBase")
   If rs.EOF Then
      MsgBox "Cuenta Contable no permitida. Verificar Plan de Cuentas de SBS", vbInformation, "Aviso"
      glAceptar = False
      Exit Sub
   End If
   rs.MoveLast
   sSql = "cCtaContCod LIKE '" & rs!cCtaContCod & "_%' "
   Set rs = clsCtaCont.CargaCtaCont(sSql, "CtaContBase")
   If rs.RecordCount > 0 Then
      If MsgBox("Cuenta Contable no permitida segun Plan de Cuentas de SBS. ¿ Desea Continuar ?", vbQuestion + vbYesNo, "!Confirmación¡") = vbNo Then
         glAceptar = False
      End If
   End If
   RSClose rs
   If Not glAceptar Then
      If txtDigDos.Enabled Then
         txtDigDos.SetFocus
      Else
         cmdCancelar.SetFocus
      End If
      Exit Sub
   End If
End If

If lvAgencia.ListItems.Count > 0 Then
   Set rs = clsCtaCont.CargaCtaCont(" cCtaContCod LIKE '" & sCta & "__' ")
   If rs.RecordCount > 0 Then
      lAgencias = True
   End If
   RSClose rs
End If

If MsgBox(" ¿ Seguro de grabar Datos ? ", vbOKCancel + vbQuestion, "Confirma grabación") = vbOk Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   If txtDigInt = "" Then
      sCod = txtDigDos
      If lNuevo Then
         clsCtaCont.InsertaCtaCont txtDigDos.Text, txtCtaContDescrip, gsMovNro, sTabla
      Else
         clsCtaCont.InsertaCtaCont sCod, txtCtaContDescrip, gsMovNro, sTabla, "6"
      End If
   Else
      If Not lConsolidado Then
         sCta = txtDigDos & txtDigInt & txtDigRest
         clsCtaCont.InsertaCtaCont sCod, txtCtaContDescrip, gsMovNro, sTabla, "6"
      Else
         sCod = ""
         For n = 0 To chkMoneda.Count - 1
            sCta = txtDigDos & IIf(n = 5, "6", n) & txtDigRest
            If chkMoneda(n).value = 1 Then
               sCod = sCod & IIf(n = 5, "6", n)
            Else
               If lvAgencia.ListItems.Count > 0 Then
                  sCta = sCta & "%"
               End If
               clsCtaCont.EliminaCtaCont sCta, sTabla
            End If
         Next
         If sCod <> "" Then
            clsCtaCont.InsertaCtaCont txtDigDos & txtDigInt & txtDigRest, txtCtaContDescrip, gsMovNro, sTabla, sCod
         End If
      End If
      'sCod = txtDigDos & "0" & txtDigRest
   End If
   If lvAgencia.ListItems.Count > 0 Then
      If lAgencias Then
         n = MsgBox(" Cuenta ya posee divisionarias de Agencias. ¿ Desea Continuar ? ", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Confirmación")
         If n = vbCancel Then
'            lTransOp = False
            Exit Sub
         End If
      Else
         n = vbYes
      End If
      If n = vbYes Then
         For n = 1 To lvAgencia.ListItems.Count
            sCta = txtDigDos & "_" & txtDigRest & Right(lvAgencia.ListItems(n).Text, 2)
            If lvAgencia.ListItems(n).Checked Then
                clsCtaCont.InsertaCtaCont sCta, lvAgencia.ListItems(n).SubItems(1), gsMovNro, sTabla, sCod
'                     GrabaCuenta sCta, lvAgencia.ListItems(N).SubItems(1), ExisteCuenta(sCta)
            Else
               clsCtaCont.EliminaCtaCont sCta, sTabla
               'GrabaCuenta sCta, lvAgencia.ListItems(N).SubItems(1), 2
            End If
         Next
      End If
   End If
   sCod = txtDigDos & txtDigInt & txtDigRest
   
   lOk = True
   Unload Me
End If
Exit Sub
errAcepta:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"

End Sub

Private Sub cmdAgencia_Click()
Dim lvItem As ListItem
Dim oAge As New DActualizaDatosArea
If cmdAgencia.Caption = "&Generar Agencias >>>" Then
    Set rs = oAge.GetAgencias(, False)
   If rs.EOF Then
      RSClose rs
      MsgBox "No se definieron Agencias en el Sistema...Consultar con Sistemas", vbInformation, "Aviso"
      Exit Sub
   End If
   cmdAgencia.Caption = "No &Generar Agencias <<<"
   Me.Height = Me.Height + 2500
   Do While Not rs.EOF
      Set lvItem = lvAgencia.ListItems.Add(, , rs!Codigo)
      lvItem.SubItems(1) = rs!Descripcion
      lvItem.Checked = True
      rs.MoveNext
   Loop
   RSClose rs
   lvAgencia.SetFocus
Else
   cmdAgencia.Caption = "&Generar Agencias >>>"
   Me.Height = Me.Height - 2500
   lvAgencia.ListItems.Clear
End If

End Sub

Private Sub CmdCancelar_Click()
Unload Me
lOk = False
End Sub

Private Sub txtDigDos_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If Len(Trim(txtDigDos.Text)) = 1 Then
   txtDigInt.SetFocus
   Exit Sub
End If
If KeyAscii = 13 Then
   If Len(Trim(txtDigDos.Text)) <> 0 Then
      txtDigInt.SetFocus
   Else
      MsgBox "El campo no puede estar vacío...", vbInformation, "Atención...!!"
   End If
End If
End Sub

Private Sub txtDigInt_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If Len(txtDigInt) = 0 Then
      chkMoneda(0).Enabled = False
      chkMoneda(1).Enabled = False
      chkMoneda(2).Enabled = False
      chkMoneda(3).Enabled = False
      chkMoneda(4).Enabled = False
      chkMoneda(5).Enabled = False
      txtCtaContDescrip.SetFocus
      Exit Sub
   End If
End If
If lConsolidado Then
   If InStr("0", Chr$(KeyAscii)) = 0 Then
      MsgBox "Valor debe ser Dígito Consolidado", vbInformation, "Atención...!!"
      KeyAscii = 0
   Else
      txtDigRest.SetFocus
   End If
ElseIf InStr("6", Chr$(KeyAscii)) = 0 Then
      MsgBox "Valor debe ser Dígito 6 de Ajuste", vbInformation, "Atención...!!"
      KeyAscii = 0
Else
   txtDigRest.SetFocus
End If
End Sub

Private Sub txtDigRest_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtCtaContDescrip.SetFocus
End If
End Sub
Private Sub txtCtaContDescrip_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   If Len(Trim(txtCtaContDescrip.Text)) <> 0 Then
      If Len(txtDigInt) = 0 Then
         cmdAceptar.SetFocus
      Else
         If fraMoneda.Enabled Then
            chkMoneda(1).SetFocus
         Else
            cmdAceptar.SetFocus
         End If
      End If
   Else
      MsgBox "El campo no puede estar vacío...", vbInformation, "Atención...!!"
   End If
End If
End Sub

Private Sub Form_Activate()
If Not lConsolidado Then
   fraMoneda.Enabled = False
End If
If lNuevo Then
   txtDigDos.SetFocus
Else
   txtDigDos.Enabled = False
   txtDigInt.Enabled = False
   txtDigRest.Enabled = False
   txtCtaContDescrip.SetFocus
End If

End Sub

Private Sub Form_Load()
CentraForm Me
Me.Top = Me.Top - 1000
txtDigDos.Text = Mid(sCod, 1, 2)
txtDigInt.Text = Mid(sCod, 3, 1)
txtDigRest.Text = Mid(sCod, 4, 19)
txtCtaContDescrip.Text = sDesc
Me.Caption = "Cuentas Contables: Mantenimiento: " & IIf(lNuevo, "Nuevo", "Modificar")
Set clsCtaCont = New DCtaCont
If Not lNuevo Then
   Set rs = clsCtaCont.CargaCtaCont("cCtaContCod LIKE '" & txtDigDos.Text & "_" & txtDigRest.Text & "'", sTabla)
   Do While Not rs.EOF
      chkMoneda(IIf(Mid(rs!cCtaContCod, 3, 1) = "6", 5, Val(Mid(rs!cCtaContCod, 3, 1)))).value = 1
      rs.MoveNext
   Loop
   RSClose rs
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMdiMain.staMain.Panels(2).Text = ""
End Sub

Public Property Get OK() As Integer
    OK = lOk
End Property
Public Property Let OK(ByVal vNewValue As Integer)
lOk = OK
End Property

Public Property Get cCtaContCod() As String
cCtaContCod = sCod
End Property

Public Property Let cCtaContCod(ByVal vNewValue As String)
sCod = cCtaContCod
End Property
