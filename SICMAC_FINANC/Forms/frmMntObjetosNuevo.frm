VERSION 5.00
Begin VB.Form frmMntObjetosNuevo 
   Caption         =   "Objetos"
   ClientHeight    =   2835
   ClientLeft      =   4905
   ClientTop       =   1635
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3540
      TabIndex        =   4
      Top             =   2340
      Width           =   1215
   End
   Begin VB.TextBox txtObjetoCod 
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
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   0
      Top             =   375
      Width           =   1935
   End
   Begin VB.TextBox txtObjetoDesc 
      BackColor       =   &H00F0FFFF&
      Height          =   855
      Left            =   240
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1170
      Width           =   6435
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
      Left            =   240
      TabIndex        =   5
      Top             =   930
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
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1995
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   6675
   End
End
Attribute VB_Name = "frmMntObjetosNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lNuevo As Boolean
Dim sCod As String, sDesc As String
Dim nObjNiv As Integer

Dim lOk As Boolean
Dim lActivaTrans As Boolean

'ARLO20170208****
Dim objPista As COMManejador.Pista
Dim lsOpera, lsAccion As String
'************

Public Sub Inicia(plNuevo As Boolean, psCod As String, psDesc As String)
lNuevo = plNuevo
sCod = psCod
sDesc = psDesc
frmMdiMain.staMain.Panels(2).Text = "Mantenimiento de Objetos"
Me.Show 1
End Sub
Private Function GetNivelObj() As Integer
Dim sSql As String, rs As New ADODB.Recordset
Dim sText As String
Dim clsObj As New DObjeto

GetNivelObj = 1
sText = txtObjetoCod
Do While True
   sText = Left(sText, Len(sText) - 1)
   Set rs = clsObj.CargaObjeto(sText)
   If Not rs.EOF Then
      GetNivelObj = rs!nObjetoNiv + 1
      Exit Function
   End If
   If Len(sText) = 1 Then
      Exit Do
   End If
Loop
Set clsObj = Nothing
End Function

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Len(txtObjetoDesc) = 0 Then
   MsgBox " Descripción del Objeto esta vacio...! ", vbCritical, "Error de Actualización"
   txtObjetoDesc.SetFocus
   Exit Function
End If
If Len(txtObjetoCod) = 0 Then
   MsgBox " Código del Objeto esta vacio...! ", vbCritical, "Error de Actualización"
   txtObjetoCod.SetFocus
   Exit Function
Else
'   If lNuevo Then
      nObjNiv = GetNivelObj()
'   End If
End If
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
Dim clsObj As DObjeto
On Error GoTo errAcepta
If Not ValidaDatos() Then
   Exit Sub
End If
If MsgBox("¿ Seguro de Grabar ?", vbOKCancel + vbQuestion, "Confirmación") = vbOk Then
   sDesc = txtObjetoDesc.Text
   sCod = txtObjetoCod.Text
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Set clsObj = New DObjeto
   If lNuevo Then
      clsObj.InsertaObjeto sCod, sDesc, nObjNiv, gsMovNro
   Else
      clsObj.ActualizaObjeto sCod, sDesc, nObjNiv, gsMovNro
   End If
            'ARLO20170208
            If lNuevo Then
            lsOpera = "Agrego"
            lsAccion = "1"
            Else: lsOpera = "Modifico"
            lsAccion = "2"
            End If
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantObjetos
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, "Se " & lsOpera & " el Objeto |Cod : " & txtObjetoCod & " |Descripción : " & txtObjetoDesc
            Set objPista = Nothing
            '*******
   lOk = True
   Unload Me
End If
Exit Sub
errAcepta:
If Err.Number = -2147467259 Then
   MsgBox TextErr(Err.Description), vbCritical, "Error de Actualización"
Else
   MsgBox " Objeto ya existe. Imposible ADICIONAR ...! ", vbExclamation, "Error de Actualización"
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
lOk = False
End Sub

Private Sub txtObjetoCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Len(Trim(txtObjetoCod.Text)) <> 0 Then
      txtObjetoDesc.SetFocus
   Else
      MsgBox "Código no puede estar vacío...", vbInformation, "Atención...!!"
   End If
End If
End Sub

Private Sub txtObjetoDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   If Len(Trim(txtObjetoDesc.Text)) <> 0 Then
      cmdAceptar.SetFocus
   Else
      MsgBox "La descripción no puede estar vacía...", vbInformation, "Atención...!!"
   End If
End If
End Sub

Private Sub Form_Activate()
If lNuevo Then
   txtObjetoCod.SetFocus
Else
   txtObjetoCod.Enabled = False
   txtObjetoDesc.SetFocus
End If
End Sub

Private Sub Form_Load()
txtObjetoCod.Text = sCod
txtObjetoDesc.Text = sDesc
Me.Caption = "Objetos: Mantenimiento: " & IIf(lNuevo, "Nuevo", "Modificar")
CentraForm Me
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

Public Property Get cObjetoCod() As String
cObjetoCod = sCod
End Property
Public Property Let cObjetoCod(ByVal vNewValue As String)
sCod = cObjetoCod
End Property

Public Property Get cObjetoDesc() As String
cObjetoDesc = sDesc
End Property
Public Property Let cObjetoDesc(ByVal vNewValue As String)
sDesc = cObjetoDesc
End Property

