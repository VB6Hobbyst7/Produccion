VERSION 5.00
Begin VB.Form frmCredPagoMotivoCancAnt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Motivo de la Cancelación Anticipada"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "frmCredPagoMotivoCancAnt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtOtros 
         Height          =   300
         Left            =   720
         TabIndex        =   6
         Top             =   670
         Visible         =   0   'False
         Width           =   5115
      End
      Begin VB.ComboBox cboMotivo 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label lblOtros 
         Caption         =   "Otros"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   1
      Top             =   1200
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3720
      TabIndex        =   0
      Top             =   1200
      Width           =   1170
   End
End
Attribute VB_Name = "frmCredPagoMotivoCancAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredPagoMotivoCancAnt
'** Descripción : Formulario para elegir el motivo del porque se esta realizando
'**               la cancelacion anticipada creado segun RFC059-2012
'** Creación : JUEZ, 20120921 09:00:00 AM
'**********************************************************************************************

Option Explicit

Private lnValida As Boolean
Private lnMotivoCanc As Integer
Private lsMotivoCancOtros As String
Property Let MotivoCancOtros(pMotivoCancOtros As String)
   lsMotivoCancOtros = pMotivoCancOtros
End Property
Property Get MotivoCancOtros() As String
    MotivoCancOtros = lsMotivoCancOtros
End Property
Property Let MotivoCanc(pMotivo As Integer)
   lnMotivoCanc = pMotivo
End Property
Property Get MotivoCanc() As Integer
    MotivoCanc = lnMotivoCanc
End Property
Property Let RegistraMotivo(pRegMotivo As String)
   lnValida = pRegMotivo
End Property
Property Get RegistraMotivo() As String
    RegistraMotivo = lnValida
End Property

Public Sub Inicia()
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Call CargaCombo(cboMotivo, 9080)
    Call CambiaTamañoCombo(cboMotivo, 390)
    lnValida = False
    lnMotivoCanc = 0
    lsMotivoCancOtros = ""
    Me.Show 1
End Sub

Private Sub cboMotivo_Click()
    If Trim(Right(cboMotivo.Text, 3)) = "14" Then
        lblOtros.Visible = True
        txtOtros.Visible = True
    Else
        lblOtros.Visible = False
        txtOtros.Visible = False
    End If
End Sub

Private Sub cmdAceptar_Click()
    If Trim(Right(cboMotivo.Text, 3)) <> "" Then
        If Trim(Right(cboMotivo.Text, 3)) = "14" And Trim(txtOtros.Text) = "" Then
            MsgBox "Detalle el motivo de la cancelacion", vbInformation, "Aviso"
            txtOtros.SetFocus
        Else
            lnValida = True
            lnMotivoCanc = CInt(Trim(Right(cboMotivo.Text, 3)))
            lsMotivoCancOtros = IIf(txtOtros.Visible = False, "", Trim(txtOtros.Text))
            Unload Me
        End If
    Else
        MsgBox "Seleccione un motivo", vbInformation, "Aviso"
    End If
End Sub

Public Sub dInsertaMotivoCancAnticipada(ByVal psCtaCod As String, ByVal pnMotivoCanc As Integer, ByVal psMotivoOtros As String)
    Dim oDCred As COMDCredito.DCOMCredActBD
    Set oDCred = New COMDCredito.DCOMCredActBD
    oDCred.dInsertaMotivoCancAnticipada psCtaCod, pnMotivoCanc, psMotivoOtros
    Set oDCred = Nothing
End Sub

Private Sub cmdCancelar_Click()
    lnValida = False
    Unload Me
End Sub

Public Sub CargaCombo(ByVal CtrlCombo As ComboBox, ByVal psConst As String)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(psConst)
    Set clsGen = Nothing
    
    CtrlCombo.Clear
    While Not rsConst.EOF
        CtrlCombo.AddItem rsConst.fields(0) & Space(200) & rsConst.fields(1)
        rsConst.MoveNext
    Wend
End Sub

Private Sub txtOtros_LostFocus()
    txtOtros.Text = UCase(Trim(txtOtros.Text))
End Sub
