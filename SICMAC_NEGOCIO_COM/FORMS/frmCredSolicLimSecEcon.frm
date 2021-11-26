VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCredSolicLimSecEcon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud fuera de Límites de Sectores"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   Icon            =   "frmCredSolicLimSecEcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstSolicitud 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   531
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Agencia"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Crédito"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Titular"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Producto"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tipo Crédito"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Moneda"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Monto"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Monto MN"
         Object.Width           =   2346
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Total MN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Sector MN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Limite %"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "% con este crèdito"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Sector"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "Autorizar"
      Height          =   360
      Left            =   7080
      TabIndex        =   2
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox txtGlosa 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   3600
      Width           =   6135
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   9480
      TabIndex        =   4
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
      Height          =   360
      Left            =   8280
      TabIndex        =   3
      Top             =   3600
      Width           =   1050
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Glosa :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   495
   End
End
Attribute VB_Name = "frmCredSolicLimSecEcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredSolicLimSecEcon
'** Descripción : Formulario para autorizar/rechazar solicitudes de límites de créditos por
'**               Sector económico creado segun TI-ERS029-2013
'** Creación : JUEZ, 20140603 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oDPersGen As COMDPersona.DCOMPersGeneral
Dim rs As ADODB.Recordset

Private Sub cmdAutorizar_Click()
AutorizarRechazarSolicitud (1)
End Sub

Private Sub cmdRechazar_Click()
AutorizarRechazarSolicitud (2)
End Sub

Private Sub AutorizarRechazarSolicitud(ByVal pnEstado As Integer)
If Trim(txtGlosa.Text) = "" Then
    MsgBox "Debe ingresar la glosa", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Sub
End If

If pnEstado = 1 Then
    If MsgBox("Esta opción autorizará sugerir un crédito que superará los límites del sector, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
Else
    If MsgBox("Esta opción rechazará la solicitud de autorización del crédito, ésto po podrá ser sugerido, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
End If
Set oDPersGen = New COMDPersona.DCOMPersGeneral
Call oDPersGen.ActualizarSolicitudAutorizacionRiesgos(LstSolicitud.SelectedItem.SubItems(2), Trim(Me.txtGlosa.Text), pnEstado)
Set oDPersGen = Nothing
MsgBox "La solicitud fue " & IIf(pnEstado = 1, "autorizada", "rechazada"), vbInformation, "Aviso"
txtGlosa.Text = ""
CargarSolicitudes
End Sub


Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
CargarSolicitudes
End Sub

Private Sub CargarSolicitudes()
Dim L As ListItem
Set oDPersGen = New COMDPersona.DCOMPersGeneral
Set rs = oDPersGen.RecuperaSolicitudAutorizacionRiesgos
Set oDPersGen = Nothing

    LstSolicitud.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
        LstSolicitud.Enabled = True
        Do While Not rs.EOF
            Set L = LstSolicitud.ListItems.Add(, , rs.Bookmark)
            L.SubItems(1) = rs!cAgeDescripcion
            L.SubItems(2) = rs!cCtaCod
            L.SubItems(3) = rs!cPersNombre
            L.SubItems(4) = rs!cTpoProdDesc
            L.SubItems(5) = rs!cTpoCredDesc
            L.SubItems(6) = rs!cMoneda
            L.SubItems(7) = Format(rs!nMontoSol, "#,##0.00")
            L.SubItems(8) = Format(rs!nMontoMN, "#,##0.00")
            L.SubItems(9) = Format(rs!nTotalMN, "#,##0.00")
            L.SubItems(10) = Format(rs!nSectorMN, "#,##0.00")
            L.SubItems(11) = Format(rs!nLimite, "#,##0.00")
            L.SubItems(12) = Format(rs!nPorcConEsteCred, "#,##0.00")
            L.SubItems(13) = rs!cSectorDesc
            rs.MoveNext
        Loop
    End If
End Sub
