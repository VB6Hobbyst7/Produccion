VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogOCDesbloqueo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desbloqueo de Ordenes Compra"
   ClientHeight    =   2670
   ClientLeft      =   2535
   ClientTop       =   2970
   ClientWidth     =   6480
   Icon            =   "frmLogOCDesbloqueo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Quitar el Bloqueo de una Orden de Compra / Servicio "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6195
      Begin VB.CommandButton CmdDesbloquear 
         Caption         =   "Desbloquear"
         Height          =   375
         Left            =   4500
         TabIndex        =   3
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox TxtOrden 
         Height          =   285
         Left            =   1980
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese Nro Orden:"
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
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Asignar Imagen de firmas en las órdenes de Compra / Servicio"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   6195
      Begin VB.CommandButton cmdImagen 
         Caption         =   "Archivo de Imagen"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1920
         Width           =   1635
      End
      Begin MSComDlg.CommonDialog dlgImg 
         Left            =   180
         Top             =   1860
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtCodUsu 
         Height          =   315
         Left            =   840
         MaxLength       =   4
         TabIndex        =   6
         Top             =   540
         Width           =   675
      End
      Begin VB.PictureBox picFirma 
         BackColor       =   &H80000005&
         Height          =   1275
         Left            =   2160
         ScaleHeight     =   1215
         ScaleWidth      =   3675
         TabIndex        =   5
         Top             =   540
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   600
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmLogOCDesbloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oConn As DConecta
Dim rsOCVerifica As New Recordset
Dim sRutaTmp As String
Dim nTipoForm As Integer

Public Sub Inicio(ByVal pnTipoForm As Integer)
nTipoForm = pnTipoForm
Me.Show 1
End Sub

Private Sub Form_Load()
Select Case nTipoForm
    Case 1
         Me.Height = 1700
         Frame1.Visible = True
    Case 2
         Me.Height = 3080
         Frame2.Visible = True
End Select
End Sub

Private Sub CmdDesbloquear_Click()
Dim valor As Integer

Set oConn = New DConecta

If oConn.AbreConexion() Then

Set rsOCVerifica = oConn.Ejecutar("if not exists(select cDocNro from MovCotizaccontrol " & _
                " where cDocNro='1212') " & _
                " select valor=1 " & _
                " Else " & _
                " begin " & _
                " if exists(select cDocNro from MovCotizaccontrol " & _
                " where cDocNro='1212') " & _
                " begin " & _
                " select valor=0 " & _
                " delete from MovCotizaccontrol where cDocNro='1212' " & _
                " End " & _
                " End ")
 
   
   If rsOCVerifica("valor") = 0 Then
        MsgBox " El Documento fue Desbloqueado Correctamente"
   Else
        MsgBox " El Documento no esta Bloqueado"
   End If
End If
End Sub



Private Sub txtCodUsu_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cmdImagen.SetFocus
End If
End Sub

Private Sub TxtOrden_Change()
If Len(TxtOrden) > 13 Then
TxtOrden.Text = Left(TxtOrden, 13)
TextoSeleccionado
End If
End Sub

Private Sub TxtOrden_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnterosOrdenCompra(KeyAscii)
End Sub

Public Sub TextoSeleccionado()
Dim i As Integer
Dim TextBox As Object

Set TextBox = Screen.ActiveControl
    If TypeName(TextBox) = "TextBox" Then
        i = Len(TextBox.Text)
        TextBox.SelStart = 0
        TextBox.SelLength = i
    End If
End Sub

Function NumerosEnterosOrdenCompra(intTecla As Integer, Optional pbNegativos As Boolean = False) As Integer
Dim cValidar As String
    If pbNegativos = False Then
        cValidar = "0123456789-"
    Else
        cValidar = "0123456789-"
    End If
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            'Beep
        End If
    End If
    NumerosEnterosOrdenCompra = intTecla
End Function


'-------------------------------------------------------------------------------


Private Sub cmdImagen_Click()
On Error GoTo SalImg

dlgImg.ShowOpen
sRutaTmp = dlgImg.FileName
Set picFirma.Picture = LoadPicture(sRutaTmp)
Exit Sub

SalImg:

End Sub

Private Sub cmdAsignar_Click()
Dim sSQL As String, RF As New ADODB.Recordset
Dim cCod As String, nItem As Integer
Dim RStream As New ADODB.Stream
Dim DBImagen As ADODB.Connection
Dim oConn As New DConecta

On Error GoTo ErrorGrabarFirmaenBD

If Len(Trim(txtCodUsu.Text)) = 0 Then
   MsgBox "Código de cargo no válido..." + Space(10), vbInformation
   Exit Sub
End If

If MsgBox("¿ Está seguro de grabar la imagen mostrada ?" + Space(10), vbQuestion + vbYesNo, "Confirme grabación") = vbYes Then

   Set DBImagen = New ADODB.Connection
   If Not oConn.AbreConexion Then
      MsgBox "No se puede establecer conexión..." + Space(10), vbInformation
      Exit Sub
   End If
   
   'sSQL = "Select * from DBCmactAux..MovDocFirmas Where cRHCargoCod = '" & txtcodusu.Text & "' "
   sSQL = "Select * from DBCmactAux..MovDocFirmas Where cCodUser = '" & txtCodUsu.Text & "' "
   
   If RF.State = adStateOpen Then RF.Close
   RF.CursorLocation = adUseClient
   DBImagen.Open oConn.CadenaConexion
   
   DBImagen.CommandTimeout = 20000
   RF.Open sSQL, DBImagen, adOpenStatic, adLockOptimistic, adCmdText

   If RF.BOF And RF.EOF Then
      RF.AddNew
      'RF.Fields("cRHCargoCod").value = txtcodusu.Text
      RF.Fields("cCodUser").value = txtCodUsu.Text
   End If
   
   RStream.Type = adTypeBinary
   RStream.Open
   RStream.LoadFromFile sRutaTmp
   If RStream.State = 0 Then
      MsgBox "Error al cargar imagen..." + Space(10), vbInformation
      RF.Close
      RStream.Close
      Set RStream = Nothing
   End If
   RF.Fields("iFirma").value = RStream.Read
   RStream.Close
   Set RStream = Nothing
   RF.Update
   RF.Close
   MsgBox "La imagen se grabó con éxito" + Space(10), vbInformation
End If
Exit Sub

ErrorGrabarFirmaenBD:
    MsgBox "Error " + Err.Description
    GrabarImagen = False
End Sub


