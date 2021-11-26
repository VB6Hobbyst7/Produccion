VERSION 5.00
Begin VB.UserControl ActXSegNumCarta 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   ScaleHeight     =   690
   ScaleWidth      =   6060
   Begin VB.Frame Frame1 
      Caption         =   "Nº de Carta"
      ForeColor       =   &H00FF0000&
      Height          =   640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCartaAnio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         MaxLength       =   4
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtCartaNum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   660
         X2              =   780
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Label lblNumCarta 
         Caption         =   "-GS-GA/CMACM"
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
         Left            =   1680
         TabIndex        =   4
         Top             =   285
         Width           =   1575
      End
   End
End
Attribute VB_Name = "ActXSegNumCarta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event KeyPress(KeyAscii As Integer)
Public Event Change()

Dim lsNumCarta As String
Dim lsPersCod As String
Dim lnIdSolicitud As Integer
Dim lsEstados  As String
Dim lnTpoSeguro As Long

Private Sub cmdBuscar_Click()
    Dim loPers As New COMDPersona.UCOMPersona
    Dim obj As New COMNCredito.NCOMGarantia
    Dim lsPersCod As String, lsPersNombre As String
    
    On Error GoTo ControlError
    
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    Set loPers = Nothing
    
    lsNumCarta = obj.ObtieneNumCartaSolicCobert(lsPersCod)
    txtCartaNum.Text = Trim(Left(lsNumCarta, 3))
    txtCartaAnio.Text = Trim(Right(lsNumCarta, 4))
    Exit Sub
    
ControlError:
        MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub txtCartaAnio_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case vbKeyReturn Or vbKeyTab
             If Len(Trim(txtCartaAnio.Text)) < 4 Then
                MsgBox "Número de Carta Incompleto", vbInformation, "Aviso"
                txtCartaAnio = ""
                If Len(Trim(txtCartaNum.Text)) < 3 Then
                    txtCartaNum.SetFocus
                Else
                    txtCartaAnio.SetFocus
                End If
             Else
                Call BuscarSolicitud(Trim(txtCartaNum.Text & txtCartaAnio.Text), Estados)
                RaiseEvent KeyPress(KeyAscii)
             End If
    End Select
End Sub

Private Sub txtCartaNum_Change()
    If Len(Trim(txtCartaNum.Text)) = 3 Then
        txtCartaAnio.SetFocus
    End If
End Sub
Private Sub txtCartaNum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
             If Len(Trim(txtCartaNum.Text)) = 3 Then
                txtCartaAnio.SetFocus
            End If
    End Select
End Sub

Private Sub UserControl_InitProperties()
    txtCartaNum.Text = ""
    txtCartaAnio.Text = ""
End Sub

Public Sub Limpiar()
    txtCartaNum.Text = ""
    txtCartaAnio.Text = ""
    txtCartaNum.Enabled = True
    txtCartaAnio.Enabled = True
    'txtCartaNum.SetFocus
    cmdBuscar.Enabled = True
End Sub

Public Sub BuscarSolicitud(ByVal psNumCarta As String, ByVal psEstados As String)
    Dim obj As New COMNCredito.NCOMGarantia
    Dim lsNumCarta As String
    lsNumCarta = obj.BuscaSolicitudCobert(psNumCarta, psEstados, lnTpoSeguro)
    If lsNumCarta = "" Then
        MsgBox "No se encontro la solicitud", vbInformation, "Aviso"
        Me.Limpiar
        txtCartaNum.SetFocus
    Else
        lnIdSolicitud = Trim(Right(lsNumCarta, 1))
        txtCartaNum.Enabled = False
        txtCartaAnio.Enabled = False
        cmdBuscar.Enabled = False
    End If
End Sub
'''PROPIEDADES
Public Property Get NumCarta() As String
    NumCarta = txtCartaNum.Text & "" & txtCartaAnio.Text
End Property
Public Property Let NumCarta(ByVal vNewValue As String)
    txtCartaNum.Text = Trim(Left(vNewValue, 3))
    txtCartaAnio.Text = Trim(Right(lsNumCarta, 4))
    PropertyChanged "NumCarta"
End Property
Public Property Get IdSolicitud() As Integer
    IdSolicitud = lnIdSolicitud
End Property
Public Property Let IdSolicitud(ByVal vNewValue As Integer)
    lnIdSolicitud = vNewValue
    PropertyChanged "IdSolicitud"
End Property
Public Property Get Estados() As String
    Estados = lsEstados
End Property
Public Property Let Estados(ByVal vNewValue As String)
    lsEstados = vNewValue
    PropertyChanged "Estados"
End Property
Public Property Get TpoSeguro() As Long
    TpoSeguro = lnTpoSeguro
End Property
Public Property Let TpoSeguro(ByVal vNewValue As Long)
    lnTpoSeguro = vNewValue
    PropertyChanged "TpoSeguro"
End Property

