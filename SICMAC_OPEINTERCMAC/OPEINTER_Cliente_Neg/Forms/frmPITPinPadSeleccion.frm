VERSION 5.00
Begin VB.Form frmPITPinPadSeleccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de PINPAD"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   ControlBox      =   0   'False
   Icon            =   "frmPITPinPadSeleccion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   45
      TabIndex        =   1
      Top             =   1185
      Width           =   4695
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   345
         Left            =   3390
         TabIndex        =   5
         Top             =   195
         Width           =   1200
      End
      Begin VB.CommandButton CmdAcpetar 
         Caption         =   "Aceptar"
         Height          =   345
         Left            =   75
         TabIndex        =   4
         Top             =   195
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   4680
      Begin VB.ComboBox CboPuerto 
         Height          =   315
         ItemData        =   "frmPITPinPadSeleccion.frx":030A
         Left            =   1800
         List            =   "frmPITPinPadSeleccion.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   690
         Width           =   2745
      End
      Begin VB.ComboBox cboTipoPinPad 
         Height          =   315
         ItemData        =   "frmPITPinPadSeleccion.frx":030E
         Left            =   1800
         List            =   "frmPITPinPadSeleccion.frx":031B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   285
         Width           =   2745
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Puerto :"
         Height          =   270
         Left            =   75
         TabIndex        =   6
         Top             =   690
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Pinpad a Usar :"
         Height          =   270
         Left            =   60
         TabIndex        =   2
         Top             =   315
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmPITPinPadSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Private Function GetMaquinaUsuario() As String  'Para obtener la Maquina del Usuario
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    GetMaquinaUsuario = Trim(Left$(buffMaq, lSizeMaq))
End Function

Private Sub CmdAcpetar_Click()
Dim lsSql As String
Dim loConec As New DConecta

    gnTipoPinPad = CInt(Right(Me.cboTipoPinPad.Text, 3))
    gnPinPadPuerto = CInt(Me.CboPuerto.Text)

    lsSql = " exec  PIT_stp_ins_RegistraPinPadDePC '" & GetMaquinaUsuario & "'," & gnTipoPinPad & "," & gnPinPadPuerto
     
    loConec.AbreConexion
    loConec.ConexionActiva.Execute lsSql
    loConec.CierraConexion
    
    Set loConec = Nothing
        
    MsgBox "Datos Registrados con Exito", vbInformation, "Aviso"

    Unload Me
   
End Sub


Public Function ObtieneCodigos(ByVal psCodigo As Integer) As ADODB.Recordset
Dim loConec As New DConecta
Dim lsSql As String
On Error GoTo ErrorObtieneCodigos
    loConec.AbreConexion
    
    lsSql = "exec DBTarjeta..stp_sel_DevuelveCodigos " & psCodigo

    Set ObtieneCodigos = loConec.CargaRecordSet(lsSql)
    
    loConec.CierraConexion
    Set loConec = Nothing
    
Exit Function
ErrorObtieneCodigos:
    err.Raise err.Number, "DCOMPersona:ErrorObtieneCodigos", err.Description
End Function



Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim rsPinPad As ADODB.Recordset

    CboPuerto.AddItem "1"
    CboPuerto.AddItem "2"
    CboPuerto.AddItem "3"
    CboPuerto.AddItem "4"
    CboPuerto.AddItem "5"
    CboPuerto.AddItem "6"
    CboPuerto.AddItem "7"
    CboPuerto.AddItem "8"
    CboPuerto.AddItem "9"
    CboPuerto.AddItem "10"
    
    Set rsPinPad = ObtieneCodigos(26) 'AMDO20150420
    Call CargaCombo(cboTipoPinPad, rsPinPad)
    
    For i = 0 To Me.cboTipoPinPad.ListCount - 1
        If CInt(Right(cboTipoPinPad.List(i), 3)) = gnTipoPinPad Then
            Exit For
        End If
    Next i
    If Me.cboTipoPinPad.ListCount <> i Then
        Me.cboTipoPinPad.ListIndex = i
    Else
        Me.cboTipoPinPad.ListIndex = -1
    End If
    
    For i = 0 To Me.CboPuerto.ListCount - 1
        If CInt(CboPuerto.List(i)) = gnPinPadPuerto Then
            Exit For
        End If
    Next i
    If Me.CboPuerto.ListCount <> i Then
        Me.CboPuerto.ListIndex = i
    Else
        Me.CboPuerto.ListIndex = -1
    End If
End Sub

