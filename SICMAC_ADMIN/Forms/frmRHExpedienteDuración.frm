VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmRHExpedienteDuración 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Expediente Duración"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   Icon            =   "frmRHExpedienteDuración.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Documento"
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4695
      Begin VB.ComboBox cboTpoDoc 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtNroDias 
         Height          =   315
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "N° Días"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Documento:"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Duración:"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblDocumento 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registro de Documento"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmRHExpedienteDuración"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CREADO POR ARLO20161221 ********************************************
'***REGISTRO DE EXPEDINTE                                           *
'********************************************************************
Option Explicit
Dim oConst As COMDConstSistema.NCOMConstSistema
Dim lnTipo As TipoOpe
Dim bError As Boolean
Dim ObtenerExpedientePersonal As Recordset
Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub
Public Sub CargarCmbTpoDoc()
    Dim Sql As String
    
    Dim rs As New ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    
    Sql = "SELECT  CON.nConsValor,CON.cConsDescripcion FROM Constante CON WHERE CON.nConsCod = '10021'"
    Sql = Sql + "AND CON.nConsValor <> '10021' AND CON.bEstado= 1 AND CON.nConsValor NOT IN (1,2,3,4,5,6,7,9,10,13,14,15,19,20) ORDER BY CON.nConsValor  ASC"
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Sub
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Set rs = Conn.CargaRecordSet(Sql)
    
    With rs
    Do Until .EOF
     cboTpoDoc.AddItem "" & rs!cConsDescripcion
     cboTpoDoc.ItemData(cboTpoDoc.NewIndex) = "" & rs!nConsValor
       .MoveNext
    Loop
    End With
    rs.Close
    
    Conn.CierraConexion
    Set Conn = Nothing
End Sub
Sub cboTpoDoc_Click()
txtNroDias.Text = ""
End Sub
Private Sub cmdCancelar_Click()
txtNroDias.Text = ""
cboTpoDoc.ListIndex = -1
End Sub

Private Sub cmdRegistrar_Click()
    Dim dFechaRegistro  As String
    Dim Conn As New DConecta
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset

    If cboTpoDoc.Text = "" Then
        'MsgBox "Ingrese el Tipo de Documento", vbCritical, "Aviso"
        MsgBox "Ingrese el Tipo de Documento", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtNroDias.Text)) = 0 Then
        'MsgBox "Ingrese un Periodo de Dias", vbCritical, "Aviso"
        MsgBox "Ingrese un Periodo de Días", vbInformation, "Aviso"
        Me.txtNroDias.SetFocus
        Exit Sub
    End If

    bError = False
    Sql = "spt_ins_RegistrarDuracionExpedienteRRHH '" & cboTpoDoc.ItemData(cboTpoDoc.ListIndex)
    Sql = Sql & "','" & txtNroDias.Text & "','" & Format(gdFecSis, "yyyyMMdd") & "'"
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        'Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Sub
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Set ObtenerExpedientePersonal = Conn.CargaRecordSet(Sql)
    Conn.CierraConexion
    Set Conn = Nothing
    'MsgBox "Nuevo expediente registrado", vbExclamation, "Aviso"
    MsgBox "Nuevo expediente registrado", vbInformation, "Aviso"
    txtNroDias.Text = ""
    cboTpoDoc.ListIndex = -1
End Sub
'Sub txtNroDias_Keypress(KeyAscii As Integer)
'        If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 13) Then
'        If (KeyAscii = 13) Then
'        cmdRegistrar.SetFocus
'        End If
'       ' Exit Sub
'        Else:
'                MsgBox "Solo se Admiten Numeros Enteros", vbCritical, "Aviso"
'        End If
'End Sub

Sub txtNroDias_Keypress(KeyAscii As Integer)
'If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
'  MsgBox "Solo se Admiten Numeros Enteros", vbCritical, "Aviso"
'  KeyAscii = 8
'End If
If KeyAscii = 13 Then
KeyAscii = 0
'SendKeys "{tab}"
cmdRegistrar.SetFocus
ElseIf KeyAscii <> 8 Then
If Not IsNumeric(Chr(KeyAscii)) Then
Beep
KeyAscii = 0
End If
End If
'If (KeyAscii = 13) Then
'cmdRegistrar.SetFocus
'End If
End Sub
Private Sub Form_Load()
Call CargarCmbTpoDoc
End Sub

Public Sub RegistrarDuracionExpediente()

End Sub

Public Function FechaVencimiento() As ADODB.Recordset
    Dim sSql As String
    Dim oConect As DConecta

    Set oConect = New DConecta
    oConect.AbreConexion
    sSql = "exec stp_sel_devolverFechaVencExp "
    Set FechaVencimiento = oConect.CargaRecordSet(sSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function
Public Function AvisarVencimientoExpediente() As ADODB.Recordset
    Dim sSql As String
    Dim oConect As DConecta

    Set oConect = New DConecta
    oConect.AbreConexion
    sSql = "exec stp_sel_ExpedientesProximoAVencer "
    Set AvisarVencimientoExpediente = oConect.CargaRecordSet(sSql)
    oConect.CierraConexion
    Set oConect = Nothing
End Function

