VERSION 5.00
Begin VB.Form frmListaAnalistas 
   Caption         =   "Lista de Analistas"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "frmListaAnalistas.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8445
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Agencia"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   8040
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lista de Analistas"
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5895
      Begin VB.ListBox lstAnalistas 
         Height          =   5910
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analista"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5895
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lblUsuario 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmListaAnalistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cUser As String
'RECO20140621 ERS095-2014*********************************
Private Sub cboAgencia_Click()
    Me.lstAnalistas.Clear
    Call LlenarLista(IIf(cboAgencia.ItemData(cboAgencia.ListIndex) < 10, "0" & cboAgencia.ItemData(cboAgencia.ListIndex), cboAgencia.ItemData(cboAgencia.ListIndex)))
End Sub
'RECO FIN*************************************************

Private Sub cmdAceptar_Click()
    'Dim o As New frmHojaRutaAnalista
    cUser = lblUsuario.Caption
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'RECO20140621 ERS095-2014*********************************
    'Dim oCred As New COMDCredito.DCOMCreditos
    'Dim oDR As New ADODB.Recordset
    'Dim i As Integer
    
    'Set oDR = oCred.ListarAnalistasXAge(gsCodAge)
    
    'For i = 0 To oDR.RecordCount - 1
    '        lstAnalistas.AddItem (UCase(oDR!cUser) & " - " & oDR!cPersNombre)
    '    oDR.MoveNext
    'Next
    'Set oDR = Nothing
    Call CargaAgencias
    'RECO FIN************************************************
End Sub

Private Sub lstAnalistas_Click()
    Dim a As String
    lblUsuario.Caption = Mid(lstAnalistas.Text, 1, 4)
    lblNombre.Caption = Mid(lstAnalistas.Text, 8, Len(lstAnalistas.Text))
    a = lstAnalistas.Text
End Sub
'RECO20140621 ERS095-2014*********************************
Private Sub CargaAgencias()
    Dim clsAge As COMDConstSistema.DCOMGeneral
    Dim rs As ADODB.Recordset
    Dim i As Long
    Dim lnAgeCodAct As Integer
    Set rs = New ADODB.Recordset
    Set clsAge = New COMDConstSistema.DCOMGeneral
    
    If gsCodCargo = "002026" Or gsCodCargo = "002036" Then
        Set rs = clsAge.DevuelveListaAgenciaJT(gsCodUser)
    Else
        Set rs = clsAge.GetNombreAgencias()
        Me.cboAgencia.Enabled = False
    End If
    Do While Not rs.EOF
        cboAgencia.AddItem "" & rs!cAgeDescripcion
        cboAgencia.ItemData(cboAgencia.NewIndex) = "" & rs!cAgeCod
        
        If rs!cAgeCod = gsCodAge Then
            lnAgeCodAct = i
        End If
        i = i + 1
        rs.MoveNext
    Loop
    cboAgencia.ListIndex = lnAgeCodAct
    Call LlenarLista(IIf(cboAgencia.ItemData(cboAgencia.ListIndex) < 10, "0" & cboAgencia.ItemData(cboAgencia.ListIndex), cboAgencia.ItemData(cboAgencia.ListIndex)))
    
    rs.Close
    Set clsAge = Nothing
    Set rs = Nothing
End Sub
Private Sub LlenarLista(ByVal psAgeCod As String)
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim oDR As New ADODB.Recordset
    Dim i As Integer
    Set oDR = oCred.ListarAnalistasXAge(psAgeCod)
    
    For i = 0 To oDR.RecordCount - 1
            lstAnalistas.AddItem (UCase(oDR!cUser) & " - " & oDR!cPersNombre)
        oDR.MoveNext
    Next
    Set oDR = Nothing
End Sub
'RECO FIN*************************************************
