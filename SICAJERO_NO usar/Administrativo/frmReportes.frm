VERSION 5.00
Begin VB.Form frmReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes para Tarjeta Afiliadas"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   390
      Left            =   3630
      TabIndex        =   6
      Top             =   1125
      Width           =   1260
   End
   Begin VB.CommandButton CmdReportes 
      Caption         =   "Generar Reporte"
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   1155
      Width           =   2220
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      Height          =   900
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   5070
      Begin VB.CheckBox ChkAge 
         Caption         =   "Incluir Todas las Agencias"
         Height          =   285
         Left            =   1410
         TabIndex        =   12
         Top             =   1500
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1095
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.TextBox txtFecFin 
         Height          =   330
         Left            =   3930
         TabIndex        =   4
         Text            =   "10/01/2008"
         Top             =   315
         Width           =   1020
      End
      Begin VB.TextBox txtfecIni 
         Height          =   330
         Left            =   1425
         TabIndex        =   2
         Text            =   "10/01/2008"
         Top             =   315
         Width           =   1020
      End
      Begin VB.TextBox TxtUsu 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1410
         TabIndex        =   8
         Top             =   705
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CheckBox ChkUsu 
         Caption         =   "Incluir Todos los Usuarios"
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label4 
         Caption         =   "Agencia :"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   1095
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Fin :"
         Height          =   225
         Left            =   2655
         TabIndex        =   3
         Top             =   345
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Inicio :"
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo Usuario :"
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   750
         Visible         =   0   'False
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    
    If Not IsDate(Me.txtfecIni.Text) Then
        MsgBox "Fecha de Inicio Incorrecta"
        txtfecIni.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Not IsDate(Me.txtFecFin.Text) Then
        MsgBox "Fecha de Fin de Reporte Incorrecta"
        txtFecFin.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
'    If Len(Trim(TxtUsu.Text)) = 0 And ChkUsu.Value = 0 Then
'        MsgBox "Usuario Incorrecto"
'        TxtUsu.SetFocus
'        ValidaDatos = False
'        Exit Function
'    End If
'
'    If CboAgencia.ListIndex = -1 And ChkAge.Value = 0 Then
'        MsgBox "Agencia Incorrecta"
'        CboAgencia.SetFocus
'        ValidaDatos = False
'        Exit Function
'    End If
    
    
End Function


Private Sub CmdReportes_Click()
Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim Cont As Integer

If Not ValidaDatos Then
    Exit Sub
End If


'sSql = "Select cNumTarjeta, dFecAfil, ISNULL(P.cPersNombre,'') cPersNombre, "
'sSql = sSql & " T.cUserAfil , A.cAgeDescripcion "
'sSql = sSql & " From tarjeta T"
'sSql = sSql & " LEFT JOIN agosto05..Persona P ON T.cPersCod = P.cPersCod "
'sSql = sSql & " Inner Join agosto05..Agencias A ON CONVERT(int,A.cAgeCod) = T.nCodAge"
'sSql = sSql & " where T.dFecAfil >= '" & Format(CDate(Me.txtfecIni.Text), "mm/dd/yyyy") & "'  and T.dFecAfil <='" & Format(CDate(Me.txtFecFin.Text), "mm/dd/yyyy") & "'"
'If Me.ChkUsu.Value <> 1 Then
'    sSql = sSql & " and T.cUserAfil='" & Me.TxtUsu.Text & "'"
'End If
'If Me.ChkAge.Value <> 1 Then
'    sSql = sSql & " and T.nCodAge=" & Mid(Me.CboAgencia.Text, 1, 2)
'End If
'


Set R = New ADODB.Recordset
sCadRep = "."

'Cabecera
sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(40) & "REPORTE DE TARJETAS AFILIADAS" & Chr(10) & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
'sCadRep = sCadRep & Space(5) & "TARJETA" & Space(10) & "FECHA AFILIACION" & Space(10) & "CLIENTE" & Space(14) & "USUARIO" & Space(7) & "AGENCIA" & Space(5) & Chr(10)
sCadRep = sCadRep & Space(5) & "TARJETA" & Space(10) & "FECHA AFILIACION" & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)

'AbrirConexion
oConec.AbreConexion

sSQL = " REP_TarjetaAfiliadas '" & Format(CDate(Me.txtfecIni.Text), "mm/dd/yyyy") & "','" & Format(CDate(Me.txtFecFin.Text), "mm/dd/yyyy") & "','" & Trim(Str(Me.ChkUsu.Value)) & "','" & Trim(Me.TxtUsu.Text) & "','" & Trim(Str(Me.ChkAge.Value)) & "','" & Trim(Me.CboAgencia.Text) & "'"
Cont = 0
R.Open sSQL, oConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Do While Not R.EOF
    sCadRep = sCadRep & Space(5) & Right(Space(16) & R!cNumTarjeta, 16) & Space(5) & Format(R!dFecAfil, "dd/mm/yyyy") & Chr(10)
    'sCadRep = sCadRep & Space(5) & Right(Space(16) & R!cNumTarjeta, 16) & Space(5) & Format(R!dFecAfil, "dd/mm/yyyy") & Space(5) & Left(R!cPersNombre & Space(30), 25) & Space(2) & Left(R!cUserAfil & Space(20), 10) & Space(2) & Left(R!cAgeDescripcion & Space(30), 20) & Space(2) & Chr(10)
    Cont = Cont + 1
    R.MoveNext
Loop
R.Close
'CerrarConexion
oConec.CierraConexion
Set R = Nothing

sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
sCadRep = sCadRep & Space(5) & "CANTIDAD : " & Str(Cont) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)

    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim R As ADODB.Recordset
Dim sSQL As String

    Set oConec = New DConecta
    'sSql = "Select cAgeCod, cAgeDescripcion from agosto05..Agencias Order by cAgeCod"
    sSQL = "ATM_DevuelveAgencias "
    
    Me.CboAgencia.Clear
    'AbrirConexion
    oConec.AbreConexion
    Set R = New ADODB.Recordset
    R.Open sSQL, oConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not R.EOF
        CboAgencia.AddItem (R!cAgeCod & Space(1) & R!cAgeDescripcion)
        
        R.MoveNext
    Loop
    R.Close
    'CerrarConexion
    oConec.CierraConexion
    Set R = Nothing
    
    Me.txtfecIni.Text = Format(Now(), "dd/mm/yyyy")
    Me.txtFecFin.Text = Format(Now(), "dd/mm/yyyy")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
