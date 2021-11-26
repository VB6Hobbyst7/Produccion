VERSION 5.00
Begin VB.Form frmRepTarjCancel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Tarjetas Canceladas"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmRepTarjCancel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   390
      Left            =   3780
      TabIndex        =   12
      Top             =   1957
      Width           =   1260
   End
   Begin VB.CommandButton CmdReportes 
      Caption         =   "Generar Reporte"
      Height          =   375
      Left            =   75
      TabIndex        =   11
      Top             =   1965
      Width           =   2220
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      Height          =   1890
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5070
      Begin VB.CheckBox ChkAge 
         Caption         =   "Incluir Todas las Agencias"
         Height          =   285
         Left            =   1410
         TabIndex        =   6
         Top             =   1500
         Width           =   2190
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1095
         Width           =   3600
      End
      Begin VB.CheckBox ChkUsu 
         Caption         =   "Incluir Todos los Usuarios"
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   720
         Width           =   2190
      End
      Begin VB.TextBox TxtUsu 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1410
         TabIndex        =   3
         Top             =   705
         Width           =   1035
      End
      Begin VB.TextBox txtFecFin 
         Height          =   330
         Left            =   3930
         TabIndex        =   2
         Text            =   "10/01/2008"
         Top             =   315
         Width           =   1020
      End
      Begin VB.TextBox txtfecIni 
         Height          =   330
         Left            =   1425
         TabIndex        =   1
         Text            =   "10/01/2008"
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label Label4 
         Caption         =   "Agencia :"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   1095
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo Usuario :"
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   750
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Fin :"
         Height          =   225
         Left            =   2655
         TabIndex        =   8
         Top             =   345
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Inicio :"
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmRepTarjCancel"
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
    
    If Len(Trim(TxtUsu.Text)) = 0 And ChkUsu.Value = 0 Then
        MsgBox "Usuario Incorrecto"
        TxtUsu.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    If CboAgencia.ListIndex = -1 And ChkAge.Value = 0 Then
        MsgBox "Agencia Incorrecta"
        CboAgencia.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    
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
'
'sSql = "Select cNumTarjeta, dFecBloqCancelTarj, ISNULL(P.cPersNombre,'') cPersNombre, "
'sSql = sSql & " T.cUserBloqCanc , A.cAgeDescripcion "
'sSql = sSql & " From tarjeta T"
'sSql = sSql & " LEFT JOIN agosto05..Persona P ON T.cPersCod = P.cPersCod "
'sSql = sSql & " Inner Join agosto05..Agencias A ON CONVERT(int,A.cAgeCod) = T.nCodAge"
'sSql = sSql & " where T.dFecBloqCancelTarj >= '" & Format(CDate(Me.txtfecIni.Text), "mm/dd/yyyy") & "'  and T.dFecBloqCancelTarj <='" & Format(CDate(Me.txtFecFin.Text), "mm/dd/yyyy") & "'"
'sSql = sSql & " AND T.nCondicion=50 "
'If Me.ChkUsu.Value <> 1 Then
'    sSql = sSql & " and T.cUserBloqCanc='" & Me.TxtUsu.Text & "'"
'End If
'If Me.ChkAge.Value <> 1 Then
'    sSql = sSql & " and T.nCodAge=" & Mid(Me.CboAgencia.Text, 1, 2)
'End If

sSQL = " REP_TarjetaCancel '" & Format(CDate(Me.txtfecIni.Text), "mm/dd/yyyy") & "','" & Format(CDate(Me.txtFecFin.Text), "mm/dd/yyyy") & "','" & Trim(Str(Me.ChkUsu.Value)) & "','" & Trim(Me.TxtUsu.Text) & "','" & Trim(Str(Me.ChkAge.Value)) & "','" & Trim(Me.CboAgencia.Text) & "'"

Set R = New ADODB.Recordset
sCadRep = "."

'Cabecera
sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(40) & "REPORTE DE TARJETAS CANCELADAS" & Chr(10) & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
sCadRep = sCadRep & Space(5) & "TARJETA" & Space(20) & "FECHA CANCEL." & Space(10) & "CLIENTE" & Space(14) & "USUARIO" & Space(7) & "AGENCIA" & Space(5) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
Cont = 0
'AbrirConexion
oConec.AbreConexion
R.Open sSQL, oConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Do While Not R.EOF
    sCadRep = sCadRep & Space(5) & Right(Space(16) & R!cNumTarjeta, 16) & Space(5) & Format(R!dFecBloqCancelTarj, "dd/mm/yyyy hh:mm:ss") & Space(5) & Left(R!cPersNombre & Space(30), 25) & Space(2) & Left(R!cUserBloqCanc & Space(10), 10) & Space(2) & Left(R!cAgeDescripcion & Space(20), 20) & Space(2) & Chr(10)
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
