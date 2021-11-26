VERSION 5.00
Begin VB.Form frmRepTarRechaz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Tarjetas Rechazadas"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmRepTarRechaz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      Height          =   930
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5070
      Begin VB.TextBox txtfecIni 
         Height          =   330
         Left            =   1425
         TabIndex        =   0
         Text            =   "10/01/2008"
         Top             =   315
         Width           =   1020
      End
      Begin VB.TextBox txtFecFin 
         Height          =   330
         Left            =   3930
         TabIndex        =   1
         Text            =   "10/01/2008"
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Inicio :"
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Fin :"
         Height          =   225
         Left            =   2655
         TabIndex        =   5
         Top             =   345
         Width           =   1305
      End
   End
   Begin VB.CommandButton CmdReportes 
      Caption         =   "Generar Reporte"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1050
      Width           =   2220
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   390
      Left            =   3795
      TabIndex        =   3
      Top             =   1050
      Width           =   1260
   End
End
Attribute VB_Name = "frmRepTarRechaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdReportes_Click()
Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim Cont As Integer
Dim loConec As New DConecta

If Not IsDate(Me.txtfecIni.Text) Or Not IsDate(Me.txtFecFin.Text) Then
    MsgBox "Fecha Incorrecta", vbInformation, "Aviso"
    Exit Sub
End If


sSQL = " REP_TarjetaRechazadas '" & Format(CDate(Me.txtfecIni.Text), "mm/dd/yyyy") & "','" & Format(CDate(Me.txtFecFin.Text), "mm/dd/yyyy") & "'"

Set R = New ADODB.Recordset
sCadRep = "."

'Cabecera
sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(40) & "REPORTE DE TARJETAS RECHAZADAS" & Chr(10) & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
sCadRep = sCadRep & Space(5) & "TARJETA" & Space(20) & "FECHA RECHAZO" & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
Cont = 0
'AbrirConexion
loConec.AbreConexion
R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Do While Not R.EOF
    sCadRep = sCadRep & Space(5) & Right(Space(16) & R!cNumTarjeta, 16) & Space(5) & Format(R!dFecRechazada, "dd/mm/yyyy hh:mm:ss") & Chr(10)
    Cont = Cont + 1
    R.MoveNext
Loop
R.Close
'CerrarConexion
loConec.CierraConexion
Set loConec = Nothing
Set R = Nothing

sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
sCadRep = sCadRep & Space(5) & "CANTIDAD : " & Str(Cont) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)

    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing
    
End Sub


Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.txtfecIni.Text = Format(Now, "dd/mm/yyyy")
    Me.txtFecFin.Text = Format(Now, "dd/mm/yyyy")
    
End Sub
