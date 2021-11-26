VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRepMovimientoBovAgencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Movimientos de Boveda de Agencia"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   380
      Left            =   4440
      TabIndex        =   4
      Top             =   2400
      Width           =   1410
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   380
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CheckBox chkUsuario 
         Caption         =   "Incluir todos los usuarios"
         Height          =   285
         Left            =   3360
         TabIndex        =   11
         Top             =   1080
         Width           =   2190
      End
      Begin VB.ComboBox CboUsuDes 
         Height          =   315
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   4515
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   4530
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   330
         Left            =   4200
         TabIndex        =   5
         Top             =   1560
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaIni 
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Fin :"
         Height          =   225
         Left            =   2880
         TabIndex        =   8
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Inicio :"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia :"
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmRepMovimientoBovAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CboAgencia_Click()
    Call CargaUsuarios
End Sub

Private Sub CmdGenerar_Click()
    
    If Me.CboAgencia.ListIndex = -1 Then
        MsgBox "Seleccione una Agencia", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If chkUsuario.Value = 0 And Me.CboUsuDes.ListIndex = -1 Then
        MsgBox "Seleccione un Usuario", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If chkUsuario.Value = 1 Then
        Call ReporteMovimientosAgenciaUsuario(Mid(Me.CboAgencia.Text, 1, 2), "", Format(CDate(Me.txtFechaIni.Text), "yyyy/mm/dd"), Format(CDate(Me.txtFechaFin.Text), "yyyy/mm/dd"))
    Else
        Call ReporteMovimientosAgenciaUsuario(Mid(Me.CboAgencia.Text, 1, 2), Mid(CboUsuDes.Text, 1, InStr(CboUsuDes.Text, "-") - 1), Format(CDate(Me.txtFechaIni.Text), "yyyy/mm/dd"), Format(CDate(Me.txtFechaFin.Text), "yyyy/mm/dd"))
    End If
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim R As ADODB.Recordset
Dim sSQL As String

    Set oConec = New DConecta
    
    sSQL = "ATM_DevuelveAgencias "
    oConec.AbreConexion
    
    Me.CboAgencia.Clear
    
    Set R = New ADODB.Recordset
    R.Open sSQL, oConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not R.EOF
        CboAgencia.AddItem (R!cAgeCod & Space(1) & R!cAgeDescripcion)
        R.MoveNext
    Loop
    R.Close

    oConec.CierraConexion
    
    Me.txtFechaFin.Text = Format(gdFecSis, "DD/MM/YYYY")
    Me.txtFechaIni.Text = Format(gdFecSis, "DD/MM/YYYY")
    
    Call CargaUsuarios
    
    Set R = Nothing
End Sub

Private Sub CargaUsuarios()
Dim R As ADODB.Recordset
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
    
    If Len(Me.CboAgencia.Text) = 0 Then Exit Sub
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(Mid(Me.CboAgencia.Text, 1, 2)))
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaUsuarios"
    
    Set R = Cmd.Execute
    CboUsuDes.Clear
    Do While Not R.EOF
         CboUsuDes.AddItem R!Codigo & "-" & R!Nombre
       R.MoveNext
    Loop
    R.Close
    oConec.CierraConexion
    
    Set R = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub


Public Sub ReporteMovimientosAgenciaUsuario(ByVal pnCodAge As Integer, ByVal psCodUsu As String, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date)
Dim loPrevio As Previo.clsPrevio
Dim lRs As ADODB.Recordset
Dim lsSQL As String
Dim lsCadRep As String
Dim lnCont As Integer
    
    lsSQL = "exec stp_sel_RepMovimientoBovedaAgencia " & pnCodAge & ",'" & psCodUsu & "','" & Format(pdFechaIni, "yyyy-mm-dd") & "','" & Format(pdFechaFin, "yyyy-mm-dd") & "'"
    
    Set lRs = New ADODB.Recordset
    lsCadRep = "."
    
    'Cabecera
    lsCadRep = lsCadRep & Space(1) & Left("CMAC MAYNAS S.A." & Space(70), 70) & " Fecha   : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    lsCadRep = lsCadRep & Space(2) & Left(gsNomAge & Space(70), 70) & " Usuario : " & gsCodUser & Chr(10)
    lsCadRep = lsCadRep & Space(30) & "REPORTE DE MOVIMIENTOS BOVEDA - AGENCIA" & Chr(10)
    lsCadRep = lsCadRep & Space(35) & "DEL" & Format(pdFechaIni, "dd/mm/yyyy") & " AL " & Format(pdFechaFin, "dd/mm/yyyy") & Chr(10) & Chr(10)
    lsCadRep = lsCadRep & Space(2) & String(100, "-") & Chr(10)
    lsCadRep = lsCadRep & Space(2) & "Agencia                    Fecha       Usuario  Tar. Habil.  Tar. Devuelt." & Chr(10)
    lsCadRep = lsCadRep & Space(2) & String(100, "-") & Chr(10)
    lnCont = 0

    oConec.AbreConexion
    lRs.Open lsSQL, oConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not lRs.EOF
        If DateDiff("d", lRs!dFecha, "1800-01-01") <> 0 And lRs!cCodUsu <> "TOTAL" Then
            lsCadRep = lsCadRep & Space(2) & Left(lRs!cNomAgeArea & Space(25), 25) & Space(2)
            lsCadRep = lsCadRep & lRs!dFecha & Space(4)
            lsCadRep = lsCadRep & lRs!cCodUsu & Space(2)
            lsCadRep = lsCadRep & Right(Space(12) & lRs!Habilitacion, 12) & Space(2)
            lsCadRep = lsCadRep & Right(Space(12) & lRs!Devolucion, 12) & Chr(10)
        ElseIf lRs!cNomAgeArea = "TOTAL" Then
            lsCadRep = lsCadRep & Space(2) & String(100, "-") & Chr(10)
            lsCadRep = lsCadRep & Space(39) & "TOTALES   " & Right(Space(12) & lRs!Habilitacion, 12) & Space(2) & Right(Space(12) & lRs!Devolucion, 12) & Chr(10)
        End If
        lnCont = lnCont + 1
        lRs.MoveNext
    Loop
    
    lRs.Close
    oConec.CierraConexion
    Set lRs = Nothing
    
    Set loPrevio = New Previo.clsPrevio
    Call loPrevio.Show(lsCadRep, "REPORTE")
    Set loPrevio = Nothing
    
End Sub

