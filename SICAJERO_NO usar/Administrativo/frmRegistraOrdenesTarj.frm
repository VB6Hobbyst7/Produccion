VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRegistraOrdenesTarj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar de Tarjetas Ordenadas"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   Icon            =   "frmRegistraOrdenesTarj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   3375
      Width           =   9420
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2355
         Top             =   255
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   480
         Left            =   7680
         TabIndex        =   4
         Top             =   240
         Width           =   1650
      End
      Begin VB.CommandButton CmdImportar 
         Caption         =   "Importar Archivo"
         Height          =   480
         Left            =   75
         TabIndex        =   3
         Top             =   240
         Width           =   1650
      End
      Begin VB.CommandButton CmdRegistrar 
         Caption         =   "Registrar Archivo"
         Height          =   480
         Left            =   5985
         TabIndex        =   2
         Top             =   240
         Width           =   1650
      End
   End
   Begin MSComctlLib.ListView LstTarjetas 
      Height          =   3165
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   5583
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TARJETA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PVV"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CADUCIDAD"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmRegistraOrdenesTarj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNomFile As String
Dim sPathFile As String

Private Sub ImportarArchivo()
Dim sSQL As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim L As ListItem

    Set cn = New ADODB.Connection
    'cn.ConnectionString = "DSN=DSNPin"
    'cn.ConnectionString = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};DBQ=C:\GlobalNET\;" '"", "", "
    cn.ConnectionString = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & sPathFile & ";" '"", "", "
    cn.Open

    'sSql = "select * from CardFileNameOutput#txt"
    sSQL = "select * from " & Replace(sNomFile, ".", "#")
    Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.Open sSQL, cn
    'Set dtgPin.DataSource = Rs

    'sCodCajMayBCR = "109"
    'dtpFecha.Value = Date
    'dtpFechaVenc.Value = Date
    Me.LstTarjetas.ListItems.Clear
    Do While Not rs.EOF
    
        Set L = Me.LstTarjetas.ListItems.Add(, , Replace(rs.Fields(0).Value, " ", ""))
        Call L.ListSubItems.Add(, , Right("0000" & rs.Fields(5), 4))
        Call L.ListSubItems.Add(, , rs.Fields(1).Value)

'        Set L = Me.LstTarjetas.ListItems.Add(, , Mid(Rs.Fields(1).Value, 1, 16))
'        Call L.ListSubItems.Add(, , Mid(Rs.Fields(1).Value, 45, 4))
'        Call L.ListSubItems.Add(, , Mid(Rs.Fields(0).Value, 21, 5))
        
        rs.MoveNext
    Loop
End Sub

Public Sub RegistrarArchivo()
Dim i As Integer
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim dFecVenc As Date
Dim loConec As New DConecta

    If Me.LstTarjetas.ListItems.Count = 0 Then
        MsgBox "No Exsten Datos para Registrar", vbInformation, "Aviso"
        Exit Sub
    End If

    If MsgBox("Se van a Actualizar los Datos de las Tarjetas de la Lista, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    loConec.AbreConexion
    For i = 1 To Me.LstTarjetas.ListItems.Count
        
        dFecVenc = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Mid(LstTarjetas.ListItems(i).SubItems(2), 4, 2) & "/20" & Mid(LstTarjetas.ListItems(i).SubItems(2), 1, 2))))
        
        Set Cmd = New ADODB.Command
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cNumTarjeta", adVarChar, adParamInput, 50, LstTarjetas.ListItems(i).Text)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nCondicion", adInteger, adParamInput, , -2)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nRetenerTarjeta", adInteger, adParamInput, , 0)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nCodAge", adInteger, adParamInput, , 0)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cUserAfil", adVarChar, adParamInput, 50, gsCodUser)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFecAfil", adDate, adParamInput, , Now)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cPersCod", adVarChar, adParamInput, 50, "")
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cPVV", adVarChar, adParamInput, 10, LstTarjetas.ListItems(i).SubItems(1))
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFecVenc", adDate, adParamInput, , dFecVenc)
        Cmd.Parameters.Append Prm
                

        Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "ATM_RegistraTarjeta"
        Cmd.Execute
        

        Set Cmd = Nothing
        Set Prm = Nothing
        
    Next i

    loConec.CierraConexion
    Set loConec = Nothing
    
    MsgBox "Datos Actualizados con Exito", vbInformation, "Aviso"

End Sub

Private Sub CmdImportar_Click()
Dim i As Integer

    CommonDialog1.ShowOpen
    sNomFile = CommonDialog1.FileName
    
    
    If Len(Trim(sNomFile)) = 0 Then
        MsgBox "Archivo Incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    
    For i = Len(sNomFile) - 1 To 1 Step -1
        If Mid(sNomFile, i, 1) = "\" Then
            sPathFile = Mid(sNomFile, 1, i)
            sNomFile = Mid(sNomFile, i + 1, Len(sNomFile) - i)
            Exit For
        End If
    Next i
    
    
    Call ImportarArchivo
    
End Sub

Private Sub CmdRegistrar_Click()
RegistrarArchivo
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    
End Sub


