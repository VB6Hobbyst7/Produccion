VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAfilCargaTarRechazadas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Archivo de Tarjetas Rechazadas"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "frmAfilCargaTarRechazadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   3285
      Width           =   9420
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   480
         Left            =   7680
         TabIndex        =   3
         Top             =   240
         Width           =   1650
      End
      Begin VB.CommandButton CmdImportar 
         Caption         =   "Importar Archivo"
         Height          =   480
         Left            =   75
         TabIndex        =   2
         Top             =   240
         Width           =   1650
      End
      Begin VB.CommandButton CmdRegistrar 
         Caption         =   "Registrar Archivo"
         Height          =   480
         Left            =   5985
         TabIndex        =   1
         Top             =   240
         Width           =   1650
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2355
         Top             =   255
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.ListView LstTarjetas 
      Height          =   3165
      Left            =   0
      TabIndex        =   4
      Top             =   0
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TARJETA"
         Object.Width           =   5644
      EndProperty
   End
End
Attribute VB_Name = "frmAfilCargaTarRechazadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNomFile As String
Dim sPathFile As String
Dim oConec As DConecta

Private Sub ImportarArchivo()
Dim sSQL As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim L As ListItem
Dim strValue As String

Me.LstTarjetas.ListItems.Clear
Open sNomFile For Input As #1
Do While Not EOF(1)
    Input #1, strValue

    Set L = Me.LstTarjetas.ListItems.Add(, , strValue)

Loop
Close #1


End Sub

Public Sub RegistrarArchivo()
Dim i As Integer
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim dFecVenc As Date

    If Me.LstTarjetas.ListItems.Count = 0 Then
        MsgBox "No Exsten Datos para Registrar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Se van a Actualizar los Datos de las Tarjetas de la Lista, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    For i = 1 To Me.LstTarjetas.ListItems.Count
                
        Set Cmd = New ADODB.Command
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cNumTarjeta", adVarChar, adParamInput, 50, LstTarjetas.ListItems(i).Text)
        Cmd.Parameters.Append Prm
                
        oConec.AbreConexion
        Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "ATM_RegistraTarjetaRechazada"
        Cmd.Execute
        
        
        Set Cmd = Nothing
        Set Prm = Nothing
        oConec.CierraConexion
        
    Next i

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


Private Sub Form_Load()
    Set oConec = New DConecta
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
