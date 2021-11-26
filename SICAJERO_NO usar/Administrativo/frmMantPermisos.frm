VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantPermisos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Permisos"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "frmMantPermisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   30
      TabIndex        =   2
      Top             =   3090
      Width           =   8700
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   405
         Left            =   7185
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   30
      TabIndex        =   0
      Top             =   105
      Width           =   8700
      Begin MSComctlLib.ListView LstGrupos 
         Height          =   1830
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   3228
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Menu"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   9172
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   810
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   8535
         Begin VB.ComboBox CboGrupos 
            Height          =   315
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   255
            Width           =   5985
         End
         Begin VB.Label Label1 
            Caption         =   "Grupos :"
            Height          =   225
            Left            =   150
            TabIndex        =   4
            Top             =   270
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frmMantPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta


Private Sub CboGrupos_Click()
    Call ActualizaListaDEMenus
End Sub


Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub ActualizaListaDEMenus()
Dim R As ADODB.Recordset
Dim i As Integer
Dim sSQL As String

    For i = 1 To LstGrupos.ListItems.Count - 1
          
        LstGrupos.ListItems(i).Checked = False
          
    Next i
        
    
    sSQL = "ATM_RecuperaPermisosDEGrupo '" & Me.CboGrupos.Text & "'"
    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    'AbrirConexion
    R.Open sSQL, oConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not R.EOF
        
        For i = 1 To LstGrupos.ListItems.Count - 1
            If UCase(LstGrupos.ListItems(i).Text) = UCase(R!cNomMenu) Then
                LstGrupos.ListItems(i).Checked = True
                Exit For
            End If
        Next i
    
        R.MoveNext
    Loop
    R.Close
    
    'CerrarConexion
    oConec.CierraConexion

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Ctl As Control
Dim L As ListItem
Dim sTipo As String

    Set oConec = New DConecta
    
    CargaControlGrupos (gsDominio)
    Me.CboGrupos.Clear
    
    For i = 0 To UBound(LstUsuarios) - 1
        Me.CboGrupos.AddItem (LstUsuarios(i))
    Next i
    
    For Each Ctl In MDIMenu.Controls
        sTipo = TypeName(Ctl)
        If sTipo = "Menu" Then
            Set L = Me.LstGrupos.ListItems.Add(, , Ctl.Name)
            Call L.ListSubItems.Add(, , Ctl.Caption)
        End If
    Next
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub LstGrupos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNomgrupo", adVarChar, adParamInput, 150, Me.CboGrupos.Text)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNomMenu", adVarChar, adParamInput, 150, Item.Text)
    Cmd.Parameters.Append Prm

    If Item.Checked Then
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@pnHab", adInteger, adParamInput, , 1)
        Cmd.Parameters.Append Prm
    Else
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@pnHab", adInteger, adParamInput, , 0)
        Cmd.Parameters.Append Prm
    End If
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistraPermiso"
    Cmd.Execute
    
    'CerrarConexion
    oConec.CierraConexion

    Set Prm = Nothing
    Set Cmd = Nothing
            
End Sub
