VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantPermisosNuevo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Permisos"
   ClientHeight    =   7050
   ClientLeft      =   2625
   ClientTop       =   1275
   ClientWidth     =   6795
   Icon            =   "frmMantPermisosNuevo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabOpc 
      Height          =   2895
      Left            =   105
      TabIndex        =   6
      Top             =   3405
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Menu"
      TabPicture(0)   =   "frmMantPermisosNuevo.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TVMenu"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Operaciones"
      TabPicture(1)   =   "frmMantPermisosNuevo.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TVOperacion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSComctlLib.TreeView TVMenu 
         Height          =   2370
         Left            =   -74880
         TabIndex        =   7
         Top             =   420
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   4180
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView TVOperacion 
         Height          =   2355
         Left            =   120
         TabIndex        =   8
         Top             =   450
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   4154
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   120
      TabIndex        =   3
      Top             =   6315
      Width           =   6510
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   435
         Left            =   4950
         TabIndex        =   5
         Top             =   180
         Width           =   1470
      End
      Begin VB.CommandButton CmdReGenerar 
         Caption         =   "&Regenerar Menu"
         Height          =   435
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Actualizar la Base de Datos con el Menu"
         Top             =   180
         Width           =   1470
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin TabDlg.SSTab SSFichas 
      Height          =   3240
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   5715
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Usuarios"
      TabPicture(0)   =   "frmMantPermisosNuevo.frx":0044
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LstUsuarios"
      Tab(0).Control(1)=   "CmdMostrarNombre"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Grupos"
      TabPicture(1)   =   "frmMantPermisosNuevo.frx":0060
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LstGrupos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CommandButton CmdMostrarNombre 
         Caption         =   "&Mostrar Nombres"
         Height          =   300
         Left            =   -70125
         TabIndex        =   9
         Top             =   0
         Width           =   1665
      End
      Begin MSComctlLib.ListView LstGrupos 
         Height          =   2535
         Left            =   135
         TabIndex        =   2
         Top             =   435
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Grupo"
            Object.Width           =   9701
         EndProperty
      End
      Begin MSComctlLib.ListView LstUsuarios 
         Height          =   2580
         Left            =   -74865
         TabIndex        =   1
         Top             =   450
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   4551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   6703
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMantPermisosNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCadMenu As String
Dim sCadMenuGrp As String
Dim oAcceso As UAcceso
Dim RsMenu As ADODB.Recordset
Dim oPista As COMManejador.Pista

Private Sub CargaTVewOperaciones(ByVal rsUsu As ADODB.Recordset)
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node
Dim n As Node


sOpePadre = "P"
Set n = TVOperacion.Nodes.Add(, , sOpePadre, "Operaciones")
TVOperacion.Nodes.Clear

Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = TVOperacion.Nodes.Add(, , sOpePadre, sOperacion)
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = TVOperacion.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion)
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = TVOperacion.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion)
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = TVOperacion.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion)
            nodOpe.Tag = sOpeCod
    End Select
    nodOpe.Expanded = False
    rsUsu.MoveNext
Loop

End Sub

Private Function TodosMarcados(Node As Node) As Boolean
    If Node.Checked = False Then
        TodosMarcados = False
    Else
        TodosMarcados = True
        If Node.Text <> Node.LastSibling Then
            TodosMarcados = TodosMarcados(Node.Next)
        End If
        If TodosMarcados = False Then
            TodosMarcados = False
            Exit Function
        End If
        If Node.Children > 0 Then
            TodosMarcados = TodosMarcados(Node.Child)
        Else
            TodosMarcados = True
        End If
    End If
End Function

Private Function CargaAccesos(ByRef Node As Node, Optional ByVal pbOPeracion As Boolean = False) As Node
        
        If Not pbOPeracion Then
        
            If InStr(sCadMenu, Node.Tag) > 0 Then 'Si es Mayor que Cero Tiene Permiso
                Node.Checked = True
            Else
                Node.Checked = False
            End If
            If InStr(sCadMenuGrp, Node.Tag) > 0 Then 'Si es Mayor que Cero Tiene Permiso
                Node.ForeColor = vbBlue
            Else
                Node.ForeColor = vbBlack
            End If
            If Node.Text <> Node.LastSibling Then
                Call CargaAccesos(Node.Next)
            End If
    
            If Node.Children > 0 Then
                Call CargaAccesos(Node.Child)
            End If
        Else
            If InStr(sCadMenu, "*" & Node.Tag) > 0 Then 'Si es Mayor que Cero Tiene Permiso
                Node.Checked = True
            Else
                Node.Checked = False
            End If
            If InStr(sCadMenuGrp, "*" & Node.Tag) > 0 Then 'Si es Mayor que Cero Tiene Permiso
                Node.ForeColor = vbBlue
            Else
                Node.ForeColor = vbBlack
            End If
            If Node.Text <> Node.LastSibling Then
                Call CargaAccesos(Node.Next, True)
            End If
    
            If Node.Children > 0 Then
                Call CargaAccesos(Node.Child, True)
            End If
        End If
End Function


Private Function ActualizaMenuActivos(ByRef nPunt As Integer) As Integer

    'Si Es Nodo Final
    If MatMenuItems(nPunt).nPuntDer = -1 And MatMenuItems(nPunt).nPuntAbajo = -1 Then
        If oAcceso.TienePermiso(Left(MatMenuItems(nPunt).sName, 11), MatMenuItems(nPunt).sIndex) Then
            ActualizaMenuActivos = 1
            MatMenuItems(nPunt).bCheck = True
        Else
            ActualizaMenuActivos = 0
            MatMenuItems(nPunt).bCheck = False
        End If
    End If
    
    'Si Tiene mas Nodos Hijos
    If MatMenuItems(nPunt).nPuntDer <> -1 Then
        ActualizaMenuActivos = ActualizaMenuActivos(MatMenuItems(nPunt).nPuntDer)
        If ActualizaMenuActivos > 0 Then
            MatMenuItems(nPunt).bCheck = True
            ActualizaMenuActivos = 1
        Else
            ActualizaMenuActivos = 0
        End If
    End If

    'Si Tiene mas Nodos Paralelos
    If MatMenuItems(nPunt).nPuntAbajo <> -1 Then
        If oAcceso.TienePermiso(Left(MatMenuItems(nPunt).sName, 11), MatMenuItems(nPunt).sIndex) Then
            MatMenuItems(nPunt).bCheck = True
            ActualizaMenuActivos = 1
        End If
        ActualizaMenuActivos = ActualizaMenuActivos + ActualizaMenuActivos(MatMenuItems(nPunt).nPuntAbajo)
    End If
End Function


Private Function UbicaMenuActivos(ByVal psName As String, ByVal psIndex As String, ByVal nPunt As Integer) As Boolean
    
       'Si lo encuentro
    If Left(MatMenuItems(nPunt).sName, 11) = psName And MatMenuItems(nPunt).sIndex = psIndex Then
        If MatMenuItems(nPunt).bCheck Then
            UbicaMenuActivos = True
        Else
            UbicaMenuActivos = False
        End If
        Exit Function
    End If
    
    'Si Tiene mas Nodos Hijos
    If MatMenuItems(nPunt).nPuntDer <> -1 Then
        UbicaMenuActivos = UbicaMenuActivos(psName, psIndex, MatMenuItems(nPunt).nPuntDer)
        If UbicaMenuActivos Then
            Exit Function
        End If
    End If
    
    'Si Tiene mas Nodos Paralelos
    If MatMenuItems(nPunt).nPuntAbajo <> -1 Then
        UbicaMenuActivos = UbicaMenuActivos(psName, psIndex, MatMenuItems(nPunt).nPuntAbajo)
    End If
    
    'Si es Nodo Final
    If MatMenuItems(nPunt).nPuntDer = -1 And MatMenuItems(nPunt).nPuntAbajo = -1 Then
        UbicaMenuActivos = False
    End If
    
End Function


Private Sub HabilitarMenuparaDise?o()
Dim nPunt As Integer
'Habilitar el menu de tal forma que si algun hijo esta activo se
'Active tambien desde su padre
    nPunt = 0
    ActualizaMenuActivos nPunt
    
End Sub

Private Sub CargaMenuMDIMain(ByVal poAcceso As UAcceso)
Dim Ctl As Control
Dim sTipo As String
Dim nPos As Integer
Dim sCadMenuTemp As String

Call HabilitarMenuparaDise?o
On Error Resume Next
    For Each Ctl In frmMdiMain.Controls
        sTipo = TypeName(Ctl)
        If sTipo = "Menu" Then

            If UbicaMenuActivos(Ctl.Name, Format(Ctl.Index, "00"), 0) Then
                Ctl.Visible = True
            Else
                Ctl.Visible = False
            End If
        End If
    Next
End Sub

Private Sub CmdMostrarNombre_Click()
Dim i As Integer
    For i = 1 To LstUsuarios.ListItems.Count
        LstUsuarios.ListItems(i).SubItems(1) = oAcceso.MostarNombre(gsDominio, LstUsuarios.ListItems(i).Text)
        DoEvents
    Next i
End Sub

Private Sub CmdReGenerar_Click()

Dim sCadTemp As String
Dim sCadTempSql As String
Dim Ctl As Control
Dim MatMenu() As String
Dim NumMenus As Integer

    On Error GoTo ERRORCmdReGenerar_Click
    If MsgBox("Esta Seguro que Desea Regenerar el Menu", vbInformation + vbYesNo, "Aviso") = vbYes Then
        
        NumMenus = 0
        For Each Ctl In frmMdiMain.Controls
            If TypeName(Ctl) = "Menu" Then
                NumMenus = NumMenus + 1
            End If
        Next
        ReDim MatMenu(NumMenus, 3)
        NumMenus = 0
        For Each Ctl In frmMdiMain.Controls
            If TypeName(Ctl) = "Menu" Then
                MatMenu(NumMenus, 0) = Ctl.Name
                MatMenu(NumMenus, 1) = Ctl.Caption
                MatMenu(NumMenus, 2) = Ctl.Index
                NumMenus = NumMenus + 1
            End If
        Next
        
        
        Call oAcceso.RegeneraMenu(MatMenu)
        If oAcceso.vError = True Then
            MsgBox oAcceso.sMsgError, vbInformation, "Aviso"
        Else
            MsgBox "Se ha Actualizado el Menu sin Problemas", vbInformation, "Aviso"
        End If
        
        Call CargaMenuNuevo
        Call LstUsuarios_Click
    End If
    Exit Sub
    
ERRORCmdReGenerar_Click:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdsalir_Click()
    Call oAcceso.CargaMenu(gsDominio, gsCodUser)
    Call CargaMenuMDIMain(oAcceso)
    Unload Me
End Sub

Private Sub Form_Load()
Dim TVItem As Node
    CentraForm Me
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Screen.MousePointer = 11
    Set oAcceso = New UAcceso
    
Set oPista = New COMManejador.Pista
gsOpeCod = 691003

    Call CargaControles
    Screen.MousePointer = 0
End Sub

Private Sub CargaControles()
Call CargaUsuariosNuevo
Call CargaGruposNuevo
Call CargaMenuNuevo
Call CargaTVewOperacionesNuevo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.txtFecha.Text = "__/__/____"
    Me.txtFecha.Enabled = False
    Set oAcceso = Nothing
End Sub

Private Sub LstGrupos_Click()

    If LstGrupos.ListItems.Count > 0 Then
        If Me.txtFecha.Enabled = False Then
            Me.txtFecha.Enabled = True
        End If
        Call oAcceso.CargaMenu(gsDominio, LstGrupos.SelectedItem.Text, "2", , , , gdFecSis)
        sCadMenu = oAcceso.sCadMenu
        sCadMenuGrp = oAcceso.sCadMenuGrp
        
        'Para Acceso a Menu
        Call CargaAccesos(TVMenu.Nodes(1))
        If TVMenu.Nodes.Count > 0 Then
            TVMenu.Nodes(1).Expanded = True
        End If
        If TVMenu.Nodes.Count > 1 Then
            TVMenu.Nodes(1).Checked = TodosMarcados(TVMenu.Nodes(1).Child)
        End If
        
        'Para Acceso a Operaciones
        Call CargaAccesos(TVOperacion.Nodes(1), True)
        If TVOperacion.Nodes.Count > 0 Then
            TVOperacion.Nodes(1).Expanded = True
        End If
        
    End If
End Sub

Private Sub LstGrupos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstGrupos.SortKey = ColumnHeader.SubItemIndex
    LstGrupos.SortOrder = lvwAscending
    LstGrupos.Sorted = True
End Sub

Private Sub LstUsuarios_Click()
    If LstUsuarios.ListItems.Count > 0 Then

        If Me.txtFecha.Enabled = False Then
            Me.txtFecha.Enabled = True
        End If
        'se agrego parametro fecha
        Call oAcceso.CargaMenu(gsDominio, LstUsuarios.SelectedItem.Text, "1", , , , gdFecSis)
        sCadMenu = oAcceso.sCadMenu
        sCadMenuGrp = oAcceso.sCadMenuGrp
        
        'Para Accesos al Menu
        If TVMenu.Nodes.Count > 1 Then
            Call CargaAccesos(TVMenu.Nodes(1).Child)
            TVMenu.Nodes(1).Expanded = True
            TVMenu.Nodes(1).Checked = TodosMarcados(TVMenu.Nodes(1).Child)
        End If
        
        'Para Acceso a Operaciones
        If TVOperacion.Nodes.Count > 1 Then
            Call CargaAccesos(TVOperacion.Nodes(1), True)
            TVOperacion.Nodes(1).Expanded = True
            TVOperacion.Nodes(1).Checked = TodosMarcados(TVOperacion.Nodes(1).Child)
        End If
    End If
End Sub

Private Sub LstUsuarios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstUsuarios.SortKey = ColumnHeader.SubItemIndex
    LstUsuarios.SortOrder = lvwAscending
    LstUsuarios.Sorted = True
End Sub
Private Sub PermisoItem(ByRef Node As Node, ByVal bChecked As Boolean)

Dim sTipoUsu As String
Dim sUsuario As String
        
        If SSFichas.Tab = 0 Then
            sTipoUsu = "1"
            sUsuario = Trim(LstUsuarios.SelectedItem.Text)
        Else
            sTipoUsu = "2"
            sUsuario = Trim(LstGrupos.SelectedItem.Text)
        End If
        

        If UCase(Node.Text) <> "MENU" Then
            If bChecked = True Then
                If Not Node.Checked Then
                    If IsDate(txtFecha) Then
                        If (txtFecha >= gdFecSis) Then
                            Call oAcceso.OtorgarOperacion(sUsuario, Node.Tag, sTipoUsu, CDate(Me.txtFecha.Text))
                        Else
                            MsgBox "La fecha ingresada debe ser Mayor o Igual a la del Sistema", vbInformation, "Aviso"
                            Exit Sub
                        End If
                    Else
                        Call oAcceso.OtorgarOperacion(sUsuario, Node.Tag, sTipoUsu)
                    End If
                    Node.Checked = True
                End If
            Else
                Call oAcceso.DenegarAcceso(sUsuario, Node.Tag, sTipoUsu)
                Node.Checked = False
            End If
        End If
        If oAcceso.vError Then
            MsgBox oAcceso.sMsgError, vbInformation, "Aviso"
        End If
        
        If Node.Text <> Node.LastSibling Then
            Call PermisoItem(Node.Next, bChecked)
        End If
        If Node.Children > 0 Then
            Call PermisoItem(Node.Child, bChecked)
        End If
End Sub

Private Sub PermisoItemOperacion(ByRef Node As Node, ByVal bChecked As Boolean)

Dim sTipoUsu As String
Dim sUsuario As String
        
        If SSFichas.Tab = 0 Then
            sTipoUsu = "1"
            sUsuario = Trim(LstUsuarios.SelectedItem.Text)
        Else
            sTipoUsu = "2"
            sUsuario = Trim(LstGrupos.SelectedItem.Text)
        End If
        

        If UCase(Node.Text) <> "MENU" Then
            If bChecked = True Then
                If Not Node.Checked Then
                    If IsDate(txtFecha) Then
                        If (txtFecha >= gdFecSis) Then
                            Call oAcceso.OtorgarOperacion(sUsuario, Node.Tag, sTipoUsu, CDate(Me.txtFecha.Text))
                        Else
                            MsgBox "La fecha ingresada debe ser Mayor o Igual a la del Sistema", vbInformation, "Aviso"
                            Exit Sub
                        End If
                    Else
                        Call oAcceso.OtorgarOperacion(sUsuario, Node.Tag, sTipoUsu)
                    End If
                    Node.Checked = True
                End If
            Else
                Call oAcceso.DenegarOperacion(sUsuario, Node.Tag, sTipoUsu)
                Node.Checked = False
            End If
        End If
        If oAcceso.vError Then
            MsgBox oAcceso.sMsgError, vbInformation, "Aviso"
        End If
        
        If Node.Text <> Node.LastSibling Then
            Call PermisoItemOperacion(Node.Next, bChecked)
        End If
        If Node.Children > 0 Then
            Call PermisoItemOperacion(Node.Child, bChecked)
        End If
End Sub
Private Sub LstUsuarios_KeyUp(KeyCode As Integer, Shift As Integer)
    Call LstUsuarios_Click
End Sub

Private Sub TVMenu_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim sTipoUsu As String
Dim sUsuario As String
Dim NodeTemp As Node
Dim oNCOMContFunciones As NContFunciones
Dim lsMovNro As String

        Screen.MousePointer = 11
        
        If SSFichas.Tab = 0 Then
            sTipoUsu = "1"
            sUsuario = Trim(LstUsuarios.SelectedItem.Text)
        Else
            sTipoUsu = "2"
            sUsuario = Trim(LstGrupos.SelectedItem.Text)
        End If

        If UCase(Node.Text) <> "MENU" Then
            If Node.Checked = True Then
              If IsDate(txtFecha) Then
                    If (txtFecha >= gdFecSis) Then
                        Call oAcceso.OtorgarOperacion(sUsuario, Node.Tag, sTipoUsu, CDate(Me.txtFecha.Text))
                        Set oNCOMContFunciones = New NContFunciones
                        lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                        Call oPista.InsertarPista(gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Otorga Permiso a " & IIf(sTipoUsu = "1", "USUARIO " & UCase(sUsuario), UCase(sUsuario)) & " la opci?n " & CStr(Node.Tag) & " - " & Replace(Node.Text, "&", "") & " cuya ruta es " & Replace(CStr(Node.FullPath), "&", ""))
                        Set oNCOMContFunciones = Nothing
                    Else
                        MsgBox "La fecha ingresada debe ser Mayor o Igual a la del Sistema", vbInformation, "Aviso"
                        Exit Sub
                    End If
                Else
                    Call oAcceso.OtorgarOperacion(sUsuario, Node.Tag, sTipoUsu)
                    Set oNCOMContFunciones = New NContFunciones
                    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Call oPista.InsertarPista(gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Otorga Permiso a " & IIf(sTipoUsu = "1", "USUARIO " & UCase(sUsuario), UCase(sUsuario)) & " la opci?n " & CStr(Node.Tag) & " - " & Replace(Node.Text, "&", "") & " cuya ruta es " & Replace(CStr(Node.FullPath), "&", ""))
                    Set oNCOMContFunciones = Nothing
                End If
            Else
                Call oAcceso.DenegarOperacion(sUsuario, Node.Tag, sTipoUsu)
                Set oNCOMContFunciones = New NContFunciones
                lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Call oPista.InsertarPista(gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Denegar Permiso a " & IIf(sTipoUsu = "1", "USUARIO " & UCase(sUsuario), UCase(sUsuario)) & " la opci?n " & CStr(Node.Tag) & " - " & Replace(Node.Text, "&", "") & " cuya ruta es " & Replace(CStr(Node.FullPath), "&", ""))
                Set oNCOMContFunciones = Nothing
            End If
        End If
        If oAcceso.vError Then
            MsgBox oAcceso.sMsgError, vbInformation, "Aviso"
        End If
        
        Screen.MousePointer = 0
End Sub

Private Sub TVOperacion_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim sTipoUsu As String
Dim sUsuario As String
Dim NodeTemp As Node
Dim oNCOMContFunciones As NContFunciones
Dim lsMovNro As String

        Screen.MousePointer = 11
        If SSFichas.Tab = 0 Then
            sTipoUsu = "1"
            sUsuario = Trim(LstUsuarios.SelectedItem.Text)
        Else
            sTipoUsu = "2"
            sUsuario = Trim(LstGrupos.SelectedItem.Text)
        End If

        If UCase(Node.Text) <> "MENU" Then
            If Node.Checked = True Then
                If IsDate(txtFecha) Then
                    Call oAcceso.OtorgarOperacion(sUsuario, Node.Tag, sTipoUsu, CDate(Me.txtFecha.Text))
                    Set oNCOMContFunciones = New NContFunciones
                    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Call oPista.InsertarPista(gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Otorga Permiso a " & IIf(sTipoUsu = "1", "USUARIO " & UCase(sUsuario), UCase(sUsuario)) & " la operaci?n " & Node.Text & " cuya ruta es " & Replace(CStr(Node.FullPath), "&", ""))
                    Set oNCOMContFunciones = Nothing
                Else
                    Call oAcceso.OtorgarOperacion(sUsuario, Node.Tag, sTipoUsu)
                    Set oNCOMContFunciones = New NContFunciones
                    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Call oPista.InsertarPista(gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Otorga Permiso a " & IIf(sTipoUsu = "1", "USUARIO " & UCase(sUsuario), UCase(sUsuario)) & " la operaci?n " & Node.Text & " cuya ruta es " & Replace(CStr(Node.FullPath), "&", ""))
                    Set oNCOMContFunciones = Nothing
                End If
            Else
                Call oAcceso.DenegarOperacion(sUsuario, Node.Tag, sTipoUsu)
                Set oNCOMContFunciones = New NContFunciones
                lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Call oPista.InsertarPista(gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Denegar Permiso a " & IIf(sTipoUsu = "1", "USUARIO " & UCase(sUsuario), UCase(sUsuario)) & " la operaci?n " & Node.Text & " cuya ruta es " & Replace(CStr(Node.FullPath), "&", ""))
                Set oNCOMContFunciones = Nothing
                If SSTabOpc.Tab = 0 Then
                    Call oAcceso.DenegarOperacion(sUsuario, Mid(Node.Tag, 1, 3) & "0000000000", sTipoUsu)
                    Set oNCOMContFunciones = New NContFunciones
                    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Call oPista.InsertarPista(gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Denegar Permiso a " & IIf(sTipoUsu = "1", "USUARIO " & UCase(sUsuario), UCase(sUsuario)) & " la operaci?n " & Node.Text & " cuya ruta es " & Replace(CStr(Node.FullPath), "&", ""))
                    Set oNCOMContFunciones = Nothing
                    Set NodeTemp = Node.Parent
                    If Not NodeTemp Is Nothing Then
                        Do
                            NodeTemp.Checked = False
                            Call oAcceso.DenegarOperacion(sUsuario, NodeTemp.Tag, sTipoUsu)
                            Set oNCOMContFunciones = New NContFunciones
                            lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                            Call oPista.InsertarPista(gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Denegar Permiso a " & IIf(sTipoUsu = "1", "USUARIO " & UCase(sUsuario), UCase(sUsuario)) & " la operaci?n " & Node.Text & " cuya ruta es " & Replace(CStr(Node.FullPath), "&", ""))
                            Set oNCOMContFunciones = Nothing
                            Set NodeTemp = NodeTemp.Parent
                            If NodeTemp Is Nothing Then
                                Exit Do
                            End If
                        Loop
                    End If
                End If
            End If
        End If
        If oAcceso.vError Then
            MsgBox oAcceso.sMsgError, vbInformation, "Aviso"
        End If
                    
        If SSTabOpc.Tab = 0 Then
            If Node.Children > 0 Then
                Call PermisoItemOperacion(Node.Child, Node.Checked)
            End If
            TVOperacion.Nodes(1).Checked = TodosMarcados(TVOperacion.Nodes(1).Child)
        End If
        Screen.MousePointer = 0
End Sub

Private Sub CargaUsuariosNuevo()
Dim i As Integer
Dim sCadTemp As String
Dim L As ListItem
    Set oAcceso = New UAcceso
        
    Call oAcceso.CargaControlUsuarios(gsDominio)
    sCadTemp = oAcceso.DameUsuario
    LstUsuarios.ListItems.Clear
    Do While sCadTemp <> ""
        Set L = LstUsuarios.ListItems.Add(, , sCadTemp)
        L.SubItems(1) = oAcceso.DameUsuarioNameFull
        sCadTemp = oAcceso.DameUsuario
    Loop
End Sub

Private Sub CargaGruposNuevo()

Dim L As ListItem
Dim sCadTemp As String
    
    Call oAcceso.CargaControlGrupos(gsDominio)
    sCadTemp = oAcceso.DameGrupo
    LstGrupos.ListItems.Clear
    Do While sCadTemp <> ""
        Set L = LstGrupos.ListItems.Add(, , sCadTemp)
        sCadTemp = oAcceso.DameUsuario
    Loop
    
End Sub

Private Sub CargaTVewOperacionesNuevo()
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node
Dim n As Node

Set rsUsu = oAcceso.DameRSOperaciones
sOpePadre = "P"
Set n = TVOperacion.Nodes.Add(, , sOpePadre, "Operaciones")
TVOperacion.Nodes.Clear

Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = TVOperacion.Nodes.Add(, , sOpePadre, sOperacion)
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = TVOperacion.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion)
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = TVOperacion.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion)
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = TVOperacion.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion)
            nodOpe.Tag = sOpeCod
    End Select
    nodOpe.Expanded = False
    rsUsu.MoveNext
Loop
RSClose rsUsu
End Sub

Private Sub CargaMenuNuevo()

Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim sTemp As String
Dim sTemp2 As String
Dim sTemp3 As String
Dim sTemp4 As String
Dim sTemp5 As String
Dim n As Node
Dim sPadre As String
Dim MatMenu As Variant
Dim R As ADODB.Recordset
Dim Y As Integer



    Set R = oAcceso.DameItemsMenu
    i = 0
    ReDim MatMenuItems(0)
    ReDim Preserve MatMenuItems(i + 1)
    MatMenuItems(i).nId = i
    MatMenuItems(i).sCodigo = Trim(R!cCodigo)
    MatMenuItems(i).sCaption = Trim(R!cDescrip)
    MatMenuItems(i).sName = Trim(R!cname)
    MatMenuItems(i).sIndex = Right(R!cname, 2)
    MatMenuItems(i).bCheck = False
    MatMenuItems(i).nNivel = 1
    MatMenuItems(i).nPuntDer = -1
    MatMenuItems(i).nPuntAbajo = -1
    i = i + 1
    Y = i
    R.MoveNext
    Call CargaMenuArbol(R, i, Y)
    R.Close
    
    TVMenu.Nodes.Clear
    Set n = TVMenu.Nodes.Add(, , "M", "Menu")
    sPadre = "M"
    For i = 0 To UBound(MatMenuItems) - 1
        If MatMenuItems(i).nNivel = 1 Then
            Set n = TVMenu.Nodes.Add("M", tvwChild, "N" & MatMenuItems(i).sCodigo, MatMenuItems(i).sCaption)
            n.Tag = MatMenuItems(i).sName
            sTemp = "N" & MatMenuItems(i).sCodigo
        Else
            If MatMenuItems(i).nNivel = 2 Then
                Set n = TVMenu.Nodes.Add(sTemp, tvwChild, "N2" & MatMenuItems(i).sCodigo, MatMenuItems(i).sCaption)
                n.Tag = MatMenuItems(i).sName
                sTemp2 = "N2" & MatMenuItems(i).sCodigo
            Else
                If MatMenuItems(i).nNivel = 3 Then
                    Set n = TVMenu.Nodes.Add(sTemp2, tvwChild, "N3" & MatMenuItems(i).sCodigo, MatMenuItems(i).sCaption)
                    n.Tag = MatMenuItems(i).sName
                    sTemp3 = "N3" & MatMenuItems(i).sCodigo
                Else
                    If MatMenuItems(i).nNivel = 4 Then
                        Set n = TVMenu.Nodes.Add(sTemp3, tvwChild, "N4" & MatMenuItems(i).sCodigo, MatMenuItems(i).sCaption)
                        n.Tag = MatMenuItems(i).sName
                        sTemp4 = "N4" & MatMenuItems(i).sCodigo
                    Else 'Nivel 5
                        Set n = TVMenu.Nodes.Add(sTemp4, tvwChild, "N5" & MatMenuItems(i).sCodigo, MatMenuItems(i).sCaption)
                        n.Tag = MatMenuItems(i).sName
                        sTemp5 = "N5" & MatMenuItems(i).sCodigo
                    End If
                End If
            End If
        End If
    Next i
    sCadMenu = oAcceso.sCadMenu
    TVMenu.Nodes(1).Expanded = True
    
End Sub
