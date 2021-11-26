VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogBieSerMant_des 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Bienes/Servicios"
   ClientHeight    =   7740
   ClientLeft      =   990
   ClientTop       =   1200
   ClientWidth     =   10725
   Icon            =   "frmLogBieSerMant_des.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgList 
      Left            =   4575
      Top             =   7230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogBieSerMant_des.frx":000C
            Key             =   "abierto"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogBieSerMant_des.frx":035E
            Key             =   "cerrado"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   360
      Left            =   75
      TabIndex        =   5
      Top             =   7335
      Width           =   1080
   End
   Begin VB.CommandButton cmdBuscarSig 
      Caption         =   "Buscar S&iguiente >>>"
      Height          =   360
      Left            =   1200
      TabIndex        =   4
      Top             =   7335
      Width           =   2085
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Bienes / Servicios "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7305
      Left            =   30
      TabIndex        =   2
      Top             =   15
      Width           =   10725
      Begin TabDlg.SSTab SSTab1 
         Height          =   1830
         Left            =   75
         TabIndex        =   7
         Top             =   5415
         Width           =   10545
         _ExtentX        =   18600
         _ExtentY        =   3228
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Detalle"
         TabPicture(0)   =   "frmLogBieSerMant_des.frx":06B0
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fgBSDet"
         Tab(0).Control(1)=   "cmdBS(0)"
         Tab(0).Control(2)=   "cmdBS(1)"
         Tab(0).Control(3)=   "cmdBS(2)"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Proveedores Mismo Producto"
         TabPicture(1)   =   "frmLogBieSerMant_des.frx":06CC
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fgBSProv"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Productos Similares"
         TabPicture(2)   =   "frmLogBieSerMant_des.frx":06E8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fgBSProvNS"
         Tab(2).ControlCount=   1
         Begin VB.CommandButton cmdBS 
            Caption         =   "Eliminar"
            Height          =   330
            Index           =   2
            Left            =   -65760
            TabIndex        =   11
            Top             =   1110
            Width           =   1230
         End
         Begin VB.CommandButton cmdBS 
            Caption         =   "Modificar"
            Height          =   330
            Index           =   1
            Left            =   -65760
            TabIndex        =   10
            Top             =   735
            Width           =   1230
         End
         Begin VB.CommandButton cmdBS 
            Caption         =   "Agregar"
            Height          =   330
            Index           =   0
            Left            =   -65760
            TabIndex        =   9
            Top             =   375
            Width           =   1230
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBSDet 
            Height          =   1335
            Left            =   -74925
            TabIndex        =   8
            Top             =   390
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   2355
            _Version        =   393216
            FixedCols       =   0
            BackColorBkg    =   16777215
            AllowBigSelection=   0   'False
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBSProvNS 
            Height          =   1335
            Left            =   -74925
            TabIndex        =   12
            Top             =   405
            Width           =   10380
            _ExtentX        =   18309
            _ExtentY        =   2355
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            TextStyleFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBSProv 
            Height          =   1365
            Left            =   75
            TabIndex        =   13
            Top             =   375
            Width           =   10380
            _ExtentX        =   18309
            _ExtentY        =   2408
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            TextStyleFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin MSComctlLib.TreeView tvwObjeto 
         Height          =   5010
         Left            =   60
         TabIndex        =   6
         Top             =   195
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   8837
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Detalle de Item seleccionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   5220
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "Imprimir"
      Height          =   360
      Index           =   3
      Left            =   8295
      TabIndex        =   1
      Top             =   7335
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9495
      TabIndex        =   0
      Top             =   7335
      Width           =   1200
   End
End
Attribute VB_Name = "frmLogBieSerMant_des"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL  As String
Dim rs As New ADODB.Recordset
Dim sCod As String, sEstado As String, sObj
'Dim nNiv As Integer
Dim nIndex As Integer
Dim nodX As Node
Dim llOk As Boolean
Dim raiz As String
Dim nObjNiv As Integer
Dim gaObj() As String
Dim lnColCod As Long
Dim lnColDesc As Long
Public psDatoCod As String
Public psDatoDesc As String
Public vbUltNiv As Boolean
Dim nBuscarPos As Integer
Dim sBuscarTexto As String

'------------------------Funcionalidad Copiada
Dim clsDBS As DLogBieSer
Dim parent As String
Public CodigoBS As String
Dim objPista As COMManejador.Pista 'ARLO 20170125



Public Sub Inicio(prs As ADODB.Recordset, sObjCod As String, Optional sRaiz As String = "")
sCod = prs(0)
sObj = sObjCod
raiz = sRaiz
Set rs = prs
Me.Show 1
End Sub

Private Sub BuscarDato(ByVal nPos As Integer, ByVal psBuscarTexto As String, TipoBusqueda As String)
Dim k As Integer
Select Case TipoBusqueda
Case "txt"
        For k = nPos + 1 To tvwObjeto.Nodes.Count
        If InStr(UCase(tvwObjeto.Nodes(k).Text), UCase(psBuscarTexto)) > 0 Then
            tvwObjeto.Nodes(k).Selected = True
            nBuscarPos = k
            Exit For
        End If
        Next
        If nPos = nBuscarPos Then
           MsgBox " Dato no encontrado,Intente Nuevamente  ", vbInformation, "Informacion no encontrada"
        End If
        tvwObjeto.SetFocus
Case "cod"
        For k = nPos + 1 To tvwObjeto.Nodes.Count
            If Trim(Left(tvwObjeto.Nodes(k).Text, InStr(1, tvwObjeto.Nodes(k).Text, "-") - 1)) = psBuscarTexto Then
            tvwObjeto.Nodes(k).Selected = True
            nBuscarPos = k
            Exit For
        End If
        Next
'        If nPos = nBuscarPos Then
'           MsgBox " ¡ Dato no encontrado ,Intente Nuevamente   ", vbInformation, "Informacion no encontrada"
'        End If
        tvwObjeto.SetFocus
End Select
End Sub


Private Sub cmdBS_Click(Index As Integer)
Dim rsBSDet As ADODB.Recordset
    Dim sBSCod As String, sBSCodDet As String, sBSNomDet As String
    Dim nResult As Integer
    Select Case Index
        Case 0:
            'Agregar
            sBSCod = Trim(fgBSDet.TextMatrix(fgBSDet.row, 0))
            'nResult = frmLogBieSerMantIngreso.Inicio("1", sBSCod)
              
            If Len(sBSCod) = 11 Then
                MsgBox "El Nuevo Codigo de Bien no puede ser Mayor a 11 Digitos ", vbInformation, "Seleccione un nivel Superior"
                Exit Sub
            End If
             
            
            CodigoBS = ""
            nResult = frmLogMantOpc.Inicio("1", "1", sBSCod)
            
            If nResult = 0 Then
                'Set fgBS.Recordset = clsDBS.CargaBS(BsTodosFlex)
                sBuscarTexto = CodigoBS
                Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, CodigoBS)
                    If rsBSDet.RecordCount > 0 Then
                       Set fgBSDet.Recordset = rsBSDet
                Else
                    fgBSDet.Rows = 2
                    fgBSDet.Clear
                End If
                Carga_Arbol
                BuscarDato nBuscarPos, sBuscarTexto, "cod"
                 
                 If tvwObjeto.SelectedItem.parent Is Nothing Then
                    parent = "1"
                 Else
                    parent = Trim(Left(tvwObjeto.SelectedItem.parent, InStr(1, tvwObjeto.SelectedItem.parent, "-") - 1))
                 End If
                 
                
                'parent = tvwObjeto.SelectedItem.parent
            End If
        Case 1:
            'Modificar
            sBSCod = fgBSDet.TextMatrix(fgBSDet.row, 0)
            sBSCodDet = fgBSDet.TextMatrix(fgBSDet.row, 0)
            sBSNomDet = fgBSDet.TextMatrix(fgBSDet.row, 2)
            If sBSCodDet = "" Then
                MsgBox "No existe bien/servicio", vbInformation, " Aviso"
                   Exit Sub
            End If
            'nResult = frmLogBieSerMantIngreso.Inicio("2", sBSCodDet)
            nResult = frmLogMantOpc.Inicio("1", "2", sBSCodDet)
            If nResult = 0 Then
                'Set fgBS.Recordset = clsDBS.CargaBS(BsTodosFlex)
                
                Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
                If rsBSDet.RecordCount > 0 Then
                    Set fgBSDet.Recordset = rsBSDet
                Else
                    fgBSDet.Rows = 2
                    fgBSDet.Clear
                End If
                Carga_Arbol
                BuscarDato nBuscarPos, sBSCodDet, "cod"
            End If
        Case 2:
            'Eliminar
            
            
            
            
            sBSCod = fgBSDet.TextMatrix(fgBSDet.row, 0)
            sBSCodDet = fgBSDet.TextMatrix(fgBSDet.row, 0)
            sBSNomDet = fgBSDet.TextMatrix(fgBSDet.row, 2)
            
            If sBSCodDet = "" Then
                MsgBox "No existe bien/servicio", vbInformation, " Aviso"
                Exit Sub
            End If
            If tvwObjeto.SelectedItem.Children > 0 Then
                MsgBox "No Puede Eliminar el Codigo de Bien " & sBSCod & " Este tiene codigos que dependen de este  ", vbInformation, " No se puede eliminar,Elimine descendencia primero"
                Exit Sub
            End If
            
            If MsgBox("¿ Estás seguro de eliminar " & sBSNomDet & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                nResult = clsDBS.EliminaBS(sBSCodDet)
                'ARLO 20170125
                gsOpeCod = LogPistaEntraSalidaBienes
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Elimino el Bien Cod : " & sBSCodDet & " | " & sBSNomDet
                Set objPista = Nothing
                '***********
                If nResult = 0 Then
                    'Set fgBSDet.Recordset = clsDBS.CargaBS(BsTodosFlex)
                    If tvwObjeto.SelectedItem.parent Is Nothing Then
                    parent = "1"
                    Else
                    parent = Trim(Left(tvwObjeto.SelectedItem.parent, InStr(1, tvwObjeto.SelectedItem.parent, "-") - 1))
                    End If
                    Carga_Arbol
                    
                    
                    
                    
                    If parent <> "1" Then
                        BuscarDato nBuscarPos, parent, "cod"
                    End If
                    
                    Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, parent)
                    'If Not (rsBSDet.EOF And rsBSDet.BOF) Then
                        If rsBSDet.RecordCount > 0 Then
                            Set fgBSDet.Recordset = rsBSDet
                        Else
                            fgBSDet.Rows = 2
                            fgBSDet.Clear
                        End If
                    'End If
                Else
                    MsgBox "No se terminó la operación", vbInformation, " Aviso "
                    
                End If
                    tvwObjeto.SetFocus
                
            End If
        Case 3:
            'IMPRIMIR
            Dim clsNImp As NLogImpre
            Dim clsPrevio As clsPrevio
            Dim sImpre As String
            MousePointer = 11
            Set clsNImp = New NLogImpre
            sImpre = clsNImp.ImpBS(gsNomAge, gdFecSis)
            Set clsNImp = Nothing
            
            MousePointer = 0
            Set clsPrevio = New clsPrevio
            clsPrevio.Show sImpre, Me.Caption, True, , gImpresora
            Set clsPrevio = Nothing
        
        Case Else
            MsgBox "Indice de comand de bien/servicio no reconocido", vbInformation, " Aviso "
    End Select
End Sub

Private Sub cmdBuscar_Click()
nBuscarPos = 0
If Me.tvwObjeto.Nodes.Count > 0 Then
   sBuscarTexto = InputBox("Descripción de Producto a Buscar ", "Busca de Bienes")
   BuscarDato nBuscarPos, sBuscarTexto, "txt"
End If
End Sub

Private Sub cmdBuscarSig_Click()
BuscarDato nBuscarPos, sBuscarTexto, "txt"
End Sub


Private Sub cmdSalir_Click()
 Set clsDBS = Nothing
 Unload Me
End Sub

Private Sub Form_Activate()
tvwObjeto.SetFocus
If rs.EOF And rs.BOF Then
   Unload Me
End If
End Sub
Private Sub Form_Load()
Dim rsBSDet As ADODB.Recordset
Set clsDBS = New DLogBieSer
Dim I As Integer
fgBSDet.ColWidth(0) = 0
fgBSDet.ColWidth(1) = 1650
fgBSDet.ColWidth(2) = 4400
fgBSDet.ColWidth(3) = 600
fgBSDet.ColWidth(4) = 850
fgBSDet.ColWidth(5) = 850
fgBSDet.ColWidth(6) = 850
fgBSDet.ColWidth(7) = 850
fgBSDet.ColWidth(8) = 500
For I = 0 To 3
    Me.cmdBS(I).Enabled = True
    Me.cmdBS(I).Visible = True
Next I
Carga_Arbol
If tvwObjeto.Nodes.Count > 0 Then
Else
    cmdBS(0).Enabled = False
    cmdBS(1).Enabled = False
    cmdBS(2).Enabled = False
    cmdBS(3).Enabled = False
    rsBSDet.Close
    Set rsBSDet = Nothing
    Exit Sub
End If
Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, CodigoBS)
If rsBSDet.RecordCount > 0 Then
    Set fgBSDet.Recordset = rsBSDet
Else
    fgBSDet.Rows = 2
    fgBSDet.Clear
End If
rsBSDet.Close
Set rsBSDet = Nothing
End Sub
Private Sub ExpandeObj()
Dim I As Integer
For I = 1 To tvwObjeto.Nodes.Count
    If InStr(sObj, Mid(tvwObjeto.Nodes(I).Key, 2, 21)) = 1 Then
       tvwObjeto.Nodes(I).Expanded = True
       tvwObjeto.Nodes(I).Selected = True
       nIndex = I
    End If
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set oDescObj = Nothing
End Sub

Private Sub mnuBuscarIni_Click()

End Sub


Private Sub tvwObjeto_Collapse(ByVal Node As MSComctlLib.Node)
Node.ForeColor = "&H80000008"
End Sub

Private Sub tvwObjeto_DblClick()
'If tvwObjeto.Nodes(nIndex).Children = 0 Then
'    cmdAceptar_Click
'End If
End Sub

Private Sub tvwObjeto_Expand(ByVal Node As MSComctlLib.Node)
Node.ForeColor = "&H8000000D"
End Sub

Private Sub tvwObjeto_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   If tvwObjeto.Nodes(nIndex).Children = 0 Then
'      cmdAceptar_Click
'   End If
'End If
End Sub
Private Sub tvwObjeto_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim clsProveedor As DLogProveedor
    Set clsProveedor = New DLogProveedor
    Dim rsProv As New ADODB.Recordset
    Dim rsProvNS As New ADODB.Recordset
    
    If InStr(1, Node.Text, "-") < 9 Then
        fgBSProv.Clear
        fgBSProv.Rows = 2
        fgBSProvNS.Clear
        fgBSProvNS.Rows = 2
        Exit Sub
    End If
    
    Set rsProv = clsProveedor.CargaProveedorBS(ProBSProveedor, Trim(Left(Node.Text, InStr(1, Node.Text, "-") - 2)), False)
    Set rsProvNS = clsProveedor.CargaProveedorBS(ProBSProveedor, Trim(Left(Node.Text, InStr(1, Node.Text, "-") - 2)), True)
    
    If rsProv.EOF And rsProv.BOF Then
        fgBSProv.Clear
        fgBSProv.Rows = 2
    Else
        Set Me.fgBSProv.Recordset = rsProv
        fgBSProv.ColWidth(0) = 1500
        fgBSProv.ColWidth(1) = 3500
        fgBSProv.ColWidth(2) = 4000
        fgBSProv.ColWidth(5) = 4000
    End If
    
    If rsProvNS.EOF And rsProvNS.BOF Then
        fgBSProvNS.Clear
        fgBSProvNS.Rows = 2
    Else
        Set Me.fgBSProvNS.Recordset = rsProvNS
        fgBSProvNS.ColWidth(0) = 1500
        fgBSProvNS.ColWidth(1) = 3500
        fgBSProvNS.ColWidth(2) = 4000
        fgBSProvNS.ColWidth(5) = 4000
    End If
    
    
    Dim rsBSDet As ADODB.Recordset
    Dim sBSCod As String
    Dim nPos As Integer
    nIndex = Node.Index
    nPos = InStr(1, Node, "-")
    sBSCod = Trim(Left(Node, nPos - 1))
    If Node.Key <> "K1" Then
    parent = Trim(Left(Node.parent, InStr(1, Node.parent, "-") - 1))
    Else
    parent = 1
    End If
    'sBSCod = Trim(fgBS.TextMatrix(fgBS.Row, 0))
    Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
    If rsBSDet.RecordCount > 0 Then
        Set fgBSDet.Recordset = rsBSDet
    Else
        fgBSDet.Rows = 2
        fgBSDet.Clear
    End If
End Sub

Public Property Get lOk() As Boolean
lOk = llOk
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
llOk = lOk
End Property

Private Sub GetDatosObjeto(nIndex As Integer)
Dim n As Integer
psDatoCod = Mid(tvwObjeto.Nodes(nIndex).Key, 2, Len(tvwObjeto.Nodes(nIndex).Key))
psDatoDesc = Mid(tvwObjeto.Nodes(nIndex).Text, InStr(tvwObjeto.Nodes(nIndex).Text, "-") + 2, 255)
End Sub
Public Property Get ColCod() As Long
ColCod = lnColCod
End Property
Public Property Let ColCod(ByVal vNewValue As Long)
lnColCod = vNewValue
End Property
Public Property Get ColDesc() As Long
ColDesc = lnColDesc
End Property
Public Property Let ColDesc(ByVal vNewValue As Long)
lnColDesc = vNewValue
End Property
Sub Carga_Arbol()
Dim sCod As String
Dim Sql As String
On Error GoTo ErrObj

llOk = False
tvwObjeto.Nodes.Clear
'CentraForm Me
Set tvwObjeto.ImageList = imgList
lnColDesc = 1
Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion
      'SQL = " Select RTrim(BS.cBSCod) Codigo, " _
      '& " Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', " _
      '& " len(BS.cBSCod) Nivel " _
      '& " From BienesServicios BS Inner Join Constante Co On CO.nConsValor = BS.nBSUnidad  And " _
      '& " CO.nConsCOd = '1019' Where  Substring(BS.cBSCod,2,2) In ('11','12','13') And bVigente = 1 ORDER BY RTRIM(BS.cBSCod) "
      
      Sql = " Select RTrim(BS.cBSCod) Codigo, " _
      & " Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])', " _
      & " len(BS.cBSCod) Nivel " _
      & " From BienesServicios BS Inner Join Constante Co On CO.nConsValor = BS.nBSUnidad  And " _
      & " CO.nConsCOd = '1019' Where   bVigente = 1 ORDER BY RTRIM(BS.cBSCod) "

      
      
Set rs = oCon.CargaRecordSet(Sql)
Set nodX = tvwObjeto.Nodes.Add()
rs.MoveFirst
If raiz = "" Then
   'lblObjeto.Caption = " Objeto: " & rs(lnColDesc)
   nObjNiv = rs(2)
   sCod = rs(lnColCod)
   nodX.Key = "K" & sCod
   nodX.Text = sCod & " - " & rs(lnColDesc)
   AsignaImgNodo nodX
   nodX.Tag = CStr(rs(2))
   rs.MoveNext
Else
   'lblObjeto.Caption = " Objeto: " & raiz
   sCod = Mid(rs(lnColCod), 1, 2) & "X"
   nObjNiv = rs(2) - 1
   nodX.Key = "K" & sCod
   nodX.Text = sCod & " - " & raiz
   AsignaImgNodo nodX
   nodX.Tag = "0"
End If
CargaNodo sCod, tvwObjeto, rs, nObjNiv, lnColCod, lnColDesc
nIndex = 1
tvwObjeto.Nodes(1).Expanded = True
If Len(sObj) > 0 Then
   ExpandeObj
End If
nBuscarPos = 1

Exit Sub
ErrObj:
   Err.Raise Err.Number, "frmDescObjeto-form-load", Err.Description
End Sub


