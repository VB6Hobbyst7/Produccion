VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLogProveedor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Proveedores."
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   Icon            =   "frmLogProveedor.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkCargarDetallleBBSS 
      Appearance      =   0  'Flat
      Caption         =   "Cargar Detalle de Bienes / Servicios"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2610
      TabIndex        =   27
      Top             =   4035
      Width           =   3435
   End
   Begin VB.TextBox txtNroMax 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10395
      TabIndex        =   25
      Text            =   "10"
      Top             =   300
      Width           =   945
   End
   Begin VB.TextBox TxtFilas 
      Height          =   285
      Left            =   13080
      TabIndex        =   24
      Text            =   "10"
      Top             =   330
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   345
      Left            =   5760
      TabIndex        =   22
      Top             =   240
      Width           =   1155
   End
   Begin VB.TextBox TxtBuscar 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   60
      TabIndex        =   21
      Top             =   240
      Width           =   5670
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   0
      TabIndex        =   10
      Top             =   3960
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Bienes / Servicios"
      TabPicture(0)   =   "frmLogProveedor.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdProvBS(1)"
      Tab(0).Control(1)=   "cmdProvBS(0)"
      Tab(0).Control(2)=   "fgProvBS"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Agencias  "
      TabPicture(1)   =   "frmLogProveedor.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Line3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LstAgencia"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "CmdAsignar"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ChkTodos"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CheckBox ChkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   155
         TabIndex        =   19
         Top             =   385
         Width           =   200
      End
      Begin VB.CommandButton cmdProvBS 
         Caption         =   "Eliminar"
         Height          =   345
         Index           =   1
         Left            =   -64920
         TabIndex        =   14
         Top             =   1080
         Width           =   1050
      End
      Begin VB.CommandButton cmdProvBS 
         Caption         =   "Agregar"
         Height          =   345
         Index           =   0
         Left            =   -64920
         TabIndex        =   13
         Top             =   600
         Width           =   1050
      End
      Begin VB.CommandButton CmdAsignar 
         Caption         =   "&Asignar"
         Height          =   330
         Left            =   10320
         TabIndex        =   12
         Top             =   360
         Width           =   1035
      End
      Begin VB.ListBox LstAgencia 
         Appearance      =   0  'Flat
         Height          =   2055
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   600
         Width           =   9825
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgProvBS 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   3625
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
      Begin VB.Line Line3 
         X1              =   1680
         X2              =   240
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   120
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   600
         X2              =   120
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                Descripción de Agencia"
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
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   6105
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "            Codigo Agencia"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Hab/Des. Consucode "
      Height          =   345
      Index           =   6
      Left            =   8400
      TabIndex        =   9
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Hab/Des. Sunat"
      Height          =   345
      Index           =   5
      Left            =   6960
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   10320
      TabIndex        =   7
      Top             =   3600
      Width           =   1035
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Imprimir"
      Height          =   345
      Index           =   4
      Left            =   2160
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Buscar"
      Height          =   345
      Index           =   3
      Left            =   3240
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Eliminar"
      Height          =   345
      Index           =   2
      Left            =   8640
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Act./Des."
      Height          =   345
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton cmdProv 
      Caption         =   "Agregar"
      Height          =   345
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   11400
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgProv 
      Height          =   2925
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5159
      _Version        =   393216
      BackColor       =   -2147483639
      Cols            =   14
      FixedCols       =   0
      BackColorBkg    =   -2147483639
      GridColor       =   -2147483633
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin VB.Label lblNroMaximo 
      Caption         =   "Nro Maximo de Proveedores :"
      Height          =   195
      Left            =   8115
      TabIndex        =   26
      Top             =   375
      Width           =   2160
   End
   Begin VB.Label Label5 
      Caption         =   "Numero Filas"
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
      Height          =   255
      Left            =   11880
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "   Ingrese Proveedor"
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
      Height          =   255
      Left            =   15
      TabIndex        =   16
      Top             =   15
      Width           =   3270
   End
End
Attribute VB_Name = "frmLogProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDProv As DLogProveedor
Dim clsNProv As NLogProveedor
Dim clsDAgencias As DLogAgencias
Dim saveClicks(1) As Integer
Dim filaGrid As Integer
Dim bFilaGrid As Boolean
Dim sTempBusqueda As String
Dim sDescripcionAgencia As String
Dim estadoAccesoLogistica As Integer

Private Sub chkTodos_Click()
Dim i As Integer
If ChkTodos.value = 1 Then
    For i = 0 To LstAgencia.ListCount - 1
        LstAgencia.Selected(i) = True
        LstAgencia.ListIndex = 0
    Next
Else
If ChkTodos.value = 0 Then
    For i = 0 To LstAgencia.ListCount - 1
        LstAgencia.Selected(i) = False
        LstAgencia.ListIndex = 0
    Next
End If
End If
End Sub

Private Sub cmdAsignar_Click()
Dim rsAsigna As ADODB.Recordset
Dim strcPersCod, strCAgeCod As String
Dim i, Fila As Integer
Dim objCn As New DConecta
Dim sSQL As String
Dim sActualiza As String

On Error GoTo ErrAsignar



strcPersCod = fgProv.TextMatrix(fgProv.RowSel, 0)
sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)

For i = 0 To LstAgencia.ListCount - 1
If LstAgencia.Selected(i) = True Then

strCAgeCod = Trim(Left(LstAgencia.List(i), 3))

sSQL = " if not exists(select cAgeCod from proveedorAG where cAgeCod='" & strCAgeCod & "' " & _
       " and cPersCod= '" & strcPersCod & "') begin insert into proveedorAG (cPersCod,cAgeCod,cUltimaActualizacion) " & _
       " values ('" & strcPersCod & "','" & strCAgeCod & "','" & sActualiza & "') end"
       
If objCn.AbreConexion Then
objCn.Ejecutar sSQL
   
End If
End If
Next i
objCn.CierraConexion
Exit Sub
ErrAsignar:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdBuscar_Click()
TxtBuscar_KeyPress (13)
End Sub

Private Sub cmdProv_Click(Index As Integer)
    Dim rsProvBS As ADODB.Recordset
    Dim obp As UPersona
    Dim nCont As Integer, nResult As Integer
    Dim sActualiza As String
    Dim sPersCod As String
    Dim sPersNombre As String
    Dim nProvEstado As Integer
    Dim pnEstSunat As Variant
    Dim pnEstConsucode As Variant
    Dim nProvAgeReten As LogProvAgenteRetencion
    Dim nProvBuenCont As LogProvBuenContribuyente
    Dim bPersDoc As Boolean
    Dim rsProv As ADODB.Recordset

    
    If fgProv.TextMatrix(fgProv.row, 0) = "" And Index <> 0 Then Exit Sub
    
    If Index > 0 Then
        If fgProv.row = 0 Then
            MsgBox "No existen proveedores", vbInformation, " Aviso"
            Exit Sub
        End If
        sPersCod = fgProv.TextMatrix(fgProv.row, 0)
        sPersNombre = fgProv.TextMatrix(fgProv.row, 1)
        nProvEstado = Val(fgProv.TextMatrix(fgProv.row, 5))
        
    End If

    Select Case Index
        Case 0:
            'Agregar proveedores
            Set obp = frmBuscaPersona.Inicio
            If obp Is Nothing Then Exit Sub
            sPersCod = Trim(obp.sPersCod)
            For nCont = 1 To fgProv.Rows - 1
                If fgProv.TextMatrix(nCont, 0) = sPersCod Then
                    MsgBox "Proveedor " & obp.sPersNombre & " ya se encuentra registrado", vbInformation, " Aviso "
                    Exit Sub
                End If
            Next
            'bPersDoc = True
            'If Trim(obp.sPersIdnroRUC) = "" Then
                'bPersDoc = False
                'For nCont = 1 To obp.DocsPers.RecordCount
                    'If obp.DocsPers!cPersIdTpo = gPersIdRUS Then
                        'bPersDoc = True
                        'Exit For
                    'End If
                    'obp.DocsPers.MoveNext
                'Next
            'End If
            If sPersCod <> "" Then
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                nResult = clsDProv.GrabaProveedor(sPersCod, sActualiza, gsCodAge)       '**** REFERENCIA A CLASE DLOGPROVEEDOR:GrabaProveedor, modificado   ****'
                If nResult = 0 Then
                    '--Set fgProv.Recordset = clsDProv.CargaProveedores(estadoAccesoLogistica, gsCodAge)
                    'sPersCod = fgProv.TextMatrix(fgProv.row, 0)
                    
                    'Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
                    'If rsProvBS.RecordCount > 0 Then
                    '    Set fgProvBS.Recordset = rsProvBS
                    'Else
                    '    fgProvBS.Rows = 2
                    '    fgProvBS.Clear
                    'End If
                    MsgBox "Agregación Realizada con exito en.  " & sDescripcionAgencia, vbInformation, "Aviso"
                Else
                    Call ErrorDesc(nResult)
                End If
            Else
                MsgBox "No se hizo la agregación", vbInformation, "Aviso"
            End If
        Case 1:
            'Activar/Desactivar
            nResult = clsNProv.ActDesProveedor(sPersCod, nProvEstado)   '**** REFERENCIA A CLASE DLOGPROVEEDOR:ActDesProveedor ****'
            If nResult = 0 Then
                'Set fgProv.Recordset = clsDProv.CargaProveedor
                
                If fgProv.TextMatrix(fgProv.RowSel, 6) = "Activado" Then
                fgProv.TextMatrix(fgProv.RowSel, 6) = "Desactivado"
                Else
                fgProv.TextMatrix(fgProv.RowSel, 6) = "Activado"
                End If
                
            Else
                Call ErrorDesc(nResult)
            End If
        Case 2:
            'Eliminar
            
            'Validar que no halla realizado ningun proceso
            If MsgBox("¿ Estás seguro de eliminar a " & sPersNombre & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                nResult = clsDProv.EliminaProveedor(sPersCod)
                If nResult = 0 Then
                    Set rsProv = New ADODB.Recordset
                    'Set rsProv = clsDProv.CargaProveedores(estadoAccesoLogistica, gsCodAge)

                    If rsProv.EOF And rsProv.BOF Then
                        fgProv.Rows = 1
                        fgProv.Rows = 2
                        fgProv.FixedRows = 1
                    Else
                        'Set fgProv.Recordset = clsDProv.CargaProveedores(estadoAccesoLogistica, gsCodAge)
                    End If
                    
                    sPersCod = fgProv.TextMatrix(fgProv.row, 0)
                    
                    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
                    If rsProvBS.RecordCount > 0 Then
                        Set fgProvBS.Recordset = rsProvBS
                    Else
                        fgProvBS.Clear
                        fgProvBS.Rows = 2
                    End If
                Else
                    Call ErrorDesc(nResult)
                End If
            End If
        Case 3:
            'BUSCAR
            'sPersNombre = Trim(frmLogMantOpc.Inicio(0, 1))
            'If sPersNombre <> "" Then
            '    For nCont = 1 To fgProv.Rows - 1
            '        If Left(fgProv.TextMatrix(nCont, 1), Len(sPersNombre)) = sPersNombre Then
            '            'fgProv.Row = nCont
            '            fgProv.TopRow = nCont
            '            Exit Sub
            '        End If
            '    Next
            '    MsgBox "Nombre no se encuentra en la lista", vbInformation, " Aviso"
            'End If
            Dim sProvBuscar As String
            
            
            sProvBuscar = InputBox("Ingrese proveedor :", "")
            For filaGrid = 1 To fgProv.Rows - 1
                If InStr(1, fgProv.TextMatrix(filaGrid, 1), Trim(UCase(sProvBuscar))) Then
                   bFilaGrid = True
                   fgProv_Scroll
                   ColorDeFila (filaGrid)
                   Exit For
                End If
            Next filaGrid
            
        Case 4:
            'IMPRIMIR
            Dim clsNImp As NLogImpre
            Dim clsPrevio As clsPrevio
            Dim sImpre As String
            MousePointer = 11
            Set clsNImp = New NLogImpre
            'sImpre = clsNImp.ImpProveedor(gsNomAge, gdFecSis, sTempBusqueda, gsCodAge) 'modifcado en los parametros sTempBusqueda
            Set clsNImp = Nothing
            MousePointer = 0
            Set clsPrevio = New clsPrevio
            clsPrevio.Show sImpre, Me.Caption, True
            Set clsPrevio = Nothing
            
        Case 5:
            'HAB/DES. SUNAT
            
            pnEstSunat = IIf(fgProv.TextMatrix(fgProv.RowSel, 11) = "Habilitado", 1, 0)
            
            nResult = clsNProv.HabDesabilitaSunat(sPersCod, pnEstSunat) '**** REFERENCIA A CLASE NLOGPROVEEDOR:HabDesabilitaSunat, creado ****'
            If nResult = 0 Then
                'Set fgProv.Recordset = clsDProv.CargaProveedor
                If fgProv.TextMatrix(fgProv.RowSel, 11) = "Habilitado" Then
                    fgProv.TextMatrix(fgProv.RowSel, 11) = "Inhabilitado"
                Else
                    fgProv.TextMatrix(fgProv.RowSel, 11) = "Habilitado"
                End If
            Else
                Call ErrorDesc(nResult)
            End If
            
        Case 6:
            'HAB/DES. CONSUCODE
            
            pnEstConsucode = IIf(fgProv.TextMatrix(fgProv.RowSel, 12) = "Habilitado", 1, 0)
            
            nResult = clsNProv.HabDesabilitaConsucode(sPersCod, pnEstConsucode) '**** REFERENCIA A CLASE NLOGPROVEEDOR:HabDesabilitaConsucode, creado ****'
            If nResult = 0 Then
                'Set fgProv.Recordset = clsDProv.CargaProveedor
                If fgProv.TextMatrix(fgProv.RowSel, 12) = "Habilitado" Then
                    fgProv.TextMatrix(fgProv.RowSel, 12) = "Inhabilitado"
                Else
                    fgProv.TextMatrix(fgProv.RowSel, 12) = "Habilitado"
                End If
            Else
                Call ErrorDesc(nResult)
            End If
            
        Case Else
            MsgBox "Indice de comand de proveedores no reconocido", vbInformation, " Aviso "
    End Select
    
    'MDISicmact.staMain.Panels(2).Text = fgProv.Rows & " registros."
End Sub




Private Sub cmdProvBS_Click(Index As Integer)
    Dim clsDBS As DLogBieSer
    Dim rsProvBS As ADODB.Recordset
    Dim nResult As Integer
    Dim sActualiza As String
    Dim sPersCod As String
    Dim sBSCod As String
    Dim sBSNombre As String
    
    If fgProv.TextMatrix(fgProv.row, 0) = "" Then Exit Sub

    
    If fgProv.row = 0 Then
        MsgBox "No existen proveedores", vbInformation, " Aviso"
        Exit Sub
    End If
    sPersCod = fgProv.TextMatrix(fgProv.row, 0)
    
    If Index > 0 Then
        If fgProvBS.row = 0 Then
            MsgBox "No existen bienes/servicios del proveedor", vbInformation, " Aviso"
            Exit Sub
        End If
        sBSCod = fgProvBS.TextMatrix(fgProvBS.row, 0)
        sBSNombre = fgProvBS.TextMatrix(fgProvBS.row, 2)
    End If
    
    Select Case Index
        Case Is = 0
            'Agregar bien/servicio de proveedor
            Dim vBS As ClassDescObjeto
            Set vBS = New ClassDescObjeto
            Set clsDBS = New DLogBieSer
            vBS.ColCod = 0
            vBS.ColDesc = 1
            vBS.Show clsDBS.CargaBS(BsTodosArbol), ""
            
            Set clsDBS = Nothing
            If vBS.lbOk Then
                sBSCod = vBS.gsSelecCod
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                
                nResult = clsDProv.GrabaProveedorBS(sPersCod, sBSCod, sActualiza)
                If nResult = 0 Then
                    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
                    If rsProvBS.RecordCount > 0 Then
                        Set fgProvBS.Recordset = rsProvBS
                    Else
                        fgProvBS.Rows = 2
                        fgProvBS.Clear
                    End If
                Else
                    Call ErrorDesc(nResult)
                End If
            Else
                MsgBox "No se hizo la agregación", vbInformation, "Aviso"
            End If
            Set vBS = Nothing
        Case Is = 1
            'Eliminar
            If sBSCod = "" Then Exit Sub

            If MsgBox("¿ Estás seguro de eliminar " & sBSNombre & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                nResult = clsDProv.EliminaProveedorBS(sPersCod, sBSCod)
                If nResult = 0 Then
                    Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
                    If rsProvBS.RecordCount > 0 Then
                        Set fgProvBS.Recordset = rsProvBS
                    Else
                        fgProvBS.Rows = 2
                        fgProvBS.Clear
                    End If
                Else
                    Call ErrorDesc(nResult)
                End If
            End If
        Case Else
            MsgBox "Indice de comand bien/servicio de proveedor no reconocido", vbInformation, " Aviso "
    End Select
End Sub

Private Sub cmdSalir_Click()
    Set clsDProv = Nothing
    Set clsNProv = Nothing
    Unload Me
End Sub


Private Sub fgProv_Click()
    Dim rsProvBS As ADODB.Recordset
    Dim sPersCod As String
    Dim nTempCol As Integer
    Dim rsAgenProv As ADODB.Recordset
    Dim filaAgencia As Integer
    Dim itemsSelect As Integer
    
    With fgProv
    If .Col <> 13 Then
        .ForeColorSel = &H80000009
    End If
    
    End With
    
    With fgProv
        If .Col = 13 Then
            .BackColorSel = &H80000009
        Else
            fgProv.BackColorSel = &H8000000D
        End If
    End With
    
    nTempCol = fgProv.Col
    ColorDeFila (fgProv.row)
    
    With fgProv
    .Col = nTempCol
        If .Col = 13 Then
            editarComentario
        Else
            If Me.chkCargarDetallleBBSS.value = 1 Then
               sPersCod = fgProv.TextMatrix(fgProv.row, 0)
               Me.MousePointer = 11
                
               Set rsProvBS = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
               If rsProvBS.RecordCount > 0 Then
                  Set fgProvBS.Recordset = rsProvBS
               Else
                  fgProvBS.Rows = 2
                  fgProvBS.Clear
               End If
               Me.MousePointer = 0
            End If
            
            'Verificar asignaciones
            For itemsSelect = 0 To LstAgencia.ListCount - 1
                LstAgencia.Selected(itemsSelect) = False
            Next itemsSelect
            
            Set rsAgenProv = clsDAgencias.DevuelveProveedorAG(fgProv.TextMatrix(fgProv.row, 0))    '**** REFERENCIA A CLASE DLOGPROVEEDOR:DevuelveProveedorAG, creado ****'
            Do While Not rsAgenProv.EOF = True
                For filaAgencia = 0 To LstAgencia.ListCount - 1
                    If Trim(Left(LstAgencia.List(filaAgencia), 3)) = rsAgenProv!cAgeCod Then
                        LstAgencia.Selected(filaAgencia) = True
                        LstAgencia.ListIndex = filaAgencia
                        Exit For
                    Else
                        LstAgencia.ListIndex = -1
                    End If
                Next
                rsAgenProv.MoveNext
            Loop
        End If
    End With
    
End Sub

Sub ColorDeFila(ByVal filaApintar As Integer)
Dim i As Integer

If saveClicks(1) <> 0 Then
quitarColorDeFila
End If
For i = 1 To fgProv.Cols - 2
    With fgProv
        .Col = i
        .row = filaApintar
        .CellBackColor = &H8000000D
        .CellForeColor = &H80000009
    End With
Next i
fgProv.Col = 0
saveClicks(1) = filaApintar
End Sub
Sub quitarColorDeFila()
Dim i As Integer

For i = 1 To fgProv.Cols - 2
    With fgProv
        .Col = i
        .row = saveClicks(1)
        .CellBackColor = &H80000009
        .CellForeColor = &H80000008
    End With
Next i
fgProv.Col = 0
saveClicks(1) = 0
End Sub

Sub editarComentario()
Text1 = fgProv
Text1.Visible = True
Text1.SetFocus

Text1.Left = fgProv.CellLeft + fgProv.Left
Text1.Top = fgProv.CellTop + fgProv.Top
Text1.Width = fgProv.CellWidth
Text1.Height = fgProv.CellHeight

End Sub

Private Sub fgProv_LeaveCell()
 
If Text1.Visible Then
    fgProv = Text1
    Text1.Visible = False
End If
End Sub

Private Sub fgProv_RowColChange()
    fgProv_Click
End Sub

Private Sub LstAgencia_Click()
LstAgencia.ListIndex = LstAgencia.ListIndex
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim sComentario As String
Dim sPersCod As String

If KeyAscii = 13 Then
    fgProv_LeaveCell
    
    With fgProv
    .ForeColorSel = &H80000008
    End With
    
    sPersCod = fgProv.TextMatrix(fgProv.RowSel, 0)
    sComentario = fgProv.TextMatrix(fgProv.RowSel, 13)
            
    Call clsNProv.IngresaQuitaComentario(sPersCod, sComentario) '**** REFERENCIA A CLASE NLOGPROVEEDOR:IngresaQuitaCometario, creado ****'
End If
End Sub
Private Sub fgProv_Scroll()
fgProv_LeaveCell
If bFilaGrid = True Then
fgProv.TopRow = filaGrid
fgProv.ScrollTrack = True
bFilaGrid = False
End If
End Sub



Private Sub Form_Load()
Dim rsProv  As ADODB.Recordset, rsProvBS  As ADODB.Recordset, rsAgencias As ADODB.Recordset
Dim sPersCod As String

Call CentraForm(Me)

Set clsDProv = New DLogProveedor
Set clsNProv = New NLogProveedor
Set clsDAgencias = New DLogAgencias

Set rsAgencias = clsDAgencias.CargaAgencias(gsCodAge)   '**** REFERENCIA A CLASE DLogAgencias:CargaAgencias, creado ****'
If rsAgencias.RecordCount > 0 Then
    sDescripcionAgencia = rsAgencias!cAgeDescripcion
Else
    sDescripcionAgencia = "AGENCIA NO IDENTIFICADA"
End If


If gbPermisoLogProveedorAG = True Then
    estadoAccesoLogistica = 0
    SSTab1.TabVisible(1) = False
    frmLogProveedor.Caption = "Mantenimiento de Proveedores.  " & sDescripcionAgencia
    AnchoColumna
    cmdProv(1).Visible = False
    cmdProv(3).Visible = False
    cmdProv(4).Move 1080
Else
    
    AnchoColumna
    estadoAccesoLogistica = 1
    SSTab1.TabVisible(1) = True
    cmdProv(1).Visible = True
    SSTab1.Tab = 1
    frmLogProveedor.Caption = "Mantenimiento de Proveedores.  " & sDescripcionAgencia
    CargaAgencias
End If
End Sub
Private Sub cabeceraGrid()
fgProv.TextMatrix(0, 0) = ""
fgProv.TextMatrix(0, 1) = "Nombre"
fgProv.TextMatrix(0, 2) = "Dirección"
fgProv.TextMatrix(0, 3) = "RUC"
fgProv.TextMatrix(0, 4) = "TELEFONO"
fgProv.TextMatrix(0, 5) = ""
fgProv.TextMatrix(0, 6) = "Estado"
fgProv.TextMatrix(0, 7) = "Cta_MN"
fgProv.TextMatrix(0, 8) = "Cta_ME"
fgProv.TextMatrix(0, 9) = ""
fgProv.TextMatrix(0, 10) = "A.Ret/B.Cont"
fgProv.TextMatrix(0, 11) = ""
fgProv.TextMatrix(0, 12) = "Habil SUNAT"
fgProv.TextMatrix(0, 13) = "Habil Consucode"
fgProv.TextMatrix(0, 14) = "COMENTARIO"
End Sub
Private Sub AnchoColumna()
fgProv.Cols = 14
fgProv.ColWidth(0) = 0
fgProv.ColWidth(1) = 3000
fgProv.ColWidth(2) = 3500
fgProv.ColWidth(3) = 1000
fgProv.ColWidth(4) = 1000
fgProv.ColWidth(5) = 0
fgProv.ColWidth(6) = 1000
fgProv.ColWidth(7) = 1000
fgProv.ColWidth(8) = 1000
fgProv.ColWidth(9) = 0
fgProv.ColWidth(10) = 1500
fgProv.ColWidth(13) = 3000
fgProvBS.ColWidth(2) = 7000

End Sub
Private Sub ErrorDesc(ByVal pnError As Integer)
    Select Case pnError
        Case 1
            MsgBox "Error al establecer la conexión", vbInformation, " Aviso "
        Case 2
            MsgBox "Registro duplicado", vbInformation, " Aviso "
    End Select
End Sub





Private Sub CargaAgencias()
Dim R As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
On Error GoTo ErrAgencia
sSQL = "select AG.cAgeCod,AG.cAgeDescripcion "
sSQL = sSQL & " from agencias AG"
sSQL = sSQL & " Order By AG.cAgeCod"
Set oConecta = New DConecta
oConecta.AbreConexion
Set R = oConecta.CargaRecordSet(sSQL)
oConecta.CierraConexion
Set oConecta = Nothing
LstAgencia.Clear
    Do While Not R.EOF
        LstAgencia.AddItem R!cAgeCod & Space(100) & R!cAgeDescripcion
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub
ErrAgencia:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
Dim rsBuscaProv As ADODB.Recordset

    KeyAscii = CaracteresFuncionales(KeyAscii)

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Len(Trim(TxtBuscar.Text)) = 0 Then
        MsgBox "Falta Ingresar el Nombre del Proveedor", vbInformation, "Aviso"
        Exit Sub
      End If
        fgProvBS.Rows = 2
        fgProvBS.Clear
        Me.MousePointer = 11
        sTempBusqueda = TxtBuscar.Text
        saveClicks(1) = 0
        '**** REFERENCIA A CLASE DLOGPROVEEDOR:CargaProveedores, creado ****'
        'estadoAccesoLogistica,gsCodAge :en gVarPublicas global estadoAccesoLogistica as integer
        Set rsBuscaProv = clsDProv.CargaProveedores(estadoAccesoLogistica, gsCodAge, TxtBuscar.Text, CInt(TxtFilas.Text))
      
      If rsBuscaProv.RecordCount = 0 Then
           MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
           fgProv.Rows = 2
           fgProv.Clear
      Else
      
         If rsBuscaProv.RecordCount > Me.txtNroMax.Text Then
            MsgBox "Sea mas específico en su Busqueda """ _
             & "Por ejemplo si usted desea buscar VELASQUEZ/PEREZ,JUAN escriba lo sgte:   " & vbCrLf & vbCrLf & "" _
             & "       VELA%PEREZ" & vbCrLf & "" _
             & "       VELA%JUAN/" & vbCrLf & "" _
             & "       PEREZ%JUAN" & vbCrLf & "" _
             & "       PEREZ,JUAN" & vbCrLf & "", vbInformation, "Aviso"
             fgProv.Rows = 2
             fgProv.Clear
         Else
            Set fgProv.Recordset = rsBuscaProv
            With fgProv
            .Col = 1
            fgProv.Sort = 7
            End With
            
            'TxtBuscar.Text = Trim(rsBuscaProv(1))
            fgProv.SetFocus
         End If
      End If
      
   Else
        KeyAscii = Letras(KeyAscii)
        
   End If
Me.MousePointer = 0
End Sub

Private Sub TxtFilas_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub TxtFilas_LostFocus()
If Len(TxtFilas) = 0 Then
TxtFilas = 10
End If
End Sub

Private Sub txtNroMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdBuscar.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
    
End Sub

Private Sub txtNroMax_LostFocus()
    If Not IsNumeric(Me.txtNroMax.Text) Then
        Me.txtNroMax.Text = "10"
    End If
End Sub
