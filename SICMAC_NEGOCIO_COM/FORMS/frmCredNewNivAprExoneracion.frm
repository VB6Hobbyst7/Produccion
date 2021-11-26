VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredNewNivAprExoneracion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizaciones"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   Icon            =   "frmCredNewNivAprExoneracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabNiveles 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6376
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Autorizaciones"
      TabPicture(0)   =   "frmCredNewNivAprExoneracion.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "cmdGrabarNiveles"
      Tab(0).Control(2)=   "feNiveles"
      Tab(0).Control(3)=   "cmdEliminarNiveles"
      Tab(0).Control(4)=   "cmdNuevoNiveles"
      Tab(0).Control(5)=   "cmdCerrarNiveles"
      Tab(0).Control(6)=   "cboTipoExonera"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Tipo de Autorizaciones"
      TabPicture(1)   =   "frmCredNewNivAprExoneracion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "feTiposExonera"
      Tab(1).Control(1)=   "cmdEliminarTipos"
      Tab(1).Control(2)=   "cmdNuevoTipos"
      Tab(1).Control(3)=   "cmdGrabarTipos"
      Tab(1).Control(4)=   "cmdCerrarTipos"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Not. Contempladas"
      TabPicture(2)   =   "frmCredNewNivAprExoneracion.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "feNotConte"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdNewNotConte"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdEliminarNotConte"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdCerrarNotConte"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdGrabarNotConte"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdGrabarNotConte 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5880
         TabIndex        =   17
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdCerrarNotConte 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7080
         TabIndex        =   16
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdEliminarNotConte 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   15
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdNewNotConte 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdCerrarTipos 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -67920
         TabIndex        =   12
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabarTipos 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         TabIndex        =   11
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdNuevoTipos 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74880
         TabIndex        =   10
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdEliminarTipos 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73710
         TabIndex        =   9
         Top             =   3120
         Width           =   1170
      End
      Begin VB.ComboBox cboTipoExonera 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCredNewNivAprExoneracion.frx":035E
         Left            =   -73080
         List            =   "frmCredNewNivAprExoneracion.frx":0368
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   6255
      End
      Begin VB.CommandButton cmdCerrarNiveles 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -67920
         TabIndex        =   3
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdNuevoNiveles 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74880
         TabIndex        =   2
         Top             =   3120
         Width           =   1170
      End
      Begin VB.CommandButton cmdEliminarNiveles 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73710
         TabIndex        =   1
         Top             =   3120
         Width           =   1170
      End
      Begin SICMACT.FlexEdit feNiveles 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   6
         Top             =   840
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   3836
         Cols0           =   6
         FixedCols       =   2
         HighLight       =   1
         EncabezadosNombres=   "-cNivExoneraCod-Nivel de Aprobacion-Desde-Hasta-Aux"
         EncabezadosAnchos=   "300-0-4480-1500-1500-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-4-X"
         ListaControles  =   "0-0-3-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-R-C"
         FormatosEdit    =   "0-1-0-2-2-0"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin VB.CommandButton cmdGrabarNiveles 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -68040
         TabIndex        =   7
         Top             =   2640
         Width           =   1170
      End
      Begin SICMACT.FlexEdit feTiposExonera 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   4471
         Cols0           =   5
         FixedCols       =   2
         HighLight       =   1
         EncabezadosNombres=   "-cExoneraCod-Autorizaciones-Tipo-Aux"
         EncabezadosAnchos=   "300-0-5900-1550-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-X"
         ListaControles  =   "0-0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-C"
         FormatosEdit    =   "0-1-0-0-0"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit feNotConte 
         Height          =   2535
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   4471
         Cols0           =   6
         FixedCols       =   2
         HighLight       =   1
         EncabezadosNombres=   "-cNivExoneraCod-Nivel de Aprobacion-Desde-Hasta-Aux"
         EncabezadosAnchos=   "300-0-4480-1500-1500-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-4-X"
         ListaControles  =   "0-0-3-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-R-C"
         FormatosEdit    =   "0-1-0-2-2-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Autorización :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   540
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCredNewNivAprExoneracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredNewNivAprExoneracion
'** Descripción : Formulario para la administración de los Tipos y Niveles de Exoneracion creado segun RFC110-2012
'** Creación : JUEZ, 20121205 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim oGen As COMDConstSistema.DCOMGeneral
Dim oDNiv As COMDCredito.DCOMNivelAprobacion
Dim oNNiv As COMNCredito.NCOMNivelAprobacion
Dim rs As ADODB.Recordset
Dim fbNuevoTE As Boolean, fbEditaTE As Boolean, fbNuevoNE As Boolean, fbEditaNE As Boolean
Dim fnRowSelectTE As Integer, fnRowSelectNE As Integer

'JOEP20210906 ERS041-2021
Dim nTpRegistro As Integer
Dim nTpConsultar As Integer
Public Enum TpOpciones
    RegistroAut = 1
    ConsultaAut = 2
End Enum
'JOEP20210906 ERS041-2021

Public Sub TiposExoneracion()
    'Me.Caption = "Tipos de Exoneraciones/Autorización" 'RECO20160526 ERS0022016
    CargaComboFlex
    
    'JOEP20210906 ERS041-2021
    'ListaTiposExoneracion 'Coemento 'JOEP20210906 ERS041-2021
    ListaTiposExoneracion (RegistroAut)
    nTpRegistro = RegistroAut
    nTpConsultar = 0
    'JOEP20210906 ERS041-2021
    
    'SSTabTiposExonera.Visible = True 'RECO20160610 ERS002-2016
    'SSTabNiveles.Visible = False 'RECO20160610 ERS002-2016
    SSTabNiveles.TabVisible(0) = False 'RECO20160610 ERS002-2016
    SSTabNiveles.TabVisible(2) = False 'JOEP20210904
    cmdCerrarNiveles.Cancel = False
    cmdCerrarTipos.Cancel = True
    fbEditaTE = False
    fbNuevoTE = False
    fnRowSelectTE = 0
    Me.Show 1
End Sub

Public Sub NivelesExoneracion()
    'Me.Caption = "Niveles de Exoneración" 'RECO20160526 ERS0022016
     
    'JOEP20210906 ERS041-2021
    SSTabNiveles.Tab = 0
    'ListaComboTiposExoneracion 'Comento 'JOEP20210906 ERS041-2021
    ListaComboTiposExoneracion (ConsultaAut)
    Call CargaGrillaAutoCont("TIP0024") 'JOEP20210906
    nTpRegistro = 0
    nTpConsultar = ConsultaAut
    'JOEP20210906 ERS041-2021
    
    'SSTabTiposExonera.Visible = False 'RECO20160610 ERS002-2016
    'SSTabNiveles.Visible = True 'RECO20160610 ERS002-2016
    SSTabNiveles.TabVisible(1) = False 'RECO20160610 ERS002-2016
    cmdCerrarNiveles.Cancel = True
    cmdCerrarTipos.Cancel = False
    fbNuevoNE = False
    fbEditaNE = False
    Me.cmdNuevoNiveles.Enabled = True
    fnRowSelectNE = 0
    Me.Show 1
End Sub

Private Sub CargaComboFlex()
    Set oGen = New COMDConstSistema.DCOMGeneral
    Set rs = oGen.GetConstante(7070)
    Set oGen = Nothing
    feTiposExonera.CargaCombo rs
End Sub

Private Sub ListaTiposExoneracion(ByVal nTpOpcion As Integer) 'JOEP20210906 ERS041-2021
'Private Sub ListaTiposExoneracion()'Comento JOEP20210906 ERS041-2021
    Dim lnFila As Integer
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    'Set rs = oDNiv.RecuperaTiposExoneraciones()'Comento 'JOEP20210906 ERS041-2021
    Set rs = oDNiv.RecuperaTiposExoneraciones(nTpOpcion) 'JOEP20210906 ERS041-2021
    Set oDNiv = Nothing
    Call LimpiaFlex(feTiposExonera)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feTiposExonera.AdicionaFila
            lnFila = feTiposExonera.row
            feTiposExonera.TextMatrix(lnFila, 1) = rs!cExoneraCod
            feTiposExonera.TextMatrix(lnFila, 2) = rs!cExoneraDesc
            feTiposExonera.TextMatrix(lnFila, 3) = rs!cTipoExoneraDesc & Space(25) & rs!nTipoExoneraCod
            rs.MoveNext
        Loop
        feTiposExonera.TopRow = 1
        feTiposExonera.row = 1
    Else
        'cmdNuevoTipos.Enabled = False
        'cmdEliminarTipos.Enabled = False
    End If
End Sub

Private Sub cmdCerrarTipos_Click()
    Unload Me
End Sub

Private Sub cmdEliminarTipos_Click()
    If feTiposExonera.TextMatrix(feTiposExonera.row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feTiposExonera.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Set oDNiv = New COMDCredito.DCOMNivelAprobacion
            Call oDNiv.dEliminaTipoExoneracion(feTiposExonera.TextMatrix(feTiposExonera.row, 1))
            feTiposExonera.EliminaFila feTiposExonera.row
            fbNuevoTE = False
            fbEditaTE = False
        End If
    End If
End Sub

Private Sub cmdGrabarTipos_Click()
    If ValidaDatosGrid(feTiposExonera, "Debe ingresar al menos un tipo de exoneración", "Faltan datos en la lista", 4) Then
        Dim rsDatos As ADODB.Recordset
        Set rsDatos = IIf(feTiposExonera.rows - 1 > 0, feTiposExonera.GetRsNew(), Nothing)
        If MsgBox("¿Está seguro de registrar los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set oNNiv = New COMNCredito.NCOMNivelAprobacion
        If feTiposExonera.TextMatrix(feTiposExonera.row, 1) = "xxx" Then
            Call oNNiv.dInsertaTiposExoneraciones(feTiposExonera.TextMatrix(feTiposExonera.row, 2), CInt(Right(Trim(feTiposExonera.TextMatrix(feTiposExonera.row, 3)), 2)))
        Else
            'ALPA 20141025***************************************
            If feTiposExonera.TextMatrix(feTiposExonera.row, 1) = "TIP0009" Then
                MsgBox "No se puede eliminar el tipo de exoneración (TASA)", vbExclamation
                Exit Sub
            End If
            '****************************************************
            Call oNNiv.dActualizaTiposExoneraciones(feTiposExonera.TextMatrix(feTiposExonera.row, 1), feTiposExonera.TextMatrix(feTiposExonera.row, 2), CInt(Right(Trim(feTiposExonera.TextMatrix(feTiposExonera.row, 3)), 2)))
        End If
        Set oNNiv = Nothing
        MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
        feTiposExonera.lbEditarFlex = False
        fbNuevoTE = False
        fbEditaTE = False
        fnRowSelectTE = 0
    End If
End Sub

Private Sub cmdNuevoTipos_Click()
    'ListaTiposExoneracion 'Comento JOEP20210906 ERS041-2021
    ListaTiposExoneracion (RegistroAut) 'JOEP20210906 ERS041-2021
    feTiposExonera.AdicionaFila
    CargaComboFlex
    fbEditaTE = False
    fbNuevoTE = True
    feTiposExonera.SetFocus
    SendKeys "{Enter}"
End Sub

Private Sub feTiposExonera_DblClick()
    EditarGrid
    If fbNuevoTE Then
        If feTiposExonera.col = 2 Then
            feTiposExonera.TextMatrix(feTiposExonera.row, 1) = "xxx"
        End If
    End If
End Sub

Private Sub feTiposExonera_GotFocus()
    feTiposExonera.col = 2
End Sub

Private Sub feTiposExonera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EditarGrid
        If fbNuevoTE Then
            If feTiposExonera.col = 2 Then
                feTiposExonera.TextMatrix(feTiposExonera.row, 1) = "xxx"
            End If
        End If
    End If
End Sub

Private Sub EditarGrid()
    If fbNuevoTE = False Then
        If feTiposExonera.TextMatrix(feTiposExonera.row, 0) <> "" Then
            If fbEditaTE = False Then
                If MsgBox("¿Desea editar los datos de la fila " & feTiposExonera.row & "? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                    feTiposExonera.lbEditarFlex = False
                Else
                    feTiposExonera.lbEditarFlex = True
                    fbEditaTE = True
                    feTiposExonera.TopRow = feTiposExonera.row
                    fnRowSelectTE = feTiposExonera.row
                End If
            Else
                If fnRowSelectTE <> feTiposExonera.row Then
                    feTiposExonera.lbEditarFlex = False
                    fbEditaTE = False
                End If
            End If
        End If
    Else
        If feTiposExonera.row = feTiposExonera.rows - 1 Then
            feTiposExonera.lbEditarFlex = True
            CargaComboFlex
        Else
            feTiposExonera.lbEditarFlex = False
        End If
    End If
End Sub

Private Sub feTiposExonera_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 2 Then
        feTiposExonera.TextMatrix(pnRow, pnCol) = UCase(feTiposExonera.TextMatrix(pnRow, pnCol))
    End If
    If pnCol = 3 Then
        If fbNuevoTE Then
            If pnRow = feTiposExonera.rows - 1 Then
                cmdGrabarTipos_Click
                feTiposExonera.TopRow = pnRow
            End If
        Else
            cmdGrabarTipos_Click
            feTiposExonera.TopRow = pnRow
        End If
        cmdNuevoTipos.SetFocus
    End If
End Sub

'Private Sub ListaComboTiposExoneracion()'Comento JOEP20210906 ERS041-2021
Private Sub ListaComboTiposExoneracion(ByVal nTpOpcion As Integer) 'JOEP20210906 ERS041-2021
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    'Set rs = oDNiv.RecuperaTiposExoneraciones()'Comento JOEP20210906 ERS041-2021
    Set rs = oDNiv.RecuperaTiposExoneraciones(nTpOpcion) 'JOEP20210906 ERS041-2021
    Set oDNiv = Nothing
    cboTipoExonera.Clear
    While Not rs.EOF
        cboTipoExonera.AddItem rs!cExoneraDesc & Space(500) & rs!cExoneraCod
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub

Private Sub CargaComboFlexNiveles()
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaNivAprListaParam(2)
    Set oDNiv = Nothing
    feNiveles.CargaCombo rs
    Set rs = Nothing
End Sub

Private Sub cboTipoExonera_Click()
    Dim lnFila As Integer
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaNivelesExoneracion(Trim(Right(cboTipoExonera.Text, 10)))
    Set oDNiv = Nothing
    Call LimpiaFlex(feNiveles)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feNiveles.AdicionaFila
            lnFila = feNiveles.row
            feNiveles.TextMatrix(lnFila, 1) = rs!cNivExoneraCod
            feNiveles.TextMatrix(lnFila, 2) = rs!cNivAprDesc & Space(100) & rs!cNivAprCod
            feNiveles.TextMatrix(lnFila, 3) = Format(rs!nNivExoneraDesde, "#,##0.00")
            feNiveles.TextMatrix(lnFila, 4) = Format(rs!nNivExoneraHasta, "#,##0.00")
            rs.MoveNext
        Loop
        feNiveles.TopRow = 1
        rs.MoveLast
    End If
End Sub

Private Sub cmdNuevoNiveles_Click()

'JOEP20210906 ERS041-2021
If Validaciones(ConsultaAut) = False Then
    Exit Sub
End If
'JOEP20210906 ERS041-2021

'If feNiveles.TextMatrix(feNiveles.Row, 2) <> "" And feNiveles.TextMatrix(feNiveles.Row, 3) <> "" And feNiveles.TextMatrix(feNiveles.Row, 4) <> "" Then
    CargaComboFlexNiveles
    feNiveles.AdicionaFila
    CargaComboFlex
    fbEditaNE = False
    fbNuevoNE = True
    cmdNuevoNiveles.Enabled = False
    If feNiveles.TextMatrix(feNiveles.rows - 2, 4) = "Hasta" Then
        feNiveles.TextMatrix(feNiveles.rows - 1, 3) = Format(0, "#,##0.00")
    Else
        feNiveles.TextMatrix(feNiveles.rows - 1, 3) = Format(CDbl(feNiveles.TextMatrix(feNiveles.rows - 2, 4)) + 0.01, "#,##0.00")
    End If
    
    feNiveles.SetFocus
    SendKeys "{Enter}"
'End If
End Sub

Private Sub cmdEliminarNiveles_Click()
'JOEP20210906 ERS041-2021
If Validaciones(ConsultaAut) = False Then
    Exit Sub
End If
'JOEP20210906 ERS041-2021

    If feNiveles.TextMatrix(feNiveles.row, 0) <> "" Then
        'ALPA 20141025***************************************
        If Trim(Right(cboTipoExonera.Text, 10)) = "TIP0009" Then
            MsgBox "No se puede eliminar el tipo de exoneración (TASA)", vbExclamation
            Exit Sub
        End If
        '****************************************************
        If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feNiveles.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Set oDNiv = New COMDCredito.DCOMNivelAprobacion
            Call oDNiv.dEliminaNivelesExoneracion(Trim(Right(cboTipoExonera.Text, 10)), feNiveles.TextMatrix(feNiveles.row, 1))
            feNiveles.EliminaFila feNiveles.row
            fbNuevoNE = False
            fbEditaNE = False
            Me.cmdNuevoNiveles.Enabled = True
            fnRowSelectNE = 0
        End If
    End If
End Sub

Private Sub feNiveles_DblClick()
    EditarGridNiveles
    If fbNuevoNE Then
        If feNiveles.col = 2 Then
            feNiveles.TextMatrix(feNiveles.row, 1) = "xxx"
        End If
    End If
End Sub

Private Sub feNiveles_GotFocus()
    feNiveles.col = 2
End Sub

Private Sub feNiveles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EditarGridNiveles
        If fbNuevoNE Then
            If feNiveles.col = 2 Then
                feNiveles.TextMatrix(feNiveles.row, 1) = "xxx"
            End If
        End If
    End If
End Sub

Private Sub EditarGridNiveles()
    If fbNuevoNE = False Then
        If feNiveles.TextMatrix(feNiveles.row, 0) <> "" Then
            If feNiveles.col <> 3 Then
                If fbEditaNE = False Then
                    If MsgBox("¿Desea editar los datos de la fila " & feNiveles.row & "? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                        feNiveles.lbEditarFlex = False
                    Else
                        feNiveles.lbEditarFlex = True
                        CargaComboFlexNiveles
                        fbEditaNE = True
                        feNiveles.TopRow = feNiveles.row
                        fnRowSelectNE = feNiveles.row
                    End If
                Else
                    If fnRowSelectNE <> feNiveles.row Then
                        feNiveles.lbEditarFlex = False
                        fbEditaNE = False
                    End If
                End If
            End If
        End If
    Else
        If feNiveles.row = feNiveles.rows - 1 Then
            feNiveles.lbEditarFlex = True
            CargaComboFlexNiveles
        Else
            feNiveles.lbEditarFlex = False
        End If
    End If
End Sub

Private Sub feNiveles_OnCellChange(pnRow As Long, pnCol As Long)
    Dim i As Integer
'    If pnCol = 2 Then
'        feNiveles.TextMatrix(pnRow, pnCol) = UCase(feNiveles.TextMatrix(pnRow, pnCol))
'    End If
    If pnCol = 2 Then
        For i = 1 To feNiveles.rows - 1
            If feNiveles.TextMatrix(i, 0) <> "" Then
                If i <> pnRow Then
                    If Trim(Right(feNiveles.TextMatrix(i, 2), 10)) = Trim(Right(feNiveles.TextMatrix(pnRow, 2), 10)) Then
                        MsgBox "El nivel ya fue ingresado", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            End If
        Next i
    End If
    If pnCol = 4 Then
        If feNiveles.TextMatrix(pnRow, 4) <> "" Then
            If fbEditaNE Then
                If pnRow < feNiveles.rows - 1 Then
                    feNiveles.TextMatrix(pnRow + 1, 3) = CDbl(feNiveles.TextMatrix(pnRow, 4)) + 0.01
                End If
            End If
            If fbNuevoNE Then
                If pnRow = feNiveles.rows - 1 Then
                    cmdGrabarNiveles_Click
                    feNiveles.TopRow = pnRow
                End If
            Else
                cmdGrabarNiveles_Click
                feNiveles.TopRow = pnRow
            End If
            cmdGrabarNiveles.SetFocus
        End If
    End If
End Sub

Private Sub cmdGrabarNiveles_Click()
    If ValidaDatosGrid(feNiveles, "Debe ingresar al menos un nivel de exoneración", "Faltan datos en la lista", 4) Then
        If Trim(cboTipoExonera.Text) = "" Then
            MsgBox "Debe elegir un tipo de exoneración", vbInformation, "Aviso"
            cboTipoExonera.SetFocus
            Exit Sub
        End If
        If ValidaMontosDesdeHasta Then
            Dim rsDatos As ADODB.Recordset
            Set rsDatos = IIf(feNiveles.rows - 1 > 0, feNiveles.GetRsNew(), Nothing)
            If MsgBox("¿Está seguro de registrar los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            Set oNNiv = New COMNCredito.NCOMNivelAprobacion
                Call oNNiv.dInsertaNivelesExoneraciones(Trim(Right(cboTipoExonera.Text, 10)), rsDatos)
            Set oNNiv = Nothing
            MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            feNiveles.lbEditarFlex = False
            fbNuevoNE = False
            fbEditaNE = False
            cmdNuevoNiveles.Enabled = True
            fnRowSelectNE = 0
        End If
    End If
End Sub

Private Sub cmdCerrarNiveles_Click()
    Unload Me
End Sub

Private Function ValidaMontosDesdeHasta() As Boolean
    Dim i As Integer
    ValidaMontosDesdeHasta = False
    
    For i = 1 To feNiveles.rows - 1
        If i <> feNiveles.row Then
            If feNiveles.TextMatrix(i, 0) <> "" Then
                If Trim(Right(feNiveles.TextMatrix(i, 2), 10)) = Trim(Right(feNiveles.TextMatrix(feNiveles.row, 2), 10)) Then
                    MsgBox "El nivel ya fue ingresado", vbInformation, "Aviso"
                    ValidaMontosDesdeHasta = False
                    Exit Function
                End If
            End If
        End If
    Next i
    
    For i = 1 To feNiveles.rows - 1
        'FRHU 20160818 Observación
        If Not IsNumeric(feNiveles.TextMatrix(i, 4)) Then
            MsgBox "El monto Desde debe ser un número valido en la fila " & i, vbInformation, "Aviso"
            ValidaMontosDesdeHasta = False
            Exit Function
        End If
        'FIN FRHU 20160818
        If CDbl(feNiveles.TextMatrix(i, 4)) < CDbl(feNiveles.TextMatrix(i, 3)) Then
            MsgBox "El monto Desde debe ser menor al monto Hasta en la fila " & i, vbInformation, "Aviso"
            ValidaMontosDesdeHasta = False
            Exit Function
        End If
    Next i
    
    ValidaMontosDesdeHasta = True
End Function

'JOEP20210906 ERS041-2021
Private Sub CargaGrillaAutoCont(ByVal TipoExonera As String)
    Dim lnFila As Integer
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaNivelesExoneracion(TipoExonera)
    Set oDNiv = Nothing
    Call LimpiaFlex(feNotConte)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feNotConte.AdicionaFila
            lnFila = feNotConte.row
            feNotConte.TextMatrix(lnFila, 1) = rs!cNivExoneraCod
            feNotConte.TextMatrix(lnFila, 2) = rs!cNivAprDesc & Space(100) & rs!cNivAprCod
            feNotConte.TextMatrix(lnFila, 3) = Format(rs!nNivExoneraDesde, "#,##0.00")
            feNotConte.TextMatrix(lnFila, 4) = Format(rs!nNivExoneraHasta, "#,##0.00")
            rs.MoveNext
        Loop
        feNotConte.TopRow = 1
        rs.MoveLast
    End If
End Sub

Private Function Validaciones(ByVal nTpOp As Integer, Optional ByVal nTab As Integer = 0) As Boolean
Dim i As Integer
Dim J As Integer

Validaciones = True

    If nTpConsultar = nTpOp Then
        
        If nTab = 1 Then
            If Trim(Right(cboTipoExonera.Text, 10)) = "" Then
                MsgBox "Seleccione el Tipo de Autorizacion", vbInformation, "Aviso"
                Validaciones = False
                Exit Function
            End If
        End If
    'Tab Not. Contemplada
        If nTab = 2 Then
            For i = 1 To feNotConte.rows - 1
                If feNotConte.TextMatrix(i, 2) = "" Or feNotConte.TextMatrix(i, 3) = "" Or feNotConte.TextMatrix(i, 4) = "" Then
                    MsgBox "Falta ingresar datos en la fila, " & feNotConte.TextMatrix(i, 0), vbInformation, "Aviso"
                    Validaciones = False
                    Exit Function
                End If
            Next i
        'rangos de montos
            For i = 1 To feNotConte.rows - 1
                If CDbl(feNotConte.TextMatrix(i, 4)) < CDbl(feNotConte.TextMatrix(i, 3)) Then
                    MsgBox "El Monto de la columna [Hasta] debe ser mayor a Columna [Desde], fila, " & feNotConte.TextMatrix(i, 0), vbInformation, "Aviso"
                    Validaciones = False
                    Exit Function
                End If
            Next i
            For i = 1 To feNotConte.rows - 2
                If CDbl(feNotConte.TextMatrix(i, 4)) <> CDbl(feNotConte.TextMatrix(i + 1, 3)) - 0.01 Then
                    MsgBox "La distribucion de montos son incorrectos, fila, " & feNotConte.TextMatrix(i + 1, 0), vbInformation, "Aviso"
                    Validaciones = False
                    Exit Function
                End If
            Next i
        'rangos de montos
        'No se repita Niveles
            For i = 1 To feNotConte.rows - 1
                For J = 1 To feNotConte.rows - 1
                    If Right(feNotConte.TextMatrix(i, 2), 10) = Right(feNotConte.TextMatrix(J, 2), 10) And i <> J Then
                        MsgBox "Nivele de aprobacion duplicados, por favor de verificar.", vbInformation, "Aviso"
                        Validaciones = False
                        Exit Function
                    End If
                Next J
            Next i
        'No se repita Niveles
        End If
    'Tab Not. Contemplada
    End If
End Function

Private Sub CargaAutConNiveles()
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaNivAprListaParam(2)
    Set oDNiv = Nothing
    feNotConte.CargaCombo rs
    Set rs = Nothing
End Sub

Private Sub cmdGrabarNotConte_Click()
    If Validaciones(nTpConsultar, 2) = False Then
        Exit Sub
    End If

    Dim rsDatos As ADODB.Recordset
    Set rsDatos = IIf(feNotConte.rows - 1 > 0, feNotConte.GetRsNew(), Nothing)
    If MsgBox("¿Está seguro de registrar los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set oNNiv = New COMNCredito.NCOMNivelAprobacion
    Call oNNiv.dInsertaNivelesExoneraciones("TIP0024", rsDatos)
    Set oNNiv = Nothing
    MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
    Call CargaGrillaAutoCont("TIP0024")
End Sub

Private Sub feNotConte_OnCellChange(pnRow As Long, pnCol As Long)
    If feNotConte.row < (feNotConte.rows - 1) Then
        If feNotConte.TextMatrix(feNotConte.row + 1, 3) <> "" Then
            feNotConte.TextMatrix(feNotConte.row + 1, 3) = Format(feNotConte.TextMatrix(feNotConte.row, 4) + 0.01, "#,#0.00")
        End If
    End If
    
    If feNotConte.TextMatrix(feNotConte.row, 1) = "" Then
        feNotConte.TextMatrix(feNotConte.row, 1) = Right(feNotConte.TextMatrix(feNotConte.row, 2), 10)
    End If
End Sub

Private Sub feNotConte_DblClick()
    If feNotConte.col = 2 Then
        CargaAutConNiveles
        feNotConte.lbEditarFlex = True
    End If
End Sub

Private Sub cmdEliminarNotConte_Click()
    If feNotConte.TextMatrix(feNotConte.row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feNotConte.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Set oDNiv = New COMDCredito.DCOMNivelAprobacion
            Call oDNiv.dEliminaNivelesExoneracion("TIP0024", feNotConte.TextMatrix(feNotConte.row, 1))
            feNotConte.EliminaFila feNotConte.row
            Me.cmdNewNotConte.Enabled = True
        End If
    End If
End Sub

Private Sub cmdNewNotConte_Click()
    If feNotConte.TextMatrix(feNotConte.row, 1) <> "" Then
        If Validaciones(nTpConsultar, 2) = False Then
            Exit Sub
        End If
    End If
    CargaAutConNiveles
    feNotConte.AdicionaFila
    cmdNewNotConte.Enabled = True

    If feNotConte.TextMatrix(feNotConte.rows - 2, 4) = "Hasta" Then
        feNotConte.TextMatrix(feNotConte.rows - 1, 3) = Format(0, "#,##0.00")
    Else
        feNotConte.TextMatrix(feNotConte.rows - 1, 3) = Format(CDbl(feNotConte.TextMatrix(feNotConte.rows - 2, 4)) + 0.01, "#,##0.00")
    End If
    
    feNotConte.SetFocus
    SendKeys "{Enter}"
End Sub

Private Sub cmdCerrarNotConte_Click()
    Unload Me
End Sub
'JOEP20210906 ERS041-2021
