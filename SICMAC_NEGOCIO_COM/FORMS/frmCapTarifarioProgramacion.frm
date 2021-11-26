VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapTarifarioProgramacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programacion de Aplicacion de Tarifarios"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12090
   Icon            =   "frmCapTarifarioProgramacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   300
      Left            =   11025
      TabIndex        =   9
      Top             =   5355
      Width           =   1005
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   300
      Left            =   9945
      TabIndex        =   6
      Top             =   5355
      Width           =   1005
   End
   Begin VB.CommandButton btnRegistrarProgramacion 
      Caption         =   "Registrar Programación"
      Height          =   300
      Left            =   7920
      TabIndex        =   5
      Top             =   5355
      Width           =   1950
   End
   Begin VB.CommandButton btnVerProgramacion 
      Caption         =   "Ver Programacion"
      Height          =   300
      Left            =   45
      TabIndex        =   4
      Top             =   5340
      Width           =   1680
   End
   Begin TabDlg.SSTab fraProgramacion 
      Height          =   5235
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   9234
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Programación"
      TabPicture(0)   =   "frmCapTarifarioProgramacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cbTipo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnSeleccionar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "grdVersiones"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dpFecha"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin MSComCtl2.DTPicker dpFecha 
         Height          =   300
         Left            =   1530
         TabIndex        =   10
         Top             =   4850
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   123404289
         CurrentDate     =   42481
      End
      Begin SICMACT.FlexEdit grdVersiones 
         Height          =   3930
         Left            =   90
         TabIndex        =   3
         Top             =   855
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   6932
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Id-tmpSel--Producto-Sub Producto-Personeria-Grupo-Version-nLista-Fecha Vigencia-tmp"
         EncabezadosAnchos=   "0-0-500-1800-1800-1800-700-4900-0-0-0"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-7-X-9-X"
         ListaControles  =   "0-0-4-0-0-0-0-3-0-2-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-L-L-C-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Id"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton btnSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   300
         Left            =   2835
         TabIndex        =   2
         Top             =   450
         Width           =   1320
      End
      Begin VB.ComboBox cbTipo 
         Height          =   315
         Left            =   495
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   2310
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Vigencia:"
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   4890
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo:"
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   495
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCapTarifarioProgramacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************
'* NOMBRE         : frmCapTarifarioProgramacion
'* DESCRIPCION    : Proyecto - Tarifario Versionado - Programacion de los Tarifarios
'* CREACION       : RIRO, 20160420 10:00 AM
'************************************************************************************************************

Option Explicit

Private oCon As COMNCaptaGenerales.NCOMCaptaDefinicion
Private rsTarifario As ADODB.Recordset
Private rsGrupos() As ADODB.Recordset
Private sProducto() As String
Private sSubProducto() As String
Private sPersoneria() As String
Private sGrupo() As String
'Private sChkProducto() As String

Private Sub CargarControles()
    'cargando tipo
    cbTipo.AddItem "Tasas" & Space(50) & 1
    cbTipo.AddItem "Comisiones" & Space(50) & 2
    cbTipo.ListIndex = 0
    dpFecha.value = DateAdd("d", 1, gdFecSis)
End Sub
Private Sub Limpiar()
    Dim rsd As Variant
    Dim sreg As Variant
    
    'Limpiando el grid
    LimpiaFlex grdVersiones
    
    'Limpiando los arreglos
    ReDim rsGrupos(0)
    ReDim sProducto(0)
    ReDim sSubProducto(0)
    ReDim sPersoneria(0)
    ReDim sGrupo(0)
'    ReDim sChkProducto(0)
  
    cbTipo.Enabled = True
    btnSeleccionar.SetFocus
    
End Sub
Private Sub CargarGrid(ByVal rsTipo As ADODB.Recordset)
Dim i As Integer, top As Integer

top = 0

'obteniendo el total de agrupaciones
If Not rsTipo Is Nothing Then
    If Not rsTipo.EOF And Not rsTipo.BOF Then
        If rsTipo.RecordCount > 0 Then
            rsTipo.MoveLast
            top = rsTipo!nLista
            rsTipo.MoveFirst
        End If
    End If
End If
'procede en caso de existir registros
If top <> 0 Then
    For i = 1 To top
        rsTipo.Filter = "nLista = " & i
        If rsTipo.RecordCount > 0 Then
            'expandiendo los arreglos
            ReDim Preserve rsGrupos(i)
            ReDim Preserve sProducto(i)
            ReDim Preserve sSubProducto(i)
            ReDim Preserve sPersoneria(i)
            ReDim Preserve sGrupo(i)
'            ReDim Preserve sChkProducto(i)
            
            'creando recordset
            Set rsGrupos(i) = New ADODB.Recordset
            rsGrupos(i).Fields.Append "cVersion", adVarChar, 150
            rsGrupos(i).Fields.Append "nIdComision", adInteger
            rsGrupos(i).CursorLocation = adUseClient
            rsGrupos(i).CursorType = adOpenStatic
            rsGrupos(i).Open
            sProducto(i) = rsTipo!cProducto
            sSubProducto(i) = rsTipo!cSubProducto
            sPersoneria(i) = rsTipo!cPersoneria
            sGrupo(i) = rsTipo!cGrupo
            
            Do While Not rsTipo.BOF And Not rsTipo.EOF
                rsGrupos(i).AddNew
                rsGrupos(i).Fields("cVersion") = rsTipo!cVersion
                rsGrupos(i).Fields("nIdComision") = rsTipo!nIdVersion
'                If rsTipo!nProgramacion > 0 Then
'                    sChkProducto(i) = "."
'                End If
                rsTipo.MoveNext
            Loop
        End If
    Next i
End If
'Llenando el grid
rsTipo.Filter = "nLista <> 999"
i = 1
For i = i To top
    If Not rsGrupos(i) Is Nothing Then
        If Not rsGrupos(i).EOF And Not rsGrupos(i).BOF Then
            If rsGrupos(i).RecordCount > 0 Then
                grdVersiones.AdicionaFila
                grdVersiones.TextMatrix(grdVersiones.Rows - 1, 0) = i
                grdVersiones.TextMatrix(grdVersiones.Rows - 1, 1) = i
'                grdVersiones.TextMatrix(grdVersiones.Rows - 1, 2) = sChkProducto(i)
                grdVersiones.TextMatrix(grdVersiones.Rows - 1, 3) = sProducto(i)
                grdVersiones.TextMatrix(grdVersiones.Rows - 1, 4) = sSubProducto(i)
                grdVersiones.TextMatrix(grdVersiones.Rows - 1, 5) = sPersoneria(i)
                grdVersiones.TextMatrix(grdVersiones.Rows - 1, 6) = sGrupo(i)
                grdVersiones.TextMatrix(grdVersiones.Rows - 1, 8) = i
            End If
        End If
    End If
Next i
End Sub
Private Function Valida() As String
Dim sMensaje As String
Dim dfecha As Date
Dim nTipo As Integer, i As Integer, nContar As Integer
Dim rs As ADODB.Recordset

nTipo = CInt(Trim(Right(cbTipo.Text, 5)))
dfecha = dpFecha.value
Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion

'verifica si ya existe una programación en la fecha seleccionada
If oCon.ObtieneProgramacionXfechaTipo(nTipo, dfecha, rs) Then
    sMensaje = "Ya existe una programacion en la fecha: " & dpFecha.value & vbNewLine
End If

'Verifica si hay algún registro seleccionado
For i = 1 To grdVersiones.Rows - 1
    If grdVersiones.TextMatrix(i, 2) = "." And val(Trim(Right(grdVersiones.TextMatrix(i, 7), 5))) > 0 Then
        nContar = nContar + 1
    End If
Next i
If nContar = 0 Then
    sMensaje = sMensaje & "Debe seleccionar una version y proceder a guardar" & vbNewLine
End If
'verificando que la fecha seleccionada sea la mayor a las fecha actual del sistema.
If gdFecSis >= dpFecha.value Then
    sMensaje = sMensaje & "La fecha seleccionada debe ser mayor a la fecha del sistema" & vbNewLine
End If
Valida = sMensaje
Set oCon = Nothing
End Function
Private Sub btnCancelar_Click()
Limpiar
cbTipo.ListIndex = 0
End Sub
Private Sub btnRegistrarProgramacion_Click()
Dim sValida As String
Dim dFechaVigencia As Date
Dim sMovNro As String
Dim lstTasaComi() As Integer, i As Integer, j As Integer
Dim oCont As COMNContabilidad.NCOMContFunciones

sValida = Valida

If Len(Trim(sValida)) = 0 Then
    Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set oCont = New COMNContabilidad.NCOMContFunciones
    sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    j = 1
    For i = 1 To grdVersiones.Rows - 1
        If Trim(grdVersiones.TextMatrix(i, 2)) = "." Then
            ReDim Preserve lstTasaComi(j)
            lstTasaComi(j) = CInt(Trim(Right(grdVersiones.TextMatrix(i, 7), 5)))
            j = j + 1
        End If
    Next i
    If MsgBox("¿Desea registrar la programación seleccionada?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        If oCon.RegistraProgramacion(dpFecha.value, sMovNro, 1, lstTasaComi, CInt(Trim(Right(cbTipo.Text, 5)))) Then
            Limpiar
            MsgBox "La programación fue registrada correctamente", vbInformation, "Aviso"
        Else
            MsgBox "Se presentaron inconvenientes", vbInformation, "Aviso"
        End If
    End If
    Set oCont = Nothing
Else
    MsgBox "Observaciones:" & vbNewLine & sValida, vbInformation, "Aviso"
End If
End Sub
Private Sub btnSalir_Click()
    If MsgBox("Desea salir del formulario de Programacion de Tarifarios?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub btnSeleccionar_Click()
'cbTipo.Enabled = False
    Dim bResp As Boolean
    If cbTipo.ListIndex >= 0 Then
        Limpiar
        bResp = False
        Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
            If CInt(Trim(Right(cbTipo.Text, 5))) = 1 Or CInt(Trim(Right(cbTipo.Text, 5))) = 2 Then  ' Tasas o Comisiones
                Set rsTarifario = oCon.ObtenerVersionesXtipo(CInt(Trim(Right(cbTipo.Text, 5))))
            Else
        End If
        If Not rsTarifario Is Nothing Then
            If Not rsTarifario.EOF And Not rsTarifario.BOF Then
                If rsTarifario.RecordCount > 0 Then
                    CargarGrid rsTarifario
                    bResp = True
                End If
            End If
        End If
        Set oCon = Nothing
        If Not bResp Then
            MsgBox "No se encontraron registros", vbInformation, "Aviso"
            btnSeleccionar.SetFocus
        Else
            cbTipo.Enabled = False
        End If
    End If
End Sub

Private Sub btnVerProgramacion_Click()
Dim ofrmVer As New frmCapTarifarioProgramacionVer
ofrmVer.Show 1

End Sub

Private Sub dpFecha_Change()

If dpFecha.value <= gdFecSis Then
    MsgBox "La fecha ingresada debe de ser mayor a la fecha del sistema", vbInformation, "Aviso"
    dpFecha.value = DateAdd("d", 1, gdFecSis)
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'controlando el Ctrl + V
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub
Private Sub Form_Load()
CargarControles
End Sub
Private Sub grdVersiones_Click()

Dim nFil As Integer
Dim nCol As Integer
Dim nList As Integer ' lista
Dim rsTmp As New ADODB.Recordset

nFil = grdVersiones.row
nCol = grdVersiones.Col
If IsNumeric(grdVersiones.TextMatrix(nFil, 8)) Then
    nList = grdVersiones.TextMatrix(nFil, 8)
Else
    nList = -1
End If
If nList > 0 Then
    If nCol = 7 Then
        grdVersiones.CargaCombo rsGrupos(nList).Clone
        SendKeys "{Enter}"
    End If
End If
End Sub
Private Sub grdVersiones_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdVersiones.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
