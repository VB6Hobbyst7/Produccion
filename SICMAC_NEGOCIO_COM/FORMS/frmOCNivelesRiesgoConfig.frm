VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOCNivelesRiesgoConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Niveles de Riesgo"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16800
   Icon            =   "frmOCNivelesRiesgoConfig.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   16800
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16620
      _ExtentX        =   29316
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   5292
      TabCaption(0)   =   "Configuración General"
      TabPicture(0)   =   "frmOCNivelesRiesgoConfig.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraVariables"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraSubVariables"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraItems"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame fraItems 
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6615
         Left            =   10440
         TabIndex        =   7
         Top             =   480
         Width           =   6015
         Begin VB.CommandButton cmdCancelarItem 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   420
            Left            =   2280
            TabIndex        =   19
            Top             =   6000
            Width           =   975
         End
         Begin VB.CommandButton cmdGrabarItem 
            Caption         =   "Grabar"
            Enabled         =   0   'False
            Height          =   420
            Left            =   1200
            TabIndex        =   18
            Top             =   6000
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarItem 
            Caption         =   "Editar"
            Enabled         =   0   'False
            Height          =   420
            Left            =   120
            TabIndex        =   17
            Top             =   6000
            Width           =   975
         End
         Begin SICMACT.FlexEdit feItem 
            Height          =   5055
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   8916
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Items-Peso-nVarSubDetCod-nVarSubCod"
            EncabezadosAnchos=   "350-4500-550-0-0"
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
            ColumnasAEditar =   "X-X-2-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C-C"
            FormatosEdit    =   "0-0-3-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label lblItems 
            Appearance      =   0  'Flat
            Caption         =   "lblItems"
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
            Height          =   375
            Left            =   1680
            TabIndex        =   9
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label3 
            Caption         =   "Sub Variable:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraSubVariables 
         Caption         =   "Sub Variables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6615
         Left            =   4680
         TabIndex        =   3
         Top             =   480
         Width           =   5655
         Begin VB.CommandButton cmdCancelarSubVar 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   420
            Left            =   2400
            TabIndex        =   16
            Top             =   6000
            Width           =   975
         End
         Begin VB.CommandButton cmdGrabarSubVar 
            Caption         =   "Grabar"
            Enabled         =   0   'False
            Height          =   420
            Left            =   1320
            TabIndex        =   15
            Top             =   6000
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarSubVar 
            Caption         =   "Editar"
            Enabled         =   0   'False
            Height          =   420
            Left            =   240
            TabIndex        =   14
            Top             =   6000
            Width           =   975
         End
         Begin SICMACT.FlexEdit feSubVariable 
            Height          =   5055
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   8916
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Sub Variables-Peso-nVarSubCod-nVarCod"
            EncabezadosAnchos=   "350-4000-550-0-0"
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
            ColumnasAEditar =   "X-X-2-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C-C"
            FormatosEdit    =   "0-0-3-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label lblSubVariable 
            Appearance      =   0  'Flat
            Caption         =   "lblSubVariable"
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
            Height          =   375
            Left            =   2160
            TabIndex        =   5
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Variable Seleccionada:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame fraVariables 
         Caption         =   "Variables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6615
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4335
         Begin VB.CommandButton cmdCancelarVar 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   420
            Left            =   2400
            TabIndex        =   13
            Top             =   6000
            Width           =   975
         End
         Begin VB.CommandButton cmdGrabarVar 
            Caption         =   "Grabar"
            Enabled         =   0   'False
            Height          =   420
            Left            =   1320
            TabIndex        =   12
            Top             =   6000
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarVar 
            Caption         =   "Editar"
            Height          =   420
            Left            =   240
            TabIndex        =   11
            Top             =   6000
            Width           =   975
         End
         Begin SICMACT.FlexEdit feVariable 
            Height          =   5535
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   9763
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Variables-Peso-nVarCod"
            EncabezadosAnchos=   "350-2500-550-0"
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
            ColumnasAEditar =   "X-X-2-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C"
            FormatosEdit    =   "0-0-3-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
End
Attribute VB_Name = "frmOCNivelesRiesgoConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************
'*** REQUERIMIENTO: TI-ERS106-2014
'*** USUARIO: FRHU
'*** FECHA CREACION: 17/09/2014
'*********************************
Option Explicit
Private bHabilitarGridVariable As Boolean
Private bHabilitarGridSubVariable As Boolean
Private bHabilitarGridItem As Boolean
Private Sub Form_Load()
    lblSubVariable.Caption = ""
    lblItems.Caption = ""
    Call CargarGridVariables
End Sub
'***********************
'*** VARIABLES
Private Function ValidarGrabarVar() As Boolean
    Dim fila As Integer, totalfila As Integer, valor As Integer, suma As Integer
    
    suma = 0
    totalfila = feVariable.Rows - 1
    For fila = 1 To totalfila
        valor = CInt(val(feVariable.TextMatrix(fila, 2)))
        suma = suma + valor
    Next fila
    If suma <> 100 Then
        MsgBox "La sumatoria de todos los pesos debe ser igual a 100"
        ValidarGrabarVar = False
        Exit Function
    End If
    ValidarGrabarVar = True
End Function
Private Sub cmdGrabarVar_Click()
    Dim nVarCod As Integer, fila As Integer, nVarValor As Integer
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    
    If Not ValidarGrabarVar Then Exit Sub
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    For fila = 1 To feVariable.Rows - 1
        nVarValor = feVariable.TextMatrix(fila, 2)
        nVarCod = feVariable.TextMatrix(fila, 3)
        Call oServ.ActualizarValorNivelesDeRiesgos(1, nVarCod, 0, 0, nVarValor)
    Next
    Set oServ = Nothing
    Call cmdCancelarVar_Click
End Sub
Private Sub cmdEditarVar_Click()
    cmdEditarVar.Enabled = False
    cmdGrabarVar.Enabled = True
    cmdCancelarVar.Enabled = True
    feVariable.lbEditarFlex = True
    bHabilitarGridVariable = True
    fraSubVariables.Enabled = False
    fraItems.Enabled = False
End Sub
Private Sub cmdCancelarVar_Click()
    cmdEditarVar.Enabled = True
    feVariable.lbEditarFlex = False
    cmdGrabarVar.Enabled = False
    cmdCancelarVar.Enabled = False
    bHabilitarGridVariable = False 'Puede seleccionar y se cargara en el siguiente grid
    fraSubVariables.Enabled = True
    fraItems.Enabled = True
    Call CargarGridVariables
End Sub
Private Sub feVariable_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If CInt(val(feVariable.TextMatrix(pnRow, pnCol))) > 100 Then
        MsgBox "Solo es posible Ingresar Valores numericos enteros menores a 100"
        Cancel = False
        Exit Sub
    End If
    Cancel = True
End Sub
Private Sub feVariable_Click()
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rs As New ADODB.Recordset
    Dim fila As Integer, nVarCod As Integer
    
    If bHabilitarGridVariable Then Exit Sub
    nVarCod = feVariable.TextMatrix(feVariable.row, 3)
    lblSubVariable.Caption = UCase(feVariable.TextMatrix(feVariable.row, 1))
    Call FormateaFlex(feItem)
    lblItems.Caption = ""
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rs = oServ.obtenerNivelesRiesgoVariables(2, nVarCod)
    Set oServ = Nothing
    
    fila = 0
    Call FormateaFlex(feSubVariable)
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            feSubVariable.AdicionaFila
            fila = fila + 1
            feSubVariable.TextMatrix(fila, 1) = UCase(rs!cVarSubDescripcion)
            feSubVariable.TextMatrix(fila, 2) = rs!nVarSubValor
            feSubVariable.TextMatrix(fila, 3) = rs!nVarSubCod
            feSubVariable.TextMatrix(fila, 4) = rs!nVarCod
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    fraItems.Enabled = False
End Sub
Private Sub CargarGridVariables()
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rs As New ADODB.Recordset
    Dim fila As Integer
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rs = oServ.obtenerNivelesRiesgoVariables(1)
    Set oServ = Nothing
    
    fila = 0
    Call FormateaFlex(feVariable)
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            feVariable.AdicionaFila
            fila = fila + 1
            feVariable.TextMatrix(fila, 1) = UCase(rs!cVarDescripcion)
            feVariable.TextMatrix(fila, 2) = rs!nVarValor
            feVariable.TextMatrix(fila, 3) = rs!nVarCod
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Sub
'**************************
'*** SUBVARIABLES
Private Function ValidarGrabarSubVar() As Boolean
    Dim fila As Integer, totalfila As Integer, valor As Integer, suma As Integer
    
    suma = 0
    totalfila = feSubVariable.Rows - 1
    For fila = 1 To totalfila
        valor = CInt(feSubVariable.TextMatrix(fila, 2))
        suma = suma + valor
    Next fila
    If suma <> 100 Then
        MsgBox "La sumatoria de todos los pesos debe ser igual a 100"
        ValidarGrabarSubVar = False
        Exit Function
    End If
    ValidarGrabarSubVar = True
End Function
Private Sub cmdGrabarSubVar_Click()
    Dim nVarCod As Integer, nVarSubCod As Integer, fila As Integer, nVarValor As Integer
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    
    nVarCod = feSubVariable.TextMatrix(1, 4)
    
    If Not ValidarGrabarSubVar Then Exit Sub
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    For fila = 1 To feSubVariable.Rows - 1
        nVarValor = feSubVariable.TextMatrix(fila, 2)
        nVarSubCod = feSubVariable.TextMatrix(fila, 3)
        Call oServ.ActualizarValorNivelesDeRiesgos(2, nVarCod, nVarSubCod, 0, nVarValor)
    Next
    Set oServ = Nothing
    
    Call cmdCancelarSubVar_Click
End Sub
Private Sub cmdEditarSubVar_Click()
    cmdEditarSubVar.Enabled = False
    cmdGrabarSubVar.Enabled = True
    cmdCancelarSubVar.Enabled = True
    feSubVariable.lbEditarFlex = True
    bHabilitarGridSubVariable = True
    fraVariables.Enabled = False
    fraItems.Enabled = False
End Sub
Private Sub cmdCancelarSubVar_Click()
    cmdEditarSubVar.Enabled = True
    cmdGrabarSubVar.Enabled = False
    cmdCancelarSubVar.Enabled = False
    feSubVariable.lbEditarFlex = False
    bHabilitarGridSubVariable = False
    fraVariables.Enabled = True
    fraItems.Enabled = True
    Call CargarGridSubVariablesDespuesDeGrabar
End Sub
Private Sub feSubVariable_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If CInt(val(feSubVariable.TextMatrix(pnRow, pnCol))) > 100 Then
        MsgBox "Solo es posible Ingresar Valores numericos enteros menores a 100"
        Cancel = False
        Exit Sub
    End If
    Cancel = True
End Sub
Private Sub CargarGridSubVariablesDespuesDeGrabar()
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rs As New ADODB.Recordset
    Dim fila As Integer, nVarCod As Integer
    
    nVarCod = feSubVariable.TextMatrix(1, 4)
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rs = oServ.obtenerNivelesRiesgoVariables(2, nVarCod)
    Set oServ = Nothing
    
    fila = 0
    Call FormateaFlex(feSubVariable)
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            feSubVariable.AdicionaFila
            fila = fila + 1
            feSubVariable.TextMatrix(fila, 1) = UCase(rs!cVarSubDescripcion)
            feSubVariable.TextMatrix(fila, 2) = rs!nVarSubValor
            feSubVariable.TextMatrix(fila, 3) = rs!nVarSubCod
            feSubVariable.TextMatrix(fila, 4) = rs!nVarCod
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Sub
Private Sub feSubVariable_Click()
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rs As New ADODB.Recordset
    Dim fila As Integer, nVarSubCod As Integer
    
    If bHabilitarGridSubVariable Then Exit Sub
    If feSubVariable.TextMatrix(feSubVariable.row, 3) = "" Then
        Exit Sub
    Else
        cmdEditarSubVar.Enabled = True
        cmdEditarItem.Enabled = True
    End If
    
    nVarSubCod = feSubVariable.TextMatrix(feSubVariable.row, 3)
    lblItems.Caption = UCase(feSubVariable.TextMatrix(feSubVariable.row, 1))
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rs = oServ.obtenerNivelesRiesgoVariables(3, nVarSubCod)
    Set oServ = Nothing
    
    fila = 0
    Call FormateaFlex(feItem)
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            feItem.AdicionaFila
            fila = fila + 1
            feItem.TextMatrix(fila, 1) = UCase(rs!cVarSubDetDescripcion)
            feItem.TextMatrix(fila, 2) = rs!nVarSubDetValor
            feItem.TextMatrix(fila, 3) = rs!nVarSubDetCod
            feItem.TextMatrix(fila, 4) = rs!nVarSubCod
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    fraItems.Enabled = True
End Sub
'******************
'*** ITEMS
Private Sub cmdGrabarItem_Click()
    Dim nVarSubDetCod As Long, nVarSubCod As Integer, fila As Integer, nVarValor As Integer
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    
    nVarSubCod = feItem.TextMatrix(1, 4)
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    For fila = 1 To feItem.Rows - 1
        nVarValor = feItem.TextMatrix(fila, 2)
        nVarSubDetCod = feItem.TextMatrix(fila, 3)
        Call oServ.ActualizarValorNivelesDeRiesgos(3, 0, nVarSubCod, nVarSubDetCod, nVarValor)
    Next
    Set oServ = Nothing
    
    Call cmdCancelarItem_Click
End Sub
Private Sub cmdEditarItem_Click()
    cmdEditarItem.Enabled = False
    cmdGrabarItem.Enabled = True
    cmdCancelarItem.Enabled = True
    feItem.lbEditarFlex = True
    bHabilitarGridItem = True
    fraVariables.Enabled = False
    fraSubVariables.Enabled = False
End Sub
Private Sub cmdCancelarItem_Click()
    cmdEditarItem.Enabled = True
    cmdGrabarItem.Enabled = False
    cmdCancelarItem.Enabled = False
    feItem.lbEditarFlex = False
    bHabilitarGridItem = False
    fraVariables.Enabled = True
    fraSubVariables.Enabled = True
    Call CargarGridItemDespuesDeGrabar
End Sub
Private Sub CargarGridItemDespuesDeGrabar()
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rs As New ADODB.Recordset
    Dim fila As Integer, nVarSubCod As Integer
    
    nVarSubCod = feItem.TextMatrix(1, 4)
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rs = oServ.obtenerNivelesRiesgoVariables(3, nVarSubCod)
    Set oServ = Nothing
    
    fila = 0
    Call FormateaFlex(feItem)
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            feItem.AdicionaFila
            fila = fila + 1
            feItem.TextMatrix(fila, 1) = UCase(rs!cVarSubDetDescripcion)
            feItem.TextMatrix(fila, 2) = rs!nVarSubDetValor
            feItem.TextMatrix(fila, 3) = rs!nVarSubDetCod
            feItem.TextMatrix(fila, 4) = rs!nVarSubCod
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Sub
Private Sub feItem_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If CInt(val(feItem.TextMatrix(pnRow, pnCol))) > 100 Then
        MsgBox "Solo es posible Ingresar Valores numericos enteros menores a 100"
        Cancel = False
        Exit Sub
    End If
    Cancel = True
End Sub
