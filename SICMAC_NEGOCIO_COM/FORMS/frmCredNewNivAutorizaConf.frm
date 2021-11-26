VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredNewNivAutorizaConf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Autorizaciones de Créditos"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frmCredNewNivAutorizaConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "Guardar datos"
      Top             =   4510
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   5880
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   4510
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4365
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   7699
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Autorizaciones"
      TabPicture(0)   =   "frmCredNewNivAutorizaConf.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feAutorizacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdConfigurar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.CommandButton cmdConfigurar 
         Caption         =   "&Configurar..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Configurar"
         Top             =   3915
         Width           =   1120
      End
      Begin SICMACT.FlexEdit feAutorizacion 
         Height          =   3435
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   6059
         Cols0           =   5
         HighLight       =   2
         EncabezadosNombres=   "N°-Codigo-Autorización-Habilitado-Configura"
         EncabezadosAnchos=   "400-0-4800-1000-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-4-0"
         EncabezadosAlineacion=   "C-C-L-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         CantEntero      =   12
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmCredNewNivAutorizaConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmCredNewNivAutorizaConf
'** Descripción : Formulario para configurar autorizaciones de créditos según ERS002-2016
'** Creación : EJVG, 20160204 04:49:00 PM
'****************************************************************************************
Option Explicit

Dim fvEscalonaConf As TEscalonamientoConf

Dim fnModificaHabilita As Integer
Dim fnModificaEscalonaConfig As Integer

Private Sub Form_Load()
    fnModificaHabilita = 0
    fnModificaEscalonaConfig = 0
    
    cargarControles
End Sub
Private Sub cargarControles()
    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim bEscalona As Boolean
    
    On Error GoTo ErrCargaControles
    Screen.MousePointer = 11
    
    FormateaFlex feAutorizacion
    Set rs = oDNiv.CargarTiposAutorizacionxConfig()
    Do While Not rs.EOF
        feAutorizacion.AdicionaFila
        i = feAutorizacion.row
        
        feAutorizacion.TextMatrix(i, 1) = rs!cAutorizaID
        feAutorizacion.TextMatrix(i, 2) = rs!cAutorizaDesc
        feAutorizacion.TextMatrix(i, 3) = IIf(rs!nEstado = 1, "1", "")
        feAutorizacion.TextMatrix(i, 4) = IIf(rs!bConfigura, 1, 0)
        
        If rs!cAutorizaID = "TIP0002" Then
            bEscalona = True
        End If
        
        rs.MoveNext
    Loop
    
    fvEscalonaConf.nMoraPorcentaje = 0#
    fvEscalonaConf.nMontoCuoMayorA = 0#
    fvEscalonaConf.nMontoCuoMenorIgual = 0#
    fvEscalonaConf.nMontoCreMayorA = 0#
    fvEscalonaConf.nMontoCreMenorIgual = 0#
    
    If bEscalona Then 'Escalonamiento
        Set rs = oDNiv.CargarAutorizacionEscalonaConfig()
        If (Not rs.EOF) Then
            fvEscalonaConf.nMoraPorcentaje = rs!nMoraPorcentaje
            fvEscalonaConf.nMontoCuoMayorA = rs!nMontoCuoMayorA
            fvEscalonaConf.nMontoCuoMenorIgual = rs!nMontoCuoMenorIgualA
            fvEscalonaConf.nMontoCreMayorA = rs!nMontoCreMayorA
            fvEscalonaConf.nMontoCreMenorIgual = rs!nMontoCreMenorIgualA
            rs.MoveNext
        End If
    End If
    
    RSClose rs
    Set oDNiv = Nothing
    
    Screen.MousePointer = 0
    Exit Sub
ErrCargaControles:
    MsgBox Err.Description, vbCritical, "Aviso"
    Screen.MousePointer = 0
End Sub
Private Sub feAutorizacion_RowColChange()
    If feAutorizacion.TextMatrix(feAutorizacion.row, 4) = "1" Then 'nConfigura
        cmdConfigurar.Enabled = True
    Else
        cmdConfigurar.Enabled = False
    End If
End Sub
Private Sub feAutorizacion_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    
    Editar = Split(feAutorizacion.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Or pnCol = 1 Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End Sub
Private Sub feAutorizacion_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    fnModificaHabilita = fnModificaHabilita + 1
    HabilitarGuardar
    
    If feAutorizacion.TextMatrix(pnRow, pnCol) = "" Then
        MsgBox "I   M   P   O   R   T   A   N   T   E   !  !  !" & Chr(13) & Chr(13) & "Al deshabilitar la autorización [" & feAutorizacion.TextMatrix(pnRow, 2) & "]," & Chr(13) & "el Sistema ya no podrá realizar la verificación si los créditos necesitan está autorización.", vbInformation, "Aviso"
    End If
End Sub
Private Sub HabilitarGuardar()
    cmdGuardar.Enabled = False
    If fnModificaHabilita > 0 Then
        cmdGuardar.Enabled = True
    End If
    If fnModificaEscalonaConfig > 0 Then
        cmdGuardar.Enabled = True
    End If
End Sub
Private Sub cmdGuardar_Click()
    Dim oNNiv As COMNCredito.NCOMNivelAprobacion
    Dim bExito As Boolean
    
    On Error GoTo ErrGuardar
    cmdGuardar.Enabled = False
        
    If MsgBox("¿Está seguro de guardar los cambios realizados?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        cmdGuardar.Enabled = True
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    cmdGuardar.Enabled = False
    Set oNNiv = New COMNCredito.NCOMNivelAprobacion
    
    bExito = oNNiv.GuardarAutorizacionConfig(GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), feAutorizacion.GetRsNew, fvEscalonaConf)
    
    Set oNNiv = Nothing
    Screen.MousePointer = 0
    cmdGuardar.Enabled = True
    
    If bExito Then
        MsgBox "Se ha guardado satisfactoriamente los datos.", vbInformation, "Aviso"
    Else
        MsgBox "Ha sucedido un error al, si el problema persiste comuniquese con el Dpto. de TI.", vbCritical, "Aviso"
    End If
    Exit Sub
ErrGuardar:
    MsgBox Err.Description, vbCritical, "Aviso"
    Screen.MousePointer = 0
    cmdGuardar.Enabled = True
End Sub
Private Sub cmdConfigurar_Click()
    Dim frmEscalonaConf As frmCredNewNivAutorizaConfEscalona
    Dim index As Integer
    Dim bEscalonaOK As Boolean
    
    index = feAutorizacion.row
    
    Select Case feAutorizacion.TextMatrix(index, 1)
        Case "TIP0002": 'Escalonamiento
            Set frmEscalonaConf = New frmCredNewNivAutorizaConfEscalona
            bEscalonaOK = frmEscalonaConf.Inicio(fvEscalonaConf)
            If bEscalonaOK Then
                fnModificaEscalonaConfig = fnModificaEscalonaConfig + 1
            End If
    End Select

    HabilitarGuardar
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
