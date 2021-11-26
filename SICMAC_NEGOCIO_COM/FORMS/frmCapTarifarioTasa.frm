VERSION 5.00
Begin VB.Form frmCapTarifarioTasa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captaciones - Tasas de Interes"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9075
   Icon            =   "frmCapTarifarioTasa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTasasDolares 
      Caption         =   "Tasas - Dolares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   90
      TabIndex        =   20
      Top             =   4725
      Width           =   8910
      Begin VB.CommandButton btnAgregaDolares 
         Caption         =   "Agregar"
         Height          =   300
         Left            =   7830
         TabIndex        =   22
         Top             =   270
         Width           =   960
      End
      Begin VB.CommandButton btnQuitaDolares 
         Caption         =   "Quitar"
         Height          =   300
         Left            =   7830
         TabIndex        =   21
         Top             =   630
         Width           =   960
      End
      Begin SICMACT.FlexEdit grdTasasDolar 
         Height          =   3075
         Left            =   135
         TabIndex        =   23
         Top             =   270
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   5424
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-nIdTasaDet-Monto Ini-Monto Fin-Plazo Ini-Plazo Fin-Tasa Interes-Tasa Ord. Pago-tmp"
         EncabezadosAnchos=   "0-0-1000-1000-1000-1000-1200-1500-0"
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
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-8"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-L"
         FormatosEdit    =   "0-0-2-2-3-3-2-2-0"
         TextArray0      =   "Nro"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CheckBox ckAplicacionInmediata 
      Alignment       =   1  'Right Justify
      Caption         =   "Aplicación Inmediata"
      Height          =   240
      Left            =   6975
      TabIndex        =   16
      Top             =   8370
      Width           =   1995
   End
   Begin VB.CommandButton btnNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   135
      TabIndex        =   14
      Top             =   8730
      Width           =   960
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1125
      TabIndex        =   13
      Top             =   8730
      Width           =   960
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7020
      TabIndex        =   12
      Top             =   8730
      Width           =   960
   End
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6030
      TabIndex        =   11
      Top             =   8730
      Width           =   960
   End
   Begin VB.CommandButton btnGuardarComo 
      Caption         =   "Guardar Como"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4410
      TabIndex        =   10
      Top             =   8730
      Width           =   1590
   End
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8010
      TabIndex        =   9
      Top             =   8730
      Width           =   960
   End
   Begin VB.Frame fraTasasSoles 
      Caption         =   "Tasas - Soles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   90
      TabIndex        =   5
      Top             =   1170
      Width           =   8910
      Begin VB.CommandButton btnQuitaSoles 
         Caption         =   "Quitar"
         Height          =   300
         Left            =   7830
         TabIndex        =   8
         Top             =   630
         Width           =   960
      End
      Begin VB.CommandButton btnAgregaSoles 
         Caption         =   "Agregar"
         Height          =   300
         Left            =   7830
         TabIndex        =   7
         Top             =   270
         Width           =   960
      End
      Begin SICMACT.FlexEdit grdTasasSoles 
         Height          =   3075
         Left            =   135
         TabIndex        =   6
         Top             =   270
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   5424
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-nIdTasaDet-Monto Ini-Monto Fin-Plazo Ini-Plazo Fin-Tasa Interes-Tasa Ord. Pago-tmp"
         EncabezadosAnchos=   "0-0-1000-1000-1000-1000-1200-1500-0"
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
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-8"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-L"
         FormatosEdit    =   "0-0-2-2-3-3-2-2-0"
         TextArray0      =   "Nro"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   8910
      Begin VB.CommandButton btnSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   300
         Left            =   7290
         TabIndex        =   18
         Top             =   270
         Width           =   1050
      End
      Begin VB.CommandButton btnExportar 
         Caption         =   "Exportar Todo"
         Height          =   300
         Left            =   2115
         TabIndex        =   17
         Top             =   630
         Width           =   1320
      End
      Begin VB.ComboBox cbGrupo 
         Height          =   315
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   1635
      End
      Begin VB.ComboBox cbProducto 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   2265
      End
      Begin VB.ComboBox cbPersoneria 
         Height          =   315
         Left            =   4995
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   2265
      End
      Begin VB.CommandButton btnExaminar 
         Caption         =   "Examinar"
         Height          =   300
         Left            =   1035
         TabIndex        =   19
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Agencias:"
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version actual: "
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   8370
      Width           =   4020
   End
End
Attribute VB_Name = "frmCapTarifarioTasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************
'* NOMBRE         : frmCapTarifarioTasa
'* DESCRIPCION    : Proyecto - Tarifario Versionado - Creacion y Mantenimiento de Tasas
'* CREACION       : RIRO, 20160420 10:00 AM
'************************************************************************************************************

Option Explicit

Private bFocoGridSoles As Boolean   'ver si el flexEdit soles tiene el foco
Private bFocoGridDolares As Boolean 'ver si el flexEdit dolares tiene el foco
Private bPresionaEnter As Boolean 'artificio para mejor experiencia de usuario en el grid
Private nTipoOperacion As Integer '0=Sin definir, 1=Nuevo, 2=Edicion
Private nProducto As Integer 'ahorros=232, plazo fijo=233, cts = 234
Private oCon As COMNCaptaGenerales.NCOMCaptaDefinicion
Private nIdTasa As Integer ' ID de la tasa seleccionada.
Private oTasa As tTasa 'Objeto Tasa, Ubicacion: COMDConstantes/DCOMValores/tTasa

Public Sub Inicio(pnProducto As Integer)
    nProducto = pnProducto
    Limpiar
    Me.Show 1
End Sub
Private Sub CargarControles()
    Dim oConstante As COMDConstSistema.DCOMGeneral
    Dim oCon As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsTmp As ADODB.Recordset
    
    'cargando los grupos
    Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsTmp = oCon.ObtenerGruposComision(1)
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF And Not rsTmp.BOF Then
            CargaCombo cbGrupo, rsTmp
            Set rsTmp = Nothing
        Else
            cbGrupo.Clear
        End If
    Else
        cbGrupo.Clear
    End If
    Set oCon = Nothing
    Set oConstante = New COMDConstSistema.DCOMGeneral
    
    'cargando los subproductos de ahorros
    Set rsTmp = oConstante.GetConstante(IIf(nProducto = 232, 2030, IIf(nProducto = 233, 2032, IIf(nProducto = 234, 2033, -1))), , "", "-")
    CargaCombo cbProducto, rsTmp
    Set rsTmp = Nothing
    
    'cargando las personerias
    Set rsTmp = oConstante.GetConstante(1002, , "'[12]'", "-")
    CargaCombo cbPersoneria, rsTmp
    Set rsTmp = Nothing
    
    'seleccionando el primer registro de los combos
    If cbGrupo.ListCount > 0 Then cbGrupo.ListIndex = 0
    If cbProducto.ListCount > 0 Then cbProducto.ListIndex = 0
    If cbPersoneria.ListCount > 0 Then cbPersoneria.ListIndex = 0
End Sub

Private Sub btnAgregaDolares_Click()
    grdTasasDolar.SetFocus
    grdTasasDolar.lbEditarFlex = True
    grdTasasDolar.AdicionaFila
    grdTasasDolar.TextMatrix(grdTasasDolar.Rows - 1, 1) = 0
    grdTasasDolar.Col = 2
    btnAgregaDolares.Default = False
    SendKeys "{F2}"
End Sub

Private Sub btnAgregaSoles_Click()
    grdTasasSoles.SetFocus
    grdTasasSoles.lbEditarFlex = True
    grdTasasSoles.AdicionaFila
    grdTasasSoles.TextMatrix(grdTasasSoles.Rows - 1, 1) = 0
    grdTasasSoles.Col = 2
    btnAgregaSoles.Default = False
    SendKeys "{F2}"
End Sub

Private Sub btnCancelar_Click()
Limpiar

End Sub

Private Sub btnEditar_Click()
nTipoOperacion = 2
BlqControles (4)
End Sub

Private Sub btnExaminar_Click()

    Dim rsVersiones As ADODB.Recordset
    Dim bVersiones As Boolean
    Dim oExaminar As frmCapTarifarioExaminar
    
    'Buscando versiones que coincidan
    Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
    bVersiones = False
    Set rsVersiones = oCon.ObtenerVersionesTasa(CStr(Trim(Left(cbGrupo.Text, 5))), 232, CInt(Trim(Right(cbProducto.Text, 5))), CInt(Trim(Right(cbPersoneria.Text, 5))))
    If Not rsVersiones Is Nothing Then
        If Not rsVersiones.EOF And Not rsVersiones.BOF Then
            If rsVersiones.RecordCount > 0 Then
                bVersiones = True
            End If
        End If
    Else
        bVersiones = False
    End If
    'mostrando lista de versiones encontradas
    If bVersiones Then
        Set oExaminar = New frmCapTarifarioExaminar
        oExaminar.nTipo = 2 ' tasas
        oExaminar.rsExaminar = rsVersiones
        oExaminar.Show 1
        If oExaminar.bRespuesta Then
            oTasa.IdTasa = oExaminar.Id
            If oCon.ObtenerTasaVersion(oTasa) Then
                PintarTasa oTasa
                BlqControles (2)
            Else
                MsgBox "No se puedo acceder a los datos de la versiòn seleccionada", vbInformation, "Aviso"
            End If
        End If
    Else
        MsgBox "No se encontraron versiones", vbInformation, "Aviso"
    End If
    Set oCon = Nothing
End Sub
Private Sub PintarTasa(pTasa As tTasa)

    'cargando detalle
    LimpiaFlex grdTasasSoles
    LimpiaFlex grdTasasDolar
    
    lblVersion.Caption = "Version Actual: " & Right(pTasa.MovNro, 4) & " - " & IIf(pTasa.Version < 10, "0" & pTasa.Version, pTasa.Version) & " " & pTasa.FechaRegistro & " - Registro Inicial"
    ckAplicacionInmediata.value = pTasa.AplicaInmediato
    With pTasa
        If Not .rsTasa Is Nothing Then
            If Not .rsTasa.EOF Then
                If .rsTasa.RecordCount > 0 Then
                    Do While Not .rsTasa.EOF
                        If .rsTasa("nMoneda") = 1 Then 'Soles
                            grdTasasSoles.AdicionaFila
                            grdTasasSoles.TextMatrix(grdTasasSoles.Rows - 1, 1) = .rsTasa("nIdTarifarioTasaDet")
                            grdTasasSoles.TextMatrix(grdTasasSoles.Rows - 1, 2) = Format(.rsTasa("nMontoIni"), "0.00")
                            grdTasasSoles.TextMatrix(grdTasasSoles.Rows - 1, 3) = Format(.rsTasa("nMontoFin"), "0.00")
                            grdTasasSoles.TextMatrix(grdTasasSoles.Rows - 1, 4) = .rsTasa("nPlazoIni")
                            grdTasasSoles.TextMatrix(grdTasasSoles.Rows - 1, 5) = .rsTasa("nPlazoFin")
                            grdTasasSoles.TextMatrix(grdTasasSoles.Rows - 1, 6) = Format(.rsTasa("nTasaInteres"), "0.00")
                            grdTasasSoles.TextMatrix(grdTasasSoles.Rows - 1, 7) = Format(.rsTasa("nTasaOrdenPago"), "0.00")
                            
                        ElseIf .rsTasa("nMoneda") = 2 Then 'Dolares
                            grdTasasDolar.AdicionaFila
                            grdTasasDolar.TextMatrix(grdTasasDolar.Rows - 1, 1) = .rsTasa("nIdTarifarioTasaDet")
                            grdTasasDolar.TextMatrix(grdTasasDolar.Rows - 1, 2) = Format(.rsTasa("nMontoIni"), "0.00")
                            grdTasasDolar.TextMatrix(grdTasasDolar.Rows - 1, 3) = Format(.rsTasa("nMontoFin"), "0.00")
                            grdTasasDolar.TextMatrix(grdTasasDolar.Rows - 1, 4) = .rsTasa("nPlazoIni")
                            grdTasasDolar.TextMatrix(grdTasasDolar.Rows - 1, 5) = .rsTasa("nPlazoFin")
                            grdTasasDolar.TextMatrix(grdTasasDolar.Rows - 1, 6) = Format(.rsTasa("nTasaInteres"), "0.00")
                            grdTasasDolar.TextMatrix(grdTasasDolar.Rows - 1, 7) = Format(.rsTasa("nTasaOrdenPago"), "0.00")
                        
                        End If
                        .rsTasa.MoveNext
                    Loop
                    .rsTasa.MoveFirst
                End If
            End If
        End If
    End With
    
End Sub
Private Function Validar() As String
Dim bValida As Boolean
Dim i As Integer, j As Integer
Dim sMensaje As String

bValida = True

    If grdTasasSoles.Rows = 2 And Len(Trim(grdTasasSoles.TextMatrix(1, 1))) = 0 Then
        sMensaje = "El Grid de tasas en Soles no contiene registros" & vbNewLine
    End If
    If grdTasasDolar.Rows = 2 And Len(Trim(grdTasasDolar.TextMatrix(1, 1))) = 0 Then
        sMensaje = sMensaje & "El Grid de tasas en Dolarres no contiene registros" & vbNewLine
    End If

If nTipoOperacion < 1 Or nTipoOperacion > 2 Then
sMensaje = sMensaje & "Operación no identificada" & vbNewLine
End If
If Len(Trim(sMensaje)) = 0 Then
    For i = 1 To grdTasasSoles.Rows - 1
        For j = 2 To 7
            If Not IsNumeric(Trim(grdTasasSoles.TextMatrix(i, j))) Then
                bValida = False
                j = 7
            End If
        Next j
        If Not bValida Then
            i = grdTasasSoles.Rows
        End If
    Next i
    If Not bValida Then
        sMensaje = "El Grid de tasas en Soles debe contener solo valores numericos" & vbNewLine
    End If
    bValida = True
    For i = 1 To grdTasasDolar.Rows - 1
        For j = 2 To 7
            If Not IsNumeric(Trim(grdTasasDolar.TextMatrix(i, j))) Then
                bValida = False
                j = 7
            End If
        Next j
        If Not bValida Then
            i = grdTasasSoles.Rows
        End If
    Next i
    If Not bValida Then
        sMensaje = sMensaje & "El Grid de tasas en Dolares debe contener solo valores numericos" & vbNewLine
    End If
End If

'validando en caso de edicion
If nTipoOperacion = 2 And Len(Trim(sMensaje)) = 0 Then
    Dim rsSol As ADODB.Recordset, rsDol As ADODB.Recordset
    Dim Lst() As Integer
    ObtenerTasasXactualizar rsSol, rsDol, Lst
    If rsSol.RecordCount = 0 And rsDol.RecordCount = 0 And UBound(Lst) = 0 And oTasa.AplicaInmediato = CInt(ckAplicacionInmediata.value) Then
        sMensaje = sMensaje & "No existen registros para la actualización" & vbNewLine
    End If
    Set rsSol = Nothing
    Set rsDol = Nothing
End If
Validar = Trim(sMensaje)

End Function
Private Sub Limpiar()
    LimpiaFlex grdTasasSoles
    LimpiaFlex grdTasasDolar
    nTipoOperacion = 0
    bPresionaEnter = False
    nIdTasa = -1
    ckAplicacionInmediata.value = 0
    BlqControles (0)
    'seleccionando el primer registro de los combos
    If cbGrupo.ListCount > 0 Then cbGrupo.ListIndex = 0
    If cbProducto.ListCount > 0 Then cbProducto.ListIndex = 0
    If cbPersoneria.ListCount > 0 Then cbPersoneria.ListIndex = 0
    'limpiando la tasa ***
    LipiarTasa oTasa
    'end de limpiar tasa
End Sub
Private Sub LipiarTasa(poTasa As tTasa)
    Set poTasa.rsTasa = Nothing
    poTasa.AplicaInmediato = 0
    poTasa.Estado = 0
    poTasa.FechaRegistro = Now
    poTasa.Glosa = ""
    poTasa.Grupo = ""
    poTasa.IdTasa = 0
    poTasa.MovNro = ""
    poTasa.Personeria = 0
    poTasa.Producto = 0
    poTasa.SubProducto = 0
    poTasa.Version = 0
End Sub
Private Sub crearRStasa(ByRef rs As Recordset)
    Set rs = New ADODB.Recordset
    rs.Fields.Append "nIdTarifarioTasaDet", adInteger
    rs.Fields.Append "nIdTarifarioTasaCab", adInteger
    rs.Fields.Append "nMontoIni", adDouble
    rs.Fields.Append "nMontoFin", adDouble
    rs.Fields.Append "nPlazoIni", adInteger
    rs.Fields.Append "nPlazoFin", adInteger
    rs.Fields.Append "nTasaInteres", adDouble
    rs.Fields.Append "nTasaOrdenPago", adDouble
    rs.Fields.Append "nMoneda", adInteger
    rs.Fields.Append "nEstado", adInteger
    rs.Fields.Append "nTipo", adInteger '1=actualizacion, 2=eliminacion
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.Open
End Sub
Private Sub llenaFila(ByRef rs As Recordset, ByVal Flex As Control, ByVal nFila As Integer, ByVal nTasaCab As Integer, ByVal nmoneda As Integer, ByVal nTipo As Integer)
    rs.AddNew
    rs.Fields("nIdTarifarioTasaDet") = Flex.TextMatrix(nFila, 1)
    rs.Fields("nIdTarifarioTasaCab") = nTasaCab
    rs.Fields("nMontoIni") = Flex.TextMatrix(nFila, 2)
    rs.Fields("nMontoFin") = Flex.TextMatrix(nFila, 3)
    rs.Fields("nPlazoIni") = Flex.TextMatrix(nFila, 4)
    rs.Fields("nPlazoFin") = Flex.TextMatrix(nFila, 5)
    rs.Fields("nTasaInteres") = Flex.TextMatrix(nFila, 6)
    rs.Fields("nTasaOrdenPago") = Flex.TextMatrix(nFila, 7)
    rs.Fields("nMoneda") = nmoneda
    rs.Fields("nEstado") = 1
    rs.Fields("nTipo") = nTipo '1=Nuevo, 2=Edicion
End Sub
Private Sub ObtenerTasasXactualizar(ByRef rsSol As ADODB.Recordset, ByRef rsDol As ADODB.Recordset, ByRef lstElim As Variant)
    
    Dim i As Integer
    Dim rsIni As ADODB.Recordset
    
    Set rsIni = oTasa.rsTasa.Clone
    crearRStasa rsSol
    crearRStasa rsDol
    
    'revisando si hay actualizaciones o nuevas tasas en SOLES S/. S/. S/.
    rsIni.Filter = "nMoneda = 1" 'Filtrando soles
    If Len(Trim(grdTasasSoles.TextMatrix(1, 1))) > 0 Then
        For i = 1 To grdTasasSoles.Rows - 1
            ' si es un registro nuevo
            If grdTasasSoles.TextMatrix(i, 1) = 0 Then
                llenaFila rsSol, grdTasasSoles, i, oTasa.IdTasa, 1, 1
    
            'si es un registro ya existente
            Else
                If Not rsIni Is Nothing Then
                    If Not rsIni.EOF And Not rsIni.BOF Then
                        If rsIni.RecordCount > 0 Then
                            Do While Not rsIni.EOF And Not rsIni.BOF
                                If grdTasasSoles.TextMatrix(i, 1) = rsIni("nIdTarifarioTasaDet") And rsIni("nMoneda") = 1 Then 'Solo monedas en soles
                                    If rsIni!nMontoIni <> CDbl(grdTasasSoles.TextMatrix(i, 2)) Or _
                                       rsIni!nMontoFin <> CDbl(grdTasasSoles.TextMatrix(i, 3)) Or _
                                       rsIni!nPlazoIni <> CDbl(grdTasasSoles.TextMatrix(i, 4)) Or _
                                       rsIni!nPlazoFin <> CDbl(grdTasasSoles.TextMatrix(i, 5)) Or _
                                       rsIni!nTasaInteres <> CDbl(grdTasasSoles.TextMatrix(i, 6)) Or _
                                       rsIni!nTasaOrdenPago <> CDbl(grdTasasSoles.TextMatrix(i, 7)) Then
                                    
                                        llenaFila rsSol, grdTasasSoles, i, oTasa.IdTasa, 1, 2
                                    End If
                                    rsIni.MoveLast
                                End If
                                rsIni.MoveNext
                            Loop
                            rsIni.MoveFirst
                        End If
                    End If
                End If
            End If
        Next i
        If Not rsSol.EOF Then rsSol.MoveFirst
    End If
    'revisando si hay actualizaciones o nuevas tasas en DOLARES $$$$$
    rsIni.Filter = "nMoneda = 2" 'Filtrando dolares
    If Len(Trim(grdTasasDolar.TextMatrix(1, 1))) > 0 Then
        For i = 1 To grdTasasDolar.Rows - 1
            ' si es un registro nuevo
            If grdTasasDolar.TextMatrix(i, 1) = 0 Then
                llenaFila rsDol, grdTasasDolar, i, oTasa.IdTasa, 2, 1
                
            ' si es un registro ya existente
            Else
                If Not rsIni Is Nothing Then
                    If Not rsIni.EOF And Not rsIni.BOF Then
                        If rsIni.RecordCount > 0 Then
                            Do While Not rsIni.EOF And Not rsIni.BOF
                                If grdTasasDolar.TextMatrix(i, 1) = rsIni("nIdTarifarioTasaDet") And rsIni("nMoneda") = 2 Then 'Solo monedas en soles
                                    If rsIni!nMontoIni <> CDbl(grdTasasDolar.TextMatrix(i, 2)) Or _
                                       rsIni!nMontoFin <> CDbl(grdTasasDolar.TextMatrix(i, 3)) Or _
                                       rsIni!nPlazoIni <> CDbl(grdTasasDolar.TextMatrix(i, 4)) Or _
                                       rsIni!nPlazoFin <> CDbl(grdTasasDolar.TextMatrix(i, 5)) Or _
                                       rsIni!nTasaInteres <> CDbl(grdTasasDolar.TextMatrix(i, 6)) Or _
                                       rsIni!nTasaOrdenPago <> CDbl(grdTasasDolar.TextMatrix(i, 7)) Then
                                    
                                        llenaFila rsDol, grdTasasDolar, i, oTasa.IdTasa, 2, 2
                                    End If
                                    rsIni.MoveLast
                                End If
                                rsIni.MoveNext
                            Loop
                            rsIni.MoveFirst
                        End If
                    End If
                End If
            End If
        Next i
        If Not rsDol.EOF Then rsDol.MoveFirst
    End If
    rsIni.Filter = "nMoneda <> 999"
    'verificando si hubieron registros eliminados
    Dim bEncontrado As Boolean
    Dim n As Integer 'Cantidad de registros
    bEncontrado = False
    n = 0
    ReDim Preserve lstElim(n)
    If Not rsIni Is Nothing Then
        If Not rsIni.EOF And Not rsIni.BOF Then
            If rsIni.RecordCount > 0 Then
                Do While Not rsIni.EOF And Not rsIni.BOF
                    bEncontrado = False
                    If rsIni!nmoneda = 1 Then
                        'buscando en todo el grid de tasas en soles
                        If Len(Trim(grdTasasSoles.TextMatrix(1, 1))) > 0 Then
                            For i = 1 To grdTasasSoles.Rows - 1
                                If rsIni!nIdTarifarioTasaDet = grdTasasSoles.TextMatrix(i, 1) Then
                                    bEncontrado = True
                                    i = grdTasasSoles.Rows
                                End If
                            Next i
                        End If
                    ElseIf rsIni!nmoneda = 2 Then
                        'buscando en todo el grid de tasas en soles
                        If Len(Trim(grdTasasDolar.TextMatrix(1, 1))) > 0 Then
                            For i = 1 To grdTasasDolar.Rows - 1
                                If rsIni!nIdTarifarioTasaDet = grdTasasDolar.TextMatrix(i, 1) Then
                                    bEncontrado = True
                                    i = grdTasasDolar.Rows
                                End If
                            Next i
                        End If
                    End If
                    'verificando si se encontro
                    If Not bEncontrado Then
                        n = n + 1
                        ReDim Preserve lstElim(n)
                        lstElim(n) = rsIni!nIdTarifarioTasaDet
                    End If
                    rsIni.MoveNext
                Loop
                rsIni.MoveFirst
            End If
        End If
    End If
End Sub
Private Sub btnGuardar_Click()
    Dim sMensaje As String
    sMensaje = Validar

    If Len(Trim(sMensaje)) = 0 Then

        Dim bRespuesta As Boolean
        Dim rsSol As ADODB.Recordset
        Dim rsDol As ADODB.Recordset
        Dim oCont As COMNContabilidad.NCOMContFunciones
        Dim oFrmGuardar As frmCapTarifarioGuardar
        Dim lstElim() As Integer

        sMensaje = IIf(nTipoOperacion = 1, "¿Desea crear una nueva version de tasas?", "¿Desea actualizar la actual versión de Tasas?")
               
        If MsgBox(sMensaje, vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then

            Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
            Set oCont = New COMNContabilidad.NCOMContFunciones

            oTasa.Producto = nProducto
            oTasa.SubProducto = CInt(Trim(Right(cbProducto.Text, 5)))
            oTasa.Personeria = CInt(Trim(Right(cbPersoneria.Text, 5)))
            oTasa.Grupo = Trim(Left(cbGrupo.Text, 5))
            oTasa.MovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            oTasa.Estado = 1
            oTasa.AplicaInmediato = CInt(ckAplicacionInmediata.value)

            If nTipoOperacion = 1 Then ' Nuevo
                'obtener ultima version de tasa
                Set oFrmGuardar = New frmCapTarifarioGuardar
                oFrmGuardar.nVersion = oCon.ObtenerUltimaVersionTasa(oTasa.Grupo, oTasa.Producto, oTasa.SubProducto, oTasa.Personeria)
                oFrmGuardar.dFechaRegistro = oTasa.FechaRegistro
                oFrmGuardar.nTipo = 2 ' tasas
                oFrmGuardar.Caption = "Guardar..."
                oFrmGuardar.Show 1
                If oFrmGuardar.bRespuesta Then
                    Set rsSol = grdTasasSoles.GetRsNew
                    Set rsDol = grdTasasDolar.GetRsNew
                    oTasa.Glosa = oFrmGuardar.sGosa
                    oTasa.Version = oFrmGuardar.nVersion
                    bRespuesta = oCon.AgregaTasaVersion(oTasa, rsSol, rsDol)
                    If bRespuesta Then
                        Limpiar
                        MsgBox "Los datos se guardaron correctamente", vbInformation, "Aviso"
                    Else
                        MsgBox "Se presentaron inconvenientes durante la grabacion.", vbInformation, "Aviso"
                    End If
                End If
                Set oFrmGuardar = Nothing
             
            ElseIf nTipoOperacion = 2 Then ' Edición
            
                ObtenerTasasXactualizar rsSol, rsDol, lstElim
                bRespuesta = oCon.ActualizaTasaVersion(oTasa, rsSol, rsDol, lstElim)
                If bRespuesta Then
                    'cargando los datos guardados
                    oCon.ObtenerTasaVersion oTasa
                    PintarTasa oTasa
                    BlqControles (2)
                    nTipoOperacion = 0
                    MsgBox "Los datos se guardaron correctamente", vbInformation, "Aviso"
                Else
                    MsgBox "Se presentaron inconvenientes durante la grabacion.", vbInformation, "Aviso"
                End If
            End If
            Set oCont = Nothing
            Set oCon = Nothing
        End If
    Else
        MsgBox "Observaciones: " & vbNewLine & sMensaje, vbInformation, "Aviso"
    End If
End Sub

Private Sub btnNuevo_Click()
nTipoOperacion = 1
LimpiaFlex grdTasasSoles
LimpiaFlex grdTasasDolar
BlqControles (3)
btnAgregaSoles.SetFocus
End Sub

Private Sub btnQuitaDolares_Click()
    Dim nFila As Integer
    nFila = grdTasasDolar.row
    If Len(Trim(grdTasasDolar.TextMatrix(nFila, 1))) = 0 Then
        MsgBox "Debe seleccionar un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿Desea eliminar el registro seleccionado?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        grdTasasDolar.EliminaFila nFila
    End If
End Sub

Private Sub btnQuitaSoles_Click()
    Dim nFila As Integer
    nFila = grdTasasSoles.row
    If Len(Trim(grdTasasSoles.TextMatrix(nFila, 1))) = 0 Then
        MsgBox "Debe seleccionar un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿Desea eliminar el registro seleccionado?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        grdTasasSoles.EliminaFila nFila
    End If
End Sub

Private Sub btnSalir_Click()
    If MsgBox("Desea salir del formulario de tasas?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub btnSeleccionar_Click()
    Dim bSeleccionar As Boolean
    bSeleccionar = True
    If cbGrupo.ListIndex < 0 Then
        bSeleccionar = False
    End If
    If cbProducto.ListIndex < 0 Then
        bSeleccionar = False
    End If
    If cbPersoneria.ListIndex < 0 Then
        bSeleccionar = False
    End If
    If Not bSeleccionar Then
        MsgBox "Debe seleccionar Grupo, Producto y Personería", vbInformation, "Aviso"
    Else
    BlqControles (1)
    End If
End Sub

Private Sub BlqControles(ByVal nTipoBloqueo As Integer)

If nTipoBloqueo = 0 Then 'Inicio

    cbGrupo.Enabled = True
    cbProducto.Enabled = True
    cbPersoneria.Enabled = True
    btnSeleccionar.Enabled = True
    btnExaminar.Enabled = False
    btnExportar.Enabled = False

    grdTasasSoles.lbEditarFlex = False
    grdTasasSoles.Enabled = False
    btnAgregaSoles.Enabled = False
    btnQuitaSoles.Enabled = False
    
    grdTasasDolar.lbEditarFlex = False
    grdTasasDolar.Enabled = False
    btnAgregaDolares.Enabled = False
    btnQuitaDolares.Enabled = False

    btnNuevo.Enabled = False
    btnEditar.Enabled = False
    ckAplicacionInmediata.Enabled = False
    btnGuardarComo.Enabled = False
    btnGuardar.Enabled = False
    btnCancelar.Enabled = True
    btnSalir.Enabled = True
    
ElseIf nTipoBloqueo = 1 Then ' Seleccionar
    
    cbGrupo.Enabled = False
    cbProducto.Enabled = False
    cbPersoneria.Enabled = False
    btnSeleccionar.Enabled = False
    btnExaminar.Enabled = True
    btnExportar.Enabled = False

    grdTasasSoles.lbEditarFlex = False
    grdTasasSoles.Enabled = False
    btnAgregaSoles.Enabled = False
    btnQuitaSoles.Enabled = False
    
    grdTasasDolar.lbEditarFlex = False
    grdTasasDolar.Enabled = False
    btnAgregaDolares.Enabled = False
    btnQuitaDolares.Enabled = False

    btnNuevo.Enabled = True
    btnEditar.Enabled = False
    ckAplicacionInmediata.Enabled = False
    btnGuardarComo.Enabled = False
    btnGuardar.Enabled = False
    btnCancelar.Enabled = True
    btnSalir.Enabled = True

ElseIf nTipoBloqueo = 2 Then 'Examinar, asumiendo que encuentran versiones anteriores

    cbGrupo.Enabled = False
    cbProducto.Enabled = False
    cbPersoneria.Enabled = False
    btnSeleccionar.Enabled = False
    btnExaminar.Enabled = False
    btnExportar.Enabled = True

    grdTasasSoles.lbEditarFlex = False
    grdTasasSoles.Enabled = True
    btnAgregaSoles.Enabled = False
    btnQuitaSoles.Enabled = False
    
    grdTasasDolar.lbEditarFlex = False
    grdTasasDolar.Enabled = True
    btnAgregaDolares.Enabled = False
    btnQuitaDolares.Enabled = False

    btnNuevo.Enabled = True
    btnEditar.Enabled = True
    ckAplicacionInmediata.Enabled = False
    btnGuardarComo.Enabled = True
    btnGuardar.Enabled = False
    btnCancelar.Enabled = True
    btnSalir.Enabled = True

ElseIf nTipoBloqueo = 3 Then ' Nuevo

    cbGrupo.Enabled = False
    cbProducto.Enabled = False
    cbPersoneria.Enabled = False
    btnSeleccionar.Enabled = False
    btnExaminar.Enabled = False
    btnExportar.Enabled = True

    grdTasasSoles.Enabled = True
    btnAgregaSoles.Enabled = True
    btnQuitaSoles.Enabled = True
    
    grdTasasDolar.Enabled = True
    btnAgregaDolares.Enabled = True
    btnQuitaDolares.Enabled = True

    btnNuevo.Enabled = False
    btnEditar.Enabled = False
    ckAplicacionInmediata.Enabled = True
    btnGuardarComo.Enabled = False
    btnGuardar.Enabled = True
    btnCancelar.Enabled = True
    btnSalir.Enabled = True

ElseIf nTipoBloqueo = 4 Then ' Editar

    cbGrupo.Enabled = False
    cbProducto.Enabled = False
    cbPersoneria.Enabled = False
    btnSeleccionar.Enabled = False
    btnExaminar.Enabled = False
    btnExportar.Enabled = True

    grdTasasSoles.lbEditarFlex = True
    grdTasasSoles.Enabled = True
    btnAgregaSoles.Enabled = True
    btnQuitaSoles.Enabled = True
    
    grdTasasDolar.lbEditarFlex = True
    grdTasasDolar.Enabled = True
    btnAgregaDolares.Enabled = True
    btnQuitaDolares.Enabled = True

    btnNuevo.Enabled = False
    btnEditar.Enabled = False
    ckAplicacionInmediata.Enabled = True
    btnGuardarComo.Enabled = True
    btnGuardar.Enabled = True
    btnCancelar.Enabled = True
    btnSalir.Enabled = True


End If
End Sub
Private Sub Form_Initialize()
bPresionaEnter = False
nIdTasa = -1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'controlando el Ctrl + V
    If KeyCode = 86 And Shift = 2 And (bFocoGridSoles Or bFocoGridDolares) Then
        KeyCode = 10
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim nFil As Integer, nCol As Integer
    If bFocoGridSoles Then
        nCol = grdTasasSoles.Col
        If KeyAscii = 13 Then
            bPresionaEnter = True
        Else
            If nCol = 2 Or nCol = 3 Or nCol = 6 Or nCol = 7 Then ' numeros decimales
                If KeyAscii <> 46 Then
                    KeyAscii = NumerosEnteros(KeyAscii)
                End If
                btnAgregaSoles.Default = False
            ElseIf nCol = 4 Or nCol = 5 Then ' numeros enteros
                KeyAscii = NumerosEnteros(KeyAscii)
                btnAgregaSoles.Default = False
            End If
        End If
    End If
    If bFocoGridDolares Then
        nCol = grdTasasDolar.Col
        If KeyAscii = 13 Then
            bPresionaEnter = True
        Else
            If nCol = 2 Or nCol = 3 Or nCol = 6 Or nCol = 7 Then ' numeros decimales
                If KeyAscii <> 46 Then
                    KeyAscii = NumerosEnteros(KeyAscii)
                End If
                btnAgregaDolares.Default = False
            ElseIf nCol = 4 Or nCol = 5 Then ' numeros enteros
                KeyAscii = NumerosEnteros(KeyAscii)
                btnAgregaDolares.Default = False
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
CargarControles
End Sub
Private Sub grdTasasDolar_GotFocus()
bFocoGridDolares = True
btnAgregaSoles.Default = False
End Sub
Private Sub grdTasasDolar_LostFocus()
bFocoGridDolares = False
End Sub
Private Sub grdTasasDolar_RowColChange()
    Dim nCol As Integer
    nCol = grdTasasDolar.Col
    If nCol = 1 Then ' si es la columna final
        btnAgregaDolares.SetFocus
        btnAgregaDolares.Default = True
    Else
        If bPresionaEnter Then
            SendKeys "{F2}"
            bPresionaEnter = False
        End If
        
    End If
End Sub
Private Sub grdTasasSoles_GotFocus()
    bFocoGridSoles = True
    btnAgregaDolares.Default = False
End Sub
Private Sub grdTasasSoles_LostFocus()
    bFocoGridSoles = False
End Sub
Private Sub grdTasasSoles_RowColChange()
    Dim nCol As Integer
    nCol = grdTasasSoles.Col
    If nCol = 1 Then ' si es la columna final
        btnAgregaSoles.SetFocus
        btnAgregaSoles.Default = True
    Else
        If bPresionaEnter Then
            SendKeys "{F2}"
            bPresionaEnter = False
        End If
        
    End If
End Sub

