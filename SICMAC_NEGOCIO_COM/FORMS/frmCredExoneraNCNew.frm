VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredExoneraNCNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizaciones no contempladas"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "frmCredExoneraNCNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   310
      Left            =   6960
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   310
      Left            =   6000
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   310
      Left            =   1200
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   310
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin SICMACT.FlexEdit feExoneraciones 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3201
      Cols0           =   8
      HighLight       =   1
      EncabezadosNombres=   "#-Autorización-Descripción-Usuario-ExoneraItem-cMovNro-TipoRegistro-nId"
      EncabezadosAnchos=   "300-7000-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-1-X-X-X-X-X-X"
      ListaControles  =   "0-3-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5106
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Autorizaciones no contempladas Solicitadas"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCredExoneraNCNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************
'AUTOR      : RECO, Renzo Cordova Lozano
'Nombre     : frmCredExoneraNCNew
'Descripcion: El formulario se creo con la finalidad de administrar la exoneraciones no contempladas
'Fecha Crea.: 25/05/2016
'****************************************************************************************************

Option Explicit
Dim fsCtaCod As String
Dim fsNivAprCod As String
Dim fbNuevoRegistro As Boolean
Dim fbExononeracion As Boolean
Dim fnTipoCargaSugerencia 'FRHU 20160820

Dim nExito As Integer 'JOEP20171120

'Public Function inicia(ByVal psCtaCod As String) As Boolean
Public Function inicia(ByVal psCtaCod As String, Optional ByVal pnTipoCarga As Integer = 1) As Boolean 'FRHU 20160820
    'FRHU 20160820
    fnTipoCargaSugerencia = pnTipoCarga
    If fnTipoCargaSugerencia = 2 Then 'lSugerTipoActRegistro = 1 , lSugerTipoActConsultar = 2
        CmdAceptar.Enabled = False
        cmdAgregar.Enabled = False
        cmdQuitar.Enabled = False
    End If
    'FIN FRHU 20160820
    fbExononeracion = False
    fsCtaCod = psCtaCod
    fbNuevoRegistro = True
    Call CargaDatosNivel
    Call CargarListaExoneraciones(psCtaCod)
    Call CargaAutorizaciones
    Me.Show 1
    inicia = fbExononeracion
End Function

Private Sub cmdAceptar_Click()
    If Guardar Then
        If nExito = 1 Then
            MsgBox "Los datos se guardaron de forma correcta", vbInformation, "Alerta"
            FEExoneraciones.lbEditarFlex = False 'FRHU 20160819
            Unload Me
        ElseIf nExito = 2 Then
            MsgBox "Los datos se guardaron de forma correcta", vbInformation, "Alerta"
            FEExoneraciones.lbEditarFlex = False 'FRHU 20160819
            Unload Me
        End If
    Else
        'MsgBox "OcurriÃ³ un error al momento de registrar, consulte a departamento de TI", vbInformation, "Alerta" 'FRHU 20160819
    End If
End Sub

Private Sub cmdAgregar_Click()
    FEExoneraciones.AdicionaFila
    FEExoneraciones.TextMatrix(FEExoneraciones.row, 3) = gsCodUser
    FEExoneraciones.TextMatrix(FEExoneraciones.row, 4) = FEExoneraciones.rows - 1
    FEExoneraciones.TextMatrix(FEExoneraciones.row, 6) = 1 'Asigna como registro nuevo
    'FRHU 20160819
    FEExoneraciones.SetFocus
    SendKeys "{ENTER}"
    FEExoneraciones.lbEditarFlex = True
    'FIN FRHU 20160819
    
End Sub

Private Sub cmdCancelar_Click()
    'If Guardar Then 'FRHU 20160819
    FEExoneraciones.lbEditarFlex = False 'FRHU 20160819
    Unload Me
    'End If 'FRHU 20160819
End Sub

Private Sub cmdQuitar_Click()
    
    FEExoneraciones.EliminaFila (FEExoneraciones.row)
    
End Sub

Private Sub CargarListaExoneraciones(ByVal psCtaCod As String)
    Dim oCredNiv As New COMNCredito.NCOMNivelAprobacion
    Dim oRS As New ADODB.Recordset
    Dim nIndice As Integer
    
    Set oRS = oCredNiv.ListaCredNivExoneracionCta(psCtaCod)
    FEExoneraciones.Clear
    FormateaFlex FEExoneraciones
    
    If Not (oRS.EOF And oRS.BOF) Then
        For nIndice = 1 To oRS.RecordCount
            FEExoneraciones.AdicionaFila
            FEExoneraciones.TextMatrix(nIndice, 1) = oRS!cExoneracion
            FEExoneraciones.TextMatrix(nIndice, 2) = oRS!cDescripcion
            FEExoneraciones.TextMatrix(nIndice, 3) = oRS!cUser
            FEExoneraciones.TextMatrix(nIndice, 4) = nIndice 'oRS!nItem
            FEExoneraciones.TextMatrix(nIndice, 5) = oRS!cmovnroreg
            FEExoneraciones.TextMatrix(nIndice, 6) = 2 'Asigna como registro existente
            FEExoneraciones.TextMatrix(nIndice, 7) = oRS!nId
            oRS.MoveNext
        Next
        fbNuevoRegistro = False
        fbExononeracion = True
    End If
End Sub

Private Function Guardar() As Boolean
    Dim oCredNiv As New COMNCredito.NCOMNivelAprobacion
    Dim nIdRegistro As Long
    Dim nIndice As Integer
    Dim sMovNro As String, sMjs As String
    
    Dim obDCredNiv As New COMDCredito.DCOMNivelAprobacion 'JOEP20171120
    Set obDCredNiv = New COMDCredito.DCOMNivelAprobacion 'JOEP20171120
    
    Call LlenarDescripcionAutorizacion 'JOEP20210930 Adecuacion al Reglamento
     
    Guardar = False
    sMjs = ValidaDatos
    'FRHU 20160819
    If sMjs <> "" Then
        MsgBox sMjs, vbInformation, "Alerta"
        Exit Function
    End If
    'FIN FRHU 20160819
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    'If fbNuevoRegistro = True And sMjs = "" Then
    If fbNuevoRegistro = True Then 'FRHU 20160819
        'FRHU 20160819
        'If sMjs <> "" Then
            'MsgBox sMjs, vbInformation, "Alerta"
            'Exit Function
        'End If
        'FIN FRHU 20160819
        nIdRegistro = oCredNiv.RegistrarCredNivExoneraCabecera(fsCtaCod, EstadoAutoExonera.gEstadoPendiente, sMovNro)
    Else
          
        nIdRegistro = IIf(FEExoneraciones.TextMatrix(1, 7) = "", 0, FEExoneraciones.TextMatrix(1, 7))
        
    End If
    
    If nIdRegistro > 0 Then
        For nIndice = 1 To FEExoneraciones.rows - 1
            'Call oCredNiv.RegistrarCredNivExoneraDetalle(nIdRegistro, feExoneraciones.TextMatrix(nIndice, 0), Replace(feExoneraciones.TextMatrix(nIndice, 1), "'", "''"), Replace(feExoneraciones.TextMatrix(nIndice, 2), "'", "''") _
             '                                           , fsNivAprCod, 1, EstadoAutoExonera.gEstadoPendiente, IIf(feExoneraciones.TextMatrix(nIndice, 6) = 1, sMovNro, feExoneraciones.TextMatrix(nIndice, 5)))
            'JOEP20210930 Adecuacion al Reglamento
            Call oCredNiv.RegistrarCredNivExoneraDetalle(nIdRegistro, FEExoneraciones.TextMatrix(nIndice, 0), Right(Replace(FEExoneraciones.TextMatrix(nIndice, 1), "'", "''"), 7), Trim(Replace(FEExoneraciones.TextMatrix(nIndice, 2), "'", "''")) _
                                                        , fsNivAprCod, 1, EstadoAutoExonera.gEstadoPendiente, IIf(FEExoneraciones.TextMatrix(nIndice, 6) = 1, sMovNro, FEExoneraciones.TextMatrix(nIndice, 5)))
            'JOEP20210930 Adecuacion al Reglamento
        Next
        fbExononeracion = True
        nExito = 1 'JOEP
    Else 'JOEP
        Call obDCredNiv.EliminaCredNivExonera(fsCtaCod, 0) 'JOEP20171120
        fbExononeracion = False 'JOEP20171120
        nExito = 2 'JOEP20171120
    End If
    
    Guardar = True
    Set obDCredNiv = Nothing 'JOEP20171120
End Function

Private Sub CargaDatosNivel()
    Dim oCons As New COMDConstSistema.DCOMGeneral
    fsNivAprCod = oCons.LeeConstSistema(527)
End Sub

Private Function ValidaDatos() As String
    Dim nIndice As Integer
    Dim i As Long
    Dim J As Long
    ValidaDatos = ""
If fbNuevoRegistro = True Then 'JOEP20171120
    For nIndice = 1 To FEExoneraciones.rows - 1
        If FEExoneraciones.TextMatrix(nIndice, 1) = "" Or FEExoneraciones.TextMatrix(nIndice, 2) = "" Or FEExoneraciones.TextMatrix(nIndice, 3) = "" Then
            'ValidaDatos = "Los datos no pueden ser vacios"
            ValidaDatos = "Los datos no pueden ser vacios en la fila (" & nIndice & ")" 'FRHU 20160819
            Exit Function
        End If
    Next
    'JOEP20210930 Adecuacion al Reglamento
    'Verificar duplicidad de autorizaciones
    For i = 1 To FEExoneraciones.rows - 1
        For J = 1 To FEExoneraciones.rows - 1
            If FEExoneraciones.TextMatrix(i, 1) = FEExoneraciones.TextMatrix(J, 1) And i <> J Then
                ValidaDatos = "Autorizaciones Duplicados"
                Exit Function
            End If
        Next J
    Next i
    'JOEP20210930 Adecuacion al Reglamento
Else
        If (FEExoneraciones.TextMatrix(nIndice, 1) = "" And FEExoneraciones.TextMatrix(nIndice, 2) = "" And FEExoneraciones.TextMatrix(nIndice, 3) = "") Then
        Else
            For nIndice = 1 To FEExoneraciones.rows - 1
                If FEExoneraciones.TextMatrix(nIndice, 1) = "" Or FEExoneraciones.TextMatrix(nIndice, 2) = "" Or FEExoneraciones.TextMatrix(nIndice, 3) = "" Then
                ValidaDatos = "Los datos no pueden ser vacios en la fila (" & nIndice & ")"
                End If
            Next
        End If
End If 'JOEP20171120
End Function

'JOEP20210930 Adecuacion al Reglamento
Private Sub CargaAutorizaciones()
    Dim objAut As New COMDCredito.DCOMNivelAprobacion
    Dim rsAut As ADODB.Recordset
    Set objAut = New COMDCredito.DCOMNivelAprobacion
    
    Set rsAut = objAut.ObtieneAutoNoContempladas
    
    If Not (rsAut.BOF And rsAut.EOF) Then
         FEExoneraciones.CargaCombo rsAut
    End If
    
    Set objAut = Nothing
    RSClose rsAut
End Sub

Private Sub feExoneraciones_DblClick()
    Call CargaAutorizaciones
End Sub

Private Sub LlenarDescripcionAutorizacion()
Dim i As Integer
    For i = 1 To FEExoneraciones.rows - 1
        FEExoneraciones.TextMatrix(i, 2) = Left(FEExoneraciones.TextMatrix(i, 1), 100)
    Next i
End Sub
'JOEP20210930 Adecuacion al Reglamento
