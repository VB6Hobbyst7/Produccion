VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredExoneraNCNew 
   Caption         =   "Exoneraciones no contempladas"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   Icon            =   "frmCredExoneraNCNew.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   8055
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
      EncabezadosNombres=   "#-Exoneración-Descripción Exoneración-Usuario-ExoneraItem-cMovNro-TipoRegistro-nId"
      EncabezadosAnchos=   "300-2500-3500-1000-0-0-0-0"
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
      ColumnasAEditar =   "X-1-2-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbEditarFlex    =   -1  'True
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
            Caption         =   "Exoneraciones Solicitadas"
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

Public Function Inicia(ByVal psCtaCod As String) As Boolean
    fbExononeracion = False
    fsCtaCod = psCtaCod
    fbNuevoRegistro = True
    Call CargaDatosNivel
    Call CargarListaExoneraciones(psCtaCod)
    Me.Show 1
    Inicia = fbExononeracion
End Function

Private Sub cmdAceptar_Click()
    If Guardar Then
        MsgBox "Los datos se guardaron de forma correcta", vbInformation, "Alerta"
        Unload Me
    Else
        MsgBox "Ocurrió un error al momento de registrar, consulte a departamento de TI", vbInformation, "Alerta"
    End If
End Sub

Private Sub cmdAgregar_Click()
    feExoneraciones.AdicionaFila
    feExoneraciones.TextMatrix(feExoneraciones.row, 3) = gsCodUser
    feExoneraciones.TextMatrix(feExoneraciones.row, 4) = feExoneraciones.Rows - 1
    feExoneraciones.TextMatrix(feExoneraciones.row, 6) = 1 'Asigna como registro nuevo
End Sub

Private Sub cmdCancelar_Click()
    If Guardar Then
        Unload Me
    End If
End Sub

Private Sub cmdQuitar_Click()
    feExoneraciones.EliminaFila (feExoneraciones.row)
End Sub

Private Sub CargarListaExoneraciones(ByVal psCtaCod As String)
    Dim oCredNiv As New COMNCredito.NCOMNivelAprobacion
    Dim ors As New ADODB.Recordset
    Dim nIndice As Integer
    
    Set ors = oCredNiv.ListaCredNivExoneracionCta(psCtaCod)
    feExoneraciones.Clear
    FormateaFlex feExoneraciones
    
    If Not (ors.EOF And ors.BOF) Then
        For nIndice = 1 To ors.RecordCount
            feExoneraciones.AdicionaFila
            feExoneraciones.TextMatrix(nIndice, 1) = ors!cExoneracion
            feExoneraciones.TextMatrix(nIndice, 2) = ors!cDescripcion
            feExoneraciones.TextMatrix(nIndice, 3) = ors!cUser
            feExoneraciones.TextMatrix(nIndice, 4) = nIndice 'oRS!nItem
            feExoneraciones.TextMatrix(nIndice, 5) = ors!cMovNroReg
            feExoneraciones.TextMatrix(nIndice, 6) = 2 'Asigna como registro existente
            feExoneraciones.TextMatrix(nIndice, 7) = ors!nId
            ors.MoveNext
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
    
    Guardar = False
    sMjs = ValidaDatos
   
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If fbNuevoRegistro = True And sMjs = "" Then
        
        If sMjs <> "" Then
            MsgBox sMjs, vbInformation, "Alerta"
            Exit Function
        End If
        
        nIdRegistro = oCredNiv.RegistrarCredNivExoneraCabecera(fsCtaCod, EstadoAutoExonera.gEstadoPendiente, sMovNro)
    Else
        nIdRegistro = IIf(feExoneraciones.TextMatrix(1, 7) = "", 0, feExoneraciones.TextMatrix(1, 7))
    End If
    
    If nIdRegistro > 0 Then
        For nIndice = 1 To feExoneraciones.Rows - 1
            Call oCredNiv.RegistrarCredNivExoneraDetalle(nIdRegistro, feExoneraciones.TextMatrix(nIndice, 4), feExoneraciones.TextMatrix(nIndice, 1), feExoneraciones.TextMatrix(nIndice, 1) _
                                                        , fsNivAprCod, 1, EstadoAutoExonera.gEstadoPendiente, IIf(feExoneraciones.TextMatrix(nIndice, 6) = 1, sMovNro, feExoneraciones.TextMatrix(nIndice, 5)))
        Next
        fbExononeracion = True
    End If
    Guardar = True
End Function

Private Sub CargaDatosNivel()
    Dim oCons As New COMDConstSistema.DCOMGeneral
    fsNivAprCod = oCons.LeeConstSistema(527)
End Sub

Private Function ValidaDatos() As String
    Dim nIndice As Integer
    ValidaDatos = ""
    For nIndice = 1 To feExoneraciones.Rows - 1
        If feExoneraciones.TextMatrix(nIndice, 1) = "" Or feExoneraciones.TextMatrix(nIndice, 2) = "" Or feExoneraciones.TextMatrix(nIndice, 3) = "" Then
            ValidaDatos = "Los datos no pueden ser vacios"
        End If
    Next
End Function
