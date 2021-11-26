VERSION 5.00
Begin VB.Form frmCredCargosTasa 
   Caption         =   "Aprobación de Solicitud de Tasa"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18165
   Icon            =   "frmCredCargosTasa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   18165
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.FlexEdit FEAprobacion 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17975
      _ExtentX        =   31697
      _ExtentY        =   4260
      Cols0           =   11
      HighLight       =   1
      EncabezadosNombres=   "CodCargo-Agencia-Tipo Crédito-Crédito-Cliente-Movimiento-Tasa Sol.-Tasa Apr.-Estado-cPersCod-cTpoProdCod"
      EncabezadosAnchos=   "0-2000-2500-2000-4000-2500-1200-1200-2000-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-7-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-L-L-C-R-R-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0"
      CantDecimales   =   4
      TextArray0      =   "CodCargo"
      lbEditarFlex    =   -1  'True
      RowHeight0      =   300
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   13920
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   375
      Left            =   15240
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "Aprobar"
      Height          =   375
      Left            =   16680
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "frmCredCargosTasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnCantidad As Integer
'MACM 20210406 PIGNORATICIO
Private Sub cmdActualizar_Click()
    Call mostrarFlex
End Sub
Private Sub cmdAnular_Click()
    If Len(Trim(FEAprobacion.TextMatrix(FEAprobacion.row, 3))) <> "18" Then
        MsgBox "Favor, debe seleccionar la petición de aprobación de tasa", vbInformation, "Aviso!"
        Exit Sub
    End If
    Dim sMovNroA As String
    Dim objCargos As COMDCredito.DCOMNivelAprobacion
    Dim ClsMov As COMNContabilidad.NCOMContFunciones
    Set objCargos = New COMDCredito.DCOMNivelAprobacion
    Set ClsMov = New COMNContabilidad.NCOMContFunciones
    sMovNroA = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call objCargos.GetActualizarColocacPermisoAprobacion(FEAprobacion.TextMatrix(FEAprobacion.row, 3), FEAprobacion.TextMatrix(FEAprobacion.row, 5), sMovNroA, FEAprobacion.TextMatrix(FEAprobacion.row, 9), gsCodPersUser, gsCodCargo, FEAprobacion.TextMatrix(FEAprobacion.row, 6), FEAprobacion.TextMatrix(FEAprobacion.row, 7), 0, 0)
    'Call objCargos.GetActualizarColocacPermisoAprobacion(FEAprobacion.TextMatrix(FEAprobacion.row, 3), FEAprobacion.TextMatrix(FEAprobacion.row, 5), sMovNroA, FEAprobacion.TextMatrix(FEAprobacion.row, 9), gsCodPersUser, gsCodCargo, FEAprobacion.TextMatrix(FEAprobacion.row, 6), FEAprobacion.TextMatrix(FEAprobacion.row, 7), 2, 2)
    MsgBox "La petición de tasa se eliminó correctamente ", vbInformation, "Aviso!"
    Call mostrarFlex
End Sub

Private Sub cmdAprobar_Click()
    If Len(Trim(FEAprobacion.TextMatrix(FEAprobacion.row, 3))) <> "18" Then
        MsgBox "Favor, debe seleccionar la petición de aprobación de tasa", vbInformation, "Aviso!"
    Exit Sub
    End If
    Dim lnTpoCred As Integer
    Dim sMovNroA As String
    Dim objCargos As COMDCredito.DCOMNivelAprobacion
    Dim ClsMov As COMNContabilidad.NCOMContFunciones
    
    Set objCargos = New COMDCredito.DCOMNivelAprobacion
    Set ClsMov = New COMNContabilidad.NCOMContFunciones
    sMovNroA = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                                                                                                                                        
    Call objCargos.GetActualizarColocacPermisoAprobacion(FEAprobacion.TextMatrix(FEAprobacion.row, 3), FEAprobacion.TextMatrix(FEAprobacion.row, 5), sMovNroA, FEAprobacion.TextMatrix(FEAprobacion.row, 9), gsCodPersUser, gsCodCargo, FEAprobacion.TextMatrix(FEAprobacion.row, 6), FEAprobacion.TextMatrix(FEAprobacion.row, 7), 2, 2)
    'MACM 20210406 PIGNORATICIO
    lnTpoCred = FEAprobacion.TextMatrix(FEAprobacion.row, 10)
    If lnTpoCred = 709 Or lnTpoCred = 705 Then
        Call objCargos.ActualizaTasaAprobadaPigno(FEAprobacion.TextMatrix(FEAprobacion.row, 3), FEAprobacion.TextMatrix(FEAprobacion.row, 7))
    End If
    MsgBox "La petición de tasa se aprobó correctamente ", vbInformation, "Aviso!"
    Call mostrarFlex
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub
'MACM 20210604 PIGNORATICIO
Private Sub FEAprobacion_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim nRaros As String
    Dim nTexto As String
    If pnCol = 7 Then
        Cancel = True
        If FEAprobacion.TextMatrix(FEAprobacion.row, 7) = "0.00" Then
            MsgBox "La tasa de aprobación no puede ser cero", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}"
        ElseIf FEAprobacion.TextMatrix(FEAprobacion.row, 7) = "." Then
            MsgBox "Por favor ingrese un numero correcto", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}"
        End If
        nRaros = "-"
        nTexto = FEAprobacion.TextMatrix(FEAprobacion.row, 7)
        
        If InStr(nTexto, nRaros) > 0 Then
            MsgBox "Por favor ingrese un numero correcto", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}"
        End If
    Else
        Cancel = False
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'controlando el Ctrl + V
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub
Private Sub Form_Load()
    Call mostrarFlex
End Sub
Private Sub mostrarFlex()
Dim ObjOCredNivel As COMDCredito.DCOMNivelAprobacion
Set ObjOCredNivel = New COMDCredito.DCOMNivelAprobacion
Dim objRS As ADODB.Recordset
Set objRS = New ADODB.Recordset
FormateaFlex FEAprobacion
lnCantidad = 0
Set objRS = ObjOCredNivel.GetCargarColocacPermisoAprobacion(Format(gdFecSis, "YYYYMMDD"), gsCodCargo)
If objRS.EOF Or objRS.BOF Then
Exit Sub
End If
Do While Not objRS.EOF
    lnCantidad = lnCantidad + 1
    FEAprobacion.AdicionaFila
    FEAprobacion.TextMatrix(objRS.Bookmark, 0) = gsCodCargo
    FEAprobacion.TextMatrix(objRS.Bookmark, 1) = objRS!cAgeDescripcion
    FEAprobacion.TextMatrix(objRS.Bookmark, 2) = objRS!cTpoCred
    FEAprobacion.TextMatrix(objRS.Bookmark, 3) = objRS!cCtaCod
    FEAprobacion.TextMatrix(objRS.Bookmark, 4) = objRS!cPersNombre
    FEAprobacion.TextMatrix(objRS.Bookmark, 5) = objRS!cMovPermisoSolicitante
    If objRS!cCtaCod = "" Then
        FEAprobacion.TextMatrix(objRS.Bookmark, 6) = ""
        FEAprobacion.TextMatrix(objRS.Bookmark, 7) = ""
        FEAprobacion.TextMatrix(objRS.Bookmark, 8) = ""
    Else
        FEAprobacion.TextMatrix(objRS.Bookmark, 6) = objRS!nTasaSol
        FEAprobacion.TextMatrix(objRS.Bookmark, 7) = objRS!nTasaApr
        FEAprobacion.TextMatrix(objRS.Bookmark, 8) = "Pendiente"
    End If
    
    
    FEAprobacion.TextMatrix(objRS.Bookmark, 9) = objRS!cPersCod
    FEAprobacion.TextMatrix(objRS.Bookmark, 10) = objRS!cTpoProdCod
    objRS.MoveNext
Loop
End Sub
