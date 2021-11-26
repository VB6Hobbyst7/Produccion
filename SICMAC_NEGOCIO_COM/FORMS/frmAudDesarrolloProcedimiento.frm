VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAudDesarrolloProcedimiento 
   Caption         =   "AUDITORIA: DESARROLLO DE PROCEDIMIENTOS"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   Icon            =   "frmAudDesarrolloProcedimiento.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4890
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Procedimientos"
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   7695
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   6360
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdSolicVal 
         Caption         =   "Solicitar Validación"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   1815
      End
      Begin SICMACT.FlexEdit grdProcedimientos 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7455
         _extentx        =   13150
         _extenty        =   4260
         cols0           =   5
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Procedimiento-Descripción-Estado-ID"
         encabezadosanchos=   "300-3400-4500-2200-700"
         font            =   "frmAudDesarrolloProcedimiento.frx":030A
         font            =   "frmAudDesarrolloProcedimiento.frx":0336
         font            =   "frmAudDesarrolloProcedimiento.frx":0362
         font            =   "frmAudDesarrolloProcedimiento.frx":038E
         font            =   "frmAudDesarrolloProcedimiento.frx":03BA
         fontfixed       =   "frmAudDesarrolloProcedimiento.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-C-C"
         formatosedit    =   "0-0-0-0-0"
         textarray0      =   "#"
         selectionmode   =   1
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actividad"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7695
      Begin VB.CommandButton cmdBuscaAct 
         Caption         =   "..."
         Height          =   315
         Left            =   7080
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblActividad 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   7215
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Desarrollo de Procedimiento"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAudDesarrolloProcedimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim sCodActividad As String
'Dim sActividadDesc As String
'Public nProcedimientoID As Integer
'Public sProcedimientoNombre As String
'
'Private Sub cmdBuscaAct_Click()
'    frmAudListaActividades.Show 1
'    sCodActividad = frmAudListaActividades.psCodActividad
'    sActividadDesc = frmAudListaActividades.psActividadDesc
'
'    lblActividad.Caption = sActividadDesc
'    'HabilitarControles
'    CargarGrilla (sCodActividad)
'End Sub
'
'Public Sub CargarGrilla(ByVal sCodigiAct As String)
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'    grdProcedimientos.Clear
'    'grdProcedimientos.Row = 2
'    grdProcedimientos.FormaCabecera
'    grdProcedimientos.rsFlex = objCOMNAuditoria.ListarProcedimientosActividad(sCodigiAct)
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdSolicVal_Click()
'    If grdProcedimientos.TextMatrix(grdProcedimientos.row, 3) = "Aceptado" Then
'        MsgBox "Ya fue aceptado una validación para este procedimiento", vbExclamation, "Aviso"
'        Exit Sub
'    End If
'    If grdProcedimientos.TextMatrix(grdProcedimientos.row, 3) = "Registrado" Or grdProcedimientos.TextMatrix(grdProcedimientos.row, 3) = "Devuelto" Then
'        sProcedimientoNombre = grdProcedimientos.TextMatrix(grdProcedimientos.row, 1)
'        nProcedimientoID = CInt(grdProcedimientos.TextMatrix(grdProcedimientos.row, 4))
'        frmAudSolicitarValidacion.Show 1
'        CargarGrilla (sCodActividad)
'    Else
'        MsgBox "Ya fue solicitada una validación para este procedimiento", vbExclamation, "Aviso"
'    End If
'End Sub
