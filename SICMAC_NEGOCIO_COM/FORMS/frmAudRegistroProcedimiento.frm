VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAudRegistroProcedimiento 
   Caption         =   "Registro de Procedimiento"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   Icon            =   "frmAudRegistroProcedimiento.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7035
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Procedimientos"
      Height          =   5055
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   8895
      Begin VB.CommandButton cmdCerrarAct 
         Caption         =   "Cerrar Actividad"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6840
         TabIndex        =   12
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   7800
         TabIndex        =   11
         Top             =   4560
         Width           =   975
      End
      Begin SICMACT.FlexEdit grdProcedimientos 
         Height          =   3255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   8535
         _extentx        =   15055
         _extenty        =   5741
         cols0           =   5
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Procedimiento-Descripción-Estado-ID"
         encabezadosanchos=   "300-3400-4500-2200-700"
         font            =   "frmAudRegistroProcedimiento.frx":030A
         font            =   "frmAudRegistroProcedimiento.frx":0336
         font            =   "frmAudRegistroProcedimiento.frx":0362
         font            =   "frmAudRegistroProcedimiento.frx":038E
         font            =   "frmAudRegistroProcedimiento.frx":03BA
         fontfixed       =   "frmAudRegistroProcedimiento.frx":03E6
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
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   6255
      End
      Begin VB.TextBox txtProcedimiento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción:"
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Procedimiento:"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actividad"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   8895
      Begin VB.CommandButton cmdBuscaAct 
         Caption         =   "..."
         Height          =   315
         Left            =   8400
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblActividad 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   8535
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11880
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registro de Procedimiento"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAudRegistroProcedimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim sCodActividad As String
'Dim sActividadDesc As String
'
'Private Sub cmdAgregar_Click()
'    If txtProcedimiento.Text <> "" And txtDescripcion.Text <> "" Then
'        If ValidaExisteGrilla(grdProcedimientos, txtProcedimiento.Text) = False Then
'            Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'            Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'            objCOMNAuditoria.AuditRegistrarProcedimientoActividad sCodActividad, txtProcedimiento.Text, txtDescripcion.Text, 1
'            txtProcedimiento.Text = ""
'            txtDescripcion.Text = ""
'            CargarGrilla (sCodActividad)
'        Else
'            MsgBox "Nombre del procedimiento ya existe", vbCritical, "Aviso"
'        End If
'    Else
'        MsgBox "Los datos no pueden ser vacios", vbCritical, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdBuscaAct_Click()
'    frmAudListaActividades.Show 1
'    sCodActividad = frmAudListaActividades.psCodActividad
'    sActividadDesc = frmAudListaActividades.psActividadDesc
'
'    lblActividad.Caption = sActividadDesc
'    HabilitarControles
'    CargarGrilla (sCodActividad)
'End Sub
'
'Public Sub HabilitarControles()
'    txtProcedimiento.Enabled = True
'    txtDescripcion.Enabled = True
'    cmdAgregar.Enabled = True
'    cmdQuitar.Enabled = True
'End Sub
'
'Public Sub desHabilitarControles()
'    txtProcedimiento.Enabled = False
'    txtDescripcion.Enabled = False
'    cmdAgregar.Enabled = False
'    cmdQuitar.Enabled = False
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
'Public Sub LimpiarForm()
'    lblActividad.Caption = ""
'    txtProcedimiento.Text = ""
'    txtDescripcion.Text = ""
'    desHabilitarControles
'End Sub
'
'Private Sub cmdCerrarAct_Click()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'    If VerificarEstadoProcedimiento = True Then
'        objCOMNAuditoria.AuditCerrarActividad (sCodActividad)
'        grdProcedimientos.Clear
'        grdProcedimientos.FormaCabecera
'        grdProcedimientos.Rows = 2
'        'CargarGrilla (sCodActividad)
'    Else
'        MsgBox "Todos los procedimientos deben estar en estado 'Aceptado'", vbCritical, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdQuitar_Click()
'    If grdProcedimientos.TextMatrix(grdProcedimientos.row, grdProcedimientos.Col) = "" Then
'        MsgBox "Dato Vacio. Seleccione opción valida", vbCritical, "Aviso"
'        Exit Sub
'    End If
'    If grdProcedimientos.TextMatrix(grdProcedimientos.row, 3) = "Registrado" Then
'        Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'        Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'        objCOMNAuditoria.QuitarProcedimiento (grdProcedimientos.TextMatrix(grdProcedimientos.row, 4))
'        CargarGrilla (sCodActividad)
'    Else
'        MsgBox "Solo se pueden quitar Procedimientos en esdato:Registrado", vbCritical, "Aviso"
'    End If
'End Sub
'
'Public Function ValidaExisteGrilla(ByVal lFlxElemento As FlexEdit, ByVal sTextoBusca As String) As Boolean
'    For i = 1 To lFlxElemento.Rows - 1
'         If grdProcedimientos.TextMatrix(1, 1) = sTextoBusca Then
'            ValidaExisteGrilla = True
'            Exit Function
'         End If
'    Next
'End Function
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Public Function VerificarEstadoProcedimiento() As Boolean
'    VerificarEstadoProcedimiento = True
'    For x = 1 To grdProcedimientos.Rows - 1
'        If grdProcedimientos.TextMatrix(x, 3) <> "Aceptado" Then
'            VerificarEstadoProcedimiento = False
'            Exit Function
'        End If
'    Next
'End Function
