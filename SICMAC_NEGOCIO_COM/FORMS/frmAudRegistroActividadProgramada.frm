VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAudRegistroActividadProgramada 
   Caption         =   "Registro de Actividades"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmAudRegistroActividadProgramada.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6615
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   7080
      TabIndex        =   19
      Tag             =   "9"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   18
      Tag             =   "10"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "Asignar"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Tag             =   "8"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mantenimiento"
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   9255
      Begin SICMACT.FlexEdit grdActividades 
         Height          =   3135
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   9015
         _extentx        =   15901
         _extenty        =   5530
         cols0           =   7
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Código-Actividad-Norma Legal-Tipo-Origen-Fec. de Reg."
         encabezadosanchos=   "300-1200-2500-1800-1200-3000-1200"
         font            =   "frmAudRegistroActividadProgramada.frx":030A
         font            =   "frmAudRegistroActividadProgramada.frx":0336
         font            =   "frmAudRegistroActividadProgramada.frx":0362
         font            =   "frmAudRegistroActividadProgramada.frx":038E
         font            =   "frmAudRegistroActividadProgramada.frx":03BA
         fontfixed       =   "frmAudRegistroActividadProgramada.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0"
         textarray0      =   "#"
         selectionmode   =   1
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registro"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9255
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Tag             =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8040
         TabIndex        =   15
         Tag             =   "7"
         Top             =   1440
         Width           =   1100
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "Registrar"
         Height          =   375
         Left            =   6840
         TabIndex        =   14
         Tag             =   "6"
         Top             =   1440
         Width           =   1100
      End
      Begin VB.ComboBox cboOrigen 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "5"
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtNormaLegal 
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Tag             =   "4"
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "2"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtActividad 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Tag             =   "1"
         Top             =   600
         Width           =   3135
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   1440
         TabIndex        =   9
         Tag             =   "3"
         Top             =   1320
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Origen:"
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Norma Legal:"
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fec. Registro"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Actividad:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11456
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Registro de Actividades"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAudRegistroActividadProgramada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Public sActividadCod As String
'Public sActividadDesc As String
'Private Sub cboTipo_Change()
'    If cboTipo.ItemData(cboTipo.ListIndex) = 1 Then
'       cboOrigen.Clear
'       cboOrigen.AddItem "" & "Programación Anual"
'       cboOrigen.ItemData(cboOrigen.NewIndex) = "" & 1
'       cboOrigen.ListIndex = 0
'    Else
'        cboOrigen.Clear
'        CargarOrigenActividad
'    End If
'End Sub
'
'Private Sub cboTipo_Click()
'    Call cboTipo_Change
'End Sub
'
'Private Sub cmdAsignar_Click()
'    sActividadCod = grdActividades.TextMatrix(grdActividades.row, grdActividades.Col)
'    If sActividadCod = "" Then
'        MsgBox "Seleccione una opción valida", vbCritical, "Aviso"
'        Exit Sub
'    End If
'    sActividadDesc = grdActividades.TextMatrix(grdActividades.row, 2)
'    frmAudAsignarActividad.Show 1
'    CargarGrid
'End Sub
'
'Private Sub cmdCancelar_Click()
'    Call limpiar
'End Sub
'
'Private Sub cmdQuitar_Click()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'    Dim sCodigo As String
'    If MsgBox("¿Esta seguro que desea quitar la Actividad " & grdActividades.TextMatrix(grdActividades.row, grdActividades.Col) & "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'        sCodigo = grdActividades.TextMatrix(grdActividades.row, grdActividades.Col)
'        objCOMNAuditoria.DarBajaActividad (sCodigo)
'        CargarGrid
'    End If
'End Sub
'
'Private Sub cmdRegistrar_Click()
'    If txtCodigo.Text <> "" And txtActividad.Text <> "" And txtNormaLegal.Text <> "" And txtFecha.Text <> "__/__/____" Then
'        Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'        Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'
'        If VerificarCodigoActividad(txtCodigo.Text) = False Then
'            objCOMNAuditoria.RegistrarActividad txtCodigo.Text, txtActividad.Text, cboTipo.ItemData(cboTipo.ListIndex), _
'                                            txtFecha.Text, txtNormaLegal.Text, cboOrigen.ItemData(cboOrigen.ListIndex)
'        Call limpiar
'        CargarGrid
'        Else
'            MsgBox "El Código de la Actividad ya se encuentra registrado. Por favor utilice otro código", vbExclamation, "Aviso"
'            txtCodigo.Text = ""
'        End If
'    Else
'        MsgBox "Los datos no pueden estar vacios", vbCritical, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    CargarTipoActividad
'    CargarOrigenActividad
'    CargarGrid
'    Call cboTipo_Change
'End Sub
'
'Public Sub CargarTipoActividad()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'
'    Dim lrDatos As ADODB.Recordset
'    Set lrDatos = New ADODB.Recordset
'    Set lrDatos = objCOMNAuditoria.ListarAuditTipoActividad
'
'    Call CargarComboBox(lrDatos, cboTipo)
'End Sub
'
'Public Sub CargarOrigenActividad()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'
'    Dim lrDatos As ADODB.Recordset
'    Set lrDatos = New ADODB.Recordset
'    Set lrDatos = objCOMNAuditoria.ListarAuditOrigenActividad
'
'    Call CargarComboBox(lrDatos, cboOrigen)
'End Sub
'
'Public Sub CargarComboBox(ByVal lrDatos As ADODB.Recordset, ByVal cboControl As ComboBox)
'    Dim lrDatosTmp As ADODB.Recordset
'    Set lrDatosTmp = New ADODB.Recordset
'
'    Do Until lrDatos.EOF
'     cboControl.AddItem "" & lrDatos!cConsDescripcion
'     cboControl.ItemData(cboControl.NewIndex) = "" & lrDatos!nConsValor
'     lrDatos.MoveNext
'    Loop
'    Set lrDatos = Nothing
'
'    cboControl.ListIndex = 0
'End Sub
'
'Public Sub limpiar()
'    txtCodigo.Text = ""
'    txtActividad.Text = ""
'    txtNormaLegal.Text = ""
'    cboTipo.ListIndex = 0
'    cboOrigen.ListIndex = 0
'    txtFecha.Text = "__/__/____"
'End Sub
'
'Public Sub CargarGrid()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'    grdActividades.Clear
'    grdActividades.FormaCabecera
'    grdActividades.rsFlex = objCOMNAuditoria.ObtenerAuditActividades
'End Sub
'
'Public Function VerificarCodigoActividad(ByVal sCodigo As String) As Boolean
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'    VerificarCodigoActividad = False
'    If objCOMNAuditoria.VerificarCodigoActividad(sCodigo).RecordCount > 0 Then
'        VerificarCodigoActividad = True
'    End If
'End Function
'
