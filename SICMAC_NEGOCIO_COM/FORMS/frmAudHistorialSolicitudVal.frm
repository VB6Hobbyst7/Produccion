VERSION 5.00
Begin VB.Form frmAudHistorialSolicitudVal 
   Caption         =   "Historial de Solicitud de Validacion de Procedimiento"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "frmAudHistorialSolicitudVal.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frActividad 
      Caption         =   "Actividad"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   5520
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin SICMACT.FlexEdit grdHistorial 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6255
         _extentx        =   11033
         _extenty        =   3201
         cols0           =   6
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Comentarios-Conclusiones-Fec.Registro-Estado-Motivo Rechazo"
         encabezadosanchos=   "450-3500-3500-1200-1200-3800"
         font            =   "frmAudHistorialSolicitudVal.frx":030A
         font            =   "frmAudHistorialSolicitudVal.frx":0336
         font            =   "frmAudHistorialSolicitudVal.frx":0362
         font            =   "frmAudHistorialSolicitudVal.frx":038E
         font            =   "frmAudHistorialSolicitudVal.frx":03BA
         fontfixed       =   "frmAudHistorialSolicitudVal.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-C-C-L"
         formatosedit    =   "0-0-0-0-0-0"
         textarray0      =   "#"
         colwidth0       =   450
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label lblProcedimiento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Procedimiento:"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAudHistorialSolicitudVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Public pnProcedimientoID As Integer
'
'Private Sub cmdsalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Dim objCOMNAudio As COMNAuditoria.NCOMRegistros
'    Set objCOMNAudio = New COMNAuditoria.NCOMRegistros
'
'    'pnProcedimientoID = frmAudSolicitarValidacion.npProcedimientoID
'    If CargarGrilla(objCOMNAudio.HistorialSolicitudValidacionProcedimento(pnProcedimientoID)) = False Then
'        MsgBox "El procedimiento no tiene un historial", vbInformation, "Aviso"
'    End If
'End Sub
'
'Public Function CargarGrilla(ByVal DR As ADODB.Recordset) As Boolean
'    Dim i As Integer
'    grdHistorial.Clear
'    grdHistorial.FormaCabecera
'    grdHistorial.AdicionaFila
'    CargarGrilla = False
'    Do Until DR.EOF
'        i = i + 1
'        frActividad.Caption = DR!cActividadDesc
'        lblProcedimiento.Caption = DR!cProcedimientoNombre
'        grdHistorial.TextMatrix(i, 1) = DR!cComentario
'        grdHistorial.TextMatrix(i, 2) = DR!cConclusion
'        grdHistorial.TextMatrix(i, 3) = DR!dFechaRegistro
'        grdHistorial.TextMatrix(i, 4) = DR!cConsDescripcion
'        grdHistorial.TextMatrix(i, 5) = DR!cMotivoRechazo
'        grdHistorial.AdicionaFila
'        DR.MoveNext
'        CargarGrilla = True
'    Loop
'End Function
