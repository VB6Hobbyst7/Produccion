VERSION 5.00
Begin VB.Form frmAudListaActividades 
   Caption         =   "Lista Actividades"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   Icon            =   "frmAudListaActividades.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   6960
      TabIndex        =   2
      Top             =   2760
      Width           =   750
   End
   Begin VB.CommandButton cmdElegir 
      Caption         =   "Elegir"
      Height          =   315
      Left            =   6000
      TabIndex        =   1
      Top             =   2760
      Width           =   750
   End
   Begin SICMACT.FlexEdit grdListaActividades 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _extentx        =   13785
      _extenty        =   4683
      cols0           =   4
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Codigo Actividad-Descripción-Asignado a"
      encabezadosanchos=   "240-2000-3000-3500"
      font            =   "frmAudListaActividades.frx":030A
      font            =   "frmAudListaActividades.frx":0336
      font            =   "frmAudListaActividades.frx":0362
      font            =   "frmAudListaActividades.frx":038E
      font            =   "frmAudListaActividades.frx":03BA
      fontfixed       =   "frmAudListaActividades.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-X"
      listacontroles  =   "0-0-0-0"
      encabezadosalineacion=   "C-C-C-C"
      formatosedit    =   "0-0-0-0"
      textarray0      =   "#"
      selectionmode   =   1
      colwidth0       =   240
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmAudListaActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public psCodActividad As String
'Public psActividadDesc As String
'Private Sub cmdElegir_Click()
'    psCodActividad = grdListaActividades.TextMatrix(grdListaActividades.row, grdListaActividades.Col)
'    psActividadDesc = grdListaActividades.TextMatrix(grdListaActividades.row, 2)
'    Unload Me
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'  Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'  Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'  If VerificarGrupo = True Then
'    grdListaActividades.rsFlex = objCOMNAuditoria.AuditDatosActividadesXUser(gsCodPersUser, 1)
'  Else
'    grdListaActividades.rsFlex = objCOMNAuditoria.AuditDatosActividadesXUser(gsCodPersUser, 2)
'  End If
'
'End Sub
'
'Public Function VerificarGrupo() As Boolean
'    Dim lsGrupoCompara As String
'    lsGrupoCompara = "GRUPO JEFE AUDITORIA"
'    Dim lsGrupo As String
'    For x = 1 To Len(gsGruposUser)
'        If Mid(gsGruposUser, x, 1) <> "," Then
'            lsGrupo = lsGrupo & Mid(gsGruposUser, x, 1)
'        Else
'            If lsGrupoCompara = lsGrupo Then
'                VerificarGrupo = True
'                Exit Function
'            End If
'            lsGrupo = ""
'        End If
'        If x = Len(gsGruposUser) Then
'            If lsGrupoCompara = lsGrupo Then
'                VerificarGrupo = True
'                Exit Function
'            End If
'        End If
'    Next
'End Function
