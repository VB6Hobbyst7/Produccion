VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAuditoriaRegistroSeguimiento 
   Caption         =   "Registro de Medidas correctivas"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   Icon            =   "frmAuditoriaRegistroSeguimiento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin MSDataListLib.DataCombo dcTipoMedida 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         Height          =   5655
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   5055
         Begin VB.TextBox txtArea 
            Height          =   405
            Left            =   1680
            TabIndex        =   22
            Top             =   4080
            Width           =   3255
         End
         Begin MSDataListLib.DataCombo dcEnte 
            Height          =   315
            Left            =   2520
            TabIndex        =   20
            Top             =   4560
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.TextBox txtSituacionComentario 
            Height          =   735
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   3240
            Width           =   3255
         End
         Begin VB.TextBox txtObservacion 
            Height          =   735
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtAccion 
            Height          =   735
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   1920
            Width           =   3255
         End
         Begin VB.ComboBox cmbSituacion 
            Height          =   315
            ItemData        =   "frmAuditoriaRegistroSeguimiento.frx":030A
            Left            =   3360
            List            =   "frmAuditoriaRegistroSeguimiento.frx":0317
            TabIndex        =   9
            Top             =   2760
            Width           =   1575
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   350
            Left            =   3960
            TabIndex        =   8
            Top             =   5040
            Width           =   975
         End
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "Actualizar"
            Height          =   350
            Left            =   2760
            TabIndex        =   7
            Top             =   5040
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtRecomendacion 
            Height          =   705
            Left            =   1680
            TabIndex        =   6
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label Label7 
            Caption         =   "Area:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   4200
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Situación a la fecha:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Recomendaciones:"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Acción correctiva:"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Ente Emisor"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   4560
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Observaciones:"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   3960
         TabIndex        =   4
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   78446593
         CurrentDate     =   40234
      End
      Begin VB.TextBox txtNroInforme 
         Height          =   350
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Emisión:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Informe Nº:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAuditoriaRegistroSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim objCOMNAuditoria As COMNAuditoria.NCOMSeguimiento
'
'Private Sub cmdAceptar_Click()
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
'
'    If MsgBox("Esta Seguro de Registar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'
'        If gNroInformeId = 0 Then
'            gNroInformeId = CInt(objCOMNAuditoria.InsertarInformeMed(txtNroInforme.Text, DTPicker1.value, dcTipoMedida.BoundText))
'            objCOMNAuditoria.InsertarMedidasCorrectivas gNroInformeId, txtObservacion.Text, txtRecomendacion.Text, txtAccion.Text, cmbSituacion.Text, txtSituacionComentario.Text, dcEnte.BoundText, txtArea.Text
'        Else
'            objCOMNAuditoria.InsertarMedidasCorrectivas gNroInformeId, txtObservacion.Text, txtRecomendacion.Text, txtAccion.Text, cmbSituacion.Text, txtSituacionComentario.Text, dcEnte.BoundText, txtArea.Text
'        End If
'
'        MsgBox "Los Datos se registraron correctamente", vbInformation
'
'        If MsgBox("Desea Seguir Registrando Medidas?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'            Limpiar
'        Else
'            gNroInformeId = 0
'            txtNroInforme.Text = ""
'            Limpiar
'        End If
'
'    End If
'End Sub
'
'Private Sub Limpiar()
'    txtObservacion.Text = ""
'    txtRecomendacion.Text = ""
'    txtAccion.Text = ""
'    txtSituacionComentario.Text = ""
'    txtArea.Text = ""
'End Sub
'
'Private Sub cmdActualizar_Click()
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
'
'    If MsgBox("Esta Seguro de Actualizar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'
'        objCOMNAuditoria.ActualizarInformeMed gNroInformeId, txtNroInforme.Text, DTPicker1.value, dcTipoMedida.BoundText
'        objCOMNAuditoria.ActualizarMedidasCorrectivas gNroInformeId, gMedidasCorrectivasId, txtObservacion.Text, txtRecomendacion.Text, txtAccion.Text, cmbSituacion.Text, txtSituacionComentario.Text, dcEnte.BoundText, txtArea.Text
'
'        MsgBox "Los Datos se Actualizaron Correctamente", vbInformation, Me.Caption
'        gNroInformeId = 0
'        txtNroInforme.Text = ""
'        Limpiar
'
'    End If
'End Sub
'
'Private Sub dcTipoMedida_Click(Area As Integer)
'    If dcTipoMedida.BoundText = "1" Then
'        Label7.Visible = True
'        Label9.Visible = True
'        txtArea.Visible = True
'        dcEnte.Visible = True
'    Else
'        Label7.Visible = False
'        Label9.Visible = False
'        txtArea.Visible = False
'        dcEnte.Visible = False
'    End If
'End Sub
'
'Private Sub Form_Load()
'    If gNroInformeId <> 0 Then
'        CargarDatosModificar
'        cmdAceptar.Visible = False
'        cmdActualizar.Visible = True
'    Else
'        DTPicker1.value = Date
'        CargarTipoMedida
'        CargarEnte
'    End If
'End Sub
'
'Private Sub CargarTipoMedida()
'    Dim rsTipoMedida As ADODB.Recordset
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
'    Set rsTipoMedida = objCOMNAuditoria.ObtenerTipoMedida
'    Set dcTipoMedida.RowSource = rsTipoMedida
'    dcTipoMedida.BoundColumn = "iTipoId"
'    dcTipoMedida.ListField = "vTipo"
'    Set objCOMNAuditoria = Nothing
'    Set rsTipoMedida = Nothing
'    dcTipoMedida.BoundText = 0
'End Sub
'
'Private Sub CargarEnte()
'    Dim rsEnte As ADODB.Recordset
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
'    Set rsEnte = objCOMNAuditoria.ObtenerEnte
'    Set dcEnte.RowSource = rsEnte
'    dcEnte.BoundColumn = "iEnteId"
'    dcEnte.ListField = "vEnte"
'    Set objCOMNAuditoria = Nothing
'    Set rsEnte = Nothing
'    dcEnte.BoundText = 0
'End Sub
'
'Public Sub CargarDatosModificar()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMSeguimiento
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
'    Dim rs As ADODB.Recordset
'    Set rs = objCOMNAuditoria.ObtenerInformeXId(gNroInformeId, gMedidasCorrectivasId)
'
'    txtNroInforme.Text = rs("vNroInforme")
'    DTPicker1.value = rs("vFechaEmision")
'
'    CargarTipoMedida
'    dcTipoMedida.BoundText = rs("vTipo")
'
'    txtObservacion.Text = rs("tObservacion")
'    txtRecomendacion.Text = rs("tRecomendacion")
'    txtAccion.Text = rs("tAccion")
'    cmbSituacion.Text = rs("vSituacion")
'    txtSituacionComentario.Text = rs("tSituacion")
'    txtArea.Text = rs("vArea")
'
'    CargarEnte
'    dcEnte.BoundText = rs("vEnte")
'
'End Sub
