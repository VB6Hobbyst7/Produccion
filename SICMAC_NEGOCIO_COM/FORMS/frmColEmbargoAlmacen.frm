VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmColEmbargoAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacenes"
   ClientHeight    =   2640
   ClientLeft      =   6540
   ClientTop       =   5100
   ClientWidth     =   6240
   Icon            =   "frmColEmbargoAlmacen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   6015
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin MSDataListLib.DataCombo dcAlmacen 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   570
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txtAlmacen 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   585
         Visible         =   0   'False
         Width           =   4815
      End
      Begin MSDataListLib.DataCombo dcAgencia 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   210
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmColEmbargoAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lsAlmacen As String
'Dim lsCodAlmacen As String
'Private Sub cmdAceptar_Click()
'    If dcAlmacen.BoundText <> "0" And dcAlmacen.BoundText <> "" Then
'        lsAlmacen = dcAlmacen.Text + Space(100) + dcAlmacen.BoundText
'        Unload Me
'    Else
'        MsgBox "Seleccione un Almacen", vbExclamation, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdCancelar_Click()
'    Me.dcAgencia.Enabled = True
'    dcAlmacen.Visible = True
'    cmdNuevo.Visible = True
'    cmdAceptar.Visible = True
'
'    cmdGrabar.Visible = False
'    cmdCancelar.Visible = False
'    cmdModificar.Visible = False
'    Me.txtAlmacen.Visible = False
'     Me.txtDescripcion.Enabled = False
'
'     Me.txtAlmacen.Text = ""
'    Me.txtDescripcion.Text = ""
'End Sub
'
'Private Sub cmdGrabar_Click()
'    If Not valida Then
'       Exit Sub
'    End If
'    Dim rsAlmacen As Recordset
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'    Dim sMovNro As String
'
'    If MsgBox("Seguro de Registrar los Datos", vbYesNo, "Aviso") = vbYes Then
'        Dim clsMov As COMNContabilidad.NCOMContFunciones
'        Dim sCorrelativo As String
'
'        Set clsMov = New COMNContabilidad.NCOMContFunciones
'
'        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'        If lsCodAlmacen = "" Or lsCodAlmacen = "0" Then
'            Set oColRec = New COMNColocRec.NCOMColRecCredito
'            sCorrelativo = oColRec.ObtenerAlmacenCorrelativo(dcAgencia.BoundText)
'            oColRec.guardarAlmacenEmbargo dcAgencia.BoundText + sCorrelativo, Me.txtAlmacen.Text, Me.txtDescripcion, sMovNro
'            lsAlmacen = Me.txtAlmacen.Text + Space(100) + dcAgencia.BoundText + sCorrelativo
'        Else
'            Set oColRec = New COMNColocRec.NCOMColRecCredito
'            oColRec.modificarAlmacenEmbargo lsCodAlmacen
'            oColRec.guardarAlmacenEmbargo lsCodAlmacen, Me.txtAlmacen.Text, Me.txtDescripcion, sMovNro
'            lsAlmacen = Me.txtAlmacen.Text + Space(100) + lsCodAlmacen
'        End If
'
'        MsgBox "Se Guardaron los datos del Almacen"
'
'        Unload Me
'    End If
'End Sub
'Private Function valida() As Boolean
'       valida = True
'        If dcAgencia.BoundText = "0" Then
'            MsgBox "Debe seleccionar una Agencia", vbInformation, "Aviso"
'            valida = False
'            Exit Function
'        End If
'        If Me.txtAlmacen.Text = "" Then
'            MsgBox "Debe Ingresar el Nombre del Almacen", vbInformation, "Aviso"
'            valida = False
'            Exit Function
'        End If
'End Function
'
'Private Sub cmdModificar_Click()
'    Me.dcAgencia.Enabled = False
'    Me.dcAlmacen.Visible = False
'    Me.cmdNuevo.Visible = False
'    Me.cmdAceptar.Visible = False
'    Me.cmdModificar.Visible = False
'
'    Me.cmdGrabar.Visible = True
'    Me.cmdCancelar.Visible = True
'    Me.txtAlmacen.Visible = True
'    Me.txtAlmacen.Text = Me.dcAlmacen.Text
'    Me.txtDescripcion.Enabled = True
'    lsCodAlmacen = dcAlmacen.BoundText
'    Me.txtAlmacen.SetFocus
'
'End Sub
'
'Private Sub cmdNuevo_Click()
'    dcAlmacen.Visible = False
'    cmdNuevo.Visible = False
'    cmdAceptar.Visible = False
'    cmdModificar.Visible = False
'
'    cmdGrabar.Visible = True
'    cmdCancelar.Visible = True
'    Me.txtAlmacen.Visible = True
'    Me.txtDescripcion.Enabled = True
'
'    Me.txtAlmacen.Text = ""
'    Me.txtDescripcion.Text = ""
'
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub dcAgencia_Click(Area As Integer)
'
'    cargarAlmacen
'    cmdModificar.Visible = False
'End Sub
'Private Sub cargarAlmacen()
'    Dim rsAlmacen As Recordset
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'
'
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsAlmacen = oColRec.ObtenerAlmacenEmbargo(dcAgencia.BoundText, "0")
'
'        dcAlmacen.BoundColumn = "cCodAlmacen"
'        dcAlmacen.DataField = "cCodAlmacen"
'        Set dcAlmacen.RowSource = rsAlmacen
'        dcAlmacen.ListField = "cNomAlmacen"
'        dcAlmacen.BoundText = 0
'       Set rsAlmacen = Nothing
'       Me.txtDescripcion.Text = ""
'
'End Sub
'Private Sub dcAgencia_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.dcAlmacen.SetFocus
'    End If
'End Sub
'
'Private Sub dcAlmacen_Change()
'    'ObtenerDetalleAlmacen
'End Sub
'Private Sub ObtenerDetalleAlmacen()
'    Dim rsAlmacen As Recordset
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsAlmacen = oColRec.ObtenerAlmacenEmbargo(dcAgencia.BoundText, dcAlmacen.BoundText)
'    If Not (rsAlmacen.BOF And rsAlmacen.EOF) Then
'        Me.txtDescripcion.Text = rsAlmacen!cDescripcion
'       Set rsAlmacen = Nothing
'       cmdModificar.Visible = True
'    End If
'End Sub
'Public Function Inicio(ByVal psCodAlmacen As String) As String
'    lsCodAlmacen = Trim(psCodAlmacen)
'    If lsCodAlmacen = "0" Or lsCodAlmacen = "" Then
'        lsAlmacen = ""
'    End If
'    Me.Show 1
'    If Trim(Right(lsAlmacen, 5)) = "0" Or lsAlmacen = "" Then
'        lsAlmacen = "[Ingrese Almacen]" + Space(100) + "0"
'    End If
'    Inicio = lsAlmacen
'End Function
'
'Private Sub dcAlmacen_Click(Area As Integer)
'    ObtenerDetalleAlmacen
'End Sub
'
'Private Sub dcAlmacen_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdAceptar.SetFocus
'    End If
'End Sub
'
'Private Sub Form_Load()
'
'    CargarAgencias
'    If lsCodAlmacen <> "0" And lsCodAlmacen <> "" Then
'        Me.dcAgencia.BoundText = Left(lsCodAlmacen, 2)
'        cargarAlmacen
'        Me.dcAlmacen.BoundText = lsCodAlmacen
'    End If
'End Sub
'Private Sub CargarAgencias()
'    Dim rsAgencia As New ADODB.Recordset
'    Dim objCOMNCredito As COMNCredito.NCOMBPPR
'    Set objCOMNCredito = New COMNCredito.NCOMBPPR
'    Set rsAgencia.DataSource = objCOMNCredito.getCargarAgencias
'    dcAgencia.BoundColumn = "cAgeCod"
'    dcAgencia.DataField = "cAgeCod"
'    Set dcAgencia.RowSource = rsAgencia
'    dcAgencia.ListField = "cAgeDescripcion"
'    dcAgencia.BoundText = 0
'End Sub
'Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.txtDescripcion.SetFocus
'    End If
'End Sub
'Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then
'        Me.cmdGrabar.SetFocus
'    End If
'End Sub
