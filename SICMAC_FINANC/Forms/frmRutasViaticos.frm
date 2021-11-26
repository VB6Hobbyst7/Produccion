VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRutasViaticos 
   Caption         =   "Rutas / Destinos para Viaticos"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   Icon            =   "frmRutasViaticos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9255
      Begin MSDataGridLib.DataGrid grdEntiOpeRecipro 
         Height          =   2775
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4895
         _Version        =   393216
         BackColor       =   -2147483629
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "nCod"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cDestino"
            Caption         =   "Destino"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5999.812
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   7575
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtCodEntidad 
      Height          =   375
      Left            =   840
      MaxLength       =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   8415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rutas :"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   510
   End
End
Attribute VB_Name = "frmRutasViaticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lConsulta As Boolean
Dim sSql As String
Dim rsConst As ADODB.Recordset
Dim rsEnti As ADODB.Recordset
Dim I As Integer
Dim lNuevo As Boolean
Dim oConst As NConstSistemas

Private Sub cmdAceptar_Click()
Dim sqlCod As String
Dim Cod As Integer
On Error GoTo AceptarErr

If Not ValidaDatos Then
   Exit Sub
End If

If MsgBox(" ¿ Seguro de grabar datos ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
'   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Select Case lNuevo
      Case True
            oConst.InsertaRutasViaticos Me.txtCodEntidad
      Case False
            oConst.ActualizaRutasViaticos rsEnti!nCod, Me.txtCodEntidad
   End Select
    CargaRutas
End If
txtCodEntidad.Text = ""

ActivaBotones True
grdEntiOpeRecipro.SetFocus
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelar_Click()

txtCodEntidad.Text = ""

ActivaBotones True
Me.grdEntiOpeRecipro.SetFocus
End Sub
 '*** PASI TI-ERS050-2014
Private Sub cmdEliminar_Click()
Dim sELIM As String
Dim lrs   As ADODB.Recordset
If rsEnti.EOF Then
   MsgBox "No existen datos para Eliminar", vbInformation, "¡Aviso!"
   Exit Sub
End If
If MsgBox(" ¿ Está seguro de Eliminar este dato ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
   'oConst.EliminaEntiOpeRecipro rsEnti!cpersidnro
   oConst.EliminaRutasViaticos rsEnti!nCod 'PASI 20140401 TI-ERS050-2014
   CargaRutas
   Set oDoc = Nothing
   RSClose lrs
End If
grdEntiOpeRecipro.SetFocus
End Sub
'End PASI

Private Sub cmdModificar_Click()
If rsEnti.EOF Then
   MsgBox "No existen datos para Modificar", vbInformation, "Aviso"
   Exit Sub
End If

txtCodEntidad.Text = rsEnti!cDestino

ActivaBotones False

lNuevo = False
End Sub

Private Sub cmdNuevo_Click()
    txtCodEntidad.Text = ""
    ActivaBotones False
    lNuevo = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set oConst = New NConstSistemas
       
    CargaRutas
    
    If lConsulta Then
        cmdNuevo.Visible = False
        cmdModificar.Visible = False
        cmdEliminar.Visible = False '*** PASI TI-ERS050-2014
    End If
    
    ActivaBotones True
    
    CentraForm Me
End Sub
Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub
Sub ActivaBotones(lActiva As Boolean)

cmdNuevo.Visible = lActiva
cmdModificar.Visible = lActiva
cmdEliminar.Visible = lActiva '*** PASI TI-ERS050-2014
cmdSalir.Visible = lActiva
cmdAceptar.Visible = Not lActiva
cmdCancelar.Visible = Not lActiva

txtCodEntidad.Enabled = Not lActiva

End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsEnti
Set rsEnti = Nothing
End Sub

Private Sub grdEntiOpeRecipro_GotFocus()
grdEntiOpeRecipro.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdEntiOpeRecipro_LostFocus()
grdEntiOpeRecipro.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub grdEntiOpeRecipro_HeadClick(ByVal ColIndex As Integer)
If Not rsEnti Is Nothing Then
   If Not rsEnti.EOF Then
      rsEnti.Sort = grdEntiOpeRecipro.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub CargaRutas()
    Set rsEnti = oConst.LeeRutasViaticos()
    Set Me.grdEntiOpeRecipro.DataSource = rsEnti
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False

If Len(Trim(txtCodEntidad.Text)) = 0 Then
   MsgBox "Ingrese una Ruta.", vbCritical, "Atención"
   Exit Function
End If

ValidaDatos = True
End Function

