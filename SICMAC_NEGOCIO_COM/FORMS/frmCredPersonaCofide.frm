VERSION 5.00
Begin VB.Form frmCredPersonaCofide 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Personas COFIDE"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   Icon            =   "frmCredPersonaCofide.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   9135
      Begin SICMACT.FlexEdit grdPersonas 
         Height          =   2805
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   4948
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Cod.Cofide"
         EncabezadosAnchos=   "250-1700-3600-1500-1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   255
         RowHeight0      =   300
      End
      Begin VB.Label Label2 
         Caption         =   "Si desea modificar, seleccionar elemento de la lista"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   9135
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7680
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraInfoCliente 
      Caption         =   "Información del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtCodCofide 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cboRelacion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Cod.COFIDE:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Persona:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCredPersonaCOFIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bNuevo As Boolean

Private Sub cmdCancelar_Click()
    LimpiaControles
End Sub
Private Sub LimpiaControles()
    TxtBCodPers.Text = ""
    lblNombre.Caption = ""
    txtCodCofide.Text = ""
    cboRelacion.ListIndex = 0
    fraInfoCliente.Enabled = False
    cmdNuevo.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdEditar.Enabled = False
    cmdNuevo.Default = True
    bNuevo = False
End Sub
Public Sub Inicio()

    Dim oCons As COMDConstSistema.DCOMGeneral
    Dim rsRel As ADODB.Recordset
    Set oCons = New COMDConstSistema.DCOMGeneral
    
    Set rsRel = oCons.GetConstante(3036, , "'1[18]'")
    While Not rsRel.EOF And Not rsRel.BOF
        cboRelacion.AddItem rsRel!cdescripcion & Space(100) & rsRel!nConsValor
        rsRel.MoveNext
    Wend
    Call LimpiaControles
    Call ListarPersonasCofide
    CentraForm Me
    Me.Show 1
End Sub

Private Sub cmdEditar_Click()
    cmdNuevo.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdEditar.Enabled = False
    cmdGrabar.Default = True
    Me.fraInfoCliente.Enabled = True
    TxtBCodPers.SetFocus
    bNuevo = False
End Sub

Private Sub cmdGrabar_Click()
    Dim oPers As COMDCredito.DCOMCredActBD
    Set oPers = New COMDCredito.DCOMCredActBD

    If ValidaDatos Then
        If MsgBox("¿Está seguro de haber ingresado correctamente los datos?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            If bNuevo = True Then
                Call oPers.AgregaNuevoPersCofide(TxtBCodPers.Text, CInt(Trim(Right(cboRelacion.Text, 2))), Me.txtCodCofide.Text)
            Else
                Call oPers.ModificaPersCofide(TxtBCodPers.Text, CInt(Trim(Right(cboRelacion.Text, 2))), Me.txtCodCofide.Text)
            End If
            MsgBox "Se registraron correctamente los datos", vbInformation, "Aviso"
            Call LimpiaControles
            Call ListarPersonasCofide
        End If
    End If
End Sub
Private Function ValidaDatos() As Boolean
    Dim oPers As COMDCredito.DCOMCredActBD
    Dim Validado As Boolean
    
    Set oPers = New COMDCredito.DCOMCredActBD
    
    Validado = True
    If lblNombre.Caption = "" Or Me.TxtBCodPers.Text = "" Then
        MsgBox "No ha ingresado la persona o empresa", vbInformation, "Aviso"
        Validado = False
    ElseIf Me.txtCodCofide = "" Then
        MsgBox "No ha ingresado el código Cofide", vbInformation, "Aviso"
        Validado = False
    ElseIf Me.cboRelacion.Text = "" Then
        MsgBox "Debe seleccionar el tipo de relación", vbInformation, "Aviso"
        Validado = False
    End If
    If oPers.ValidaExistePersCofide(TxtBCodPers.Text, CInt(Trim(Right(cboRelacion.Text, 2)))) And bNuevo = True Then
        MsgBox "La relación ya fue ingresada", vbInformation, "Aviso"
        Validado = False
    End If
    If oPers.ValidaExistePersCofide(txtCodCofide.Text, CInt(Trim(Right(cboRelacion.Text, 2)))) Then
        MsgBox "La código Cofide ya existe", vbInformation, "Aviso"
        Validado = False
    End If
    ValidaDatos = Validado
End Function
Private Sub ListarPersonasCofide()
    Dim oPers As COMDCredito.DCOMCredActBD
    Set oPers = New COMDCredito.DCOMCredActBD
    Set grdPersonas.Recordset = oPers.ListaPersonasCofide
End Sub
Private Sub cmdNuevo_Click()
    cmdNuevo.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdEditar.Enabled = False
    cmdGrabar.Default = True
    Me.fraInfoCliente.Enabled = True
    TxtBCodPers.SetFocus
    bNuevo = True
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub grdPersonas_Click()
    Dim nFila As Integer
    cmdEditar.Enabled = True
    
    cmdNuevo.Enabled = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = True
    cmdEditar.Enabled = True
    Me.fraInfoCliente.Enabled = False
    cmdEditar.Default = True
    
    nFila = grdPersonas.Row
    TxtBCodPers.Text = grdPersonas.TextMatrix(nFila, 1)
    lblNombre.Caption = grdPersonas.TextMatrix(nFila, 2)
    txtCodCofide.Text = grdPersonas.TextMatrix(nFila, 4)
    cboRelacion.ListIndex = IndiceListaCombo(cboRelacion, IIf(grdPersonas.TextMatrix(nFila, 3) = "OPERADOR", 18, 11))
End Sub
Private Sub TxtBCodPers_EmiteDatos()
    lblNombre.Caption = TxtBCodPers.psDescripcion
    txtCodCofide.SetFocus
End Sub

Private Sub txtCodCofide_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboRelacion.SetFocus
    End If
End Sub
