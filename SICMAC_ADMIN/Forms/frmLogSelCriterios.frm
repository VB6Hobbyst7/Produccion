VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogSelCriterios 
   Caption         =   "Criterios de Evaluacion"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   Icon            =   "frmLogSelCriterios.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   10065
   Begin VB.TextBox txtpuntaje 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   5115
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtUltMov 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6450
      TabIndex        =   8
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5595
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   4500
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1215
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   900
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   780
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   5054
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   2310
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3405
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogSelCriterios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCodigo As String
Dim lbEditado  As Boolean
Dim lnTipo As TipoOpe
Dim clsDGAdqui As DLogAdquisi
Dim oCri As NActualizaProcesoSelecLog

Private Sub CmdCancelar_Click()
    lbEditado = True

    Limpia
    cmdEditar.Enabled = True
    cmdNuevo.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdEliminar.Enabled = True
    cmdImprimir.Enabled = True
    CargaDatos
    txtdescripcion.Enabled = False
    txtpuntaje.Enabled = False
End Sub

Private Sub cmdEditar_Click()
    
    
    If Me.Flex.TextMatrix(Me.Flex.Row, 1) = "" Then Exit Sub
    
    If txtCodigo.Text = "" Then
        MsgBox "Seleccione un Criterio de  Evaluacion Tecnico para su Modificacion", vbInformation, "Seleccione un Criterio de Evaluacion"
        Exit Sub
    End If
    
    'Activa True
    lbEditado = True
    txtdescripcion.Enabled = True
    txtpuntaje.Enabled = True
    Me.txtdescripcion.SetFocus
    
    cmdEditar.Enabled = False
    cmdNuevo.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = False
    cmdImprimir.Enabled = False
    

End Sub

Private Sub cmdEliminar_Click()
    If Me.Flex.TextMatrix(Flex.Row, 0) = "" Then Exit Sub
    
    If MsgBox("Desea Eliminar los el registro.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If Not clsDGAdqui.TipoUsado(Me.Flex.TextMatrix(Flex.Row, 0)) Then
        oCri.EliminaSeleccionMantCriterios Me.Flex.TextMatrix(Flex.Row, 0)
    Else
        MsgBox "No puede Eliminar porque el Criterio esta Uso.", vbInformation, "No se Puede Eliminar"
    End If
    CargaDatos

End Sub

Private Sub cmdGrabar_Click()
    
    Dim sactualiza As String
    sactualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
    If Me.txtdescripcion.Text = "" Then
        MsgBox "Debe ingresar una descripcion.", vbInformation, "Aviso"
        Me.txtdescripcion.SetFocus
        Exit Sub
    End If
    If lbEditado Then
        'oMer.ModificaMerDemTabla Me.txtCodigo.Text, Me.txtDescripcion.Text, Me.TxtBuscar.Text, GetMovNro(gsCodUser, gsCodAge)
        oCri.ActualizaSeleccionMantCriterios Me.txtCodigo.Text, Me.txtdescripcion.Text, txtpuntaje.Text, sactualiza
    Else
        'oMer.AgregaMerDemTabla Me.txtCodigo.Text, Me.txtDescripcion.Text, Me.TxtBuscar.Text, GetMovNro(gsCodUser, gsCodAge)
        oCri.AgregaSeleccionMantCriterios Me.txtCodigo.Text, UCase(Me.txtdescripcion.Text), txtpuntaje.Text, sactualiza
    End If
    lbEditado = True
    'Activa False
    Limpia
    CargaDatos
    
    cmdEditar.Enabled = True
    cmdNuevo.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdEliminar.Enabled = True
    cmdImprimir.Enabled = True
    txtdescripcion.Enabled = False
    txtpuntaje.Enabled = False

End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim oMer As NMeritosDemeritos
    Set oMer = New NMeritosDemeritos
    Dim lsCadena As String
    lsCadena = oMer.GetReporteMerDemTabla(gsNomAge, gsEmpresa, gdFecSis)
    lsCadena = oCri.GetReporteLogSelCriteriosTecnicos(gsNomAge, gsEmpresa, gdFecSis)
    oPrevio.Show lsCadena, Me.Caption, True, 66
    Set oPrevio = Nothing
End Sub

Private Sub cmdNuevo_Click()
    Limpia
    'Me.txtCodigo.Text = "01"
    'genrar
    Me.txtCodigo.Text = clsDGAdqui.CargaSelNumCriteriosTecnicos
    lbEditado = False
    'Activa True
        cmdEditar.Enabled = False
        cmdNuevo.Enabled = False
        cmdGrabar.Enabled = True
        cmdCancelar.Enabled = True
        cmdEliminar.Enabled = False
        cmdImprimir.Enabled = False
        txtdescripcion.Enabled = True
        txtpuntaje.Enabled = True
    'Me.txtDescripcion.SetFocus
End Sub

Private Sub Limpia()
    Me.txtCodigo.Text = ""
    Me.txtdescripcion.Text = ""
    Me.txtpuntaje.Text = ""
    Me.txtUltMov.Text = ""
End Sub
Private Sub CargaDatos()
    Dim rsCri As ADODB.Recordset
    Set rsCri = New ADODB.Recordset
    Set rsCri = clsDGAdqui.CargaSelCriteriosTecnicos(1)
    If rsCri.EOF And rsCri.BOF Then
        Flex.Clear
        Flex.Rows = 2
        Flex.Cols = 4
        Flex.TextMatrix(0, 0) = "Codigo"
        Flex.TextMatrix(0, 1) = "Descripcion"
        Flex.TextMatrix(0, 2) = "Descripcion"
        Flex.TextMatrix(0, 3) = "UltActualizacion"
        txtCodigo.Text = ""
        txtdescripcion.Text = ""
        txtUltMov.Text = ""
    Else
        Set Flex.DataSource = rsCri
    End If
    Me.Flex.ColWidth(0) = 1000
    Me.Flex.ColWidth(1) = 4000
    Me.Flex.ColWidth(2) = 1000
    Me.Flex.ColWidth(3) = 3000
    
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Flex_Click()
If Flex.Row <= 0 Then Exit Sub
txtCodigo.Text = Flex.TextMatrix(Flex.Row, 0)
txtdescripcion.Text = Flex.TextMatrix(Flex.Row, 1)
txtpuntaje.Text = Flex.TextMatrix(Flex.Row, 2)
txtUltMov.Text = Flex.TextMatrix(Flex.Row, 3)

End Sub

Private Sub Form_Load()
    Me.Width = 8235
    Me.Height = 4305
    Set clsDGAdqui = New DLogAdquisi
    Set oCri = New NActualizaProcesoSelecLog
    CargaDatos
    cmdEditar.Enabled = True
    cmdNuevo.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdEliminar.Enabled = True
    cmdImprimir.Enabled = True
End Sub
