VERSION 5.00
Begin VB.Form frmCredConsultarFichaSobreEnd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar : Ficha de Sobreendeudado"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   Icon            =   "frmCredConsultarFichaSobreEnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.TextBox txtAno 
         Height          =   285
         Left            =   5040
         TabIndex        =   3
         Top             =   380
         Width           =   1095
      End
      Begin VB.ComboBox cmbMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   380
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   7695
      End
      Begin SICMACT.TxtBuscar txtCodPers 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
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
         EnabledText     =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Año :"
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Mes :"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre y Apellido"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
   End
   Begin SICMACT.FlexEdit fsEvaluacion 
      Height          =   1815
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   3201
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "N°-Código-Evaluación-Plan de Mitigación"
      EncabezadosAnchos=   "400-1300-2500-5000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "R-L-L-C"
      FormatosEdit    =   "3-1-0-0"
      TextArray0      =   "N°"
      lbEditarFlex    =   -1  'True
      Enabled         =   0   'False
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      TipoBusPersona  =   2
   End
End
Attribute VB_Name = "frmCredConsultarFichaSobreEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Inicio(ByVal psTipoRegMant As Integer) As Boolean
Dim oComboFecha As COMDCredito.DCOMCredito
Dim rsComboFecha As ADODB.Recordset

Dim PermisoFichaSobrEnd As COMNCredito.NCOMCredito
Set PermisoFichaSobrEnd = New COMNCredito.NCOMCredito
    
'(2: Gerencia de Riesgo)
fnTipoPermiso = PermisoFichaSobrEnd.ObtieneTipoPermisoFichaSobreEnd(psTipoRegMant, gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo
    
    If CargaControlesTipoPermiso(fnTipoPermiso) Then
    
        Set oComboFecha = New COMDCredito.DCOMCredito
        Set rsComboFecha = oComboFecha.MostrarComboFecha()
        
        CargarComboBox rsComboFecha, cmbMes
        
        txtAno.Text = Format(gdFecSis, "yyyy")
    Else
        Exit Function
    End If
     Me.Show 1
End Function

Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer) As Boolean
    '1: JefeAgencia o JefeNegocio->
    If TipoPermiso = 2 Then
        'Call HabilitaControles(False)
        CargaControlesTipoPermiso = True
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        'Call HabilitaControles(False)
        CargaControlesTipoPermiso = False
    End If
End Function

Private Sub cmdBuscar_Click()

Dim lnFila As Integer

Dim oBuscar As COMDCredito.DCOMCredito
Dim rsObtBus As ADODB.Recordset
Set oBuscar = New COMDCredito.DCOMCredito

If Validar Then
    Set rsObtBus = oBuscar.ObtenerBusquedaFichaSobEnd(txtCodPers.Text, (cmbMes.ItemData(cmbMes.ListIndex)), txtAno.Text)
        
        If Not (rsObtBus.EOF And rsObtBus.BOF) Then
            fsEvaluacion.Clear
            fsEvaluacion.FormaCabecera
            
            Call LimpiaFlex(fsEvaluacion)
                Do While Not rsObtBus.EOF
                    fsEvaluacion.AdicionaFila
                    lnFila = fsEvaluacion.row
                    fsEvaluacion.TextMatrix(lnFila, 1) = rsObtBus!sCodigo
                    fsEvaluacion.TextMatrix(lnFila, 2) = rsObtBus!sEval
                    fsEvaluacion.TextMatrix(lnFila, 3) = rsObtBus!cPlanmitigacion
                    rsObtBus.MoveNext
                Loop
            rsObtBus.Close
            Set rsObtBus = Nothing
        Else
            MsgBox "No Existe Cliente", vbInformation, "Aviso"
            Call Limpiar
        End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub

Private Sub txtCodPers_EmiteDatos()
Dim oDPersonaS As COMDPersona.DCOMPersonas
Dim oRs As ADODB.Recordset

Call Limpiar

If Trim(txtCodPers.Text) = "" Then Exit Sub
       sPersCod = Trim(txtCodPers.Text)
       Set oDPersonaS = New COMDPersona.DCOMPersonas
       Set oRs = oDPersonaS.BuscaCliente(sPersCod, BusquedaCodigo)
       Set oDPersonaS = Nothing
                                   
       If Not oRs.EOF And Not oRs.BOF Then
        txtNombre.Text = oRs!cPersNombre
       End If
       
End Sub

Private Sub Limpiar()
    LimpiaFlex fsEvaluacion
End Sub

Private Function Validar() As Boolean

Validar = True

'Codigo Persona
    If txtCodPers.Text = "" Then
        MsgBox "Busque al Cliente", vbInformation, "Aviso"
        txtCodPers.SetFocus
        Validar = False
        Exit Function
    End If
'Mes
    If (cmbMes.ListIndex) = -1 Then
        MsgBox "Seleccione el Mes", vbInformation, "Aviso"
        cmbMes.SetFocus
        Validar = False
        Exit Function
    End If
'Ano
    If txtAno.Text = "" Then
        MsgBox "Ingrese el Año", vbInformation, "Aviso"
        txtAno.SetFocus
        Validar = False
        Exit Function
    End If
'Nombre Persona
    If txtNombre.Text = "" Then
        MsgBox "Seleccion al Cliente", vbInformation, "Aviso"
        Validar = False
        Exit Function
    End If
End Function

Private Sub txtCodPers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmbMes
    End If
End Sub

Private Sub cmbMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtAno
    End If
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmdBuscar
    End If
End Sub
