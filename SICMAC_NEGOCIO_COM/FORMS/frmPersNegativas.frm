VERSION 5.00
Begin VB.Form frmPersNegativas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Preventivo del LAFT"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmPersNegativas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtfuentes 
      Height          =   570
      Left            =   12720
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   26
      Tag             =   "txtPrincipal"
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame fracontrol 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4905
      Left            =   8760
      TabIndex        =   30
      Top             =   0
      Width           =   2700
      Begin VB.CommandButton cmdCargaLote 
         Caption         =   "Carga Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2295
         Left            =   120
         TabIndex        =   43
         Top             =   2520
         Width           =   2415
         Begin VB.OptionButton optCondicion 
            Caption         =   "Esta Persona"
            Enabled         =   0   'False
            Height          =   255
            Index           =   7
            Left            =   720
            TabIndex        =   24
            Top             =   1920
            Width           =   1575
         End
         Begin VB.OptionButton optCondicion 
            Caption         =   "Lista ONU"
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   23
            Top             =   1680
            Width           =   1575
         End
         Begin VB.OptionButton optCondicion 
            Caption         =   "Lista OFAC"
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   22
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton optCondicion 
            Caption         =   "Vinculados-NT"
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   21
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton optCondicion 
            Caption         =   "PEPS"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   20
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton optCondicion 
            Caption         =   "Fraudulentos"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   19
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optCondicion 
            Caption         =   "Negativo"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   18
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optCondicion 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.CommandButton cmdImprimir 
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
            Left            =   120
            Picture         =   "frmPersNegativas.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Imprimir Solicitud"
            Top             =   240
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   14
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eli&minar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   16
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.Frame FraTipoRea 
      Caption         =   "Tipo de Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   8400
      Begin VB.OptionButton OptTR 
         Caption         =   "Persona Jurídica"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   4320
         TabIndex        =   29
         Top             =   240
         Width           =   1980
      End
      Begin VB.OptionButton OptTR 
         Caption         =   "Persona Natural"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   1950
      End
   End
   Begin VB.TextBox txtJustificacion 
      Height          =   570
      Left            =   12360
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   25
      Tag             =   "txtPrincipal"
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Justificación y Fuente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   120
      TabIndex        =   32
      Top             =   5040
      Width           =   11295
      Begin VB.CommandButton cmdVisitasEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   33
         Top             =   2040
         Width           =   1140
      End
      Begin VB.CommandButton cmdVisitasNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   1140
      End
      Begin SICMACT.FlexEdit feJustificanegativa 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2937
         Cols0           =   10
         HighLight       =   1
         EncabezadosNombres=   "#-Delito-Juzgado-Oficio Multiple-Oficio Juzgado-Expediente-Departamento-Tipo-Comentario-Aux"
         EncabezadosAnchos=   "400-2500-2500-2500-2500-2300-2000-2000-5000-0"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-6-7-8-X"
         ListaControles  =   "0-0-0-0-0-0-3-3-0-0"
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4335
      Left            =   120
      TabIndex        =   34
      Top             =   600
      Width           =   8415
      Begin VB.TextBox txtApeCasada 
         Height          =   405
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1680
         Width           =   2970
      End
      Begin VB.CommandButton cmdBuscarPersona 
         Caption         =   "..."
         Height          =   315
         Left            =   6480
         TabIndex        =   45
         Top             =   285
         Width           =   315
      End
      Begin VB.TextBox txtNomRazSocial 
         Height          =   405
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   1
         Tag             =   "txtPrincipal"
         Top             =   720
         Width           =   2970
      End
      Begin VB.TextBox txtApePat 
         Height          =   405
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "txtPrincipal"
         Top             =   1200
         Width           =   2970
      End
      Begin VB.TextBox txtApeMat 
         Height          =   405
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "txtPrincipal"
         Top             =   1680
         Width           =   2970
      End
      Begin VB.ComboBox cboCondicion 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2520
         Width           =   3090
      End
      Begin VB.TextBox txtcargo 
         Height          =   405
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "txtPrincipal"
         Top             =   3480
         Width           =   4995
      End
      Begin VB.TextBox txtinstitucion 
         Height          =   405
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   7
         Tag             =   "txtPrincipal"
         Top             =   3000
         Width           =   4995
      End
      Begin VB.CheckBox Chkestado 
         Caption         =   "Descativar"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtNumDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   18
         TabIndex        =   0
         Tag             =   "txtPrincipal"
         Top             =   240
         Width           =   2970
      End
      Begin VB.CommandButton cmdexaminar 
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
         Height          =   315
         Left            =   5280
         TabIndex        =   10
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Casada"
         Height          =   195
         Left            =   5280
         TabIndex        =   46
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
         Height          =   195
         Left            =   720
         TabIndex        =   42
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Num.Doc.ID.:"
         Height          =   195
         Left            =   720
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   195
         Left            =   720
         TabIndex        =   40
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
         Height          =   195
         Left            =   720
         TabIndex        =   39
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Condición"
         Height          =   195
         Left            =   720
         TabIndex        =   38
         Top             =   2520
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Institucion"
         Height          =   195
         Left            =   720
         TabIndex        =   37
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   195
         Left            =   720
         TabIndex        =   36
         Top             =   3600
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   720
         TabIndex        =   35
         Top             =   2160
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmPersNegativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim nTipoOperacion As Integer '0 Nuevo...1 Modificar
Dim i As Integer
Dim FEPatVehPersNoMoverdeFila As Integer
Dim lcMovNro As String
Dim lnCondicion As Integer
Dim indice As Integer
Dim feJustificanegativaNoMoverdeFila As Integer 'WIOR 20120329
Dim oPersona As UPersona_Cli 'WIOR 20120710
Dim fnTpoDoc As Integer 'EJVG20121002
Dim fnFrmPersona As Integer 'JGPA20191217 - Obs. [NAGL]
Dim fbSalir As Boolean 'JGPA20191217 - Obs. [NAGL]

Sub Limpiar_Controles()

    Me.txtNumDoc.Text = ""
    Me.txtNomRazSocial.Text = ""
    Me.txtApePat.Text = ""
    Me.txtApeMat.Text = ""
    Me.txtJustificacion.Text = ""
    Me.txtfuentes.Text = ""
    Me.txtcargo.Text = ""
    Me.txtinstitucion.Text = ""
    Me.cboCondicion.ListIndex = -1
    Me.Chkestado.value = False
    Call LimpiaFlex(Me.feJustificanegativa)
    lcMovNro = ""
    Me.optCondicion(7).Enabled = False  'JACA 20110217
    Me.optCondicion(0).value = True 'JACA 20110217
    Me.txtApeCasada.Text = "" 'WIOR 20120710
    Call OptTR_Click(0) 'WIOR 20120710
    OptTR.Item(0).value = True 'WIOR 20120710

End Sub

Private Sub cboCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub Chkestado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub
'WIOR 20120710*******************************************************
Private Sub cmdBuscarPersona_Click()
Dim oPersBuscada As COMDPersona.UCOMPersona
Set oPersBuscada = New COMDPersona.UCOMPersona
Dim Persona As COMDPersona.UCOMPersona
Set Persona = New COMDPersona.UCOMPersona
Set oPersBuscada = frmBuscaPersona.Inicio
Dim lnCondicion As Integer
If oPersBuscada Is Nothing Then Exit Sub

If Persona.ValidaEnListaNegativaCondicion(IIf(IsNull(oPersBuscada.sPersIdnroDNI), "", oPersBuscada.sPersIdnroDNI), IIf(IsNull(oPersBuscada.sPersIdnroRUC), "", oPersBuscada.sPersIdnroRUC), lnCondicion, oPersBuscada.sPersNombre) Then
    MsgBox "Persona ya se encuentra en la base negativa.", vbInformation, "Aviso"
    Exit Sub
End If

If Cargar_Datos_Persona(Trim(oPersBuscada.sPersCod)) = False Then
    MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
    Exit Sub
End If
End Sub

Function Cargar_Datos_Persona(ByVal pcPersCod As String) As Boolean
    
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli
    oPersona.sCodAge = gsCodAge
    
    Cargar_Datos_Persona = True
    
    Call oPersona.RecuperaPersona(pcPersCod, , gsCodUser)
    
    If oPersona.PersCodigo = "" Then
        Cargar_Datos_Persona = False
        Exit Function
    End If
    Call CargaDatos
End Function
Private Sub CargaDatos()
  If oPersona.Personeria = gPersonaNat Then
        Me.txtApePat.Text = oPersona.ApellidoPaterno
        Me.txtApeMat.Text = oPersona.ApellidoMaterno
        Me.txtNomRazSocial.Text = oPersona.Nombres
        Me.txtApeCasada.Text = oPersona.ApellidoCasada
        Call OptTR_Click(0)
        OptTR.Item(0).value = True
    Else
        txtNomRazSocial.Text = oPersona.NombreCompleto
        Call OptTR_Click(1)
        OptTR.Item(1).value = True
    End If
End Sub
'WIOR - FIN***************************************************
Private Sub cmdCancelar_Click()

    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    'cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    cmdexaminar.Enabled = True
    
    Call Limpiar_Controles
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
     Call borra_variables
     feJustificanegativaNoMoverdeFila = -1 'WIOR 20120329
     Me.cmdBuscarPersona.Visible = False 'WIOR 20120710
     cmdCargaLote.Enabled = True 'WIOR 20130810
End Sub
'WIOR 20130808 ************************
Private Sub cmdCargaLote_Click()
frmPersNegativasLote.Show 1
End Sub
'WIOR FIN *****************************

Private Sub CmdEditar_Click()
nTipoOperacion = 1
    If Not Me.txtNumDoc.Text = "" Then
        Call Habilita_Grabar(True)
        Call Habilita_Modifica_Datos(True)
        Me.txtNomRazSocial.SetFocus
    Else
        Call Habilita_Grabar(True)
        Me.txtNumDoc.Enabled = True
        Call Habilita_Modifica_Datos(True)
        Me.txtNumDoc.SetFocus
    End If
    cmdGrabar.Enabled = True
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    'cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    cmdexaminar.Enabled = False
    feJustificanegativa.lbEditarFlex = True 'JACA 20110217
    feJustificanegativaNoMoverdeFila = feJustificanegativa.row 'WIOR 20120329
    cmdCargaLote.Enabled = False 'WIOR 20130810
End Sub

Private Sub CmdEliminar_Click()
Dim oPers As COMDPersona.DCOMPersonas
If MsgBox("¿Está seguro que desea eliminar este Registro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    Set oPers = New COMDPersona.DCOMPersonas
    '**********************MADM 20090928
    If Me.txtNumDoc.Text = "" Then
        Call oPers.eliminaperslistanegativa_Gral(IIf(Me.OptTR(0) = True, 1, 2), Trim(Me.txtApePat.Text), Trim(Me.txtApeMat.Text), Trim(Me.txtNomRazSocial.Text), lcMovNro) 'JACA 20110530 SE AGREGO lcMovNro
    Else
        'Call oPers.eliminaperslistanegativa(IIf(Me.OptTR(0) = True, 1, 2), Trim(Me.txtNumDoc.Text), lcMovNro) 'JACA 20110530 SE AGREGO lcMovNro
        Call oPers.eliminaperslistanegativa(IIf(fnTpoDoc > 0, fnTpoDoc, IIf(Me.OptTR(0) = True, 1, 2)), Trim(Me.txtNumDoc.Text), lcMovNro) 'JACA 20110530 SE AGREGO lcMovNro
    End If
    '**********************MADM 20090928
   
    Set oPers = Nothing
    
'    cmdEliminar.Enabled = False
'    cmdEditar.Enabled = False

    Call Me.Limpiar_Controles
End If
End Sub

Private Sub borra_variables()
frmBuscaPersonaNegativa.lnTipoDocId = 0
frmBuscaPersonaNegativa.cMovNro = ""
frmBuscaPersonaNegativa.snumdoc = ""
End Sub

Private Sub cmdexaminar_Click()
Dim oPers As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim i As Integer
    
Set oPers = New COMDPersona.DCOMPersonas
Call frmBuscaPersonaNegativa.Inicio
'**********MADM 20090928
If Not frmBuscaPersonaNegativa.snumdoc = "" Then
    Set rs = oPers.CargaDatosPersNegativas(frmBuscaPersonaNegativa.lnTipoDocId, frmBuscaPersonaNegativa.snumdoc, frmBuscaPersonaNegativa.cMovNro) 'JACA 20110530 SE AGREGO frmBuscaPersonaNegativa.cMovNro
ElseIf Not frmBuscaPersonaNegativa.cMovNro = "" Then
    Set rs = oPers.CargaDatosPersNegativas_MovNro(frmBuscaPersonaNegativa.cMovNro)
Else
Exit Sub
End If
'**********MADM 20090928
If Not (rs.BOF And rs.EOF) Then
    
    Me.txtNumDoc.Text = rs!cNumId
    Me.txtNomRazSocial.Text = Trim(rs!cNombre)
    Me.txtApePat.Text = Trim(rs!cApePat)
    Me.txtApeMat.Text = Trim(rs!capemat)
    Me.txtfuentes.Text = rs!cFuente
    ''cboCondicion.ListIndex = IndiceListaCombo(cboCondicion, Trim(str(rs!nCondicion)), 2)
    cboCondicion.ListIndex = IndiceListaCombo(cboCondicion, Trim(str(rs!nCondicion))) 'marg
    'madm 20100510
    Me.txtinstitucion.Text = IIf(IsNull(rs!cInstitucion), "", rs!cInstitucion)
    Me.txtcargo.Text = IIf(IsNull(rs!cCargo), "", rs!cCargo)
    Me.txtApeCasada.Text = IIf(IsNull(rs!cApeCasada), "", rs!cApeCasada) 'WIOR 20120710
    'end madm
    'madm 20110115
    If rs!nEstado = 0 Then
        Me.Chkestado.value = 0
    Else
        Me.Chkestado.value = 1
    End If
    'end madm
    
    'JACA 20110217
    'Me.txtJustificacion.Text = rs!cJustificacion
    Set rs1 = oPers.CargaDatosPersNegativasJustificacion_MovNro(frmBuscaPersonaNegativa.cMovNro)
    lcMovNro = frmBuscaPersonaNegativa.cMovNro
    
     If Not (rs1.BOF And rs1.EOF) Then
        feJustificanegativa.lbEditarFlex = True
        Call LimpiaFlex(feJustificanegativa)
            For i = 0 To rs1.RecordCount - 1
                feJustificanegativa.AdicionaFila
                feJustificanegativa.TextMatrix(i + 1, 0) = i + 1
                feJustificanegativa.TextMatrix(i + 1, 1) = rs1!cDelito
                feJustificanegativa.TextMatrix(i + 1, 2) = rs1!cJuzgado
                feJustificanegativa.TextMatrix(i + 1, 3) = rs1!cOfMultiple
                feJustificanegativa.TextMatrix(i + 1, 4) = rs1!cOfJuzgado
                feJustificanegativa.TextMatrix(i + 1, 5) = rs1!cExpediente
                'WIOR 20120329-INICIO
                feJustificanegativa.TextMatrix(i + 1, 6) = rs1!Departamento & Space(75) & rs1!CodDepartamento
                feJustificanegativa.TextMatrix(i + 1, 7) = rs1!Tipo & Space(75) & rs1!CodTipo
                feJustificanegativa.TextMatrix(i + 1, 8) = rs1!cComentario
                'WIOR -FIN
                rs1.MoveNext
            Next i
        feJustificanegativa.lbEditarFlex = False
    End If
    'JACA END
    
    If rs!nTipoPers = 1 Then
        Me.OptTR(0).value = True
        Me.OptTR(1).value = False
    Else
        Me.OptTR(0).value = False
        Me.OptTR(1).value = True
    End If

    Me.cmdEditar.Enabled = True
    Me.cmdEliminar.Enabled = True
    Me.optCondicion(7).Enabled = True 'JACA 20110217
    Me.optCondicion(7).value = True 'JACA 20110217
Else
    Me.cmdEditar.Enabled = False
    Me.cmdEliminar.Enabled = False
    Me.Limpiar_Controles
End If

Set oPers = Nothing
End Sub

Private Sub cmdGrabar_Click()
Dim oPers As COMDPersona.DCOMPersonas
    Set oPers = New COMDPersona.DCOMPersonas
Dim RSVerifica As ADODB.Recordset
fbSalir = False 'JGPA20191217
If Valida_Datos = False Then Exit Sub

If MsgBox("¿Está seguro de registrar los datos ingresados?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

'Set oPol = New COMDCredito.DCOMPoliza

'*****************************MADM 20090928
 
If Me.txtNumDoc.Text = "" Then
Set RSVerifica = oPers.VerificaPersListaNegativa_Gral(IIf(Me.OptTR(0), 1, 2), Trim(Me.txtApePat), Trim(Me.txtApeMat), Trim(Me.txtNomRazSocial), lcMovNro, Trim(Me.txtApeCasada.Text)) 'JACA 20110530 SE AGREGO lcMovNro
'WIOR 20120710 AGREGO TRIM(ME.txtApeCasada.Text )
Else
'Set RSVerifica = oPers.VerificaPersListaNegativa(IIf(Me.OptTR(0), 1, 2), IIf(Me.OptTR(0), 1, 2), Trim(Me.txtNumDoc), lcMovNro) 'JACA 20110530 SE AGREGO lcMovNro
    Set RSVerifica = oPers.VerificaPersListaNegativa(IIf(Me.OptTR(0), 1, 2), IIf(fnTpoDoc > 0, fnTpoDoc, IIf(Me.OptTR(0), 1, 2)), Trim(Me.txtNumDoc), lcMovNro) 'EJVG20121002
End If
'*****************************MADM 20090928
If nTipoOperacion = 0 Then
    If Not (RSVerifica.EOF And RSVerifica.BOF) Then
        MsgBox "Esta persona ya fue registrada, verifique por favor dando clic en el boton buscar.", vbInformation, "Mensaje"
        cmdexaminar.Enabled = True
        Exit Sub
    End If
        Dim pmMovNum As String
        pmMovNum = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        'Call oPers.RegistraPersNegativo(IIf(Me.OptTR(0), 1, 2), IIf(Me.OptTR(0), 1, 2), Me.txtNumDoc, Trim(Me.txtNomRazSocial), Trim(Me.txtApePat), Trim(Me.txtApeMat), IIf(Me.txtJustificacion.Visible = False, "", Me.txtJustificacion), IIf(Me.txtfuentes.Visible = False, "", Me.txtfuentes), Trim(Me.txtinstitucion.Text), Trim(Me.txtcargo.Text), pmMovNum, CInt(Trim(Right(cboCondicion.Text, 2))), IIf(Me.Chkestado.value, 1, 0), Trim(Me.txtApeCasada.Text))
        'WIOR 20120710 SE AGREGO EL PARAMETRO Me.txtApeCasada.Text.
        Call oPers.RegistraPersNegativo(IIf(Me.OptTR(0), 1, 2), IIf(fnTpoDoc > 0, fnTpoDoc, IIf(Me.OptTR(0), 1, 2)), Me.txtNumDoc, Trim(Me.txtNomRazSocial), Trim(Me.txtApePat), Trim(Me.txtApeMat), IIf(Me.txtJustificacion.Visible = False, "", Me.txtJustificacion), IIf(Me.txtfuentes.Visible = False, "", Me.txtfuentes), Trim(Me.txtinstitucion.Text), Trim(Me.txtcargo.Text), pmMovNum, CInt(Trim(Right(cboCondicion.Text, 2))), IIf(Me.Chkestado.value, 1, 0), Trim(Me.txtApeCasada.Text)) 'EJVG20121002
        'JACA 20110217
        For i = 1 To Me.feJustificanegativa.Rows - 1
            If feJustificanegativa.TextMatrix(i, 1) <> "" Or feJustificanegativa.TextMatrix(i, 2) <> "" Then
               'Call oPers.ModificaPersNegativoJustifica(pmMovNum, feJustificanegativa.TextMatrix(i, 1), feJustificanegativa.TextMatrix(i, 2), feJustificanegativa.TextMatrix(i, 3), feJustificanegativa.TextMatrix(i, 4), feJustificanegativa.TextMatrix(i, 5))
               Call oPers.ModificaPersNegativoJustifica(pmMovNum, feJustificanegativa.TextMatrix(i, 1), feJustificanegativa.TextMatrix(i, 2), feJustificanegativa.TextMatrix(i, 3), feJustificanegativa.TextMatrix(i, 4), feJustificanegativa.TextMatrix(i, 5), IIf(feJustificanegativa.TextMatrix(i, 6) = "", "", Trim(Right(feJustificanegativa.TextMatrix(i, 6), 20))), IIf(feJustificanegativa.TextMatrix(i, 7) = "", 0, Trim(Right(feJustificanegativa.TextMatrix(i, 7), 20))), feJustificanegativa.TextMatrix(i, 8)) 'WIOR 20120329
            
            End If
        Next
        'JACA END
Else
     'EDITAR
     'JACA 20110217
     If (RSVerifica.EOF = False And RSVerifica.BOF = False) Then 'SI TIENE REGISTRO ENTRA (PERTENECE A OTRO)
         If (RSVerifica!cMovNro <> lcMovNro) Then
            MsgBox "Los datos ingresados se encuentran registrados con anterioridad, verifique!!!.", vbInformation, "Mensaje"
         Else
            Call ModificarPersNegativo
         End If
     Else
           Call ModificarPersNegativo
     End If
     'END JACA
    
End If
Set oPers = Nothing
    
    cmdGrabar.Enabled = False
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    'cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    cmdexaminar.Enabled = True
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
    Call Me.Limpiar_Controles
    
'    If (Me.txtfuentes.Visible = False) Then
'        Unload Me
'    End If
Me.cmdBuscarPersona.Visible = False 'WIOR 20120710
cmdCargaLote.Enabled = True 'WIOR 20130810
'JGPA20191217 Obs. [NAGL]
If fnFrmPersona = 1 Then
    fbSalir = True
End If
If fnFrmPersona = 1 And fbSalir = True Then
    Unload Me
    Exit Sub
End If
'End JGPA20191217
End Sub

'JACA 20110217
Private Sub ModificarPersNegativo()
    Dim oPers As COMDPersona.DCOMPersonas
    Set oPers = New COMDPersona.DCOMPersonas
    Call oPers.EliminaPersNegativoJustifica(lcMovNro)
           
     For i = 1 To Me.feJustificanegativa.Rows - 1
         If feJustificanegativa.TextMatrix(i, 1) <> "" Or feJustificanegativa.TextMatrix(i, 2) <> "" Then
         ' Call oPers.ModificaPersNegativoJustifica(lcMovNro, feJustificanegativa.TextMatrix(i, 1), feJustificanegativa.TextMatrix(i, 2), feJustificanegativa.TextMatrix(i, 3), feJustificanegativa.TextMatrix(i, 4), feJustificanegativa.TextMatrix(i, 5))
            Call oPers.ModificaPersNegativoJustifica(lcMovNro, feJustificanegativa.TextMatrix(i, 1), feJustificanegativa.TextMatrix(i, 2), feJustificanegativa.TextMatrix(i, 3), feJustificanegativa.TextMatrix(i, 4), feJustificanegativa.TextMatrix(i, 5), IIf(feJustificanegativa.TextMatrix(i, 6) = "", "", Trim(Right(feJustificanegativa.TextMatrix(i, 6), 20))), IIf(feJustificanegativa.TextMatrix(i, 7) = "", 0, Trim(Right(feJustificanegativa.TextMatrix(i, 7), 20))), feJustificanegativa.TextMatrix(i, 8)) 'WIOR 20120329
         End If
     Next
     'If Me.txtNumDoc.Text = "" Then
      
          '*******************MADM 20090928
          'MODIFICA X NOMBRE
           'Call oPers.ModificaPersNegativo_Gral(IIf(Me.OptTR(0), 1, 2), IIf(Me.OptTR(0), 1, 2), Me.txtNumDoc, Trim(Me.txtNomRazSocial), Trim(Me.txtApePat), Trim(Me.txtApeMat), IIf(Me.txtJustificacion.Visible = False, "", Me.txtJustificacion), IIf(Me.txtfuentes.Visible = False, "", Me.txtfuentes), Trim(Me.txtinstitucion.Text), Trim(Me.txtcargo.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), CInt(Trim(Right(cboCondicion.Text, 2))), IIf(Me.Chkestado.value, 1, 0), lcMovNro)
          '*******************MADM 20090928
     'Else
        'MODIFICA X NUMERO DE DOCUMENTO
        'Call oPers.ModificaPersNegativo(IIf(Me.OptTR(0), 1, 2), IIf(Me.OptTR(0), 1, 2), Me.txtNumDoc, Trim(Me.txtNomRazSocial), Trim(Me.txtApePat), Trim(Me.txtApeMat), IIf(Me.txtJustificacion.Visible = False, "", Me.txtJustificacion), IIf(Me.txtfuentes.Visible = False, "", Me.txtfuentes), Trim(Me.txtinstitucion.Text), Trim(Me.txtcargo.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), CInt(Trim(Right(cboCondicion.Text, 2))), IIf(Me.Chkestado.value, 1, 0), lcMovNro, Trim(Me.txtApeCasada.Text))
        Call oPers.ModificaPersNegativo(IIf(Me.OptTR(0), 1, 2), IIf(fnTpoDoc > 0, fnTpoDoc, IIf(Me.OptTR(0), 1, 2)), Me.txtNumDoc, Trim(Me.txtNomRazSocial), Trim(Me.txtApePat), Trim(Me.txtApeMat), IIf(Me.txtJustificacion.Visible = False, "", Me.txtJustificacion), IIf(Me.txtfuentes.Visible = False, "", Me.txtfuentes), Trim(Me.txtinstitucion.Text), Trim(Me.txtcargo.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), CInt(Trim(Right(cboCondicion.Text, 2))), IIf(Me.Chkestado.value, 1, 0), lcMovNro, Trim(Me.txtApeCasada.Text)) 'EJVG20121002
        'WIOR 20120710 AGREGO Trim(Me.txtApeCasada.Text )
     'End If
          'END MADM
End Sub
'JACA END

Private Sub cmdImprimir_Click()

    Dim sCadImp As String
    Dim oPrev As previo.clsprevio
    'Dim oNPers As COMNPersona.NCOMPersona
    Dim oNPers As clases.npersona

    Set oPrev = New previo.clsprevio
    'Set oNPers = New COMNPersona.NCOMPersona
    Set oNPers = New clases.npersona
    
    sCadImp = oNPers.ImprimePersListaNegativa(IIf(lnCondicion = 7, 0, lnCondicion), IIf(lnCondicion = 7, lcMovNro, 0)) 'JACA 20110217
    
    Set oNPers = Nothing
    
    previo.Show sCadImp, "Registro de Clientes - Lista Negativa", False
    Set oPrev = Nothing
    
End Sub

Private Sub cmdNuevo_Click()
Dim oPol As COMDCredito.DCOMPoliza
Set oPol = New COMDCredito.DCOMPoliza

nTipoOperacion = 0
Call Limpiar_Controles

'txtNumPoliza.Text = oPol.RecuperaNumeroPoliza
Set oPol = Nothing

Call Habilita_Grabar(True)
Call Habilita_Datos(True)
    Me.cmdBuscarPersona.Visible = True 'WIOR 20120710
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    cmdSalir.Enabled = False
    
    cmdGrabar.Enabled = True
    'cmdCancelar.Enabled = True
    cmdexaminar.Enabled = False
    cmdCargaLote.Enabled = False 'WIOR 20130810
    txtNumDoc.SetFocus
 Call borra_variables
End Sub

Private Sub cmdsalir_Click()
     Call borra_variables
    Unload Me
End Sub

'JACA 20110217
Private Sub cmdVisitasEliminar_Click()
    If MsgBox("¿¿Está seguro de eliminar la selección actual??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
       If feJustificanegativa.TextMatrix(feJustificanegativa.row, 2) <> "" Then
           feJustificanegativa.EliminaFila feJustificanegativa.row
       End If
    End If
End Sub

Private Sub cmdVisitasNuevo_Click()
    feJustificanegativa.lbEditarFlex = True
    feJustificanegativa.AdicionaFila
    feJustificanegativa.SetFocus
    feJustificanegativaNoMoverdeFila = feJustificanegativa.Rows - 1 'WIOR 20120329
End Sub

Private Sub cmdVisitasNuevo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    feJustificanegativa.lbEditarFlex = True
    feJustificanegativa.AdicionaFila
    feJustificanegativa.SetFocus
    feJustificanegativaNoMoverdeFila = -1 'WIOR 20120329
End If
End Sub

'JACA END
Private Sub Form_Load()
Dim oCons As COMDConstantes.DCOMConstantes
Dim o1Cons As COMDConstantes.DCOMAgencias

Dim rs As ADODB.Recordset
Dim R1 As ADODB.Recordset
Dim R2 As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes
Set R1 = oCons.RecuperaConstantes(1052)
Set R2 = oCons.RecuperaConstantes(9072)
Set oCons = Nothing

Set o1Cons = New COMDConstantes.DCOMAgencias
Set rs = o1Cons.ObtieneAgencias()
Set o1Cons = Nothing

Call CentraForm(Me)

Call Llenar_Combo_con_Recordset(R2, cboCondicion)
'Call Cargar_Datos_Objetos_PersonasNegativas 'WIOR 20120321
Call Habilita_Grabar(False)
Call Habilita_Datos(False)
 Me.cmdBuscarPersona.Visible = False 'WIOR 20120710
End Sub

Sub Habilita_Grabar(ByVal pbHabilita As Boolean)
    cmdGrabar.Visible = pbHabilita
    cmdNuevo.Visible = Not pbHabilita
End Sub

'WIOR 20120322-INICIO
Private Sub feJustificanegativa_RowColChange()
Dim oPersonas As COMDPersona.DCOMPersonas
Dim oConstante As COMDConstantes.DCOMConstantes
feJustificanegativaNoMoverdeFila = feJustificanegativa.row
   If feJustificanegativa.lbEditarFlex Then
        If feJustificanegativaNoMoverdeFila <> -1 Then
            feJustificanegativa.row = feJustificanegativaNoMoverdeFila
        End If
        Set oPersonas = New COMDPersona.DCOMPersonas
        Set oConstante = New COMDConstantes.DCOMConstantes
        Select Case feJustificanegativa.Col
            Case 6 'Departamento
                feJustificanegativa.CargaCombo oPersonas.CargarUbicacionesGeograficas(True, 5, "04028")
            Case 7 'Tipo de Justificacion y Fuente
                feJustificanegativa.CargaCombo oConstante.RecuperaConstantes(9991)
        End Select
        Set oConstante = Nothing
        Set oPersonas = Nothing
    End If
End Sub
'WIOR - FIN

Function Valida_Datos() As Boolean
Valida_Datos = True

If Me.txtNumDoc.Enabled Then
    If Len(Me.txtNumDoc.Text) = 0 Then
        If MsgBox("¿Está seguro de NO registrar el Número del documento?", vbQuestion + vbYesNo) = vbNo Then
            Me.txtNumDoc.SetFocus
            Valida_Datos = False
            Exit Function
        End If
    End If
End If

If Me.OptTR(0).value = True Then
    If Len(Trim(Me.txtNomRazSocial.Text)) = 0 Or Len(Trim(Me.txtApePat.Text)) = 0 And (Len(Trim(Me.txtApeCasada.Text)) = 0 Or Len(Trim(Me.txtApeMat.Text)) = 0) Then
        MsgBox "Debe indicar el Nombre y los apellidos", vbInformation, "Mensaje"
        txtNomRazSocial.SetFocus
        Valida_Datos = False
        Exit Function
    End If
Else
    If Len(Me.txtNomRazSocial.Text) = 0 Then
        MsgBox "Debe indicar la razón social", vbInformation, "Mensaje"
        txtNomRazSocial.SetFocus
        Valida_Datos = False
        Exit Function
    End If
End If

'MADM 20100524
If Not (Me.txtfuentes.Visible = False) Then
'    If Len(Me.txtfuentes.Text) = 0 Then
'        MsgBox "Debe indicar la justificacion", vbInformation, "Mensaje"
'        txtJustificacion.SetFocus
'        Valida_Datos = False
'        Exit Function
'    End If
    
    If Len(Me.txtfuentes.Text) = 0 Then
        MsgBox "Debe indicar la fuente.", vbInformation, "Mensaje"
        txtfuentes.SetFocus
        Valida_Datos = False
        Exit Function
    End If
End If
'MADM 20100524

If cboCondicion.ListIndex = -1 Then
    MsgBox "Debe indicar la Condición", vbInformation, "Mensaje"
    cboCondicion.SetFocus
    Valida_Datos = False
    Exit Function
End If

If Len(Me.txtinstitucion.Text) = 0 And cboCondicion.ListIndex = 2 Then
    MsgBox "Debe indicar la Institucion.", vbInformation, "Mensaje"
    txtinstitucion.SetFocus
    Valida_Datos = False
    Exit Function
End If

If Len(Me.txtcargo.Text) = 0 And cboCondicion.ListIndex = 2 Then
    MsgBox "Debe indicar el Cargo.", vbInformation, "Mensaje"
    txtcargo.SetFocus
    Valida_Datos = False
    Exit Function
End If

End Function

Sub Habilita_Datos(ByVal pbHabilita As Boolean)
  
    Me.FraTipoRea.Enabled = pbHabilita
    Me.OptTR(0).Enabled = pbHabilita
    Me.OptTR(1).Enabled = pbHabilita
    Me.Chkestado.Enabled = pbHabilita
    Me.txtNumDoc.Enabled = pbHabilita
    Me.txtNomRazSocial.Enabled = pbHabilita
    Me.txtApePat.Enabled = pbHabilita
    Me.txtApeMat.Enabled = pbHabilita
    Me.txtJustificacion.Enabled = pbHabilita
    Me.txtfuentes.Enabled = pbHabilita
    Me.cboCondicion.Enabled = pbHabilita
    Me.txtcargo.Enabled = pbHabilita
    Me.txtinstitucion.Enabled = pbHabilita
    Me.feJustificanegativa.Enabled = pbHabilita 'JACA 20110217
    Me.cmdVisitasEliminar.Enabled = pbHabilita
    Me.cmdVisitasNuevo.Enabled = pbHabilita
    lcMovNro = ""
    Me.txtApeCasada.Enabled = pbHabilita 'WIOR 20120710
End Sub
'JGPA20191217 - Obs. [NAGL]
Private Sub Form_Unload(Cancel As Integer)
    If fnFrmPersona = 1 Then
        If fbSalir = False Then
            MsgBox ("Aún no completó el Registro Preventivo LAFT")
            Cancel = 1
         End If
    End If
End Sub

Sub Habilita_Modifica_Datos(ByVal pbHabilita As Boolean)

    Me.txtNomRazSocial.Enabled = pbHabilita
    Me.txtApePat.Enabled = pbHabilita
    Me.txtApeMat.Enabled = pbHabilita
    Me.txtJustificacion.Enabled = pbHabilita
    Me.txtfuentes.Enabled = pbHabilita
    Me.cboCondicion.Enabled = pbHabilita
    Me.txtcargo.Enabled = pbHabilita
    Me.txtinstitucion.Enabled = pbHabilita
    Me.Chkestado.Enabled = pbHabilita
    Me.feJustificanegativa.Enabled = pbHabilita 'JACA 20110217
    Me.cmdVisitasEliminar.Enabled = pbHabilita 'JACA 20110217
    Me.cmdVisitasNuevo.Enabled = pbHabilita 'JACA 20110217
    Me.txtApeCasada.Enabled = pbHabilita 'WIOR 20120710
End Sub



'JACA 20110217
Private Sub optCondicion_Click(Index As Integer)
lnCondicion = Index
End Sub
'JACA END

Private Sub OptTR_Click(Index As Integer)
    If Index = 0 Then
        Me.txtApePat.Visible = True
        Me.txtApeMat.Visible = True
        Me.Label3.Visible = True
        Me.Label5.Visible = True
        'WIOR 20120710
        Me.txtApeCasada.Visible = True
        Me.Label6.Visible = True
        'WIOR FIN
    Else
        Me.txtApePat.Visible = False
        Me.txtApeMat.Visible = False
        Me.txtApePat.Text = ""
        Me.txtApeMat.Text = ""
        'WIOR 20120710
        Me.txtApeCasada.Visible = False
        Me.Label6.Visible = False
        Me.txtApeCasada.Text = ""
        'WIOR FIN
        Me.Label3.Visible = False
        Me.Label5.Visible = False
    End If
End Sub


'WIOR 20120710********************
Private Sub txtApeCasada_Change()
If txtApeCasada.SelStart > 0 Then
        i = Len(Mid(txtApeCasada.Text, 1, txtApeCasada.SelStart))
    End If
    txtApeCasada.Text = UCase(txtApeCasada.Text)
    txtApeCasada.SelStart = i
End Sub
'WIOR FIN *************************
Private Sub txtApeMat_Change()
    'JACA 20111207*****************************************
    If txtApeMat.SelStart > 0 Then
        i = Len(Mid(txtApeMat.Text, 1, txtApeMat.SelStart))
    End If
    'JACA END**********************************************
    txtApeMat.Text = UCase(txtApeMat.Text)
    'i = Len(txtApeMat.Text)Comentado by JACA 20111207
    txtApeMat.SelStart = i
End Sub

Private Sub txtApeMat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
Else
'    KeyAscii = SoloLetras(KeyAscii)'Comentado by JACA 20111207
End If
End Sub

Private Sub txtApePat_Change()
    'JACA 20111207*****************************************
    If txtApePat.SelStart > 0 Then
        i = Len(Mid(txtApePat.Text, 1, txtApePat.SelStart))
    End If
    'JACA END**********************************************
    
    txtApePat.Text = UCase(txtApePat.Text)
    'i = Len(txtApePat.Text)Comentado by JACA 20111207
    txtApePat.SelStart = i
End Sub

Private Sub txtApePat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
Else
'    KeyAscii = SoloLetras(KeyAscii)'Comentado by JACA 20111207
End If
End Sub

Private Sub txtcargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtfuentes_Change()
    txtfuentes.Text = UCase(txtfuentes.Text)
    i = Len(txtfuentes.Text)
    txtfuentes.SelStart = i
End Sub

Private Sub txtfuentes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtinstitucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtJustificacion_Change()
    txtJustificacion.Text = UCase(txtJustificacion.Text)
    i = Len(txtJustificacion.Text)
    txtJustificacion.SelStart = i
End Sub

Private Sub txtJustificacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtNomRazSocial_Change()
    'JACA 20111207*****************************************
    If txtNomRazSocial.SelStart > 0 Then
        i = Len(Mid(txtNomRazSocial.Text, 1, txtNomRazSocial.SelStart))
    End If
    'JACA END**********************************************
    
    
    txtNomRazSocial.Text = UCase(txtNomRazSocial.Text)
'    i = Len(txtNomRazSocial.Text)Comentado by JACA 20111207
    txtNomRazSocial.SelStart = i
    
End Sub

Private Sub txtNomRazSocial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
Else
'    KeyAscii = SoloLetras(KeyAscii) 'MADM 20090928'Comentado by JACA 20111207
End If
End Sub

Private Sub txtNumDoc_Change()
    
    'JACA 20111207*****************************************
    If txtNumDoc.SelStart > 0 Then
        i = Len(Mid(txtNumDoc.Text, 1, txtNumDoc.SelStart))
    End If
    'JACA END**********************************************
    
    txtNumDoc.Text = UCase(txtNumDoc.Text)
    'i = Len(txtNumDoc.Text) comentado by JACA 20111207
    txtNumDoc.SelStart = i
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
Else
    KeyAscii = NumerosEnteros(KeyAscii) 'MADM 20090928
End If
End Sub

Public Sub Inicio(ByVal NumDoc As String, ByVal Tipo As Integer, Optional ByVal Nombres As String = "", Optional ByVal ApePat As String = "", Optional ByVal ApeMat As String = "", Optional ByVal ApeCas As String = "", Optional ByVal bfrmPersonaMant As Integer = 0) 'WIOR 20130201 AGREGO ApeCas 'JGPA20191217 Added bfrmPersonaMant Obs. [NAGL]"", Optional ByVal ApePat As String = "", Optional ByVal ApeMat As String = "", Optional ByVal ApeCas As String = "") 'WIOR 20130201 AGREGO ApeCas
    Dim R2 As ADODB.Recordset
    Dim oCons As COMDConstantes.DCOMConstantes
    
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R2 = oCons.RecuperaConstantes(9072)
    fnFrmPersona = bfrmPersonaMant 'JGPA20191217 - Obs. [NAGL]
        
    If Tipo = 2 Then
        Me.txtApePat.Visible = False
        Me.txtApeMat.Visible = False
        Me.txtinstitucion.Visible = False
        Me.txtcargo.Visible = False
        Me.txtApePat.Text = ""
        Me.txtApeMat.Text = ""
        Me.txtinstitucion.Text = ""
        Me.txtcargo.Text = ""
        Me.Label3.Visible = False
        Me.Label5.Visible = False
        Me.Label8.Visible = False
        Me.Label9.Visible = False
        Me.txtApeCasada.Visible = False 'WIOR 20130201
        Me.Label6.Visible = False 'WIOR 20130201
    Else
        
        Me.txtApePat.Visible = True
        Me.txtApeMat.Visible = True
        Me.Label3.Visible = True
        Me.Label5.Visible = True
        Me.txtApeCasada.Visible = True 'WIOR 20130201
        Me.Label6.Visible = True 'WIOR 20130201
    End If
    
    Me.txtNumDoc.Text = NumDoc
    Me.txtNomRazSocial.Text = Nombres
    Me.txtApePat.Text = ApePat
    Me.txtApeMat.Text = ApeMat
    Me.txtApeCasada.Text = ApeCas 'WIOR 20130201
    fnTpoDoc = Tipo
    
    Call Llenar_Combo_con_Recordset(R2, cboCondicion)
    
    Call Habilita_Grabar(True)
    Call Habilita_Datos(False)

    Me.cboCondicion.Enabled = False
    Me.cboCondicion.ListIndex = 2
    Me.txtcargo.Enabled = True
    Me.txtinstitucion.Enabled = True
    
    Me.txtJustificacion.Visible = False
    Me.txtfuentes.Visible = False
'    Me.Label6.Visible = False
'    Me.Label7.Visible = False
    
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    cmdSalir.Enabled = False
    
    cmdGrabar.Enabled = True
    'cmdCancelar.Enabled = True
    cmdexaminar.Enabled = False

 'JGPA20191217 - Obs. [NALG] No se debe cancelar si viene del registro de persona
    If fnFrmPersona = 1 Then
       cmdCancelar.Enabled = False
       cmdCargaLote.Enabled = False
    End If
    'End JGPA20191217---------

     Me.Show 1
End Sub
