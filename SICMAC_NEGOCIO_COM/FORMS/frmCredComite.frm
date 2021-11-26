VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredComite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Comites"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "frmCredComite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   4320
      TabIndex        =   18
      ToolTipText     =   "Cancelar Todos los cambios Realizados"
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
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
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Buscar datos de comites"
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   2880
      TabIndex        =   17
      ToolTipText     =   "Grabar y/o Editar Todos los cambios Realizados"
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
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
      Left            =   5640
      TabIndex        =   20
      ToolTipText     =   "Salir de la interface"
      Top             =   4200
      Width           =   975
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
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Restablecer controles"
      Top             =   4200
      Width           =   975
   End
   Begin TabDlg.SSTab SSTDatosGen 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   -2147483648
      TabCaption(0)   =   "Datos &Generales"
      TabPicture(0)   =   "frmCredComite.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFecIncRuc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtFec"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboagencia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtcomite"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbocoordinador"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbocomite"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&Analistas"
      TabPicture(1)   =   "frmCredComite.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCargar"
      Tab(1).Control(1)=   "Check1"
      Tab(1).Control(2)=   "LstAnalista"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Datos Comites"
      TabPicture(2)   =   "frmCredComite.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCodigo"
      Tab(2).Control(1)=   "txtDescripcionCampanas"
      Tab(2).Control(2)=   "cmdEliminar"
      Tab(2).Control(3)=   "MshComite"
      Tab(2).Control(4)=   "lblmensaje"
      Tab(2).Control(5)=   "Label10"
      Tab(2).Control(6)=   "Label9"
      Tab(2).ControlCount=   7
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtDescripcionCampanas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73680
         TabIndex        =   37
         Top             =   3180
         Width           =   3945
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69600
         TabIndex        =   15
         ToolTipText     =   "Eliminar Comites"
         Top             =   3120
         Width           =   1020
      End
      Begin VB.ComboBox cbocomite 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1620
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69480
         TabIndex        =   11
         Top             =   1020
         Width           =   1020
      End
      Begin VB.ComboBox cbocoordinador 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3360
         Width           =   4575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Todos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.ListBox LstAnalista 
         Enabled         =   0   'False
         Height          =   2985
         Left            =   -74880
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   900
         Width           =   5265
      End
      Begin VB.TextBox txtcomite 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   1680
         Width           =   3255
      End
      Begin VB.ComboBox cboagencia 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
         Width           =   3375
      End
      Begin VB.TextBox txtPersDireccDomicilio 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73740
         MaxLength       =   100
         TabIndex        =   28
         Top             =   2190
         Width           =   5200
      End
      Begin VB.ComboBox cmbPersDireccCondicion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73740
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2550
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ubicacion Geografica"
         Height          =   1665
         Left            =   -74865
         TabIndex        =   6
         Top             =   420
         Width           =   7365
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "frmCredComite.frx":035E
            Left            =   2235
            List            =   "frmCredComite.frx":0360
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1140
            Width           =   2190
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "frmCredComite.frx":0362
            Left            =   4680
            List            =   "frmCredComite.frx":0364
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   540
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            ItemData        =   "frmCredComite.frx":0366
            Left            =   4680
            List            =   "frmCredComite.frx":0368
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1155
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmCredComite.frx":036A
            Left            =   2250
            List            =   "frmCredComite.frx":036C
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   540
            Width           =   2175
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmCredComite.frx":036E
            Left            =   210
            List            =   "frmCredComite.frx":0370
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   525
            Width           =   1815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   2235
            TabIndex        =   26
            Top             =   900
            Width           =   600
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   4680
            TabIndex        =   25
            Top             =   915
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            Height          =   195
            Left            =   210
            TabIndex        =   24
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            Height          =   195
            Left            =   2265
            TabIndex        =   23
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            Height          =   195
            Left            =   4695
            TabIndex        =   22
            Top             =   285
            Width           =   750
         End
      End
      Begin MSMask.MaskEdBox txtFec 
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Top             =   2520
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshComite 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   14
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblmensaje 
         Caption         =   "Para eliminar escoja Agencia en la pestaña Datos Grales, y escoja de la lista un elemento..."
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   3720
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74460
         TabIndex        =   40
         Top             =   2820
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   39
         Top             =   3240
         Width           =   1125
      End
      Begin VB.Label Label8 
         Caption         =   "Coordinador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblFecIncRuc 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Registro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   35
         Top             =   2160
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblPersDireccDomicilio 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   -74820
         TabIndex        =   32
         Top             =   2220
         Width           =   630
      End
      Begin VB.Label lblPersDireccCondicion 
         AutoSize        =   -1  'True
         Caption         =   "Condicion"
         Height          =   195
         Left            =   -74820
         TabIndex        =   31
         Top             =   2625
         Width           =   705
      End
      Begin VB.Label Label13 
         Caption         =   "Valor Comercial U$"
         Height          =   240
         Left            =   -71580
         TabIndex        =   30
         Top             =   2625
         Width           =   1440
      End
      Begin VB.Label lblRefDomicilio 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   195
         Left            =   -74820
         TabIndex        =   29
         Top             =   3030
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmCredComite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Opc As Integer
Dim est As Integer
'Para flexEdit de Relacion con Personas
Dim cmdPersRelaEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FERelPersNoMoverdeFila As Integer

Private Sub CboAgencia_Click()
Dim rs As ADODB.Recordset
Dim oGen As COMDCredito.DCOMCreditos
Dim nro As Integer
Dim str As String
  
  If CboAgencia.ListIndex <> -1 And Opc = 2 Then
        CargarComites 'llena cbocomite, siempre y fuerzo a poner en blanco responsable y lista
        If est <> 1 Then
            CargarComites_grilla 'llena grilla muestra el tab y el boton eliminar
        End If
  ElseIf CboAgencia.ListIndex <> -1 And Opc = 1 Then 'para la opcion nuevo
         
         Set oGen = New COMDCredito.DCOMCreditos
         Set rs = New ADODB.Recordset
         Set rs = oGen.Num_Comite_Agencia(CInt(Right((Me.CboAgencia.Text), 2))) 'devuleve el numero de comite x agencia
         Set oGen = Nothing
        
        If Not rs.EOF And Not rs.BOF Then
            nro = rs!Total
        End If
        
        If nro >= 1 Then
           CargaAnalistas_filtro 'filtrado de analistas
           Else
           CargaAnalistas 'todos los analistas
        End If
        CargaAnalistas_cbo ' llena los responsables
  End If
  
End Sub

Private Sub cboAgencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If CboAgencia.ListIndex <> -1 Then
           If Opc = 2 Then
            Me.txtcomite.SetFocus
           End If
        Else
            CboAgencia.SetFocus
        End If
    End If
End Sub

Private Sub cbocomite_click()
If Opc = 2 And Me.cbocomite.ListIndex <> -1 Then 'para buscar
      CargaDatos_Comite
      CargarComites_grilla_1
End If

End Sub
'busca
Sub CargaDatos_Comite()
Dim rs As ADODB.Recordset
Dim oGen As COMDCredito.DCOMCreditos
Dim Res As String
Dim i As Integer
Dim Pos As Integer
     
     Set oGen = New COMDCredito.DCOMCreditos
     Set rs = New ADODB.Recordset
     'Set rs = oGen.CargaDatos_Comite(CInt(Right((Me.cbocomite.Text), 2)))
     Set rs = oGen.CargaDatos_Comite(CInt(Trim(Right((Me.cbocomite.Text), 4)))) 'FRHU 20150326
     Set oGen = Nothing
       
     If Not rs.EOF And Not rs.BOF Then
        Me.cmdCargar.Enabled = True
        Me.Check1.Enabled = False
        txtFec.Enabled = True
        txtFec.Text = Format(rs!fec_comite, "dd/mm/yyyy")
        Res = rs!res_comite
        CargaAnalistas_cbo ' llena responsable
        'ANALISTAS Q NO ESTEN EN OTRO COMITE ************************************************************
        limpia_lista
        CargaAnalistas_filtro1 'llena lista
        '------------------------------------------------------------------------------------------------
        Do While Not rs.EOF
                     For i = 0 To Me.cbocoordinador.ListCount - 1
                        If (Trim(Right(Me.cbocoordinador.List(i), 18)) = Res) Then
                           Pos = i
                        End If
                     Next i
            rs.MoveNext
        Loop
        Me.cbocoordinador.ListIndex = Pos
    End If
    rs.Close
End Sub

Private Sub Check1_Click()
Dim bCheck As Boolean
Dim i As Integer
       
    If Check1.value = 0 Then
        bCheck = False
    Else
        bCheck = True
    End If
    
    If LstAnalista.ListCount <= 0 Then
        Exit Sub
    End If
    
    For i = 0 To LstAnalista.ListCount - 1
        LstAnalista.Selected(i) = bCheck
    Next i
End Sub


Sub estado_controles(xb As Boolean)
        If Opc = 1 Then
        Me.txtcomite.Enabled = xb
        Else
        Me.cbocomite.Enabled = xb
        End If
        Me.txtFec.Enabled = xb
        Me.CboAgencia.Enabled = xb
        Me.cbocoordinador.Enabled = xb
        Me.LstAnalista.Enabled = xb
        Me.Check1.Enabled = xb
        Me.cmdCargar.Enabled = xb
End Sub

Sub limpia_controles_texts()
        If Opc = 1 Then
        Me.txtcomite.Text = ""
        Else
        cbocomite.ListIndex = -1
        End If
        CboAgencia.ListIndex = -1
        cbocoordinador.ListIndex = -1
        Me.txtFec.Text = "__/__/____"
        Me.Check1.value = False
        
End Sub

Private Sub cmdBuscar_Click()
    Opc = 2 ' se define busqueda
    SSTDatosGen.Tab = 0
    Me.cmdGrabar.Caption = "&Editar"
    Me.CmdBuscar.Enabled = False
    Me.CmdNuevo.Enabled = False
    Me.cmdGrabar.Enabled = True
    Me.cmdCancelar.Enabled = True
    'habilitar controles para busqueda
    Me.cbocomite.Visible = True
    Me.cbocomite.Enabled = True
    Me.txtcomite.Visible = False
    Me.cbocoordinador.Enabled = True
    Me.cbocoordinador.Clear
    Me.cbocomite.Clear
    Me.CboAgencia.Enabled = True
    Me.CboAgencia.SetFocus
    
     CargaAgencia
End Sub

Private Sub cmdCancelar_Click()
'vuelve a la normalidad
limpia_controles_texts
estado_controles (False)
est = 0
SSTDatosGen.Tab = 0
SSTDatosGen.TabVisible(2) = False
Me.txtCodigo.Text = ""
Me.txtDescripcionCampanas.Text = ""
'Me.FERelPers.Clear
CmdEliminar.Enabled = False
LstAnalista.Clear

        If Opc = 1 Then
            Me.cmdGrabar.Enabled = False
         ElseIf Opc = 2 Then
            Me.cbocomite.Clear
            Me.CboAgencia.Clear
            Me.cbocoordinador.Clear
            Me.cbocomite.Visible = False
            Me.txtcomite.Visible = True
            Me.cmdGrabar.Caption = "&Grabar"
            Me.cmdGrabar.Enabled = False
        End If
        Me.CmdBuscar.Enabled = True
        Me.CmdNuevo.Enabled = True
        Me.CmdNuevo.SetFocus
        
End Sub

Private Sub cmdCargar_Click()
Dim i As Integer
Dim R As ADODB.Recordset
Dim oGen As COMDCredito.DCOMCreditos
Dim pcPersCod As String

 If Me.cbocomite.ListIndex < 0 Then
        MsgBox "Seleccione o Cambie de Comite"
        Exit Sub
 End If
    
Set R = New ADODB.Recordset
Set oGen = New COMDCredito.DCOMCreditos

'Set R = oGen.CargaAnalistasAgencia_Det(CInt(Right(CboAgencia.Text, 2)), CInt(Right(Me.cbocomite.Text, 2)))
Set R = oGen.CargaAnalistasAgencia_Det(CInt(Right(CboAgencia.Text, 2)), CInt(Trim(Right(Me.cbocomite.Text, 4)))) 'FRHU 20150326
 Me.LstAnalista.Enabled = True
 Me.cmdCargar.Enabled = False
 Me.Check1.Enabled = True
 
    If Not R.EOF And Not R.BOF Then 'indica que hay alguno asignado
        Do While Not R.EOF
            pcPersCod = R!cPersCod
                    For i = 0 To LstAnalista.ListCount - 1
                        If (Trim(Right(LstAnalista.List(i), 18)) = pcPersCod) Then
                            LstAnalista.Selected(i) = True
                        End If
                    Next i
            R.MoveNext
        Loop
    End If
    
End Sub

Private Sub cmdEliminar_Click()
 Dim oCredD As COMDCredito.DCOMCreditos
 Set oCredD = New COMDCredito.DCOMCreditos
  If txtCodigo <> "" And txtDescripcionCampanas <> "" Then
        If MsgBox("Esta Seguro que Desea Eliminar el Comite", vbQuestion + vbYesNo, "Aviso") = vbYes Then
               If oCredD.EliminacionComitedet(CInt(txtCodigo)) Then
                       If oCredD.EliminacionComite(CInt(txtCodigo)) Then
                            MsgBox "Los datos se eliminaron Correctamente "
                           Set oCredD = Nothing
                       End If
               Else
                       MsgBox "Existe un error en la elimininacion " & err.Description, vbInformation, "AVISO"
               End If
               Call cmdCancelar_Click
           End If
Else
    MsgBox "No ha seleccionado ningun registro", vbInformation, "AVISO"
End If
End Sub

Private Sub CmdGrabar_Click()
Dim i As Integer
Dim xc As Boolean
Dim id As String
Dim rs As ADODB.Recordset
Dim oGen As COMDCredito.DCOMCreditos
Set rs = New ADODB.Recordset
Set oGen = New COMDCredito.DCOMCreditos
xc = False
Dim nContador As Integer
'VALIDACION DE LOS DATOS DE COMITE
If est = 1 Then
        MsgBox ("Esta agencia no tiene registrado Comite, Cancele y Registre !!")
        Me.cmdCancelar.SetFocus
Else
         If Me.txtFec.Text = "" Then
                    MsgBox "Fecha no válida", vbInformation, "Aviso"
                    Exit Sub
        ElseIf Not IsDate(txtFec) Then
                    MsgBox "Fecha no válida", vbInformation, "Aviso"
                    Exit Sub
        End If
            
        If Opc = 1 Then 'nuevo
            If Me.txtcomite.Text = "" Or Me.txtFec.Text = "" Or Me.cbocoordinador.ListIndex = -1 Or Me.CboAgencia.ListIndex = -1 Then
                cmdNuevo_Click
                MsgBox ("Complete Datos, Verifique !!")
                Exit Sub
            End If
            
        Else
            If Me.CboAgencia.ListIndex = -1 Or Me.cbocomite.ListIndex = -1 Or Me.cbocoordinador.ListIndex = -1 Or Me.txtFec.Text = "" Then
                MsgBox ("seleccione Agencia y/o Comite para iniciar búsqueda !!")
                Exit Sub
            End If
        End If
        
        Dim oCredD As COMDCredito.DCOMCreditos
         Set oCredD = New COMDCredito.DCOMCreditos
            If Opc = 1 Then
            If LstAnalista.ListCount = 0 Then
                    oCredD.Insertar_Datos_Comite CInt(Right(CboAgencia.Text, 2)), Trim(Me.txtcomite.Text), Trim(Right(Me.cbocoordinador.Text, 15)), Me.txtFec.Text
                    SSTDatosGen.Tab = 0
                    MsgBox ("Comite Insertado con Exito")
                    limpia_controles_texts
                    LstAnalista.Clear
                    estado_controles (False)
                    Me.cmdGrabar.Enabled = False
                    Me.cmdCancelar.Enabled = False
                    Me.CmdNuevo.Enabled = True
                    Me.CmdBuscar.Enabled = True
                    Me.CmdNuevo.SetFocus
                     Exit Sub
            Else
                    oCredD.Insertar_Datos_Comite CInt(Right(CboAgencia.Text, 2)), Trim(Me.txtcomite.Text), Trim(Right(Me.cbocoordinador.Text, 15)), Me.txtFec.Text
                    Set rs = oGen.CargaDatos_Comite_ult(CInt(Right(CboAgencia.Text, 2)), Trim(Me.txtcomite.Text))
                    If Not rs.EOF And Not rs.BOF Then
                          id = rs("id_comite")
                   Else
                          Exit Sub
                   End If
                   
                   For i = 0 To LstAnalista.ListCount - 1
                       If LstAnalista.Selected(i) = True Then
                          nContador = nContador + 1
                          oCredD.Insertar_Datos_Comite_det CInt(id), CInt(Right(CboAgencia.Text, 2)), Trim(Right(LstAnalista.List(i), 18)), nContador
                          xc = True
                      End If
                    Next i
                   nContador = 0
                   If xc Then
                       MsgBox ("Comite Insertado con Analistas Satisfactoriamente")
                       SSTDatosGen.Tab = 0
                       Me.cmdCancelar.SetFocus
                   Else
                       MsgBox ("Comite Insertado con Exito")
                       SSTDatosGen.Tab = 0
                       Me.cmdCancelar.SetFocus
                   End If
                   
                   limpia_controles_texts
                   LstAnalista.Clear
                   estado_controles (False)
                   Me.cmdGrabar.Enabled = False
                   Me.cmdCancelar.Enabled = False
                   Me.CmdNuevo.Enabled = True
                   Me.CmdBuscar.Enabled = True
                   Me.CmdNuevo.SetFocus
            End If
            
            Else 'edicion y no borrar los datos de los texts
                   'oCredD.Actualizar_Datos_Comite CInt(Right(Me.cbocomite.Text, 2)), CInt(Right(CboAgencia.Text, 2)), Trim(Me.cbocomite.Text), Trim(Right(Me.cbocoordinador.Text, 15)), Me.txtFec.Text
                   oCredD.Actualizar_Datos_Comite CInt(Trim(Right(Me.cbocomite.Text, 4))), CInt(Right(CboAgencia.Text, 2)), Trim(Me.cbocomite.Text), Trim(Right(Me.cbocoordinador.Text, 15)), Me.txtFec.Text 'FRHU 20150326
                   CargarComites_grilla_1
                    If LstAnalista.ListCount <= 0 Then
                            MsgBox ("La lista no tiene elementos, Cambie de Comite u Agencia")
                            Exit Sub
                    Else
                            
                            For i = 0 To LstAnalista.ListCount - 1
                                If LstAnalista.Selected(i) = True Then
                                    nContador = nContador + 1
                                         'oCredD.Insertar_Datos_Comite_det CInt(Right(Me.cbocomite.Text, 2)), CInt(Right(CboAgencia.Text, 2)), Trim(Right(LstAnalista.List(i), 18)), nContador
                                         oCredD.Insertar_Datos_Comite_det CInt(Trim(Right(Me.cbocomite.Text, 4))), CInt(Right(CboAgencia.Text, 2)), Trim(Right(LstAnalista.List(i), 18)), nContador 'FRHU 20150326
                                End If
                            Next i
                                    If nContador = 0 Or Me.cmdCargar.Enabled = True Then
                                         If MsgBox("No cargo Analistas y/o Seleccionó alguno, se procederá a borrar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                                            'If oCredD.EliminacionComitedet(CInt(Right(Me.cbocomite.Text, 2))) Then
                                            If oCredD.EliminacionComitedet(CInt(Trim(Right(Me.cbocomite.Text, 4)))) Then 'FRHU 20150326
                                                    MsgBox ("Comite Modificado Correctamente")
                                                    Me.cmdCargar.Enabled = True
                                                    limpia_lista
                                                    Me.LstAnalista.Enabled = False
                                                    Me.Check1.value = False
                                                    Me.Check1.Enabled = False
                                                    Me.cmdCancelar.SetFocus
                                                    SSTDatosGen.Tab = 0
                                                    Exit Sub
                                            End If
                                        End If
                                    End If
                                nContador = 0
                                MsgBox ("Comite Insertado con Analistas Satisfactoriamente")
                                limpia_lista
                                Me.cmdCargar.Enabled = True
                                Me.LstAnalista.Enabled = False
                                Me.Check1.value = False
                                Me.Check1.Enabled = False
                                Me.cmdCancelar.SetFocus
                                SSTDatosGen.Tab = 0
                    End If
                End If
End If
End Sub
Sub limpia_lista()
Dim i As Integer
For i = 0 To LstAnalista.ListCount - 1
     LstAnalista.Selected(i) = False
Next i
End Sub

Private Sub cmdNuevo_Click()
    Opc = 1
    est = 0
    estado_controles (True)
    CargaAgencia
    CargaAnalistas_cbo
    Me.cmdGrabar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.CmdNuevo.Enabled = False
    Me.CmdBuscar.Enabled = False
    Me.cmdCargar.Enabled = False
    CmdEliminar.Enabled = False
    If Not LstAnalista.ListCount <= 0 Then
        LstAnalista.Clear
    End If
    
End Sub

Private Sub Command3_Click()
Me.Height = 7215
End Sub

Public Sub LimpiaFlex(ByRef Flex As Control)
    Dim i As Integer
    Flex.Rows = 2
    For i = 0 To Flex.Cols - 1
        Flex.TextMatrix(1, i) = ""
    Next i
End Sub

Private Sub cmdsalir_Click()
    Opc = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Opc = 0
    Me.Width = 7005
    Me.Height = 5220
    SSTDatosGen.TabVisible(2) = False
    ConfigurarMShComite
End Sub

Sub ConfigurarMShComite()
 MshComite.Clear
    MshComite.Cols = 3
    MshComite.Rows = 2
    
    With MshComite
        .TextMatrix(0, 0) = "Codigo"
        .TextMatrix(0, 1) = "Descripcion"
        .TextMatrix(0, 2) = "Responsable"
        
        .ColWidth(1) = 2500
        .ColWidth(2) = 3000
    End With
End Sub

Private Sub CargaAgencia()
'Dim loCargaAg As COMDColocPig.DCOMColPFunciones
Dim loCargaAg As COMDConstantes.DCOMAgencias 'FRHU 20150326
Dim lrAgenc As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    'FRHU 20150326
    'Set loCargaAg = New COMDColocPig.DCOMColPFunciones
    'Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = New COMDConstantes.DCOMAgencias
    Set lrAgenc = loCargaAg.ObtieneAgencias()
    'FIN FRHU 20150326
    Set loCargaAg = Nothing
    Call llenar_cbo_agencia(lrAgenc, CboAgencia)
    Exit Sub

ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Sub llenar_cbo_agencia(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    'pcboObjeto.AddItem Trim(pRs!cAgeDescripcion) & Space(100) & Trim(str(pRs!cAgeCod))
    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & Space(100) & pRs!nConsValor 'FRHU 20150326
    pRs.MoveNext
Loop
pRs.Close
End Sub

Sub llenar_cbo_analista(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cPersNombre) & Space(100) & Trim(str(pRs!cPersCod))
    pRs.MoveNext
Loop
pRs.Close
End Sub
Sub llenar_cbo_analista_p(pRs As ADODB.Recordset, pcboObjeto As ComboBox, P As String)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cPersNombre) & Space(100) & Trim(str(pRs!cPersCod))
    pRs.MoveNext
Loop
    pcboObjeto.ListIndex = CInt(P)
pRs.Close
End Sub

Sub llenar_cbo_comite(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!nom_comite) & Space(100) & Trim(str(pRs!id_comite))
    pRs.MoveNext
Loop
pRs.Close
End Sub

Public Sub CargarComites_grilla()
Dim rs As ADODB.Recordset
Dim oGen As COMDCredito.DCOMCreditos

    On Error GoTo ERRORCargarComites
    
    Set oGen = New COMDCredito.DCOMCreditos
        Set rs = New ADODB.Recordset
        Set rs = oGen.CargaComiteAgencia(CInt(Right(CboAgencia.Text, 2)))
    Set oGen = Nothing
    
     If Not (rs.EOF And rs.BOF) Then
        Call llenar_grilla_comite(rs)
        Me.CmdEliminar.Enabled = True
        SSTDatosGen.TabVisible(2) = True
        SSTDatosGen.Tab = 0
        lblMensaje.Visible = False
    Else
        Me.CmdEliminar.Enabled = False
        SSTDatosGen.TabVisible(2) = False
        SSTDatosGen.Tab = 0
        lblMensaje.Visible = False
    End If
    Set rs = Nothing
    Exit Sub
ERRORCargarComites:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub
Public Sub CargarComites_grilla_1()
Dim rs As ADODB.Recordset
Dim oGen  As COMDCredito.DCOMCreditos

    On Error GoTo ERRORCargarComites
    
    Set oGen = New COMDCredito.DCOMCreditos
        Set rs = New ADODB.Recordset
        'Set rs = oGen.CargaComiteAgenciaxComite(CInt(Right(CboAgencia.Text, 2)), CInt(Right(Me.cbocomite.Text, 2)))
        Set rs = oGen.CargaComiteAgenciaxComite(CInt(Right(CboAgencia.Text, 2)), CInt(Trim(Right(Me.cbocomite.Text, 4)))) 'FRHU 20150326
    Set oGen = Nothing

    Me.txtCodigo.Text = ""
    Me.txtDescripcionCampanas.Text = ""
    Me.CmdEliminar.Enabled = False
    lblMensaje.Visible = True
    
     If Not (rs.EOF And rs.BOF) Then
        Call llenar_grilla_comite(rs)
    End If
    Set rs = Nothing
    Exit Sub
ERRORCargarComites:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub

Public Sub CargarComites()
Dim rs As ADODB.Recordset
Dim oGen As COMDCredito.DCOMCreditos

    On Error GoTo ERRORCargarComites
    
    Set oGen = New COMDCredito.DCOMCreditos
        Set rs = New ADODB.Recordset
        Set rs = oGen.CargaComiteAgencia(CInt(Right(CboAgencia.Text, 2)))
    Set oGen = Nothing
    
     If Not (rs.EOF And rs.BOF) Then
        Call llenar_cbo_comite(rs, Me.cbocomite)
        If cbocoordinador.ListIndex <> -1 Then ' fuerzo a poner en blanco responsable y lista
            cbocoordinador.ListIndex = -1
            Me.txtFec.Text = "__/__/____"
            Me.Check1.value = False
        End If
        If Me.LstAnalista.ListCount > 0 Then
            Me.LstAnalista.Clear
        End If
        est = 0
    Else
        MsgBox ("Esta agencia no tiene registrado Comite, Cancele y Registre !!")
'        Me.FERelPers.Clear
        SSTDatosGen.TabVisible(2) = False
        SSTDatosGen.Tab = 0
        Me.cbocomite.Clear
        Me.cbocoordinador.Clear
        Me.LstAnalista.Clear
        Me.cmdCancelar.SetFocus
        Me.Check1.Enabled = False
        Me.txtFec.Text = "__/__/____"
        est = 1
        Exit Sub
    End If
    Set rs = Nothing
    Exit Sub
ERRORCargarComites:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub
Sub llenar_grilla_comite(ByVal rs As ADODB.Recordset)
Dim i As Integer
ConfigurarMShComite
      Do Until rs.EOF
        With MshComite
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = rs!id_comite
            .TextMatrix(.Rows - 2, 1) = rs!nom_comite
            .TextMatrix(.Rows - 2, 2) = rs!cPersNombre
        End With
        rs.MoveNext
    Loop
 rs.Close
 
End Sub
'modificada para no cargar responsable de comites
Private Sub CargaAnalistas()
Dim rs As ADODB.Recordset
Dim sAnalistas As String
Dim oGen As COMDCredito.DCOMCreditos

    On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New COMDCredito.DCOMCreditos
        Set rs = New ADODB.Recordset
        Set rs = oGen.CargaAnalistasAgencia(Right(CboAgencia.Text, 2))
    Set oGen = Nothing
    LstAnalista.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            LstAnalista.AddItem rs!cPersNombre & Space(100) & rs!cPersCod
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
'PARA FILTRAR POR AGENCIA -----------------------
Private Sub CargaAnalistas_filtro()
Dim rs As ADODB.Recordset
Dim sAnalistas As String
Dim oGen As COMDCredito.DCOMCreditos

    On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New COMDCredito.DCOMCreditos
        Set rs = New ADODB.Recordset
        Set rs = oGen.CargaAnalistasAgencia(Right(CboAgencia.Text, 2), True)
    Set oGen = Nothing
    LstAnalista.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            LstAnalista.AddItem rs!cPersNombre & Space(100) & rs!cPersCod
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
'PARA FILTRAR POR COMITE (analistas y comites)-----------------------
Private Sub CargaAnalistas_filtro1()
Dim rs As ADODB.Recordset
Dim sAnalistas As String
Dim oGen As COMDCredito.DCOMCreditos

    On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New COMDCredito.DCOMCreditos
        Set rs = New ADODB.Recordset
        'Set rs = oGen.CargaAnalistasAgenciasinComite(Right(Me.CboAgencia.Text, 2), Right(Me.cbocomite.Text, 2))
        Set rs = oGen.CargaAnalistasAgenciasinComite(Right(Me.CboAgencia.Text, 2), Trim(Right(Me.cbocomite.Text, 4))) 'FRHU 20150326
    Set oGen = Nothing
    LstAnalista.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            LstAnalista.AddItem rs!cPersNombre & Space(100) & rs!cPersCod
            rs.MoveNext
        Loop
        Me.LstAnalista.Enabled = False
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

'cargar el combo
Private Sub CargaAnalistas_cbo()
Dim rs As ADODB.Recordset
Dim sAnalistas As String
Dim oGen  As COMDConstSistema.DCOMGeneral

    On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New COMDConstSistema.DCOMGeneral
        Set rs = New ADODB.Recordset
        'Set rs = oGen.CargaAnalistas()
        Set rs = oGen.GetResponsableComite() 'FRHU 20150326
    Set oGen = Nothing
    Call llenar_cbo_analista(rs, Me.cbocoordinador)
'    rs.Close
    Set rs = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
'poner el responsable
Private Sub CargaAnalistas_cbo_p(cod As String)
Dim rs As ADODB.Recordset
Dim sAnalistas As String
Dim oGen  As COMDConstSistema.DCOMGeneral

    On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New COMDConstSistema.DCOMGeneral
        Set rs = New ADODB.Recordset
        Set rs = oGen.CargaAnalistas()
    Set oGen = Nothing
    Call llenar_cbo_analista(rs, Me.cbocoordinador)
    Set rs = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub


Private Sub MshComite_Click()
 On Error GoTo ErrHandler
        With Me.MshComite
            txtCodigo.Text = .TextMatrix(.row, 0)
            txtDescripcionCampanas.Text = .TextMatrix(.row, 1)
        End With
Exit Sub
ErrHandler:
    MsgBox "Se ha producido un error al seleccionar el registro", vbInformation, "AVISO"
End Sub

