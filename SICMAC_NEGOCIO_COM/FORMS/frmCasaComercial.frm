VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCasaComercial 
   Caption         =   "Convenio Casa Comercial"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   Icon            =   "frmCasaComercial.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCasaComercial.frx":030A
   ScaleHeight     =   4320
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin TabDlg.SSTab SSTDatosGen 
         Height          =   3855
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6800
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BackColor       =   -2147483648
         TabCaption(0)   =   "Casa &Comercial"
         TabPicture(0)   =   "frmCasaComercial.frx":0614
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label11"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblFecIncRuc"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label8"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Line1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtfec1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtFec"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtdescripcion"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtnombre"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdGrabar"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Cmdnuevo"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "CmdBuscar"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "CmdSalir"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtempresa"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "CmdAsignar"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "&Asignar Convenio"
         TabPicture(1)   =   "frmCasaComercial.frx":0630
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label12"
         Tab(1).Control(1)=   "LstCampana"
         Tab(1).Control(2)=   "cboconvenio"
         Tab(1).Control(3)=   "Cmdguardar"
         Tab(1).Control(4)=   "CmdSalir1"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "&Mantenimiento"
         TabPicture(2)   =   "frmCasaComercial.frx":064C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FEEmpresas"
         Tab(2).ControlCount=   1
         Begin VB.CommandButton CmdSalir1 
            Caption         =   "Salir"
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
            Left            =   -69720
            TabIndex        =   14
            ToolTipText     =   "Registrar Convenio Casa Comercial"
            Top             =   1680
            Width           =   1020
         End
         Begin VB.CommandButton Cmdguardar 
            Caption         =   "Grabar"
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
            Left            =   -69720
            TabIndex        =   13
            ToolTipText     =   "Registrar Convenio Casa Comercial"
            Top             =   960
            Width           =   1020
         End
         Begin VB.CommandButton CmdAsignar 
            Caption         =   "Asignar"
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
            Left            =   3840
            TabIndex        =   9
            ToolTipText     =   "Asignar Convenio Casa Comercial"
            Top             =   3240
            Width           =   1020
         End
         Begin VB.TextBox txtempresa 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   1080
            Width           =   4335
         End
         Begin VB.CommandButton CmdSalir 
            Caption         =   "Salir"
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
            Left            =   5040
            TabIndex        =   10
            ToolTipText     =   "Salir"
            Top             =   3240
            Width           =   1020
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
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
            Left            =   2640
            TabIndex        =   8
            ToolTipText     =   "Busca Casa Comercial"
            Top             =   3240
            Width           =   1020
         End
         Begin VB.CommandButton Cmdnuevo 
            Caption         =   "Nuevo"
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
            Left            =   240
            TabIndex        =   1
            ToolTipText     =   "Nueva Casa Comercial"
            Top             =   3240
            Width           =   1020
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "Grabar"
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
            Left            =   1440
            TabIndex        =   7
            ToolTipText     =   "Registrar Casa Comercial"
            Top             =   3240
            Width           =   1020
         End
         Begin VB.ComboBox cboconvenio 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   -73680
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   3135
         End
         Begin VB.ListBox LstCampana 
            Enabled         =   0   'False
            Height          =   2760
            Left            =   -74760
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   960
            Width           =   4905
         End
         Begin VB.TextBox txtnombre 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   2
            Top             =   600
            Width           =   4335
         End
         Begin VB.TextBox txtPersDireccDomicilio 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73740
            MaxLength       =   100
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   2550
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            Caption         =   "Ubicacion Geografica"
            Height          =   1665
            Left            =   -74865
            TabIndex        =   17
            Top             =   420
            Width           =   7365
            Begin VB.ComboBox cmbPersUbiGeo 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               ItemData        =   "frmCasaComercial.frx":0668
               Left            =   2235
               List            =   "frmCasaComercial.frx":066A
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   1140
               Width           =   2190
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               ItemData        =   "frmCasaComercial.frx":066C
               Left            =   4680
               List            =   "frmCasaComercial.frx":066E
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   540
               Width           =   2430
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               ItemData        =   "frmCasaComercial.frx":0670
               Left            =   4680
               List            =   "frmCasaComercial.frx":0672
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   1155
               Width           =   2430
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               ItemData        =   "frmCasaComercial.frx":0674
               Left            =   2250
               List            =   "frmCasaComercial.frx":0676
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   540
               Width           =   2175
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               ItemData        =   "frmCasaComercial.frx":0678
               Left            =   210
               List            =   "frmCasaComercial.frx":067A
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   525
               Width           =   1815
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Distrito :"
               Height          =   195
               Left            =   2235
               TabIndex        =   27
               Top             =   900
               Width           =   600
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Zona : "
               Height          =   195
               Left            =   4680
               TabIndex        =   26
               Top             =   915
               Width           =   540
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Pais : "
               Height          =   195
               Left            =   210
               TabIndex        =   25
               Top             =   285
               Width           =   435
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Departamento :"
               Height          =   195
               Left            =   2265
               TabIndex        =   24
               Top             =   285
               Width           =   1095
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Provincia :"
               Height          =   195
               Left            =   4695
               TabIndex        =   23
               Top             =   285
               Width           =   750
            End
         End
         Begin VB.TextBox txtdescripcion 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   675
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   1560
            Width           =   4335
         End
         Begin MSMask.MaskEdBox txtFec 
            Height          =   300
            Left            =   1320
            TabIndex        =   5
            Top             =   2400
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
         Begin MSMask.MaskEdBox txtfec1 
            Height          =   300
            Left            =   4440
            TabIndex        =   6
            Top             =   2400
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
         Begin SICMACT.FlexEdit FEEmpresas 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   15
            Top             =   600
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   5106
            Cols0           =   4
            HighLight       =   2
            EncabezadosNombres=   "-Id-Nombre-Institucion"
            EncabezadosAnchos=   "50-400-2500-3000"
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
            ColumnasAEditar =   "X-X-X-X"
            ListaControles  =   "0-0-0-0"
            BackColor       =   -2147483644
            EncabezadosAlineacion=   "C-L-L-C"
            FormatosEdit    =   "0-0-0-0"
            ColWidth0       =   45
            RowHeight0      =   300
            CellBackColor   =   -2147483644
         End
         Begin VB.Label Label1 
            Caption         =   "Empresa"
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
            Left            =   120
            TabIndex        =   39
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Line Line1 
            X1              =   360
            X2              =   5880
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label Label12 
            Caption         =   "Comercial"
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
            Left            =   -74760
            TabIndex        =   38
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Descripcion"
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
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label lblFecIncRuc 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
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
            TabIndex        =   36
            Top             =   2400
            Width           =   660
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
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblPersDireccDomicilio 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            Height          =   195
            Left            =   -74820
            TabIndex        =   34
            Top             =   2220
            Width           =   630
         End
         Begin VB.Label lblPersDireccCondicion 
            AutoSize        =   -1  'True
            Caption         =   "Condicion"
            Height          =   195
            Left            =   -74820
            TabIndex        =   33
            Top             =   2625
            Width           =   705
         End
         Begin VB.Label Label13 
            Caption         =   "Valor Comercial U$"
            Height          =   240
            Left            =   -71580
            TabIndex        =   32
            Top             =   2625
            Width           =   1440
         End
         Begin VB.Label lblRefDomicilio 
            AutoSize        =   -1  'True
            Caption         =   "Referencia"
            Height          =   195
            Left            =   -74820
            TabIndex        =   31
            Top             =   3030
            Width           =   780
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
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
            Left            =   3600
            TabIndex        =   30
            Top             =   2400
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frmCasaComercial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nCantVin As Integer
Dim sMatCasa() As String
Dim lsCodCasa As String
Dim nPost As Integer
Dim loDatos As ADODB.Recordset
Dim loDatos1 As ADODB.Recordset
Dim nContador As Integer
Dim filaFE As Integer

Public Enum OpeMant1
    lRegistro = 1
    lMantenimiento = 2
    lConsultas = 3
    lAsignar = 4
    lCancelar = 5
End Enum
Dim nOperacion As OpeMant1

Sub llenar_casa(pRs As ADODB.Recordset, pcboObjeto As ComboBox, P As Integer)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cNombre) & Space(100) & Trim(str(pRs!id))
    pRs.MoveNext
Loop
pcboObjeto.ListIndex = P
pRs.Close
End Sub

Sub CargarCampanha()
    Dim odCamp As COMDCredito.DCOMCampanas
    Dim rs As ADODB.Recordset
    LstCampana.Clear
    
    Set odCamp = New COMDCredito.DCOMCampanas
    Set rs = odCamp.CargarCampanas
    Set odCamp = Nothing
    Do Until rs.EOF
        LstCampana.AddItem rs!cDescripcion & Space(100) & rs!idCampana
'        LstCampana.AddItem rs!cDescripcion
'        LstCampana.ItemData(LstCampana.NewIndex) = CInt(rs!IdCampana)
        rs.MoveNext
    Loop
End Sub
Sub limpiar_lista(ByVal x As Boolean)
Dim i As Integer
 For i = 0 To LstCampana.ListCount - 1
     LstCampana.Selected(i) = x
 Next
End Sub

Public Function verificar_lista(ByVal x As Boolean) As Boolean
 Dim contadorz As Integer
 Dim i As Integer
 contadorz = 0
 For i = 0 To LstCampana.ListCount - 1
     If LstCampana.Selected(i) = x Then
        contadorz = contadorz + 1
     End If
 Next
 If contadorz > 0 Then
    verificar_lista = True
 Else
    verificar_lista = False
 End If
End Function

Private Sub cboconvenio_Click()
Dim oCredD As COMDCredito.DCOMCreditos
Dim rs As ADODB.Recordset
Dim i As Integer
 Set oCredD = New COMDCredito.DCOMCreditos
 Set rs = oCredD.DevolverDatos_ConvenioCasaComercial(Right(Me.cboconvenio.Text, 2))
 nContador = 0
 limpiar_lista (False)
 
    Do While Not rs.EOF
        For i = 0 To LstCampana.ListCount - 1
             If CInt(rs!id_Camp) = CInt(Trim(Right(LstCampana.List(i), 3))) Then
                LstCampana.Selected(i) = True
                 nContador = nContador + 1
                Exit For
             End If
         Next i
    rs.MoveNext
 Loop
 Cmdguardar.Enabled = True
 
End Sub

Sub llenar_cbos_convenio(x As Integer)
    Dim oCredD As COMDCredito.DCOMCreditos
    Dim rs As ADODB.Recordset
    Set oCredD = New COMDCredito.DCOMCreditos
    Set rs = oCredD.DevolverDatos_CasaComercial(x)
    Call llenar_casa(rs, cboconvenio, x)
End Sub

Private Sub CmdAsignar_Click()
    nOperacion = lAsignar
    habilita_textos (True)
 
End Sub

Private Sub CmdBuscar_Click()
 Dim oCredD As COMDCredito.DCOMCreditos
 Set oCredD = New COMDCredito.DCOMCreditos
 
 nOperacion = lConsultas

     If CmdBuscar.Caption = "Eliminar" Then
        If valida_datos Then
            Call oCredD.Eliminar_Datos_CasaComercial(lsCodCasa)
            If filaFE > 0 Then
                FEEmpresas.EliminaFila (filaFE)
                MsgBox "Convenio y Comercial eliminada correctamente", vbInformation, "Mantenimiento de Convenios"
                limpia_textos
                habilita_textos (False)
            Else
                MsgBox "Convenio y Comercial No se puede eliminar", vbInformation, "Mantenimiento de Convenios"
            End If
        End If
     Else
        
         Call oCredD.ObtenerDatosCasaComercial(loDatos, 999)
         If loDatos.EOF Then
            MsgBox "No cuenta con Ningun registro, Haga Clic en el boton Nuevo", vbInformation, "Mantenimiento de Convenios"
            limpia_textos
            habilita_textos (False)
            Me.Cmdnuevo.SetFocus
         Else
            habilita_textos (True)
            SSTDatosGen.TabVisible(1) = False
            SSTDatosGen.TabVisible(2) = True
            SSTDatosGen.Tab = 2
            MostrarEmpresa
            Me.cmdGrabar.Caption = "Actualizar"
            Me.CmdBuscar.Caption = "Eliminar"
            Me.CmdSalir.Caption = "Cancelar"
            
         End If
    End If

End Sub

Private Sub MostrarEmpresa()
    Dim k As Integer
    Dim i As Integer
    Dim j As Integer
    Dim kg As Integer
    k = 0
    Dim nEncontrado As Integer
        If nPost > 0 Then
            For i = 1 To nPost
                FEEmpresas.EliminaFila (1)
            Next i
        End If
        
        nPost = 0
        j = 0
        k = 0
        kg = 0
        
        If loDatos.EOF Or loDatos.BOF Then
            Exit Sub
        End If
        nEncontrado = 0
        Do Until loDatos.EOF
            If nEncontrado = 0 Then
                k = k + 1
                FEEmpresas.AdicionaFila
                'FEEmpresas.TextMatrix(K, 0) = ""
                FEEmpresas.TextMatrix(k, 1) = loDatos!id
                FEEmpresas.TextMatrix(k, 2) = loDatos!cNombre
                FEEmpresas.TextMatrix(k, 3) = loDatos!cInstitucion
            End If
            nPost = j
            nEncontrado = 0
       loDatos.MoveNext
    Loop

End Sub


Private Sub CmdGrabar_Click()
  Dim oCredD As COMDCredito.DCOMCreditos
  Dim rs As ADODB.Recordset
  
  Set oCredD = New COMDCredito.DCOMCreditos
  Set rs = New ADODB.Recordset
  
        If valida_datos Then
              If nOperacion = lRegistro Then
                    If MsgBox("Esta Seguro de registrar el convenio ?", vbQuestion + vbYesNo, "Mantenimiento de Convenios") = vbYes Then
                      
                      Set rs = oCredD.DevolverDatos_ValidaCasaComercial(Trim(Me.txtnombre.Text), Trim(Me.txtempresa.Text))
                          If rs.EOF And rs.BOF Then
                            oCredD.Insertar_Datos_CasaComercial Trim(Me.txtnombre.Text), Trim(Me.txtempresa.Text), Trim(Me.txtdescripcion.Text), CDate(Me.txtFec.Text), CDate(Me.txtfec1.Text)
                            MsgBox "Comercial registrado correctamente", vbInformation, "Mantenimiento de Convenios"
                            limpia_textos
                            habilita_textos (False)
                            Me.Cmdnuevo.SetFocus
                          Else
                             MsgBox "Datos ya registrados, Verífique", vbInformation, "Mantenimiento de Convenios"
                          End If
                    Else
                      MsgBox "Datos cancelados", vbInformation, "Mantenimiento de Convenios"
                    End If
               ElseIf nOperacion = lConsultas Then
                    'Set rs = oCredD.VerificarCasaComercialFechaVenc(lsCodCasa, CDate(Trim(Me.txtfec1.Text)))
                    '    If Not rs.EOF And rs.BOF Then
                            If MsgBox("Esta Seguro de Actualizar datos ?", vbQuestion + vbYesNo, "Mantenimiento de Convenios") = vbYes Then
                               oCredD.Actualizar_Datos_CasaComercial lsCodCasa, Trim(Me.txtnombre.Text), Trim(Me.txtempresa.Text), Trim(Me.txtdescripcion.Text), CDate(Me.txtFec.Text), CDate(Me.txtfec1.Text)
                               MsgBox "Comercial actualizado correctamente", vbInformation, "Mantenimiento de Convenios"
                            End If
                    '    Else
                    '        MsgBox "Fecha de Vencimiento no Valida", vbInformation, "Mantenimiento de Convenios"
                    '    End If
               End If
        Else
              Exit Sub
        End If
  Set oCredD = Nothing
  Set rs = Nothing
End Sub
Public Function valida_datos() As Boolean
valida_datos = True

If Me.txtnombre.Text = "" Or Me.txtempresa.Text = "" Or Me.txtdescripcion.Text = "" Then
    valida_datos = False
    MsgBox "Complete los Datos para Registrar", vbInformation, "Mantenimiento de Convenios"
End If

If Not IsDate(Me.txtFec) Then
    valida_datos = False
    MsgBox "La fecha de inicio no es correcta", vbInformation, "Mantenimiento de Convenios"
End If

If Not IsDate(Me.txtfec1) Then
    valida_datos = False
    MsgBox "La fecha final no es correcta", vbInformation, "Mantenimiento de Convenios"
End If

End Function

Private Sub Cmdguardar_Click()
 Dim oCredD As COMDCredito.DCOMCreditos
 Set oCredD = New COMDCredito.DCOMCreditos
 Dim i As Integer
 If Not (verificar_lista(True)) Then
         MsgBox "No seleccionó ninguna Opción", vbInformation, "Mantenimiento de Convenios"
         If MsgBox("Esta Seguro de No Asignar a Ninguna Campaña ?", vbQuestion + vbYesNo, "Mantenimiento de Convenios") = vbYes Then
            oCredD.Eliminar_Datos_ConvenioCasaComercial (CInt(Right(Me.cboconvenio.Text, 2)))
         Else
            Exit Sub
         End If
 ElseIf (Me.cboconvenio.ListIndex = 0) Then
        MsgBox "Seleccione una Casa Comercial Vàlida", vbInformation, "Mantenimiento de Convenios"
        Exit Sub
 Else
         For i = 0 To LstCampana.ListCount - 1
            If LstCampana.Selected(i) = True Then
                nContador = nContador + 1
                oCredD.Insertar_Datos_ConvenioCasaComercial CInt(Right(Me.cboconvenio.Text, 2)), CInt(Trim(Right(LstCampana.List(i), 3))), nContador
           End If
        Next i
         If i = LstCampana.ListCount And nContador > 0 Then
            MsgBox "Asignacion(es) registradas correctamente", vbInformation, "Mantenimiento de Convenios"
         End If
End If
nContador = 0
End Sub

Private Sub cmdNuevo_Click()
nOperacion = lRegistro
limpia_textos
habilita_textos (True)
Me.txtnombre.SetFocus
Me.CmdSalir.Caption = "Cancelar"
End Sub

Sub habilita_textos(val As Boolean)
 Dim oCredD As COMDCredito.DCOMCreditos
 Set oCredD = New COMDCredito.DCOMCreditos
    If nOperacion = lRegistro Then
        Me.txtnombre.Enabled = val
        Me.txtempresa.Enabled = val
        Me.cboconvenio.Enabled = Not val
        Me.LstCampana.Enabled = Not val
        Me.Cmdnuevo.Enabled = Not val
        Me.CmdBuscar.Enabled = Not val
        Me.CmdAsignar.Enabled = Not val
        Me.CmdSalir.Enabled = val
        Me.cmdGrabar.Enabled = val
    ElseIf nOperacion = lConsultas Then
        Me.txtnombre.Enabled = Not val
        Me.txtempresa.Enabled = Not val
        Me.cboconvenio.Enabled = val
        Me.LstCampana.Enabled = val
        Me.Cmdnuevo.Enabled = Not val
        Me.CmdAsignar.Enabled = Not val
        Me.cmdGrabar.Enabled = val
        Me.CmdSalir.Enabled = val
        If Not val Then
            'Me.FEEmpresas.Clear
            Me.CmdSalir.Enabled = Not val
            Me.CmdSalir.Caption = "Salir"
            Me.CmdBuscar.Caption = "Buscar"
            Me.cmdGrabar.Caption = "Grabar"
            Me.CmdBuscar.Enabled = Not val
            SSTDatosGen.TabVisible(1) = val
            SSTDatosGen.TabVisible(2) = val
            SSTDatosGen.Tab = 0
        End If
    ElseIf nOperacion = lAsignar Then
        If (oCredD.DevolverDatos_NumeroCasaComercial < 1) Then
            MsgBox "No cuenta con Ningun registro, Haga Clic en el boton Nuevo", vbInformation, "Mantenimiento de Convenios"
            limpia_textos
            SSTDatosGen.TabVisible(0) = val
            SSTDatosGen.TabVisible(1) = Not val
            SSTDatosGen.TabVisible(2) = Not val
            SSTDatosGen.Tab = 0
            Me.Cmdnuevo.SetFocus
            Exit Sub
        Else
            SSTDatosGen.TabVisible(0) = False
            SSTDatosGen.TabVisible(2) = False
            SSTDatosGen.TabVisible(1) = True
            llenar_cbos_convenio (0)
            SSTDatosGen.Tab = 1
        End If
        Me.cboconvenio.Enabled = val
        Me.LstCampana.Enabled = val
        Me.CmdSalir.Enabled = val
        Me.Cmdnuevo.Enabled = Not val
        Me.CmdBuscar.Enabled = Not val
        Me.CmdSalir1.Enabled = val
    ElseIf nOperacion = lCancelar Then
        Me.txtnombre.Enabled = val
        Me.txtempresa.Enabled = val
        Me.cboconvenio.Enabled = val
        Me.LstCampana.Enabled = val
        Me.Cmdnuevo.Enabled = Not val
        Me.CmdAsignar.Enabled = Not val
        Me.cmdGrabar.Enabled = val
        Me.CmdSalir.Enabled = Not val
        Me.CmdBuscar.Enabled = Not val
        Me.Cmdnuevo.SetFocus
    End If
        Me.txtdescripcion.Enabled = val
        Me.txtFec.Enabled = val
        Me.txtfec1.Enabled = val
End Sub

Sub limpia_textos()
    Me.txtnombre.Text = ""
    Me.txtdescripcion.Text = ""
    Me.txtempresa.Text = ""
    Me.txtFec.Text = "__/__/____"
    Me.txtfec1.Text = "__/__/____"
End Sub

Private Sub cmdSalir_Click()
   
    If Me.CmdSalir.Caption = "Cancelar" Then
        limpia_textos
        SSTDatosGen.TabVisible(0) = True
        SSTDatosGen.TabVisible(1) = False
        SSTDatosGen.TabVisible(2) = False
        SSTDatosGen.Tab = 0
        nOperacion = lCancelar
        Me.cmdGrabar.Caption = "Grabar"
        Me.CmdSalir.Caption = "Salir"
        Me.CmdBuscar.Caption = "Buscar"
        habilita_textos (False)
        Me.Cmdnuevo.SetFocus
    Else
        nOperacion = 2
    End If
    
    If nOperacion = lConsultas Or nOperacion = lRegistro Then
        limpia_textos
        habilita_textos (False)
        Me.Cmdnuevo.Enabled = True
        Me.Cmdnuevo.SetFocus
    Else
        If Not nOperacion = lCancelar Then
            Unload Me
        End If
    End If
    
End Sub

Sub obtener_casa(cod As String)
 Dim oCredD As COMDCredito.DCOMCreditos
 Set oCredD = New COMDCredito.DCOMCreditos
 Call oCredD.ObtenerDatosCasaComercial(loDatos1, CInt(cod))
    Me.txtnombre.Text = loDatos1!cNombre
    Me.txtdescripcion.Text = loDatos1!cDescripcion
    Me.txtempresa.Text = loDatos1!cInstitucion
    Me.txtFec.Text = Format(loDatos1!f_inicio, "DD/MM/YYYY")
    Me.txtfec1.Text = Format(loDatos1!f_fin, "DD/MM/YYYY")
 habilita_textos True
 Set oCredD = Nothing
End Sub

Private Sub CmdSalir1_Click()
    limpia_textos
    SSTDatosGen.TabVisible(0) = True
    SSTDatosGen.TabVisible(1) = False
    SSTDatosGen.TabVisible(2) = False
    SSTDatosGen.Tab = 0
    
    nOperacion = lCancelar
    Me.cmdGrabar.Caption = "Grabar"
    Me.CmdSalir.Caption = "Salir"
    Me.CmdBuscar.Caption = "Buscar"
    habilita_textos (False)
    Me.Cmdnuevo.SetFocus
    
'    Me.CmdNuevo.Enabled = True
'    Me.CmdAsignar.Enabled = True
'    Me.CmdBuscar.Enabled = True
'
'    Me.cmdSalir.Enabled = True
End Sub

Private Sub FEEmpresas_OnCellChange(pnRow As Long, pnCol As Long)
   filaFE = pnRow
   FEEmpresas.row = pnRow
        FEEmpresas.Col = pnCol
        lsCodCasa = FEEmpresas.TextMatrix(FEEmpresas.row, 1)
If FEEmpresas.TextMatrix(FEEmpresas.row, 2) <> "" Then
    Call obtener_casa(lsCodCasa)
End If
End Sub

Private Sub FEEmpresas_RowColChange()
    Call FEEmpresas_OnCellChange(FEEmpresas.row, FEEmpresas.Col)
End Sub

Private Sub Form_Load()
    llenar_cbos_convenio (0)
    nOperacion = lRegistro
    'Set rs = Nothing
    'Set oCons = Nothing
    CargarCampanha
    SSTDatosGen.Tab = 0
    SSTDatosGen.TabVisible(1) = False
    SSTDatosGen.TabVisible(2) = False
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub



