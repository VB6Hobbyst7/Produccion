VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersonaActualizacionDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización y Autorización de Datos"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8970
   Icon            =   "frmPersonaActualizacionDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "Vista Previa"
      Height          =   375
      Left            =   7680
      TabIndex        =   39
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1320
      TabIndex        =   38
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   6960
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos del Cliente"
      TabPicture(0)   =   "frmPersonaActualizacionDatos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblApellidoCasada"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label17"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label18"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtNombre"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtApellidoPaterno"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtApellidoMaterno"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtApellidoCasada"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbPersNatSexo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbPersNatEstCiv"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbNacionalidad"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmbTipoDoi"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtDoi"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkModificarDatosSensibles"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtDireccion"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtReferencia"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtCelular"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtTelefono"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtCorreo"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Frame1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmbPersUbiGeo(0)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmbPersUbiGeo(1)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmbPersUbiGeo(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmbPersUbiGeo(3)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmbPersUbiGeo(4)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   4
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   3
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   2
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3960
         Width           =   1935
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   1
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   3600
         Width           =   1935
      End
      Begin VB.ComboBox cmbPersUbiGeo 
         Height          =   315
         Index           =   0
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Autorizar Uso de Datos"
         Height          =   735
         Left            =   1560
         TabIndex        =   34
         Top             =   5760
         Width           =   2175
         Begin VB.OptionButton optNo 
            Caption         =   "NO"
            Height          =   375
            Left            =   1200
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optSi 
            Caption         =   "SI"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox txtCorreo 
         Height          =   285
         Left            =   1560
         TabIndex        =   32
         Top             =   5400
         Width           =   4935
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   31
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox txtCelular 
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   22
         Top             =   5040
         Width           =   1935
      End
      Begin VB.TextBox txtReferencia 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   3240
         Width           =   5055
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   2880
         Width           =   5055
      End
      Begin VB.CheckBox chkModificarDatosSensibles 
         Caption         =   "Modificar Datos Sensibles"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtDoi 
         Height          =   285
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   18
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox cmbTipoDoi 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cmbNacionalidad 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cmbPersNatEstCiv 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbPersNatSexo 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtApellidoCasada 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtApellidoMaterno 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtApellidoPaterno 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label18 
         Caption         =   "Correo:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Celular:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Zona:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Distrito:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Departamento:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Referencia:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   2100
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo DOI:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Nacionalidad:"
         Height          =   255
         Left            =   5520
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Estado Civil:"
         Height          =   255
         Left            =   5520
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Sexo:"
         Height          =   255
         Left            =   5520
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblApellidoCasada 
         Caption         =   "Apellido Casada:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido Materno:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Apellido Paterno:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombres:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPersonaActualizacionDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TiposBusquedaNombre
    BusqApellidoPaterno = 1
    BusqApellidoMaterno = 2
    BusqApellidoCasada = 3
    BusqNombres = 4
End Enum
Private Enum TTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum
Dim bEstadoCargando As Boolean
Dim nPos As Integer
Dim oPersona As New UPersona_Cli
Dim MatPersona(1 To 2) As TActAutDatos
Dim fsPersCod As String
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim nMovVistoElec As Long
Dim fsNombreCompleto As String
'
Dim Nivel1() As String
Dim ContNiv1 As Long
Dim Nivel2() As String
Dim ContNiv2 As Long
Dim Nivel3() As String
Dim ContNiv3 As Long
Dim Nivel4() As String
Dim ContNiv4 As Long
Dim Nivel5() As String
Dim ContNiv5 As Long
Dim cOpe As String ' add pti1 ERS070-2018 06/12/2018
Dim cNue As String
Dim cAut As String
Dim sCiudadAgencia As String
Dim nTamanio As Integer
Dim Spac As Integer
Dim Index As Integer
Dim Princ As Integer
Dim CantCarac As Integer
Dim txtcDescrip As String
Dim contador As Integer
Dim nCentrar As Integer
Dim nTamLet As Integer
Dim spacvar As Integer ' fin add pti1 ERS070-2018 06/12/2018
 




Public Sub Inicio(ByVal psPersCod As String, _
Optional ByVal psCodOpe As String = "0", _
Optional ByVal bAut As String = "", _
Optional ByVal bSoloActDat As String = "")
'Public Sub Inicio(ByVal psPersCod As String) 'comentado por pti1 ERS070-2018 06/12/2018
    'add pti1 ers070-2018 **************************************
    'Call HabilitarDeshabilitar(False)'comentado por pti1 ERS070-2018
    'bAut 1: es antiguo 'bSoloActDat 0: no autorizó datos 1: ya autorizo datos
    'bAut 0: es nuevo
    cOpe = psCodOpe 'add pti1 ers070-2018
    cAut = bAut 'add pti1 ers070-2018
    cNue = bSoloActDat 'add pti1 ers070-2018
    Call HabilitarDeshabilitar(True)
    Call CargarDatosObjetoPersona
    Call ObtenerDatos(psPersCod)
    Call ObtenerDistritoAgencia(gsCodAge)
    If bAut = 1 Then 'add pti1 ers070-2018 22/12/2018 si es  bAut=1 ya se registro con anterioridad
    'ya se registro con anterioridad
            If bSoloActDat = 0 Then
            'no acepto el uso de sus datos
                Frame1.Visible = True
                optNo = True
                fsPersCod = psPersCod
                Me.Show 1
            Else
                Frame1.Visible = False
                fsPersCod = psPersCod
                Me.Show 1
            End If
    Else
           If bAut = 0 Then
           'bAut=0 (la persona aun no registro la autorización de  sus datos)
                Frame1.Visible = True
                fsPersCod = psPersCod
                Me.Show 1
            Else
            'esta opción no debería usarse y si es que se da la opcion solo se actualizara los datos mas no la autorización
                fsPersCod = psPersCod
                Frame1.Visible = False
                Me.Show 1
            End If
    End If
    'fin add pti1 ers070-2018 ****************************************
    
'    Call HabilitarDeshabilitar(False) 'comentado por pti1
'    Call CargarDatosObjetoPersona
'    Call ObtenerDatos(psPersCod)
'    fsPersCod = psPersCod
'    Me.Show 1
End Sub
Private Sub HabilitarDeshabilitar(ByVal pbValor As Boolean)
    txtNombre.Enabled = pbValor
    txtApellidoPaterno.Enabled = pbValor
    txtApellidoMaterno.Enabled = pbValor
    If txtApellidoCasada.Visible = True Then
        txtApellidoCasada.Enabled = pbValor
    End If
    cmbTipoDoi.Enabled = pbValor
    txtDoi.Enabled = pbValor
    cmbPersNatSexo.Enabled = pbValor
    cmbPersNatEstCiv.Enabled = pbValor
    cmbNacionalidad.Enabled = pbValor
End Sub

Private Sub CargarDatosObjetoPersona()
    Dim oDPersonaS As New comdpersona.DCOMPersonas
    Dim oConstante As New COMDConstantes.DCOMConstantes
    Dim lrsEstCivil As New ADODB.Recordset
    Dim lrsUbiGeo As New ADODB.Recordset
    Dim lrsTipoDoi As New ADODB.Recordset
    Dim i As Integer
    
    bEstadoCargando = True
    'Carga Tipos de Sexo de Personas
    cmbPersNatSexo.AddItem "FEMENINO" & Space(50) & "F"
    cmbPersNatSexo.AddItem "MASCULINO" & Space(50) & "M"
    'Carga Estado Civil
    Set lrsEstCivil = oConstante.RecuperaConstantes(gPersEstadoCivil)
    Call Llenar_Combo_con_Recordset(lrsEstCivil, cmbPersNatEstCiv)
    'Ubicaciones Geograficas
    Set lrsUbiGeo = oDPersonaS.CargarUbicacionesGeograficas(True, 0)
    While Not lrsUbiGeo.EOF
        cmbPersUbiGeo(0).AddItem Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
        cmbNacionalidad.AddItem Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
        lrsUbiGeo.MoveNext
    Wend
    If lrsUbiGeo.RecordCount > 0 Then lrsUbiGeo.MoveFirst
    For i = 0 To lrsUbiGeo.RecordCount
        If Trim(lrsUbiGeo!cUbiGeoCod) = "04028" Then
            nPos = i
        End If
    Next i
    cmbPersUbiGeo(0).ListIndex = nPos
    cmbNacionalidad.ListIndex = nPos
    'Tipo Doi
    Set lrsTipoDoi = oConstante.RecuperaConstantes(gPersIdTipo)
    Call Llenar_Combo_con_Recordset(lrsTipoDoi, cmbTipoDoi)
End Sub
Private Sub ObtenerDatos(ByVal psPersCod As String)
    Dim oNPersona As New COMNPersona.NCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim sUbicGeografica As String
    Dim sNombres As String, sApePat As String, sApeMat As String, sApeCas As String
    Dim sSexo As String, sEstadoCivil As String
    Dim sDomicilio As String, cNacionalidad As String, sRefDomicilio As String
    Dim sTelefonos As String, sCelular As String, sEmail As String
    Dim sPersIDTpo As String, sPersIDnro As String
    
    bEstadoCargando = True
    Set rsPersona = oNPersona.ObtenerDatosParaActAutDeCliente(psPersCod)
    If Not (rsPersona.EOF And rsPersona.BOF) Then
        sApePat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoPaterno)
        sApeMat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoMaterno)
        sNombres = BuscaNombre(rsPersona!cPersNombre, BusqNombres)
        sApeCas = BuscaNombre(rsPersona!cPersNombre, BusqApellidoCasada)
        sSexo = Trim(IIf(IsNull(rsPersona!cPersnatSexo), "", rsPersona!cPersnatSexo))
        sEstadoCivil = Trim(IIf(IsNull(rsPersona!nPersNatEstCiv), "", rsPersona!nPersNatEstCiv))
        sDomicilio = rsPersona!cPersDireccDomicilio
        sUbicGeografica = rsPersona!cPersDireccUbiGeo
        sTelefonos = IIf(IsNull(rsPersona!cPersTelefono), "", rsPersona!cPersTelefono)
        sCelular = IIf(IsNull(rsPersona!cPersCelular), "", rsPersona!cPersCelular)
        sEmail = IIf(IsNull(rsPersona!cEmail), "", rsPersona!cEmail)
        cNacionalidad = Trim(IIf(IsNull(rsPersona!cNacionalidad), "", rsPersona!cNacionalidad))
        sRefDomicilio = Trim(IIf(IsNull(rsPersona!cPersRefDomicilio), "", rsPersona!cPersRefDomicilio))
        sPersIDTpo = Trim(IIf(IsNull(rsPersona!cPersIDTpo), "", rsPersona!cPersIDTpo))
        sPersIDnro = Trim(IIf(IsNull(rsPersona!cPersIDnro), "", rsPersona!cPersIDnro))
        
        txtNombre.Text = sNombres
        txtApellidoPaterno.Text = sApePat
        txtApellidoMaterno.Text = sApeMat
        txtDireccion.Text = sDomicilio
        txtReferencia.Text = sRefDomicilio
        oPersona.Sexo = sSexo
        If sSexo = "F" Then
            txtApellidoCasada.Text = sApeCas
            cmbPersNatSexo.ListIndex = 0
            If sApeCas = "" Then
                Call DistribuyeApellidos(False)
            Else
                Call DistribuyeApellidos(True)
            End If
        Else
            cmbPersNatSexo.ListIndex = 1
            Call DistribuyeApellidos(False)
        End If
        'Carga Ubicacion Georgrafica
        If Len(Trim(sUbicGeografica)) = 12 Then
            cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "04028")
            cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "1" & Mid(sUbicGeografica, 2, 2) & String(9, "0"))
            cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "2" & Mid(sUbicGeografica, 2, 4) & String(7, "0"))
            cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & "3" & Mid(sUbicGeografica, 2, 6) & String(5, "0"))
            cmbPersUbiGeo(4).ListIndex = IndiceListaCombo(cmbPersUbiGeo(4), Space(30) & sUbicGeografica)
        Else
            cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & sUbicGeografica)
            cmbPersUbiGeo(1).Clear
            cmbPersUbiGeo(1).AddItem cmbPersUbiGeo(0).Text
            cmbPersUbiGeo(1).ListIndex = 0
            cmbPersUbiGeo(2).Clear
            cmbPersUbiGeo(2).AddItem cmbPersUbiGeo(0).Text
            cmbPersUbiGeo(2).ListIndex = 0
            cmbPersUbiGeo(3).Clear
            cmbPersUbiGeo(3).AddItem cmbPersUbiGeo(0).Text
            cmbPersUbiGeo(3).ListIndex = 0
            cmbPersUbiGeo(4).Clear
            cmbPersUbiGeo(4).AddItem cmbPersUbiGeo(0).Text
            cmbPersUbiGeo(4).ListIndex = 0
        End If
        txtDoi.Text = sPersIDnro
        cmbTipoDoi.ListIndex = IndiceListaCombo(cmbTipoDoi, sPersIDTpo)
        'FRHU 20151205 OBSERVACION
        If Trim(Right(cmbTipoDoi.Text, 2)) = "1" Then
            txtDoi.MaxLength = 8
        Else
            txtDoi.MaxLength = 15
        End If
        'FIN FRHU
        cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, sEstadoCivil)
        cmbNacionalidad.ListIndex = IndiceListaCombo(cmbNacionalidad, Space(30) & cNacionalidad)
        txtCelular.Text = sCelular
        txtTelefono.Text = sTelefonos
        txtCorreo.Text = sEmail
        oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
        MatPersona(1).sNombres = sNombres
        MatPersona(1).sApePat = sApePat
        MatPersona(1).sApeMat = sApeMat
        MatPersona(1).sApeCas = sApeCas
        MatPersona(1).sPersIDTpo = sPersIDTpo
        MatPersona(1).sPersIDnro = sPersIDnro
        MatPersona(1).sSexo = sSexo
        MatPersona(1).sEstadoCivil = sEstadoCivil
        MatPersona(1).cNacionalidad = cNacionalidad
        MatPersona(1).sDomicilio = sDomicilio
        MatPersona(1).sRefDomicilio = sRefDomicilio
        MatPersona(1).sUbicGeografica = sUbicGeografica
        MatPersona(1).sCelular = sCelular
        MatPersona(1).sTelefonos = sTelefonos
        MatPersona(1).sEmail = sEmail
  
    End If
    bEstadoCargando = False
End Sub
'add pti1 ers070-2018 28/12/2018
Private Sub ObtenerDistritoAgencia(ByVal sCodAge As String)
    Dim oNPersona As New COMNPersona.NCOMPersona
    
    sCiudadAgencia = oNPersona.ObtenerDistritoAgencia(sCodAge)
    
End Sub
Private Sub DistribuyeApellidos(ByVal bApellCasada As Boolean)
   If bApellCasada = True Then
        'lblApCasada.Visible = True
        'txtApellidoCasada.Visible = True
               
        'lblPersNombreAM.Top = 840
        'txtApellidoMaterno.Top = 840
        
        'lblApCasada.Top = 1200
        'txtApellidoCasada.Top = 1200
        
        'lblPersNombreN.Top = 1560
        'txtNombre.Top = 1560
        lblApellidoCasada.Visible = True
        txtApellidoCasada.Visible = True
    Else
        'lblPersNombreAM.Top = 960
        'txtApellidoMaterno.Top = 960
                
        'lblApCasada.Visible = False
        'txtApellidoCasada.Visible = False
        
        'lblPersNombreN.Top = 1440
        'txtNombre.Top = 1440
        lblApellidoCasada.Visible = False
        txtApellidoCasada.Visible = False
    End If
End Sub
Private Function BuscaNombre(ByVal psNombre As String, ByVal nTipoBusqueda As TiposBusquedaNombre) As String
Dim sCadTmp As String
Dim PosIni As Integer
Dim PosFin As Integer
Dim PosIni2 As Integer
    sCadTmp = ""
    Select Case nTipoBusqueda
        Case 1 'Busqueda de Apellido Paterno
            If Mid(psNombre, 1, 1) <> "/" And Mid(psNombre, 1, 1) <> "\" And Mid(psNombre, 1, 1) <> "," Then
                PosIni = 1
                PosFin = InStr(1, psNombre, "/")
                If PosFin = 0 Then
                    PosFin = InStr(1, psNombre, "\")
                    If PosFin = 0 Then
                        PosFin = InStr(1, psNombre, ",")
                        If PosFin = 0 Then
                            PosFin = Len(psNombre)
                        End If
                    End If
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 2 'Apellido materno
           PosIni = InStr(1, psNombre, "/")
           If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = InStr(1, psNombre, "\")
                If PosFin = 0 Then
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 3 'Apellido de casada
           PosIni = InStr(1, psNombre, "\")
           If PosIni <> 0 Then
                PosIni2 = InStr(1, psNombre, "VDA")
                If PosIni2 <> 0 Then
                    PosIni = PosIni2 + 3
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                Else
                    PosIni = PosIni + 1
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Trim(Mid(psNombre, PosIni, PosFin - PosIni))
            Else
                sCadTmp = ""
            End If
        Case 4 'Nombres
            PosIni = InStr(1, psNombre, ",")
            If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = Len(psNombre)
                sCadTmp = Mid(psNombre, PosIni, (PosFin + 1) - PosIni)
            Else
                sCadTmp = ""
            End If
            
    End Select
    BuscaNombre = sCadTmp
End Function
'COMENTADO POR PTI1 ERS070-1018 05-12-2018
'Private Sub chkModificarDatosSensibles_Click()
'    If chkModificarDatosSensibles.value = 1 Then
'        Call HabilitarDeshabilitar(True)
'    End If
'    If chkModificarDatosSensibles.value = 0 Then
'        Call HabilitarDeshabilitar(False)
'        txtNombre.Text = MatPersona(1).sNombres
'        txtApellidoPaterno.Text = MatPersona(1).sApePat
'        txtApellidoMaterno.Text = MatPersona(1).sApeMat
'        cmbTipoDoi.ListIndex = IndiceListaCombo(cmbTipoDoi, MatPersona(1).sPersIDTpo)
'        txtDoi.Text = MatPersona(1).sPersIDnro
'        If MatPersona(1).sSexo = "F" Then
'            txtApellidoCasada.Text = MatPersona(1).sApeCas
'            cmbPersNatSexo.ListIndex = 0
'            If MatPersona(1).sApeCas = "" Then
'                Call DistribuyeApellidos(False)
'            Else
'                Call DistribuyeApellidos(True)
'            End If
'        Else
'            cmbPersNatSexo.ListIndex = 1
'            Call DistribuyeApellidos(False)
'        End If
'        cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, MatPersona(1).sEstadoCivil)
'        cmbNacionalidad.ListIndex = IndiceListaCombo(cmbNacionalidad, Space(30) & MatPersona(1).cNacionalidad)
'    End If
'End Sub

Private Sub cmbPersNatEstCiv_Change()
    If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
  End If
End Sub

Private Sub cmbPersNatEstCiv_Click()
    If Not bEstadoCargando Then
'        If oPersona.TipoActualizacion <> PersFilaNueva Then
'            oPersona.TipoActualizacion = PersFilaModificada
'        End If
        oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
        If (CInt(oPersona.EstadoCivil) = gPersEstadoCivilCasado Or CInt(oPersona.EstadoCivil) = gPersEstadoCivilViudo) And oPersona.Sexo = "F" Then
            Call DistribuyeApellidos(True)
        Else
            Call DistribuyeApellidos(False)
            txtApellidoCasada.Text = ""
        End If
    End If
End Sub

Private Sub cmbPersNatSexo_Change()
    If Not bEstadoCargando Then
'    If oPersona.TipoActualizacion <> PersFilaNueva Then
'        oPersona.TipoActualizacion = PersFilaModificada
'    End If
    oPersona.Sexo = Trim(Right(cmbPersNatSexo.Text, 10))
  End If
End Sub

Private Sub cmbPersNatSexo_Click()
    If Not bEstadoCargando Then
'        If oPersona.TipoActualizacion <> PersFilaNueva Then
'            oPersona.TipoActualizacion = PersFilaModificada
'        End If
        oPersona.Sexo = Trim(Right(cmbPersNatSexo.Text, 10))
        If oPersona.EstadoCivil = "" Then
            Exit Sub
        End If
        If (CInt(oPersona.EstadoCivil) = gPersEstadoCivilCasado Or CInt(oPersona.EstadoCivil) = gPersEstadoCivilViudo) And oPersona.Sexo = "F" Then
            Call DistribuyeApellidos(True)
        Else
            Call DistribuyeApellidos(False)
            txtApellidoCasada.Text = ""
        End If
    End If
End Sub

Private Sub cmbPersUbiGeo_Change(Index As Integer)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
    End If
End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
    Dim oUbic As comdpersona.DCOMPersonas
    Dim rs As ADODB.Recordset
    Dim i As Integer

    If Index <> 4 Then

        Set oUbic = New comdpersona.DCOMPersonas
        Set rs = oUbic.CargarUbicacionesGeograficas(True, Index + 1, Trim(Right(cmbPersUbiGeo(Index).Text, 15)))

    If Trim(Right(cmbPersUbiGeo(0).Text, 12)) <> "04028" Then
        If Index = 0 Then
            For i = 1 To cmbPersUbiGeo.count - 1
                cmbPersUbiGeo(i).Clear
                cmbPersUbiGeo(i).AddItem Trim(Trim(cmbPersUbiGeo(0).Text)) & Space(50) & Trim(Right(cmbPersUbiGeo(0).Text, 12))
            Next i
         End If
    Else
        For i = Index + 1 To cmbPersUbiGeo.count - 1
        cmbPersUbiGeo(i).Clear
        Next
        
        While Not rs.EOF
            cmbPersUbiGeo(Index + 1).AddItem Trim(rs!cUbiGeoDescripcion) & Space(50) & Trim(rs!cUbiGeoCod)
            rs.MoveNext
        Wend
    End If
    Set oUbic = Nothing
    End If

'    If Not bEstadoCargando Then
'        If oPersona.TipoActualizacion <> PersFilaNueva Then
'            oPersona.TipoActualizacion = PersFilaModificada
'        End If
'        oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
'    End If
End Sub

Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 4 Then
            cmbPersUbiGeo(Index + 1).SetFocus
        Else
            txtCelular.SetFocus
        End If
    End If
End Sub
Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TTipoCombo)
Dim i As Long
Dim sCodigo As String

    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboDpto
            cmbPersUbiGeo(1).Clear
            If sCodigo = "04028" Then
                For i = 0 To ContNiv2 - 1
                    cmbPersUbiGeo(1).AddItem Nivel2(i)
                Next i
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboProv
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(2).Clear
                For i = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel3(i), 15)), 2, 2) Then
                        cmbPersUbiGeo(2).AddItem Nivel3(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel4(i), 15)), 2, 4) Then
                        cmbPersUbiGeo(3).AddItem Nivel4(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(4).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv5 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel5(i), 15)), 2, 6) Then
                        cmbPersUbiGeo(4).AddItem Nivel5(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(4).AddItem psValor
            End If
    End Select
End Sub
Public Function VistoElectronico() As Boolean
    Dim loVistoElectronico As New frmVistoElectronico
    Dim lbVistoVal As Boolean
    
    lbVistoVal = loVistoElectronico.Inicio(14, nOperacion)
                       
    If Not lbVistoVal Then
        MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones.", vbInformation, "Mensaje del Sistema"
        VistoElectronico = False
        Exit Function
    End If
    VistoElectronico = True
    Call loVistoElectronico.RegistraVistoElectronico(0, nMovVistoElec)
End Function
'FRHU 20151205 OBSERVACION
Private Sub cmbTipoDoi_Click()
    If Trim(Right(cmbTipoDoi.Text, 2)) = "1" Then
        txtDoi.MaxLength = 8
    Else
        txtDoi.MaxLength = 15
    End If
End Sub
Private Sub cmbTipoDoi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Right(cmbTipoDoi.Text, 2)) = "1" Then
            txtDoi.MaxLength = 8
        Else
            txtDoi.MaxLength = 15
        End If
    End If
End Sub
'FIN FRHU
Private Sub cmdAceptar_Click()
    Dim chkModDs As Integer
    Dim sMovNro As String
    Dim oCont As New COMNContabilidad.NCOMContFunciones
    Dim oNPersona As New COMNPersona.NCOMPersona
    Dim nSalir As Integer
    Dim nAutorizaUsoDatos As Integer
    
    If Not ValidarDatos Then Exit Sub 'FRHU 20151205 Observacion 'descomentado por pti1 30/01/2019 LPDP
    'Screen.MousePointer = 11
    nSalir = ValidarNombresApellidos(txtApellidoPaterno.Text, Len(txtApellidoPaterno.Text), nSalir, 1)
    nSalir = ValidarNombresApellidos(txtApellidoMaterno.Text, Len(txtApellidoMaterno.Text), nSalir, 2)
    nSalir = ValidarNombresApellidos(txtNombre.Text, Len(txtNombre.Text), nSalir, 3)
    If nSalir <> 0 Then
        MsgBox ("Nombre tiene caracteres no aceptados")
        If nSalir = 1 Then
            txtApellidoPaterno.SetFocus
        ElseIf nSalir = 2 Then
            txtApellidoMaterno.SetFocus
        ElseIf nSalir = 3 Then
            txtNombre.SetFocus
        End If
        Screen.MousePointer = 0
        Exit Sub
    End If
    
'     If optSi.value = 0 And optNo.value = 0 Then 'comentado por PTI1 ERS070-2018
'        MsgBox ("Debe seleccionar si Autoriza o no, el Uso de Datos")
'        Exit Sub
'    End If 'FIN COMENTADO POR PTI1
    
    If cAut = 0 Then 'ADD PTI1 ERS070-2018
         If optSi.value = 0 And optNo.value = 0 And Frame1.Visible Then
            MsgBox ("Debe seleccionar si autoriza o no, el uso de sus datos"), vbInformation, "AVISO"
            Exit Sub
        End If
    End If 'FIN PTI1 ERS070-2018
    

    'FRHU 20151217 INCIDENTE 'TIC1512170001'
    'If Trim(Right(cmbPersNatSexo.Text, 2)) <> "F" And Len(Trim(txtApellidoCasada.Text)) > 0 Then
    '    fsNombreCompleto = txtApellidoPaterno.Text & "/" & txtApellidoMaterno.Text & "\" & txtApellidoCasada.Text & "," & txtNombre.Text
    'Else
    '    fsNombreCompleto = txtApellidoPaterno.Text & "/" & txtApellidoMaterno.Text & "," & txtNombre.Text
    'End If
    If Trim(Right(cmbPersNatSexo.Text, 2)) = "F" And Len(Trim(txtApellidoCasada.Text)) > 0 Then
        If Trim(Right(cmbPersNatEstCiv.Text, 2)) = "3" Then
            fsNombreCompleto = txtApellidoPaterno.Text & "/" & txtApellidoMaterno.Text & "\VDA " & txtApellidoCasada.Text & "," & txtNombre.Text
        Else
            fsNombreCompleto = txtApellidoPaterno.Text & "/" & txtApellidoMaterno.Text & "\" & txtApellidoCasada.Text & "," & txtNombre.Text
        End If
    Else
        fsNombreCompleto = txtApellidoPaterno.Text & "/" & txtApellidoMaterno.Text & "," & txtNombre.Text
    End If
    'FIN FRHU
    
    MatPersona(2).sNombres = txtNombre.Text
    MatPersona(2).sApePat = txtApellidoPaterno.Text
    MatPersona(2).sApeMat = txtApellidoMaterno.Text
    MatPersona(2).sApeCas = txtApellidoCasada.Text
    MatPersona(2).sPersIDTpo = Trim(Right(cmbTipoDoi.Text, 2))
    MatPersona(2).sPersIDnro = txtDoi.Text
    MatPersona(2).sSexo = Trim(Right(cmbPersNatSexo.Text, 2))
    MatPersona(2).sEstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 2))
    'MatPersona(2).cNacionalidad = Trim(Right(cmbNacionalidad.Text, 2))
    MatPersona(2).cNacionalidad = Trim(Right(cmbNacionalidad.Text, 12)) 'FRHU 20151206 OBSERVACION
    MatPersona(2).sDomicilio = txtDireccion.Text
    MatPersona(2).sRefDomicilio = txtReferencia.Text
    MatPersona(2).sUbicGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 12))
    MatPersona(2).sCelular = txtCelular.Text
    MatPersona(2).sTelefonos = txtTelefono.Text
    MatPersona(2).sEmail = txtCorreo.Text
    
    If Not ValidarDatos Then Exit Sub 'FRHU 20151205 Observacion
    
    'If chkModificarDatosSensibles.value = 1 Then ' INICIO comentado por pti1 ERS070-2018 05122018
        'If Not VistoElectronico Then Exit Sub
        'chkModDs = 1
    'Else
        'chkModDs = 0
    'End If 'FIN por pti1 ERS070-2018 05122018
    
    'ADD PTI1 ERS070-2018 ***********************************************
    
    
       If cAut = 1 Then
            If cNue = 1 Then
                nAutorizaUsoDatos = cNue
                'YA SE REGISTRÓ CON ANTERIORIDAD Y YA PASO 6 MESES  Y EL CLIENTE AUTORIZÓ SUS DATOS POR LO TAL SOLO SE ACTUALIZARÁ
                sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
                'Call oNPersona.InsertarPersActAutDatos(sMovNro, chkModDs, fsPersCod, MatPersona(), nAutorizaUsoDatos, 1, 1, nMovVistoElec)
                Call oNPersona.ActualizarPersActAutDatos(fsPersCod, fsNombreCompleto, MatPersona(2).sSexo, MatPersona(2).sEstadoCivil, MatPersona(2).cNacionalidad, MatPersona(2).sPersIDTpo, _
                                     MatPersona(2).sPersIDnro, MatPersona(2).sDomicilio, MatPersona(2).sRefDomicilio, MatPersona(2).sUbicGeografica, MatPersona(2).sCelular, MatPersona(2).sTelefonos, _
                                     MatPersona(2).sEmail, sMovNro)
            
                Call MsgBox("Los datos se actualizaron correctamente", vbInformation, "AVISO")
                Call ImprimirPdfCartilla
                Unload Me
            Else
                'YA SE REGISTRÓ CON ANTERIORIDAD Y YA PASO 6 MESES  EL SISTEMA ACTULIZARA SUS DATOS Y HARA SU NUEVA AUTORIZACIÓN DE DATOS
                If optSi.value = True And optNo.value = False Then nAutorizaUsoDatos = 1
                If optSi.value = False And optNo.value = True Then nAutorizaUsoDatos = 0
            
                sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Call oNPersona.InsertarPersActAutDatos(sMovNro, chkModDs, fsPersCod, MatPersona(), nAutorizaUsoDatos, 1, 1, nMovVistoElec, 1, cOpe)
                Call oNPersona.ActualizarPersActAutDatos(fsPersCod, fsNombreCompleto, MatPersona(2).sSexo, MatPersona(2).sEstadoCivil, MatPersona(2).cNacionalidad, MatPersona(2).sPersIDTpo, _
                                     MatPersona(2).sPersIDnro, MatPersona(2).sDomicilio, MatPersona(2).sRefDomicilio, MatPersona(2).sUbicGeografica, MatPersona(2).sCelular, MatPersona(2).sTelefonos, _
                                     MatPersona(2).sEmail, sMovNro)
            
                Call MsgBox("Los datos se actualizaron correctamente", vbInformation, "AVISO")
                Call ImprimirPdfCartillaAutorizacion
                Unload Me
               
            End If
    Else
           If cAut = 0 Then
           'NUEVO REGISTRO DE AUTORIZACIÓN
                If optSi.value = True And optNo.value = False Then nAutorizaUsoDatos = 1
                If optSi.value = False And optNo.value = True Then nAutorizaUsoDatos = 0

                sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Call oNPersona.InsertarPersActAutDatos(sMovNro, chkModDs, fsPersCod, MatPersona(), nAutorizaUsoDatos, 1, 1, nMovVistoElec, 1, cOpe)
                Call oNPersona.ActualizarPersActAutDatos(fsPersCod, fsNombreCompleto, MatPersona(2).sSexo, MatPersona(2).sEstadoCivil, MatPersona(2).cNacionalidad, MatPersona(2).sPersIDTpo, _
                                     MatPersona(2).sPersIDnro, MatPersona(2).sDomicilio, MatPersona(2).sRefDomicilio, MatPersona(2).sUbicGeografica, MatPersona(2).sCelular, MatPersona(2).sTelefonos, _
                                     MatPersona(2).sEmail, sMovNro)

                Call MsgBox("Los datos se actualizaron correctamente", vbInformation, "AVISO")
                Call ImprimirPdfCartillaAutorizacion
                Call ImprimirPdfCartilla
                Unload Me
            Else
                'If optSi.value = True And optNo.value = False Then nAutorizaUsoDatos = 1
                'If optSi.value = False And optNo.value = True Then nAutorizaUsoDatos = 0

                sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                'Call oNPersona.InsertarPersActAutDatos(sMovNro, chkModDs, fsPersCod, MatPersona(), nAutorizaUsoDatos, 1, 1, nMovVistoElec)
                Call oNPersona.ActualizarPersActAutDatos(fsPersCod, fsNombreCompleto, MatPersona(2).sSexo, MatPersona(2).sEstadoCivil, MatPersona(2).cNacionalidad, MatPersona(2).sPersIDTpo, _
                                     MatPersona(2).sPersIDnro, MatPersona(2).sDomicilio, MatPersona(2).sRefDomicilio, MatPersona(2).sUbicGeografica, MatPersona(2).sCelular, MatPersona(2).sTelefonos, _
                                     MatPersona(2).sEmail, sMovNro)

                Call MsgBox("Los datos se actualizaron correctamente", vbInformation, "AVISO")
                Call ImprimirPdfCartilla
                Unload Me
            End If
    End If
    
 
'ADD PTI1 ERS070-2108 *****************************************
'    If optSi.value = True And optNo.value = False Then nAutorizaUsoDatos = 1 'inicio comentado por pti1 ERS070-2018 06/12/2018
'    If optSi.value = False And optNo.value = True Then nAutorizaUsoDatos = 0
'
'    sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'    Call oNPersona.InsertarPersActAutDatos(sMovNro, chkModDs, fsPersCod, MatPersona(), nAutorizaUsoDatos, 1, 1, nMovVistoElec)
'    Call oNPersona.ActualizarPersActAutDatos(fsPersCod, fsNombreCompleto, MatPersona(2).sSexo, MatPersona(2).sEstadoCivil, MatPersona(2).cNacionalidad, MatPersona(2).sPersIDTpo, _
'                         MatPersona(2).sPersIDnro, MatPersona(2).sDomicilio, MatPersona(2).sRefDomicilio, MatPersona(2).sUbicGeografica, MatPersona(2).sCelular, MatPersona(2).sTelefonos, _
'                         MatPersona(2).sEmail, sMovNro)
'
'    Call MsgBox("Los datos se actualizaron correctamente", vbInformation, "AVISO")
'    Call ImprimirPdfCartilla
'    Unload Me 'FIN COMENTADO PTI1
End Sub
'FRHU 20151204 ERS077-2015
Private Function ValidarDatos() As Boolean
    ValidarDatos = True
   ' If chkModificarDatosSensibles.value = 1 Then 'COMENTADO POR PTI1 ERS070-2018
        If txtNombre.Text = "" Then
            MsgBox "Falta ingresar el Nombre", vbInformation, "AVISO"
            ValidarDatos = False
            txtNombre.SetFocus
            Exit Function
        End If
        If txtApellidoPaterno.Text = "" Then
            MsgBox "Falta ingresar el Apellido Paterno", vbInformation, "AVISO"
            ValidarDatos = False
            txtApellidoPaterno.SetFocus
            Exit Function
        End If
        If txtDoi.Text = "" Then
            MsgBox "Falta ingresar el DOI", vbInformation, "AVISO"
            ValidarDatos = False
            txtDoi.SetFocus
            Exit Function
        End If
        If txtDireccion.Text = "" Then
            MsgBox "Falta ingresar la dirección", vbInformation, "AVISO"
            ValidarDatos = False
            txtDireccion.SetFocus
            Exit Function
        End If
        'If MatPersona(1).cNacionalidad <> MatPersona(2).cNacionalidad Then 'OBSERVACION 'COMENTADO POR PTI1 ERS070-2018
        If MatPersona(1).cNacionalidad <> Trim(Right(cmbNacionalidad.Text, 18)) Then  'ADD PTI1 ERS070-2018
            MsgBox "Se intenta cambiar la nacionalidad, favor coordinar con el area respectiva.", vbInformation, "AVISO"
            ValidarDatos = False
            cmbNacionalidad.SetFocus
            Exit Function
        End If
        If Trim(Right(cmbTipoDoi.Text, 2)) = "1" Then
            If Len(txtDoi.Text) <> 8 Then
                MsgBox "El Dni debe tener 8 caracteres", vbInformation, "AVISO"
                ValidarDatos = False
                txtDoi.SetFocus
                Exit Function
            End If
        End If
     
   ' End If COMENTADO POR PTI1 ERS070-2018
       'ADD PTI1 ERS070-2018 22/12/2018**************************
        If txtCelular.Text = "" Then
            MsgBox "Falta ingresar el número de celular", vbInformation, "AVISO"
            ValidarDatos = False
            txtCelular.SetFocus
            Exit Function
        End If
        If Len(txtCelular.Text) <> 9 Or Len(txtCelular.Text) < 9 Then
                  MsgBox "El número de celular debe de tener 9 caracteres", vbInformation, "AVISO"
                  ValidarDatos = False
            txtCelular.SetFocus
            Exit Function
                  Exit Function
        End If
          If txtTelefono.Text = "" Then
            MsgBox "Falta ingresar el número de telefono", vbInformation, "AVISO"
            ValidarDatos = False
            txtTelefono.SetFocus
            Exit Function
        End If
        
        If Len(txtTelefono.Text) > 10 Or Len(txtTelefono.Text) < 6 Then
            MsgBox "El número de telefono debe de tener 9 caracteres", vbInformation, "AVISO"
            ValidarDatos = False
            txtTelefono.SetFocus
            Exit Function
        End If
        'FIN PTI1 ****************************
End Function
'FIN FRHU
Private Sub ImprimirPdfCartillaAutorizacion() 'add PTI1 ERS070-2018 11/12/2018
    Dim sParrafoUno As String
    Dim sParrafoDos As String
    Dim sParrafoTres As String
    Dim sParrafoCuatro As String
    Dim sParrafoCinco As String
    Dim sParrafoSeis As String
    Dim sParrafoSiete As String
    Dim sParrafoOcho As String
    Dim oDoc As cPDF
    Dim nAltura As Integer
    
    Set oDoc = New cPDF
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Cartilla Autorización y Actualización de datos personales"
    oDoc.Title = "Cartilla Autorización y Actualización de datos personales"
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaAutorizacionActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
    'oDoc.LoadImageFromFile App.path & "\logo_cmacmaynas.bmp", "Logo"
    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo" 'O
    
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '<body>
    nAltura = 20
    oDoc.WTextBox 10, 10, 780, 575, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    'oDoc.WImage 60, 480, 35, 100, "Logo"
    oDoc.WImage 70, 460, 50, 100, "Logo" 'O
    oDoc.WTextBox 90, 50, 15, 500, "AUTORIZACIÓN PARA EL TRATAMIENTO DE DATOS PERSONALES", "F2", 11, hCenter 'agregado por pti1 ers070-2018 05/12/2018
     



    'sParrafoUno = "Yo " & MatPersona(2).sNombres & " " & MatPersona(2).sApePat & " " & MatPersona(2).sApeMat & IIf(Len(MatPersona(2).sApeCas) = 0, "", " " & IIf(MatPersona(2).sEstadoCivil = "2", "DE", "VDA") & " " & MatPersona(2).sApeCas) & " con DOI N° " & Trim(MatPersona(2).sPersIDnro) & " autorizo y otorgo por tiempo  " & _
    '              "indefinido, mi consentimiento libre, previo, expreso, inequívoco e informado a la CAJA MUNICIPAL DE AHORRO Y CRÉDITO DE MAYNAS S.A. (en adelante, LA CAJA), " &
    '              "legislación  vigente. Datos personales que la Caja queda autorizada por el cliente a mantenerlos en su(s) base(s) de datos, así como para que sean almacenados, " &
    '              "servicio, así como resultado de la suscripción de contratos, formularios, y a los recopilados anteriormente, actualmente y/o por recopilar por LA CAJA. Asimismo, otorgo " &
    '              "otorgo mi autorización para el envío de información promocional y/o publicitaria de los servicios y productos que LA CAJA ofrece, a través de cualquier medio " &
    '              "de comunicación que se considere apropiado para su difusión, y para su uso en la gestión administrativa y comercial de LA CAJA que guarde relación con su objeto social. " &
    '              "En ese sentido, autorizo a LA CAJA al uso de mis datos personales para tratamientos que supongan el" & _
    '              "desarrollo de acciones y actividades comerciales, incluyendo la realización de estudios de mercado, elaboración de perfiles de compra y " & _
     '             "evaluaciones financieras. El uso y tratamiento de mis datos personales, se sujetan a lo establecido por el artículo 13° de la Ley N° 29733 - Ley de Protección de Datos Personales. "
    'sParrafoUno = "Yo " & String(78, vbTab) & "  con DOI N° " & String(20, vbTab) & "autorizo y otorgo por tiempo  "
    oDoc.WTextBox 125, 56, 360, 520, (MatPersona(2).sNombres & " " & MatPersona(2).sApePat & " " & MatPersona(2).sApeMat & IIf(Len(MatPersona(2).sApeCas) = 0, "", " " & IIf(MatPersona(2).sEstadoCivil = "2", "DE", "VDA") & " " & MatPersona(2).sApeCas)), "F1", 11, hjustify
    oDoc.WTextBox 125, 484, 360, 520, (Trim(MatPersona(2).sPersIDnro)), "F1", 11, hjustify
    oDoc.WTextBox 125, 56, 360, 520, ("___________________________________________________________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 481, 360, 520, ("____________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 35, 10, 520, ("Yo, " & String(120, vbTab) & "  con DOI N° " & String(22, vbTab) & ""), "F1", 11, hjustify
    
 
     sParrafoUno = "autorizo y otorgo por tiempo indefinido, " & String(0.52, vbTab) & "mi consentimiento libre, previo, expreso, inequívoco e informado a" & Chr$(13) & _
                   "la " & String(0.52, vbTab) & "CAJA MUNICIPAL DE AHORRO Y CRÉDITO DE MAYNAS " & String(0.52, vbTab) & "S.A. " & String(0.52, vbTab) & "(en " & String(0.52, vbTab) & "adelante," & String(0.52, vbTab) & " ""LA CAJA""), " & String(0.51, vbTab) & " para " & String(0.51, vbTab) & " el" & Chr$(13) & _
                   "tratamiento de mis datos personales proporcionados " & String(0.7, vbTab) & " en contexto de la contratación de cualquier producto " & Chr$(13) & _
                   "(activo y/o pasivo)" & String(0.52, vbTab) & " o" & String(0.51, vbTab) & " servicio, " & String(0.52, vbTab) & " así " & String(0.52, vbTab) & "como " & String(0.52, vbTab) & "resultado" & String(0.52, vbTab) & "de " & String(0.52, vbTab) & " la suscripción de contratos, " & String(0.52, vbTab) & " formularios, " & String(0.52, vbTab) & " y a los " & Chr$(13) & _
                   "recopilados anteriormente, actualmente y/o por recopilar por " & String(0.52, vbTab) & "LA CAJA. " & String(0.53, vbTab) & "Asimismo, " & String(0.53, vbTab) & "otorgo " & String(0.53, vbTab) & "mi autorización" & Chr$(13) & _
                   "para el envío de información  promocional y/o publicitaria de los servicios y productos que" & String(0.53, vbTab) & " LA CAJA ofrece, " & Chr$(13) & _
                   "a tráves de cualquier medio de comunicación que se considere apropiado para su difusión, " & String(0.53, vbTab) & "y " & String(0.52, vbTab) & "para" & String(0.53, vbTab) & " su uso " & Chr$(13) & _
                   "en la gestión administrativa " & String(0.53, vbTab) & " y " & String(0.5, vbTab) & " comercial de  " & String(0.53, vbTab) & "LA  " & String(0.53, vbTab) & "CAJA " & String(0.53, vbTab) & " que guarde relación con su objeto social.  " & String(0.53, vbTab) & "En " & String(0.52, vbTab) & "ese " & Chr$(13) & _
                   "sentido, autorizo a LA CAJA al uso de mis datos personales para tratamientos que supongan el " & String(0.52, vbTab) & "desarrollo" & Chr$(13) & _
                   "de acciones y actividades comerciales, incluyendo la realización de estudios  de  mercado, " & String(0.53, vbTab) & " elaboración " & String(0.52, vbTab) & "de" & Chr$(13) & _
                   "perfiles de compra " & String(0.53, vbTab) & " y evaluaciones financieras. " & String(0.54, vbTab) & " El uso y tratamiento de mis datos personales, " & String(0.54, vbTab) & "se sujetan" & Chr$(13) & _
                   "a lo establecido por el artículo 13° de la Ley N° 29733 - Ley de Protección de Datos Personales."
    
  
    sParrafoDos = "Declaro conocer el compromiso de " & String(0.52, vbTab) & "LA CAJA " & String(0.52, vbTab) & " por garantizar el mantenimiento de la confidencialidad" & String(0.52, vbTab) & " y " & String(0.52, vbTab) & "el " & Chr$(13) & _
                  "tratamiento seguro de mis datos personales, incluyendo el resguardo en las transferencias de " & String(0.52, vbTab) & "los mismos, " & Chr$(13) & _
                  "que se realicen " & String(0.53, vbTab) & "en cumplimiento de la " & String(0.55, vbTab) & " Ley N° 29733 - Ley de Protección " & String(0.53, vbTab) & " de Datos Personales. De" & String(0.53, vbTab) & "igual " & Chr$(13) & _
                  "manera, declaro " & String(0.52, vbTab) & "conocer que los datos personales " & String(0.55, vbTab) & "proporcionados por mi persona serán incorporados " & String(0.52, vbTab) & "al " & Chr$(13) & _
                  "Banco de Datos de Clientes de  " & String(0.6, vbTab) & " LA CAJA, el cual  " & String(0.55, vbTab) & "se encuentra debidamente registrado ante la" & String(0.52, vbTab) & " Dirección " & Chr$(13) & _
                  "Nacional  " & String(0.55, vbTab) & " de  " & String(0.55, vbTab) & " Protección de Datos " & String(0.55, vbTab) & "Personales, para lo cual " & String(0.55, vbTab) & " autorizo a LA CAJA " & String(0.52, vbTab) & "que " & String(0.55, vbTab) & " recopile, registre, " & Chr$(13) & _
                  "organice, " & String(0.55, vbTab) & "almacene, " & String(0.55, vbTab) & "conserve, bloquee, suprima, extraiga, consulte, utilice, transfiera, exporte, importe" & String(0.52, vbTab) & " o " & Chr$(13) & _
                  "procese de cualquier otra forma mis datos personales, con las limitaciones que prevé la Ley."
                 
                 
    sParrafoTres = "Del mismo modo, y siempre que así lo estime necesario, declaro conocer que podré ejercitar mis derechos " & Chr$(13) & _
                   "de " & String(0.55, vbTab) & " acceso, " & String(0.56, vbTab) & " rectificación, " & String(0.58, vbTab) & " cancelación " & String(0.55, vbTab) & " y " & String(0.55, vbTab) & " oposición relativos a este tratamiento, de conformidad " & String(0.52, vbTab) & "con lo " & Chr$(13) & _
                   "establecido" & String(0.51, vbTab) & " en " & String(0.5, vbTab) & "el " & String(0.6, vbTab) & " Titulo" & String(0.54, vbTab) & " III " & String(0.54, vbTab) & " de la Ley N° 29733 - Ley de Protección de Datos " & String(0.52, vbTab) & " Personales" & String(0.52, vbTab) & " acercándome " & Chr$(13) & _
                   "a cualquiera de las Agencias de LA CAJA a nivel nacional."

   sParrafoCuatro = "Asimismo, " & String(1.4, vbTab) & " declaro " & String(1.4, vbTab) & " conocer " & String(1.4, vbTab) & " el " & String(1.4, vbTab) & "compromiso " & String(1.4, vbTab) & " de " & String(1.4, vbTab) & " LA " & String(1.4, vbTab) & "CAJA " & String(1.4, vbTab) & " por " & String(1.4, vbTab) & "respetar " & String(1.4, vbTab) & "los " & String(1.4, vbTab) & "principios " & String(1.4, vbTab) & "de " & String(1.4, vbTab) & " legalidad, " & Chr$(13) & _
                    "consentimiento, finalidad, proporcionalidad, calidad, disposición de recurso, y nivel de protección adecuado," & Chr$(13) & _
                    "conforme lo dispone la Ley N° 29733 - Ley de Protección de Datos Personales," & String(1.4, vbTab) & " para " & String(1.4, vbTab) & "el " & String(1.4, vbTab) & "tratamiento de los" & Chr$(13) & _
                    "datos personales otorgados por mi persona."
                  
    sParrafoCinco = "Esta autorización es" & String(1.5, vbTab) & " indefinida y se mantendrá inclusive" & String(0.5, vbTab) & " después de terminada(s) la(s) operación(es)" & String(0.52, vbTab) & " y/o " & Chr$(13) & _
                    "el(los) Contrato(s) que tenga" & String(1.5, vbTab) & " o pueda tener con LA CAJA" & String(1.3, vbTab) & " sin perjuicio de " & String(0.5, vbTab) & "poder ejercer mis derechos " & String(0.52, vbTab) & "de " & Chr$(13) & _
                    "acceso, rectificación, cancelación y oposición mencionados en el presente documento."
                  
     Dim cfecha  As String 'pti1 add
     cfecha = Choose(Month(gdFecSis), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
                                        
 
            nTamanio = Len(sParrafoUno)
            spacvar = 23
            Spac = 138
            Index = 1
            Princ = 1
            CantCarac = 0
            
            nTamLet = 6: contador = 0: nCentrar = 80
            
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoUno, Index, CantCarac)
                        oDoc.WTextBox Spac, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoUno, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoUno, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoDos)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoDos, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoDos, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoDos, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoTres)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoTres, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoTres, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoTres, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoCuatro)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCuatro, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCuatro, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCuatro, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            
            nTamanio = Len(sParrafoCinco)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCinco, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCinco, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCinco, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
    
                  'oDoc.WTextBox 125, 30, 360, 520, sParrafoUno, "F1", 11, hjustify
                  'oDoc.WTextBox 277, 30, 360, 520, sParrafoDos, "F1", 11, hjustify
                  'oDoc.WTextBox 376, 30, 360, 520, sParrafoTres, "F1", 11, hjustify
                  'oDoc.WTextBox 432, 30, 360, 520, sParrafoCuatro, "F1", 11, hjustify
                  'oDoc.WTextBox 484, 30, 360, 520, sParrafoCinco, "F1", 11, hjustify
    


    oDoc.WTextBox 610, 35, 60, 300, ("En " & sCiudadAgencia & " a los " & Day(gdFecSis) & " días del mes de " & cfecha & " de " & Year(gdFecSis) & "."), "F1", 11, hLeft 'O  agregado  por pti1
    oDoc.WTextBox 670, 35, 90, 200, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 730, 35, 60, 180, "________________________________________", "F1", 8, hCenter
    oDoc.WTextBox 745, 90, 60, 80, "Firma", "F1", 10, hCenter
    
    sParrafoSeis = "¿Autorizas a Caja Maynas para el tratamiento de sus datos personales?"
    
    oDoc.WTextBox 670, 280, 60, 250, sParrafoSeis, "F1", 11, hLeft 'O  agregado  por pti1
   
   
    oDoc.WTextBox 712, 300, 15, 20, "SI", "F1", 8, hCenter
    oDoc.WTextBox 742, 300, 15, 20, "NO", "F1", 8, hCenter
    
    oDoc.WTextBox 690, 420, 70, 80, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    If optSi.value = True And optNo.value = False Then
        oDoc.WTextBox 710, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    ElseIf optSi.value = False And optNo.value = True Then
        oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 740, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    End If
    

            
    oDoc.PDFClose
    oDoc.Show
    '</body>
End Sub
Private Sub ImprimirPdfCartilla()
    Dim sParrafoUno As String
    Dim sParrafoDos As String
    Dim oDoc As cPDF
    Dim nAltura As Integer
    
    Set oDoc = New cPDF
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Cartilla Autorización y Actualización de datos personales"
    oDoc.Title = "Cartilla Autorización y Actualización de datos personales"
    
    'If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaAutorizacionActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then 'comentado por pti1 ers070-2018 27/12/2018
    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
    'oDoc.LoadImageFromFile App.path & "\logo_cmacmaynas.bmp", "Logo"
    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo" 'O
    
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '<body>
    nAltura = 20
    oDoc.WTextBox 10, 10, 780, 575, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    'oDoc.WImage 60, 480, 35, 100, "Logo"
    oDoc.WImage 65, 480, 50, 100, "Logo" 'O
    oDoc.WTextBox 80, 50, 15, 500, "CARTILLA DE ACTUALIZACIÓN DE DATOS PERSONALES", "F2", 11, hCenter 'add pti1
    'oDoc.WTextBox 80, 50, 15, 500, "CARTILLA AUTORIZACIÓN Y ACTUALIZACIÓN DE DATOS PERSONALES", "F2", 11, hCenter 'comentado por pti1
    
    If (Day(gdFecSis)) > 9 Then
        oDoc.WTextBox 119, 32, 20, 500, (Day(gdFecSis)), "F1", 9, hLeft 'ADD POR PTI1
    Else
        oDoc.WTextBox 119, 32, 20, 500, "0" & (Day(gdFecSis)), "F1", 9, hLeft
    End If
    If (Month(gdFecSis)) > 9 Then
        oDoc.WTextBox 119, 62, 20, 500, (Month(gdFecSis)), "F1", 9, hLeft
    Else
        oDoc.WTextBox 119, 62, 20, 500, "0" & (Month(gdFecSis)), "F1", 9, hLeft 'FIN ADD POR PTI1
    End If
    oDoc.WTextBox 119, 95, 20, 500, (Year(gdFecSis)), "F1", 9, hLeft 'ADD POR PTI1
    oDoc.WTextBox 120, 20, 20, 500, "______/______/________", "F1", 9, hLeft 'DESCOMENTADO POR PTI1
    oDoc.WTextBox 130 + nAltura, 20, 20, 50, "Nombre(s):", "F1", 9, hLeft
    oDoc.WTextBox 130 + nAltura, 70, 20, 500, MatPersona(2).sNombres, "F1", 9, hLeft '100 caracteres
    oDoc.WTextBox 130 + nAltura, 70, 20, 500, "____________________________________________________________________________________________________", "F1", 9, hLeft
    oDoc.WTextBox 160 + nAltura, 20, 20, 50, "Apellidos: ", "F1", 9, hLeft
    oDoc.WTextBox 160 + nAltura, 70, 20, 400, MatPersona(2).sApePat & " " & MatPersona(2).sApeMat & IIf(Len(MatPersona(2).sApeCas) = 0, "", " " & IIf(MatPersona(2).sEstadoCivil = "2", "DE", "VDA") & " " & MatPersona(2).sApeCas), "F1", 9, hLeft '64 caracteres
    oDoc.WTextBox 160 + nAltura, 70, 20, 400, "________________________________________________________________", "F1", 9, hLeft
    oDoc.WTextBox 160 + nAltura, 400, 20, 55, "Estado Civil: ", "F1", 9, hLeft
    oDoc.WTextBox 160 + nAltura, 455, 20, 250, Trim(Left(cmbPersNatEstCiv.Text, 23)), "F1", 9, hLeft '23 caracteres
    oDoc.WTextBox 160 + nAltura, 455, 20, 250, "_______________________", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 20, 20, 140, "Documento de Identificación:", "F1", 9, hLeft
    
    oDoc.WTextBox 190 + nAltura, 150, 20, 30, "D.N.I.", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 180, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    If MatPersona(2).sPersIDTpo = "1" Then oDoc.WTextBox 190 + nAltura, 180, 10, 20, "X", "F1", 8, hCenter
    
    oDoc.WTextBox 190 + nAltura, 210, 20, 80, "Carnet Extranjeria", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 290, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    If MatPersona(2).sPersIDTpo = "4" Then oDoc.WTextBox 190 + nAltura, 290, 10, 20, "X", "F1", 8, hCenter
    
    oDoc.WTextBox 190 + nAltura, 320, 20, 50, "Pasaporte", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 370, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    If MatPersona(2).sPersIDTpo = "11" Then oDoc.WTextBox 190 + nAltura, 370, 10, 20, "X", "F1", 8, hCenter
    
    oDoc.WTextBox 190 + nAltura, 420, 20, 10, "Nº ", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 435, 20, 200, Trim(MatPersona(2).sPersIDnro), "F1", 9, hLeft '27 caracteres
    oDoc.WTextBox 190 + nAltura, 435, 20, 200, "___________________________", "F1", 9, hLeft
    oDoc.WTextBox 220 + nAltura, 20, 20, 100, "Dirección de domicilio:", "F1", 9, hLeft
    oDoc.WTextBox 220 + nAltura, 115, 20, 500, Left(Trim(MatPersona(2).sDomicilio), 90), "F1", 9, hLeft '91 caracteres
    oDoc.WTextBox 220 + nAltura, 115, 20, 500, "___________________________________________________________________________________________", "F1", 9, hLeft
    oDoc.WTextBox 250 + nAltura, 20, 20, 50, "Referencia:", "F1", 9, hLeft
    oDoc.WTextBox 250 + nAltura, 70, 20, 500, Left(Trim(MatPersona(2).sRefDomicilio), 100), "F1", 9, hLeft '100 caracteres
    oDoc.WTextBox 250 + nAltura, 70, 20, 500, "____________________________________________________________________________________________________", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 20, 20, 60, "Departamento:", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 85, 20, 130, Left(cmbPersUbiGeo(1).Text, 25), "F1", 9, hLeft '26
    oDoc.WTextBox 280 + nAltura, 85, 20, 130, "__________________________", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 220, 20, 50, "Provincia:", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 265, 20, 130, Left(cmbPersUbiGeo(2).Text, 25), "F1", 9, hLeft '26
    oDoc.WTextBox 280 + nAltura, 265, 20, 130, "__________________________", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 400, 20, 60, "Distrito:", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 435, 20, 135, Left(cmbPersUbiGeo(3).Text, 25), "F1", 9, hLeft '27
    oDoc.WTextBox 280 + nAltura, 435, 20, 135, "___________________________", "F1", 9, hLeft
    oDoc.WTextBox 310 + nAltura, 20, 20, 50, "Celular:", "F1", 9, hLeft
    oDoc.WTextBox 310 + nAltura, 70, 20, 120, MatPersona(2).sCelular, "F1", 9, hLeft '24
    oDoc.WTextBox 310 + nAltura, 70, 20, 120, "________________________", "F1", 9, hLeft
    oDoc.WTextBox 310 + nAltura, 195, 20, 50, "Teléfono:", "F1", 9, hLeft
    oDoc.WTextBox 310 + nAltura, 240, 20, 120, MatPersona(2).sTelefonos, "F1", 9, hLeft '24
    oDoc.WTextBox 310 + nAltura, 240, 20, 120, "________________________", "F1", 9, hLeft
    oDoc.WTextBox 340 + nAltura, 20, 20, 80, "Correo electrónico:", "F1", 9, hLeft
    oDoc.WTextBox 340 + nAltura, 100, 20, 260, Trim(MatPersona(2).sEmail), "F1", 9, hLeft '52
    oDoc.WTextBox 340 + nAltura, 100, 20, 260, "____________________________________________________", "F1", 9, hLeft
 'INICIO COMENTADO POR PTI1
'    sParrafoUno = "Autorización para Recopilación y Tratamiento de Datos: Ley de protección de datos personales - N° 29733 (en adelante la Ley): por el presente documento el cliente " & _
'                  "entrega a la Caja datos personales que lo identifican y/o lo hacen identificable, y que son considerados datos personales conforme a las disposiciones de la Ley y la " & _
'                  "legislación  vigente. Datos personales que la Caja queda autorizada por el cliente a mantenerlos en su(s) base(s) de datos, así como para que sean almacenados, " & _
'                  "sistematizados y utilizados para los fines que se detallan en el presente documento. El cliente también autoriza a la Caja que: (i) la información podrá ser conservada " & _
'                  "por la Caja de forma indefinida e independientemente de la relación contractual que mantenga o no con la Caja, (ii) su información está protegida por las leyes " & _
'                  "aplicables y procedimientos que la Caja tiene implementados para el ejercicio de sus derechos, con el objeto que se evite la alteración, pérdida o acceso de personas " & _
'                  "personales con terceras personas, dentro o fuera del país, vinculadas o no a la Caja, exclusivamente para la tercerización de tratamientos autorizados y de " & _
'                  "conformidad con las medidas de seguridad exigidas por la Ley, v) la Caja transfiera o comparta sus datos personales con empresas vinculadas a la Caja y/o terceros, " & _
'                  "para fines de publicidad, mercadeo y similares, vi) le envíe, a través de mensajes de texto a su teléfono celular (SMS), llamadas telefónicas a su teléfono fijo o celular, " & _
'                  "mensajes de correo electrónico a su correo personal o comunicaciones enviadas a su domicilio, promociones e información relacionada a los servicios y productos " & _
'                  "que la Caja, sus subsidiarias o afiliadas ofrecen directa o indirectamente a través de las distintas asociaciones comerciales que la Caja pueda tener, e inclusive " & _
'                  "requerimientos de cobranza, directamente o a través de terceros, respecto de las deudas que pueda mantener el cliente con la Caja.  Asimismo, conforme a lo " & _
'                  "estipulado en la ley, el cliente tiene conocimiento que cuenta con el derecho de actualizar, incluir, rectificar y suprimir sus datos personales, así como a oponerse a su " & _
'                  "tratamiento para los fines antes indicados. El cliente también conoce que en cualquier momento, puede revocar la presente autorización para tratar sus datos " & _
'                  "personales, lo cual surtirá efectos en un plazo no mayor de 5 días calendario contados desde el día siguiente de recibida la comunicación. La revocación no surtirá " & _
'                  "efecto frente a hechos cumplidos, ni frente al tratamiento que sea necesario para la ejecución de una relación contractual vigente o sus consecuencias legales, ni " & _
'                  "podrá oponerse a tratamientos permitidos por ley. Para ejercer el derecho de revocatoria o cualquier otro que la Ley establezca con relación a sus datos personales, " & _
'                  "el cliente deberá dirigir una comunicación escrita por cualquiera de los canales de atención proporcionados por la Caja conforme a la Ley."
'    oDoc.WTextBox 400, 20, 360, 555, sParrafoUno, "F1", 7, hjustify
'
'    sParrafoDos = "Nota.- El cliente declara que, antes de suscribir el presente documento, ha sido informado que tiene derecho a no proporcionar a la Caja la autorización para el " & _
'                  "tratamiento de sus datos personales y que si no la proporciona la Caja no podrá tratar sus datos personales en la forma explicada en éste documento, lo que no " & _
'                  "impide su uso para la ejecución y cumplimiento de cualquier relación contractual que mantenga el cliente con la Caja."
'    oDoc.WTextBox 540, 20, 60, 555, sParrafoDos, "F1", 7, hjustify
' FIN COMENTADO POR PTI1

     Dim cfecha  As String 'pti1 add
     cfecha = Choose(Month(gdFecSis), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
                                        
                                        
    If (Day(gdFecSis)) > 9 Then
        oDoc.WTextBox 410, 25, 20, 200, (Day(gdFecSis)), "F1", 9, hLeft 'ADD POR PTI1
    Else
        oDoc.WTextBox 410, 25, 20, 200, "0" & (Day(gdFecSis)), "F1", 9, hLeft
    End If
    oDoc.WTextBox 410, 65, 80, 200, (cfecha), "F1", 9, hLeft  ' ADD PTI1
    oDoc.WTextBox 410, 155, 110, 200, Right(Year(gdFecSis), 2), "F1", 9, hLeft ' ADD PTI1
    oDoc.WTextBox 410, 20, 60, 200, "____ de ______________ del 20____", "F1", 9, hLeft ' DESCOMENTADO POR PTI1
    'oDoc.WTextBox 580, 20, 60, 200, ArmaFecha(gdFecSis), "F1", 9, hLeft 'O ' COMENTADO POR PTI1
    oDoc.WTextBox 490, 20, 60, 50, "Firma:", "F1", 9, hLeft
    oDoc.WTextBox 490, 50, 60, 150, "___________________________", "F1", 9, hLeft
    
    oDoc.WTextBox 420, 200, 80, 70, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    oDoc.WTextBox 635, 300, 60, 100, "Acepto", "F1", 9, hLeft ' INICIO COMENTADO POR PTI1
'    oDoc.WTextBox 650, 300, 60, 150, "Autorizar uso de mis datos", "F1", 9, hLeft
'
'    oDoc.WTextBox 630, 420, 15, 20, "SI", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'    oDoc.WTextBox 630, 440, 15, 20, "NO", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'
'    If optSi.value = True And optNo.value = False Then
'        oDoc.WTextBox 645, 420, 15, 20, "X", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        oDoc.WTextBox 645, 440, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    ElseIf optSi.value = False And optNo.value = True Then
'        oDoc.WTextBox 645, 420, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        oDoc.WTextBox 645, 440, 15, 20, "X", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    End If ' CFIN COMENTADO POR ERS070
            
    oDoc.PDFClose
    oDoc.Show
    '</body>
End Sub

Private Sub ImprimirPdfCartillaVistaPrueba()
    Dim sParrafoUno As String
    Dim sParrafoDos As String
    Dim sParrafoTres As String ' add pti1
    Dim sParrafoCuatro As String ' add pti1
    Dim sParrafoCinco As String ' add pti1
     Dim sParrafoSeis As String
    Dim oDoc As cPDF
    Dim nAltura As Integer
    
    Set oDoc = New cPDF
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Cartilla Autorización y Actualización de datos personales"
    oDoc.Title = "Cartilla Autorización y Actualización de datos personales"
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaAutorizacionActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo"
    
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '<body>
    nAltura = 20
'    oDoc.WTextBox 10, 10, 780, 575, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack 'comentado por pti1
'    oDoc.WImage 65, 480, 50, 100, "Logo"
'    oDoc.WTextBox 80, 50, 15, 500, "CARTILLA AUTORIZACIÓN Y ACTUALIZACIÓN DE DATOS PERSONALES", "F2", 11, hCenter 'fin comentado por pti1
    
    oDoc.WTextBox 10, 10, 780, 575, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack 'agregado por pti1 ers070-2018 05/12/2018
    oDoc.WImage 70, 460, 50, 100, "Logo" 'agregado por pti1 ers070-2018 05/12/2018
    oDoc.WTextBox 90, 50, 15, 500, "AUTORIZACIÓN PARA EL TRATAMIENTO DE DATOS PERSONALES", "F2", 11, hCenter 'agregado por pti1 ers070-2018 05/12/2018
    
    'agregado por pti1 ers070-2018 05/12/2018 ********************
                 
    'sParrafoUno = "Yo " & Trim(txtNombre.Text) & " " & Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & IIf(Len(Trim(txtApellidoCasada.Text)) = 0, "", " " & IIf(Trim(Right(cmbTipoDOI.Text, 2)) = "2", "DE", "VDA") & " " & Trim(txtApellidoCasada.Text)) &
    '             "" & IIf(Trim(Right(cmbTipoDOI.Text, 2)) = "1", " con D.N.I N° " & Trim(txtDOI.Text), " ") & " " &
    '           "" & IIf(Trim(Right(cmbTipoDOI.Text, 2)) = "4", " con carnet de extranjería N° " & Trim(txtDOI.Text), " ") & " " &
    '           "" & IIf(Trim(Right(cmbTipoDOI.Text, 2)) = "11", " con pasaporte N° " & Trim(txtDOI.Text), "") & " " &
    '           " autorizo y otorgo por tiempo  " & _
    '           "indefinido, mi consentimiento libre, previo, expreso, inequívoco e informado a la CAJA MUNICIPAL DE AHORRO Y CRÉDITO DE MAYNAS S.A. (en adelante, LA CAJA), " &
    '           "legislación  vigente. Datos personales que la Caja queda autorizada por el cliente a mantenerlos en su(s) base(s) de datos, así como para que sean almacenados, " &
    '           "servicio, así como resultado de la suscripción de contratos, formularios, y a los recopilados anteriormente, actualmente y/o por recopilar por LA CAJA. Asimismo, otorgo " &
    '           "otorgo mi autorización para el envío de información promocional y/o publicitaria de los servicios y productos que LA CAJA ofrece, a través de cualquier medio " &
    '           "de comunicación que se considere apropiado para su difusión, y para su uso en la gestión administrativa y comercial de LA CAJA que guarde relación con su objeto social. " &
    '           "En ese sentido, autorizo a LA CAJA al uso de mis datos personales para tratamientos que supongan el" & _
    '           "desarrollo de acciones y actividades comerciales, incluyendo la realización de estudios de mercado, elaboración de perfiles de compra y " & _
    '           "evaluaciones financieras. El uso y tratamiento de mis datos personales, se sujetan a lo establecido por el artículo 13° de la Ley N° 29733 - Ley de Protección de Datos Personales. "
    
   
    oDoc.WTextBox 125, 56, 360, 520, (Trim(txtNombre.Text) & " " & Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & IIf(Len(Trim(txtApellidoCasada.Text)) = 0, "", " " & IIf(Trim(Right(cmbPersNatEstCiv.Text, 2)) = "2", "DE", "VDA") & " " & Trim(txtApellidoCasada.Text))), "F1", 11, hjustify
    oDoc.WTextBox 125, 484, 360, 520, (Trim(txtDoi.Text)), "F1", 11, hjustify
    oDoc.WTextBox 125, 56, 360, 520, ("___________________________________________________________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 481, 360, 520, ("____________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 35, 10, 520, ("Yo, " & String(120, vbTab) & "  con DOI N° " & String(22, vbTab) & ""), "F1", 11, hjustify
 
    sParrafoUno = "autorizo y otorgo por tiempo indefinido, " & String(0.52, vbTab) & "mi consentimiento libre, previo, expreso, inequívoco e informado a" & Chr$(13) & _
                   "la " & String(0.52, vbTab) & "CAJA MUNICIPAL DE AHORRO Y CRÉDITO DE MAYNAS " & String(0.52, vbTab) & "S.A. " & String(0.52, vbTab) & "(en " & String(0.52, vbTab) & "adelante," & String(0.52, vbTab) & " ""LA CAJA""), " & String(0.51, vbTab) & " para " & String(0.51, vbTab) & " el" & Chr$(13) & _
                   "tratamiento de mis datos personales proporcionados " & String(0.7, vbTab) & " en contexto de la contratación de cualquier producto " & Chr$(13) & _
                   "(activo y/o pasivo)" & String(0.52, vbTab) & " o" & String(0.51, vbTab) & " servicio, " & String(0.52, vbTab) & " así " & String(0.52, vbTab) & "como " & String(0.52, vbTab) & "resultado" & String(0.52, vbTab) & "de " & String(0.52, vbTab) & " la suscripción de contratos, " & String(0.52, vbTab) & " formularios, " & String(0.52, vbTab) & " y a los " & Chr$(13) & _
                   "recopilados anteriormente, actualmente y/o por recopilar por " & String(0.52, vbTab) & "LA CAJA. " & String(0.53, vbTab) & "Asimismo, " & String(0.53, vbTab) & "otorgo " & String(0.53, vbTab) & "mi autorización" & Chr$(13) & _
                   "para el envío de información  promocional y/o publicitaria de los servicios y productos que" & String(0.53, vbTab) & " LA CAJA ofrece, " & Chr$(13) & _
                   "a tráves de cualquier medio de comunicación que se considere apropiado para su difusión, " & String(0.53, vbTab) & "y " & String(0.52, vbTab) & "para" & String(0.53, vbTab) & " su uso " & Chr$(13) & _
                   "en la gestión administrativa " & String(0.53, vbTab) & " y " & String(0.5, vbTab) & " comercial de  " & String(0.53, vbTab) & "LA  " & String(0.53, vbTab) & "CAJA " & String(0.53, vbTab) & " que guarde relación con su objeto social.  " & String(0.53, vbTab) & "En " & String(0.52, vbTab) & "ese " & Chr$(13) & _
                   "sentido, autorizo a LA CAJA al uso de mis datos personales para tratamientos que supongan el " & String(0.52, vbTab) & "desarrollo" & Chr$(13) & _
                   "de acciones y actividades comerciales, incluyendo la realización de estudios  de  mercado, " & String(0.53, vbTab) & " elaboración " & String(0.52, vbTab) & "de" & Chr$(13) & _
                   "perfiles de compra " & String(0.53, vbTab) & " y evaluaciones financieras. " & String(0.54, vbTab) & " El uso y tratamiento de mis datos personales, " & String(0.54, vbTab) & "se sujetan" & Chr$(13) & _
                   "a lo establecido por el artículo 13° de la Ley N° 29733 - Ley de Protección de Datos Personales."
    
  
    sParrafoDos = "Declaro conocer el compromiso de " & String(0.52, vbTab) & "LA CAJA " & String(0.52, vbTab) & " por garantizar el mantenimiento de la confidencialidad" & String(0.52, vbTab) & " y " & String(0.52, vbTab) & "el " & Chr$(13) & _
                  "tratamiento seguro de mis datos personales, incluyendo el resguardo en las transferencias de " & String(0.52, vbTab) & "los mismos, " & Chr$(13) & _
                  "que se realicen " & String(0.53, vbTab) & "en cumplimiento de la " & String(0.55, vbTab) & " Ley N° 29733 - Ley de Protección " & String(0.53, vbTab) & " de Datos Personales. De" & String(0.53, vbTab) & "igual " & Chr$(13) & _
                  "manera, declaro " & String(0.52, vbTab) & "conocer que los datos personales " & String(0.55, vbTab) & "proporcionados por mi persona serán incorporados " & String(0.52, vbTab) & "al " & Chr$(13) & _
                  "Banco de Datos de Clientes de  " & String(0.6, vbTab) & " LA CAJA, el cual  " & String(0.55, vbTab) & "se encuentra debidamente registrado ante la" & String(0.52, vbTab) & " Dirección " & Chr$(13) & _
                  "Nacional  " & String(0.55, vbTab) & " de  " & String(0.55, vbTab) & " Protección de Datos " & String(0.55, vbTab) & "Personales, para lo cual " & String(0.55, vbTab) & " autorizo a LA CAJA " & String(0.52, vbTab) & "que " & String(0.55, vbTab) & " recopile, registre, " & Chr$(13) & _
                  "organice, " & String(0.55, vbTab) & "almacene, " & String(0.55, vbTab) & "conserve, bloquee, suprima, extraiga, consulte, utilice, transfiera, exporte, importe" & String(0.52, vbTab) & " o " & Chr$(13) & _
                  "procese de cualquier otra forma mis datos personales, con las limitaciones que prevé la Ley."
                 
                 
    sParrafoTres = "Del mismo modo, y siempre que así lo estime necesario, declaro conocer que podré ejercitar mis derechos " & Chr$(13) & _
                   "de " & String(0.55, vbTab) & " acceso, " & String(0.56, vbTab) & " rectificación, " & String(0.58, vbTab) & " cancelación " & String(0.55, vbTab) & " y " & String(0.55, vbTab) & " oposición relativos a este tratamiento, de conformidad " & String(0.52, vbTab) & "con lo " & Chr$(13) & _
                   "establecido" & String(0.51, vbTab) & " en " & String(0.5, vbTab) & "el " & String(0.6, vbTab) & " Titulo" & String(0.54, vbTab) & " III " & String(0.54, vbTab) & " de la Ley N° 29733 - Ley de Protección de Datos " & String(0.52, vbTab) & " Personales" & String(0.52, vbTab) & " acercándome " & Chr$(13) & _
                   "a cualquiera de las Agencias de LA CAJA a nivel nacional."

   sParrafoCuatro = "Asimismo, " & String(1.4, vbTab) & " declaro " & String(1.4, vbTab) & " conocer " & String(1.4, vbTab) & " el " & String(1.4, vbTab) & "compromiso " & String(1.4, vbTab) & " de " & String(1.4, vbTab) & " LA " & String(1.4, vbTab) & "CAJA " & String(1.4, vbTab) & " por " & String(1.4, vbTab) & "respetar " & String(1.4, vbTab) & "los " & String(1.4, vbTab) & "principios " & String(1.4, vbTab) & "de " & String(1.4, vbTab) & " legalidad, " & Chr$(13) & _
                    "consentimiento, finalidad, proporcionalidad, calidad, disposición de recurso, y nivel de protección adecuado," & Chr$(13) & _
                    "conforme lo dispone la Ley N° 29733 - Ley de Protección de Datos Personales," & String(1.4, vbTab) & " para " & String(1.4, vbTab) & "el " & String(1.4, vbTab) & "tratamiento de los" & Chr$(13) & _
                    "datos personales otorgados por mi persona."
                  
    sParrafoCinco = "Esta autorización es" & String(1.5, vbTab) & " indefinida y se mantendrá inclusive" & String(0.5, vbTab) & " después de terminada(s) la(s) operación(es)" & String(0.52, vbTab) & " y/o " & Chr$(13) & _
                    "el(los) Contrato(s) que tenga" & String(1.5, vbTab) & " o pueda tener con LA CAJA" & String(1.3, vbTab) & " sin perjuicio de " & String(0.5, vbTab) & "poder ejercer mis derechos " & String(0.52, vbTab) & "de " & Chr$(13) & _
                    "acceso, rectificación, cancelación y oposición mencionados en el presente documento."
                    
     Dim cfecha  As String 'pti1 add
     cfecha = Choose(Month(gdFecSis), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
                                        
              nTamanio = Len(sParrafoUno)
            spacvar = 23
            Spac = 138
            Index = 1
            Princ = 1
            CantCarac = 0
            
            nTamLet = 6: contador = 0: nCentrar = 80
            
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoUno, Index, CantCarac)
                        oDoc.WTextBox Spac, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoUno, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoUno, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoDos)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoDos, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoDos, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoDos, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoTres)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoTres, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoTres, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoTres, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoCuatro)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCuatro, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCuatro, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCuatro, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            
            nTamanio = Len(sParrafoCinco)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCinco, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCinco, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCinco, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
    
    
                  'oDoc.WTextBox 125, 30, 360, 520, sParrafoUno, "F1", 11, hjustify
                  'oDoc.WTextBox 277, 30, 360, 520, sParrafoDos, "F1", 11, hjustify
                  'oDoc.WTextBox 376, 30, 360, 520, sParrafoTres, "F1", 11, hjustify
                  'oDoc.WTextBox 432, 30, 360, 520, sParrafoCuatro, "F1", 11, hjustify
                  'oDoc.WTextBox 484, 30, 360, 520, sParrafoCinco, "F1", 11, hjustify
    
    oDoc.WTextBox 610, 35, 60, 520, ("En " & sCiudadAgencia & " a los " & Day(gdFecSis) & " días del mes de " & cfecha & " de " & Year(gdFecSis) & "."), "F1", 11, hLeft 'O  agregado  por pti1
    oDoc.WTextBox 670, 35, 90, 200, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 730, 35, 60, 180, "________________________________________", "F1", 8, hCenter
    oDoc.WTextBox 745, 90, 60, 80, "Firma", "F1", 10, hCenter
    
    sParrafoSeis = "¿Autorizas a Caja Maynas para el tratamiento de sus datos personales?"
    
    oDoc.WTextBox 670, 280, 60, 250, sParrafoSeis, "F1", 11, hLeft 'O  agregado  por pti1
   
   
    oDoc.WTextBox 712, 300, 15, 20, "SI", "F1", 8, hCenter
    oDoc.WTextBox 742, 300, 15, 20, "NO", "F1", 8, hCenter
    
    oDoc.WTextBox 690, 420, 70, 80, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    If optSi.value = True And optNo.value = False Then
        oDoc.WTextBox 710, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    ElseIf optSi.value = False And optNo.value = True Then
        oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 740, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    End If
         
  
' fin agregado por pti1 ers070-2018 05/12/2018

' inicio comentado por pti1 ers070-2018 11/12/2018
'    oDoc.WTextBox 130 + nAltura, 20, 20, 50, "Nombre(s):", "F1", 9, hLeft
'    oDoc.WTextBox 130 + nAltura, 70, 20, 500, Trim(txtNombre.Text), "F1", 9, hLeft '100 caracteres
'    oDoc.WTextBox 130 + nAltura, 70, 20, 500, "____________________________________________________________________________________________________", "F1", 9, hLeft
'    oDoc.WTextBox 160 + nAltura, 20, 20, 50, "Apellidos: ", "F1", 9, hLeft
'    oDoc.WTextBox 160 + nAltura, 70, 20, 400, txtApellidoPaterno.Text & " " & txtApellidoMaterno.Text & IIf(Len(txtApellidoCasada.Text) = 0, "", " " & IIf(Trim(Right(cmbPersNatEstCiv.Text, 2)) = "2", "DE", "VDA") & " " & txtApellidoCasada.Text), "F1", 9, hLeft '64 caracteres
'    oDoc.WTextBox 160 + nAltura, 70, 20, 400, "________________________________________________________________", "F1", 9, hLeft
'    oDoc.WTextBox 160 + nAltura, 400, 20, 55, "Estado Civil: ", "F1", 9, hLeft
'    oDoc.WTextBox 160 + nAltura, 455, 20, 250, Trim(Left(cmbPersNatEstCiv.Text, 23)), "F1", 9, hLeft '23 caracteres
'    oDoc.WTextBox 160 + nAltura, 455, 20, 250, "_______________________", "F1", 9, hLeft
'    oDoc.WTextBox 190 + nAltura, 20, 20, 140, "Documento de Identificación:", "F1", 9, hLeft
'
'    oDoc.WTextBox 190 + nAltura, 150, 20, 30, "D.N.I.", "F1", 9, hLeft
'    oDoc.WTextBox 190 + nAltura, 180, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'    If Trim(Right(cmbTipoDOI.Text, 2)) = "1" Then oDoc.WTextBox 190 + nAltura, 180, 10, 20, "X", "F1", 8, hCenter
'
'    oDoc.WTextBox 190 + nAltura, 210, 20, 80, "Carnet Extranjeria", "F1", 9, hLeft
'    oDoc.WTextBox 190 + nAltura, 290, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'    If Trim(Right(cmbTipoDOI.Text, 2)) = "4" Then oDoc.WTextBox 190 + nAltura, 290, 10, 20, "X", "F1", 8, hCenter
'
'    oDoc.WTextBox 190 + nAltura, 320, 20, 50, "Pasaporte", "F1", 9, hLeft
'    oDoc.WTextBox 190 + nAltura, 370, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'    If Trim(Right(cmbTipoDOI.Text, 2)) = "11" Then oDoc.WTextBox 190 + nAltura, 370, 10, 20, "X", "F1", 8, hCenter
'
'    oDoc.WTextBox 190 + nAltura, 420, 20, 10, "Nº ", "F1", 9, hLeft
'    oDoc.WTextBox 190 + nAltura, 435, 20, 200, Trim(txtDOI.Text), "F1", 9, hLeft '27 caracteres
'    oDoc.WTextBox 190 + nAltura, 435, 20, 200, "___________________________", "F1", 9, hLeft
'    oDoc.WTextBox 220 + nAltura, 20, 20, 100, "Dirección de domicilio:", "F1", 9, hLeft
'    oDoc.WTextBox 220 + nAltura, 115, 20, 500, Left(Trim(txtDireccion.Text), 90), "F1", 9, hLeft '91 caracteres
'    oDoc.WTextBox 220 + nAltura, 115, 20, 500, "___________________________________________________________________________________________", "F1", 9, hLeft
'    oDoc.WTextBox 250 + nAltura, 20, 20, 50, "Referencia:", "F1", 9, hLeft
'    oDoc.WTextBox 250 + nAltura, 70, 20, 500, Left(Trim(txtReferencia.Text), 100), "F1", 9, hLeft '100 caracteres
'    oDoc.WTextBox 250 + nAltura, 70, 20, 500, "____________________________________________________________________________________________________", "F1", 9, hLeft
'    oDoc.WTextBox 280 + nAltura, 20, 20, 60, "Departamento:", "F1", 9, hLeft
'    oDoc.WTextBox 280 + nAltura, 85, 20, 130, Left(cmbPersUbiGeo(1).Text, 25), "F1", 9, hLeft '26
'    oDoc.WTextBox 280 + nAltura, 85, 20, 130, "__________________________", "F1", 9, hLeft
'    oDoc.WTextBox 280 + nAltura, 220, 20, 50, "Provincia:", "F1", 9, hLeft
'    oDoc.WTextBox 280 + nAltura, 265, 20, 130, Left(cmbPersUbiGeo(2).Text, 25), "F1", 9, hLeft '26
'    oDoc.WTextBox 280 + nAltura, 265, 20, 130, "__________________________", "F1", 9, hLeft
'    oDoc.WTextBox 280 + nAltura, 400, 20, 60, "Distrito:", "F1", 9, hLeft
'    oDoc.WTextBox 280 + nAltura, 435, 20, 135, Left(cmbPersUbiGeo(3).Text, 25), "F1", 9, hLeft '27
'    oDoc.WTextBox 280 + nAltura, 435, 20, 135, "___________________________", "F1", 9, hLeft
'    oDoc.WTextBox 310 + nAltura, 20, 20, 50, "Celular:", "F1", 9, hLeft
'    oDoc.WTextBox 310 + nAltura, 70, 20, 120, Trim(txtCelular.Text), "F1", 9, hLeft '24
'    oDoc.WTextBox 310 + nAltura, 70, 20, 120, "________________________", "F1", 9, hLeft
'    oDoc.WTextBox 310 + nAltura, 195, 20, 50, "Teléfono:", "F1", 9, hLeft
'    oDoc.WTextBox 310 + nAltura, 240, 20, 120, Trim(txtTelefono.Text), "F1", 9, hLeft '24
'    oDoc.WTextBox 310 + nAltura, 240, 20, 120, "________________________", "F1", 9, hLeft
'    oDoc.WTextBox 340 + nAltura, 20, 20, 80, "Correo electrónico:", "F1", 9, hLeft
'    oDoc.WTextBox 340 + nAltura, 100, 20, 260, Trim(txtCorreo.Text), "F1", 9, hLeft '52
'    oDoc.WTextBox 340 + nAltura, 100, 20, 260, "____________________________________________________", "F1", 9, hLeft
'
'    sParrafoUno = "Autorización para Recopilación y Tratamiento de Datos: Ley de protección de datos personales - N° 29733 (en adelante la Ley): por el presente documento el cliente " & _
'                  "entrega a la Caja datos personales que lo identifican y/o lo hacen identificable, y que son considerados datos personales conforme a las disposiciones de la Ley y la " & _
'                  "legislación  vigente. Datos personales que la Caja queda autorizada por el cliente a mantenerlos en su(s) base(s) de datos, así como para que sean almacenados, " & _
'                  "sistematizados y utilizados para los fines que se detallan en el presente documento. El cliente también autoriza a la Caja que: (i) la información podrá ser conservada " & _
'                  "por la Caja de forma indefinida e independientemente de la relación contractual que mantenga o no con la Caja, (ii) su información está protegida por las leyes " & _
'                  "aplicables y procedimientos que la Caja tiene implementados para el ejercicio de sus derechos, con el objeto que se evite la alteración, pérdida o acceso de personas " & _
'                  "personales con terceras personas, dentro o fuera del país, vinculadas o no a la Caja, exclusivamente para la tercerización de tratamientos autorizados y de " & _
'                  "conformidad con las medidas de seguridad exigidas por la Ley, v) la Caja transfiera o comparta sus datos personales con empresas vinculadas a la Caja y/o terceros, " & _
'                  "para fines de publicidad, mercadeo y similares, vi) le envíe, a través de mensajes de texto a su teléfono celular (SMS), llamadas telefónicas a su teléfono fijo o celular, " & _
'                  "mensajes de correo electrónico a su correo personal o comunicaciones enviadas a su domicilio, promociones e información relacionada a los servicios y productos " & _
'                  "que la Caja, sus subsidiarias o afiliadas ofrecen directa o indirectamente a través de las distintas asociaciones comerciales que la Caja pueda tener, e inclusive " & _
'                  "requerimientos de cobranza, directamente o a través de terceros, respecto de las deudas que pueda mantener el cliente con la Caja.  Asimismo, conforme a lo " & _
'                  "estipulado en la ley, el cliente tiene conocimiento que cuenta con el derecho de actualizar, incluir, rectificar y suprimir sus datos personales, así como a oponerse a su " & _
'                  "tratamiento para los fines antes indicados. El cliente también conoce que en cualquier momento, puede revocar la presente autorización para tratar sus datos " & _
'                  "personales, lo cual surtirá efectos en un plazo no mayor de 5 días calendario contados desde el día siguiente de recibida la comunicación. La revocación no surtirá " & _
'                  "efecto frente a hechos cumplidos, ni frente al tratamiento que sea necesario para la ejecución de una relación contractual vigente o sus consecuencias legales, ni " & _
'                  "podrá oponerse a tratamientos permitidos por ley. Para ejercer el derecho de revocatoria o cualquier otro que la Ley establezca con relación a sus datos personales, " & _
'                  "el cliente deberá dirigir una comunicación escrita por cualquiera de los canales de atención proporcionados por la Caja conforme a la Ley."
'    oDoc.WTextBox 400, 20, 360, 555, sParrafoUno, "F1", 7, hjustify
'
'    sParrafoDos = "Nota.- El cliente declara que, antes de suscribir el presente documento, ha sido informado que tiene derecho a no proporcionar a la Caja la autorización para el " & _
'                  "tratamiento de sus datos personales y que si no la proporciona la Caja no podrá tratar sus datos personales en la forma explicada en éste documento, lo que no " & _
'                  "impide su uso para la ejecución y cumplimiento de cualquier relación contractual que mantenga el cliente con la Caja."
'    oDoc.WTextBox 540, 20, 60, 555, sParrafoDos, "F1", 7, hjustify
'
'    oDoc.WTextBox 580, 20, 60, 200, ArmaFecha(gdFecSis), "F1", 9, hLeft
'    'oDoc.WTextBox 580, 20, 60, 200, "____ de ______________ del ______", "F1", 9, hLeft
'    oDoc.WTextBox 650, 20, 60, 50, "Firma:", "F1", 9, hLeft
'    oDoc.WTextBox 650, 50, 60, 150, "___________________________", "F1", 9, hLeft
'
'    oDoc.WTextBox 610, 200, 50, 50, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    oDoc.WTextBox 635, 300, 60, 100, "Acepto", "F1", 9, hLeft
'    oDoc.WTextBox 650, 300, 60, 150, "Autorizar uso de mis datos", "F1", 9, hLeft
'
'    oDoc.WTextBox 630, 420, 15, 20, "SI", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'    oDoc.WTextBox 630, 440, 15, 20, "NO", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'
'    If optSi.value = True And optNo.value = False Then
'        oDoc.WTextBox 645, 420, 15, 20, "X", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        oDoc.WTextBox 645, 440, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    ElseIf optSi.value = False And optNo.value = True Then
'        oDoc.WTextBox 645, 420, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        oDoc.WTextBox 645, 440, 15, 20, "X", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    End If ' comentado por fin pti1 acta ers0-70 2018 11/12/2018
            
    oDoc.PDFClose
    oDoc.Show
    '</body>
End Sub
Private Sub cmdCerrar_Click()
'    Unload Me ' comentado por pti1 ers070-2018
'add pti1 ers 070-2018
If MsgBox("¿Desea salir sin guardar los cambios?", vbYesNo + vbQuestion, "Atención") = vbNo Then
    MsgBox "Se procederá a guardar los cambios y/o autorización", vbInformation, "AVISO"
    cmdAceptar_Click
    Else
    Unload Me
End If
End Sub
Private Sub cmdVistaPrevia_Click()
    Call ImprimirPdfCartillaVistaPrueba
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras(KeyAscii, True)
    If KeyAscii = 13 Then
        txtApellidoPaterno.SetFocus
    End If
End Sub
Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras(KeyAscii, True)
    If KeyAscii = 13 Then
        txtApellidoMaterno.SetFocus
    End If
End Sub
Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras2(KeyAscii, True)
    If KeyAscii = 13 Then
        If txtApellidoCasada.Visible Then
            txtApellidoCasada.SetFocus
        Else
            txtDoi.SetFocus
        End If
    End If
End Sub
Private Sub txtApellidoCasada_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras2(KeyAscii, True)
    If KeyAscii = 13 Then
        txtDoi.SetFocus
    End If
End Sub
Private Sub txtDOI_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersNatSexo.SetFocus
    End If
End Sub
Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtReferencia.SetFocus
    End If
End Sub
Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersUbiGeo(1).SetFocus
    End If
End Sub
Private Sub txtCelular_KeyPress(KeyAscii As Integer)
    If txtCelular.Text = "" Then
        If Not DigitoRPM(KeyAscii) Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                If (validarTelefonoControl(txtCelular)) Then
                   txtTelefono.SetFocus
                End If
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            If (validarTelefonoControl(txtCelular)) Then
                txtTelefono.SetFocus
            End If
        End If
    End If
End Sub
Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If (validarTelefonoControl(txtTelefono)) Then
           txtCorreo.SetFocus
        End If
    End If
End Sub
Private Function validarTelefonoControl(ctrControl As Control) As Boolean
    Dim x As New RegExp
    x.Pattern = "^(\*|\#)?\d+$"
    
    ctrControl.Text = RTrim(LTrim(ctrControl.Text))
    
    If Len(ctrControl.Text) = 0 Then
        validarTelefonoControl = True
        Exit Function
    ElseIf Not (x.Test(ctrControl.Text)) Then
        MsgBox "El Telefono debe contener solo numeros ó comenzar con # ó *", vbInformation, "Aviso"
        ctrControl.SetFocus
        validarTelefonoControl = False
        Exit Function
    Else
        Dim Numero As String
        Dim digito As String
        Dim flat As Boolean
        Dim n As Integer
        n = 6
        
        Numero = ctrControl.Text
        digito = Mid(Numero, 1, 1)
        
        If digito = "#" Or digito = "*" Then
            digito = (Mid(Numero, 2, 1))
            n = n + 1
        End If
        
        If (Len(Numero) < n) Then
            MsgBox "El Telefono deben tener por lo menos 6 digitos", vbInformation, "Aviso"
            ctrControl.SetFocus
            validarTelefonoControl = False
            Exit Function
        End If
        
        x.Pattern = "^(\*|\#)?" + digito + "+$"
        
        flat = Not (x.Test(ctrControl.Text))
        
        If Not flat Then
            MsgBox "No todos los digitos deben ser Iguales", vbInformation, "Aviso"
            ctrControl.SetFocus
        End If
        
        validarTelefonoControl = flat
    End If
End Function
Public Function ValidarNombresApellidos(ByVal sNombreApeC As String, ByVal nLargoNombreApeC As Integer, ByRef nSalir As Integer, ByVal nTipo) As Integer
Dim nI As Integer
Dim nJ As Integer
Dim sNombreApe As String
Dim nLargoNombreApe As Integer

sNombreApe = sNombreApeC
nLargoNombreApe = nLargoNombreApeC
If nSalir = 0 Then
For nI = 1 To nLargoNombreApe
    For nJ = 192 To 254
        If Mid(sNombreApe, nI, 1) = Chr(nJ) And ((nJ >= 192 And nJ <= 208) Or (nJ >= 210 And nJ <= 219) Or (nJ >= 221 And nJ <= 240) Or (nJ >= 242 And nJ <= 251) Or (nJ >= 253 And nJ <= 254)) Then
            If nTipo = 1 Then
                nSalir = 1
            ElseIf nTipo = 2 Then
                nSalir = 2
            ElseIf nTipo = 3 Then
                nSalir = 3
            End If
            Exit For
        End If
    Next nJ
    If nSalir = 1 Then
        Exit For
    End If
Next nI
End If
ValidarNombresApellidos = nSalir
End Function
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
