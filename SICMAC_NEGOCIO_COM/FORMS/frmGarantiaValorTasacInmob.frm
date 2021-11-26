VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaValorTasacInmob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TASACIÓN INMOBILIARIA"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "frmGarantiaValorTasacInmob.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmdET 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      ToolTipText     =   "Aceptar"
      Top             =   4530
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5025
      TabIndex        =   22
      ToolTipText     =   "Cancelar"
      Top             =   4530
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tasación"
      TabPicture(0)   =   "frmGarantiaValorTasacInmob.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraUbigeo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraValor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDescripcion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraTasador"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame fraTasador 
         Caption         =   "Tasación"
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
         Height          =   1040
         Left            =   120
         TabIndex        =   42
         Top             =   3160
         Width           =   9495
         Begin VB.TextBox txtTasacionTC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8520
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   240
            Width           =   840
         End
         Begin SICMACT.TxtBuscar txtTasadorCod 
            Height          =   255
            Left            =   960
            TabIndex        =   17
            Top             =   645
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
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
         Begin VB.TextBox txtTasadorNombre 
            Height          =   285
            Left            =   2475
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   650
            Width           =   3165
         End
         Begin VB.TextBox txtTasadorDNI 
            Height          =   285
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   645
            Width           =   1000
         End
         Begin VB.TextBox txtTasadorREPEV 
            Height          =   285
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   650
            Width           =   1200
         End
         Begin MSMask.MaskEdBox txtTasacionFecha 
            Height          =   330
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tasador :"
            Height          =   195
            Left            =   150
            TabIndex        =   47
            Top             =   660
            Width           =   675
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   150
            TabIndex        =   46
            Top             =   285
            Width           =   540
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "T/C Tasación :"
            Height          =   195
            Left            =   7320
            TabIndex        =   45
            Top             =   285
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "DNI : "
            Height          =   195
            Left            =   5790
            TabIndex        =   44
            Top             =   660
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "REPEV : "
            Height          =   195
            Left            =   7440
            TabIndex        =   43
            Top             =   660
            Width           =   675
         End
      End
      Begin VB.Frame fraDescripcion 
         Caption         =   "Descripción"
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
         Height          =   1040
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   9495
         Begin VB.TextBox txtAnioConstruccion 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   8700
            MaxLength       =   4
            TabIndex        =   14
            Tag             =   "txtPrincipal"
            Text            =   "2015"
            Top             =   650
            Width           =   650
         End
         Begin VB.TextBox txtNroSotanos 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            MaxLength       =   2
            TabIndex        =   13
            Tag             =   "txtPrincipal"
            Text            =   "0"
            Top             =   650
            Width           =   650
         End
         Begin VB.TextBox txtNroPisos 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4035
            MaxLength       =   2
            TabIndex        =   12
            Tag             =   "txtPrincipal"
            Text            =   "0"
            Top             =   650
            Width           =   650
         End
         Begin VB.TextBox txtNroLocales 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   11
            Tag             =   "txtPrincipal"
            Text            =   "0"
            Top             =   650
            Width           =   650
         End
         Begin VB.ComboBox cmbInmuebleClase 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   3000
         End
         Begin VB.ComboBox cmbInmuebleCate 
            Height          =   315
            Left            =   6360
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   3000
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Año de Construcción: "
            Height          =   195
            Left            =   7080
            TabIndex        =   41
            Top             =   660
            Width           =   1575
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "N° de Sótanos : "
            Height          =   195
            Left            =   5070
            TabIndex        =   40
            Top             =   660
            Width           =   1170
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Categoría :"
            Height          =   195
            Left            =   5400
            TabIndex        =   39
            Top             =   285
            Width           =   795
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Clase de Inmueble :"
            Height          =   195
            Left            =   150
            TabIndex        =   38
            Top             =   285
            Width           =   1395
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "N° de Pisos : "
            Height          =   195
            Left            =   2640
            TabIndex        =   37
            Top             =   640
            Width           =   960
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "N° de Locales :"
            Height          =   195
            Left            =   150
            TabIndex        =   36
            Top             =   660
            Width           =   1095
         End
      End
      Begin VB.Frame fraValor 
         Caption         =   "Valor"
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
         Height          =   1600
         Left            =   6240
         TabIndex        =   29
         Top             =   360
         Width           =   3375
         Begin VB.TextBox txtVRM 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   1230
            Width           =   1560
         End
         Begin VB.TextBox txtValorEdificacion 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   920
            Width           =   1560
         End
         Begin VB.TextBox txtValorComercial 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   600
            Width           =   1560
         End
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "V.R.M :"
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   1275
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Edificación y Obras :"
            Height          =   195
            Left            =   150
            TabIndex        =   32
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   150
            TabIndex        =   31
            Top             =   290
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Comercial :"
            Height          =   195
            Left            =   150
            TabIndex        =   30
            Top             =   630
            Width           =   1185
         End
      End
      Begin VB.Frame fraUbigeo 
         Caption         =   "Ubigeo"
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
         Height          =   1600
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   6015
         Begin VB.ComboBox cmdLT 
            Height          =   315
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox cmbMZ 
            Height          =   315
            ItemData        =   "frmGarantiaValorTasacInmob.frx":0326
            Left            =   3120
            List            =   "frmGarantiaValorTasacInmob.frx":0328
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox cmbUbicacionGeografica 
            Height          =   315
            Index           =   4
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "Zona"
            Top             =   680
            Width           =   2100
         End
         Begin VB.ComboBox cmbUbicacionGeografica 
            Height          =   315
            Index           =   3
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Tag             =   "Distrito"
            Top             =   680
            Width           =   2100
         End
         Begin VB.ComboBox cmbUbicacionGeografica 
            Height          =   315
            Index           =   2
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Tag             =   "Provincia"
            Top             =   240
            Width           =   2100
         End
         Begin VB.TextBox txtDireccion 
            Height          =   285
            Left            =   960
            MaxLength       =   255
            TabIndex        =   4
            Top             =   1080
            Width           =   1700
         End
         Begin VB.ComboBox cmbUbicacionGeografica 
            Height          =   315
            Index           =   1
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Tag             =   "Departamento"
            Top             =   240
            Width           =   2100
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "ET: "
            Height          =   195
            Left            =   4800
            TabIndex        =   53
            Top             =   1125
            Width           =   300
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "LT: "
            Height          =   195
            Left            =   3840
            TabIndex        =   49
            Top             =   1125
            Width           =   285
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Mz : "
            Height          =   195
            Left            =   2760
            TabIndex        =   48
            Top             =   1125
            Width           =   345
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   1125
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   150
            TabIndex        =   28
            Top             =   725
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   3120
            TabIndex        =   27
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Dpto :"
            Height          =   195
            Left            =   150
            TabIndex        =   26
            Top             =   290
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Prov :"
            Height          =   195
            Left            =   3120
            TabIndex        =   25
            Top             =   285
            Width           =   420
         End
      End
   End
End
Attribute VB_Name = "frmGarantiaValorTasacInmob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************************
'** Nombre : frmGarantiaValorTasacInmob
'** Descripción : Para registro/edición/consulta de Valorización Inmobiliaria creado segun TI-ERS063-2014
'** Creación : EJVG, 20150204 09:36:01 AM
'*******************************************************************************************************
Option Explicit

Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fbPrimero As Boolean
Dim fbOk As Boolean

Dim fnMoneda As Moneda
Dim fvValorTasacionInmobiliaria As tValorTasacionInmobiliaria
Dim fvValorTasacionInmobiliaria_ULT As tValorTasacionInmobiliaria
Dim fnTpoBienContrato As Integer 'CTI5 ERS0012020
Dim fsEtapa As String 'CTI5 ERS0012020
Dim fsTpoDoc As String 'CTI5 ERS0012020

'CTI5 ERS0012020****************************
Private Sub cmbMZ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdLT
    End If
End Sub

Private Sub cmbMZ_LostFocus()
  EnfocaControl cmdLT
End Sub

Private Sub cmdLT_LostFocus()
    If cmbMoneda.Enabled And cmbMoneda.Visible Then
        EnfocaControl cmbMoneda
    Else
        EnfocaControl txtValorComercial
    End If
End Sub
Private Sub cmdLT_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If cmbMoneda.Enabled And cmbMoneda.Visible Then
        EnfocaControl cmbMoneda
    Else
        EnfocaControl txtValorComercial
    End If
  End If
End Sub
Private Sub cmdET_LostFocus()
    EnfocaControl cmbMoneda
End Sub
Private Sub cmdET_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        EnfocaControl cmbMoneda
 End If
End Sub
'*******************************************

Private Sub cmbUbicacionGeografica_Click(Index As Integer)
    Dim oUbic As COMDPersona.DCOMPersonas
    Dim rs As ADODB.Recordset
    Dim i As Integer

    If Index <> 4 Then
        Set oUbic = New COMDPersona.DCOMPersonas
        Set rs = oUbic.CargarUbicacionesGeograficas(True, Index + 1, Trim(Right(cmbUbicacionGeografica(Index).Text, 15)))

        For i = Index + 1 To cmbUbicacionGeografica.count
            cmbUbicacionGeografica(i).Clear
        Next
        
        While Not rs.EOF
            cmbUbicacionGeografica(Index + 1).AddItem Trim(rs!cUbiGeoDescripcion) & Space(50) & Trim(rs!cUbiGeoCod)
            rs.MoveNext
        Wend
        Set oUbic = Nothing
    End If
End Sub
Private Sub cmbUbicacionGeografica_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 4 Then
            EnfocaControl Me.cmbUbicacionGeografica(Index + 1)
        Else
            EnfocaControl txtDireccion
        End If
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsCtaCod As String
    Dim lsValida As String
    
    On Error GoTo ErrAceptar
    If Not validarDatos Then Exit Sub
    
    fvValorTasacionInmobiliaria.sUbicacionGeografica = Trim(Right(cmbUbicacionGeografica(4).Text, 15))
    fvValorTasacionInmobiliaria.sDireccion = Trim(txtDireccion.Text)
    fnMoneda = CInt(Right(cmbMoneda, 2))
    fvValorTasacionInmobiliaria.nValorComercial = CCur(txtValorComercial.Text)
    fvValorTasacionInmobiliaria.nValorEdificacion = CCur(txtValorEdificacion.Text)
    fvValorTasacionInmobiliaria.nVRM = CCur(txtVRM.Text)
    fvValorTasacionInmobiliaria.nCategoria = CInt(Trim(Right(cmbInmuebleCate.Text, 2)))
    fvValorTasacionInmobiliaria.nClase = CInt(Trim(Right(cmbInmuebleClase.Text, 2)))
    fvValorTasacionInmobiliaria.nNroPisos = CInt(txtNroPisos.Text)
    fvValorTasacionInmobiliaria.nNroLocales = CInt(txtNroLocales.Text)
    fvValorTasacionInmobiliaria.nNroSotanos = CInt(txtNroSotanos.Text)
    fvValorTasacionInmobiliaria.nAnioConstruccion = CInt(txtAnioConstruccion.Text)
    fvValorTasacionInmobiliaria.dTasacion = CDate(txtTasacionFecha.Text)
    fvValorTasacionInmobiliaria.nTasacionTC = CCur(txtTasacionTC.Text)
    fvValorTasacionInmobiliaria.sTasadorCod = txtTasadorCod.psCodigoPersona
    fvValorTasacionInmobiliaria.sTasadorNombre = Trim(txtTasadorNombre.Text)
    fvValorTasacionInmobiliaria.sTasadorDNI = Trim(txtTasadorDNI.Text)
    fvValorTasacionInmobiliaria.sTasadorREPEV = Trim(txtTasadorREPEV.Text)
    
    'CTI5 ERS012020*****************************************************
    fvValorTasacionInmobiliaria.sMz = Trim(Right(cmbMZ.Text, 3))
    fvValorTasacionInmobiliaria.sLt = Trim(Right(cmdLT.Text, 3))
    fvValorTasacionInmobiliaria.Set = Trim(Right(cmdET.Text, 3))
    '*******************************************************************

    fbOk = True
    Unload Me
    Exit Sub
ErrAceptar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub



Private Sub Form_Load()
    fbOk = False
    Screen.MousePointer = 11
    
    cargarControles
    LimpiarControles

    If fbEditar Or fbConsultar Then
        cmbUbicacionGeografica(1).ListIndex = IndiceListaCombo(cmbUbicacionGeografica(1), Space(30) & "1" & Mid(fvValorTasacionInmobiliaria.sUbicacionGeografica, 2, 2) & String(9, "0"))
        cmbUbicacionGeografica(2).ListIndex = IndiceListaCombo(cmbUbicacionGeografica(2), Space(30) & "2" & Mid(fvValorTasacionInmobiliaria.sUbicacionGeografica, 2, 4) & String(7, "0"))
        cmbUbicacionGeografica(3).ListIndex = IndiceListaCombo(cmbUbicacionGeografica(3), Space(30) & "3" & Mid(fvValorTasacionInmobiliaria.sUbicacionGeografica, 2, 6) & String(5, "0"))
        cmbUbicacionGeografica(4).ListIndex = IndiceListaCombo(cmbUbicacionGeografica(4), Space(30) & fvValorTasacionInmobiliaria.sUbicacionGeografica)
        txtDireccion.Text = fvValorTasacionInmobiliaria.sDireccion
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
        txtValorComercial.Text = Format(fvValorTasacionInmobiliaria.nValorComercial, "#,##0.00")
        txtValorEdificacion.Text = Format(fvValorTasacionInmobiliaria.nValorEdificacion, "#,##0.00")
        txtVRM.Text = Format(fvValorTasacionInmobiliaria.nVRM, "#,##0.00")
        cmbInmuebleCate.ListIndex = IndiceListaCombo(cmbInmuebleCate, fvValorTasacionInmobiliaria.nCategoria)
        cmbInmuebleClase.ListIndex = IndiceListaCombo(cmbInmuebleClase, fvValorTasacionInmobiliaria.nClase)
        txtNroLocales.Text = fvValorTasacionInmobiliaria.nNroLocales
        txtNroPisos.Text = fvValorTasacionInmobiliaria.nNroPisos
        txtNroSotanos.Text = fvValorTasacionInmobiliaria.nNroSotanos
        txtAnioConstruccion.Text = fvValorTasacionInmobiliaria.nAnioConstruccion
        txtTasacionFecha.Text = Format(fvValorTasacionInmobiliaria.dTasacion, gsFormatoFechaView)
        txtTasacionTC.Text = Format(fvValorTasacionInmobiliaria.nTasacionTC, gsFormatoNumeroView)
        txtTasadorCod.Text = fvValorTasacionInmobiliaria.sTasadorCod
        txtTasadorCod.psCodigoPersona = fvValorTasacionInmobiliaria.sTasadorCod
        txtTasadorNombre.Text = fvValorTasacionInmobiliaria.sTasadorNombre
        txtTasadorDNI.Text = fvValorTasacionInmobiliaria.sTasadorDNI
        txtTasadorREPEV.Text = fvValorTasacionInmobiliaria.sTasadorREPEV
        
        'CTI5 ERS012020*****************************************************
        If fnTpoBienContrato = gTpoBienFuturo Then
        fvValorTasacionInmobiliaria.Set = fsEtapa
            If fbEditar Then
                If fsTpoDoc = 1 Then
                    cmdET.Enabled = True
                Else
                    cmdET.Enabled = False
                End If
            End If
        End If
        cmbMZ.ListIndex = IndiceListaCombo(cmbMZ, fvValorTasacionInmobiliaria.sMz)
        cmdLT.ListIndex = IndiceListaCombo(cmdLT, fvValorTasacionInmobiliaria.sLt)
        cmdET.ListIndex = IndiceListaCombo(cmdET, fvValorTasacionInmobiliaria.Set)
        '*******************************************************************

        
        If fbConsultar Then
            fraUbigeo.Enabled = False
            fraValor.Enabled = False
            fraDescripcion.Enabled = False
            fraTasador.Enabled = False
            CmdAceptar.Enabled = False
        End If
    End If
    
    If fbPrimero Then
        fnMoneda = gMonedaNacional
    End If
    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
    cmbMoneda.Enabled = False
    
    If fbRegistrar Then
        cmbUbicacionGeografica(1).ListIndex = IndiceListaCombo(cmbUbicacionGeografica(1), Space(30) & "1" & Mid(fvValorTasacionInmobiliaria_ULT.sUbicacionGeografica, 2, 2) & String(9, "0"))
        cmbUbicacionGeografica(2).ListIndex = IndiceListaCombo(cmbUbicacionGeografica(2), Space(30) & "2" & Mid(fvValorTasacionInmobiliaria_ULT.sUbicacionGeografica, 2, 4) & String(7, "0"))
        cmbUbicacionGeografica(3).ListIndex = IndiceListaCombo(cmbUbicacionGeografica(3), Space(30) & "3" & Mid(fvValorTasacionInmobiliaria_ULT.sUbicacionGeografica, 2, 6) & String(5, "0"))
        cmbUbicacionGeografica(4).ListIndex = IndiceListaCombo(cmbUbicacionGeografica(4), Space(30) & fvValorTasacionInmobiliaria_ULT.sUbicacionGeografica)
        txtDireccion.Text = fvValorTasacionInmobiliaria_ULT.sDireccion
        'cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
        txtValorComercial.Text = Format(fvValorTasacionInmobiliaria_ULT.nValorComercial, "#,##0.00")
        txtValorEdificacion.Text = Format(fvValorTasacionInmobiliaria_ULT.nValorEdificacion, "#,##0.00")
        txtVRM.Text = Format(fvValorTasacionInmobiliaria_ULT.nVRM, "#,##0.00")
        cmbInmuebleCate.ListIndex = IndiceListaCombo(cmbInmuebleCate, fvValorTasacionInmobiliaria_ULT.nCategoria)
        cmbInmuebleClase.ListIndex = IndiceListaCombo(cmbInmuebleClase, fvValorTasacionInmobiliaria_ULT.nClase)
        txtNroLocales.Text = IIf(fvValorTasacionInmobiliaria_ULT.nNroLocales = 0, 1, fvValorTasacionInmobiliaria_ULT.nNroLocales)
        txtNroPisos.Text = IIf(fvValorTasacionInmobiliaria_ULT.nNroPisos = 0, 1, fvValorTasacionInmobiliaria_ULT.nNroPisos)
        txtNroSotanos.Text = fvValorTasacionInmobiliaria_ULT.nNroSotanos
        txtAnioConstruccion.Text = IIf(fvValorTasacionInmobiliaria_ULT.nAnioConstruccion = 0, Year(gdFecSis), fvValorTasacionInmobiliaria_ULT.nAnioConstruccion)
        txtTasacionFecha.Text = Format(IIf(Year(fvValorTasacionInmobiliaria_ULT.dTasacion) <= 1950, gdFecSis, fvValorTasacionInmobiliaria_ULT.dTasacion), gsFormatoFechaView)
        txtTasacionTC.Text = Format(fvValorTasacionInmobiliaria_ULT.nTasacionTC, gsFormatoNumeroView)
        txtTasadorCod.Text = fvValorTasacionInmobiliaria_ULT.sTasadorCod
        txtTasadorCod.psCodigoPersona = fvValorTasacionInmobiliaria_ULT.sTasadorCod
        txtTasadorNombre.Text = fvValorTasacionInmobiliaria_ULT.sTasadorNombre
        txtTasadorDNI.Text = fvValorTasacionInmobiliaria_ULT.sTasadorDNI
        txtTasadorREPEV.Text = fvValorTasacionInmobiliaria_ULT.sTasadorREPEV
        'CTI5 ERS012020*****************************************************
        If fnTpoBienContrato = gTpoBienFuturo Then
            fvValorTasacionInmobiliaria.Set = fsEtapa
            cmbMZ.ListIndex = IndiceListaCombo(cmbMZ, fvValorTasacionInmobiliaria.sMz)
            cmdLT.ListIndex = IndiceListaCombo(cmdLT, fvValorTasacionInmobiliaria.sLt)
            cmdET.ListIndex = IndiceListaCombo(cmdET, fvValorTasacionInmobiliaria.Set)
            If fsTpoDoc = 1 Then
                cmdET.Enabled = True
            Else
                cmdET.Enabled = False
            End If
        End If
        '*******************************************************************
    End If
    
    If fbRegistrar Then
        Caption = "TASACIÓN INMOBILIARIA [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = "TASACIÓN INMOBILIARIA [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = "TASACIÓN INMOBILIARIA [ EDITAR ]"
    End If
    
    Call CambiaTamañoCombo(cmbUbicacionGeografica(4), 200)
    Screen.MousePointer = 0
End Sub
Public Function Registrar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef pvValorTasacionInmobiliaria As tValorTasacionInmobiliaria, ByRef pvValorTasacionInmobiliaria_ULT As tValorTasacionInmobiliaria, Optional pnTpoBienContrato As Integer = 1, Optional psEtapa As String, Optional psTpoDoc As String) As Boolean
    fbRegistrar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvValorTasacionInmobiliaria = pvValorTasacionInmobiliaria
    fvValorTasacionInmobiliaria_ULT = pvValorTasacionInmobiliaria_ULT
    
    'CTI5 ERS012020
    fsEtapa = psEtapa
    fsTpoDoc = psTpoDoc
    fnTpoBienContrato = pnTpoBienContrato
    If (pnTpoBienContrato = 2) Or (Trim(fvValorTasacionInmobiliaria.sMz) <> "" Or Trim(fvValorTasacionInmobiliaria.Set) <> "") Then
        cmbMZ.Visible = True
        cmdLT.Visible = True
        cmdET.Visible = True
        Label20.Visible = True
        Label22.Visible = True
        Label23.Visible = True
        txtDireccion.Width = 1700
    Else
        cmbMZ.Visible = False
        cmdLT.Visible = False
        cmdET.Visible = False
        Label20.Visible = False
        Label22.Visible = False
        Label23.Visible = False
        txtDireccion.Width = 4815
    End If
    '****************************
        
    Show 1
    pnMoneda = fnMoneda
    pvValorTasacionInmobiliaria = fvValorTasacionInmobiliaria
    
    Registrar = fbOk
End Function
Public Function Editar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef pvValorTasacionInmobiliaria As tValorTasacionInmobiliaria, ByRef pvValorTasacionInmobiliaria_ULT As tValorTasacionInmobiliaria, Optional pnTpoBienContrato As Integer = 1, Optional psEtapa As String, Optional psTpoDoc As String) As Boolean
    fbEditar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fvValorTasacionInmobiliaria = pvValorTasacionInmobiliaria
    fvValorTasacionInmobiliaria_ULT = pvValorTasacionInmobiliaria_ULT
    'CTI5 ERS012020
    fsEtapa = psEtapa
    fsTpoDoc = psTpoDoc
    fnTpoBienContrato = pnTpoBienContrato
    If (pnTpoBienContrato = 2) Or (Trim(fvValorTasacionInmobiliaria.sMz) <> "" Or Trim(fvValorTasacionInmobiliaria.Set) <> "") Then
        cmbMZ.Visible = True
        cmdLT.Visible = True
        cmdET.Visible = True
        Label20.Visible = True
        Label22.Visible = True
        Label23.Visible = True
    Else
        cmbMZ.Visible = False
        cmdLT.Visible = False
        cmdET.Visible = False
        Label20.Visible = False
        Label22.Visible = False
        Label23.Visible = False
    End If
    '****************************
    
    Show 1
    pnMoneda = fnMoneda
    pvValorTasacionInmobiliaria = fvValorTasacionInmobiliaria
    
    Editar = fbOk
End Function
Public Sub Consultar(ByVal pnMoneda As Moneda, ByRef pvValorTasacionInmobiliaria As tValorTasacionInmobiliaria, Optional pnTpoBienContrato As Integer = 1, Optional psEtapa As String, Optional psTpoDoc As String)
    fbConsultar = True
    fnMoneda = pnMoneda
    fvValorTasacionInmobiliaria = pvValorTasacionInmobiliaria
    'CTI5 ERS012020
    fsEtapa = psEtapa
    fsTpoDoc = psTpoDoc
    fnTpoBienContrato = pnTpoBienContrato
    If (pnTpoBienContrato = 2) Or (Trim(fvValorTasacionInmobiliaria.sMz) <> "" Or Trim(fvValorTasacionInmobiliaria.Set) <> "") Then
        cmbMZ.Visible = True
        cmdLT.Visible = True
        cmdET.Visible = True
        Label20.Visible = True
        Label22.Visible = True
        Label23.Visible = True
    Else
        cmbMZ.Visible = False
        cmdLT.Visible = False
        cmdET.Visible = False
        Label20.Visible = False
        Label22.Visible = False
        Label23.Visible = False
    End If
    '****************************
    Show 1
End Sub
Private Sub cargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim oUbic As New COMDPersona.DCOMPersonas
    Dim rs As New ADODB.Recordset
    
    Set rs = oCons.RecuperaConstantes(1011)
    Call Llenar_Combo_con_Recordset(rs, cmbMoneda)
    
    Set rs = oCons.RecuperaConstantes(gGarantiaClaseTasacInmobiliaria)
    Call Llenar_Combo_con_Recordset(rs, cmbInmuebleClase)
    
    Set rs = oCons.RecuperaConstantes(gGarantiaCategoriaTasacInmobiliaria)
    Call Llenar_Combo_con_Recordset(rs, cmbInmuebleCate)
    
    Set rs = oUbic.CargarUbicacionesGeograficas(True, 1, "04028")
    Call Llenar_Combo_con_Recordset_New(rs, cmbUbicacionGeografica(1))
    
    Call llenarManzanas 'CTI3 ERS0012020
    Call llenarLote 'CTI3 ERS0012020
    Call llenarEtapa 'CTI3 ERS0012020
    RSClose rs
    Set oCons = Nothing
    Set oUbic = Nothing
End Sub
'CTI3 ERS0012020
Private Sub llenarEtapa()
    cmdET.AddItem "I" & Space(100) & "I"
    cmdET.AddItem "II" & Space(100) & "II"
    cmdET.AddItem "III" & Space(100) & "III"
    cmdET.AddItem "IV" & Space(100) & "IV"
End Sub
Private Sub llenarLote()
    Dim i As Integer
    i = 0
    For i = 1 To 150
        cmdLT.AddItem Trim(CStr(i)) & Space(100) & Trim(CStr(i))
    Next i
End Sub
Private Sub llenarManzanas()
    cmbMZ.AddItem "A" & Space(100) & "A"
    cmbMZ.AddItem "B" & Space(100) & "B"
    cmbMZ.AddItem "C" & Space(100) & "C"
    cmbMZ.AddItem "D" & Space(100) & "D"
    cmbMZ.AddItem "E" & Space(100) & "E"
    cmbMZ.AddItem "F" & Space(100) & "F"
    cmbMZ.AddItem "G" & Space(100) & "G"
    cmbMZ.AddItem "H" & Space(100) & "H"
    cmbMZ.AddItem "I" & Space(100) & "I"
    cmbMZ.AddItem "J" & Space(100) & "J"
    cmbMZ.AddItem "K" & Space(100) & "K"
    cmbMZ.AddItem "H" & Space(100) & "H"
    cmbMZ.AddItem "I" & Space(100) & "I"
    cmbMZ.AddItem "J" & Space(100) & "J"
    cmbMZ.AddItem "K" & Space(100) & "K"
    'Added by TORE 20210622: Peticion LITO
    cmbMZ.AddItem "L" & Space(100) & "L"
    cmbMZ.AddItem "M" & Space(100) & "M"
    cmbMZ.AddItem "N" & Space(100) & "N"
    cmbMZ.AddItem "O" & Space(100) & "O"
    cmbMZ.AddItem "P" & Space(100) & "P"
    cmbMZ.AddItem "Q" & Space(100) & "Q"
    cmbMZ.AddItem "R" & Space(100) & "R"
    cmbMZ.AddItem "S" & Space(100) & "S"
    cmbMZ.AddItem "T" & Space(100) & "T"
    cmbMZ.AddItem "U" & Space(100) & "U"
    cmbMZ.AddItem "V" & Space(100) & "V"
    cmbMZ.AddItem "X" & Space(100) & "X"
    cmbMZ.AddItem "Y" & Space(100) & "Y"
    cmbMZ.AddItem "Z" & Space(100) & "Z"
    'Added by TORE 20210622
End Sub
'END CTI3
Private Sub cmdCancelar_Click()
    fbOk = False
    Unload Me
End Sub
Private Sub LimpiarControles()
    cmbUbicacionGeografica(1).ListIndex = -1
    cmbUbicacionGeografica(2).ListIndex = -1
    cmbUbicacionGeografica(3).ListIndex = -1
    cmbUbicacionGeografica(4).ListIndex = -1
    txtDireccion.Text = ""
    cmbMoneda.ListIndex = -1
    txtValorComercial.Text = "0.00"
    txtValorEdificacion.Text = "0.00"
    txtVRM.Text = "0.00"
    cmbInmuebleClase.ListIndex = -1
    cmbInmuebleCate.ListIndex = -1
    txtNroLocales.Text = "1"
    txtNroPisos.Text = "1"
    txtNroSotanos.Text = "0"
    txtAnioConstruccion.Text = Year(gdFecSis)
    txtTasacionFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    txtTasacionTC.Text = "0.00"
    txtTasadorCod.Text = ""
    txtTasadorNombre.Text = ""
    txtTasadorDNI.Text = ""
    txtTasadorREPEV.Text = ""
    'CTI5 ERS001-2020******************************
    cmbMZ.ListIndex = -1
    cmdLT.ListIndex = -1
    cmdET.ListIndex = -1
    '**********************************************
End Sub
Private Sub txtAnioConstruccion_LostFocus()
    txtAnioConstruccion.Text = Format(txtAnioConstruccion.Text, "0000")
End Sub
Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        If cmbMZ.Visible Then
            EnfocaControl cmbMZ
        Else
            If cmbMoneda.Enabled And cmbMoneda.Visible Then
                EnfocaControl cmbMoneda
            Else
                EnfocaControl txtValorComercial
            End If
        End If
    End If
End Sub
Private Sub cmbMoneda_Click()
    Dim lnMoneda As Moneda
    Dim lsColor As Long
    lnMoneda = val(Trim(Right(cmbMoneda.Text, 3)))
    If lnMoneda = gMonedaNacional Then
        lsColor = &H80000005
    Else
        lsColor = &HC0FFC0
    End If
    
    txtValorComercial.BackColor = lsColor
    txtValorEdificacion.BackColor = lsColor
    txtVRM.BackColor = lsColor
End Sub
Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtValorComercial
    End If
End Sub
Private Sub txtDireccion_LostFocus()
    txtDireccion.Text = UCase(txtDireccion.Text)
    If cmbMZ.Visible Then
            EnfocaControl cmbMZ
        Else
            If cmbMoneda.Enabled And cmbMoneda.Visible Then
                EnfocaControl cmbMoneda
            Else
                EnfocaControl txtValorComercial
            End If
    End If
End Sub
Private Sub TxtValorEdificacion_LostFocus()
    txtValorEdificacion.Text = Format(txtValorEdificacion.Text, "#,##0.00")
End Sub
Private Sub txtTasadorCod_EmiteDatos()
    Dim oPersona As New COMDPersona.DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim lsDNI As String, lsREPEV As String
    
    Screen.MousePointer = 11
    
    txtTasadorNombre.Text = ""
    txtTasadorDNI.Text = ""
    txtTasadorREPEV.Text = ""
    If txtTasadorCod.Text <> "" Then
        Set rsPersona = oPersona.RecuperaDatosPersonaxGarantia(txtTasadorCod.psCodigoPersona)
        If Not rsPersona.EOF Then
            lsDNI = rsPersona!Dni
            lsREPEV = rsPersona!REPEV
        End If
        txtTasadorNombre.Text = txtTasadorCod.psDescripcion
        txtTasadorDNI.Text = lsDNI
        txtTasadorREPEV.Text = lsREPEV
    End If
    
    RSClose rsPersona
    Set oPersona = Nothing
    
    Screen.MousePointer = 0
End Sub
Private Sub txtTasadorNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtTasadorDNI
    End If
End Sub
Private Sub txtTasadorDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtTasadorREPEV
    End If
End Sub
Private Sub txtTasadorREPEV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl CmdAceptar
    End If
End Sub
Private Sub txtValorComercial_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtValorComercial, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        EnfocaControl txtValorEdificacion
    End If
End Sub
Private Sub TxtValorEdificacion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtValorEdificacion, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        EnfocaControl txtVRM
    End If
End Sub
Private Sub txtValorComercial_LostFocus()
    txtValorComercial.Text = Format(txtValorComercial.Text, "#,##0.00")
End Sub
Private Sub TxtVRM_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVRM, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        EnfocaControl cmbInmuebleClase
    End If
End Sub
Private Sub cmbInmuebleClase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmbInmuebleCate
    End If
End Sub
Private Sub cmbInmuebleCate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNroLocales
    End If
End Sub
Private Sub txtNroLocales_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtNroPisos
    End If
End Sub
Private Sub txtNroPisos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtNroSotanos
    End If
End Sub
Private Sub txtNroSotanos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtAnioConstruccion
    End If
End Sub
Private Sub txtAnioConstruccion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtTasacionFecha
    End If
End Sub
Private Sub txtTasacionFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtTasacionTC
    End If
End Sub
Private Sub txtTasacionTC_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTasacionTC, KeyAscii, 6, 4)
    If KeyAscii = 13 Then
        EnfocaControl txtTasadorCod
    End If
End Sub
Private Sub txtTasacionTC_LostFocus()
    txtTasacionTC.Text = Format(txtTasacionTC.Text, gsFormatoNumeroView)
End Sub
Private Sub txtTasadorCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl CmdAceptar
    End If
End Sub
Private Sub TxtVRM_LostFocus()
    txtVRM.Text = Format(txtVRM.Text, "#,##0.00")
End Sub
Private Function validarDatos() As Boolean
    Dim i As Integer
    Dim lsFecha As String
        
    For i = 1 To cmbUbicacionGeografica.count
        If cmbUbicacionGeografica(i).ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar el dato [" & Me.cmbUbicacionGeografica(i).Tag & "] de la Ubicación Geográfica", vbInformation, "Aviso"
            EnfocaControl cmbUbicacionGeografica(i)
            Exit Function
        End If
    Next
    If Len(Trim(txtDireccion.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Dirección", vbInformation, "Aviso"
        EnfocaControl txtDireccion
        Exit Function
    End If
    'CTI5 ERS012020******************************
   
    If fnTpoBienContrato = gTpoBienFuturo Then
        If cmbMZ.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar la manzana", vbInformation, "Aviso"
            EnfocaControl cmbMZ
            Exit Function
        End If
        If cmdLT.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar el lote", vbInformation, "Aviso"
            EnfocaControl cmdLT
            Exit Function
        End If
        If cmdET.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar la etapa", vbInformation, "Aviso"
            EnfocaControl cmdET
            Exit Function
        End If
    End If
    '********************************************
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la moneda", vbInformation, "Aviso"
        EnfocaControl cmbMoneda
        Exit Function
    End If
    If Not IsNumeric(txtValorComercial.Text) Then
        MsgBox "Ud. debe de especificar el Valor Comercial", vbInformation, "Aviso"
        EnfocaControl txtValorComercial
        Exit Function
    Else
        If CCur(txtValorComercial.Text) <= 0 Then
            MsgBox "El Valor Comercial debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtValorComercial
            Exit Function
        End If
    End If
    If Not IsNumeric(txtValorEdificacion.Text) Then
        MsgBox "Ud. debe de especificar el Valor de Edificación y Obras", vbInformation, "Aviso"
        EnfocaControl txtValorEdificacion
        Exit Function
    End If
    If Not IsNumeric(txtVRM.Text) Then
        MsgBox "Ud. debe de especificar el Valor de Realización", vbInformation, "Aviso"
        EnfocaControl txtVRM
        Exit Function
    Else
        If CCur(txtVRM.Text) <= 0 Then
            MsgBox "El Valor de Realización debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtVRM
            Exit Function
        End If
        If CCur(txtVRM.Text) > CCur(txtValorComercial.Text) Then
            MsgBox "El Valor de Realización no puede ser mayor al Valor Comercial", vbInformation, "Aviso"
            EnfocaControl txtVRM
            Exit Function
        End If
    End If
    If cmbInmuebleClase.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Clase del Inmueble", vbInformation, "Aviso"
        EnfocaControl cmbInmuebleClase
        Exit Function
    End If
    If cmbInmuebleCate.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Categoría del Inmueble", vbInformation, "Aviso"
        EnfocaControl cmbInmuebleCate
        Exit Function
    End If
    If Not IsNumeric(txtNroLocales.Text) Then
        MsgBox "Ud. debe de especificar el N° de Locales", vbInformation, "Aviso"
        EnfocaControl txtNroLocales
        Exit Function
    Else
        If CInt(txtNroLocales.Text) <= 0 Then
            MsgBox "El N° de Locales debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtNroLocales
            Exit Function
        End If
    End If
    If Not IsNumeric(txtNroPisos.Text) Then
        MsgBox "Ud. debe de especificar el N° de Pisos", vbInformation, "Aviso"
        EnfocaControl txtNroPisos
        Exit Function
    Else
        If CInt(txtNroPisos.Text) <= 0 Then
            MsgBox "El N° de Pisos debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtNroPisos
            Exit Function
        End If
    End If
    If Not IsNumeric(txtNroSotanos.Text) Then 'Nro de Sotanos puede ser cero
        MsgBox "Ud. debe de especificar el N° de Sotanos", vbInformation, "Aviso"
        EnfocaControl txtNroSotanos
        Exit Function
    End If
    If Not IsNumeric(txtAnioConstruccion.Text) Then
        MsgBox "Ud. debe de especificar el Año de Construcción", vbInformation, "Aviso"
        EnfocaControl txtAnioConstruccion
        Exit Function
    Else
        If CInt(txtAnioConstruccion.Text) <= 1890 Then
            MsgBox "Ud. debe de especificar el Año de Construcción", vbInformation, "Aviso"
            EnfocaControl txtAnioConstruccion
            Exit Function
        End If
        If CInt(txtAnioConstruccion.Text) > Year(gdFecSis) Then
            MsgBox "El Año de Construcción no debe ser mayor al año del sistema", vbInformation, "Aviso"
            EnfocaControl txtAnioConstruccion
            Exit Function
        End If
    End If
    lsFecha = ValidaFecha(txtTasacionFecha.Text)
    If Len(Trim(lsFecha)) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        EnfocaControl txtTasacionFecha
        Exit Function
    Else
        If Not fvValorTasacionInmobiliaria_ULT.bMigrado Then
            If CDate(txtTasacionFecha.Text) <= fvValorTasacionInmobiliaria_ULT.dTasacion Then
                MsgBox "La actual fecha de Tasación no puede ser menor o igual a la última Tasación: " & Format(fvValorTasacionInmobiliaria_ULT.dTasacion, gsFormatoFechaView), vbInformation, "Aviso"
                EnfocaControl txtTasacionFecha
                Exit Function
            End If
        End If
        If CDate(txtTasacionFecha.Text) > gdFecSis Then
            MsgBox "La fecha de Tasación no puede ser mayor a la fecha de Sistema", vbInformation, "Aviso"
            EnfocaControl txtTasacionFecha
            Exit Function
        End If
    End If
    If Not IsNumeric(txtTasacionTC.Text) Then
        MsgBox "Ud. debe de especificar el Tipo de Cambio de Tasación", vbInformation, "Aviso"
        EnfocaControl txtTasacionTC
        Exit Function
    Else
        If CCur(txtTasacionTC.Text) <= 0 Then
            MsgBox "El Tipo de Cambio de Tasación debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtTasacionTC
            Exit Function
        End If
        If CCur(txtTasacionTC.Text) > 10 Then
            MsgBox "Verifique el Tipo de Cambio de Tasación", vbInformation, "Aviso"
            EnfocaControl txtTasacionTC
            Exit Function
        End If
    End If
    If Len(Trim(txtTasadorCod.Text)) <> 13 Or Len(txtTasadorCod.psCodigoPersona) <> 13 Then
        MsgBox "Ud. debe especificar al Tasador", vbInformation, "Aviso"
        EnfocaControl txtTasadorCod
        Exit Function
    Else
        If txtTasadorCod.PersPersoneria <> gPersonaNat Then
            MsgBox "El Tasador debe ser una persona Natural", vbInformation, "Aviso"
            EnfocaControl txtTasadorCod
            Exit Function
        Else
            If Len(Trim(txtTasadorDNI.Text)) <> 8 Then
                MsgBox "El Tasador no cuenta con Documento DNI", vbInformation, "Aviso"
                EnfocaControl txtTasadorCod
                Exit Function
            End If
            If Len(Trim(txtTasadorREPEV.Text)) = 0 Then
                MsgBox "El Tasador no cuenta con Documento REPEV", vbInformation, "Aviso"
                EnfocaControl txtTasadorCod
                Exit Function
            End If
        End If
    End If
        
    validarDatos = True
End Function
