VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredBPPConfJefesAgeYT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Configuración de Jefes de Agencia y Jefes Territoriales"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   Icon            =   "frmCredBPPConfJefesAgeYT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Jefes de Agencia"
      TabPicture(0)   =   "frmCredBPPConfJefesAgeYT.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTopes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Jefes Territoriales"
      TabPicture(1)   =   "frmCredBPPConfJefesAgeYT.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraTopes 
         Caption         =   "Configuración del Jefe de Agencia"
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
         Height          =   3495
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   8175
         Begin VB.ComboBox cmbAgenciasCJA 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdMostrarCJA 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   6960
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbMesesCJA 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdGuardarCJA 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   6840
            TabIndex        =   8
            Top             =   2880
            Width           =   1095
         End
         Begin Spinner.uSpinner uspAnioCJA 
            Height          =   315
            Left            =   2880
            TabIndex        =   12
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Max             =   9999
            Min             =   1900
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin SICMACT.FlexEdit feJefeAgencia 
            Height          =   1815
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   7680
            _ExtentX        =   13547
            _ExtentY        =   3201
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "#-Usuario-Nombre-Estado-Aux"
            EncabezadosAnchos=   "0-1200-5000-1000-0"
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
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-4-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-C-L"
            FormatosEdit    =   "0-0-0-0-0"
            CantEntero      =   15
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
            Height          =   195
            Left            =   3960
            TabIndex        =   15
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Mes - Año :"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   420
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Configuración del Jefe Territorial"
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
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8295
         Begin VB.CommandButton cmdGuardarJT 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   6840
            TabIndex        =   4
            Top             =   2760
            Width           =   1095
         End
         Begin VB.ComboBox cmbMesesJT 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdMostrarJT 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   4560
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
         Begin Spinner.uSpinner uspAnioJT 
            Height          =   315
            Left            =   3480
            TabIndex        =   5
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Max             =   9999
            Min             =   1900
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin SICMACT.FlexEdit feJefeTerritorial 
            Height          =   1815
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3201
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "#-Zona-Usuario-Nombre-Aux"
            EncabezadosAnchos=   "0-1200-1000-5000-0"
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
            ColumnasAEditar =   "X-X-2-3-X"
            ListaControles  =   "0-0-3-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-C"
            FormatosEdit    =   "0-0-0-0-0"
            CantEntero      =   15
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Mes - Año :"
            Height          =   195
            Left            =   720
            TabIndex        =   6
            Top             =   420
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPConfJefesAgeYT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private i As Integer
'
'Private Sub cmbAgenciasCJA_Click()
'LimpiaDatos
'End Sub
'
'
'Private Sub cmbMesesCJA_Click()
'LimpiaDatos
'End Sub
'
'
'Private Sub cmbMesesJT_Click()
'LimpiaDatos 1
'End Sub
'
'Private Sub cmdGuardarCJA_Click()
'On Error GoTo Error
'If ValidaDatos(1) Then
'    If MsgBox("Esta seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim oBPP As COMNCredito.NCOMBPPR
'        Dim lsAgeCod As String
'        Dim lnMes As Integer
'        Dim lnAnio As Integer
'        Dim lsFecha As String
'        Dim lnEstado As Integer
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        lsAgeCod = Trim(Right(cmbAgenciasCJA.Text, 4))
'        lnMes = CInt(Trim(Right(cmbMesesCJA.Text, 4)))
'        lnAnio = CInt(uspAnioCJA.valor)
'
'        For i = 1 To feJefeAgencia.Rows - 1
'            lnEstado = IIf(Trim(feJefeAgencia.TextMatrix(i, 3)) = ".", 1, 0)
'            Call oBPP.OpeJefAgenciaAsig(lsAgeCod, lnMes, lnAnio, Trim(feJefeAgencia.TextMatrix(i, 4)), lnEstado, gsCodUser, lsFecha)
'        Next i
'
'        Set oBPP = Nothing
'        cmdMostrarCJA.Enabled = True
'        cmdGuardarCJA.Enabled = False
'        MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'    End If
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarJT_Click()
'On Error GoTo Error
'If ValidaDatos(1, 1) Then
'    If MsgBox("Esta seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim oBPP As COMNCredito.NCOMBPPR
'        Dim lnMes As Integer
'        Dim lnAnio As Integer
'        Dim lsFecha As String
'        Dim lnZona As Integer
'        Dim lsPersCod As String
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        lnMes = CInt(Trim(Right(cmbMesesJT.Text, 4)))
'        lnAnio = CInt(uspAnioJT.valor)
'
'        For i = 1 To feJefeTerritorial.Rows - 1
'            lsPersCod = feJefeTerritorial.TextMatrix(i, 2)
'            lsPersCod = Trim(Right(Mid(lsPersCod, 5, Len(lsPersCod)), 215))
'            lsPersCod = Trim(Left(lsPersCod, 13))
'            lnZona = CInt(Trim(feJefeTerritorial.TextMatrix(i, 4)))
'
'            Call oBPP.OpeJefTerritorialAsig(lnZona, lnMes, lnAnio, lsPersCod, gsCodUser, lsFecha)
'        Next i
'
'        Set oBPP = Nothing
'        cmdMostrarJT.Enabled = True
'        cmdGuardarJT.Enabled = False
'        MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'    End If
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdMostrarCJA_Click()
'If ValidaDatos Then
'    Dim lsAgeCod As String
'    Dim lnMes As Integer
'    Dim lnAnio As Integer
'
'    lsAgeCod = Trim(Right(cmbAgenciasCJA.Text, 4))
'    lnMes = CInt(Trim(Right(cmbMesesCJA.Text, 4)))
'    lnAnio = CInt(uspAnioCJA.valor)
'
'    Call CargaDatosJA(lnMes, lnAnio, lsAgeCod)
'
'    cmdGuardarCJA.Enabled = True
'    cmdMostrarCJA.Enabled = False
'End If
'End Sub
'
'Private Sub cmdMostrarJT_Click()
'If ValidaDatos(0, 1) Then
'    Dim lnMes As Integer
'    Dim lnAnio As Integer
'
'    lnMes = CInt(Trim(Right(cmbMesesJT.Text, 4)))
'    lnAnio = CInt(uspAnioJT.valor)
'
'    Call CargaDatosJT(lnMes, lnAnio)
'
'    cmdGuardarJT.Enabled = True
'    cmdMostrarJT.Enabled = False
'End If
'End Sub
'
'Private Sub feJefeTerritorial_OnChangeCombo()
'Dim lsNombre As String
'
'lsNombre = feJefeTerritorial.TextMatrix(feJefeTerritorial.row, 2)
'lsNombre = Trim(Right(Mid(lsNombre, 5, Len(lsNombre)), 215))
'lsNombre = Mid(lsNombre, 15, Len(lsNombre))
'feJefeTerritorial.TextMatrix(feJefeTerritorial.row, 3) = lsNombre
'End Sub
'
'Private Sub Form_Load()
'    uspAnioCJA.valor = Year(gdFecSis)
'    uspAnioJT.valor = Year(gdFecSis)
'    CargaComboMeses cmbMesesCJA
'    CargaComboMeses cmbMesesJT
'    CargaComboAgencias cmbAgenciasCJA
'    cmdGuardarCJA.Enabled = False
'    cmdGuardarJT.Enabled = False
'End Sub
'
'Private Sub LimpiaDatos(Optional pnPestana As Integer = 0)
'If pnPestana = 0 Then
'    LimpiaFlex feJefeAgencia
'    cmdGuardarCJA.Enabled = False
'    cmdMostrarCJA.Enabled = True
'Else
'    LimpiaFlex feJefeTerritorial
'    cmdGuardarJT.Enabled = False
'    cmdMostrarJT.Enabled = True
'End If
'End Sub
'
'Private Sub uspAnioCJA_Change()
'LimpiaDatos
'End Sub
'
'Private Sub uspAnioJT_Change()
'LimpiaDatos 1
'End Sub
'
'Private Function ValidaDatos(Optional pnTipo As Integer = 0, Optional ByVal pnPestana As Integer = 0) As Boolean
'Dim Contador As Integer
'ValidaDatos = True
'If pnPestana = 0 Then
'    If Trim(cmbMesesCJA.Text) = "" Then
'        MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'        ValidaDatos = False
'        cmbMesesCJA.SetFocus
'        Exit Function
'    End If
'
'    If Trim(cmbAgenciasCJA.Text) = "" Then
'        MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'        ValidaDatos = False
'        cmbAgenciasCJA.SetFocus
'        Exit Function
'    End If
'
'    If pnTipo = 1 Then
'        Contador = 0
'        For i = 1 To feJefeAgencia.Rows - 1
'            If Trim(feJefeAgencia.TextMatrix(i, 3)) = "." Then
'                Contador = Contador + 1
'            End If
'        Next i
'
'        If Contador = 0 Then
'            MsgBox "Favor de Selecionar un Jefe de Agencia", vbInformation, "Aviso"
'            ValidaDatos = False
'            feJefeAgencia.SetFocus
'            Exit Function
'        End If
'
'        If Contador > 1 Then
'            MsgBox "Favor de Selecionar solo un Jefe de Agencia", vbInformation, "Aviso"
'            ValidaDatos = False
'            feJefeAgencia.SetFocus
'            Exit Function
'        End If
'    End If
'Else
'    If Trim(cmbMesesJT.Text) = "" Then
'        MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'        ValidaDatos = False
'        cmbMesesJT.SetFocus
'        Exit Function
'    End If
'
'
'    If pnTipo = 1 Then
'        For i = 1 To feJefeTerritorial.Rows - 1
'            If Trim(feJefeTerritorial.TextMatrix(i, 2)) = "-" Then
'                MsgBox "Favor de Seleccionar un Jefe Territorial para la " & Trim(feJefeTerritorial.TextMatrix(i, 1)), vbInformation, "Aviso"
'                ValidaDatos = False
'                feJefeTerritorial.SetFocus
'                Exit Function
'            End If
'        Next i
'    End If
'End If
'
'End Function
'
'Private Sub CargaDatosJA(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psAgeCod As String)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.ObtenerJefAgenciaAct(pnMes, pnAnio, psAgeCod)
'LimpiaFlex feJefeAgencia
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    For i = 1 To rsBPP.RecordCount
'        feJefeAgencia.AdicionaFila
'        feJefeAgencia.TextMatrix(i, 1) = rsBPP!Usuario
'        feJefeAgencia.TextMatrix(i, 2) = rsBPP!Nombre
'        feJefeAgencia.TextMatrix(i, 3) = IIf(Trim(rsBPP!nEstado) = "0", "", "1")
'        feJefeAgencia.TextMatrix(i, 4) = Trim(rsBPP!cPersCod)
'        rsBPP.MoveNext
'    Next i
'End If
'
'Set oBPP = Nothing
'Set rsBPP = Nothing
'End Sub
'Private Sub CargaDatosJT(ByVal pnMes As Integer, ByVal pnAnio As Integer)
'
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'Set oBPP = New COMNCredito.NCOMBPPR
'
'LimpiaFlex feJefeTerritorial
''Combito
'Set rsBPP = oBPP.ObtenerUsuariosJefTerritorialAct(pnMes, pnAnio)
'feJefeTerritorial.CargaCombo rsBPP
'
''Datos
'Set rsBPP = Nothing
'Set rsBPP = oBPP.ObtenerJefTerritorialAct(pnMes, pnAnio)
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    For i = 1 To rsBPP.RecordCount
'        feJefeTerritorial.AdicionaFila
'        feJefeTerritorial.TextMatrix(i, 1) = Trim(rsBPP!Zona)
'        feJefeTerritorial.TextMatrix(i, 2) = Trim(rsBPP!cUser) & Space(75) & Trim(rsBPP!cPersCod) & "-" & Trim(rsBPP!Nombre)
'        feJefeTerritorial.TextMatrix(i, 3) = Trim(rsBPP!Nombre)
'        feJefeTerritorial.TextMatrix(i, 4) = Trim(rsBPP!CodZona)
'        rsBPP.MoveNext
'    Next i
'End If
'
'Set oBPP = Nothing
'Set rsBPP = Nothing
'End Sub
