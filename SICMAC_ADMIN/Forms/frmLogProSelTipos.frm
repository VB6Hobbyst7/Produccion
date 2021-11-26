VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogProSelTipos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipos de Proceso de Seleccion"
   ClientHeight    =   5760
   ClientLeft      =   1875
   ClientTop       =   2235
   ClientWidth     =   8250
   Icon            =   "frmLogProSelTipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8250
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   5040
      Width           =   1275
   End
   Begin TabDlg.SSTab sstReg 
      Height          =   3555
      Left            =   120
      TabIndex        =   4
      Top             =   2100
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   6271
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   670
      TabCaption(0)   =   "Limites de Procesos de Selección    "
      TabPicture(0)   =   "frmLogProSelTipos.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrRegLimites"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frListLimites"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Etapas de Procesos de Selección     "
      TabPicture(1)   =   "frmLogProSelTipos.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraEtaReg"
      Tab(1).Control(1)=   "fraEtaVis"
      Tab(1).ControlCount=   2
      Begin VB.Frame frListLimites 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   80
         TabIndex        =   5
         Top             =   400
         Width           =   7815
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1440
            TabIndex        =   37
            Top             =   2540
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   2540
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDet 
            Height          =   2175
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   3836
            _Version        =   393216
            Cols            =   12
            FixedCols       =   0
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483633
            WordWrap        =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   2
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   12
         End
      End
      Begin VB.Frame FrRegLimites 
         Caption         =   "Registro de Limites de Procesos de Seleccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   80
         TabIndex        =   29
         Top             =   400
         Visible         =   0   'False
         Width           =   7815
         Begin VB.TextBox txtconsucode 
            Height          =   315
            Left            =   6060
            TabIndex        =   56
            Top             =   360
            Width           =   1470
         End
         Begin VB.TextBox txtAbreviatura 
            Height          =   315
            Left            =   3060
            TabIndex        =   54
            Top             =   360
            Width           =   1470
         End
         Begin VB.Frame Frame3 
            Caption         =   "Obras"
            Height          =   1215
            Left            =   5280
            TabIndex        =   41
            Top             =   1200
            Width           =   2415
            Begin VB.TextBox txtObrasMax 
               Height          =   315
               Left            =   960
               TabIndex        =   51
               Top             =   720
               Width           =   1275
            End
            Begin VB.TextBox txtobrasMin 
               Height          =   315
               Left            =   960
               TabIndex        =   50
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Maximo:"
               Height          =   195
               Left            =   240
               TabIndex        =   53
               Top             =   780
               Width           =   585
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Minimo:"
               Height          =   195
               Left            =   240
               TabIndex        =   52
               Top             =   420
               Width           =   540
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Servicios"
            Height          =   1215
            Left            =   2760
            TabIndex        =   40
            Top             =   1200
            Width           =   2415
            Begin VB.TextBox txtServiciosMax 
               Height          =   315
               Left            =   960
               TabIndex        =   47
               Top             =   720
               Width           =   1275
            End
            Begin VB.TextBox txtServiciosMin 
               Height          =   315
               Left            =   960
               TabIndex        =   46
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Maximo:"
               Height          =   195
               Left            =   240
               TabIndex        =   49
               Top             =   780
               Width           =   585
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Minimo:"
               Height          =   195
               Left            =   240
               TabIndex        =   48
               Top             =   420
               Width           =   540
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Bienes"
            Height          =   1215
            Left            =   240
            TabIndex        =   39
            Top             =   1200
            Width           =   2415
            Begin VB.TextBox txtBienesMax 
               Height          =   315
               Left            =   900
               TabIndex        =   43
               Top             =   720
               Width           =   1275
            End
            Begin VB.TextBox txtBienesMin 
               Height          =   315
               Left            =   900
               TabIndex        =   42
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Maximo:"
               Height          =   195
               Left            =   180
               TabIndex        =   45
               Top             =   780
               Width           =   585
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Minimo:"
               Height          =   195
               Left            =   180
               TabIndex        =   44
               Top             =   420
               Width           =   540
            End
         End
         Begin VB.TextBox txtCodLimite 
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton cmdCancelarLimites 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6360
            TabIndex        =   32
            Top             =   2540
            Width           =   1335
         End
         Begin VB.CommandButton cmdGrabarLimites 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   4980
            TabIndex        =   31
            Top             =   2540
            Width           =   1335
         End
         Begin VB.TextBox txtDescLimite 
            Height          =   315
            Left            =   1200
            TabIndex        =   30
            Top             =   720
            Width           =   6315
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Consucode Cod"
            Height          =   195
            Left            =   4800
            TabIndex        =   57
            Top             =   420
            Width           =   1140
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Abreviatura"
            Height          =   195
            Left            =   2040
            TabIndex        =   55
            Top             =   420
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   480
            TabIndex        =   35
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   180
            TabIndex        =   34
            Top             =   780
            Width           =   840
         End
      End
      Begin VB.Frame fraEtaReg 
         Caption         =   "Agregar Etapas al Proceso de Selección  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -74880
         TabIndex        =   9
         Top             =   540
         Visible         =   0   'False
         Width           =   7755
         Begin VB.ComboBox cboResp 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1620
            Width           =   6075
         End
         Begin VB.TextBox txtOrden 
            Height          =   315
            Left            =   7020
            MaxLength       =   3
            TabIndex        =   17
            Top             =   420
            Width           =   375
         End
         Begin VB.ComboBox cboEtapa 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   420
            Width           =   4875
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   675
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   840
            Width           =   6075
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   4680
            TabIndex        =   11
            Top             =   2100
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6060
            TabIndex        =   10
            Top             =   2100
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   1680
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Orden"
            Height          =   195
            Left            =   6420
            TabIndex        =   16
            Top             =   480
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Etapa"
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   480
            Width           =   420
         End
      End
      Begin VB.Frame fraEtaVis 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2835
         Left            =   -74880
         TabIndex        =   6
         Top             =   600
         Width           =   7755
         Begin VB.CommandButton cmdQuitarEtapa 
            Caption         =   "Quitar Etapa"
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   2340
            Width           =   1515
         End
         Begin VB.CommandButton cmdAgregaEtapa 
            Caption         =   "Agregar Etapa"
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   2340
            Width           =   1515
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSEta 
            Height          =   2295
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483633
            WordWrap        =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
      End
   End
   Begin VB.Frame fraCab 
      Caption         =   "Procesos de Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton cmdQuitarProceso 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   6600
         TabIndex        =   3
         Top             =   780
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgragarProceso 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSCab 
         Height          =   1455
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Registro de Procesos de Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txtDescProceso 
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   840
         Width           =   6555
      End
      Begin VB.CommandButton cmdGrabarProceso 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5220
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarProceso 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6540
         TabIndex        =   24
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtCodProceso 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   540
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmLogProSelTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String, gnOrden As Integer

Private Sub cmdAgragarProceso_Click()
    Dim oCon As DConecta, Rs As ADODB.Recordset
    fraReg.Visible = True
    fraCab.Visible = False
    sstReg.Enabled = False
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet("Select nMaxNro=Max(nProSelTpoCod) from LogProSelTpo ")
        If Not Rs.EOF Then
           txtCodProceso.Text = Rs!nMaxNro + 1
        Else
           txtCodProceso.Text = 1
        End If
        oCon.CierraConexion
    End If
End Sub

Private Sub cmdAgregar_Click()
    Dim oCon As DConecta, Rs As New ADODB.Recordset
    frListLimites.Visible = False
    FrRegLimites.Visible = True
    sstReg.TabEnabled(1) = False
    cmdSalir.Visible = False
    fraCab.Enabled = False
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet("Select nMaxNro=coalesce(Max(nProSelSubTpo),0) from LogProSelTpoRangos where nProSelTpoCod = " & CInt(Val(MSCab.TextMatrix(MSCab.row, 1))) & " ")
        If Not Rs.EOF Then
           txtCodLimite.Text = Rs!nMaxNro + 1
        Else
           txtCodLimite.Text = 1
        End If
        oCon.CierraConexion
    End If
End Sub

Private Sub cmdCancelarLimites_Click()
    frListLimites.Visible = True
    FrRegLimites.Visible = False
    sstReg.TabEnabled(1) = True
    cmdSalir.Visible = True
    fraCab.Enabled = True
    txtBienesMax.Text = ""
    txtBienesMin.Text = ""
    txtServiciosMax.Text = ""
    txtServiciosMin.Text = ""
    txtObrasMax.Text = ""
    txtobrasMin.Text = ""
    txtCodLimite.Text = ""
    txtDescLimite.Text = ""
    txtAbreviatura.Text = ""
    txtconsucode.Text = ""
End Sub

Private Sub cmdCancelarProceso_Click()
    fraReg.Visible = False
    fraCab.Visible = True
    sstReg.Enabled = True
    txtCodProceso.Text = ""
    txtDescProceso.Text = ""
End Sub

Private Sub cmdGrabarLimites_Click()
On Error GoTo cmdGrabarLimitesErr
    Dim oCon As DConecta, sSQL As String
    If Val(txtBienesMax.Text) < Val(txtBienesMin.Text) Then
        MsgBox "Rango de Bienes Incorrecto...", vbInformation, "Aviso"
        Exit Sub
    End If
    If Val(txtServiciosMax.Text) < Val(txtServiciosMin.Text) Then
        MsgBox "Rango de Servicios Incorrecto...", vbInformation, "Aviso"
        Exit Sub
    End If
    If Val(txtObrasMax.Text) < Val(txtobrasMin.Text) Then
        MsgBox "Rango de Obras Incorrecto...", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "insert into LogProSelTpoRangos(nProSelTpoCod,nProSelSubTpo, cProSelSubTpo, cAbreviatura, nConsucodeCod, nBienesMin, nBienesMax, nObrasMin, nObrasMax, nServiMin,nServiMax) " & _
                " values(" & MSCab.TextMatrix(MSCab.row, 1) & "," & txtCodLimite.Text & ",'" & txtDescLimite.Text & "','" & txtAbreviatura.Text & "'," & txtconsucode.Text & "," & txtBienesMin.Text & "," & txtBienesMax.Text & "," & txtobrasMin.Text & "," & txtObrasMax.Text & "," & txtServiciosMin.Text & "," & txtServiciosMax.Text & ")"
        oCon.Ejecutar sSQL
        oCon.CierraConexion
    End If
    ListaTiposProcesosSeleccion MSCab.TextMatrix(MSCab.row, 1)
    cmdCancelarLimites_Click
    Exit Sub
cmdGrabarLimitesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cmdGrabarProceso_Click()
On Error GoTo cmdGrabarProcesoErr
    Dim oCon As DConecta, sSQL As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "insert into LogProSelTpo(nProSelTpoCod, cProSelTpoDescripcion) " & _
                " values(" & txtCodProceso.Text & ",'" & txtDescProceso.Text & "')"
        oCon.Ejecutar sSQL
        MsgBox "Proceso Creado Correctamente", vbInformation
        oCon.CierraConexion
        ListaProcesosSeleccion
    End If
    cmdCancelarProceso_Click
    Exit Sub
cmdGrabarProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cmdQuitar_Click()
On Error GoTo cmdQuitarErr
    Dim oCon As DConecta, sSQL As String
    
    If Len(MSCab.TextMatrix(MSCab.row, 1)) = 0 Or Len(MSDet.TextMatrix(MSDet.row, 1)) = 0 Then Exit Sub
    
    If MsgBox("Desea Eliminar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "delete LogProSelTpoRangos where nProSelTpoCod = " & MSCab.TextMatrix(MSCab.row, 1) & " and nProSelSubTpo= " & MSDet.TextMatrix(MSDet.row, 1)
        oCon.Ejecutar sSQL
        MsgBox "Limite Eliminado Correctamente", vbInformation
        oCon.CierraConexion
    End If
    ListaTiposProcesosSeleccion MSCab.TextMatrix(MSCab.row, 1)
    Exit Sub
cmdQuitarErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cmdQuitarProceso_Click()
On Error GoTo cmdQuitarProcesoErr
    Dim oCon As DConecta, sSQL As String
    If MsgBox("Desea Eliminar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "delete LogProSelTpo where nProSelTpoCod = " & MSCab.TextMatrix(MSCab.row, 1)
        oCon.Ejecutar sSQL
        MsgBox "Proceso Eliminado Correctamente", vbInformation
        oCon.CierraConexion
        ListaProcesosSeleccion
    End If
    Exit Sub
cmdQuitarProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub Form_Load()
CentraForm Me
ListaProcesosSeleccion
sstReg.Tab = 0
gnOrden = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub ListaProcesosSeleccion()
Dim Rs As New ADODB.Recordset, oConn As New DConecta, i As Integer
If Not oConn.AbreConexion Then
   Exit Sub
End If
LimpiaFlexCab
sSQL = "select nProSelTpoCod, cProSelTpoDescripcion from LogProSelTpo"
Set Rs = oConn.CargaRecordSet(sSQL)
If Not Rs.EOF Then
   Do While Not Rs.EOF
      i = i + 1
      InsRow MSCab, i
      MSCab.TextMatrix(i, 1) = Rs!nProSelTpoCod
      MSCab.TextMatrix(i, 2) = Rs!cProSelTpoDescripcion
      Rs.MoveNext
   Loop
End If
End Sub

Sub LimpiaFlexCab()
MSCab.Clear
MSCab.Rows = 2
MSCab.RowHeight(0) = 320
MSCab.RowHeight(1) = 8
MSCab.ColWidth(0) = 0
MSCab.ColWidth(1) = 260:  MSCab.ColAlignment(1) = 4
MSCab.ColWidth(2) = 5800
MSCab.ColWidth(3) = 0
MSCab.ColWidth(4) = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelTipos = Nothing
End Sub

Private Sub MSCab_GotFocus()
If sstReg.Tab = 0 Then
   ListaTiposProcesosSeleccion CInt(Val(MSCab.TextMatrix(MSCab.row, 1)))
Else
   ListaEtapasxProceso CInt(Val(MSCab.TextMatrix(MSCab.row, 1)))
End If
End Sub

Private Sub MSCab_RowColChange()
If sstReg.Tab = 0 Then
   ListaTiposProcesosSeleccion MSCab.TextMatrix(MSCab.row, 1)
Else
   ListaEtapasxProceso MSCab.TextMatrix(MSCab.row, 1)
End If
End Sub

Sub ListaTiposProcesosSeleccion(vProSelCod As Integer)
Dim Rs As New ADODB.Recordset, oConn As New DConecta, i As Integer
       
LimpiaFlexDet
If Not oConn.AbreConexion Then
   Exit Sub
End If
sSQL = "select nProSelSubTpo, cProSelSubTpo," & _
       "  nBienesMin, nBienesMax, nObrasMin, nObrasMax, nServiMin,nServiMax" & _
       " " & _
       "  from LogProSelTpoRangos " & _
       " where nProSelTpoCod = " & vProSelCod & " "
Set Rs = oConn.CargaRecordSet(sSQL)
If Not Rs.EOF Then
   Do While Not Rs.EOF
      i = i + 1
      InsRow MSDet, i
      MSDet.RowHeight(i) = 460
      MSDet.TextMatrix(i, 1) = Rs!nProSelSubTpo
      MSDet.TextMatrix(i, 2) = Rs!cProSelSubTpo
      MSDet.TextMatrix(i, 3) = FNumero(Rs!nBienesMin)
      MSDet.TextMatrix(i, 4) = FNumero(Rs!nBienesMax)
      MSDet.TextMatrix(i, 5) = FNumero(Rs!nObrasMin)
      MSDet.TextMatrix(i, 6) = FNumero(Rs!nObrasMax)
      MSDet.TextMatrix(i, 7) = FNumero(Rs!nServiMin)
      MSDet.TextMatrix(i, 8) = FNumero(Rs!nServiMax)
      Rs.MoveNext
   Loop
End If
End Sub

Sub LimpiaFlexDet()
MSDet.Clear
MSDet.Rows = 2
MSDet.RowHeight(0) = 420
MSDet.RowHeight(1) = 8
MSDet.ColWidth(0) = 0
MSDet.ColWidth(1) = 260:  MSDet.ColAlignment(1) = 4
MSDet.ColWidth(2) = 1800
MSDet.ColWidth(3) = 900
MSDet.ColWidth(4) = 900
MSDet.ColWidth(5) = 900:  MSDet.TextMatrix(0, 3) = "Bienes Mínimo"
MSDet.ColWidth(6) = 900:  MSDet.TextMatrix(0, 4) = "Bienes Máximo"
MSDet.ColWidth(7) = 900:  MSDet.TextMatrix(0, 5) = "Obras Mínimo"
MSDet.ColWidth(8) = 900:  MSDet.TextMatrix(0, 6) = "Obras Máximo"
MSDet.ColWidth(9) = 900:  MSDet.TextMatrix(0, 7) = "Serv. Mínimo"
MSDet.ColWidth(10) = 900: MSDet.TextMatrix(0, 8) = "Serv. Máximo"
End Sub

'*************************************************************************

Sub ListaEtapasxProceso(vProSelCod As Integer)
Dim Rs As New ADODB.Recordset, oConn As New DConecta, i As Integer
If Not oConn.AbreConexion Then
   Exit Sub
End If

LimpiaFlexEta

sSQL = " Select t.nOrden, t.nEtapaCod, t.cDescripcion, e.cDescripcion cEtapa, c.cCargo, t.nPlazo " & _
       " from LogProSelTpoEtapa t " & _
       "  LEFT OUTER join LogProSelTpoEtapaCargo c on c.nProSelTpoCod = t.nProSelTpoCod and c.nEtapaCod=t.nEtapaCod" & _
       " inner join LogEtapa e on t.nEtapaCod = e.nEtapaCod and e.nEstado = 1 " & _
       " where t.nProSelTpoCod = " & vProSelCod & " " & _
       " order by t.nOrden "
       
       '" inner join (SELECT nConsValor as nEtapaResp, cConsDescripcion as cResponsable FROM Constante WHERE nConsCod = " & gcResponsableProceso & " AND nConsCod <> nConsValor) r on c.cCargo = r.nEtapaResp " & _

Set Rs = oConn.CargaRecordSet(sSQL)
If Not Rs.EOF Then
   Do While Not Rs.EOF
      i = i + 1
      InsRow MSEta, i
      MSEta.RowHeight(i) = 400
      MSEta.TextMatrix(i, 0) = Rs!nEtapaCod
      MSEta.TextMatrix(i, 1) = Rs!nOrden
      MSEta.TextMatrix(i, 2) = Rs!cEtapa
      MSEta.TextMatrix(i, 3) = IIf(IsNull(Rs!cCargo), "", Rs!cCargo)
      MSEta.TextMatrix(i, 4) = Rs!cDescripcion
      Rs.MoveNext
   Loop
End If
End Sub

Sub LimpiaFlexEta()
MSEta.Clear
MSEta.Rows = 2
MSEta.RowHeight(0) = 320
MSEta.RowHeight(1) = 8
MSEta.ColWidth(0) = 0
MSEta.ColWidth(1) = 260:   MSEta.TextMatrix(0, 1) = "Nro": MSEta.ColAlignment(1) = 4
MSEta.ColWidth(2) = 4000:  MSEta.TextMatrix(0, 2) = "Etapa"
MSEta.ColWidth(3) = 2400:  MSEta.TextMatrix(0, 3) = "Responsable": MSEta.ColAlignment(3) = 4
MSEta.ColWidth(4) = 4000:  MSEta.TextMatrix(0, 4) = "Comentarios"
End Sub

Private Sub sstReg_Click(PreviousTab As Integer)
If sstReg.Tab = 1 Then
   ListaEtapasxProceso CInt(Val(MSCab.TextMatrix(MSCab.row, 1)))
End If
End Sub

Private Sub cmdAgregaEtapa_Click()
Dim oConn As New DConecta, Rs As New ADODB.Recordset
Dim nProSelTpo As Integer
Dim i As Integer, N As Integer
Dim cConsulta As String
Dim nNro As Integer

cboEtapa.Clear
fraEtaVis.Visible = False
fraEtaReg.Visible = True
txtDescripcion.Text = ""
cmdSalir.Visible = False
sstReg.TabEnabled(0) = False
fraCab.Enabled = False

If oConn.AbreConexion Then
   nProSelTpo = VNumero(MSCab.TextMatrix(MSCab.row, 1))
   Set Rs = oConn.CargaRecordSet("Select nSgte = coalesce(Max(nOrden),0) from LogProSelTpoEtapa where nProSelTpoCod = " & nProSelTpo & " ")
   If Not Rs.EOF Then
      txtOrden.Text = Rs!nSgte + 1
      gnOrden = txtOrden.Text
   End If
End If

nNro = 0
cConsulta = " AND nEtapaCod NOT IN ("
N = MSEta.Rows - 1
For i = 1 To N
    If VNumero(MSEta.TextMatrix(i, 0)) > 0 Then
       cConsulta = cConsulta + MSEta.TextMatrix(i, 0) + ","
       nNro = nNro + 1
    End If
Next
cConsulta = Mid(cConsulta, 1, Len(cConsulta) - 1) + ") "
If nNro = 0 Then cConsulta = ""

sSQL = "SELECT nEtapaCod, cDescripcion " & _
       "FROM LogEtapa WHERE nEstado = 1 " + cConsulta
       
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboEtapa.AddItem Rs!cDescripcion
         cboEtapa.ItemData(cboEtapa.ListCount - 1) = Rs!nEtapaCod
         Rs.MoveNext
      Loop
      cboEtapa.ListIndex = 0
   End If
End If

sSQL = "SELECT nConsValor, cConsDescripcion " & _
       "FROM Constante WHERE (nConsCod = " & gcResponsableProceso & ") AND (nConsCod <> nConsValor) " '+ cConsulta
       
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cboResp.AddItem Rs!cConsDescripcion
         cboResp.ItemData(cboResp.ListCount - 1) = Rs!nConsValor
         Rs.MoveNext
      Loop
      cboResp.ListIndex = 0
   End If
End If

End Sub

Private Sub cmdCancelar_Click()
fraEtaReg.Visible = False
fraEtaVis.Visible = True
fraEtaReg.Visible = False
fraEtaVis.Visible = True
cmdSalir.Visible = True
sstReg.TabEnabled(0) = True
fraCab.Enabled = True
gnOrden = 0
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo cmdGrabarErr
Dim oConn As New DConecta
Dim nEtapaCod As Integer, nProSelTpo As Integer, nOrden As Integer, nProSelSubTpo As Integer
Dim Rs As New ADODB.Recordset

If Val(txtOrden.Text) = 0 Then
    MsgBox "Error Debe Ingresar un Numero Mayor a Cero para el Orden", vbInformation, "Aviso"
    Exit Sub
End If

If Val(txtOrden.Text) > gnOrden Then
    MsgBox "Error Debe Ingresar un Numero Menor o Igual a " & gnOrden & " para el Orden", vbInformation, "Aviso"
    txtOrden.Text = gnOrden
    txtOrden.SetFocus
    Exit Sub
End If

nProSelTpo = CInt(VNumero(MSCab.TextMatrix(MSCab.row, 1)))
nProSelSubTpo = CInt(VNumero(MSDet.TextMatrix(MSDet.row, 1)))
nEtapaCod = cboEtapa.ItemData(cboEtapa.ListIndex)
nOrden = CInt(VNumero(txtOrden))

'If nFuncion = 1 Then
   If oConn.AbreConexion Then
      Set Rs = oConn.CargaRecordSet("Select * from LogProSelTpoEtapa WHERE nProSelTpoCod = " & nProSelTpo & "  and nEtapaCod = " & nEtapaCod & " ")
      If Not Rs.EOF Then
         MsgBox "Ya existe un registro de la Etapa en el Proceso de Selección..." + Space(10), vbInformation
         Exit Sub
      End If
      oConn.CierraConexion
   End If
'End If

If MsgBox("¿ Está seguro de agregar la Etapa ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If oConn.AbreConexion Then
      oConn.BeginTrans
      
      sSQL = "update LogProSelTpoEtapa set nOrden=nOrden+1 where nProSelTpoCod = " & nProSelTpo & " and nProSelSubTpo = " & nProSelSubTpo & " and nOrden >= " & nOrden
      oConn.Ejecutar sSQL
                         'nProSelTpoCod  nEtapaCod  nProSelSubTpo
      sSQL = "INSERT INTO LogProSelTpoEtapa (nProSelTpoCod, nEtapaCod, nProSelSubTpo, cDescripcion, nOrden, nPlazo ) " & _
             "       VALUES (" & nProSelTpo & "," & nEtapaCod & "," & nProSelSubTpo & ",'" & txtDescripcion.Text & "'," & nOrden & ",0)  "
      oConn.Ejecutar sSQL
                                                 'nProSelTpoCod nEtapaCod   nProSelSubTpo cCargo bObligatorio
      sSQL = "INSERT INTO LogProSelTpoEtapaCargo (nProSelTpoCod,nEtapaCod,nProSelSubTpo,cCargo) " & _
             "       VALUES (" & nProSelTpo & "," & nEtapaCod & "," & nProSelSubTpo & ",'" & cboResp.Text & "')  "
      oConn.Ejecutar sSQL
      
      oConn.CommitTrans
      
      fraEtaReg.Visible = False
      fraEtaVis.Visible = True
      cmdSalir.Visible = True
      sstReg.TabEnabled(0) = True
      fraCab.Enabled = True
      oConn.CierraConexion
      ListaEtapasxProceso MSCab.TextMatrix(MSCab.row, 1)
      gnOrden = 0
   End If
End If
Exit Sub
cmdGrabarErr:
    oConn.RollbackTrans
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdQuitarEtapa_Click()
On Error GoTo cmdQuitarEtapaErr
Dim oConn As New DConecta
Dim nEtapaCod As Integer
Dim nProSelTpo As Integer, nProSelSubTpo As Integer
Dim nOrden As Integer

If Len(MSEta.TextMatrix(MSEta.row, 0)) = 0 Then Exit Sub

nEtapaCod = MSEta.TextMatrix(MSEta.row, 0)
nProSelTpo = MSCab.TextMatrix(MSCab.row, 1)
nProSelSubTpo = MSDet.TextMatrix(MSDet.row, 1)
nOrden = MSEta.TextMatrix(MSEta.row, 1)

If MsgBox("¿ Está seguro de quitar la Etapa indicada ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If oConn.AbreConexion Then
      oConn.BeginTrans
      
      sSQL = "DELETE FROM LogProSelTpoEtapaCargo WHERE nProSelTpoCod = " & nProSelTpo & " AND nEtapaCod = " & nEtapaCod & " "
      oConn.Ejecutar sSQL
      
      sSQL = "update LogProSelTpoEtapa set nOrden=nOrden-1 where nProSelTpoCod = " & nProSelTpo & " and nProSelSubTpo = " & nProSelSubTpo & " and nOrden >= " & nOrden
      oConn.Ejecutar sSQL
      
      sSQL = "DELETE FROM LogProSelTpoEtapa WHERE nProSelTpoCod = " & nProSelTpo & " AND nEtapaCod = " & nEtapaCod & " "
      oConn.Ejecutar sSQL
      
      oConn.CommitTrans
      ListaEtapasxProceso nProSelTpo
   End If
End If
    Exit Sub
cmdQuitarEtapaErr:
    oConn.RollbackTrans
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub txtBienesMax_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtBienesMax, KeyAscii)
End Sub

Private Sub txtBienesMin_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtBienesMin, KeyAscii)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtObrasMax_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtObrasMax, KeyAscii)
End Sub

Private Sub txtobrasMin_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtobrasMin, KeyAscii)
End Sub

Private Sub txtOrden_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Sub txtServiciosMax_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtServiciosMax, KeyAscii)
End Sub

Private Sub txtServiciosMin_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtServiciosMin, KeyAscii)
End Sub
