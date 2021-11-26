VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRHCargos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6510
   ClientLeft      =   1635
   ClientTop       =   2730
   ClientWidth     =   8625
   Icon            =   "frmCurriculum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7620
      TabIndex        =   7
      Top             =   6105
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   6105
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "E&liminar"
      Height          =   375
      Left            =   2175
      TabIndex        =   5
      Top             =   6105
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Height          =   375
      Left            =   1095
      TabIndex        =   4
      Top             =   6105
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "N&uevo"
      Height          =   375
      Left            =   15
      TabIndex        =   1
      Top             =   6105
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6060
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   10689
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "&Niveles"
      TabPicture(0)   =   "frmCurriculum.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Cargos"
      TabPicture(1)   =   "frmCurriculum.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4830
         Left            =   120
         TabIndex        =   23
         Top             =   1125
         Width           =   8370
         Begin TabDlg.SSTab SSTab 
            Height          =   2115
            Left            =   90
            TabIndex        =   25
            Top             =   2655
            Width           =   8205
            _ExtentX        =   14473
            _ExtentY        =   3731
            _Version        =   393216
            TabHeight       =   520
            ForeColor       =   8388608
            TabCaption(0)   =   "Datos"
            TabPicture(0)   =   "frmCurriculum.frx":0342
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraDatos"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Areas"
            TabPicture(1)   =   "frmCurriculum.frx":035E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraAreas"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Def Planeamiento"
            TabPicture(2)   =   "frmCurriculum.frx":037A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame4"
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               ForeColor       =   &H80000008&
               Height          =   1665
               Left            =   -74910
               TabIndex        =   36
               Top             =   345
               Width           =   8010
               Begin VB.ComboBox cboDirInd 
                  Height          =   315
                  Left            =   1545
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   1065
                  Width           =   5445
               End
               Begin VB.ComboBox cboGrupoPla 
                  Height          =   315
                  Left            =   1545
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   495
                  Width           =   5445
               End
               Begin VB.Label Label1 
                  Caption         =   "Grupo Planea :"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   41
                  Top             =   1125
                  Width           =   1290
               End
               Begin VB.Label lblGrupoPla 
                  Caption         =   "Grupo Planea :"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   40
                  Top             =   555
                  Width           =   1290
               End
            End
            Begin VB.Frame fraAreas 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               ForeColor       =   &H80000008&
               Height          =   1665
               Left            =   -74910
               TabIndex        =   35
               Top             =   330
               Width           =   8010
               Begin VB.CommandButton cmdAreaEliminar 
                  Caption         =   "Eliminar"
                  Height          =   300
                  Left            =   6930
                  TabIndex        =   43
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CommandButton cmdAreaNuevo 
                  Caption         =   "Nuevo"
                  Height          =   300
                  Left            =   6915
                  TabIndex        =   42
                  Top             =   345
                  Width           =   1020
               End
               Begin Sicmact.FlexEdit FlexArea 
                  Height          =   1395
                  Left            =   60
                  TabIndex        =   37
                  Top             =   195
                  Width           =   6750
                  _ExtentX        =   13838
                  _ExtentY        =   2487
                  Cols0           =   3
                  EncabezadosNombres=   "#-Area-Descripcion"
                  EncabezadosAnchos=   "300-2000-4000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ColumnasAEditar =   "X-1-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-1-0"
                  EncabezadosAlineacion=   "C-L-L"
                  FormatosEdit    =   "0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbFlexDuplicados=   0   'False
                  lbPuntero       =   -1  'True
                  Appearance      =   0
                  ColWidth0       =   300
                  RowHeight0      =   300
               End
            End
            Begin VB.Frame fraDatos 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               ForeColor       =   &H80000008&
               Height          =   1665
               Left            =   90
               TabIndex        =   26
               Top             =   345
               Width           =   8010
               Begin VB.TextBox txtcCarCod 
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   1080
                  TabIndex        =   30
                  Top             =   360
                  Width           =   1575
               End
               Begin VB.TextBox txtcCarDes 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   1080
                  MaxLength       =   255
                  TabIndex        =   29
                  Top             =   720
                  Width           =   6795
               End
               Begin VB.TextBox txtGrado 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   7155
                  TabIndex        =   28
                  Top             =   360
                  Width           =   705
               End
               Begin VB.CheckBox chkCtrAsist 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000000&
                  Caption         =   "Control de Asistencia"
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
                  Height          =   225
                  Left            =   5445
                  TabIndex        =   27
                  Top             =   1155
                  Width           =   2385
               End
               Begin Sicmact.EditMoney txtCarSue 
                  Height          =   375
                  Left            =   1080
                  TabIndex        =   31
                  Top             =   1080
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Appearance      =   0
                  Text            =   "0"
               End
               Begin VB.Label lblSueldo 
                  Caption         =   "Sueldo"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   34
                  Top             =   1140
                  Width           =   735
               End
               Begin VB.Label lblCargo 
                  Caption         =   "Cargos:"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   33
                  Top             =   390
                  Width           =   615
               End
               Begin VB.Label lblcCarDes 
                  Caption         =   "Descrip."
                  Height          =   255
                  Left            =   240
                  TabIndex        =   32
                  Top             =   750
                  Width           =   615
               End
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
            Height          =   2340
            Left            =   90
            TabIndex        =   24
            Top             =   225
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   4128
            _Version        =   393216
            FixedCols       =   0
            ForeColorFixed  =   8388608
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nivel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   795
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   8370
         Begin VB.ComboBox cmbCargos 
            Height          =   315
            Left            =   810
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   315
            Width           =   6480
         End
         Begin VB.Label lblNivel 
            Caption         =   "Nivel :"
            Height          =   255
            Left            =   90
            TabIndex        =   22
            Top             =   330
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nivel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5595
         Left            =   -74895
         TabIndex        =   8
         Top             =   345
         Width           =   8355
         Begin VB.TextBox txtCardesc 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   13
            Top             =   4020
            Width           =   6930
         End
         Begin VB.TextBox txtOrden 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   5190
            Width           =   615
         End
         Begin VB.ComboBox cmbCatCod 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   4830
            Width           =   6900
         End
         Begin Sicmact.EditMoney txtSueldoMaximo 
            Height          =   375
            Left            =   6540
            TabIndex        =   9
            Top             =   4395
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin Sicmact.EditMoney txtSueldoNivel 
            Height          =   375
            Left            =   1320
            TabIndex        =   10
            Top             =   4395
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexNivel 
            Height          =   3720
            Left            =   90
            TabIndex        =   14
            Top             =   225
            Width           =   8160
            _ExtentX        =   14393
            _ExtentY        =   6562
            _Version        =   393216
            FixedCols       =   0
            ForeColorFixed  =   8388608
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblCargos 
            Caption         =   "Descripcripción"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   4065
            Width           =   1215
         End
         Begin VB.Label lblSueldoMinimo 
            Caption         =   "Sueldo.Niv"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   4455
            Width           =   855
         End
         Begin VB.Label lblOrden 
            Caption         =   "Orden"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   5220
            Width           =   495
         End
         Begin VB.Label lblMaximo 
            Caption         =   "Sueldo.Max."
            Height          =   255
            Left            =   5340
            TabIndex        =   16
            Top             =   4455
            Width           =   975
         End
         Begin VB.Label lblCadCod 
            Caption         =   "Categorias"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   4860
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   15
      TabIndex        =   2
      Top             =   6105
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1095
      TabIndex        =   3
      Top             =   6105
      Width           =   975
   End
End
Attribute VB_Name = "frmRHCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lbEditado As Boolean
Dim lbEditadoGen As Boolean
Dim lsCodigo As String
Dim lnIniOpe As Integer
Dim lnTipoOpe As TipoOpe

Private Sub cmbCargos_Click()
    Dim oCargos As DActualizadatosCargo
    Set oCargos = New DActualizadatosCargo
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oCargos.GetCargos(False, Right(Me.cmbCargos, 3))
    
    If rs.EOF And rs.BOF Then Exit Sub
    
    
    Set Me.Flex.DataSource = rs
    Set oCargos = Nothing
    
    Flex.ColWidth(0) = 1
    Flex.ColWidth(1) = 1100
    Flex.ColWidth(2) = 4500
    Flex.ColWidth(3) = 1200
    Flex.ColWidth(4) = 1000
    Flex.ColAlignment(3) = 7
    
    Flex.Row = 1
    Flex_EnterCell
End Sub

Private Sub cmbCatCod_Click()
    'cmbCatCod_Change
End Sub

Private Sub cmbCatCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtOrden.SetFocus
End Sub

Private Sub cmdAreaEliminar_Click()
    If MsgBox("Desea Eliminar esta area " & FlexArea.TextMatrix(FlexArea.Row, 2) & " ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    FlexArea.EliminaFila FlexArea.Row
End Sub

Private Sub cmdAreaNuevo_Click()
    FlexArea.AdicionaFila
End Sub

Private Sub cmdGrabar_Click()
    Dim oCargos As NActualizaDatosCargo
    Set oCargos = New NActualizaDatosCargo

    If Me.SSTab1.Tab = 0 Then
        If Me.txtSueldoNivel.value < 0 Then
            MsgBox "Monto no valido.", vbInformation, "Aviso"
            txtSueldoNivel.SetFocus
            Exit Sub
        ElseIf Me.txtSueldoMaximo.value < 0 Then
            MsgBox "Monto no valido.", vbInformation, "Aviso"
            txtSueldoMaximo.SetFocus
            Exit Sub
        ElseIf Trim(Me.txtCardesc) = "" Then
            MsgBox "Debe ingresar una descripción.", vbInformation, "Aviso"
            txtCardesc.SetFocus
            Exit Sub
        ElseIf Trim(Me.cmbCatCod) = "" Then
            MsgBox "Debe ingresar una Categoria.", vbInformation, "Aviso"
            Me.cmbCatCod.SetFocus
            Exit Sub
        End If
        
        If SSTab1.Tab <> lnIniOpe Then
            If lnIniOpe = 0 Then
                MsgBox "Ud. ha elegido actualizar un Nivel, para poder actualizar un Nivel, debe terminar antes la operacion antes iniciada.", vbInformation, "Aviso"
            Else
                MsgBox "Ud. ha elegido actualizar un Cargo, para poder actualizar un Nivel, debe terminar antes la operacion antes iniciada.", vbInformation, "Aviso"
            End If
            SSTab1.Tab = lnIniOpe
            Exit Sub
        Else
            lnIniOpe = -1
        End If
        
        If lbEditadoGen = False Then
            oCargos.AgregaCargo lsCodigo, Me.txtCardesc.Text, Me.txtSueldoNivel.value, Me.txtSueldoMaximo.value, Right(Me.cmbCatCod, 1), GetMovNro(gsCodUser, gsCodAge), Me.txtOrden, 0, 1, 0, 0
        Else
            oCargos.ModificaCargo Me.FlexNivel.TextMatrix(FlexNivel.Row, 1), Me.txtCardesc.Text, Me.txtSueldoNivel.value, Me.txtSueldoMaximo.value, Right(Me.cmbCatCod, 1), GetMovNro(gsCodUser, gsCodAge), Me.txtOrden, 0, 1, 0, 0
        End If

        lbEditadoGen = False
        Me.txtCardesc.Text = ""
        txtSueldoNivel.value = 0
        txtSueldoMaximo.value = 0
        Me.txtOrden.Text = "0"
        Me.FlexNivel.Enabled = True
        Refresca_C
        
        Activa True
        Me.FlexNivel.SetFocus
        FlexNivel.Row = 1
    Else
        If Trim(Me.cmbCargos) = "" Then
            MsgBox "Debe elegir un nivel.", vbInformation, "Aviso"
            SSTab.Tab = 0
            Me.cmbCargos.SetFocus
            Exit Sub
        ElseIf Trim(txtcCarDes) = "" Then
            MsgBox "Debe ingresar una descripción.", vbInformation, "Aviso"
            SSTab.Tab = 0
            txtcCarDes.SetFocus
            Exit Sub
        ElseIf Not IsNumeric(Me.txtGrado.Text) Then
            MsgBox "Debe ingresar un grado de aprobacion valido.", vbInformation, "Aviso"
            SSTab.Tab = 0
            txtGrado.SetFocus
            Exit Sub
        ElseIf Trim(Me.cboGrupoPla.Text) = "" Then
            MsgBox "Debe ingresar un grupo vaido.", vbInformation, "Aviso"
            SSTab.Tab = 2
            cboGrupoPla.SetFocus
            Exit Sub
        ElseIf Trim(Me.cboDirInd.Text) = "" Then
            MsgBox "Debe ingresar un tipo de gasto valido.", vbInformation, "Aviso"
            SSTab.Tab = 2
            cboDirInd.SetFocus
            Exit Sub
        End If
        
        If SSTab1.Tab <> lnIniOpe Then
            If lnIniOpe = 0 Then
                MsgBox "Ud. ha elegido actualizar un Nivel, para poder actualizar un Nivel, debe terminar antes la operacion antes iniciada.", vbInformation, "Aviso"
            Else
                MsgBox "Ud. ha elegido actualizar un Cargo, para poder actualizar un Nivel, debe terminar antes la operacion antes iniciada.", vbInformation, "Aviso"
            End If
            SSTab1.Tab = lnIniOpe
            Exit Sub
        Else
            lnIniOpe = -1
        End If
        
        If Not lbEditado Then
            oCargos.AgregaCargo Me.txtcCarCod.Text, Me.txtcCarDes.Text, Me.txtCarSue.value, 0, "", GetMovNro(gsCodUser, gsCodAge), 0, Me.txtGrado.Text, Me.chkCtrAsist.value, Right(Me.cboGrupoPla.Text, 5), Right(Me.cboDirInd.Text, 5)
        Else
            oCargos.ModificaCargo Me.txtcCarCod.Text, Me.txtcCarDes.Text, Me.txtCarSue.value, 0, "", GetMovNro(gsCodUser, gsCodAge), 0, Me.txtGrado.Text, Me.chkCtrAsist.value, Right(Me.cboGrupoPla.Text, 5), Right(Me.cboDirInd.Text, 5)
        End If
        oCargos.SetAreasCargo FlexArea.GetRsNew, Me.txtcCarCod.Text
        
        txtcCarCod.Text = ""
        txtcCarDes.Text = ""
        txtCarSue.value = 0
        txtcCarDes.SetFocus
        
        Activa True
        Me.Flex.SetFocus
        Flex.Row = 1
        cmbCargos_Click
    End If

    
End Sub

Private Sub CmdCancelar_Click()
    If Me.SSTab1.Tab = 0 Then
        lbEditadoGen = False
        Me.txtCardesc.Text = ""
        Me.txtSueldoNivel.value = 0
        Me.txtSueldoMaximo.value = 0
        Me.txtOrden.Text = "0"
        Me.FlexNivel.Enabled = True
        Me.cmbCatCod.ListIndex = -1
    Else
        txtcCarDes = ""
        txtCarSue.value = 0
        If txtCarSue.Enabled Then
            txtCarSue.SetFocus
        End If
        Flex_EnterCell
    End If
    
    Activa True
End Sub

Private Sub cmdEditar_Click()
    Activa False
    lnIniOpe = Me.SSTab1.Tab
    If Me.SSTab1.Tab = 0 Then
        If Me.txtCardesc.Text = "" Then
            Exit Sub
        Else
            Me.txtCardesc.SetFocus
        End If
    Else
        If Me.Flex.TextMatrix(Flex.Row, 2) = "" Then
            Exit Sub
        Else
            Me.txtcCarDes.SetFocus
        End If
    End If
    
    lbEditado = True
    lbEditadoGen = True
End Sub

Private Sub cmdEliminar_Click()
    Dim sqlD As String
    Dim oCargo As NActualizaDatosCargo
    Set oCargo = New NActualizaDatosCargo
    
    If Me.SSTab1.Tab = 0 Then
        If Not oCargo.VerificaCargoUsado(Me.FlexNivel.TextMatrix(FlexNivel.Row, 1)) Then
            If MsgBox("Se eliminara todos los niveles y sus cargos. Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            oCargo.EliminaCargo Me.FlexNivel.TextMatrix(FlexNivel.Row, 1)
            Refresca_C
        Else
            MsgBox "El Nivel esta siedo usado, no se puede eliminar.", vbInformation, "Aviso"
        End If
    Else
        If Not oCargo.VerificaCargoUsado(Me.Flex.TextMatrix(Me.Flex.Row, 1)) Then
            If MsgBox("Se eliminara el cargo. Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            oCargo.EliminaCargo Me.Flex.TextMatrix(Flex.Row, 1)
            cmbCargos_Click
        Else
            MsgBox "El Nivel esta siedo usado, no se puede eliminar.", vbInformation, "Aviso"
        End If
    End If
    
End Sub

Private Sub cmdImprimir_Click()
    Dim oCargo As NActualizaDatosCargo
    Dim lsCadena As String
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Set oCargo = New NActualizaDatosCargo
    
    lsCadena = oCargo.GetReporte(gsNomAge, gsEmpresa, gdFecSis)
    
    oPrevio.Show lsCadena, "Niveles y Cargos.", True, 66
End Sub

Private Sub cmdNuevo_Click()
    Dim lsNumCur As String
    
    lnIniOpe = SSTab1.Tab
    If Me.SSTab1.Tab = 0 Then
        'Me.FlexNivel.Enabled = False
        lbEditadoGen = False
        Me.txtCardesc.Text = ""
        Me.txtSueldoNivel.value = 0
        Me.txtSueldoMaximo.value = 0
        Me.txtOrden.Text = "0"
        Me.cmbCatCod.ListIndex = -1
        If Not IsNumeric(Me.FlexNivel.TextMatrix(FlexNivel.Rows - 1, 1)) Then
            lsNumCur = "001"
        Else
            lsNumCur = FillNum(Trim(Str(CCur(FlexNivel.TextMatrix(FlexNivel.Rows - 1, 1) + 1))), 3, "0")
        End If
        lsCodigo = lsNumCur
        Me.txtCardesc.SetFocus
    Else
        If Not IsNumeric(Flex.TextMatrix(Flex.Rows - 1, 1)) Then
            lsNumCur = "1"
        Else
            lsNumCur = Trim(Str(CCur(Mid(Flex.TextMatrix(Flex.Rows - 1, 1), 4, 3) + 1)))
        End If
        txtcCarCod = Right(Me.cmbCargos, 3) & FillNum(lsNumCur, 3, "0")
        txtcCarDes = ""
        txtCarSue.value = 0
        lbEditado = False
        
        Me.txtcCarDes.SetFocus
    End If
    Activa False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub txtcCarDes_GotFocus()
    txtcCarDes.SelStart = 0
    txtcCarDes.SelLength = 50
End Sub

Private Sub txtGrado_GotFocus()
    txtGrado.SelStart = 0
    txtGrado.SelLength = 50
End Sub

Private Sub txtSueldoNivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtSueldoMaximo.SetFocus
End Sub

Private Sub txtSueldoMaximo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmbCatCod.SetFocus
End Sub

Private Sub Flex_DblClick()
    If lnTipoOpe = gTipoOpeConsulta Then Exit Sub
    cmdEditar_Click
End Sub

Private Sub Flex_EnterCell()
    Dim oCar As DActualizadatosCargo
    Set oCar = New DActualizadatosCargo
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If Me.Flex.TextMatrix(Flex.Row, 1) = "" Then Exit Sub
    
    Me.txtcCarDes.Text = Me.Flex.TextMatrix(Flex.Row, 2)
    Me.txtCarSue.value = Me.Flex.TextMatrix(Flex.Row, 3)
    Me.txtcCarCod = Me.Flex.TextMatrix(Flex.Row, 1)
    Me.txtGrado.Text = Me.Flex.TextMatrix(Flex.Row, 4)
    Me.chkCtrAsist.value = Me.Flex.TextMatrix(Flex.Row, 5)
    
    UbicaCombo Me.cboGrupoPla, Me.Flex.TextMatrix(Flex.Row, 6)
    UbicaCombo Me.cboDirInd, Me.Flex.TextMatrix(Flex.Row, 7)
    
    Set rs = oCar.GetAreasCargo(Me.txtcCarCod.Text)
    
    If rs.EOF And rs.BOF Then
        FlexArea.Clear
        FlexArea.Rows = 2
        FlexArea.FormaCabecera
    Else
        FlexArea.rsFlex = rs
    End If
    
End Sub

Private Sub FlexNivel_EnterCell()
    Me.txtCardesc.Text = Me.FlexNivel.TextMatrix(FlexNivel.Row, 2)
    Me.txtSueldoNivel.value = Me.FlexNivel.TextMatrix(FlexNivel.Row, 3)
    Me.txtSueldoMaximo.value = Me.FlexNivel.TextMatrix(FlexNivel.Row, 4)
    Me.txtOrden.Text = Me.FlexNivel.TextMatrix(FlexNivel.Row, 5)
    UbicaCombo Me.cmbCatCod, Right(Me.FlexNivel.TextMatrix(FlexNivel.Row, 6), 2)
End Sub

Private Sub Form_Load()
    Dim oCons As DConstantes
    Dim rsE As ADODB.Recordset
    Dim oArea As DActualizaDatosArea
    Set oCons = New DConstantes
    Set rsE = New ADODB.Recordset
    Set oArea = New DActualizaDatosArea
    
    If lnTipoOpe = gTipoOpeConsulta Then
        Me.cmdCancelar.Visible = False
        Me.cmdEditar.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdImprimir.Visible = False
    Else
        Set rsE = oCons.GetConstante(6042)
        CargaCombo rsE, Me.cboGrupoPla
        Set rsE = oCons.GetConstante(6043)
        CargaCombo rsE, Me.cboDirInd
        
        Set rsE = oArea.GetAreas
        Me.FlexArea.rsTextBuscar = rsE

    End If
    Refresca_C
    Me.SSTab1.Tab = 0
End Sub

Private Sub FlexNivel_DblClick()
    If lnTipoOpe = gTipoOpeConsulta Then Exit Sub
        
    cmdEditar_Click
End Sub

Private Sub SSTab1_DblClick()
    Refresca_C
End Sub

Private Sub SSTab1_GotFocus()
    Refresca_C
End Sub

Private Sub txtCardesc_GotFocus()
    txtCardesc.SelStart = 0
    txtCardesc.SelLength = 50
End Sub

Private Sub txtCardesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtSueldoNivel.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtCarSue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmdGrabar.SetFocus
End Sub

Private Sub txtcCarDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCarSue.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub ClearScreen(Optional pbBan As Boolean = False)
    Flex.Rows = 1
    Flex.Rows = 2
    Flex.FixedRows = 1
    Flex.ColWidth(0) = 1
    Flex.ColWidth(1) = 4300
    Flex.ColWidth(2) = 1150
    
    Flex.TextMatrix(0, 0) = "Código"
    Flex.TextMatrix(0, 1) = "Descripción"
    

    If pbBan Then
        cmbCatCod = ""
        txtCarSue.value = 0
        txtcCarCod = ""
        txtCardesc = ""
        txtcCarDes = ""
    End If

End Sub

Private Sub txtOrden_GotFocus()
    txtOrden.SelStart = 0
    txtOrden.SelLength = 100
End Sub

Private Sub txtOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdGrabar.Visible Then Me.cmdGrabar.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub Refresca_C()
    Dim oCargo As DActualizadatosCargo
    Dim oConstante As DConstantes
    Dim rsA As ADODB.Recordset
    Set oCargo = New DActualizadatosCargo
    Set oConstante = New DConstantes
    Set rsA = New ADODB.Recordset
    
    Set FlexNivel.DataSource = oCargo.GetCargos(True)
    CargaCombo FlexNivel.DataSource, Me.cmbCargos, 150, 2, 1
    Set rsA = oConstante.GetConstante(gViaticosCateg)
    CargaCombo rsA, cmbCatCod, 150
    
    FlexNivel.ColWidth(0) = 1
    FlexNivel.ColWidth(1) = 700
    FlexNivel.ColWidth(2) = 3200
    FlexNivel.ColWidth(3) = 800
    FlexNivel.ColWidth(4) = 800
    FlexNivel.ColWidth(5) = 600
    FlexNivel.ColWidth(6) = 1500
    FlexNivel.ColAlignment(3) = 7
    FlexNivel.ColAlignment(4) = 7
    FlexNivel.ColAlignment(5) = 7
    
    Set oCargo = Nothing
    Set oConstante = Nothing
    ClearScreen
End Sub

Private Sub Activa(psActiva As Boolean)
    cmdGrabar.Visible = Not psActiva
    cmdCancelar.Visible = Not psActiva
    cmdNuevo.Visible = psActiva
    cmdEditar.Visible = psActiva
    cmdEliminar.Enabled = psActiva
    FlexNivel.Enabled = psActiva
    Flex.Enabled = psActiva
    Me.txtSueldoNivel.Enabled = Not psActiva
    Me.txtSueldoMaximo.Enabled = Not psActiva
    Me.FlexArea.lbEditarFlex = Not psActiva
    Me.Frame4.Enabled = Not psActiva
    Me.txtCarSue.Enabled = Not psActiva
End Sub

Public Sub Ini(pnTipoOpe As TipoOpe, psCaption As String)
    lnTipoOpe = pnTipoOpe
    Caption = psCaption
    Me.Show 1
End Sub
