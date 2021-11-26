VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCajeroGrupoOpe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Grupo de Operaciones de Cajero"
   ClientHeight    =   5610
   ClientLeft      =   1080
   ClientTop       =   2010
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajeroGrupoOpe.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   75
      TabIndex        =   17
      Top             =   30
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   9604
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmCajeroGrupoOpe.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatosGen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Asignar Operaciones"
      TabPicture(1)   =   "frmCajeroGrupoOpe.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Lista de Grupos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   90
         TabIndex        =   20
         Top             =   1980
         Width           =   9015
         Begin VB.CommandButton cmdbuscar 
            Caption         =   "&Buscar"
            Height          =   360
            Left            =   7605
            TabIndex        =   31
            Top             =   1875
            Width           =   1275
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   360
            Left            =   7590
            TabIndex        =   13
            Top             =   1425
            Width           =   1275
         End
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            Height          =   360
            Left            =   7590
            TabIndex        =   14
            Top             =   2805
            Width           =   1275
         End
         Begin MSDataGridLib.DataGrid fgGrupoOpe 
            Height          =   2985
            Left            =   150
            TabIndex        =   0
            Top             =   240
            Width           =   7290
            _ExtentX        =   12859
            _ExtentY        =   5265
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "cGrupoCod"
               Caption         =   "Código"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "cGrupoNombre"
               Caption         =   "Descripción"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "nOrden"
               Caption         =   "Orden"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "cIngEgr"
               Caption         =   "Ingreso"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "nEfectivo"
               Caption         =   "Efectivo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "cOPSuma"
               Caption         =   "Suma"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "cConsDescripcion"
               Caption         =   "Tipo Grupo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2610.142
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   629.858
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   720
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnWidth     =   854.929
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1844.787
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   360
            Left            =   7590
            TabIndex        =   9
            Top             =   330
            Width           =   1275
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "&Grabar"
            Height          =   360
            Left            =   7590
            TabIndex        =   10
            Top             =   330
            Width           =   1275
         End
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Height          =   360
            Left            =   7590
            TabIndex        =   11
            Top             =   690
            Width           =   1275
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   7590
            TabIndex        =   12
            Top             =   690
            Width           =   1275
         End
      End
      Begin VB.Frame fraDatosGen 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   90
         TabIndex        =   19
         Top             =   495
         Width           =   9015
         Begin VB.CheckBox chkSuma 
            Caption         =   "Suma"
            Height          =   285
            Left            =   5145
            TabIndex        =   6
            Top             =   600
            Value           =   1  'Checked
            Width           =   945
         End
         Begin Sicmact.TxtBuscar txtBuscarTipo 
            Height          =   345
            Left            =   1515
            TabIndex        =   7
            Top             =   975
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   609
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
            sTitulo         =   ""
         End
         Begin VB.CheckBox chkEfectivo 
            Caption         =   "Usa Efectivo"
            Height          =   330
            Left            =   3735
            TabIndex        =   5
            Top             =   585
            Value           =   1  'Checked
            Width           =   1260
         End
         Begin VB.CheckBox chkIngreso 
            Caption         =   "Ingreso"
            Height          =   300
            Left            =   2760
            TabIndex        =   4
            Top             =   600
            Value           =   1  'Checked
            Width           =   855
         End
         Begin Spinner.uSpinner spnOrden 
            Height          =   345
            Left            =   1530
            TabIndex        =   3
            Top             =   585
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
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
         Begin VB.TextBox txtNombre 
            Height          =   345
            Left            =   3300
            MaxLength       =   80
            TabIndex        =   2
            Top             =   188
            Width           =   5490
         End
         Begin VB.TextBox txtCodGrupo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   1
            Top             =   188
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Grupo :"
            Height          =   210
            Left            =   120
            TabIndex        =   24
            Top             =   1020
            Width           =   1110
         End
         Begin VB.Label lblTipoDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2475
            TabIndex        =   8
            Top             =   975
            Width           =   5115
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Orden Grupo :"
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   645
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre :"
            Height          =   210
            Left            =   2625
            TabIndex        =   22
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código de Grupo :"
            Height          =   210
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Asignación de Operaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4830
         Left            =   -74910
         TabIndex        =   18
         Top             =   480
         Width           =   9015
         Begin VB.CommandButton cmdElimina 
            Height          =   465
            Left            =   4185
            Picture         =   "frmCajeroGrupoOpe.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2625
            Width           =   540
         End
         Begin VB.CommandButton cmdAgregar 
            Height          =   465
            Left            =   4170
            Picture         =   "frmCajeroGrupoOpe.frx":0F84
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2055
            Width           =   540
         End
         Begin Sicmact.FlexEdit fgOpeAsig 
            Height          =   3750
            Left            =   135
            TabIndex        =   27
            Top             =   885
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   6615
            Cols0           =   3
            HighLight       =   2
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Codigo-Descripción"
            EncabezadosAnchos=   "350-800-2500"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L"
            FormatosEdit    =   "0-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "N°"
            SelectionMode   =   1
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin Sicmact.FlexEdit fgOpeNoAsig 
            Height          =   3750
            Left            =   4770
            TabIndex        =   28
            Top             =   885
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   6615
            Cols0           =   3
            HighLight       =   2
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Codigo-Descripción"
            EncabezadosAnchos=   "350-800-2500"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L"
            FormatosEdit    =   "0-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "N°"
            SelectionMode   =   1
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Operaciones No Asignadas :"
            Height          =   210
            Left            =   4785
            TabIndex        =   30
            Top             =   630
            Width           =   2085
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Operaciones Asignadas :"
            Height          =   210
            Left            =   135
            TabIndex        =   29
            Top             =   645
            Width           =   1845
         End
         Begin VB.Label Label5 
            Caption         =   "Grupo :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   240
            TabIndex        =   26
            Top             =   225
            Width           =   600
         End
         Begin VB.Label lblOpeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   870
            TabIndex        =   25
            Top             =   225
            Width           =   6345
         End
      End
   End
End
Attribute VB_Name = "frmCajeroGrupoOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsOpe As ADODB.Recordset
Attribute rsOpe.VB_VarHelpID = -1
Dim lbNuevo As Boolean
Dim oGrupo As nGrupoOpe
Dim oGen As DGeneral
Dim lsOpeCod As String
Private Sub chkEfectivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkSuma.SetFocus
End If
End Sub
Private Sub chkIngreso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkEfectivo.SetFocus
End If
End Sub

Private Sub chkSuma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBuscarTipo.SetFocus
End If
End Sub

Private Sub cmdAgregar_Click()
If Len(Trim(txtCodGrupo)) = 0 Then
    MsgBox "Codigo de Grupo no Ingresado o no válido", vbInformation, "Aviso"
    SSTab1.Tab = 0
    Exit Sub
End If
If fgOpeNoAsig.TextMatrix(1, 0) = "" Then Exit Sub
If oGrupo.VerificaGruposOpe(txtCodGrupo, fgOpeNoAsig.TextMatrix(fgOpeNoAsig.Row, 1)) Then
    MsgBox "Operación ya se encuentra registrada en grupo seleccionado", vbInformation, "Aviso"
    Exit Sub
End If
oGrupo.GrabaOPeGrupo fgOpeNoAsig.TextMatrix(fgOpeNoAsig.Row, 1), txtCodGrupo
'fgOpeNoAsig.EliminaFila fgOpeNoAsig.Row
Set fgOpeAsig.Recordset = oGrupo.GetOpeGrupo(txtCodGrupo)
End Sub

Private Sub cmdbuscar_Click()
Dim oDescObj As ClassDescObjeto
Set oDescObj = New ClassDescObjeto
oDescObj.BuscarDato rsOpe, 1, "Codigo"

End Sub

Private Sub cmdCancelar_Click()
SSTab1.TabEnabled(1) = True
Habilitar False
txtCodGrupo.Enabled = True
If lbNuevo Then
    CargaDatos
Else
    rsOpe.Find rsOpe(0).Name & "='" & lsOpeCod & "'"
End If
fgGrupoOpe.SetFocus
End Sub
Private Sub cmdEditar_Click()
lsOpeCod = txtCodGrupo
lbNuevo = False
Habilitar True
txtCodGrupo.Enabled = False
End Sub
Private Sub cmdElimina_Click()
If Len(Trim(txtCodGrupo)) = 0 Then
    MsgBox "Codigo de Grupo no Ingresado o no válido", vbInformation, "Aviso"
    SSTab1.Tab = 0
    Exit Sub
End If
If fgOpeAsig.TextMatrix(1, 0) = "" Then Exit Sub
oGrupo.EliminaGruposOpe fgOpeAsig.TextMatrix(fgOpeAsig.Row, 1), txtCodGrupo
fgOpeAsig.EliminaFila fgOpeAsig.Row
Set fgOpeAsig.Recordset = oGrupo.GetOpeGrupo(txtCodGrupo)
'Set fgOpeNoAsig.Recordset = oGrupo.GetOpeGrupo
End Sub

Private Sub cmdEliminar_Click()
If fgOpeAsig.TextMatrix(1, 0) <> "" Then
    MsgBox "Grupo Posee Operaciones Asignadas", vbInformation, "Aviso"
    Exit Sub
End If
If Len(Trim(txtCodGrupo)) = 0 Then
    MsgBox "Codigo de Grupo no válido", vbInformation, "Aviso"
    Exit Sub
End If
If MsgBox("Desea Eliminar el Grupo seleccionado??", vbYesNo + vbInformation, "Aviso") = vbYes Then
    oGrupo.DeleteOpeGrupo Trim(txtCodGrupo)
    CargaDatos
End If
End Sub

Private Sub cmdGrabar_Click()
If Valida = False Then Exit Sub
If MsgBox("Desea Grabar Grupo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsOpeCod = txtCodGrupo
    If lbNuevo Then
        oGrupo.InsertaOpeGrupo txtCodGrupo, txtNombre, Val(spnOrden.Valor), _
         IIf(chkIngreso = 1, "I", "E"), chkEfectivo.Value, txtBuscarTipo, IIf(chkSuma = 1, "S", "N")
    Else
        oGrupo.ActualizaOpeGrupo txtCodGrupo, txtNombre, spnOrden.Valor, _
         IIf(chkIngreso = 1, "I", "E"), chkEfectivo.Value, txtBuscarTipo, IIf(chkSuma = 1, "S", "N")
    End If
    cmdCancelar.Value = True
    rsOpe.Find rsOpe(0).Name & "='" & lsOpeCod & "'"
End If
End Sub
Sub Habilitar(Optional ByVal pbHab As Boolean = True)
cmdNuevo.Visible = Not pbHab
cmdEditar.Visible = Not pbHab
cmdGrabar.Visible = pbHab
cmdCancelar.Visible = pbHab
cmdEliminar.Enabled = Not pbHab
fraDatosGen.Enabled = pbHab
SSTab1.TabEnabled(1) = Not pbHab
fgGrupoOpe.Enabled = Not pbHab
End Sub
Function Valida() As Boolean
Valida = True
If Len(Trim(txtCodGrupo)) = 0 Then
    MsgBox "Código de Grupo no Ingresado o Incompleto", vbInformation, "Aviso"
    Valida = False
    txtCodGrupo.SetFocus
    Exit Function
End If
If Len(Trim(txtNombre)) = 0 Then
    MsgBox "Nombre o Descripción de Grupo no válido", vbInformation, "Aviso"
    Valida = False
    txtNombre.SetFocus
    Exit Function
End If
If Val(spnOrden.Valor) = 0 Then
    MsgBox "Srden de Grupo no válido", vbInformation, "Aviso"
    Valida = False
    spnOrden.SetFocus
    Exit Function
End If
If Len(Trim(txtBuscarTipo)) = 0 Then
    MsgBox "Tipo de Grupo no Ingresado", vbInformation, "Aviso"
    Valida = False
    txtBuscarTipo.SetFocus
    Exit Function
End If


End Function
Private Sub cmdNuevo_Click()
fraDatosGen.Enabled = True
SSTab1.TabEnabled(1) = False
lbNuevo = True
Habilitar True
Limpiar
txtCodGrupo.SetFocus
End Sub
Sub Limpiar()
txtBuscarTipo = ""
lblTipoDesc = ""
txtCodGrupo = ""
txtNombre = ""
chkEfectivo = 1
chkIngreso = 1
chkSuma = 1
spnOrden.Valor = 0
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub fgGrupoOpe_GotFocus()
fgGrupoOpe.MarqueeStyle = dbgHighlightRow
End Sub
Private Sub fgGrupoOpe_LostFocus()
fgGrupoOpe.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub Form_Load()
Set oGrupo = New nGrupoOpe
Set oGen = New DGeneral

lbNuevo = False
CentraForm Me
Set fgOpeNoAsig.Recordset = oGrupo.GetOperaciones
txtBuscarTipo.psRaiz = "Tipo de Grupo"
txtBuscarTipo.rs = oGen.GetConstanteNiv(1031)
CargaDatos
Habilitar False
End Sub
Sub CargaDatos()
Set rsOpe = New ADODB.Recordset
Set rsOpe = oGrupo.GetGrupoOpe
Set fgGrupoOpe.DataSource = rsOpe
End Sub

Private Sub rsOpe_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If Not pRecordset Is Nothing Then
    If Not pRecordset.EOF Then
        'cGrupoCod, cGrupoNombre, nOrden, cIngEgr,nEfectivo,cOPSuma,cConsDescripcion, nTipoGrupo
        txtCodGrupo = pRecordset!cGrupoCod
        txtNombre = pRecordset!cGrupoNombre
        chkEfectivo = pRecordset!nEfectivo
        chkIngreso = IIf(pRecordset!cIngEgr = "I", 1, 0)
        chkSuma = IIf(pRecordset!cOPSuma = "S", 1, 0)
        spnOrden.Valor = pRecordset!nOrden
        txtBuscarTipo = pRecordset!nTipoGrupo
        lblTipoDesc = pRecordset!cConsDescripcion
        lblOpeDesc = pRecordset!cGrupoCod + " - " + pRecordset!cGrupoNombre
        fgOpeAsig.Clear
        fgOpeAsig.FormaCabecera
        fgOpeAsig.Rows = 2
        Set rs = oGrupo.GetOpeGrupo(pRecordset!cGrupoCod)
        If Not rs.EOF And Not rs.BOF Then
            Set fgOpeAsig.Recordset = rs
        End If
    End If
End If
End Sub

Private Sub spnOrden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkIngreso.SetFocus
End If
End Sub

Private Sub txtBuscarTipo_EmiteDatos()
lblTipoDesc = txtBuscarTipo.psDescripcion
End Sub

Private Sub txtBuscarTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub txtCodGrupo_GotFocus()
fEnfoque txtCodGrupo
End Sub

Private Sub txtCodGrupo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtNombre.SetFocus
End If
End Sub
Private Sub txtNombre_GotFocus()
fEnfoque txtNombre
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras(KeyAscii)
If KeyAscii = 13 Then
    spnOrden.SetFocus
End If
End Sub
