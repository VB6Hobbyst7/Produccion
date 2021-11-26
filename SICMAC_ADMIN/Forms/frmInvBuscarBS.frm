VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInvBuscarBS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Bienes"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvBuscarBS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "LISTA DE BIENES"
      Height          =   4575
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   9735
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   8400
         TabIndex        =   14
         Top             =   3960
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   3615
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
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
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "nMovNro"
            Caption         =   "nMovNro"
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
         BeginProperty Column01 
            DataField       =   "cSerie"
            Caption         =   "Cod. Inventario"
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
            DataField       =   "Dep_Historica"
            Caption         =   "Depre"
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
            DataField       =   "cDescripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "vMarca"
            Caption         =   "Marca"
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
            DataField       =   "vModelo"
            Caption         =   "Modelo"
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
            DataField       =   "vSerie"
            Caption         =   "Serie"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "nBSValor"
            Caption         =   "nBSValor"
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
            Size            =   182
            BeginProperty Column00 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3899.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMensaje 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "BUSCAR BIENES:"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.Frame Frame1 
         Caption         =   "Fecha de Depreciacion"
         Height          =   1335
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtPeriodo 
            Height          =   285
            Left            =   3000
            TabIndex        =   18
            Top             =   840
            Width           =   855
         End
         Begin VB.ComboBox cmbMes 
            Height          =   315
            ItemData        =   "frmInvBuscarBS.frx":030A
            Left            =   1680
            List            =   "frmInvBuscarBS.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   2220
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblMes 
            Caption         =   "Mes :"
            Height          =   210
            Left            =   480
            TabIndex        =   19
            Top             =   360
            Width           =   510
         End
      End
      Begin VB.TextBox txtCodInv 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtBien 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "F. Ingreso:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   4455
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   600
            TabIndex        =   3
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   59834369
            CurrentDate     =   39878
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   2760
            TabIndex        =   4
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   59834369
            CurrentDate     =   39878
         End
         Begin VB.Label Label5 
            Caption         =   "A:"
            Height          =   255
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "De:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   8040
         TabIndex        =   1
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Inventario:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Bien:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmInvBuscarBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmInvBuscarBS
'** Descripción : Formulario para la Busqueda de los Activos Fijos
'** Creación : MAVM, 20090218 8:59:25 AM
'** Modificación:
'********************************************************************

Option Explicit

Private Sub Check1_Click()
    If Check1.value = 1 Then
        Frame2.Enabled = True
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    Else
        Frame2.Enabled = False
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    Dim ldFecha As Date
    If cmbMes.Text <> "" Then
        ldFecha = CDate(Trim(Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & "01" & "/" & txtPeriodo.Text))
    Else
        MsgBox "Debe escoger Mes", vbCritical
        Exit Sub
    End If
    
    Dim rs As ADODB.Recordset
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Set rs = oInventario.ObtenerBienes(txtCodInv.Text, IIf(Check1.value = 0, "", Format(DTPicker1.value, "yyyymmdd")), IIf(Check1.value = 0, "", Format(DTPicker2.value, "yyyymmdd")), txtBien.Text, ldFecha)
    
    If rs.RecordCount <> "0" Then
        lblMensaje.Visible = False
        dgBuscar.Visible = True
        Set dgBuscar.DataSource = rs
        dgBuscar.Refresh
        Screen.MousePointer = 0
        dgBuscar.SetFocus
    Else
        Set dgBuscar.DataSource = Nothing
        dgBuscar.Refresh
        lblMensaje.Visible = True
        lblMensaje.Caption = "No Existen Datos"
        dgBuscar.Visible = False
    End If
    
    Set rs = Nothing
    Set oInventario = Nothing
End Sub

Private Sub Command2_Click()
    If frmInvTransferenciaBS.ValidarDatos(dgBuscar.Columns(0).Text, dgBuscar.Columns(1).Text) = False Then
        frmInvTransferenciaBS.ActualizaFG frmInvTransferenciaBS.Index, dgBuscar.Columns(1).Text, dgBuscar.Columns(3).Text, dgBuscar.Columns(0).Text, dgBuscar.Columns(7).Text, dgBuscar.Columns(2).Text
        Unload Me
    Else
        MsgBox "El Activo Fijo ya esta Añadido!", vbCritical
    End If
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    DTPicker1.value = Date
    DTPicker2.value = Date
    txtPeriodo.Text = Format(gdFecSis, "yyyy")
    CargarMes
    Frame2.Enabled = False
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
End Sub

Private Sub CargarMes()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Set rs = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        If IIf(Len(rs.Fields(1)) = 1, "0" & rs.Fields(1), rs.Fields(1)) = Format(gdFecSis, "MM") Then
            cmbMes.Text = rs.Fields(0) & Space(50) & rs.Fields(1)
        End If
        rs.MoveNext
    Wend
End Sub
