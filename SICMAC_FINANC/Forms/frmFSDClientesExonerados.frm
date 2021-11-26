VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFSDClientesExonerados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LISTADO DE CLIENTES EXONERADOS"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14580
   Icon            =   "frmFSDClientesExonerados.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Listado de Registros Encontrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3825
      Left            =   45
      TabIndex        =   14
      Top             =   1800
      Width           =   14490
      Begin MSDataGridLib.DataGrid dgExonerados 
         Height          =   3045
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   14250
         _ExtentX        =   25135
         _ExtentY        =   5371
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "cPersTipExon"
            Caption         =   "Tipo Exon."
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
            DataField       =   "cPersCod"
            Caption         =   "cCodPers"
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
            DataField       =   "cPersCod"
            Caption         =   "Código"
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
            DataField       =   "cPersNombre"
            Caption         =   "Cliente"
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
            DataField       =   "cPersNotas"
            Caption         =   "Nota"
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
            DataField       =   "cUsuario"
            Caption         =   "Usuario"
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
            DataField       =   "cFecha"
            Caption         =   "Fecha"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   345.26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3539.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3585.26
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCantidad 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   150
         TabIndex        =   16
         Top             =   3345
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Registro(s) Encontrado(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1335
         TabIndex        =   15
         Top             =   3375
         Width           =   2130
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1650
      Left            =   8760
      TabIndex        =   12
      Top             =   120
      Width           =   1635
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
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
         Height          =   390
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1365
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
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
         Height          =   390
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1650
      Left            =   960
      TabIndex        =   6
      Top             =   90
      Width           =   7770
      Begin VB.TextBox txtNotas 
         Height          =   315
         Left            =   1560
         MaxLength       =   49
         TabIndex        =   2
         Top             =   1200
         Width           =   5295
      End
      Begin VB.ComboBox comTipo 
         Height          =   315
         ItemData        =   "frmFSDClientesExonerados.frx":030A
         Left            =   1575
         List            =   "frmFSDClientesExonerados.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   810
         Width           =   5280
      End
      Begin VB.CommandButton CmdBuscar 
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
         Left            =   6480
         TabIndex        =   0
         Top             =   285
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nota"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Exoneración"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   825
         Width           =   1245
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7680
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   435
         Width           =   570
      End
      Begin VB.Label lblNomPers 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   3780
      End
      Begin VB.Label LblPersCod 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   870
         TabIndex        =   7
         Top             =   375
         Width           =   1755
      End
   End
   Begin MSComCtl2.Animation Logo 
      Height          =   645
      Left            =   135
      TabIndex        =   13
      Top             =   180
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1138
      _Version        =   393216
      FullWidth       =   45
      FullHeight      =   43
   End
End
Attribute VB_Name = "frmFSDClientesExonerados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAgregar_Click()
Dim oExonerados As NAnx_FSD
'***Agregado por ELRO el 20130507, según TI-ERS019-2013****
Dim oNContFunciones As NContFunciones
Dim lsMovNro As String
'***Fin Agregado por ELRO el 20130507, según TI-ERS019-2013

On Error GoTo ErrorAgregar
If Len(Trim(LblPersCod.Caption)) = 0 Then
    MsgBox "Seleccione un código para procesar", vbExclamation, "Aviso!!!"
    CmdBuscar.SetFocus
    Exit Sub
Else
    If Len(Trim(comTipo.Text)) = 0 Then
        MsgBox "Seleccione un tipo de exoneración", vbExclamation, "Aviso!!!"
        comTipo.SetFocus
        Exit Sub
    End If
End If

'***Agregado por ELRO el 20130507, según TI-ERS019-2013****
If Len(Trim(txtNotas)) = 0 Then
    MsgBox "Ingrese la nota", vbExclamation, "Aviso!!!"
    txtNotas.SetFocus
    Exit Sub
End If
Set oNContFunciones = New NContFunciones
lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'***Fin Agregado por ELRO el 20130507, según TI-ERS019-2013

Set oExonerados = New NAnx_FSD

If oExonerados.GetExisteExonerado(Trim(LblPersCod.Caption)) = True Then
    MsgBox "El cliente ya esta exonerado", vbExclamation, "Aviso!!!"
    Exit Sub
Else
    oExonerados.InsertaNuevoExonerado Trim(LblPersCod.Caption), Val(Right(comTipo.Text, 5)), txtNotas.Text, lsMovNro
End If
Listar
MsgBox "Registros de exonerados actualizados satisfactoriamente", vbInformation, "Aviso!!!"
Blanquea
cmdAgregar.Enabled = False
Set oNContFunciones = Nothing '***Agregado por ELRO el 20130507, según TI-ERS019-2013
Exit Sub

ErrorAgregar:
    MsgBox Err.Description, vbExclamation, "Aviso!!!"
End Sub

Private Sub Blanquea()

    LblPersCod.Caption = ""
    lblNomPers.Caption = ""
    comTipo.ListIndex = -1
    txtNotas.Text = ""

End Sub

Private Sub cmdQuitar_Click()
Dim oExonerados As NAnx_FSD
'***Agregado por ELRO el 20130507, según TI-ERS019-2013****
Dim oNContFunciones As NContFunciones
Dim lsMovNro As String
'***Fin Agregado por ELRO el 20130507, según TI-ERS019-2013

On Error GoTo ErrorAgregar

If Val(lblCantidad.Caption) = 0 Then
    MsgBox "No existen registros que retirar", vbExclamation, "Aviso!!!"
    CmdBuscar.SetFocus
    Exit Sub
End If

If MsgBox("Desea retirar el registro seleccionado", vbQuestion + vbYesNo, "Aviso!!!") = vbYes Then
        
    Set oExonerados = New NAnx_FSD
    
    '***Agregado por ELRO el 20130507, según TI-ERS019-2013****
    Set oNContFunciones = New NContFunciones
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    '***Fin Agregado por ELRO el 20130507, según TI-ERS019-2013
    
    oExonerados.EliminaExonerado Trim(dgExonerados.Columns(1).Text), lsMovNro
    
    Listar
    MsgBox "Registros de exonerados actualizados satisfactoriamente", vbInformation, "Aviso!!!"
    Blanquea
    cmdAgregar.Enabled = False
    Set oNContFunciones = Nothing '***Agregado por ELRO el 20130507, según TI-ERS019-2013

End If
Exit Sub

ErrorAgregar:
    MsgBox Err.Description, vbExclamation, "Aviso!!!"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub comTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNotas.SetFocus
End If
End Sub

Private Sub Form_Load()

CentraForm Me
frmFondoSeguroDep.Enabled = False
 
Logo.AutoPlay = True
Logo.Open App.path & "\videos\LogoA.avi"
LlenaCmb
Listar
End Sub

Private Sub LlenaCmb()
Dim oCon As DConstantes
Dim rs As ADODB.Recordset
Dim sCad As String
On Error GoTo ErrLlena_

Set oCon = New DConstantes
Set rs = oCon.CargaConstante(gFSDTipoExonerados)
Do While Not rs.EOF
    comTipo.AddItem rs!cConsDescripcion & Space(100) & rs!nConsValor
    rs.MoveNext
Loop
Set rs = Nothing
Set oCon = Nothing
Exit Sub

ErrLlena_:
    MsgBox Err.Description, vbInformation, "Aviso!!!"

End Sub

Private Sub cmdbuscar_Click()
Dim oPersona As UPersona
Dim sPersCod As String

    
    Set oPersona = frmBuscaPersona.Inicio
    If Not oPersona Is Nothing Then
        If gbBitCentral Then
            LblPersCod.Caption = oPersona.sPersCod
        Else
            LblPersCod.Caption = Mid(oPersona.sPersCod, 4, 10)
        End If
        lblNomPers.Caption = oPersona.sPersNombre
        comTipo.ListIndex = -1
        txtNotas.Text = ""
    Else
        Exit Sub
    End If
    sPersCod = oPersona.sPersCod
    Set oPersona = Nothing
    
    
    If sPersCod <> "" Then
        cmdAgregar.Enabled = True
        comTipo.SetFocus
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFondoSeguroDep.Enabled = True
 
End Sub

Private Sub Listar()
Dim oExonerados As NAnx_FSD
Dim rs As ADODB.Recordset
Set oExonerados = New NAnx_FSD
Set rs = New ADODB.Recordset

Set rs = oExonerados.GetListaExonerados
Set dgExonerados.DataSource = rs
lblCantidad.Caption = rs.RecordCount
If Val(lblCantidad.Caption) = 0 Then
    cmdQuitar.Enabled = False
Else
    cmdQuitar.Enabled = True
End If
Set rs.ActiveConnection = Nothing
Set oExonerados = Nothing

End Sub

Private Sub txtNotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAgregar.SetFocus
End If
End Sub
