VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredNivAprCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Niveles de Aprobacion"
   ClientHeight    =   5880
   ClientLeft      =   1290
   ClientTop       =   1920
   ClientWidth     =   9465
   Icon            =   "frmCredNivAprCred.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   7935
      TabIndex        =   4
      Top             =   5385
      Width           =   1185
   End
   Begin TabDlg.SSTab SSTabs 
      Height          =   5265
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   9287
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Niveles de Aprobacion"
      TabPicture(0)   =   "frmCredNivAprCred.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Administracion"
      TabPicture(1)   =   "frmCredNivAprCred.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Niveles"
      TabPicture(2)   =   "frmCredNivAprCred.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   4065
         Left            =   180
         TabIndex        =   24
         Top             =   480
         Width           =   8835
         Begin VB.PictureBox Picture1 
            Height          =   615
            Left            =   165
            ScaleHeight     =   555
            ScaleWidth      =   8430
            TabIndex        =   28
            Top             =   3300
            Width           =   8490
            Begin VB.CommandButton CmdNivNuevo 
               Caption         =   "&Nuevo"
               Height          =   390
               Left            =   45
               TabIndex        =   33
               Top             =   90
               Width           =   1035
            End
            Begin VB.CommandButton CmdNivEditar 
               Caption         =   "&Editar"
               Height          =   390
               Left            =   1095
               TabIndex        =   32
               Top             =   90
               Width           =   1035
            End
            Begin VB.CommandButton CmdNivEliminar 
               Caption         =   "&Eliminar"
               Height          =   390
               Left            =   2145
               TabIndex        =   31
               Top             =   90
               Width           =   1035
            End
            Begin VB.CommandButton CmdNivAceptar 
               Caption         =   "&Aceptar"
               Height          =   390
               Left            =   6255
               TabIndex        =   30
               Top             =   105
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.CommandButton CmdNivCancelar 
               Caption         =   "&Cancelar"
               Height          =   390
               Left            =   7305
               TabIndex        =   29
               Top             =   105
               Visible         =   0   'False
               Width           =   1035
            End
         End
         Begin MSDataGridLib.DataGrid DGNiveles 
            Height          =   2070
            Left            =   195
            TabIndex        =   25
            Top             =   285
            Width           =   8490
            _ExtentX        =   14975
            _ExtentY        =   3651
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "nNivel"
               Caption         =   "Nivel"
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
            BeginProperty Column02 
               DataField       =   ""
               Caption         =   "codigo"
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
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   5295.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   14.74
               EndProperty
            EndProperty
         End
         Begin VB.TextBox TxtNumNivel 
            Height          =   285
            Left            =   615
            MaxLength       =   1
            TabIndex        =   26
            Top             =   2490
            Width           =   720
         End
         Begin VB.TextBox TxtNivDescripcion 
            Height          =   285
            Left            =   1335
            TabIndex        =   27
            Top             =   2490
            Width           =   4440
         End
         Begin VB.Shape Shape2 
            Height          =   465
            Left            =   180
            Top             =   2415
            Width           =   5880
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3525
         Left            =   -74730
         TabIndex        =   14
         Top             =   1080
         Width           =   7965
         Begin MSComctlLib.TreeView TVNiveles 
            Height          =   3195
            Left            =   120
            TabIndex        =   15
            Top             =   210
            Width           =   7680
            _ExtentX        =   13547
            _ExtentY        =   5636
            _Version        =   393217
            HideSelection   =   0   'False
            LabelEdit       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4650
         Left            =   -74820
         TabIndex        =   5
         Top             =   450
         Width           =   8880
         Begin VB.PictureBox Picture2 
            Height          =   615
            Left            =   120
            ScaleHeight     =   555
            ScaleWidth      =   8520
            TabIndex        =   9
            Top             =   3945
            Width           =   8580
            Begin VB.CommandButton CmdMntCancelar 
               Caption         =   "&Cancelar"
               Height          =   390
               Left            =   7395
               TabIndex        =   13
               Top             =   90
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.CommandButton CmdMntAceptar 
               Caption         =   "&Aceptar"
               Height          =   390
               Left            =   6345
               TabIndex        =   12
               Top             =   90
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.CommandButton CmdMntEliminar 
               Caption         =   "&Eliminar"
               Height          =   390
               Left            =   1095
               TabIndex        =   11
               Top             =   90
               Width           =   1035
            End
            Begin VB.CommandButton CmdMntNuevo 
               Caption         =   "&Nuevo"
               Height          =   390
               Left            =   45
               TabIndex        =   10
               Top             =   90
               Width           =   1035
            End
         End
         Begin MSDataGridLib.DataGrid DGNivel 
            Height          =   2055
            Left            =   165
            TabIndex        =   6
            Top             =   240
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
               DataField       =   "cCodNiv"
               Caption         =   "Codigo"
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
               DataField       =   "cNivel"
               Caption         =   "Nivel"
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
               DataField       =   "cProducto"
               Caption         =   "Producto"
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
               DataField       =   "cTipoCred"
               Caption         =   "Tipo"
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
               DataField       =   "nMontoMin"
               Caption         =   "Monto Inicial"
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
               DataField       =   "nMontoMax"
               Caption         =   "Monto Final"
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
               DataField       =   "cMoneda"
               Caption         =   "Moneda"
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
               MarqueeStyle    =   3
               BeginProperty Column00 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1755.213
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1709.858
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   989.858
               EndProperty
            EndProperty
         End
         Begin VB.ComboBox CmbMoneda 
            Height          =   315
            Left            =   4575
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   2910
            Width           =   1860
         End
         Begin VB.TextBox TxtMontoIni 
            Height          =   285
            Left            =   1215
            TabIndex        =   7
            Top             =   3285
            Width           =   1635
         End
         Begin VB.TextBox TxtMontoFin 
            Height          =   285
            Left            =   4575
            TabIndex        =   8
            Top             =   3285
            Width           =   1545
         End
         Begin VB.ComboBox CmbNiveles 
            Height          =   315
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2490
            Width           =   2010
         End
         Begin VB.ComboBox CmbProducto 
            Height          =   315
            Left            =   4575
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2490
            Width           =   2265
         End
         Begin VB.ComboBox CmbTipo 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2910
            Width           =   1155
         End
         Begin VB.Shape Shape1 
            Height          =   1305
            Left            =   165
            Top             =   2400
            Width           =   8550
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   3510
            TabIndex        =   35
            Top             =   2970
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Monto Max :"
            Height          =   195
            Left            =   3510
            TabIndex        =   23
            Top             =   3330
            Width           =   885
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Monto Min :"
            Height          =   195
            Left            =   315
            TabIndex        =   22
            Top             =   3330
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo :"
            Height          =   195
            Left            =   285
            TabIndex        =   20
            Top             =   2970
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Producto :"
            Height          =   195
            Left            =   3495
            TabIndex        =   18
            Top             =   2550
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nivel :"
            Height          =   195
            Left            =   300
            TabIndex        =   16
            Top             =   2550
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         Height          =   585
         Left            =   -74730
         TabIndex        =   1
         Top             =   495
         Width           =   7980
         Begin VB.ComboBox CmbAnalista 
            Height          =   315
            Left            =   1440
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   195
            Width           =   6435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Administradores :"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   225
            Width           =   1200
         End
      End
   End
End
Attribute VB_Name = "frmCredNivAprCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RNivelApr As ADODB.Recordset
Dim RNivel As ADODB.Recordset
Dim RNivelPers As ADODB.Recordset
Dim CmdEjecutarMnt As Integer
Dim sCadPerm As String

Enum TTipoMuestraNiveles
    MuestraNivelesConsulta = 1
    MuestraNivelesActualizar = 2
End Enum
Dim nTipoMuestra As TTipoMuestraNiveles

'Para el manejo de los Componentes
Dim MatNiveles() As String

Public Sub Inicio(ByVal pTipo As TTipoMuestraNiveles)
    nTipoMuestra = pTipo
    Me.Show 1
End Sub

Private Function ValidaDatosNivel() As Boolean
    ValidaDatosNivel = True
    If Len(Trim(TxtNumNivel.Text)) = 0 Then
        ValidaDatosNivel = False
        TxtNumNivel.SetFocus
        MsgBox "Ingrese Numero de Nivel", vbInformation, "Aviso"
        Exit Function
    End If
    If Len(Trim(TxtNivDescripcion.Text)) = 0 Then
        ValidaDatosNivel = False
        TxtNivDescripcion.SetFocus
        MsgBox "Ingrese la Descripcion del Nivel", vbInformation, "Aviso"
        Exit Function
    End If
    If CmdEjecutarMnt = 1 Then
        RNivel.Find " nNivel = " & Trim(TxtNumNivel.Text)
        If Not RNivel.EOF Then
            ValidaDatosNivel = False
            TxtNivDescripcion.SetFocus
            MsgBox "Nivel ya Existe", vbInformation, "Aviso"
            Exit Function
        End If
    End If
End Function

Private Function ValidaDatosAdmin() As Boolean
    ValidaDatosAdmin = True
    If CmbNiveles.ListIndex = -1 Then
        ValidaDatosAdmin = False
        MsgBox "Seleccione un Nivel", vbInformation, "Aviso"
        CmbNiveles.SetFocus
        Exit Function
    End If
    If CmbProducto.ListIndex = -1 Then
        ValidaDatosAdmin = False
        MsgBox "Seleccione un Producto", vbInformation, "Aviso"
        CmbProducto.SetFocus
        Exit Function
    End If
    If CmbTipo.ListIndex = -1 Then
        ValidaDatosAdmin = False
        MsgBox "Seleccione un Tipo de Credito", vbInformation, "Aviso"
        CmbTipo.SetFocus
        Exit Function
    End If
    If Len(Trim(TxtMontoIni.Text)) = 0 Then
        ValidaDatosAdmin = False
        MsgBox "Ingrese el Monto Minimo", vbInformation, "Aviso"
        TxtMontoIni.SetFocus
        Exit Function
    End If
    If Len(Trim(TxtMontoFin.Text)) = 0 Then
        ValidaDatosAdmin = False
        MsgBox "Ingrese el Monto Maximo", vbInformation, "Aviso"
        TxtMontoFin.SetFocus
        Exit Function
    End If
End Function

Private Function ActualizaNodosPadres(ByRef Nodo As Node) As Boolean
    If Nodo.Children > 0 Then
        If Nodo.Text <> Nodo.LastSibling Then
            If ActualizaNodosPadres(Nodo.Next) Then
                Nodo.Checked = ActualizaNodosPadres(Nodo.Child)
                ActualizaNodosPadres = Nodo.Checked
            Else
                ActualizaNodosPadres = False
                Nodo.Checked = ActualizaNodosPadres(Nodo.Child)
            End If
        Else
            Nodo.Checked = ActualizaNodosPadres(Nodo.Child)
            ActualizaNodosPadres = Nodo.Checked
        End If
    Else
        If Nodo.Text <> Nodo.LastSibling Then
            If ActualizaNodosPadres(Nodo.Next) Then
                ActualizaNodosPadres = Nodo.Checked
            Else
                ActualizaNodosPadres = False
            End If
        Else
            ActualizaNodosPadres = Nodo.Checked
        End If
    End If
    
End Function

Private Sub CargaPermisos(ByVal Nodo As Node)
    If InStr(sCadPerm, Mid(Nodo.Tag, 2, 3) & Mid(Nodo.Tag, 1, 1) & Mid(Nodo.Tag, 5, 1) & Mid(Nodo.Tag, 6, 1)) > 0 Then   'Si es Mayor que Cero Tiene Permiso
        Nodo.Checked = True
    Else
        Nodo.Checked = False
    End If
    If Nodo.Text <> Nodo.LastSibling Then
        Call CargaPermisos(Nodo.Next)
    End If
    If Nodo.Children > 0 Then
        Call CargaPermisos(Nodo.Child)
    End If
End Sub

Private Sub SeleccionaAnalista(ByVal psCodPers As String)
Dim oCredito As COMDCredito.DCOMCreditos

On Error GoTo ErrorSeleccionaAnalista
    Set oCredito = New COMDCredito.DCOMCreditos
    Set RNivelPers = oCredito.RecuperaNivAprPersona(psCodPers)
    Set oCredito = Nothing
    sCadPerm = ""
    Do While Not RNivelPers.EOF
        sCadPerm = sCadPerm & RNivelPers!cCodNiv & ","
        RNivelPers.MoveNext
    Loop
    RNivelPers.Close
    If Not TVNiveles.Nodes(1).Child Is Nothing Then
        Call CargaPermisos(TVNiveles.Nodes(1).Child)
    End If
    
    Exit Sub
    
ErrorSeleccionaAnalista:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub PermisoNodo(ByVal Nodo As Node, ByVal pbChecked As Boolean, ByVal pbUltNivel As Boolean, ByVal NodoName As String)
                        
'Dim oCredito As COMDCredito.DCOMCreditos
Dim nNumNiveles As Integer
Dim sNodo As String
Dim sAnalista As String
    
    If Nodo.Children > 0 Then
        Call PermisoNodo(Nodo.Child, pbChecked, pbUltNivel, NodoName)
        If NodoName = Nodo.Text Then
            Exit Sub
        End If
        If Nodo.Text <> Nodo.LastSibling Then
            Call PermisoNodo(Nodo.Next, pbChecked, pbUltNivel, NodoName)
        End If
        Exit Sub
    End If
    
    If Trim(Nodo.Tag) <> "X" Then
        'Set oCredito = New COMDCredito.DCOMCreditos
        nNumNiveles = UBound(MatNiveles)
        sNodo = Mid(Nodo.Tag, 2, 3) & Mid(Nodo.Tag, 1, 1) & Mid(Nodo.Tag, 5, 1) & Mid(Nodo.Tag, 6, 1)
         
        If pbChecked Then
            nNumNiveles = nNumNiveles + 2
            ReDim Preserve MatNiveles(nNumNiveles)
            MatNiveles(nNumNiveles - 2) = "E" & sNodo
            MatNiveles(nNumNiveles - 1) = "N" & sNodo
            'Call oCredito.EliminarPermisoNivelApr(Mid(Nodo.Tag, 2, 3) & Mid(Nodo.Tag, 1, 1) & Mid(Nodo.Tag, 5, 1) & Mid(Nodo.Tag, 6, 1), Trim(Right(CmbAnalista.Text, 15)))
            'Call oCredito.NuevoPermisoNivelApr(Mid(Nodo.Tag, 2, 3) & Mid(Nodo.Tag, 1, 1) & Mid(Nodo.Tag, 5, 1) & Mid(Nodo.Tag, 6, 1), Trim(Right(CmbAnalista.Text, 15)))
            Nodo.Checked = True
        Else
            nNumNiveles = nNumNiveles + 1
            ReDim Preserve MatNiveles(nNumNiveles)
            MatNiveles(nNumNiveles - 1) = "E" & sNodo
            'Call oCredito.EliminarPermisoNivelApr(Mid(Nodo.Tag, 2, 3) & Mid(Nodo.Tag, 1, 1) & Mid(Nodo.Tag, 5, 1) & Mid(Nodo.Tag, 6, 1), Trim(Right(CmbAnalista.Text, 15)))
            Nodo.Checked = False
        End If
        If pbUltNivel Then
            Exit Sub
        End If
    End If
    If Nodo.Text <> Nodo.LastSibling Then
        Call PermisoNodo(Nodo.Next, pbChecked, pbUltNivel, NodoName)
    End If
    
End Sub

Private Sub HabilitaIngresoNiveles(ByVal pbHabilita As Boolean)
    DGNiveles.Enabled = Not pbHabilita
    TxtNumNivel.Enabled = pbHabilita
    If RNivel.RecordCount > 0 And pbHabilita = True Then
        'RNivel.MoveLast
        'TxtNumNivel.Text = RNivel!nNivel + 1
        TxtNumNivel.Text = RNivel.RecordCount + 1
        
    Else
        TxtNumNivel.Text = ""
    End If
    TxtNivDescripcion.Enabled = pbHabilita
    TxtNivDescripcion.Text = ""
    CmdNivNuevo.Visible = Not pbHabilita
    CmdNivEditar.Visible = Not pbHabilita
    CmdNivEliminar.Visible = Not pbHabilita
    CmdNivAceptar.Visible = pbHabilita
    CmdNivCancelar.Visible = pbHabilita
    DGNiveles.Height = IIf(pbHabilita, 2070, 2610)
    
End Sub

Private Sub HabilitaMantenimiento(ByVal pbHabilita As Boolean)
    
    DGNivel.Enabled = Not pbHabilita
    TxtMontoIni.Enabled = pbHabilita
    TxtMontoIni.Text = ""
    TxtMontoFin.Enabled = pbHabilita
    TxtMontoFin.Text = ""
    CmbNiveles.ListIndex = -1
    CmbProducto.ListIndex = -1
    CmbTipo.ListIndex = -1
    CmbMoneda.ListIndex = -1
    CmdMntNuevo.Enabled = Not pbHabilita
    CmdMntNuevo.Visible = Not pbHabilita
    CmdMntEliminar.Enabled = Not pbHabilita
    CmdMntEliminar.Visible = Not pbHabilita
    CmdMntAceptar.Enabled = pbHabilita
    CmdMntAceptar.Visible = pbHabilita
    CmdMntCancelar.Enabled = pbHabilita
    CmdMntCancelar.Visible = pbHabilita
    DGNivel.Height = IIf(pbHabilita, 2040, 3525)
End Sub

Private Sub CargaDatos()
'    Call CargaAnalistas
'    Call CargaArbol
'    Call CargaNiveles
'    Call CargaAdministracion

    Call Cargar_Datos_Generales
    
End Sub

Private Sub Cargar_Datos_Generales()

'Los sgtes Recordset son Globales
'RNivelApr
'RNivel
'RNivelPers

Dim oCreditos As COMDCredito.DCOMCreditos

'Cargar Analista
Dim rsAnalista As ADODB.Recordset
Dim sApoderados As String

'Cargar Arbol
Dim N As Node
Dim sNivTmp As String
Dim sProd As String
Dim sMoneda As String
Dim sTipoCred As String

'Cargar Administracion
Dim rsNivAdm As ADODB.Recordset
Dim rsProducto As ADODB.Recordset
Dim rsTipoCred As ADODB.Recordset
Dim rsMoneda As ADODB.Recordset

On Error GoTo ERRORCargar_Datos_Generales
        
    Set oCreditos = New COMDCredito.DCOMCreditos
    
    Set RNivel = Nothing
    Set RNivelApr = Nothing
    
    Call oCreditos.Cargar_Datos_Generales(sApoderados, rsAnalista, RNivelApr, RNivel, _
                                        rsNivAdm, rsProducto, rsTipoCred, rsMoneda)
            
    CmbAnalista.Clear
    Do While Not rsAnalista.EOF
        CmbAnalista.AddItem PstaNombre(rsAnalista!cPersNombre, False) & Space(100) & rsAnalista!cPersCod
        rsAnalista.MoveNext
    Loop
    
'Carga Arbol

    TVNiveles.Nodes.Clear
    TVNiveles.Nodes.Add , , "M", "Niveles"
    TVNiveles.Nodes(1).Tag = "X"
    sNivTmp = ""
    sProd = ""
    sMoneda = ""
    Do While Not RNivelApr.EOF
        If sNivTmp <> Mid(RNivelApr!cCodNiv, 1, 1) Then
            sNivTmp = Mid(RNivelApr!cCodNiv, 1, 1)
            Set N = TVNiveles.Nodes.Add("M", tvwChild, "N" & sNivTmp, Trim(Left(RNivelApr!cNivel, 40)))
            'N.Tag = sNivTmp
            N.Tag = "X"
        End If
        If sProd <> Mid(RNivelApr!cCodNiv, 1, 4) Then
            sProd = Mid(RNivelApr!cCodNiv, 1, 4)
            Set N = TVNiveles.Nodes.Add("N" & sNivTmp, tvwChild, "P" & sProd, Trim(Left(RNivelApr!cProducto, 40)))
            'N.Tag = sProd
            N.Tag = "X"
        End If
        If sTipoCred <> Mid(RNivelApr!cCodNiv, 1, 5) Then
            sTipoCred = Mid(RNivelApr!cCodNiv, 1, 5)
            Set N = TVNiveles.Nodes.Add("P" & sProd, tvwChild, "V" & sTipoCred, Trim(Left(RNivelApr!cTipoCred, 40)))
            'N.Tag = sProd
            N.Tag = "X"
        End If
        Set N = TVNiveles.Nodes.Add("V" & sTipoCred, tvwChild, "S" & RNivelApr!cCodNiv, Trim(Left(RNivelApr!cMoneda, 40)))
        N.Tag = RNivelApr!cCodNiv
        RNivelApr.MoveNext
    Loop
    
    If TVNiveles.Nodes.Count > 0 Then
        TVNiveles.Nodes(1).Expanded = True
    End If
    
    If TVNiveles.Nodes.Count = 0 Then
        TVNiveles.Enabled = False
    Else
        TVNiveles.Enabled = True
    End If

    
'Cargar Niveles
    Set DGNiveles.DataSource = RNivel
    DGNiveles.Refresh
    
'Cargar Administracion

    'Carga Grid
    Set DGNivel.DataSource = RNivelApr
    DGNivel.Refresh
    
    'Carga Combo Nivels
    Call CambiaTamañoCombo(CmbNiveles)
    
    CmbNiveles.Clear
    Do While Not rsNivAdm.EOF
        CmbNiveles.AddItem rsNivAdm!cDescripcion & Space(50) & Trim(Str(rsNivAdm!nNivel))
        rsNivAdm.MoveNext
    Loop
        
    'Carga Producto
    Call CambiaTamañoCombo(CmbProducto)
    
    CmbProducto.Clear
    Do While Not rsProducto.EOF
        CmbProducto.AddItem rsProducto!cConsDescripcion & Space(50) & rsProducto!nConsValor
        rsProducto.MoveNext
    Loop
    
    'Carga Tipos de Creditos
    Call CambiaTamañoCombo(CmbTipo)
    Call Llenar_Combo_con_Recordset(rsTipoCred, CmbTipo)
    
    'Carga Moneda
    Call Llenar_Combo_con_Recordset(rsMoneda, CmbMoneda)
    
    Set rsAnalista = Nothing
    Set rsNivAdm = Nothing
    Set rsProducto = Nothing
    Set rsTipoCred = Nothing
    Set rsMoneda = Nothing
    
    Set oCreditos = Nothing
    
    Exit Sub
    
ERRORCargar_Datos_Generales:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub CmbAnalista_Click()
    Call SeleccionaAnalista(Trim(Right(CmbAnalista.Text, 15)))
    TVNiveles.Nodes(1).Checked = ActualizaNodosPadres(TVNiveles.Nodes(1))
End Sub

Private Sub CmbNiveles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbProducto.SetFocus
    End If
End Sub

Private Sub CmbProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbTipo.SetFocus
    End If
End Sub

Private Sub CmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoIni.SetFocus
    End If
End Sub

Private Sub CmdMntAceptar_Click()
Dim oCredito As COMDCredito.DCOMCreditos
Dim sTmpCodNiv As String
    If Not ValidaDatosAdmin Then
        Exit Sub
    End If
    Set oCredito = New COMDCredito.DCOMCreditos
    sTmpCodNiv = Trim(Right(CmbProducto.Text, 5)) & Trim(Right(CmbNiveles.Text, 3)) & Trim(Right(CmbTipo.Text, 3)) & Trim(Right(CmbMoneda.Text, 2))
    If oCredito.ExisteNivelAprobacion(sTmpCodNiv, Trim(Right(CmbProducto.Text, 10)), Trim(Right(CmbTipo.Text, 10)), CInt(Trim(Right(CmbNiveles.Text, 10))), CInt(Trim(Right(CmbMoneda.Text, 2)))) Then
        MsgBox "Nivel de Aprobacion ya Existe", vbInformation, "Aviso"
        Set oCredito = Nothing
        Exit Sub
    End If
    Call oCredito.NuevoNivelAprobacion(sTmpCodNiv, Trim(Right(CmbProducto.Text, 10)), Trim(Right(CmbTipo.Text, 10)), Trim(Right(CmbNiveles.Text, 10)), CDbl(TxtMontoFin.Text), CDbl(TxtMontoIni.Text), CInt(Trim(Right(CmbMoneda.Text, 2))))
    Set oCredito = Nothing
    Call HabilitaMantenimiento(False)
    Call CargaDatos
End Sub

Private Sub CmdMntCancelar_Click()
    Call HabilitaMantenimiento(False)
End Sub


Private Sub CmdMntEliminar_Click()
Dim oCredito As COMDCredito.DCOMCreditos
    If MsgBox("Se va a Eliminar el Registro, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oCredito = New COMDCredito.DCOMCreditos
        Call oCredito.EliminaNivelAprobacion(Mid(RNivelApr!cCodNiv, 2, 3) & Mid(RNivelApr!cCodNiv, 1, 1) & Mid(RNivelApr!cCodNiv, 5, 1) & Mid(RNivelApr!cCodNiv, 6, 1))
        Set oCredito = Nothing
        Call CargaDatos
    End If
End Sub

Private Sub CmdMntNuevo_Click()
    Call HabilitaMantenimiento(True)
    CmbNiveles.SetFocus
End Sub

Private Sub CmdNivAceptar_Click()
Dim oCredito As COMDCredito.DCOMCreditos
    On Error GoTo ERRORCmdNivAceptar_Click
    If Not ValidaDatosNivel Then
        Exit Sub
    End If
    Set oCredito = New COMDCredito.DCOMCreditos
    If CmdEjecutarMnt = 1 Then
        Call oCredito.Nuevonivel(CInt(Trim(TxtNumNivel.Text)), Trim(TxtNivDescripcion.Text))
    Else
        Call oCredito.ActualizaNivel(CInt(Trim(TxtNumNivel.Text)), Trim(TxtNivDescripcion.Text))
    End If
    Set oCredito = Nothing
    
    Call CargaDatos
    Call HabilitaIngresoNiveles(False)
    CmdEjecutarMnt = -1
    Exit Sub
    
ERRORCmdNivAceptar_Click:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdNivCancelar_Click()
    Call HabilitaIngresoNiveles(False)
    CmdEjecutarMnt = -1
End Sub

Private Sub CmdNivEditar_Click()
    CmdEjecutarMnt = 2
    Call HabilitaIngresoNiveles(True)
    TxtNumNivel.Text = RNivel!nNivel
    TxtNumNivel.Enabled = False
    TxtNivDescripcion.Text = RNivel!cDescripcion
    TxtNivDescripcion.SetFocus
End Sub

Private Sub CmdNivEliminar_Click()

Dim oCredito As COMDCredito.DCOMCreditos
    
    Set oCredito = New COMDCredito.DCOMCreditos
    If oCredito.TieneNivelAprobAsignado(Trim(Str(RNivel!nNivel))) Then
        MsgBox "No se puede Eliminar este Nivel esta siendo utilizado", vbInformation, "Aviso"
        Set oCredito = Nothing
        Exit Sub
    End If
    
    If MsgBox("Se Va eliminar el Nivel y Todos los Permisos Relacionados con Este, Desea Contnuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Call oCredito.EliminaNivel(RNivel!nNivel)
        Call CargaDatos
    End If
    Set oCredito = Nothing
End Sub

Private Sub CmdNivNuevo_Click()
    CmdEjecutarMnt = 1
    Call HabilitaIngresoNiveles(True)
    TxtNumNivel.SetFocus
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DGNivel.Height = 3525
    DGNiveles.Height = 2610
    Call CargaDatos
    If CmbAnalista.ListCount > 0 Then
        CmbAnalista.ListIndex = 0
    End If
    CmdEjecutarMnt = -1
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub


Private Sub SSTabs_Click(PreviousTab As Integer)
    If nTipoMuestra = MuestraNivelesConsulta Then
        SSTabs.Tab = 0
    End If
End Sub

Private Sub TVNiveles_NodeCheck(ByVal Node As MSComctlLib.Node)

Dim oCredito As COMDCredito.DCOMCreditos

    If nTipoMuestra = MuestraNivelesConsulta Then
        MsgBox "Solo Puede Consultar la Pantalla", vbInformation, "Aviso"
        Exit Sub
    End If
    If CmbAnalista.ListIndex <> -1 Then
        ReDim MatNiveles(0)
        Call PermisoNodo(Node, Node.Checked, IIf(Trim(Node.Tag) <> "X", True, False), Node.Text)
        Set oCredito = New COMDCredito.DCOMCreditos
        Call oCredito.ActualizaNivelesEnLote(MatNiveles, Trim(Right(CmbAnalista.Text, 15)))
        Set oCredito = Nothing
        TVNiveles.Nodes(1).Checked = ActualizaNodosPadres(TVNiveles.Nodes(1))
    Else
        MsgBox "Seleccione un Analista", vbInformation, "Aviso"
        Node.Checked = False
    End If
    
End Sub

Private Sub TxtMontoFin_GotFocus()
    fEnfoque TxtMontoFin
End Sub

Private Sub TxtMontoFin_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMontoFin, KeyAscii)
    If KeyAscii = 13 Then
        CmdMntAceptar.SetFocus
    End If
End Sub

Private Sub TxtMontoFin_LostFocus()
    TxtMontoFin.Text = Format(IIf(Trim(TxtMontoFin.Text) = "", 0, TxtMontoFin.Text), "#0.00")
End Sub

Private Sub TxtMontoIni_GotFocus()
    fEnfoque TxtMontoIni
End Sub

Private Sub TxtMontoIni_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMontoIni, KeyAscii)
    If KeyAscii = 13 Then
        TxtMontoFin.SetFocus
    End If
End Sub

Private Sub TxtMontoIni_LostFocus()
    TxtMontoIni.Text = Format(IIf(Trim(TxtMontoIni.Text) = "", 0, TxtMontoIni.Text), "#0.00")
End Sub

Private Sub TxtNivDescripcion_GotFocus()
    fEnfoque TxtNivDescripcion
End Sub

Private Sub TxtNivDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        CmdNivAceptar.SetFocus
    End If
End Sub

Private Sub TxtNumNivel_GotFocus()
    fEnfoque TxtNumNivel
End Sub

Private Sub TxtNumNivel_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtNivDescripcion.SetFocus
    End If
End Sub

'******************* Rutinas anteriores de Cargado de Datos

'Private Sub CargaAnalistas()
'Dim R As ADODB.Recordset
'Dim sSQL As String, sApoderados As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim oGen As DGeneral
'    On Error GoTo ERRORCargaAnalistas
'
'    Set oGen = New DGeneral
'    sApoderados = oGen.LeeConstSistema(gConstSistRHCargoCodApoderados)
'    Set oGen = Nothing
'
'    sSQL = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
'    sSQL = sSQL & " AND nRHEstado = 201 "
'    sSQL = sSQL & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
'    sSQL = sSQL & " where  RC.cRHCargoCod in (" & sApoderados & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod)"
'    sSQL = sSQL & " order by P.cPersNombre "
'
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set R = oConecta.CargaRecordSet(sSQL)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    cmbAnalista.Clear
'    Do While Not R.EOF
'        cmbAnalista.AddItem PstaNombre(R!cPersNombre, False) & Space(100) & R!cPersCod
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    Exit Sub
'ERRORCargaAnalistas:
'    MsgBox Err.Description, vbCritical, "Aviso"
'End Sub

'Private Sub CargaArbol()
'Dim RNiv As ADODB.Recordset
'Dim oCreditos As COMDCredito.DCOMCreditos
'Dim N As Node
'Dim sNivTmp As String
'Dim sProd As String
'Dim sMoneda As String
'Dim sTipoCred As String
'
'    On Error GoTo ERRORCargaArbol
'
'    Set oCreditos = New COMDCredito.DCOMCreditos
'    Set RNiv = oCreditos.RecuperaNivelesAprobacion
'    TVNiveles.Nodes.Clear
'    TVNiveles.Nodes.Add , , "M", "Niveles"
'    TVNiveles.Nodes(1).Tag = "X"
'    sNivTmp = ""
'    sProd = ""
'    sMoneda = ""
'    Do While Not RNiv.EOF
'        If sNivTmp <> Mid(RNiv!cCodNiv, 1, 1) Then
'            sNivTmp = Mid(RNiv!cCodNiv, 1, 1)
'            Set N = TVNiveles.Nodes.Add("M", tvwChild, "N" & sNivTmp, Trim(Left(RNiv!cNivel, 40)))
'            'N.Tag = sNivTmp
'            N.Tag = "X"
'        End If
'        If sProd <> Mid(RNiv!cCodNiv, 1, 4) Then
'            sProd = Mid(RNiv!cCodNiv, 1, 4)
'            Set N = TVNiveles.Nodes.Add("N" & sNivTmp, tvwChild, "P" & sProd, Trim(Left(RNiv!cProducto, 40)))
'            'N.Tag = sProd
'            N.Tag = "X"
'        End If
'        If sTipoCred <> Mid(RNiv!cCodNiv, 1, 5) Then
'            sTipoCred = Mid(RNiv!cCodNiv, 1, 5)
'            Set N = TVNiveles.Nodes.Add("P" & sProd, tvwChild, "V" & sTipoCred, Trim(Left(RNiv!cTipoCred, 40)))
'            'N.Tag = sProd
'            N.Tag = "X"
'        End If
'        Set N = TVNiveles.Nodes.Add("V" & sTipoCred, tvwChild, "S" & RNiv!cCodNiv, Trim(Left(RNiv!cMoneda, 40)))
'        N.Tag = RNiv!cCodNiv
'        RNiv.MoveNext
'    Loop
'    RNiv.Close
'    Set RNiv = Nothing
'    If TVNiveles.Nodes.Count > 0 Then
'        TVNiveles.Nodes(1).Expanded = True
'    End If
'
'    If TVNiveles.Nodes.Count = 0 Then
'        TVNiveles.Enabled = False
'    Else
'        TVNiveles.Enabled = True
'    End If
'    Exit Sub
'
'ERRORCargaArbol:
'    MsgBox Err.Description, vbCritical, "Aviso"
'End Sub

'Private Sub CargaNiveles()
'Dim oCredito As COMDCredito.DCOMCreditos
'
'    On Error GoTo ERRORCargaNiveles
'    Set oCredito = New COMDCredito.DCOMCreditos
'    Set RNivel = Nothing
'    Set RNivel = oCredito.RecuperaNiveles
'    Set DGNiveles.DataSource = RNivel
'    DGNiveles.Refresh
'    Set oCredito = Nothing
'    Exit Sub
'
'ERRORCargaNiveles:
'    MsgBox Err.Description, vbCritical, "Aviso"
'End Sub

'Private Sub CargaAdministracion()
'Dim oCredito As COMDCredito.DCOMCreditos
''Dim oConecta As COMConecta.DCOMConecta
'Dim sSQL As String
''Dim R As ADODB.Recordset
'
'Dim rsNiveles As ADODB.Recordset
'Dim rsProducto As ADODB.Recordset
'
'On Error GoTo ERRORCargaAdministracion
'
'    'Carga Grid
'    Set oCredito = New COMDCredito.DCOMCreditos
'    Set RNivelApr = oCredito.RecuperaNivelesAprobacion
'    Set DGNivel.DataSource = RNivelApr
'    DGNivel.Refresh
'    Set oCredito = Nothing
'
'    'Carga Combo Nivels
'    Call CambiaTamañoCombo(CmbNiveles)
''    sSQL = "Select nNivel,cDescripcion from ColocCredNivelesTipos"
''    Set oConecta = New COMConecta.DCOMConecta
''    oConecta.AbreConexion
''    Set R = oConecta.CargaRecordSet(sSQL)
''    oConecta.CierraConexion
''    Set oConecta = Nothing
'    CmbNiveles.Clear
'    Do While Not R.EOF
'        CmbNiveles.AddItem R!cDescripcion & Space(50) & Trim(Str(R!nNivel))
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    'Carga Producto
'    Call CambiaTamañoCombo(CmbProducto)
''    sSQL = "select nConsValor, cConsDescripcion from Constante where convert(varchar(15),nConsValor) not like '23_' AND nConsValor <> 305 AND nConsValor <> nConsCod AND nConsCod = " & gProducto
''    Set oConecta = New COMConecta.DCOMConecta
''    oConecta.AbreConexion
''    Set R = oConecta.CargaRecordSet(sSQL)
''    oConecta.CierraConexion
''    Set oConecta = Nothing
'
'    CmbProducto.Clear
'    Do While Not R.EOF
'        CmbProducto.AddItem R!cConsDescripcion & Space(50) & R!nConsValor
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    'Carga Tipos de Creditos
'    Call CambiaTamañoCombo(CmbTipo)
'    Call CargaComboConstante(gColocCredTipo, CmbTipo)
'
'    'Carga Moneda
'    Call CargaComboConstante(gMoneda, CmbMoneda)
'
'    Exit Sub
'
'ERRORCargaAdministracion:
'    MsgBox Err.Description, vbCritical, "Aviso"
'End Sub

