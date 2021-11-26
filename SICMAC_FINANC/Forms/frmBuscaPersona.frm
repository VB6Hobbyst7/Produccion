VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Persona"
   ClientHeight    =   2775
   ClientLeft      =   2370
   ClientTop       =   2535
   ClientWidth     =   7695
   Icon            =   "frmBuscaPersona.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNomPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      TabIndex        =   0
      Tag             =   "1"
      Top             =   450
      Width           =   3990
   End
   Begin VB.TextBox txtDocPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      MaxLength       =   15
      TabIndex        =   8
      Tag             =   "3"
      Top             =   450
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   6
      Top             =   2265
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   5
      Top             =   1890
      Width           =   1230
   End
   Begin VB.CommandButton cmdNewCli 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   4
      Top             =   1515
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Frame frabusca 
      Caption         =   "Buscar por ...."
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
      Height          =   1170
      Left            =   120
      TabIndex        =   9
      Top             =   105
      Width           =   1800
      Begin VB.OptionButton optOpcion 
         Caption         =   "A&pellido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Có&digo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   1245
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Nº Docu&mento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   825
         Width           =   1635
      End
   End
   Begin VB.TextBox txtCodPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      MaxLength       =   13
      TabIndex        =   7
      Tag             =   "2"
      Top             =   450
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid dbgrdPersona 
      Height          =   1815
      Left            =   2085
      TabIndex        =   12
      Top             =   840
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   3201
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "cPersNombre"
         Caption         =   "Nombre  o Razon Social"
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
         DataField       =   "cPersDireccDomicilio"
         Caption         =   "Dirección"
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
         DataField       =   "cPersIDnroDNI"
         Caption         =   "Doc. Natural"
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
         DataField       =   "cPersIDnroRUC"
         Caption         =   "Doc. Juridico"
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
         DataField       =   "cPersTelefono"
         Caption         =   "Telefono"
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
         DataField       =   "dPersNacCreac"
         Caption         =   "Fecha Nac."
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
      BeginProperty Column07 
         DataField       =   "cCodZon"
         Caption         =   "Zona"
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
      BeginProperty Column08 
         DataField       =   "cPersPersoneria"
         Caption         =   "Tipo Persona"
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
            ColumnWidth     =   4515.024
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin VB.Label LblDoc 
      Height          =   195
      Left            =   4095
      TabIndex        =   11
      Top             =   525
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese Dato a Buscar :"
      Height          =   195
      Left            =   2070
      TabIndex        =   10
      Top             =   150
      Width           =   1680
   End
End
Attribute VB_Name = "frmBuscaPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Persona As UPersona
Dim R As ADODB.Recordset
Dim bBuscarEmpleado As Boolean
Public Function Inicio(Optional ByVal pbBuscarEmpleado As Boolean = False) As UPersona
   bBuscarEmpleado = pbBuscarEmpleado
   Me.Show 1
   Set Inicio = Persona
   Set Persona = Nothing
End Function

Private Sub cmdAceptar_Click()
   Set Persona = New UPersona
   If R Is Nothing Then
        MsgBox "Seleccione un Cliente", vbInformation, "Aviso"
        Exit Sub
   Else
        If R.RecordCount = 0 Then
            MsgBox "Seleccione un Cliente", vbInformation, "Aviso"
            Exit Sub
        End If
   End If
   Call Persona.CargaDatos(R!cPersCod, R!cPersNombre, IIf(IsNull(Format(R!dPersNacCreac, "dd/mm/yyyy")), "", Format(R!dPersNacCreac, "dd/mm/yyyy")), IIf(IsNull(R!cPersDireccDomicilio), "", R!cPersDireccDomicilio), IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono), R!nPersPersoneria, IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), IIf(IsNull(R!cPersIDnro), "", R!cPersIDnro), IIf(IsNull(R!cPersNatSexo), "", R!cPersNatSexo))
'   Call Persona.CargaDatos(R!cPersCod, R!cPersNombre, Format(R!dPersNacCreac, "dd/mm/yyyy"), R!cPersDireccDomicilio, IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono), R!nPersPersoneria, IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), IIf(IsNull(R!cPersIDnro), "", R!cPersIDnro))
   Set R = Nothing
   Screen.MousePointer = 0
   Unload Me
End Sub

Private Sub cmdClose_Click()
    Set R = Nothing
    Set Persona = Nothing
    Unload Me
End Sub


Private Sub cmdNewCli_Click()
Dim sCodigo As String
Dim sNombre As String
Dim RNew As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String

    sCodigo = frmPersona.PersonaNueva
    If sCodigo <> "" Then
        sNombre = Mid(sCodigo, 14, Len(sCodigo) - 13)
        sCodigo = Mid(sCodigo, 1, 13)
        optOpcion(1).value = True
        txtCodPer.Text = sCodigo
        Call txtCodPer_KeyPress(13)
        cmdAceptar_Click
        Unload Me
    End If
End Sub

Private Sub dbgrdPersona_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If R Is Nothing Then
    Exit Sub
End If
If R.EOF Then
    Exit Sub
End If
    If txtNomPer.Visible Then
        txtNomPer.Text = dbgrdPersona.Columns(0)
    Else
        If txtCodPer.Visible Then
            txtCodPer.Text = dbgrdPersona.Columns(2)
        Else
            If txtDocPer.Visible Then
                LblDoc.Caption = Trim(IIf(IsNull(R!cTipo), "", R!cTipo))
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
CentraForm Me
    dbgrdPersona.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub optOpcion_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda por Nombre
            txtNomPer.Text = ""
            txtNomPer.Visible = True
            txtNomPer.SetFocus
            txtCodPer.Text = ""
            txtCodPer.Visible = False
            txtDocPer.Text = ""
            txtDocPer.Visible = False
            LblDoc.Visible = False
        Case 1 'Busqueda por Codigo
            txtCodPer.Text = ""
            txtCodPer.Visible = True
            txtCodPer.SetFocus
            txtNomPer.Text = ""
            txtNomPer.Visible = False
            txtDocPer.Text = ""
            txtDocPer.Visible = False
            LblDoc.Visible = False
        Case 2 'Busqueda por Documento
            txtDocPer.Text = ""
            txtDocPer.Visible = True
            LblDoc.Visible = True
            txtDocPer.SetFocus
            txtCodPer.Text = ""
            txtCodPer.Visible = False
            txtNomPer.Text = ""
            txtNomPer.Visible = False
    End Select
End Sub

Private Sub txtCodPer_GotFocus()
    fEnfoque txtCodPer
End Sub

Private Sub txtCodPer_KeyPress(KeyAscii As Integer)
Dim ClsPersona As DPersonas
    If KeyAscii = 13 Then
        If Len(Trim(txtCodPer.Text)) = 0 Then
            MsgBox "Falta Ingresar el Codigo de la Persona", vbInformation, "Aviso"
            Exit Sub
        End If
        Screen.MousePointer = 11
        Set ClsPersona = New DPersonas
        If bBuscarEmpleado Then
            Set R = ClsPersona.BuscaCliente(txtCodPer.Text, BusquedaEmpleadoCodigo)
        Else
            Set R = ClsPersona.BuscaCliente(txtCodPer.Text, BusquedaCodigo)
        End If
        Set dbgrdPersona.DataSource = R
        dbgrdPersona.Refresh
        Screen.MousePointer = 0
        If R.RecordCount = 0 Then
            MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
            txtCodPer.SetFocus
            cmdAceptar.Default = False
        Else
            dbgrdPersona.SetFocus
            cmdAceptar.Default = True
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub txtDocPer_GotFocus()
    fEnfoque txtDocPer
End Sub

Private Sub txtDocPer_KeyPress(KeyAscii As Integer)
Dim ClsPersona As DPersonas
    If KeyAscii = 13 Then
        If Len(Trim(txtDocPer.Text)) = 0 Then
            MsgBox "Falta Ingresar el Documento de la Persona", vbInformation, "Aviso"
            Exit Sub
        End If
        Screen.MousePointer = 11
        Set ClsPersona = New DPersonas
        If bBuscarEmpleado Then
            Set R = ClsPersona.BuscaCliente(txtDocPer.Text, BusquedaEmpleadoDocumento)
        Else
            Set R = ClsPersona.BuscaCliente(txtDocPer.Text, BusquedaDocumento)
        End If
        Set dbgrdPersona.DataSource = R
        dbgrdPersona.Refresh
        Screen.MousePointer = 0
        If R.RecordCount = 0 Then
            MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
            txtDocPer.SetFocus
            cmdAceptar.Default = False
        Else
            dbgrdPersona.SetFocus
            cmdAceptar.Default = True
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub txtNomPer_GotFocus()
    fEnfoque txtNomPer
End Sub

Private Sub txtNomPer_KeyPress(KeyAscii As Integer)
Dim ClsPersona As DPersonas
   If KeyAscii = 13 Then
      If Len(Trim(txtNomPer.Text)) = 0 Then
        MsgBox "Falta Ingresar el Nombre de la Persona", vbInformation, "Aviso"
        Exit Sub
      End If
      Screen.MousePointer = 11
      Set ClsPersona = New DPersonas
      If bBuscarEmpleado Then
        Set R = ClsPersona.BuscaCliente(txtNomPer.Text, BusquedaEmpleadoNombre)
      Else
        Set R = ClsPersona.BuscaCliente(txtNomPer.Text)
      End If
      Set dbgrdPersona.DataSource = R
      dbgrdPersona.Refresh
      Screen.MousePointer = 0
      If R.RecordCount = 0 Then
        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
        txtNomPer.SetFocus
        cmdAceptar.Default = False
      Else
        cmdAceptar.Default = True
        txtNomPer.Text = Trim(R!cPersNombre)
        dbgrdPersona.SetFocus
      End If
      
   Else
        KeyAscii = Letras(KeyAscii)
        cmdAceptar.Default = False
   End If
End Sub

