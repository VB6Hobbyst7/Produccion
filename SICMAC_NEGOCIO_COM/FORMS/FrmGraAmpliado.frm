VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmGraAmpliado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gravament de Creditos Ampliados"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   9255
      Begin TabDlg.SSTab SSTab1 
         Height          =   3405
         Left            =   150
         TabIndex        =   12
         Top             =   1440
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   6006
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Registro"
         TabPicture(0)   =   "FrmGraAmpliado.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DGRel"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DGGarantia"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Mantenimiento"
         TabPicture(1)   =   "FrmGraAmpliado.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LstGarantia"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid DGGarantia 
            Height          =   1215
            Left            =   150
            TabIndex        =   13
            Top             =   600
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   2143
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DGRel 
            Height          =   1215
            Left            =   120
            TabIndex        =   14
            Top             =   2130
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   2143
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView LstGarantia 
            Height          =   2685
            Left            =   -74910
            TabIndex        =   17
            Top             =   540
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   4736
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ColHdrIcons     =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Credito"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Monto"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Garantias"
            Height          =   195
            Left            =   210
            TabIndex        =   16
            Top             =   330
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Relacion del Credito Ampliado con la Garantia"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   1890
            Width           =   3240
         End
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   150
         TabIndex        =   5
         Top             =   4860
         Width           =   8775
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   3180
            TabIndex        =   10
            Top             =   180
            Width           =   1005
         End
         Begin VB.CommandButton CmdSalir 
            Caption         =   "Salir"
            Height          =   375
            Left            =   7530
            TabIndex        =   9
            Top             =   210
            Width           =   1005
         End
         Begin VB.CommandButton CmdEliminar 
            Caption         =   "Eliminar"
            Height          =   375
            Left            =   2130
            TabIndex        =   8
            Top             =   180
            Width           =   1005
         End
         Begin VB.CommandButton CmdNuevo 
            Caption         =   "Nuevo"
            Height          =   375
            Left            =   90
            TabIndex        =   7
            Top             =   180
            Width           =   1005
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   1110
            TabIndex        =   11
            Top             =   180
            Width           =   1005
         End
      End
      Begin VB.CommandButton CmdExaminar 
         Caption         =   "Examinar"
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   510
         Width           =   1305
      End
      Begin SICMACT.ActXCodCta ActXCodCta1 
         Height          =   375
         Left            =   150
         TabIndex        =   1
         Top             =   480
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   661
         Texto           =   "Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin MSComctlLib.ListView Lst 
         Height          =   885
         Left            =   5340
         TabIndex        =   4
         Top             =   540
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   1561
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Credito"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Credito Ampliado"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lista de Creditos"
         Height          =   195
         Left            =   5370
         TabIndex        =   3
         Top             =   300
         Width           =   1170
      End
   End
End
Attribute VB_Name = "FrmGraAmpliado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGarantia As ADODB.Recordset
Dim rsRelGarant As ADODB.Recordset
Dim sGarant As String

Private Sub ActXCodCta1_KeyPress(KeyAscii As Integer)
    Dim oAmpliacion As COMDCredito.DCOMAmpliacion
    If KeyAscii = 13 Then
        Set oAmpliacion = New COMDCredito.DCOMAmpliacion
        If oAmpliacion.ValidarCreditoGarantizar(Me.ActXCodCta1.NroCuenta) = False Then
            Set oAmpliacion = Nothing
            MsgBox "Este credito no esta en un estado de relacionar su garantia", vbInformation, "AVISO"
        Else
            Set oAmpliacion = Nothing
            CargarDatosCredito (Me.ActXCodCta1.NroCuenta)
            CargarGarantiaAmpliada (Me.ActXCodCta1.NroCuenta)
        End If
    End If
End Sub

Private Sub cmdcancelar_Click()
    Form_Load
End Sub

Private Sub CmdEliminar_Click()
    Dim oAmpliado As COMDCredito.DCOMAmpliacion
    If MsgBox("Esta seguro  que desea eliminar toda la operacion", vbInformation + vbYesNo, "AVISO") = vbYes Then
        Set oAmpliado = New COMDCredito.DCOMAmpliacion
        If oAmpliado.EliminarGarantiaAmpliada(Me.ActXCodCta1.NroCuenta) = True Then
           MsgBox "Se elimino correctamente la garantia", vbInformation, "AVISO"
           Form_Load
        Else
            MsgBox "Error al eliminar la garantia", vbInformation, "AVISO"
        End If
    End If
End Sub

Private Sub cmdExaminar_Click()
    FrmListaCreditosAmpliado.Show vbModal
End Sub

Public Sub CargarCredito(ByVal psCtaCod As String)
    Me.ActXCodCta1.CMAC = "108"
    Me.ActXCodCta1.Age = gsCodAge
    Me.ActXCodCta1.NroCuenta = psCtaCod
    
    CargarDatosCredito psCtaCod
    CargarGarantiaAmpliada psCtaCod
End Sub

Sub CargarDatosCredito(ByVal psCtaCod As String)
    Dim oAmpliado As COMDCredito.DCOMAmpliacion
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    
    Lst.ListItems.Clear
    
    Set oAmpliado = New COMDCredito.DCOMAmpliacion
    Set rs = oAmpliado.ListaCreditosBycCtaCodNew(psCtaCod)
    Set oAmpliado = Nothing
    
    Do Until rs.EOF
        Set Item = Lst.ListItems.Add(, , rs!cCtaCodAmp)
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub


Private Sub cmdGrabar_Click()
    Dim oAmpliacion As COMDCredito.DCOMAmpliacion
    Dim rs As ADODB.Recordset
    'proceso para guardar la informacion
    Set rs = rsRelGarant.Clone
    If rs.RecordCount > 0 Then
        If Not rs.EOF And Not rs.BOF Then
            rs.MoveFirst
            If rs(0) <> "" And rs(1) <> "" And rs(3) <> "" Then
               ' If VerificarMonto = True Then
                    Set oAmpliacion = New COMDCredito.DCOMAmpliacion
                    If oAmpliacion.InsertarGarantiaAmpliado(rs) = True Then
                        MsgBox "Se ha guardado de forma correcta", vbInformation, "AVISO"
                        Form_Load
                    Else
                        MsgBox "Error al momento de guardar", vbInformation, "AVISO"
                    End If
                    Set oAmpliacion = Nothing
                'Else
                    MsgBox "El monto no cubre la garantia", vbInformation, "AVISO"
                'End If
            Else
                MsgBox "No existe informacion para guardar", vbInformation, "AVISO"
            End If
        Else
            MsgBox "No existe informacion para guardar", vbInformation, "AVISO"
        End If
    Else
        MsgBox "No existe informacion para guardar", vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdNuevo_Click()
    If sGarant <> "" Then
        If VerificarRegistro = False Then
            rsRelGarant.AddNew
            rsRelGarant(0) = rsGarantia(0)
            rsRelGarant(1) = Me.ActXCodCta1.NroCuenta
            rsRelGarant(2) = Lst.ListItems(Lst.SelectedItem.Index)
            rsRelGarant(3) = rsGarantia(3)
            rsRelGarant.Update
            'CmdGrabar.Visible = True
            'CmdNuevo.Visible = False
        Else
            MsgBox "Ya existe es registro", vbInformation, "AVISO"
        End If
    Else
        MsgBox "Debe seleccionar la garantia", vbInformation, "AVISO"
    End If
End Sub

Function VerificarRegistro() As Boolean
    Dim rs As ADODB.Recordset
    Dim bEstado As Boolean
    
    bEstado = False
    Set rs = rsRelGarant.Clone
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
        Do Until rs.EOF
            If rs(0) = rsGarantia(0) And rs(1) = Me.ActXCodCta1.NroCuenta And rs(2) = rsGarantia(2) Then
                bEstado = True
                Exit Do
            Else
                bEstado = False
            End If
           rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    VerificarRegistro = bEstado
End Function

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub DGGarantia_Click()
    sGarant = rsGarantia(0)
End Sub


Private Sub DGRel_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim oAmpliacion As COMDCredito.DCOMAmpliacion
    Dim nMonto As Double
    
 If ColIndex = 3 Then
        Set oAmpliacion = New COMDCredito.DCOMAmpliacion
        nMonto = oAmpliacion.ObteneMontoGarant(DGRel.Columns(0).value, DGRel.Columns(2).value)
        If nMonto < DGRel.Columns(3).value Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Load()
    ConfigurarDG
    ConfigurarDGRelacion
    CmdNuevo.Visible = True
    'CmdEliminar.Visible = False
    sGarant = ""
    CmdNuevo.Enabled = True
    CmdGrabar.Enabled = True
    Lst.ListItems.Clear
    LstGarantia.ListItems.Clear
    Me.ActXCodCta1.Texto = ""
End Sub

Private Sub Lst_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rs As ADODB.Recordset
    Dim oAmpliado As COMDCredito.DCOMAmpliacion
    
    If Not Lst.SelectedItem Is Nothing Then
      Set oAmpliado = New COMDCredito.DCOMAmpliacion
      Set rs = oAmpliado.ListaGarantiaCtaAmpliado(Me.ActXCodCta1.NroCuenta, Lst.ListItems(Lst.SelectedItem.Index))
      Set oAmpliado = Nothing
      ConfigurarDG
      Do Until rs.EOF
         rsGarantia.AddNew
         rsGarantia(0) = rs!cNumGarant
         rsGarantia(1) = rs!cDescripcion
         rsGarantia(2) = rs!cConsDescripcion
         rsGarantia(3) = rs!nGravado
         rs.MoveNext
      Loop
    End If
End Sub

Sub ConfigurarDG()
    Set rsGarantia = New ADODB.Recordset
    
    With rsGarantia.Fields
         .Append "Codigo", adChar, 8
         .Append "Descripcion", adVarChar, 50
         .Append "Tipo Garantia", adVarChar, 40
         .Append "Gravado", adCurrency
    End With
    
    rsGarantia.Open
    Set DGGarantia.DataSource = rsGarantia
    
    DGGarantia.Columns(1).Width = 2500
    DGGarantia.Columns(2).Width = 2500
    DGGarantia.Columns(3).NumberFormat = "#0.00"
End Sub

Sub ConfigurarDGRelacion()
    Set rsRelGarant = New ADODB.Recordset
    
    With rsRelGarant.Fields
         .Append "Codigo", adChar, 8
         .Append "Cuenta", adVarChar, 18
         .Append "Cuenta Ampliada", adVarChar, 18
         .Append "Gravado", adCurrency
    End With
    
    rsRelGarant.Open
    Set DGRel.DataSource = rsRelGarant
    
    'DGRel.Columns(1).Width = 2500
    'DGRel.Columns(2).Width = 2500
    DGRel.Columns(3).NumberFormat = "#0.00"
End Sub


Function VerificarMonto() As Boolean
    Dim oGeneral As COMDConstSistema.DCOMGeneral
    Dim nTipoCambioFijo As Double
    Dim oAmpliacion As COMDCredito.DCOMAmpliacion
    Dim nMontoCol As Double
    Dim nMontoG As Double
    Dim nMonedaG As Integer
    Dim rs As ADODB.Recordset
    Dim oCred As COMDCredito.DCOMCredito
    
    Set rs = rsRelGarant.Clone
    rs.MoveFirst
    
    Set oGeneral = New COMDConstSistema.DCOMGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(gdFecSis, TCFijoMes)
    nTipoCambioFijo = CDbl(Format(nTipoCambioFijo, "#0.00"))
    Set oGeneral = Nothing
    
    Set oAmpliacion = New COMDCredito.DCOMAmpliacion
    nMontoCol = oAmpliacion.ObtenerMontoColocar(Me.ActXCodCta1.NroCuenta)
    
    Set oAmpliacion = Nothing
    
    
    'determinar la moneda del credito
    nMontoG = 0
    If Mid(Me.ActXCodCta1.NroCuenta, 9, 1) = "1" Then
        Do Until rs.EOF
            Set oAmpliacion = New COMDCredito.DCOMAmpliacion
                If oAmpliacion.ObtenerMonedaGarantia(rs(0)) = 1 Then
                    nMontoG = nMontoG + Val(rs(3))
                Else
                    nMontoG = nMontoG + Val(rs(3)) * nTipoCambioFijo
                End If
            Set oAmpliacion = Nothing
            rs.MoveNext
        Loop
    Else
        Do Until rs.EOF
            Set oAmpliacion = New COMDCredito.DCOMAmpliacion
                If oAmpliacion.ObtenerMonedaGarantia(rs(0)) = 1 Then
                    nMontoG = nMontoG + (Val(rs(3)) / nTipoCambioFijo)
                Else
                    nMontoG = nMontoG + Val(rs(3))
                End If
            Set oAmpliacion = Nothing
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    
    Set oCred = New COMDCredito.DCOMCredito
    Set rs = oCred.RecuperaColocGarantia(Me.ActXCodCta1.NroCuenta)
    Do While Not rs.EOF
        If Mid(Me.ActXCodCta1.NroCuenta, 9, 1) = "1" Then
            If rs!nmoneda = 1 Then
                nMontoG = nMontoG + rs!nGravado
            Else
                nMontoG = nMontoG + Val(rs!nGravado) * nTipoCambioFijo
            End If
        Else
            If rs!nmoneda = 1 Then
                nMontoG = nMontoG + Val(rs!nGravado) / nTipoCambioFijo
            Else
                nMontoG = nMontoG + rs!nGravado
            End If
        End If
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    If nMontoCol > nMontoG Then
        VerificarMonto = False
    Else
        VerificarMonto = True
    End If
    
End Function

Sub CargarGarantiaAmpliada(ByVal psCtaCod As String)
    Dim oAmpliado As COMDCredito.DCOMAmpliacion
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    Dim bEstado As Boolean
    
    Set oAmpliado = New COMDCredito.DCOMAmpliacion
    Set rs = oAmpliado.ObtenerListaGarantiasAmpliadas(psCtaCod)
    Set oAmpliado = Nothing
    
    bEstado = False
    
     Do Until rs.EOF
        Set Item = LstGarantia.ListItems.Add(, , rs!cNumGarant)
        Item.SubItems(1) = rs!cCtaCod
        Item.SubItems(2) = Format(rs!nMontoGravado, "#0.00")
        rs.MoveNext
        bEstado = True
     Loop
     Set rs = Nothing
     
     If bEstado = True Then
        SSTab1.Tab = 1
        CmdNuevo.Enabled = False
        CmdGrabar.Enabled = False
     End If
End Sub

