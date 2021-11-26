VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingInformeTecCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contigencias: Informes Técnicos"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   Icon            =   "frmContingInformeTecCons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   17
      Top             =   5070
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9675
      TabIndex        =   6
      Top             =   5070
      Width           =   1050
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   5
      Top             =   5070
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8520
      TabIndex        =   4
      Top             =   5070
      Visible         =   0   'False
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTabConting 
      Height          =   4860
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   8573
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Informes Técnicos Registrados"
      TabPicture(0)   =   "frmContingInformeTecCons.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Resultado"
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   10335
         Begin MSDataGridLib.DataGrid DBGrdInformesTec 
            Height          =   2655
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   4683
            _Version        =   393216
            AllowUpdate     =   0   'False
            ColumnHeaders   =   -1  'True
            HeadLines       =   2
            RowHeight       =   17
            RowDividerStyle =   4
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
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "nIdInfTec"
               Caption         =   "Item"
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
               DataField       =   "dFechaRegInfTec"
               Caption         =   "Fecha Informe"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "cNroInforme"
               Caption         =   "Nº Informe"
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
               Caption         =   "Personal a Cargo"
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
               DataField       =   "cCalif"
               Caption         =   "Calificacion"
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
            BeginProperty Column06 
               DataField       =   "nMontoRealGasto"
               Caption         =   "Monto Real Pérdida (Gasto)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "cMonedaDem"
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
            BeginProperty Column08 
               DataField       =   "nMontoDem"
               Caption         =   "Monto Demanda (Control)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               SizeMode        =   1
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Size            =   800
               BeginProperty Column00 
                  ColumnWidth     =   0
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column02 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   1904.882
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   1800
               EndProperty
               BeginProperty Column05 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   2475.213
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   780.095
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   2280.189
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Generales"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   10335
         Begin VB.Label Label9 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   3600
            TabIndex        =   16
            Top             =   645
            Width           =   1215
         End
         Begin VB.Label lblDescripcion 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   4920
            TabIndex        =   15
            Top             =   600
            Width           =   5235
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Registro:"
            Height          =   255
            Left            =   3600
            TabIndex        =   14
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblFecRegistro 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4920
            TabIndex        =   13
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Usuario:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   640
            Width           =   1095
         End
         Begin VB.Label lblUsuarioReg 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1440
            TabIndex        =   11
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label Label3 
            Caption         =   "Contingencia:"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblContingTipo 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1440
            TabIndex        =   9
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Nº de Registro:"
            Height          =   255
            Left            =   7680
            TabIndex        =   8
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblNroRegistro 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9000
            TabIndex        =   7
            Top             =   270
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmContingInformeTecCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingenciaCons
'** Descripción : Consulta de Contingencias por Areas creado segun RFC056-2012
'** Creación : JUEZ, 20120618 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim sNumRegistro As String
Dim nNumItem As String

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Public Function Consulta(ByVal psNumRegistro As String, Optional PrimerIngreso As Integer = 0)
    If CargaDatos(psNumRegistro, True, False) Then
        Me.Caption = "Contingencias: Consulta Informes Técnicos"
        If PrimerIngreso = 1 Then
            Me.Show 1
        End If
    End If
End Function
Public Function Extorno(ByVal psNumRegistro As String, Optional PrimerIngreso As Integer = 0)
    If CargaDatos(psNumRegistro, True, True) Then
        Me.Caption = "Contingencias: Ver Informes Técnicos"
        If PrimerIngreso = 1 Then
            Me.Show 1
        End If
    End If
End Function

Private Function CargaDatos(ByVal psNumRegistro As String, ByVal pbHabilitaImprimir As Boolean, pbHabilitaExtornar As Boolean) As Boolean
    Dim rsConting As ADODB.Recordset
    Set oConting = New DContingencia
    Set rsConting = oConting.BuscaContigenciaSeleccionada(psNumRegistro)
    
    lblContingTipo.Caption = IIf(rsConting!cTpoConting = "Activo Contingente", "Contingencia Activa", "Contingencia Pasiva")
    lblUsuarioReg.Caption = rsConting!cUserReg
    lblFecRegistro.Caption = Format(rsConting!dFechaReg, "dd/mm/yyyy")
    lblDescripcion.Caption = " " & rsConting!cContigDesc
    lblNroRegistro.Caption = rsConting!cNumRegistro
    Set rsConting = Nothing
    
    Set rs = oConting.BuscaInfsTecnicosxConting(psNumRegistro)
    Set DBGrdInformesTec.DataSource = rs
    DBGrdInformesTec.Refresh
    Screen.MousePointer = 0
    If rs.RecordCount = 0 Then
      MsgBox "No se Encontraron Informes Tecnicos", vbInformation, "Aviso"
      cmdImprimir.Visible = False
      cmdExtornar.Visible = False
      cmdEditar.Visible = False
      CargaDatos = False
      Exit Function
    Else
      cmdImprimir.Visible = pbHabilitaImprimir
      cmdExtornar.Visible = pbHabilitaExtornar
      cmdEditar.Visible = pbHabilitaExtornar
    End If
    CargaDatos = True
End Function

Private Sub cmdEditar_Click()
    nNumItem = CInt(DBGrdInformesTec.Columns(0))
    If nNumItem <> 0 Then
        Set oConting = New DContingencia
        Set rs = oConting.BuscaContigenciaSeleccionada(lblNroRegistro.Caption)
        If rs!nestado = 3 Then
            MsgBox "No puede editar el IT, la contingencia ya está liberada", vbInformation, "Aviso"
        ElseIf rs!nestado = 4 Then
            MsgBox "No puede editar el IT, la contingencia ya está desestimada", vbInformation, "Aviso"
        Else
            If Left(Trim(lblNroRegistro.Caption), 1) = gActivoContingente Then
                frmContingInformeTecEdit.EditarActivo nNumItem, lblNroRegistro.Caption
            Else
                frmContingInformeTecEdit.EditarPasivo nNumItem, lblNroRegistro.Caption
            End If
        End If
    End If
End Sub

Private Sub cmdExtornar_Click()
    nNumItem = CInt(DBGrdInformesTec.Columns(0))
    If nNumItem <> 0 Then
        Set oConting = New DContingencia
        Set rs = oConting.BuscaContigenciaSeleccionada(lblNroRegistro.Caption)
        If rs!nestado = 3 Then
            MsgBox "No puede extornar el IT, la contingencia ya está liberada", vbInformation, "Aviso"
        ElseIf rs!nestado = 4 Then
            MsgBox "No puede extornar el IT, la contingencia ya está desestimada", vbInformation, "Aviso"
        Else
            If MsgBox("Está seguro de extornar el IT de la Contingencia? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            
            Call oConting.ExtornaITConting(nNumItem, lblNroRegistro.Caption)
            
            MsgBox "Se ha extornado con exito el IT elegido", vbInformation, "Aviso"
            Call Extorno(lblNroRegistro.Caption)
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim lsCadImp As String
    'sNumRegistro = DBGrdConting.Columns(0)
    'If sNumRegistro <> "" Then
    Screen.MousePointer = 11
    lsCadImp = ImprimirRegistroITContingencias
    If Len(Trim(lsCadImp)) > 0 Then
        EnviaPrevio lsCadImp, "Informes Técnicos", gnLinPage, False
    Else
        MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
    End If
    Screen.MousePointer = 0
End Sub

Public Function ImprimirRegistroITContingencias() As String
    Dim lsCad As String
    Dim cAreaNombre As String
    Dim oAreas As DActualizaDatosArea
    Set oConting = New DContingencia
    Set rs = oConting.BuscaInfsTecnicosxConting(lblNroRegistro.Caption)
            
    Dim psTitulo As String
    Dim pnAnchoLinea As Integer
    Dim RDatosUser As ADODB.Recordset
    Dim lnNegritaON As String
    Dim lnNegritaOFF As String
    
    psTitulo = "INFORMES TÉCNICOS"
    pnAnchoLinea = 125

    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    lsCad = CON & BON & Chr$(10) & Chr$(10) & Centra(" " & psTitulo & " ", pnAnchoLinea) & Chr$(10)
    lsCad = lsCad & Centra(" " & gdFecSis & " ", pnAnchoLinea) & Chr$(10) & Chr$(10)
    lsCad = lsCad & "Contingencia: Contingencia " & IIf(Left(lblNroRegistro.Caption, 1) = "1", "Activa", "Pasiva")
    lsCad = lsCad & Space(pnAnchoLinea - 59) & "Fecha Registro: " & FillText(Trim(lblFecRegistro.Caption), 11, " ") & BOFF & Chr$(10)
    lsCad = lsCad & "Descripción: " & FillText(Me.lblDescripcion.Caption, 50, " ")
    lsCad = lsCad & Space(pnAnchoLinea - 76) & "Usuario: " & FillText(Me.lblUsuarioReg.Caption, 5, " ") & BOFF & Chr$(10)
    lsCad = lsCad & String(pnAnchoLinea, "-") & Chr$(10)
    lsCad = lsCad & "F.Informe   Nro Informe          Personal a Cargo                     Calificacion               Moneda    Provisión" & Chr$(10)
    lsCad = lsCad & String(pnAnchoLinea, "-") & Chr$(10)
    Do While Not rs.EOF
        lsCad = lsCad & Left(rs!dFechaRegInfTec & String(10, " "), 10) & Space(2)
        lsCad = lsCad & Left(rs!cNroInforme & String(19, " "), 19) & Space(2)
        lsCad = lsCad & Left(rs!cPersNombre & String(35, " "), 35) & Space(2)
        lsCad = lsCad & Left(rs!cCalif & String(22, " "), 22) & Space(2)
        lsCad = lsCad & Left(rs!cmoneda & String(7, " "), 7) & Space(3)
        lsCad = lsCad & Right(String(14, " ") & Format(rs!nProvision, "#,##0.00"), 12) & Space(2)
        lsCad = lsCad & Chr$(10)
        rs.MoveNext
    Loop
    lsCad = lsCad & String(pnAnchoLinea, "-") & COFF & Chr$(10)
    
    ImprimirRegistroITContingencias = lsCad
End Function

