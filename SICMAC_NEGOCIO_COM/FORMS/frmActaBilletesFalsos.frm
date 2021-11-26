VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmActaBilletesFalsos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acta de Inacautación de Monedas y  Billetes Falsos"
   ClientHeight    =   7680
   ClientLeft      =   3465
   ClientTop       =   2040
   ClientWidth     =   9285
   Icon            =   "frmActaBilletesFalsos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   660
      Left            =   180
      TabIndex        =   19
      Top             =   45
      Width           =   8910
      Begin VB.TextBox TxtNroActa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1095
         MaxLength       =   10
         TabIndex        =   37
         Top             =   210
         Width           =   1410
      End
      Begin VB.CommandButton cmdFiltro 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2550
         TabIndex        =   23
         Top             =   210
         Width           =   390
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Registo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3615
         TabIndex        =   22
         Top             =   285
         Width           =   1305
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Acta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   270
         Width           =   885
      End
      Begin VB.Label lblFechaReg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4980
         TabIndex        =   20
         Top             =   210
         Width           =   1380
      End
   End
   Begin VB.Frame fradatos 
      Caption         =   "Datos de Acta"
      Height          =   6240
      Left            =   180
      TabIndex        =   0
      Top             =   810
      Width           =   8925
      Begin VB.OptionButton OptMoneda 
         Caption         =   "Dólares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   7110
         TabIndex        =   33
         Top             =   1455
         Width           =   1155
      End
      Begin VB.OptionButton OptMoneda 
         Caption         =   "Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   6030
         TabIndex        =   32
         Top             =   1455
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.CommandButton cmdMenos 
         Caption         =   "&-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8370
         TabIndex        =   25
         Top             =   4170
         Width           =   390
      End
      Begin VB.CommandButton cmdMas 
         Caption         =   "&+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8355
         TabIndex        =   24
         Top             =   3480
         Width           =   390
      End
      Begin VB.CommandButton cmdVerif 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1755
         TabIndex        =   17
         Top             =   2145
         Width           =   390
      End
      Begin SICMACT.FlexEdit grdDetalle 
         Height          =   2925
         Left            =   135
         TabIndex        =   15
         Top             =   2805
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   5159
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-SERIE Y NUMERACION-FECHA  EMISION-DENOMINACION"
         EncabezadosAnchos=   "800-3500-1500-2250"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3"
         ListaControles  =   "0-0-2-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R"
         FormatosEdit    =   "0-0-0-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   795
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2355
         TabIndex        =   8
         Top             =   675
         Width           =   390
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   2160
         TabIndex        =   28
         Top             =   1395
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LblNroActa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6240
         TabIndex        =   35
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Acta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   34
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5190
         TabIndex        =   31
         Top             =   1470
         Width           =   690
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6090
         TabIndex        =   30
         Top             =   5775
         Width           =   2190
      End
      Begin VB.Label LblEtqMon 
         AutoSize        =   -1  'True
         Caption         =   "Total S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5130
         TabIndex        =   29
         Top             =   5850
         Width           =   780
      End
      Begin VB.Label lblUsuDetector 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2160
         TabIndex        =   27
         Top             =   1770
         Width           =   6270
      End
      Begin VB.Label LblUsuVerificador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2160
         TabIndex        =   26
         Top             =   2145
         Width           =   6270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DETALLE DE INFORMACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2595
         TabIndex        =   16
         Top             =   2595
         Width           =   2490
      End
      Begin VB.Label lblDICli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5565
         TabIndex        =   14
         Top             =   660
         Width           =   2250
      End
      Begin VB.Label lblCodigoCli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Top             =   660
         Width           =   1380
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Documento Identidad:"
         Height          =   195
         Left            =   3900
         TabIndex        =   12
         Top             =   690
         Width           =   1575
      End
      Begin VB.Label lblNomCli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   1005
         Width           =   7485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   1065
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Verificado Por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   2220
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Detectado Por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inacutación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   1470
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Portador:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   4
         Top             =   435
         Width           =   795
      End
   End
   Begin SICMACT.Usuario ctlusuario 
      Left            =   2550
      Top             =   5685
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   420
      Left            =   1245
      TabIndex        =   18
      Top             =   7140
      Width           =   1020
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   420
      Left            =   195
      TabIndex        =   3
      Top             =   7140
      Width           =   1020
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   8100
      TabIndex        =   2
      Top             =   7140
      Width           =   1020
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   420
      Left            =   6915
      TabIndex        =   1
      Top             =   7140
      Width           =   1020
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   330
      Left            =   4665
      TabIndex        =   36
      Top             =   6720
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   582
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmActaBilletesFalsos.frx":030A
   End
End
Attribute VB_Name = "frmActaBilletesFalsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String, lsDni As String
Dim lsEstados As String


'On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
'    If loPers Is Nothing Then
'          FraCli.Visible = False
'          FraReferencia.Visible = True
'          txtDI.Text = ""
'          TxtNombre.Text = ""
'          txtDI.SetFocus
'
'       Exit Sub
'    End If
    
   If loPers Is Nothing Then
       Exit Sub
    End If
    
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    lsDni = loPers.sPersIdnroDNI

Dim i As Integer
If lsPersCod <> "" Then
    'Me.lbCaption = lsPersCod
    Me.lblCodigoCli.Caption = lsPersCod
    Me.lblNomCli.Caption = lsPersNombre
    Me.lblDICli.Caption = lsDni
    txtFecha.SetFocus
        
End If

Set loPers = Nothing

End Sub

Private Sub cmdDetec_Click()

End Sub

Private Sub cmdFiltro_Click()
 frmActaLista.Inicia (1)
 fradatos.Enabled = True
 LLenaDatos
 fradatos.Enabled = False
 cmdGrabar.Enabled = False
 cmdImprimir.Enabled = True
 cmdNuevo.Enabled = True
End Sub
Private Sub LLenaDatos()
Dim CLSSERV As dCapServicios
Dim rstmp As Recordset, rstmpdet As Recordset

Set CLSSERV = New dCapServicios

    Set rstmp = CLSSERV.GetInfoActaCompleta(Trim(lblUsuDetector.Tag), Trim(TxtNroActa.Text))
    If Not rstmp.EOF Then
        LblNroActa.Caption = Trim(TxtNroActa.Text)
        lblNomCli.Caption = rstmp!PORTADOR
        lblDICli.Caption = rstmp!DNI
        lblCodigoCli.Caption = rstmp!CODPORTADOR
        txtFecha.Text = Format(rstmp!Fecha, "dd/MM/yyyy")
        LblUsuVerificador.Caption = rstmp!VERIFICADOR
        LblUsuVerificador.Tag = rstmp!CODusuVERIFICADOR
        lblFechaReg.Caption = Format(rstmp!fechaReg, "dd/MM/yyyy")
        
        If rstmp!cmoneda = 1 Then
            OptMoneda(0).value = True
            LblEtqMon.Caption = "Total S/."
            lblTotal.BackColor = &HC0FFFF
        Else
            OptMoneda(1).value = True
            LblEtqMon.Caption = "Total U$."
            lblTotal.BackColor = &HC0FFC0
        End If
                
        Set rstmpdet = CLSSERV.GetInfoDetActa(Trim(TxtNroActa.Text))
        Set grdDetalle.Recordset = rstmpdet
        Dim i As Integer, suma As Double
        suma = 0
        For i = 1 To grdDetalle.Rows - 1
            suma = suma + Val(grdDetalle.TextMatrix(i, 3))
        Next i
        lblTotal.Caption = Format(suma, "#,##0.00")
        
    Else
        MsgBox "No se encontró información para esta búsqueda. ", vbOKOnly + vbExclamation, "AVISO"
    
    End If

Set CLSSERV = Nothing
End Sub

Private Sub cmdGrabar_Click()
Dim CLSSERV As NCapServicios, bResultado As Boolean
Set CLSSERV = New NCapServicios
Dim sNroActa As String

If Trim(lblCodigoCli.Caption) = "" Then
    MsgBox "Cliente no válido", vbInformation + vbOKOnly, "Aviso"
    cmdBuscar.SetFocus
    Exit Sub
End If

If Not IsDate(txtFecha) Then
    MsgBox "Fecha no válido", vbInformation + vbOKOnly, "Aviso"
    Me.txtFecha.SetFocus
    Exit Sub
End If

If Trim(LblUsuVerificador.Tag) = "" Then
    MsgBox "Incluir persona que verificó información", vbInformation + vbOKOnly, "Aviso"
    cmdBuscar.SetFocus
    Exit Sub
End If

If NoBlancos = True Then
   MsgBox "Existe información en blanco en el detalle del Acta.", vbInformation + vbOKOnly, "AVISO"
   Exit Sub
End If

Dim smovnro As String, clsMov As NContFunciones
Set clsMov = New NContFunciones
smovnro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

bResultado = CLSSERV.GrabaActa(lblCodigoCli.Caption, CDate(Format(txtFecha.Text, "yyyy-mm-dd")), lblUsuDetector.Tag, LblUsuVerificador.Tag, smovnro, IIf(OptMoneda(0).value, "1", "2"), grdDetalle.GetRsNew, sNroActa)

If bResultado = False Then
   MsgBox "NO SE GRABO INFORMACION" & vbCrLf & " COMUNICARSE CON EL AREA DE SISTEMAS", vbOKOnly + vbExclamation, "AVISO"
   
Else
   LblNroActa.Caption = sNroActa
   cmdImprimir.Enabled = True
   cmdGrabar.Enabled = False
   fradatos.Enabled = False
End If

Set CLSSERV = Nothing
End Sub

Private Function NoBlancos() As Boolean
Dim i As Integer

NoBlancos = False

    For i = 1 To grdDetalle.Rows - 1
        If Trim(grdDetalle.TextMatrix(i, 1)) = "" Or Trim(grdDetalle.TextMatrix(i, 2)) = "" Or Trim(grdDetalle.TextMatrix(i, 3)) = "" Then
            NoBlancos = True
            Exit Function
        End If
    Next i


End Function



Private Sub cmdImprimir_Click()
Dim sNomA As String
'Dim lsCadImprimir  As String
'Dim loPrevio As Previo.clsPrevio
 sNomA = App.path & "\FormatoCarta\ActaBilletesFalsos.doc"
 
    
    lsCadImprimir = ""
    rtfCartas.FileName = sNomA
  
                
  '     lsCadImprimir = FomaCarta(rtfCartas.Text)
    Call CartaWORD(sNomA, LblNroActa.Caption)

'        Set loPrevio = New Previo.clsPrevio
'            loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", False
'        Set loPrevio = Nothing




End Sub
Private Function FomaCarta(ByVal psTextoCarta As String) As String
Dim lsCartaModelo As String, lsCadImp As String, i As Integer
lsCadImp = ""
Do
   
        lsCartaModelo = psTextoCarta
        lsCartaModelo = Replace(lsCartaModelo, "<<NRO ACTA>>", LblNroActa.Caption, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE DEL PORTADOR>>", Trim(lblNomCli.Caption), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECHA INCAUTACION>>", Me.txtFecha, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<USUARIO OPERADOR>>", Mid(Trim(lblUsuDetector.Caption), 6), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<USUARIO VERIFICADOR>>", Mid(Trim(LblUsuVerificador.Caption), 6), , 1, vbTextCompare)
                       
        For i = 1 To grdDetalle.Rows - 1
            lsCartaModelo = Replace(lsCartaModelo, "<SERIE" & CStr(i) & ">", grdDetalle.TextMatrix(i, 1), , 1, vbTextCompare)
            lsCartaModelo = Replace(lsCartaModelo, "<FECHA" & CStr(i) & ">", grdDetalle.TextMatrix(i, 2), , 1, vbTextCompare)
            lsCartaModelo = Replace(lsCartaModelo, "<DENOMINACION" & CStr(i) & ">", grdDetalle.TextMatrix(i, 3), , 1, vbTextCompare)
        Next i
        
        lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
        FomaCarta = lsCadImp
Loop Until MsgBox("Desea Reimprimir Acta Nro: " & LblNroActa.Caption, vbOKOnly + vbExclamation, "AVISO")



End Function



Private Sub CartaWORD(ByVal psNomPlantilla As String, ByVal NroActa As String)
Dim aLista() As String, i As Integer
Dim vFilas As Integer
 
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
     
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=psNomPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
                
    'Crea Nuevo Documento
    wApp.Documents.Add
        
    wApp.Application.Selection.TypeParagraph


'    With wApp.Selection.PageSetup
'        .TopMargin = 120 'CentimetersToPoints(10)
'        .BottomMargin = 60 'CentimetersToPoints(3)
'        .LeftMargin = 140 'CentimetersToPoints(10)
'        .RightMargin = 80 'CentimetersToPoints(5)
'    End With
        
        wApp.Application.Selection.Paste
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd

    With wApp.Selection.Find
            .Text = "<<NRO ACTA>>"
            .Replacement.Text = Trim(LblNroActa.Caption)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
     End With
     wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
     With wApp.Selection.Find
            .Text = "<<NOMBRE DEL PORTADOR>>"
            .Replacement.Text = Trim(lblNomCli.Caption)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
       End With
       wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<FECHA INCAUTACION>>"
            .Replacement.Text = Trim(txtFecha.Text)
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<USUARIO OPERADOR>>"
            .Replacement.Text = Trim(Mid(lblUsuDetector.Caption, 6))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<USUARIO VERIFICADOR>>"
            .Replacement.Text = Trim(Mid(LblUsuVerificador.Caption, 6))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
    For i = 1 To grdDetalle.Rows - 1
        With wApp.Selection.Find
            .Text = "<SERIE" & CStr(i) & ">"
            .Replacement.Text = Trim(grdDetalle.TextMatrix(i, 1))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<FECHA" & CStr(i) & ">"
            .Replacement.Text = Trim(grdDetalle.TextMatrix(i, 2))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<DENOMINACION" & CStr(i) & ">"
            '.Replacement.Text = Format(Trim(grdDetalle.TextMatrix(i, 3)), IIf(OptMoneda(0).value, "S/.", "U$") & "#,##0.00")
            .Replacement.Text = IIf(OptMoneda(0).value, "S/.", "U$") & Space(15 - Len(Trim(JDNum(Trim(grdDetalle.TextMatrix(i, 3)), 12, True, 9, 2)))) & JDNum(Trim(grdDetalle.TextMatrix(i, 3)), 12, True, 9, 2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
   Next i
   If i < 18 Then
           For i = i To 18
                    With wApp.Selection.Find
                        .Text = "<SERIE" & CStr(i) & ">"
                        .Replacement.Text = "               "
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                      End With
                    wApp.Selection.Find.Execute Replace:=wdReplaceAll
                    
                    With wApp.Selection.Find
                        .Text = "<FECHA" & CStr(i) & ">"
                        .Replacement.Text = "          "
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                      End With
                    wApp.Selection.Find.Execute Replace:=wdReplaceAll
                    
                    With wApp.Selection.Find
                        .Text = "<DENOMINACION" & CStr(i) & ">"
                        .Replacement.Text = "            "
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                      End With
                    wApp.Selection.Find.Execute Replace:=wdReplaceAll
            Next i
   
   End If
   
    With wApp.Selection.Find
            .Text = "<TOTAL>"
            '.Replacement.Text = Format(lblTotal.Caption, IIf(OptMoneda(0).value, "S/.", "U$") & "#,##0.00")
            .Replacement.Text = IIf(OptMoneda(0).value, "S/.", "U$") & Space(15 - Len(Trim(JDNum(Trim(Val(lblTotal.Caption)), 12, True, 9, 2)))) & JDNum(Trim(Val(lblTotal.Caption)), 12, True, 9, 2)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
    End With
    wApp.Selection.Find.Execute Replace:=wdReplaceAll
                  
       
       
 
          wApp.GoBack
          
      
 

wAppSource.ActiveDocument.Close
wApp.Visible = True

End Sub





Private Sub cmdMas_Click()
If grdDetalle.Row < 18 Then
  grdDetalle.AdicionaFila , , True
  grdDetalle.SetFocus
Else
    MsgBox "SOLO PUEDE INGRESAR INFORMACION PARA 18 ITEMS", vbOKOnly + vbExclamation, "AVISO"
End If
End Sub

Private Sub cmdMenos_Click()
If grdDetalle.Row >= 1 Then
    If (grdDetalle.TextMatrix(grdDetalle.Row, 1) <> "" And grdDetalle.TextMatrix(grdDetalle.Row, 2) <> "" And grdDetalle.TextMatrix(grdDetalle.Row, 3) <> "") Or (grdDetalle.Row <> 1) Then
        grdDetalle.EliminaFila grdDetalle.Row, True
    End If
End If
End Sub

Private Sub cmdNuevo_Click()
 LimpiaInformacion
 txtFecha.Text = Format(gdFecSis, "dd/MM/yyyy")
 fradatos.Enabled = True
 cmdImprimir.Enabled = False
 cmdGrabar.Enabled = True
 TxtNroActa.Text = ""
 Me.lblFechaReg.Caption = ""
 
End Sub

Private Sub LimpiaInformacion()
  LblNroActa.Caption = ""
  lblFechaReg.Caption = ""
  lblCodigoCli.Caption = ""
  lblNomCli.Caption = ""
  lblDICli.Caption = ""
  LblUsuVerificador.Tag = ""
  LblUsuVerificador.Caption = ""
  OptMoneda(0).value = True
  grdDetalle.Clear
  grdDetalle.FormaCabecera
  lblTotal.Caption = ""
  LblEtqMon.Caption = "Total S/."
  lblTotal.BackColor = &HC0FFFF
  txtFecha.Mask = "##/##/####"
  If grdDetalle.Rows > 2 Then
    Dim i As Integer
    
    For i = grdDetalle.Rows - 1 To i = 1 Step -1
        grdDetalle.EliminaFila i
    Next i
  
  End If
  
  
End Sub
Private Sub CmdSalir_Click()
 Unload Me
End Sub

Private Sub cmdVerif_Click()
frmActaLista.Inicia (2)
cmdMas.SetFocus
End Sub

Private Sub Form_Load()
Call ctlusuario.Inicio(gsCodUser)

txtFecha.Text = Format(gdFecSis, "dd/MM/yyyy")

lblUsuDetector.Caption = gsCodUser & ": " & ctlusuario.UserNom
lblUsuDetector.Tag = ctlusuario.PersCod

End Sub

Private Sub grdDetalle_EnterCell()
Dim i As Integer, suma As Double
suma = 0

   Select Case grdDetalle.Col
   
   
   
    Case 3
         For i = 1 To grdDetalle.Rows - 1
            If grdDetalle.TextMatrix(i, 3) <> "" Then
                suma = suma + CDbl(grdDetalle.TextMatrix(i, 3))
            End If
         Next i
         
         lblTotal.Caption = Format(suma, "#,##0.00")
   End Select
   

End Sub



Private Sub grdDetalle_OnCellChange(pnRow As Long, pnCol As Long)
Dim i As Integer, suma As Double
suma = 0

   Select Case pnCol
   
    Case 3
         For i = 1 To grdDetalle.Rows - 1
            If grdDetalle.TextMatrix(i, 3) <> "" Then
                suma = suma + CDbl(grdDetalle.TextMatrix(i, 3))
            End If
         Next i
         
         lblTotal.Caption = Format(suma, "#,##0.00")
   End Select
End Sub

Private Sub grdDetalle_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim i As Integer, suma As Double
suma = 0

'If Trim(grdDetalle.TextMatrix(pnRow, pnCol)) = "" Then
'    MsgBox "No pue"
'    Cancel = True
'Else
'   Select Case pnCol
'    Case 3
'         For I = 1 To grdDetalle.Rows - 1
'            suma = suma + Val(grdDetalle.TextMatrix(I, 3))
'         Next I
'
'         lblTotal.Caption = Format(suma, "#,##0.00")
'
'   End Select
'   Cancel = False
'End If

End Sub

Private Sub OptMoneda_Click(Index As Integer)
   If Index = 0 Then
     LblEtqMon.Caption = "Total S/."
     lblTotal.BackColor = &HC0FFFF
   Else
     LblEtqMon.Caption = "Total U$."
     lblTotal.BackColor = &HC0FFC0
   End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.OptMoneda(0).value = True
    Me.OptMoneda(0).SetFocus
End If
End Sub

Private Sub TxtNroActa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    LimpiaInformacion
    LLenaDatos
    fradatos.Enabled = False
    cmdGrabar.Enabled = False
    cmdImprimir.Enabled = True
    cmdNuevo.Enabled = True
ElseIf KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
End If

End Sub
