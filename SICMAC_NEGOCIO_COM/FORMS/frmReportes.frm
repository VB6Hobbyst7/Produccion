VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   6885
   ClientLeft      =   1920
   ClientTop       =   1740
   ClientWidth     =   8445
   HelpContextID   =   210
   Icon            =   "frmReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFTP 
      Caption         =   "FTP"
      Height          =   330
      Left            =   6840
      TabIndex        =   43
      Top             =   6405
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdArchivo 
      Caption         =   "SUCAVE"
      Height          =   330
      Left            =   5880
      TabIndex        =   39
      Top             =   6405
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   330
      Left            =   7680
      TabIndex        =   38
      Top             =   6405
      Width           =   660
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   330
      Left            =   4965
      TabIndex        =   37
      Top             =   6405
      Width           =   840
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   5895
      TabIndex        =   36
      Top             =   6405
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame fraFechaRango 
      Caption         =   "Rango de Fechas"
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
      Height          =   675
      Left            =   5010
      TabIndex        =   26
      Top             =   870
      Visible         =   0   'False
      Width           =   3360
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   300
         Left            =   510
         TabIndex        =   27
         Top             =   255
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   315
         Left            =   2040
         TabIndex        =   28
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   1770
         TabIndex        =   29
         Top             =   315
         Width           =   135
      End
   End
   Begin VB.Frame frmMoneda 
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
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   4965
      TabIndex        =   23
      Top             =   90
      Width           =   3360
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &Nacional"
         Height          =   345
         Index           =   0
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   210
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &Extranjera"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha"
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
      Height          =   675
      Left            =   5010
      TabIndex        =   20
      Top             =   870
      Visible         =   0   'False
      Width           =   1695
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   390
         TabIndex        =   21
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   135
      End
   End
   Begin VB.Frame fraAge 
      Caption         =   "Agencia"
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
      Height          =   1575
      Left            =   4965
      TabIndex        =   3
      Top             =   3225
      Visible         =   0   'False
      Width           =   3360
      Begin VB.ListBox lstAge 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   3105
      End
   End
   Begin MSComctlLib.TreeView tvOpe 
      Height          =   6660
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   11748
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imglstFiguras"
      BorderStyle     =   1
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
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   180
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":030A
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":065C
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":09AE
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":0D00
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraFuentes 
      Caption         =   "Fuentes de Financiamiento"
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
      Height          =   1545
      Left            =   4965
      TabIndex        =   1
      Top             =   4815
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ListBox lstFuentes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   300
         Width           =   3105
      End
   End
   Begin VB.Frame fraTCambio 
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
      Height          =   690
      Left            =   4995
      TabIndex        =   14
      Top             =   1875
      Visible         =   0   'False
      Width           =   3330
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
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
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "A&justado"
         Height          =   255
         Index           =   3
         Left            =   4500
         TabIndex        =   16
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtTipCambio2 
         Alignment       =   1  'Right Justify
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
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "T.C. Cierre Mes"
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
         Left            =   60
         TabIndex        =   19
         Top             =   705
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "T.C. del Mes"
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
         Left            =   60
         TabIndex        =   18
         Top             =   255
         Width           =   1020
      End
   End
   Begin VB.Frame fraTC 
      Caption         =   "Tipo de Cambio"
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
      Height          =   1395
      Left            =   5010
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   3360
      Begin VB.TextBox TxtTipCamFijAnt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   360
         TabIndex        =   9
         Text            =   "0"
         Top             =   390
         Width           =   1125
      End
      Begin VB.TextBox txtTipCamCompraSBS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   360
         TabIndex        =   8
         Text            =   "0"
         Top             =   945
         Width           =   1125
      End
      Begin VB.TextBox txtTipCamFij 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   1920
         TabIndex        =   7
         Text            =   "0"
         Top             =   390
         Width           =   1005
      End
      Begin VB.TextBox txtTipCamVentaSBS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   1920
         TabIndex        =   6
         Text            =   "0"
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label lblTipcambiAnt 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Fijo.Ant"
         Height          =   195
         Left            =   510
         TabIndex        =   13
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Comp.SBS"
         Height          =   195
         Left            =   390
         TabIndex        =   12
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label lblTipcamFij 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Fijo.Act"
         Height          =   195
         Left            =   2010
         TabIndex        =   11
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Venta.SBS"
         Height          =   195
         Left            =   1875
         TabIndex        =   10
         Top             =   735
         Width           =   1080
      End
   End
   Begin VB.Frame fraEuros 
      Caption         =   "Euros"
      Height          =   750
      Left            =   5040
      TabIndex        =   40
      Top             =   990
      Visible         =   0   'False
      Width           =   3360
      Begin VB.TextBox txtTCEuros 
         Height          =   285
         Left            =   1215
         TabIndex        =   41
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "T.C. Euros"
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
         Left            =   165
         TabIndex        =   42
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Periodo"
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
      Height          =   1035
      Left            =   5010
      TabIndex        =   31
      Top             =   870
      Visible         =   0   'False
      Width           =   3360
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1575
         MaxLength       =   4
         TabIndex        =   33
         Top             =   615
         Width           =   1095
      End
      Begin VB.ComboBox cboMes 
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
         ItemData        =   "frmReportes.frx":1052
         Left            =   870
         List            =   "frmReportes.frx":107A
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   255
         Width           =   1815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   285
         TabIndex        =   35
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   285
         TabIndex        =   34
         Top             =   330
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lExpand  As Boolean
Dim lExpandO As Boolean
Dim sArea  As String
Dim sTipoRepo As Integer
Dim Progress As clsProgressBar
'Dim Index760010 As Integer ' *** MAVM: Auditoria ANALISIS DE CUENTAS
'
'Dim Index760300 As Integer ' *** MAVM: Auditoria REPORTE DE PAGO A PROVEEDORES
'Dim Index760303 As Integer ' *** MAVM: Auditoria REPORTE DE PAGO A PROVEEDORES
'
'Dim Index461010 As Integer ' *** MAVM: Auditoria CONSULTA DE SALDOS
'Dim Index461011 As Integer ' *** MAVM: Auditoria CONSULTA DE SALDOS
'
'Dim Index461090 As Integer ' *** MAVM: Auditoria CONSULTA DE SALDOS ADEUDOS
'Dim Index461091 As Integer ' *** MAVM: Auditoria CONSULTA DE SALDOS ADEUDOS
'
'Dim PiOperacion As Integer ' *** MAVM: Auditoria

Public Sub Inicio(sObj As String, ByVal iOperacion As Integer, Optional plExpandO As Boolean = False)
'    PiOperacion = iOperacion
    sArea = sObj
    lExpandO = plExpandO
    Me.Show 0, MDISicmact
End Sub

Private Function validaDatos() As Boolean
Dim i As Integer
validaDatos = False
   If fraFechaRango.Visible Then
      If Not ValFecha(txtFechaDel) Then
         txtFechaDel.SetFocus: Exit Function
      End If
      If Not ValFecha(txtFechaAl) Then
         txtFechaAl.SetFocus: Exit Function
      End If
   End If
   If fraFecha.Visible Then
      If Not ValFecha(txtFecha) Then
         txtFecha.SetFocus: Exit Function
      End If
   End If
   If fraPeriodo.Visible Then
      If nVal(txtAnio) = 0 Then
         MsgBox "Ingrese Año para generar Reporte...", vbInformation, "¡Aviso!"
         txtAnio.SetFocus
         Exit Function
      End If
      If cboMes.ListIndex = -1 Then
        MsgBox "Selecciones Mes para generar Reporte...", vbInformation, "¡Aviso!"
        cboMes.SetFocus
        Exit Function
      End If
   End If
   If Not tvOpe.SelectedItem.Child Is Nothing Then
        MsgBox "Seleccione Reporte de último Nivel", vbInformation, "¡Aviso!"
        tvOpe.SetFocus
        Exit Function
   End If
   
   validaDatos = True
    
   If fraAge.Visible = True Then
        validaDatos = False
        For i = 0 To lstAge.ListCount - 1
            If lstAge.Selected(i) Then
                validaDatos = True
            End If
        Next
        If validaDatos = False Then
            MsgBox "Ud. debe seleccionar al menos una agencia", vbInformation, "Aviso"
            lstAge.SetFocus
            Exit Function
        End If
          
'        lstAge.Selected(lstAge.ListCount - 1) = False
          
    End If
   
    If validaDatos = False Then
        Exit Function
    End If
    
    If fraFuentes.Visible = True Then
        validaDatos = False
        For i = 0 To lstFuentes.ListCount - 1
            If lstFuentes.Selected(i) Then
                validaDatos = True
            End If
        Next
        If validaDatos = False Then
            MsgBox "Ud. debe seleccionar al menos una fuente", vbInformation, "Aviso"
            lstFuentes.SetFocus
            Exit Function
        End If
    End If
End Function

Private Sub CboMes_Click()
If Me.fraTCambio.Visible Then
    txtTipCambio = TipoCambioCierre(txtAnio, cboMes.ListIndex + 1)
End If
End Sub

Public Function TipoCambioCierre(pnAnio As Integer, pnMes As Integer, Optional pbMesCerrado As Boolean = True) As Currency
Dim oCambio As nTipoCambio
Dim sFecha  As Date
    If pnMes <= 0 Or pnMes > 12 Or pnAnio < 1900 Then
        Exit Function
    End If
    sFecha = CDate("01/" & Format(pnMes, "00") & "/" & Trim(pnAnio))
    sFecha = DateAdd("m", 1, sFecha)
    If Not pbMesCerrado Then
       sFecha = sFecha - 1
    End If
    Set oCambio = New nTipoCambio
    TipoCambioCierre = Format(oCambio.EmiteTipoCambio(sFecha, TCFijoMes), "#,##0.0000")
    Set oCambio = Nothing
End Function

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fraTCambio.Visible Then
            txtTipCambio = TipoCambioCierre(txtAnio, cboMes.ListIndex + 1)
        End If
        txtAnio.SetFocus
    End If
End Sub

'Private Sub cmdFTP_Click()
'Dim Opt As Integer
'
'    Select Case Mid(gsOpeCod, 1, 6)
'        Case RepCGEncBCRObligacion
'            frmAnxEncajeBCR.GeneraAnx01MN txtFechaAl
'
'        Case RepCGEncBCRObligacionME
'            frmAnxEncajeBCR.GeneraAnx01ME txtFechaAl
'
'        Case RepCGEncBCRCredDeposi
''            Opt = MsgBox("Desea Configurar el Anexo para su exportacion", vbInformation + vbYesNo, "AVISO")
''            If Opt = vbNo Then
'                frmAnxEncajeBCR.GeneraAnx02MN txtFechaAl
''            Else
''                FrmAnxencajeBCRAnexo2.lblFecha = txtFechaAl
''                FrmAnxencajeBCRAnexo2.OptSoles.value = True
''                FrmAnxencajeBCRAnexo2.Show
''            End If
'
'
'        Case RepCGEncBCRCredDeposiME
''            Opt = MsgBox("Desea Configurar el Anexo para su exportacion", vbInformation + vbYesNo, "AVISO")
''            If Opt = vbNo Then
'                frmAnxEncajeBCR.GeneraAnx02ME txtFechaAl
''            Else
''                FrmAnxencajeBCRAnexo2.lblFecha = txtFechaAl
''                FrmAnxencajeBCRAnexo2.OptDolares.value = True
''                FrmAnxencajeBCRAnexo2.Show
''            End If
'
'
'        Case RepCGEncBCRCredRecibi
'            frmAnxEncajeBCR.GeneraAnx03MN txtFechaAl
'
'        Case RepCGEncBCRCredRecibiME
'            frmAnxEncajeBCR.GeneraAnx03ME txtFechaAl
'
'        Case RepCGEncBCRObligaExon
'            frmAnxEncajeBCR.GeneraAnx04MN txtFechaAl
'
'        Case RepCGEncBCRObligaExonME
'            frmAnxEncajeBCR.GeneraAnx04ME txtFechaAl
'
'    End Select
'End Sub

Private Sub cmdGenerar_Click()
Dim lsMoneda As String
Dim ldFecha  As Date
Dim sFuente As String
Dim i As Integer
'Dim oEstadistica As COMNAuditoria.NEstadistica


'On Error GoTo cmdGenerarErr
    If tvOpe.Nodes.Count = 0 Then
        MsgBox "Lista de operaciones se encuentra vacia", vbInformation, "Aviso"
        Exit Sub
    End If
    If tvOpe.SelectedItem.Tag = "1" Then
        MsgBox "Operación seleccionada no valida...!", vbInformation, "Aviso"
        tvOpe.SetFocus
        Exit Sub
    End If
    If Not validaDatos Then Exit Sub
    
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    lsMoneda = IIf(optMoneda(0).value, "1", "2")
    If Me.fraPeriodo.Visible Then
        ldFecha = CDate(DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio, "0000"))) - 1
    End If
    If lsMoneda = "1" Then
        gsSimbolo = gcMN
    Else
        gsSimbolo = gcME
    End If
    If Left(tvOpe.SelectedItem.Key, 1) <> "P" Then
        gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescHijo = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescPadre = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60))
    Else
      gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    End If
    
    'gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    
'    Select Case Mid(gsOpeCod, 1, 4)
'        Case Mid(gContRepBaseFormula, 1, 4)
'            frmRepBaseFormula.Inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo
'
'    End Select

    Select Case Mid(gsOpeCod, 1, 6)
'
'        '***************** CARTAS FIANZA ********************************
'        Case OpeCGCartaFianzaRepIngreso, OpeCGCartaFianzaRepIngresoME
'            ImprimeCartasFianza txtFechaDel, txtFechaAl, True
'        Case OpeCGCartaFianzaRepSalida, OpeCGCartaFianzaRepSalidaME
'            ImprimeCartasFianza txtFechaDel, txtFechaAl, False
'
'        '********************* REPORTES DE CAJA GENERAL **************************
'        Case OpeCGRepFlujoDiarioResMN, OpeCGRepFlujoDiarioResME
'            ResumenFlujoDiario Me.txtFecha
'        Case OpeCGRepFlujoDiarioDetMN, OpeCGRepFlujoDiarioDetME
'            DetalleFlujoDiario txtFechaDel, txtFechaAl
'
        Case OpeCGRepRepBancosFlujoMensMN, OpeCGRepRepBancosFlujoMensME
            frmCajaGenRepFlujos.Show 0, Me
'        Case OpeCGRepRepBancosFlujoPFMN, OpeCGRepRepBancosFlujoPFME
'            frmAdeudRepDet.Inicio False
'        Case OpeCGRepRepBancosResumenPFMN, OpeCGRepRepBancosResumenPFME
'             frmPFReportes.Show 1
'
'        Case OpeCGRepRepBancosConcentFdos
'             ImprimeConcentracionFondos txtFecha.Text, Val(TxtTipCambio.Text)
'        Case OpeCGRepRepCMACSFlujoMensMN, OpeCGRepRepCMACSFlujoMensME
'            frmCajaGenRepFlujos.Show 0, Me
'        Case OpeCGRepRepCMACSFlujoPFMN, OpeCGRepRepCMACSFlujoPFME
'            frmAdeudRepDet.Inicio False
'
'        Case OpeCGRepRepCMACSResumenPFMN, OpeCGRepRepCMACSResumenPFME
'            frmPFReportes.Show 1
'
'        'Orden de Pago
'        Case OpeCGRepRepOPGirMN, OpeCGRepRepOPGirME
'             ReporteOrdenesPago txtFechaDel, txtFechaAl, gsOpeCod
'
'        'A rendir Cuenta
'        Case OpeCGRepArendirLibroAuxMN, OpeCGRepArendirLibroAuxME
'             ReporteArendirCuentaLibro txtFechaDel, txtFechaAl, gsOpeCod
'
'        Case OpeCGRepArendirViaticoLibroAuxMN, OpeCGRepArendirViaticoLibroAuxME
'             ReporteArendirCuentaViaticosLibro txtFechaDel, txtFechaAl, gsOpeCod
'
'        Case OpeCGRepArendirPendienteMN, OpeCGRepArendirPendienteME
'             ReporteARendirCuentaPendientes gsOpeCod, CInt(lsMoneda), txtFecha
'
'        Case OpeCGRepArendirViaticosMN, OpeCGRepArendirViaticosME
'             frmRepViaticos.Show 1
'
'        'Cheques
'        Case OpeCGRepRepChqRecDetMN, OpeCGRepRepChqRecDetME
'            ResumenCheques txtFecha, Mid(gsOpeCod, 1, 6), True
'        'Remesas Cheques
'        Case OpeCGRepChequesEnvMN, OpeCGRepChequesEnvME
'            ResumenChequesRem txtFechaDel, txtFechaAl, gsOpeCod
'
'        Case OpeCGRepChequesAnulMN, OpeCGRepChequesAnulME
'            ResumenChequesAnul txtFechaDel, txtFechaAl, gsOpeCod
'
'        Case OpeCGRepChequesCobMN, OpeCGRepChequesCobME
'            ResumenChequesCob txtFechaDel, txtFechaAl, gsOpeCod
'
'        Case OpeCGRepRepChqRecResMN, OpeCGRepRepChqRecResME
'            ResumenCheques txtFecha, Mid(gsOpeCod, 1, 6)
'
'        Case OpeCGRepChqRecibidoCajaMN, OpeCGRepChqRecibidoCajaME
'            ResumenChqRecibidos txtFecha, gsOpeCod
'        Case OpeCGRepChqDepositadoMN, OpeCGRepChqDepositadoME
'            ResumenChqDepositados txtFecha, gsOpeCod
'
'        'ADEUDADOS
        Case 461091, 462091
            'frmAdeudRepGen.Inicio
'        Case OpeCGAdeudRepDetalleMN, OpeCGADeudRepDetalleME
'            frmAdeudRepDet.Inicio True
'        Case OpeCGAdeudRepSaldLinFinancDescalceMN
'            ReporteSaldosLineaFinanciamientoDescalce gdFecSis
'        Case OpeCGAdeudRepCortoPlazoMN, OpeCGAdeudRepCortoPlazoME
'            frmAdeudRepGen.Inicio True
'        Case OpeCGAdeudRepxFecVenc, OpeCGAdeudRepxFecVencME
'            frmAdeudRepVenc.Inicio
'
'        'Adeudados vinculados
'        Case OpeCGAdeudRepVinculadosMN
'             ReporteAdeudadosVinculados OpeCGAdeudRepVinculadosMN
'
'        Case OpeCGAdeudRepVinculadosME
'             If txtTCEuros.Text = "" Then
'                MsgBox "Ingrese Tipo de Cambio Euros", vbInformation, "Aviso"
'                txtTCEuros.SetFocus
'                Exit Sub
'             End If
'             ReporteAdeudadosVinculados OpeCGAdeudRepVinculadosME, txtTCEuros.Text
'
'
'
'        Case OpeCGRepPresuFlujoCaja, OpeCGRepPresuFlujoCajaME
'            frmpresFlujoCaja.Show 1
'        Case OpeCGRepPresuServDeuda, OpeCGRepPresuServDeudaME
'            frmPresServDeuda.ImprimeServicioDeuda txtAnio, ldFecha, TxtTipCambio
'        Case OpeCGRepPresuFinancia, OpeCGRepPresuFinanciaME
'            frmPresFinanExtInt.ImprimeReporteFinanciamientoIE txtAnio, ldFecha, TxtTipCambio
'        Case OpeCGRepOtrBilletesFalsosMN, OpeCGRepOtrBilletesFalsosME
'            'frmReporteMonedaFalsa.lsMoneda = lsMoneda
'            frmReporteMonedaFalsa.Show 1
'        'ENCAJE
'        Case OpeCGRepEncajeConsolSdoEnc, OpeCGRepEncajeConsolSdoEncME
'            frmCajaGenReportes.ConsolidaSdoEnc gsOpeCod, txtFechaDel, txtFechaAl
'        Case OpeCGRepEncajeAgencia, OpeCGRepEncajeAgenciaME
'            frmSdoEncaje.CalculaSdoEnc lsMoneda, txtFechaDel, txtFechaAl
'
'        Case OpeCGRepEncajeConsolPosLiq, OpeCGRepEncajeConsolPosLiqME
'            frmCajaGenReportes.ConsolidaSdoEnc gsOpeCod, txtFechaDel, txtFechaAl, 2
'
'        Case OpeCGRepEstadOpeUsuario, OpeCGRepEstadOpeUsuarioME
'            frmRepOpeDiaUsuario.Show 1
'
'        Case OpeEstEncajeSimulacionPlaEncajeMN, OpeEstEncajeSimulacionPlaEncajeME
'            frmCGSimuladorPlanillaEncaje.Show 1
'
'
'
'        'Informe de Encaje al BCR
'         Case RepCGEncBCRObligacion, RepCGEncBCRObligacionME, _
'              RepCGEncBCRCredDeposi, RepCGEncBCRCredDeposiME, _
'              RepCGEncBCRCredRecibi, RepCGEncBCRCredRecibiME, _
'              RepCGEncBCRObligaExon, RepCGEncBCRObligaExonME, _
'              RepCGEncBCRLinCredExt, RepCGEncBCRLinCredExtME
'            frmRepEncajeBCR.ImprimeEncajeBCR gsOpeCod, txtFechaDel, txtFechaAl, CDbl(TxtTipCambio.Text), CDbl(txtTipCambio2.Text)
'
'        ' Saldos de Caja - Bancos y Agencias
'        Case OpeCGRepSaldoBcos, OpeCGRepSaldoCajAge
'            frmRptDiaLiquidez.Inicio Mid(gsOpeCod, 1, 6)

        
        '************************* CNTABILIDAD **********************************
'        Case gContLibroDiario
'               frmContabDiario.Show 0, Me
'        Case 760010
'               frmContabMayor.Show 0, Me
        Case 760010
               'frmContabMayorDet.Show 0, Me
'        Case gContLibInvBal
'            frmLibroInventBalanc.Show 0, Me
'        Case gContLibroCaja
'            frmLibroCaja.Show 0, Me
'
'        Case gConvFoncodes
'              frmRptFoncodes.Show 0, Me
'
'        Case gContRegCompraGastos
'               frmRegCompraGastos.Show 0, Me
'        Case 760201 'LIMA
'               frmRegCompraGastos.PorDocumento True
'               frmRegCompraGastos.Show 0, Me
'
'        Case gContRegVentas
'               frmRegVenta.Show 0, Me
'
'        Case gContRepEstadIngGastos
'               EstadisticaIngresoGasto cboMes.ListIndex + 1, txtAnio
'        Case gContRepPlanillaPagoProv
'               frmRepPagProv.Show 0, Me
'        Case gContRepControlGastoProv
'               frmProveeConsulMov.Show 1, Me
        Case 760303
               frmRepProvCaja.Show 1, Me
'        Case gContRepCompraVenta
'            frmRepResCVenta.Show 0, Me
'
'        Case gContRepEstadProv
'
'            frmRepEstadProv.Show 1, Me
'
'        'Otros Ajustes
'        Case gContAjReclasiCartera
'            frmAjusteReCartera.Show 0, Me
'        Case gContAjReclasiGaranti
'            frmAjusteGarantias.Show 0, Me
'        Case gContAjInteresDevenga
'            frmAjusteIntDevengado.Inicio True
'        Case gContAjInteresSuspens
'            frmAjusteIntDevengado.Inicio False
'
'        Case 701228
'            frmAjusteIntDevengado.Inicio 11
'        Case 702228
'            frmAjusteIntDevengado.Inicio 11
'
'
'        'Riesgos
'        Case gRiesgoCalfCarCred:
'                    Call frmRiesgosReportes.Inicio(gRiesgoCalfCarCred, gdFecSis)
'                    frmRiesgosReportes.Show 1
'        Case gRiesgoCalfAltoRiesgo:
'                    Call frmRiesgosReportes.Inicio(gRiesgoCalfAltoRiesgo, gdFecSis)
'                    frmRiesgosReportes.Show 1
'        Case gRiesgoConceCarCred:
'                    Call frmRiesgosReportes.Inicio(gRiesgoConceCarCred, gdFecSis)
'                    frmRiesgosReportes.Show 1
'        Case gRiesgoEstratDepPlazo:
'                    Call frmRiesgosReportes.Inicio(gRiesgoEstratDepPlazo, gdFecSis)
'                    frmRiesgosReportes.Show 1
'        Case gRiesgoPrincipClientesAhorros:
'                    Call frmRiesgosReportes.Inicio(gRiesgoPrincipClientesAhorros, gdFecSis)
'                    frmRiesgosReportes.Show 1
'        Case gRiesgoPrincipClientesCreditos:
'                    Call frmRiesgosReportes.Inicio(gRiesgoPrincipClientesCreditos, gdFecSis)
'                    frmRiesgosReportes.Show 1
'
'
'        '''captaciones
'        Case gContRepCaptacCVMonExtr
'            Consolida763201 gbBitCentral, Me.tvOpe.SelectedItem.Text, Me.lstAge, CDate(txtFechaDel.Text), CDate(txtFechaAl.Text)
'        Case gContRepCaptacSituacCaptac
'            Imprime763202 gbBitCentral, Me.tvOpe.SelectedItem.Text, Me.lstAge, CDate(txtFecha.Text)
'        Case gContRepCaptacMovCV
'            Imprime763203 Me.tvOpe.SelectedItem.Text, CDate(txtFechaDel.Text), TxtTipCamFijAnt, txtTipCamCompraSBS, txtTipCamVentaSBS, gbBitCentral
'
'        Case gContRepCaptacIngPagos, gContRepCaptacCredDesem
'            lstAge.Selected(lstAge.ListCount - 1) = False
'            For i = 0 To lstFuentes.ListCount - 1
'                If lstFuentes.Selected(i) = True Then
'                    If Len(Trim(sFuente)) = 0 Then
'                        sFuente = "'" & Right(lstFuentes.List(i), 1) & "'"
'                    Else
'                        sFuente = sFuente & ", '" & Right(lstFuentes.List(i), 1) & "'"
'                    End If
'                End If
'            Next
'            If Mid(gsOpeCod, 1, 6) = gContRepCaptacIngPagos Then
'                Imprime763204 Me.tvOpe.SelectedItem.Text, Trim(txtFecha.Text), sFuente, Me.lstAge, gbBitCentral
'            Else
'                Imprime763205 Me.tvOpe.SelectedItem.Text, Trim(txtFechaDel.Text), Trim(txtFechaAl.Text), sFuente, lstAge, gbBitCentral
'            End If
'
'        'Movimientos Inusuales
'        Case gContRepInusuales
'             frmRepRiesgos.Show 1
'
'        'Otros Reportes
'        Case gProvContxCtasCont
'             ProvContxCtasCont txtFechaDel, txtFechaAl
'
'        Case gProvContxOpe
'             ProvContxOperacion txtFechaDel, txtFechaAl
'
'        Case gIntDevPFxFxAG
'             InteresDevvengadoPFxFxAG txtFechaDel, txtFechaAl
'
'        Case gIntDevCTSxFxAG
'             InteresCTSxFxAG txtFechaDel, txtFechaAl
'
'        Case gPlazoFijoRango
'             PlazoFijoxRango Me.txtFecha, lsMoneda
'
'        Case gRepInstPubliResgos
'             RepInstPubliResgos
'
'        Case gITFQuincena
'             ITFQuincena txtFechaDel, txtFechaAl
'
'        Case gPlazoFijoIntCash    'By Capi 04082008
'             ImprimePlazoFijoIntCash lsMoneda
'
'        'By Capi 08022008
'        Case gCarteraDetxIntDSP
'             ImprimeCarteraDetxIntDSP Val(TxtTipCambio)
'        Case gCarteraResxLineas
'             ImprimeCarteraResxLineas Val(TxtTipCambio)
'
'        'End By
'
'        'JEOM
'        'Reportes Balance Contabilidad
'        Case gPignoraticiosVencidos
'             PignoraticiosVencidos
'
'        Case gCarteraCreditos
'             CarteraCreditos Me.txtFecha
'
'        Case gCarteraInteres
'             InteresCreditos Me.txtFecha
'
'        Case gCreditosCastigados
'             CreditosCastigados Me.txtFecha
'
'        Case gInteresesDiferidos
'             InteresesDiferidos Me.txtFecha
'
'        Case gPignoraticiosVigentes
'             PignoraticiosVigentes
'
'
'        Case gCreditosCondonados
'             CreditosCondonados Me.txtFecha
'
'        Case gProvisionCreditos
'             ProvisionCreditos
'
'        Case gIntCreditosRefinanciados
'             IntCreditosRefinanciados Me.txtFechaDel, Me.txtFechaAl
'
'        Case gProvisionCartasFianzas
'              ProvisionCartasFianzas
'
'        Case gDetalleGarantias
'             DetalleGarantias Me.txtFecha, TxtTipCambio, lsMoneda
'
'        'JEOM
'        'Reportes Balance Planeamiento
'        Case gCarteraCredPla
'             CarteraCreditosPla Me.txtFecha
'
'        Case gCredDesembolsosPla
'             CredDesembolsosPla Me.txtFecha, Me.TxtTipCambio
'
'        Case gCarteraVencidaPla
'             CarteraVencidaPla Me.txtFecha
'
'        Case gCarteraRefinanciadaPla
'             CarteraRefinanciadaPla Me.txtFecha
'
'        Case gCarteraJudicialesPla
'             CarteraJudicialesPla Me.txtFecha
'
'
'        'By Capi 18122007 Para Planeamiento
'        Case gCarteraRecupCapital
'             SayCarteraRecupCapital Me.txtFecha
'
'        'JEOM
'        'Reportes Balance Riesgos
'        Case gCredVigentesRiesgos
'             CreditosVigentes Me.txtFecha
'
'        Case gCredRefinanciadosRiesgos
'             CreditosRefinanciados Me.txtFecha
'
'        Case gPlazoFijoRiesgos
'             PlazoFijoRiesgos Me.txtFecha, lsMoneda
'
'
'        'ANEXOS
'        Case gContAnx02CredTpoGarantia 'Creditos Directos por Tipo de Garantia
'            frmAnx02CreDirGarantia.GeneraAnx02CreditosTipoGarantia txtAnio, cboMes.ListIndex + 1, nVal(TxtTipCambio), cboMes.Text, 1
'        Case gContAnx03FujoCrediticio
'            frmAnx02CreDirGarantia.GeneraAnx03FlujoCrediticioPorTipoCred gbBitCentral, txtAnio, cboMes.ListIndex + 1, nVal(TxtTipCambio), cboMes.Text, 1
'
'        Case gContAnx07N
'            frmAnexo7RiesgoInteres.Inicio True
'        Case gContAnx10DepColocaPer 'Depositos, Colocaciones y Persona por Oficinas
'            frmAnx10DepColocPers.Show 0, frmReportes
'
'        '''''''''''''''''''
'         Case gContAnx09
'             frmRepBaseFormula.Inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo
'
'         Case gContAnx11MovDepsMonto
'
'            'Anterior txt
'            '------------
'            ''''ImprimeAnx11 gbBitCentral, Me.tvOpe.SelectedItem.Text, CDate(txtFecha.Text), Me.lstAge, Val(TxtTipCamFijAnt.Text), Val(txtTipCamFij.Text), Val(txtTipCamCompraSBS.Text), Val(txtTipCamVentaSBS.Text)
'
'            ImprimeAnx11xls gbBitCentral, Me.tvOpe.SelectedItem.Text, CDate(txtFecha.Text), Me.lstAge, Val(TxtTipCamFijAnt.Text), Val(txtTipCamFij.Text), Val(txtTipCamCompraSBS.Text), Val(txtTipCamVentaSBS.Text)
'
'        Case gContAnx13DepsSEscMonto
'
'            'Anterior txt
'            '------------
'            ''''ImprimeAnx13 Me.tvOpe.SelectedItem.Text, Val(txtTipCamFij.Text), gbBitCentral
'
'            ImprimeAnx13XLS Me.tvOpe.SelectedItem.Text, Val(Me.TxtTipCambio.Text), gbBitCentral, ldFecha
'
'        Case gContAnx13DepsSEscMonto_Nuevo
'
'            'Nuevo
'            '------------
'            ImprimeAnx13XLS Me.tvOpe.SelectedItem.Text, Val(Me.TxtTipCambio.Text), gbBitCentral, ldFecha
'
'
'        Case gContAnx15A_Estad      'Informe Estadístico
'            frmAnx15AEstadDia.ImprimeEstadisticaDiaria gsOpeCod, lsMoneda, txtFecha
'        Case gContAnx15A_Efect      'Descomposición de Efectivo
'            frmAnx15AEfectivoCaja.ImprimeEfectivoCaja gsOpeCod, lsMoneda, txtFecha
'        Case gContAnx15A_Banco      'Consolidado Bancos
'            frmAnx15AConsolBancos.ImprimeConsolidaBancos gsOpeCod, lsMoneda, txtFecha
'        Case gContAnx15A_Repor      'Anexo 15A
'            'frmAnx15AReporte.ImprimeAnexo15A gsOpeCod, lsMoneda, txtfecha
'            frmAnx15AReporteNew.ImprimeAnexo15A gsOpeCod, lsMoneda, txtFecha
'        Case gContAnx15B
'            frmAnexo15BPosicionLiquidez.Inicia 1, 0, Me
'        Case gContAnx16LiqVenc
'            frmAnexo16LiquidezVenc.Inicio False
'        Case gContAnx16A
'            frmAnexo16ALiquidezVenc.Inicio False
'        Case gContAnx16B
'            frmAnexo16LiquidezVenc.Inicio False
'        Case gContAnx17A_FSD
'            frmFondoSeguroDep.Inicio txtFechaDel, txtFechaAl
'        Case 770250
'            Anexo6RFA Format(cboMes.ListIndex + 1, "00"), txtAnio, cboMes
'
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Case gContAnx17CtasFuncionarios
'            Imprime770175 ldFecha, gbBitCentral, Me.tvOpe.SelectedItem.Text
'          Case gContAnx17ListadoFSD
'             Imprime770130 gbBitCentral, Me.tvOpe.SelectedItem.Text, Val(txtTipCamFij.Text), txtFecha.Text
'        Case gContAnx17ListadoGenCtas
'
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        'REPORTES
'        Case gRiesgoSBSA02A
'             frmRepBaseFormula.Inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo, "CAR", False
'        Case gRiesgoSBSA02B
'             frmRepBaseFormula.Inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo
'
'        Case gRiesgoSBSA050
'
'            'Anterior txt
'            '------------
''            ImprimeRep05SBS cboMes.ListIndex + 1, txtAnio
'
'            ImprimeRep05SBSXls cboMes.ListIndex + 1, txtAnio
'
'        Case gContRep06Crediticio
'             frmRep6Crediticio.ImprimeAnexo6Crediticio gsOpeCod, ldFecha
'
'        Case gRepCreditosIncumplidos 'Reporte 14
'             frmRep14CredIncump.ImprimeReporte14 gsOpeCod, ldFecha, Val(TxtTipCambio.Text)
'
'
'        Case gRiesgoSBSA190
'            Call frmRiesgosGrupEconom.Inicio(gRiesgoSBSA190, gdFecSis)
'            frmRiesgosGrupEconom.Show 1
'
'        Case gRiesgoSBSA191
'            Call frmRiesgosGrupEconom.Inicio(gRiesgoSBSA191, gdFecSis)
'            frmRiesgosGrupEconom.Show 1
'
'        Case gRiesgoSBSA200
'            Call frmRiesgosGrupEconom.Inicio(gRiesgoSBSA200, gdFecSis)
'            frmRiesgosGrupEconom.Show 1
'
'        Case gRiesgoSBSA201
'            Call frmRiesgosGrupEconom.Inicio(gRiesgoSBSA201, gdFecSis)
'            frmRiesgosGrupEconom.Show 1
'
'        Case gRiesgoSBSA210
'            Call frmRiesgosGrupEconom.Inicio(gRiesgoSBSA210, gdFecSis)
'            frmRiesgosGrupEconom.Show 1
'
'        Case 780130
'            frmReporte16.Show 1
'
'        Case gPatrEfecAjxInfla
'            ImprimeRep3SBS_PatrimEfectAjustxInfl 760112, CInt(txtAnio.Text), cboMes.ListIndex + 1
'        Case gContAnx18
'            ImprimeAnx18SBS_InmMovEquip 760181, CInt(txtAnio.Text), cboMes.ListIndex + 1
'        Case gContAnx24
'            ImprimeAnx24_CTS gbBitCentral, cboMes.ListIndex + 1, Val(txtAnio.Text), Val(TxtTipCambio.Text)
'        Case 780220
'            Set Progress = New clsProgressBar
'            Progress.ShowForm Me
'            Progress.Max = 100
'            Progress.Progress 10, "Procesando Reporte de Estadistica Adelantada....."
'            Set oEstadistica = New COMNAuditoria.NEstadistica
'            Call oEstadistica.MuestraEstadisticaAdelantada(gdFecSis, Me.txtTipCambio)
'            Progress.Max = 100
'            Progress.Progress 100, "Procesando Reporte de Estadistica Adelantada....."
'            Set oEstadistica = Nothing
'            Progress.CloseForm Me
'            MsgBox "Reporte Generado Satisfactoriamente.....", vbInformation, "Aviso"

        'By Capi 18122007 Para Planeamiento
'        Case gCarteraRecupCapital
'             SayCarteraRecupCapital Me.txtFecha
        
  
    End Select
Exit Sub
cmdGenerarErr:
    MsgBox TextErr(err.Description), vbInformation, "Aviso"
End Sub



Private Sub cmdsalir_Click()
    Unload Me
End Sub
 
Private Sub cmdArchivo_Click()
Dim lsMoneda As String
Dim ldFecha  As Date
Dim sFuente As String
Dim i As Integer

On Error GoTo cmdArchivoErr
    If tvOpe.Nodes.Count = 0 Then
        MsgBox "Lista de operaciones se encuentra vacia", vbInformation, "Aviso"
        Exit Sub
    End If
    If tvOpe.SelectedItem.Tag = "1" Then
        MsgBox "Operación seleccionada no valida...!", vbInformation, "Aviso"
        tvOpe.SetFocus
        Exit Sub
    End If
    If Not validaDatos Then Exit Sub
    
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    lsMoneda = IIf(optMoneda(0).value, "1", "2")
    If Me.fraPeriodo.Visible Then
        ldFecha = CDate(DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio, "0000"))) - 1
    End If
    If lsMoneda = "1" Then
        gsSimbolo = gcMN
    Else
        gsSimbolo = gcME
    End If
    If Left(tvOpe.SelectedItem.Key, 1) <> "P" Then
        gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescHijo = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescPadre = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60))
    Else
      gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    End If
   

Exit Sub
cmdArchivoErr:
    MsgBox TextErr(err.Description), vbInformation, "Aviso"
End Sub
 

Private Sub Form_Activate()
    If tvOpe.Enabled And tvOpe.Visible Then
        tvOpe.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim sCod As String
    On Error GoTo error
    CentraForm Me
    'frmMdiMain.Enabled = False
    
     
    Dim oCambio As COMDConstSistema.NCOMTipoCambio
    Set oCambio = New COMDConstSistema.NCOMTipoCambio
    GetTipCambio gdFecSis
    TxtTipCamFijAnt = Format(gnTipCambio, "#,##0.0000")
    txtTipCamFij = Format(oCambio.EmiteTipoCambio(DateAdd("m", 1, gdFecSis), TCFijoMes), "#,##0.0000")
    txtTipCambio = txtTipCamFij
    txtTipCamCompraSBS = "0.0000"
    txtTipCamVentaSBS = "0.0000"
    
    Dim oConst As New NConstSistemas
    If Not lExpandO Then
       sCod = oConst.LeeConstSistema(gConstSistContraerListaOpe)
       If sCod <> "" Then
         lExpand = IIf(UCase(Trim(sCod)) = "FALSE", False, True)
       End If
    Else
       lExpand = lExpandO
    End If
    
    
    LoadOpeUsu "2"
    LoadAgencia
    txtFecha = gdFecSis
    txtAnio = Year(gdFecSis)
    cboMes.ListIndex = Month(gdFecSis) - 1
    
    txtFechaDel = CDate("01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000"))
    txtFechaAl = gdFecSis
    
    If gbBitCentral = True Then
       'txtFecha = oConst.LeeConstSistema(gConstSistCierreMesNegocio)
       txtFecha = gdFecSis
    Else
        Dim oCon As New DConecta
        Dim rsCierre As ADODB.Recordset
        oCon.AbreConexion
        'If oCon.AbreConexionRemota(gsCodAge, False, False) Then
            sCod = "Select cNomVar, cValorVar From VarSistema Where (cCodProd = 'AHO' And cNomVar IN ('cDBAhoCont','cServAhoCont'))  OR (cCodProd = 'ADM' And cNomVar IN ('dFecCierreMes'))"
            Set rsCierre = oCon.CargaRecordSet(sCod)
            If Not rsCierre.EOF Then
                txtFecha = CDate(Trim(rsCierre!cValorVar))
            End If
        'End If
        'oCon.CierraConexion
        oCon.AbreConexion
    End If
    
    cboMes.ListIndex = Month(txtFecha) - 1
    txtAnio = Year(txtFecha)
    RSClose rsCierre
    
    CentraForm Me
    
    Exit Sub
error:
    MsgBox TextErr(err.Description), vbExclamation, Me.Caption
End Sub


Sub LoadAgencia()
    Dim sqlAge As String
    Dim rsAge As ADODB.Recordset
    Dim oCon As DConecta
    Dim gcCentralCom As String

    Set oCon = New DConecta

    oCon.AbreConexion
    gcCentralCom = "DBCOMUNES.."

    Set rsAge = New ADODB.Recordset
    rsAge.CursorLocation = adUseClient

    If gbBitCentral = True Then
        sqlAge = "Select cAgeDescripcion cNomtab, cAgeCod cValor From Agencias"
    Else
        sqlAge = "Select cAgeDescripcion cNomtab, '112' + cAgeCod cValor From Agencias"
    End If
    Set rsAge = oCon.CargaRecordSet(sqlAge)


    lstAge.Clear

    If Not RSVacio(rsAge) Then
        While Not rsAge.EOF
            lstAge.AddItem Trim(rsAge!cNomtab) & Space(500) & Trim(rsAge!cValor)
            rsAge.MoveNext
        Wend
        lstAge.AddItem "Consolidado" & Space(500) & "CONSOL"
    End If

    rsAge.Close
    Set rsAge = Nothing

    'Fuentes de Financiamiento

    If Not gbBitCentral = True Then
        Set rsAge = New ADODB.Recordset
        rsAge.CursorLocation = adUseClient

        sqlAge = "select cvalor, cNomTab from dbcomunes..TablaCod where ccodtab like '22%'  and cvalor<>''"
        Set rsAge = oCon.CargaRecordSet(sqlAge)


        lstFuentes.Clear
        While Not rsAge.EOF
            lstFuentes.AddItem Trim(rsAge!cNomtab) & Space(500) & Trim(rsAge!cValor)
            rsAge.MoveNext
        Wend
        rsAge.Close
        Set rsAge = Nothing
    Else
        Set rsAge = New ADODB.Recordset
        rsAge.CursorLocation = adUseClient

        sqlAge = "select clineacred as cvalor, cdescripcion as cNomTab from coloclineacredito where  len(clineacred)=2  "
        
''''        sqlAge = "SELECT P.cPersCod as cValor, cPersNombre as cNomTab "
''''        sqlAge = sqlAge & " FROM InstitucionFinanc IFc INNER JOIN PERSONA P "
''''        sqlAge = sqlAge & " ON IFc.cPersCod=P.cPersCod "
''''        sqlAge = sqlAge & " where IFc.cIFTpo='05' Order by cPersNombre "
        
        Set rsAge = oCon.CargaRecordSet(sqlAge)
 
        lstFuentes.Clear
             While Not rsAge.EOF
                lstFuentes.AddItem Trim(rsAge!cNomtab) & Space(500) & Trim(rsAge!cValor)
                rsAge.MoveNext
            Wend
        rsAge.Close
        Set rsAge = Nothing
    End If
  
End Sub

'***MAVM: Modulo de Auditoria 20/08/2008
' Para Mostrar seleccioando El Reporte de ANALISIS DE CUENTAS
' por defecto en el Modulo de Auditoria
'Public Sub Inicializar_operacion()
'    tvOpe.Nodes(Index760010).Selected = True
'    tvOpe_NodeClick tvOpe.SelectedItem
'    tvOpe.Nodes(Index760010).Expanded = True
'
'    tvOpe.Enabled = False
'    tvOpe.HideSelection = False
'End Sub
'***MAVM: Modulo de Auditoria 20/08/2008

'***MAVM: Modulo de Auditoria 02/10/2008
' Para Mostrar seleccioando El Reporte de PAGO PROVEEDORES
' por defecto en el Modulo de Auditoria
'Public Sub Inicializar_Operacion_PagoProveedores()
'    tvOpe.Nodes(Index760300).Selected = True
'    tvOpe_NodeClick tvOpe.SelectedItem
'    tvOpe.Nodes(Index760300).Expanded = True
'
'    tvOpe.Nodes(Index760303).Selected = True
'    tvOpe_NodeClick tvOpe.SelectedItem
'    tvOpe.Nodes(Index760303).Expanded = True
'
'    tvOpe.Enabled = False
'    tvOpe.HideSelection = False
'End Sub

'Public Sub Inicializar_Operacion_ConsultaSaldos()
'    tvOpe.Nodes(Index461010).Selected = True
'    tvOpe_NodeClick tvOpe.SelectedItem
'    tvOpe.Nodes(Index461010).Expanded = True
'
'    tvOpe.Nodes(Index461011).Selected = True
'    tvOpe_NodeClick tvOpe.SelectedItem
'    tvOpe.Nodes(Index461011).Expanded = True
'
'    tvOpe.Enabled = False
'    tvOpe.HideSelection = False
'End Sub

'Public Sub Inicializar_Operacion_ConsultaSaldosAdeudos()
'    tvOpe.Nodes(Index461090).Selected = True
'    tvOpe_NodeClick tvOpe.SelectedItem
'    tvOpe.Nodes(Index461090).Expanded = True
'
'    tvOpe.Nodes(Index461091).Selected = True
'    tvOpe_NodeClick tvOpe.SelectedItem
'    tvOpe.Nodes(Index461091).Expanded = True
'
'    tvOpe.Enabled = False
'    tvOpe.HideSelection = False
'End Sub

'*************************

Sub LoadOpeUsu(psMoneda As String)

Dim clsGen As COMDAuditoria.DCOMRevision
Dim i As Integer
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node

Set clsGen = New COMDAuditoria.DCOMRevision
Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, sArea, psMoneda)
Set clsGen = Nothing
tvOpe.Nodes.Clear
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
            nodOpe.Tag = sOpeCod
    End Select
    i = i + 1 ' ***MAVM:Auditoria
'    If sOpeCod = "760010" Then Index760010 = i ' ***MAVM:Auditoria
'
'    If sOpeCod = "760300" Then Index760300 = i ' ***MAVM:Auditoria
'    If sOpeCod = "760303" Then Index760303 = i ' ***MAVM:Auditoria
    
'    If sOpeCod = "461010" Or sOpeCod = "462010" Then Index461010 = i ' ***MAVM:Auditoria
'    If sOpeCod = "461011" Or sOpeCod = "462011" Then Index461011 = i ' ***MAVM:Auditoria
'
'    If sOpeCod = "461090" Or sOpeCod = "462090" Then Index461090 = i ' ***MAVM:Auditoria
'    If sOpeCod = "461091" Or sOpeCod = "462091" Then Index461091 = i ' ***MAVM:Auditoria
    
    nodOpe.Expanded = lExpand
    rsUsu.MoveNext
Loop
RSClose rsUsu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmMdiMain.Enabled = True
End Sub


Private Sub OptMoneda_Click(Index As Integer)
    Dim sDig As String
    Dim sCod As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    On Error GoTo error
    If optMoneda(0) Then
        sDig = "2"
    Else
        sDig = "1"
    End If
    oConec.AbreConexion
    LoadOpeUsu sDig
    oConec.CierraConexion
'    If PiOperacion = 1 Then Inicializar_operacion
'    If PiOperacion = 2 Then Inicializar_Operacion_PagoProveedores
'    If PiOperacion = 3 Then Inicializar_Operacion_ConsultaSaldos
'    If PiOperacion = 4 Then Inicializar_Operacion_ConsultaSaldosAdeudos
    Exit Sub
error:
    MsgBox TextErr(err.Description), vbExclamation, Me.Caption
End Sub

Private Sub ActivaControles(Optional plFechaRango As Boolean = True, _
                           Optional plFechaAl As Boolean = False, _
                           Optional plFechaPeriodo As Boolean = False, _
                           Optional plTpoCambio As Boolean = False, _
                           Optional plMoneda As Boolean = True, _
                           Optional plTC As Boolean = False, _
                           Optional plAgencia As Boolean = False, _
                           Optional plFuentes As Boolean = False, _
                           Optional cmdSUCAVE As Boolean = False, _
                           Optional plUsuario As Boolean = False _
                           )
fraFechaRango.Visible = plFechaRango
fraFecha.Visible = plFechaAl
fraPeriodo.Visible = plFechaPeriodo
fraTCambio.Visible = plTpoCambio
fraTCambio.Height = 690
txtTipCambio2.Visible = False

If plFechaAl Then
    txtFecha.SetFocus
End If
If plFechaRango Then
    txtFechaDel.SetFocus
End If
fraPeriodo.Enabled = True
frmMoneda.Visible = plMoneda
If gsOpeCod = "770030" Then
    cmdArchivo.Visible = cmdSUCAVE
End If
fraTC.Visible = plTC
fraAge.Visible = plAgencia
fraFuentes.Visible = plFuentes
Blanquea
End Sub

Private Sub Blanquea()
Dim i As Integer

For i = 0 To lstAge.ListCount - 1
    lstAge.Selected(i) = False
Next
For i = 0 To lstFuentes.ListCount - 1
    lstFuentes.Selected(i) = False
Next

End Sub

Private Sub tvOpe_Click()
On Error Resume Next
tvOpe.SelectedItem.ForeColor = vbRed
End Sub

Private Sub tvOpe_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 42 Or KeyCode = 38 Or KeyCode = 40 Then
    tvOpe.SelectedItem.ForeColor = "&H80000008"
End If
End Sub

Private Sub tvOpe_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    tvOpe.SelectedItem.ForeColor = "&H80000008"
End Sub

Private Sub tvOpe_NodeClick(ByVal Node As MSComctlLib.Node)
    tvOpe.SelectedItem.ForeColor = vbRed
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    cmdGenerar.Caption = "&Generar"
    cmdImprimir.Visible = False
    fraEuros.Visible = False
    cmdArchivo.Visible = False
    cmdFTP.Visible = False
    
    Select Case Mid(gsOpeCod, 1, 4)
        Case Mid(gContRepBaseFormula, 1, 4)
            ActivaControles False, False, True
    Case Else
    Select Case Mid(gsOpeCod, 1, 6)
        Case OpeCGRepFlujoDiarioResMN, OpeCGRepFlujoDiarioResME
            ActivaControles False, True
        Case OpeCGRepFlujoDiarioDetMN, OpeCGRepFlujoDiarioDetME
            ActivaControles True
        
        '***************** CARTAS FIANZA ********************************
        Case OpeCGCartaFianzaRepIngreso, OpeCGCartaFianzaRepIngresoME
            ActivaControles True
            txtFechaDel.SetFocus
            
        Case OpeCGCartaFianzaRepSalida, OpeCGCartaFianzaRepSalidaME
            ActivaControles True
            txtFechaDel.SetFocus
            
        '********************* REPORTES DE CAJA GENERAL **************************
        Case OpeCGRepRepBancosFlujoMensMN, OpeCGRepRepBancosFlujoMensME
            ActivaControles False
        Case OpeCGRepRepBancosFlujoPFMN, OpeCGRepRepBancosFlujoPFME
            ActivaControles False
'        Case OpeCGRepRepBancosResumenPFMN, OpeCGRepRepBancosResumenPFME
'            ActivaControles False
'        Case OpeCGRepRepBancosConcentFdos
'            ActivaControles False, True, , True, False
        Case OpeCGRepRepCMACSFlujoMensMN, OpeCGRepRepCMACSFlujoMensME
            ActivaControles False
        Case OpeCGRepRepCMACSFlujoPFMN, OpeCGRepRepCMACSFlujoPFME
            ActivaControles False
'        Case OpeCGRepRepCMACSResumenPFMN, OpeCGRepRepCMACSResumenPFME
'            ActivaControles False
        Case OpeCGRepRepOPGirMN, OpeCGRepRepOPGirME
            ActivaControles True
        'A rendir Cuenta
'        Case OpeCGRepArendirLibroAuxMN, OpeCGRepArendirLibroAuxME
'            ActivaControles True, False
'        Case OpeCGRepArendirPendienteMN, OpeCGRepArendirPendienteME
'            ActivaControles False, True
            
'        Case OpeCGRepArendirViaticoLibroAuxMN, OpeCGRepArendirViaticoLibroAuxME
'            ActivaControles True, False
'        Descomentar
'        Case 461045 'Reporten Sustentacion Viaticos GITU
'            frmReporteSustViaticos.Show 1
        'Remesa de Cheques
'        Case OpeCGRepChequesEnvMN, OpeCGRepChequesEnvME
'            ActivaControles True
            
'        Case OpeCGRepChequesAnulMN, OpeCGRepChequesAnulME
'            ActivaControles True
        
        Case OpeCGRepRepChqRecDetMN, OpeCGRepRepChqRecDetME
            ActivaControles False, True
        Case OpeCGRepRepChqRecResMN, OpeCGRepRepChqRecResME
            ActivaControles False, True
        
        Case OpeCGRepRepChqValDetMN, OpeCGRepRepChqValDetME
            ActivaControles False
        Case OpeCGRepRepChqValResMN, OpeCGRepRepChqValResME
            ActivaControles False, True
        
        Case OpeCGRepRepChqValorizadosDetMN, OpeCGRepRepChqValorizadosDetME
            ActivaControles False
        Case OpeCGRepRepChqValorizadosResMN, OpeCGRepRepChqValorizadosResME
            ActivaControles False
        
        Case OpeCGRepRepChqAnulDetMN, OpeCGRepRepChqAnulDetME
            ActivaControles False
        Case OpeCGRepRepChqAnulResMN, OpeCGRepRepChqAnulResME
            ActivaControles False

        Case OpeCGRepRepChqObsDetMN, OpeCGRepRepChqObsDetME
            ActivaControles False
        Case OpeCGRepRepChqObsResMN, OpeCGRepRepChqObsResME
            ActivaControles False
        Case OpeCGRepChqRecibidoCajaMN, OpeCGRepChqRecibidoCajaME
            ActivaControles False, True
'        Case OpeCGRepChqDepositadoMN, OpeCGRepChqDepositadoME
'            ActivaControles False, True
        
        'Adeudados
'        Case OpeCGAdeudRepSaldLinFinancDescalceMN
'            ActivaControles False, False
         
'        Case OpeCGAdeudRepVinculadosME
'             ActivaControles False, False
'             fraEuros.Visible = True
'             txtTCEuros.SetFocus
             'txtTCEuros.Text = gnTipoCambioEuro
 
             
        'Presupuesto
'        Case OpeCGRepPresuFlujoCaja, OpeCGRepPresuFlujoCajaME
'            ActivaControles False, False, False
'        Case OpeCGRepPresuServDeuda, OpeCGRepPresuServDeudaME
'            ActivaControles False, False, True, True
'        Case OpeCGRepPresuFinancia, OpeCGRepPresuFinanciaME
'            ActivaControles False, False, True, True
            
         'ENCAJE
        Case OpeCGRepEncajeConsolSdoEnc, OpeCGRepEncajeConsolSdoEncME, OpeCGRepEncajeAgencia, OpeCGRepEncajeAgenciaME, OpeCGRepEncajeConsolPosLiq, OpeCGRepEncajeConsolPosLiqME
            ActivaControles True
            txtFechaDel.SetFocus
        Case OpeCGRepEncajeAgencia, OpeCGRepEncajeAgenciaME
            ActivaControles True
            txtFechaDel.SetFocus
                               
       
'        Case OpeEstEncajeSimulacionPlaEncajeMN, OpeEstEncajeSimulacionPlaEncajeME
''              frmCGSimuladorPlanillaEncaje.Show 1
            
            
        'Informe de Encaje al BCR
         Case RepCGEncBCRObligacion, RepCGEncBCRObligacionME, RepCGEncBCRCredDeposi, RepCGEncBCRCredDeposiME, RepCGEncBCRCredRecibi, RepCGEncBCRCredRecibiME, RepCGEncBCRObligaExon, RepCGEncBCRObligaExonME, RepCGEncBCRLinCredExt, RepCGEncBCRLinCredExtME
            If Mid(gsOpeCod, 1, 6) = RepCGEncBCRObligacion Or Mid(gsOpeCod, 1, 6) = RepCGEncBCRCredDeposi Or Mid(gsOpeCod, 1, 6) = RepCGEncBCRCredRecibi Or Mid(gsOpeCod, 1, 6) = RepCGEncBCRObligaExon Or Mid(gsOpeCod, 1, 6) = RepCGEncBCRLinCredExt Then
                ActivaControles True
                cmdFTP.Visible = True
                If gsOpeCod = 761201 Then
                   cmdArchivo.Visible = True
                End If
            Else
                ActivaControles True, , , True
                fraTCambio.Height = 1095
                txtTipCambio2.Visible = True
                cmdFTP.Visible = True
                If gsOpeCod = 762201 Then
                   cmdArchivo.Visible = True
                End If
            End If
            
            txtFechaDel.SetFocus
         ' saldos de caja - bancos y agencias
'         Case OpeCGRepSaldoBcos, OpeCGRepSaldoCajAge
'               ActivaControles False, False, False, False, False, False, False, False, False

        '************************* CONTABILIDAD **********************************
'        Case gContLibroDiario
'            ActivaControles False
'        Case gContLibroMayor
'            ActivaControles False
        Case gContLibroMayCta
            ActivaControles False
'        Case gConvFoncodes
'            ActivaControles False, False, False, False, False, False, False, False, False
       
        Case gContRegCompraGastos
            ActivaControles False
        Case gContRegVentas
            ActivaControles False
        
        Case gContRepEstadIngGastos
            ActivaControles False, False, True

'        Case gContRepCompraVenta
'            ActivaControles True
            
        Case gContRepPlanillaPagoProv
            ActivaControles False
        Case gContRepControlGastoProv
            ActivaControles False
        
'        Case gContRepCompraVenta
'            ActivaControles False

        'Otros Ajustes
'        Case gContAjReclasiCartera
'            ActivaControles False
'        Case gContAjReclasiGaranti
'            ActivaControles False
'        Case gContAjInteresDevenga
'            ActivaControles False
'        Case gContAjInteresSuspens
'            ActivaControles False
            
                    
        'CAPTACIONES
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
'        Case gContRepCaptacCVMonExtr
'            ActivaControles True, False, False, False, False, True, True
'            cmdGenerar.Caption = "&Consolida"
'            cmdImprimir.Visible = True
'        Case gContRepCaptacSituacCaptac
'            ActivaControles False, True, False, False, False, True, False
'        Case gContRepCaptacMovCV
'            ActivaControles True, False, False, False, False, True, False
'        Case gContRepCaptacIngPagos
'            ActivaControles False, True, False, False, False, False, True, True
'        Case gContRepCaptacCredDesem
'            ActivaControles True, False, False, False, False, False, True, True
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' OTROS REPORTES
'       Case gProvContxCtasCont
'            ActivaControles True, False, False, False, , , , , False
             
'        Case gProvContxOpe
'            ActivaControles True, False, False, False, , , , , False
             
'        Case gIntDevPFxFxAG
'            ActivaControles True, False, False, False, , , , , False
             
'        Case gIntDevCTSxFxAG
'            ActivaControles True, False, False, False, , , , , False
             
'        Case gPlazoFijoRango
'            ActivaControles False, True, False, False, True, , , , False
'        Case gRepInstPubliResgos
'            ActivaControles False, False, False, False, False, , , , False
'
'        Case gPlazoFijoRango
'            ActivaControles False, True, False, False, True, , , , False
'
'        Case gITFQuincena
'            ActivaControles True, False, False, False, False, , , , False
            
        'By Capi 04012008
'        Case gPlazoFijoIntCash
'            ActivaControles False, False, False, False, True, , , , False
        
        'By Capi 08022008
'        Case gCarteraDetxIntDSP
'            ActivaControles False, False, False, True, False, , , , False
'        Case gCarteraResxLineas
'            ActivaControles False, False, False, True, False, , , , False
        'End by
       'JEOM Reportes Balance contabilidad
'        Case gCarteraCreditos
'             ActivaControles False, True, False, False, False, , , , False
'
'        Case gCarteraInteres
'             ActivaControles False, True, False, False, False, , , , False
'
'        Case gCreditosCastigados
'             ActivaControles False, True, False, False, False, , , , False
'
'        Case gInteresesDiferidos
'             ActivaControles False, True, False, False, False, , , , False
''
''        Case gPignoraticiosVigentes
''             ActivaControles True, False, False, False, False, , , , False
'
'        Case gCreditosCondonados
'             ActivaControles False, True, False, False, False, , , , False
'
'        Case gIntCreditosRefinanciados
'             ActivaControles True, False, False, False, False, , , , False
'
'        Case gDetalleGarantias
'             ActivaControles False, True, False, True, True, , , , False
'        ''''''''''''''''''''''''''''''''''
'        'JEOM Reportes Balance Planeamiento
'        Case gCredDesembolsosPla
'             ActivaControles False, True, False, True, False, , , , False
'
'        Case gCarteraVencidaPla
'             ActivaControles False, True, False, False, False, , , , False
'
'        Case gCarteraRefinanciadaPla
'             ActivaControles False, True, False, False, False, , , , False
'
'        Case gCarteraJudicialesPla
'             ActivaControles False, True, False, False, False, , , , False
'
'        'By Capi 18122007 Para Planeamiento
'        Case gCarteraRecupCapital
'             ActivaControles False, True, False, False, False, , , , False
'
'                ''''''''''''''''''''''''''''''''''
'        'JEOM Reportes Balance Riesgos
'        Case gCredVigentesRiesgos
'             ActivaControles False, True, False, False, False, , , , False
'
'        Case gCredRefinanciadosRiesgos
'             ActivaControles False, True, False, False, False, , , , False
'
'        Case gPlazoFijoRiesgos
'             ActivaControles False, True, False, False, True, , , , False
'
'        'ANEXOS
'        Case gContAnx02CredTpoGarantia, gContAnx03FujoCrediticio
'
'            If Mid(gsOpeCod, 1, 6) = gContAnx02CredTpoGarantia Then
'                ActivaControles False, False, True, True, , , , , True
'            Else
'                ActivaControles False, False, True, True, , , , , True
'            End If
'
'        Case gContAnx07
'            ActivaControles False
'        Case gContAnx09
'            ActivaControles False, False, True
'
'        Case gContAnx10DepColocaPer
'            ActivaControles False
'
'        '''''''''''''''''''''''''
'        Case gContAnx11MovDepsMonto
'            ActivaControles False, True, False, False, False, True, True
'
'        Case gContAnx13DepsSEscMonto
'            ActivaControles False, False, True, True, False, False
'
'        Case gContAnx13DepsSEscMonto_Nuevo
'            ActivaControles False, False, True, True, False, False
'
'        Case gContAnx17ListadoFSD
'            ActivaControles False, True, False, False, False, True
'        Case gContAnx17ListadoGenCtas
'
'        Case gContAnx17CtasFuncionarios
'            ActivaControles False, False, True, False, False, False, False
'
'        '''''''''''''''''''''''''
'
'        Case gContAnx15A_Estad, gContAnx15A_Efect, gContAnx15A_Banco, gContAnx15A_Repor
'            ActivaControles False, True, , , , , , , True
'            txtFecha.SetFocus
'        Case gContAnx15B
'            ActivaControles False, False, , , , , , , True
'        Case gContAnx16LiqVenc
'            ActivaControles False
'        Case gContAnx16A
'            ActivaControles False
'        Case gContAnx16B
'            ActivaControles False
'        Case gContAnx17A_FSD
'            ActivaControles True, , , False
'            txtFechaDel.SetFocus
'        Case gContAnx18
'            ActivaControles False, False, True, False, False
'            cboMes.SetFocus
'        Case 770250
'            ActivaControles False, False, True, False, False
'            cboMes.SetFocus
'
'        'Reportes SBS
'        Case gRiesgoSBSA02A, gRiesgoSBSA02B
'            ActivaControles False, False, True, False, False, False, False, False
'
'        Case gRiesgoSBSA050
'            ActivaControles False, False, True, False, False, False, False, False
'
'        Case gContRep06Crediticio
'            ActivaControles False, False, True, , , , , , True
'
'        Case gRepCreditosIncumplidos 'Reporte 14
'            ActivaControles False, False, True, True, False, False, False, False, True
'
'        Case gPatrEfecAjxInfla
'            ActivaControles False, False, True, False
'            cboMes.SetFocus
'        Case gContAnx24
'            ActivaControles False, False, True, True, False
'
'        Case 780220
'            ActivaControles False, False, False, True, False, False, False, False, False
            
        Case Else
            ActivaControles False, False, False
   End Select
   End Select

End Sub

Private Sub tvOpe_Collapse(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H80000008"
End Sub

Private Sub tvOpe_Expand(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H8000000D"
End Sub

Private Sub tvOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdGenerar_Click
    End If
End Sub

Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
Dim oCambio As nTipoCambio
Dim sFecha  As Date
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If fraTCambio.Visible Then
        sFecha = "01/" & IIf(Len(Trim(cboMes.ListIndex + 1)) = 1, "0" & Trim(str(cboMes.ListIndex + 1)), Trim(cboMes.ListIndex + 1)) & "/" & Trim(txtAnio.Text)
        Set oCambio = New nTipoCambio
        If Len(Trim(cboMes.Text)) > 0 And val(txtAnio.Text) > 1900 Then
            txtTipCambio.Text = Format(oCambio.EmiteTipoCambio(sFecha, TCFijoDia), "#,##0.0000")
        End If
        txtTipCambio.SetFocus
    Else
        cmdGenerar.SetFocus
    End If
End If
End Sub
 

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) = True Then
       If fraTCambio.Visible = True Then
        txtTipCambio.SetFocus
       Else
        cmdGenerar.SetFocus
    End If
    End If
End If
End Sub

Private Sub txtFecha_LostFocus()
Dim oCambio As nTipoCambio
Dim sFecha As String
If fraTCambio.Visible = True Then
    'sFecha = DateAdd("m", 1, "01/" & IIf(Len(Trim(CboMes.ListIndex + 1)) = 1, "0" & Trim(Str(CboMes.ListIndex + 1)), Trim(CboMes.ListIndex + 1)) & "/" & Trim(txtAnio.Text))
    'sFecha = DateAdd("d", -1, sFecha)
    If Not IsDate(txtFecha) Then
        Me.txtFecha.SetFocus
        Exit Sub
    End If
    Set oCambio = New nTipoCambio
    txtTipCambio.Text = Format(oCambio.EmiteTipoCambio(Trim(txtFecha.Text), TCFijoDia), "#,##0.0000")
ElseIf fraTC.Visible = True Then
    'txtTipCamFij = Format(gnTipCambio, "#,##0.0000")
    'TxtTipCamFijAnt = Format(oCambio.EmiteTipoCambio(DateAdd("m", -1, gdFecSis), TCFijoMes), "#,##0.0000")
    If Not IsDate(txtFecha) Then
        Me.txtFecha.SetFocus
        Exit Sub
    End If
    Set oCambio = New nTipoCambio
    txtTipCamFij = Format(oCambio.EmiteTipoCambio(txtFecha, TCFijoMes), "#,##0.0000")
    TxtTipCamFijAnt = Format(oCambio.EmiteTipoCambio(DateAdd("m", -1, txtFecha), TCFijoMes), "#,##0.0000")
End If
End Sub

Private Sub txtFechaAl_Validate(Cancel As Boolean)
Dim oTC As New nTipoCambio
If IsDate(txtFechaAl.Text) Then
Else
    txtFechaAl = Format(gdFecSis, "dd/MM/YYYY")
End If
    txtTipCambio = Format(oTC.EmiteTipoCambio(txtFechaAl, TCFijoMes), "#,##0.00###")
    txtTipCambio2 = Format(oTC.EmiteTipoCambio(CDate(txtFechaAl) + 1, TCFijoMes), "#,##0.00###")
      
End Sub

Private Sub txtFechaDel_GotFocus()
fEnfoque txtFechaDel
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaDel) = True Then
       txtFechaAl.SetFocus
    End If
End If
End Sub

Private Sub txtFechaAl_GotFocus()
    fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaAl) = True Then
       cmdGenerar.SetFocus
    End If
End If
End Sub

Private Sub txtTCEuros_KeyPress(KeyAscii As Integer)
 KeyAscii = NumerosDecimales(txtTCEuros, KeyAscii, 14, 6)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub

Private Sub txtTipCambio_GotFocus()
fEnfoque txtTipCambio
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 14, 5)
If KeyAscii = 13 Then
    If Not txtTipCambio2.Visible Then
        cmdGenerar.SetFocus
    Else
        txtTipCambio2.SetFocus
    End If
End If
End Sub

Private Sub txtTipCambio2_GotFocus()
    fEnfoque txtTipCambio2
End Sub

Private Sub txtTipCambio2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTipCambio2, KeyAscii, , 3)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub



