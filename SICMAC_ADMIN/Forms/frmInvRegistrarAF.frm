VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmInvRegistrarAF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Activo Fijo"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvRegistrarAF.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   8775
      TabIndex        =   12
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "ACTIVAR BIEN:"
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   11415
      Begin VB.CheckBox chkMejora 
         Caption         =   "Mejora"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDepHist 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   36
         Text            =   "0"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtDepreAcum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   33
         Text            =   "0"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtMesTDeprecia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   28
         Text            =   "0"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtAgeDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         TabIndex        =   25
         Top             =   840
         Width           =   3720
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   9000
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         TabIndex        =   9
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtModelo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         TabIndex        =   8
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtMarca 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         TabIndex        =   7
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtTipoBien 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   3615
      End
      Begin Sicmact.TxtBuscar txtPersona 
         Height          =   315
         Left            =   8640
         TabIndex        =   30
         Top             =   1230
         Width           =   1530
         _extentx        =   2699
         _extenty        =   556
         appearance      =   0
         appearance      =   0
         font            =   "frmInvRegistrarAF.frx":030A
         appearance      =   0
         tipobusqueda    =   3
      End
      Begin MSMask.MaskEdBox mskFecIng 
         Height          =   285
         Left            =   3600
         TabIndex        =   37
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar txtArea 
         Height          =   315
         Left            =   8640
         TabIndex        =   38
         Top             =   480
         Width           =   1530
         _extentx        =   2699
         _extenty        =   556
         appearance      =   0
         appearance      =   0
         font            =   "frmInvRegistrarAF.frx":0336
         appearance      =   0
      End
      Begin VB.TextBox txtMejora 
         Height          =   285
         Left            =   3120
         TabIndex        =   41
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblCategoriaBien 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   43
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblCodInventario 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblnMovItem 
         Caption         =   "Label15"
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblDep 
         Caption         =   "Dep. Hist :"
         Height          =   225
         Left            =   240
         TabIndex        =   35
         Top             =   3240
         Width           =   915
      End
      Begin VB.Label Label14 
         Caption         =   "Depre. Acumulada:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Persona:"
         Height          =   255
         Left            =   5520
         TabIndex        =   32
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblPersonaG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6480
         TabIndex        =   31
         Top             =   1680
         Width           =   3720
      End
      Begin VB.Label Label11 
         Caption         =   "Meses Total Depre."
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblcBSCod 
         Caption         =   "Label11"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblnMovNro 
         Caption         =   "Label11"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Serie:"
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Modelo:"
         Height          =   255
         Left            =   5520
         TabIndex        =   22
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Marca:"
         Height          =   255
         Left            =   5520
         TabIndex        =   21
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Cod. Inventario:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo Bien:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   5520
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "F. Ingreso:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activar Bien"
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11415
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
         Height          =   2265
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   3995
         _Version        =   393216
         Rows            =   21
         Cols            =   10
         ForeColorSel    =   -2147483643
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483637
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   1
         Appearance      =   0
         RowSizingMode   =   1
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Valor de la UIT:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   44
      Top             =   3120
      Width           =   1515
   End
   Begin VB.Label lblValorUIT 
      AutoSize        =   -1  'True
      Caption         =   "ValorUIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   42
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   7560
      TabIndex        =   13
      Top             =   3000
      Width           =   1230
   End
End
Attribute VB_Name = "frmInvRegistrarAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmInvRegistrarAF
'** Descripción : Formulario Registrar los Activos Fijos
'** Creación : MAVM, 20090218 8:59:25 AM
'** Modificación:
'********************************************************************

Option Explicit
Dim sCodAgencia As String
Dim sAgencia As String
Dim iCantidad As Integer
Dim iCantidadInventario As Integer
Dim sBan As String
Dim lcTextoCategoriaBien As String, lcCodCategoriaBien As String '*** PEAC 20100506

Private Sub chkMejora_Click()
    If chkMejora.value = 1 Then
        txtMejora.Visible = True
        lblCodInventario.Visible = False
    Else
        txtMejora.Visible = False
        lblCodInventario.Visible = True
    End If
End Sub

Private Sub Command1_Click()
    Dim lsMovNro As String
    Dim lsBSCod As String
    Dim liMovItem As Integer
    If Mid((fgDetalle.TextMatrix(fgDetalle.Row, 1)), 2, 2) = "12" Then
        lsMovNro = fgDetalle.TextMatrix(fgDetalle.Row, 8)
        lsBSCod = fgDetalle.TextMatrix(fgDetalle.Row, 1)
        liMovItem = CInt(fgDetalle.TextMatrix(fgDetalle.Row, 9))
        If ValidarActivado(lsMovNro, lsBSCod, liMovItem) = False Then
            'GenerarCodInventario lsMovNro, lsBSCod '*** PEAC 20100511
            CargarDatosXAF lsMovNro, lsBSCod, liMovItem
            GenerarCodInventario lsMovNro, lsBSCod '*** PEAC 20100511
        Else
            MsgBox "El Activo Fijo está Registrado"
        End If
    Else
        MsgBox "No es un ActivoFijo", vbCritical
    End If
End Sub

Private Function ValidarActivado(ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Dim rsValidar As ADODB.Recordset
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Set rs = oInventario.DarDatosBBS(sMovNro, sBSCod, iMovItem)
    Set rsValidar = oInventario.DarActivadoXBBS(sMovNro, sBSCod, iMovItem)
    If rs!nMovCant = rsValidar.RecordCount Then
        ValidarActivado = True
    Else
        ValidarActivado = False
    End If
    Set rs = Nothing
    Set rsValidar = Nothing
    Set oInventario = Nothing
End Function

Private Function DevolverCtaCont(ByVal lsBSCod As String) As String
    Dim sCodInventario As String
    sCodInventario = ""
    'Mobiliario
    If lsBSCod = "11200" Then
        sCodInventario = "181301"
    End If
    
    'Equipo de Computo
    If lsBSCod = "11201" Then
        sCodInventario = "181302"
    End If
    
    'Vehiculos
    If lsBSCod = "11202" Then
        sCodInventario = "181401"
    End If
    
    'Terrenos
    If lsBSCod = "11203" Then
        sCodInventario = "181101"
    End If
    
    'Edificios
    If lsBSCod = "11204" Then
        sCodInventario = "181201"
    End If
       
    'Instalaciones
    If lsBSCod = "11205" Then
        sCodInventario = "181202"
    End If
    
    'Mejoras Locales Propios (Edificios)
    If lsBSCod = "11206" Then
        sCodInventario = "181201"
    End If
        
    'If lsBSCod = "11207" Then
    '    sCodInventario = "181201"
    'End If
    
    'Instalaciones en Locales Alquiladas
    If lsBSCod = "11208" Then
        sCodInventario = "181701"
    End If
    
    'mejoras Locales Alquilados
    If lsBSCod = "11209" Then
        sCodInventario = "181702"
    End If
    
    'If lsBSCod = "11210" Then
    '    sCodInventario = "181702"
    'End If
    
    'Maquinarias
    If lsBSCod = "11211" Then
        sCodInventario = "181402"
    End If
    
    DevolverCtaCont = sCodInventario
End Function

Private Function DevolverBan(ByVal lsBSCod As String) As String
    Dim sCodBan As String
    sCodBan = ""
    'Mobiliario
    If lsBSCod = "11200" Then
        sCodBan = "3"
    End If
    
    'Equipo de Computo
    If lsBSCod = "11201" Then
        sCodBan = "2"
    End If
    
    'Vehiculos
    If lsBSCod = "11202" Then
        sCodBan = "1"
    End If
    
    'Terrenos
    If lsBSCod = "11203" Then
        sCodBan = "4"
    End If
    
    'Edificios
    If lsBSCod = "11204" Then
        sCodBan = "5"
    End If
       
    'Instalaciones
    If lsBSCod = "11205" Then
        sCodBan = "6"
    End If
    
    'Mejoras Locales Propios (Edificios)
    If lsBSCod = "11206" Then
        sCodBan = "5"
    End If
        
    'If lsBSCod = "11207" Then
    '    sCodBan = "181201"
    'End If
    
    'Instalaciones en Locales Alquiladas
    If lsBSCod = "11208" Then
        sCodBan = "9"
    End If
    
    'mejoras Locales Alquilados
    If lsBSCod = "11209" Then
        sCodBan = "7"
    End If
    
    'If lsBSCod = "11210" Then
    '    sCodBan = "181702"
    'End If
    
    'Maquinarias
    If lsBSCod = "11211" Then
        sCodBan = "8"
    End If
    DevolverBan = sCodBan
End Function

Private Function DevolverCorrelativo(ByVal lsBSCod As String) As String
    Dim sCorrelativo As String
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Dim rs As ADODB.Recordset
    sCorrelativo = ""
    Set rs = oInventario.DarCorrelativo(lsBSCod)
    sCorrelativo = rs!Maximo + 1
    DevolverCorrelativo = sCorrelativo
    Set rs = Nothing
End Function

Private Sub GenerarCodInventario(ByVal lsMovNro As String, ByVal lsBSCod As String)
    Dim rs As ADODB.Recordset
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Dim sCodInventario As String
    
    Set rs = oInventario.DarAgencia(lsMovNro)
    sCodAgencia = rs!cAreaCod
    sAgencia = rs!cAreaDes
    
    '*** PEAC 20100511
    'sCodInventario = DevolverCtaCont(Mid(lsBSCod, 1, 5)) & "0" & IIf(Mid(rs!cAreaCod, 4, 2) = "", "01", Mid(rs!cAreaCod, 4, 2)) & Format(DevolverCorrelativo(Mid(lsBSCod, 1, 5)), "00000")
    If lcCodCategoriaBien = "1" Then
        sCodInventario = DevolverCtaCont(Mid(lsBSCod, 1, 5)) & "0" & IIf(Mid(rs!cAreaCod, 4, 2) = "", "01", Mid(rs!cAreaCod, 4, 2)) & Format(DevolverCorrelativo(Mid(lsBSCod, 1, 5)), "00000")
    Else
        sCodInventario = "45130111" & "0" & IIf(Mid(rs!cAreaCod, 4, 2) = "", "01", Mid(rs!cAreaCod, 4, 2)) & Format(DevolverCorrelativoNoDepre(Mid(lsBSCod, 1, 5)), "00000")
    End If
    
    '*** FIN PEAC
    
    lblCodInventario.Caption = sCodInventario
       
    Set rs = Nothing
    Set oInventario = Nothing
End Sub

Private Sub CargarDatosXAF(ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer)
    Dim rs As ADODB.Recordset
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Set rs = oInventario.DarDatosBBS(sMovNro, sBSCod, iMovItem)
    iCantidad = rs!nMovCant
    txtNombre.Text = rs!cDescrip
    txtCantidad.Text = "1"
    If rs!nMovCant <> 0 Then
        txtMonto.Text = Format(Round(rs!nMovImporte / rs!nMovCant, 2), gcFormView)
    End If
    txtArea.Text = sCodAgencia
    txtAgeDesc.Text = sAgencia
    lblnMovNro.Caption = sMovNro
    lblcBSCod.Caption = sBSCod
    lblnMovItem.Caption = iMovItem
    
    '*** PEAC 20100506
    If CDbl(txtMonto.Text) >= (CDbl(Me.lblValorUIT.Caption) * rs!nPorcenUIT) Then
        lcTextoCategoriaBien = "BIEN DEPRECIABLE"
        lcCodCategoriaBien = "1"
    Else
        lcTextoCategoriaBien = "BIEN NO DEPRECIABLE"
        lcCodCategoriaBien = "0"
    End If
    Me.lblCategoriaBien.Caption = lcTextoCategoriaBien
    '*** FIN PEAC
    
    CargarTipoBien sBSCod
    sBan = DevolverBan(Mid(sBSCod, 1, 5))
    Set rs = Nothing
    Set oInventario = Nothing
End Sub

Private Sub CargarTipoBien(ByVal sBSCod As String)
    Dim rs As ADODB.Recordset
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Set rs = oInventario.DarTipoBien(Mid(sBSCod, 1, 5))
    txtTipoBien.Text = rs!cBSDescripcion
    Set rs = Nothing
    Set oInventario = Nothing
End Sub

Private Sub Command2_Click()
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    
    Dim rs1 As ADODB.Recordset
    
    If lcCodCategoriaBien = "1" Then 'sie es bien depreciable
        If (txtMesTDeprecia.Text <> "" And txtMesTDeprecia.Text <> "0") Then
            If MsgBox("Esta Seguro de Registar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                oInventario.InsertarAF IIf(chkMejora.value = 1, txtMejora.Text, lblCodInventario.Caption), txtNombre.Text, txtAgeDesc.Text, txtMarca.Text, txtModelo.Text, txtSerie.Text, mskFecIng.Text, lblnMovNro.Caption, lblcBSCod.Caption, CInt(lblnMovItem.Caption)
                RegistrarAF
                LLenarDatos
                MsgBox "Se Registraron Correctamente"
                
                'Dim rs1 As ADODB.Recordset
                Set rs1 = oInventario.DarActivadoXBBS(lblnMovNro.Caption, lblcBSCod.Caption, lblnMovItem.Caption)
                If iCantidad = rs1.RecordCount Then
                    Limpiar
                    Set oInventario = Nothing
                Else
                    If MsgBox("Desea Seguir Registrando?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                        GenerarCodInventario gsMvoNro, Mid(lblcBSCod.Caption, 1, 5)
                        LimpiarAlgunos
                    Else
                        Limpiar
                        Set oInventario = Nothing
                    End If
                End If
            End If
        Else
            MsgBox "Debe Completar Los Datos", vbCritical
        End If
    Else
        If MsgBox("Esta Seguro de Registar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            oInventario.InsertarAF IIf(chkMejora.value = 1, txtMejora.Text, lblCodInventario.Caption), txtNombre.Text, txtAgeDesc.Text, txtMarca.Text, txtModelo.Text, txtSerie.Text, mskFecIng.Text, lblnMovNro.Caption, lblcBSCod.Caption, CInt(lblnMovItem.Caption)
            RegistrarAF
            LLenarDatos
            MsgBox "Se Registraron Correctamente"
            
            'Dim rs1 As ADODB.Recordset
            Set rs1 = oInventario.DarActivadoXBBS(lblnMovNro.Caption, lblcBSCod.Caption, lblnMovItem.Caption)
            If iCantidad = rs1.RecordCount Then
                Limpiar
                Set oInventario = Nothing
            Else
                If MsgBox("Desea Seguir Registrando?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                    GenerarCodInventario gsMvoNro, Mid(lblcBSCod.Caption, 1, 5)
                    LimpiarAlgunos
                Else
                    Limpiar
                    Set oInventario = Nothing
                End If
            End If
        End If
    End If
End Sub

Private Sub RegistrarAF()
    Dim oAF As DMov
    Set oAF = New DMov
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim lnMovNroAjus As Long
    Dim lsMovNroAjus As String
    Dim lsCtaCont As String
    Dim lsAgencia As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim lsCtaOpeBS As String
    Dim lsTipo As String
      
    If Mid(txtArea.Text, 4, 2) = "" Then
        lsAgencia = "01"
    Else
        lsAgencia = Mid(txtArea.Text, 4, 2)
    End If
        
    lsMovNro = oAF.GeneraMovNro(mskFecIng.Text, , gsCodUser)
    oAF.InsertaMov lsMovNro, gnDepAF, "Depre. de ActivoFijo " & Left(Trim(txtNombre.Text), 275)
    lnMovNro = oAF.GetnMovNro(lsMovNro)
    lsMovNroAjus = oAF.GeneraMovNro(mskFecIng.Text, , gsCodUser, lsMovNro)
    oAF.InsertaMov lsMovNroAjus, gnDepAjusteAF, "Deprr. Ajustada de ActivoFijo " & Left(Trim(txtNombre.Text), 272)
    lnMovNroAjus = oAF.GetnMovNro(lsMovNroAjus)
    
    If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaOpeBS = oAF.GetCtaDep(lblcBSCod.Caption)
    oAF.InsertaMovBSActivoFijoUnico Year(mskFecIng.Text), lnMovNro, lblcBSCod.Caption, IIf(chkMejora.value = 1, txtMejora.Text, lblCodInventario.Caption), txtMonto.Text, txtDepreAcum.Text, "0", CDate(mskFecIng.Text), Left(Me.txtArea.Text, 3), lsAgencia, Me.txtMesTDeprecia.Text, "0.00", Left(Trim(Me.txtNombre.Text), 256), "1", "0", "1", lblCodInventario.Caption, lblCodInventario.Caption, CDate(mskFecIng.Text), CDate(mskFecIng.Text), sBan, "0", txtPersona.Text, lcCodCategoriaBien, 0
    ', txtMarca.Text, txtModelo.Text, txtSerie.Text, IIf(chkMejora.value = 1, 1, 0)
    
    'Depre. Historica
    oAF.InsertaMovBSAF Year(gdFecSis), lnMovNro, 1, lblcBSCod.Caption, lblCodInventario.Caption, lnMovNro, sBan
    lsCtaCont = oAF.GetOpeCtaCtaOtro(gnDepAF, lsCtaOpeBS, "", False)
    If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNro, 1, lsCtaCont, CCur(txtDepHist.Text) * -1
    
    If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaCont = oAF.GetOpeCtaCta(gnDepAF, lsCtaOpeBS, "")
    If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNro, 2, lsCtaCont, CCur(Me.txtDepHist.Text)
    
    'Depre. Ajsutada
    If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaCont = oOpe.EmiteOpeCta(gnDepAF, "D")
    oAF.InsertaMovBSAF Year(gdFecSis), lnMovNro, 1, lblcBSCod.Caption, lblCodInventario.Caption, lnMovNroAjus, sBan
    If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNroAjus, 1, lsCtaCont, CCur(txtDepreAcum.Text) * -1
    
    lsCtaCont = oAF.GetOpeCtaCta(gnDepAF, lsCtaOpeBS, "")
    lsCtaCont = Left(lsCtaCont, 2) & "6" & Mid(lsCtaCont, 4, 100)
    If Val(txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNroAjus, 2, lsCtaCont, CCur(txtDepreAcum.Text)
    
    If txtPersona.Text <> "" Then
        oAF.InsertaMovGasto lnMovNro, txtPersona.Text, ""
        oAF.InsertaMovGasto lnMovNroAjus, txtPersona.Text, ""
    End If
            
    Set oAF = Nothing
End Sub

Private Sub LimpiarAlgunos()
    txtPersona.Text = ""
    lblPersonaG.Caption = ""
    txtMarca.Text = ""
    txtModelo.Text = ""
    txtSerie.Text = ""
    chkMejora.value = 0
    txtMejora.Text = ""
    txtMejora.Visible = False
    lblCodInventario.Visible = True
End Sub

Private Sub Limpiar()
    lblCodInventario.Caption = ""
    Me.lblCategoriaBien.Caption = "" '*** PEAC 20100506
    txtNombre.Text = ""
    txtCantidad.Text = ""
    txtMonto.Text = ""
    txtArea.Text = ""
    txtAgeDesc.Text = ""
    txtTipoBien.Text = ""
    txtMarca.Text = ""
    txtModelo.Text = ""
    txtSerie.Text = ""
    lblnMovNro.Caption = ""
    lblcBSCod.Caption = ""
    lblnMovItem.Caption = 0
    iCantidad = 0
    mskFecIng.Text = Date
    sBan = ""
    txtMesTDeprecia.Text = "0"
    txtDepreAcum.Text = "0"
    txtDepHist.Text = "0"
    txtMejora.Text = ""
    chkMejora.value = 0
    txtPersona.Text = ""
    lblPersonaG.Caption = ""
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    mskFecIng.Text = Date
    CargarCabecera
End Sub

Private Function LLenarActivado(ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer) As Integer
    Dim rsInventario As ADODB.Recordset
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Set rsInventario = oInventario.DarActivadoXBBS(sMovNro, sBSCod, iMovItem)
    If rsInventario.RecordCount <> "0" Then
        LLenarActivado = CInt(rsInventario.RecordCount)
    Else
        LLenarActivado = 0
    End If
End Function

Public Sub LLenarDatos()
    Dim rs As ADODB.Recordset
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Dim n As Integer
    Dim i As Integer
    Me.Show
    
    Set rs = oInventario.DarBBSXOC(gsMvoNro)
    
    Me.lblValorUIT.Caption = Format(IIf(gsMvoNro = "", 0, rs!ValorUIT), "#,#0.00")
    
    n = 0
    For i = 0 To rs.RecordCount - 1
        n = n + 1
        fgDetalle.TextMatrix(n, 0) = n
        fgDetalle.TextMatrix(n, 1) = rs!cBSCod
        If Not IsNull(rs!cDescrip) Then
                fgDetalle.TextMatrix(n, 2) = rs!cDescrip
        End If
        fgDetalle.TextMatrix(n, 3) = rs!cConsDescripcion
        fgDetalle.TextMatrix(n, 4) = rs!nMovCant
        If rs!nMovCant <> 0 Then
                fgDetalle.TextMatrix(n, 5) = Format(Round(rs!nMovImporte / rs!nMovCant, 2), gcFormView)
        End If
        fgDetalle.TextMatrix(n, 6) = Format(rs!nMovImporte, gcFormView)
        Dim svar As String
        svar = Mid(rs!cBSCod, 2, 2)
        If svar <> "12" Then
            fgDetalle.TextMatrix(n, 7) = 0
            FlexBackColor fgDetalle, n, "&H00E0E0E0"
        Else
            fgDetalle.TextMatrix(n, 7) = LLenarActivado(gsMvoNro, rs!cBSCod, rs!nMovItem)
        End If
        fgDetalle.TextMatrix(n, 8) = rs!nMovNro
        fgDetalle.TextMatrix(n, 9) = rs!nMovItem
    rs.MoveNext
    Next i
    SumasDoc
    Set rs = Nothing
    Set oInventario = Nothing
End Sub

Private Sub SumasDoc()
    Dim n As Integer
    Dim nTot As Currency
    For n = 1 To fgDetalle.Rows - 1
        If fgDetalle.TextMatrix(n, 6) <> "" Then
           nTot = nTot + Val(Format(fgDetalle.TextMatrix(n, 6), gcFormDato))
        End If
    Next
    If nTot > 0 Then
       txtTot = Format(nTot, gcFormView)
    Else
       txtTot = ""
    End If
End Sub

Private Sub CargarCabecera()
    fgDetalle.TextMatrix(0, 0) = "#"
    fgDetalle.TextMatrix(0, 1) = "Código"
    fgDetalle.TextMatrix(0, 2) = "Descripción"
    fgDetalle.TextMatrix(0, 3) = "Unidad"
    fgDetalle.TextMatrix(0, 4) = "Cantidad"
    fgDetalle.TextMatrix(0, 5) = "P.Unitario"
    fgDetalle.TextMatrix(0, 6) = "Sub Total"
    fgDetalle.TextMatrix(0, 7) = "Activado"
    fgDetalle.TextMatrix(0, 8) = "nMovNro"
    fgDetalle.TextMatrix(0, 9) = "nMovItem"
    
    fgDetalle.ColWidth(0) = 350
    fgDetalle.ColWidth(1) = 1500
    fgDetalle.ColWidth(2) = 3615
    fgDetalle.ColWidth(3) = 900
    fgDetalle.ColWidth(4) = 900
    fgDetalle.ColWidth(5) = 1200
    fgDetalle.ColWidth(6) = 1200
    fgDetalle.ColWidth(7) = 900
    fgDetalle.ColWidth(8) = 0
    fgDetalle.ColWidth(9) = 0
    
    fgDetalle.ColAlignmentFixed(0) = 4
    fgDetalle.ColAlignment(1) = 1
    fgDetalle.ColAlignmentFixed(4) = 7
    fgDetalle.ColAlignmentFixed(5) = 7
    fgDetalle.ColAlignmentFixed(6) = 7
    fgDetalle.ColAlignmentFixed(7) = 7
    fgDetalle.ColAlignmentFixed(8) = 7
    fgDetalle.ColAlignmentFixed(9) = 7
    
    fgDetalle.RowHeight(-1) = 285
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gsMvoNro = ""
End Sub

Private Sub txtPersona_EmiteDatos()
    Me.lblPersonaG.Caption = Me.txtPersona.psDescripcion
End Sub

Private Sub mskFecIng_GotFocus()
    Me.mskFecIng.SelStart = 0
    Me.mskFecIng.SelLength = 50
End Sub

Private Sub mskFecIng_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtNombre.SetFocus
    End If
End Sub

'Private Sub txtAge_EmiteDatos()
'    Me.lblAgeG.Caption = txtAge.psDescripcion
'End Sub
'
'Private Sub txtAge_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.txtPersona.SetFocus
'    End If
'End Sub

'*** PEAC 20100511
Private Function DevolverCorrelativoNoDepre(ByVal lsBSCod As String) As String
    Dim sCorrelativo As String
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Dim rs As ADODB.Recordset
    sCorrelativo = ""
    Set rs = oInventario.DarCorrelativoNoDepre(lsBSCod)
    sCorrelativo = rs!Maximo + 1
    DevolverCorrelativoNoDepre = sCorrelativo
    Set rs = Nothing
End Function

