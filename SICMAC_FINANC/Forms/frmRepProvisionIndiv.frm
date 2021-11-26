VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRepProvisionIndiv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Provisión Individualizada"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12465
   Icon            =   "frmRepProvisionIndiv.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   12465
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstProvisiones 
      Height          =   4455
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7858
      _Version        =   393216
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "Interés Devengado"
      TabPicture(0)   =   "frmRepProvisionIndiv.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblMontoIntDev"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMontoText"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "feIntDevengado"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Interés de Suspenso"
      TabPicture(1)   =   "frmRepProvisionIndiv.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape2"
      Tab(1).Control(1)=   "lblMontoIntSus"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "feIntSuspenso"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Provisión de Cartera"
      TabPicture(2)   =   "frmRepProvisionIndiv.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape3"
      Tab(2).Control(1)=   "lblMontoProvCart"
      Tab(2).Control(2)=   "Label8"
      Tab(2).Control(3)=   "feProvCartera"
      Tab(2).ControlCount=   4
      Begin Sicmact.FlexEdit feIntDevengado 
         Height          =   3495
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   6165
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Crédito-Cta. Debe-Cta. Haber-Fecha-Interes Dev."
         EncabezadosAnchos=   "300-2500-2500-2500-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit feIntSuspenso 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   6165
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Crédito-Cta. Debe-Cta. Haber-Fecha-Interes Susp."
         EncabezadosAnchos=   "300-2500-2500-2500-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit feProvCartera 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   22
         Top             =   480
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   6165
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Crédito-Cta. Debe-Cta. Haber-Fecha-Prov. Cartera"
         EncabezadosAnchos=   "300-2500-2500-2500-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -66015
         TabIndex        =   21
         Top             =   4050
         Width           =   570
      End
      Begin VB.Label lblMontoProvCart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -64905
         TabIndex        =   20
         Top             =   4020
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   345
         Left            =   -66120
         Top             =   4000
         Width           =   2955
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -66015
         TabIndex        =   19
         Top             =   4050
         Width           =   570
      End
      Begin VB.Label lblMontoIntSus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -64905
         TabIndex        =   18
         Top             =   4020
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   345
         Left            =   -66120
         Top             =   4000
         Width           =   2955
      End
      Begin VB.Label lblMontoText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   9000
         TabIndex        =   16
         Top             =   4050
         Width           =   570
      End
      Begin VB.Label lblMontoIntDev 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   10110
         TabIndex        =   15
         Top             =   4020
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   345
         Left            =   8900
         Top             =   4000
         Width           =   2955
      End
   End
   Begin VB.Frame frafiltro 
      Caption         =   "Filtro"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar"
         Height          =   375
         Left            =   10930
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame fraPeriodoFin 
         Caption         =   "Periodo de Fin"
         Height          =   735
         Left            =   7400
         TabIndex        =   7
         Top             =   120
         Width           =   3495
         Begin VB.ComboBox cboMesFin 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   270
            Width           =   1335
         End
         Begin VB.TextBox txtAnioFin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   11
            Top             =   280
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Mes"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Año"
            Height          =   255
            Left            =   2040
            TabIndex        =   10
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame FraPeriodoIni 
         Caption         =   "Periodo de Inicio"
         Height          =   735
         Left            =   3840
         TabIndex        =   2
         Top             =   120
         Width           =   3495
         Begin VB.TextBox txtAnioIni 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   6
            Top             =   280
            Width           =   975
         End
         Begin VB.ComboBox cboMesIni 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Año"
            Height          =   255
            Left            =   2040
            TabIndex        =   5
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Mes"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   375
         End
      End
      Begin Sicmact.ActXCodCta actCodigoCuenta 
         Height          =   465
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   820
         Texto           =   "Cuenta:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
   End
End
Attribute VB_Name = "frmRepProvisionIndiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdMostrar_Click()
Dim rsIntDev As New ADODB.Recordset
Dim rsIntSus As New ADODB.Recordset
Dim rsRevInt As New ADODB.Recordset
Dim rsProvCart As New ADODB.Recordset
Dim nTotalIntDev As Currency
Dim nTotalIntSus As Currency
Dim nTotalProvCar As Currency

Dim oAjuste As New DAjusteCont
    If Not ValidaDatos Then Exit Sub
    
    Set rsIntDev = oAjuste.InteresDevengadoDetXCta(actCodigoCuenta.NroCuenta, DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Format(Right(cboMesIni.Text, 2), "00") & "/" & txtAnioIni.Text))), DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Format(Right(cboMesFin.Text, 2), "00") & "/" & txtAnioFin.Text))))
    LimpiaFlex feIntDevengado
    nTotalIntDev = 0
    Do While Not rsIntDev.EOF
        feIntDevengado.AdicionaFila
        feIntDevengado.TextMatrix(feIntDevengado.row, 1) = rsIntDev!cCtaCod
        feIntDevengado.TextMatrix(feIntDevengado.row, 2) = rsIntDev!ctaContDebe
        feIntDevengado.TextMatrix(feIntDevengado.row, 3) = rsIntDev!ctaContHaber
        feIntDevengado.TextMatrix(feIntDevengado.row, 4) = rsIntDev!dFecha
        feIntDevengado.TextMatrix(feIntDevengado.row, 5) = Format(rsIntDev!nMonto, "#,#0.00")
        nTotalIntDev = nTotalIntDev + rsIntDev!nMonto
        rsIntDev.MoveNext
    Loop
    lblMontoIntDev.Caption = Format(nTotalIntDev, "#,#0.00")
    
    Set rsIntSus = oAjuste.InteresSuspensoDetXCta(actCodigoCuenta.NroCuenta, DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Format(Right(cboMesIni.Text, 2), "00") & "/" & txtAnioIni.Text))), DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Format(Right(cboMesFin.Text, 2), "00") & "/" & txtAnioFin.Text))))
    LimpiaFlex feIntSuspenso
    nTotalIntSus = 0
    Do While Not rsIntSus.EOF
        feIntSuspenso.AdicionaFila
        feIntSuspenso.TextMatrix(feIntSuspenso.row, 1) = rsIntSus!cCtaCod
        feIntSuspenso.TextMatrix(feIntSuspenso.row, 2) = rsIntSus!ctaContDebe
        feIntSuspenso.TextMatrix(feIntSuspenso.row, 3) = rsIntSus!ctaContHaber
        feIntSuspenso.TextMatrix(feIntSuspenso.row, 4) = rsIntSus!dFecha
        feIntSuspenso.TextMatrix(feIntSuspenso.row, 5) = Format(rsIntSus!nMonto, "#,#0.00")
        nTotalIntSus = nTotalIntSus + rsIntSus!nMonto
        rsIntSus.MoveNext
    Loop
    lblMontoIntSus.Caption = Format(nTotalIntSus, "#,#0.00")
    
    Set rsProvCart = oAjuste.ProvisionCarteraDetXCta(actCodigoCuenta.NroCuenta, DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Format(Right(cboMesIni.Text, 2), "00") & "/" & txtAnioIni.Text))), DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Format(Right(cboMesFin.Text, 2), "00") & "/" & txtAnioFin.Text))))
    LimpiaFlex feProvCartera
    nTotalProvCar = 0
    Do While Not rsProvCart.EOF
        feProvCartera.AdicionaFila
        feProvCartera.TextMatrix(feProvCartera.row, 1) = rsProvCart!cCtaCod
        feProvCartera.TextMatrix(feProvCartera.row, 2) = rsProvCart!ctaContDebe
        feProvCartera.TextMatrix(feProvCartera.row, 3) = rsProvCart!ctaContHaber
        feProvCartera.TextMatrix(feProvCartera.row, 4) = rsProvCart!dFecha
        feProvCartera.TextMatrix(feProvCartera.row, 5) = Format(rsProvCart!nMonto, "#,#0.00")
        nTotalProvCar = nTotalProvCar + rsProvCart!nMonto
        rsProvCart.MoveNext
    Loop
    lblMontoProvCart.Caption = Format(nTotalProvCar, "#,#0.00")
    
    If (rsIntDev.RecordCount + rsIntSus.RecordCount + rsProvCart.RecordCount) = 0 Then
        MsgBox "No se encontraron datos para este crédito", vbInformation, "Aviso"
    End If
    
End Sub
Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    If Len(actCodigoCuenta.Age) <> 2 Or Len(actCodigoCuenta.Prod) <> 3 Or Len(actCodigoCuenta.Cuenta) <> 10 Then
        MsgBox "Por favor, ingresar un numero de cuenta valido.", vbInformation, "Aviso"
        actCodigoCuenta.SetFocus
        Exit Function
    End If
    If cboMesIni.ListIndex = -1 Then
        MsgBox "Por favor, seleccione un mes valido.", vbInformation, "Aviso"
        cboMesIni.SetFocus
        Exit Function
    End If
    If Not Len(txtAnioIni.Text) = 4 Then
        MsgBox "Por favor, ingresar un anio valido.", vbInformation, "Aviso"
        txtAnioIni.SetFocus
        Exit Function
    End If
    If cboMesFin.ListIndex = -1 Then
        MsgBox "Por favor, seleccione un mes valido.", vbInformation, "Aviso"
        cboMesFin.SetFocus
        Exit Function
    End If
    If Not Len(txtAnioFin.Text) = 4 Then
        MsgBox "Por favor, ingresar un anio valido.", vbInformation, "Aviso"
        txtAnioFin.SetFocus
        Exit Function
    End If
    If DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Format(Right(cboMesIni.Text, 2), "00") & "/" & txtAnioIni.Text))) > DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Format(Right(cboMesFin.Text, 2), "00") & "/" & txtAnioFin.Text))) Then
        MsgBox "El periodo de inicio no puede ser mayor al periodo final. Verifique", vbInformation, "Aviso"
        txtAnioIni.SetFocus
        Exit Function
    End If
     ValidaDatos = True
End Function
Private Sub Form_Load()
    CargaControles
End Sub
Private Sub CargaControles()
    'CargaMeses
    cboMesIni.Clear
    cboMesIni.AddItem "ENERO" & Space(150) & "1"
    cboMesIni.AddItem "FEBRERO" & Space(150) & "2"
    cboMesIni.AddItem "MARZO" & Space(150) & "3"
    cboMesIni.AddItem "ABRIL" & Space(150) & "4"
    cboMesIni.AddItem "MAYO" & Space(150) & "5"
    cboMesIni.AddItem "JUNIO" & Space(150) & "6"
    cboMesIni.AddItem "JULIO" & Space(150) & "7"
    cboMesIni.AddItem "AGOSTO" & Space(150) & "8"
    cboMesIni.AddItem "SEPTIEMBRE" & Space(150) & "9"
    cboMesIni.AddItem "OCTUBRE" & Space(150) & "10"
    cboMesIni.AddItem "NOVIEMBRE" & Space(150) & "11"
    cboMesIni.AddItem "DICIEMBRE" & Space(150) & "12"
    
    cboMesFin.Clear
    cboMesFin.AddItem "ENERO" & Space(150) & "1"
    cboMesFin.AddItem "FEBRERO" & Space(150) & "2"
    cboMesFin.AddItem "MARZO" & Space(150) & "3"
    cboMesFin.AddItem "ABRIL" & Space(150) & "4"
    cboMesFin.AddItem "MAYO" & Space(150) & "5"
    cboMesFin.AddItem "JUNIO" & Space(150) & "6"
    cboMesFin.AddItem "JULIO" & Space(150) & "7"
    cboMesFin.AddItem "AGOSTO" & Space(150) & "8"
    cboMesFin.AddItem "SEPTIEMBRE" & Space(150) & "9"
    cboMesFin.AddItem "OCTUBRE" & Space(150) & "10"
    cboMesFin.AddItem "NOVIEMBRE" & Space(150) & "11"
    cboMesFin.AddItem "DICIEMBRE" & Space(150) & "12"
    
    txtAnioIni.Text = Year(gdFecSis)
    txtAnioFin.Text = Year(gdFecSis)
    
End Sub
Private Sub txtAnioFin_GotFocus()
    fEnfoque txtAnioFin
End Sub
Private Sub txtAnioFin_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
       Me.cmdMostrar.SetFocus
    End If
End Sub
Private Sub txtAnioIni_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
       Me.cboMesFin.SetFocus
    End If
End Sub
Private Sub txtAnioIni_GotFocus()
    fEnfoque txtAnioIni
End Sub
Private Sub actCodigoCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMesIni.SetFocus
    End If
End Sub
Private Sub cboMesFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboMesFin.ListIndex <> -1 Then
            txtAnioFin.SetFocus
        End If
    End If
End Sub
Private Sub cboMesIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboMesIni.ListIndex <> -1 Then
            txtAnioIni.SetFocus
        End If
    End If
End Sub
