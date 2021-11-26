VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepResponsability 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Mensual para RESPONSABILITY"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmRepResponsability.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "G&enerar"
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   5700
      TabIndex        =   30
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "&Consulta"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame fraRepMensualRespon 
      Caption         =   "Reporte Mensual para RESPONSABILITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   8055
      Begin VB.TextBox txtRMRExplica 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00800000&
         Height          =   885
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2400
         Width           =   2895
      End
      Begin VB.ComboBox CboRMRCovenant 
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
         Height          =   315
         ItemData        =   "frmRepResponsability.frx":030A
         Left            =   5040
         List            =   "frmRepResponsability.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtRMRCalif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5040
         TabIndex        =   24
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtRMRAgeCalif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5040
         TabIndex        =   20
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboRMRInfEmitido 
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
         Height          =   315
         ItemData        =   "frmRepResponsability.frx":0392
         Left            =   5040
         List            =   "frmRepResponsability.frx":039C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   195
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtRMRFecRep 
         Height          =   300
         Left            =   5040
         TabIndex        =   22
         Top             =   960
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblRep_Campo6 
         Caption         =   "Si la respuesta es afirmativa, favor explicar"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label lblRep_Campo5 
         Caption         =   "¿Está su institución rompiendo algún covenant y/o en default con algún acreedor (si/no)?"
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   1725
         Width           =   4695
      End
      Begin VB.Label lblRep_Campo4 
         Caption         =   "Calificación:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblRep_Campo3 
         Caption         =   "Fecha del Reporte:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblRep_Campo2 
         Caption         =   "Agencia de Calificación:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblRep_Campo1 
         Caption         =   "Informe emitido/recibido en el periodo (Si/No):"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fraDATAB 
      Caption         =   "DATA B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   8055
      Begin VB.TextBox txtDBRatioSufCap 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5040
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtDBCredSubDivML 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5040
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtDBCredSubML 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5040
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtDBCredDivML 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtDBCredMLIFI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblDataB_Campo5 
         Caption         =   "Ratio de suficiencia del capital:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblDataB_Campo4 
         Caption         =   "Créditos subordinados en divisas (expresado en moneda local):"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label lblDataB_Campo3 
         Caption         =   "Créditos subordinados en moneda local:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lblDataB_Campo2 
         Caption         =   "Créditos en divisas (expresado en moneda local):"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblDataB_Campo1 
         Caption         =   "Créditos en Moneda Local de IFIs, Bancos y otros:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3615
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
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   4
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox txtMes 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Top             =   200
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Año:"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblPeriodo_Campo1 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRepResponsability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmRepResponsability
'Descripcion:Formulario para Generar el Reporte de Responsability
'Creacion: PASI TI-ERS087-2014
'*****************************
Option Explicit
Dim ldFecPeriodo As Date
Dim bEstadoForm As Boolean
Private Type TCtaCont
    CuentaContable As String
    Saldo As Currency
    bSaldoA As Boolean
    bSaldoD As Boolean
End Type

Private Sub cmdConsulta_Click()
    frmRepResponsabilityConsulta.Show 1
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo ErrorGuardar
    Dim oDResponblt As DResponsability
    Dim nIdRep As Integer
    Dim nIdRepDet As Integer
    Dim nIdRepDetAlt As Integer
    Dim oDMov As DMov
    Set oDMov = New DMov
    Dim lsMovNro As String
    Dim btrans As Boolean
    
    If Not ValidaDatos Then Exit Sub
    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set oDResponblt = New DResponsability
    Screen.MousePointer = 11
    lsMovNro = oDMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    nIdRep = 0
    nIdRepDet = 0
    oDResponblt.dBeginTrans
    btrans = True
    nIdRep = oDResponblt.RegistrarConfigRepResponsability(Format(DatePart("M", ldFecPeriodo), "00"), txtAnio.Text, lsMovNro, 1)
    'Detalle del Reporte
    'nIdRepDet = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, "DATA A", "", 0)
    'nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblDataA_Campo1.Caption), CStr(Trim(txtDAnumOficinas.Text)), nIdRepDet)
    nIdRepDet = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, "DATA B", "", 0)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblDataB_Campo1.Caption), CStr(Trim(txtDBCredMLIFI.Text)), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblDataB_Campo2.Caption), CStr(Trim(txtDBCredDivML.Text)), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblDataB_Campo3.Caption), CStr(Trim(txtDBCredSubML.Text)), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblDataB_Campo4.Caption), CStr(Trim(txtDBCredSubDivML.Text)), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblDataB_Campo5.Caption), CStr(Trim(txtDBRatioSufCap.Text)), nIdRepDet)
    nIdRepDet = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, "REPORTE MENSUAL PARA RESPONSABILITY", "", 0)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblRep_Campo1.Caption), CStr(Trim(Left(cboRMRInfEmitido.Text, 2))), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblRep_Campo2.Caption), CStr(Trim(txtRMRAgeCalif.Text)), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblRep_Campo3.Caption), CStr(Trim(txtRMRFecRep.Text)), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblRep_Campo4.Caption), CStr(Trim(txtRMRCalif.Text)), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(lblRep_Campo5.Caption), CStr(Trim(Left(CboRMRCovenant.Text, 2))), nIdRepDet)
    nIdRepDetAlt = oDResponblt.RegistrarConfigRepResponsabilityDet(nIdRep, Trim(Replace(Replace(lblRep_Campo6.Caption, Chr(10), ""), Chr(13), "")), CStr(Trim(txtRMRExplica.Text)), nIdRepDet)
    oDResponblt.dCommitTrans
    
    Screen.MousePointer = 0
    MsgBox "Configuración grabada satisfactoriamente", vbInformation, "Aviso"
    btrans = False
    LimpiarDatos
    Exit Sub
ErrorGuardar:
    Screen.MousePointer = 0
    If btrans Then
        oDResponblt.dRollbackTrans
    End If
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
End Sub
Private Sub LimpiarDatos()
    txtDBCredMLIFI.Text = ""
    txtDBCredDivML.Text = ""
    txtDBCredSubML.Text = ""
    txtDBCredSubDivML.Text = ""
    txtDBRatioSufCap.Text = ""
    cboRMRInfEmitido.ListIndex = -1
    txtRMRAgeCalif.Text = ""
    txtRMRFecRep.Text = "__/__/__"
    txtRMRCalif.Text = ""
    CboRMRCovenant.ListIndex = -1
    txtRMRExplica.Text = ""
End Sub
Private Function ValidaDatos() As Boolean
Dim oDResponblt As DResponsability
Set oDResponblt = New DResponsability
    ValidaDatos = False
     If oDResponblt.ExisteConfigRepResponsability(Format(DatePart("M", ldFecPeriodo), "00"), txtAnio.Text) Then
        MsgBox "Ya existe la configuración para el periodo seleccionado.", vbInformation, "Aviso"
        Exit Function
    End If
'    If Len(Trim(txtDAnumOficinas.Text)) = 0 Then
'        MsgBox "No se ha ingresado DATA A: 'Número de Oficinas'", vbInformation, "Aviso"
'        txtDAnumOficinas.SetFocus
'        Exit Function
'    End If
    If Len(Trim(txtDBCredMLIFI.Text)) = 0 Then
         MsgBox "No se ha ingresado el valor de DATA B: 'Créditos en Moneda Local de IFIs, Bancos y otros':", vbInformation, "Aviso"
        txtDBCredMLIFI.SetFocus
        Exit Function
    End If
    If Len(Trim(txtDBCredDivML.Text)) = 0 Then
         MsgBox "No se ha ingresado el valor de 'DATA B: 'Créditos en divisas'", vbInformation, "Aviso"
        txtDBCredDivML.SetFocus
        Exit Function
    End If
     If Len(Trim(txtDBCredSubML.Text)) = 0 Then
         MsgBox "No se ha ingresado el valor de 'DATA B: 'Créditos subordinados en moneda local'", vbInformation, "Aviso"
        txtDBCredSubML.SetFocus
        Exit Function
    End If
    If Len(Trim(txtDBCredSubDivML.Text)) = 0 Then
      MsgBox "No se ha ingresado el valor de 'DATA B: 'Créditos subordinados en divisas'", vbInformation, "Aviso"
     txtDBCredSubDivML.SetFocus
     Exit Function
    End If
    If Len(Trim(txtDBRatioSufCap.Text)) = 0 Then
      MsgBox "No se ha ingresado el valor de 'DATA B: 'Ratio de suficiencia del capital'", vbInformation, "Aviso"
     txtDBRatioSufCap.SetFocus
     Exit Function
    End If
    If cboRMRInfEmitido.ListIndex = -1 Then
        MsgBox "No se ha seleccionado 'Informe emitido /recibido en el periodo'", vbInformation, "Aviso"
        cboRMRInfEmitido.SetFocus
     Exit Function
    End If
    If Len(Trim(txtRMRAgeCalif.Text)) = 0 Then
      MsgBox "No se ha ingresado el valor de 'Agencia de Calificación'", vbInformation, "Aviso"
     txtRMRAgeCalif.SetFocus
     Exit Function
    End If
    If (txtRMRFecRep.Text = "__/__/__") Then
        MsgBox "No se ha ingresado el valor de 'Fecha del Reporte'", vbInformation, "Aviso"
        txtRMRFecRep.SetFocus
     Exit Function
    End If
    If Len(Trim(txtRMRCalif.Text)) = 0 Then
        MsgBox "No se ha ingresado el valor de 'Calificación'", vbInformation, "Aviso"
        txtRMRCalif.SetFocus
        Exit Function
    End If
    If CboRMRCovenant.ListIndex = -1 Then
        MsgBox "No se ha seleccionado 'Incumplimiento de algún covenant...'", vbInformation, "Aviso"
        CboRMRCovenant.SetFocus
     Exit Function
    End If
    If Len(Trim(Replace(Replace(txtRMRExplica.Text, Chr(10), ""), Chr(13), ""))) = 0 And txtRMRExplica.Enabled = True Then
        MsgBox "No se ha Ingresado la explicacion a la respuesta afirmativa...", vbInformation, "Aviso"
        txtRMRExplica = Trim(Replace(Replace(txtRMRExplica.Text, Chr(10), ""), Chr(13), ""))
        txtRMRExplica.SetFocus
        Exit Function
    End If
    ValidaDatos = True
End Function
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    bEstadoForm = False
    EstablecePeriodo
    bEstadoForm = True
End Sub
Private Sub EstablecePeriodo()
    ldFecPeriodo = DateAdd("D", -1, "01/" & Format(DatePart("M", gdFecSis), "00") & "/" & Format(DatePart("YYYY", gdFecSis), "0000"))
    'ldFecPeriodo = DateAdd("D", -1, DateAdd("M", 1, "01/" & Format(DatePart("M", gdFecSis), "00") & "/" & Format(DatePart("YYYY", gdFecSis), "0000")))
    txtMes.Text = UCase(Devuelvemes(DatePart("M", ldFecPeriodo)))
    txtAnio.Text = Format(DatePart("YYYY", ldFecPeriodo), "0000")
    cboRMRInfEmitido.ListIndex = 1
    CboRMRCovenant.ListIndex = 1
End Sub
Private Sub txtDAnumOficinas_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumeros(KeyAscii)
     If KeyAscii = 13 Then
        Me.txtDBCredMLIFI.SetFocus
    End If
End Sub
Private Sub txtDBCredMLIFI_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
     If KeyAscii = 13 Then
        Me.txtDBCredDivML.SetFocus
    End If
End Sub
Private Sub txtDBCredMLIFI_LostFocus()
    If Trim(txtDBCredMLIFI.Text) = "" Then
        txtDBCredMLIFI.Text = "0.00"
    End If
    txtDBCredMLIFI.Text = Format(txtDBCredMLIFI.Text, "#0.00")
End Sub
Private Sub txtDBCredDivML_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
     If KeyAscii = 13 Then
        Me.txtDBCredSubML.SetFocus
    End If
End Sub
Private Sub txtDBCredDivML_LostFocus()
    If Trim(txtDBCredDivML.Text) = "" Then
        txtDBCredDivML.Text = "0.00"
    End If
    txtDBCredDivML.Text = Format(txtDBCredDivML.Text, "#0.00")
End Sub
Private Sub txtDBCredSubML_KeyPress(KeyAscii As Integer)
     KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
     If KeyAscii = 13 Then
        Me.txtDBCredSubDivML.SetFocus
    End If
End Sub
Private Sub txtDBCredSubML_LostFocus()
    If Trim(txtDBCredSubML.Text) = "" Then
        txtDBCredSubML.Text = "0.00"
    End If
    txtDBCredSubML.Text = Format(txtDBCredSubML.Text, "#0.00")
End Sub
Private Sub txtDBCredSubDivML_KeyPress(KeyAscii As Integer)
     KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
     If KeyAscii = 13 Then
        Me.txtDBRatioSufCap.SetFocus
    End If
End Sub
Private Sub txtDBCredSubDivML_LostFocus()
    If Trim(txtDBCredSubDivML.Text) = "" Then
        txtDBCredSubDivML.Text = "0.00"
    End If
    txtDBCredSubDivML.Text = Format(txtDBCredSubDivML.Text, "#0.00")
End Sub
Private Sub txtDBRatioSufCap_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
        Me.cboRMRInfEmitido.SetFocus
    End If
End Sub
Private Sub txtDBRatioSufCap_LostFocus()
    If Trim(txtDBRatioSufCap.Text) = "" Then
        txtDBRatioSufCap.Text = "0.00"
    End If
    txtDBRatioSufCap.Text = Format(txtDBRatioSufCap.Text, "#0.00")
End Sub
Private Sub cboRMRInfEmitido_Click()
    If cboRMRInfEmitido.ListIndex <> -1 And bEstadoForm Then
        txtRMRAgeCalif.SetFocus
    End If
End Sub
Private Sub txtRMRAgeCalif_KeyPress(KeyAscii As Integer)
    'KeyAscii = TextBox_SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtRMRFecRep.SetFocus
    End If
End Sub
Private Sub txtRMRFecRep_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtRMRCalif.SetFocus
    End If
End Sub
Private Sub txtRMRCalif_KeyPress(KeyAscii As Integer)
    'KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 Then
        Me.CboRMRCovenant.SetFocus
    End If
End Sub
Private Sub CboRMRCovenant_Click()
    If CboRMRCovenant.ListIndex <> -1 And bEstadoForm Then
        If Trim(Right(CboRMRCovenant.Text, 2)) = "1" Then
            txtRMRExplica.Enabled = True
            txtRMRExplica.SetFocus
        Else
            txtRMRExplica.Enabled = False
            txtRMRExplica.Text = ""
            cmdGuardar.SetFocus
        End If
    End If
End Sub
Private Sub txtRMRExplica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGuardar.SetFocus
    End If
End Sub
Private Sub cmdGenerar_Click()
    Dim oDResponblt As DResponsability
    Dim celda As Excel.Range

    Set oDResponblt = New DResponsability
    If Not oDResponblt.ExisteConfigRepResponsability(Format(DatePart("M", ldFecPeriodo), "00"), txtAnio.Text) Then
        MsgBox "Aún no se ha generado la información para este periodo. Verifique.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Dim sPathFormatoResponsability As String
    
    Dim fs As New Scripting.FileSystemObject
    Dim obj_excel As Object, Libro As Object, Hoja As Object
    
    On Error GoTo error_sub
    
    sPathFormatoResponsability = App.path & "\Spooler\RepResponsability_" + Format(ldFecPeriodo, "yyyymmdd") + ".xlsx"
    If fs.FileExists(sPathFormatoResponsability) Then
        If ArchivoEstaAbierto(sPathFormatoResponsability) Then
            If MsgBox("Debe Cerrar el Archivo: " + fs.GetFileName(sPathFormatoResponsability) + " para continuar", vbRetryCancel) = vbCancel Then
                Me.MousePointer = vbDefault
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
        End If
        fs.DeleteFile sPathFormatoResponsability, True
    End If
    sPathFormatoResponsability = App.path & "\FormatoCarta\FormatoResponsabiility.xlsx"
    If Len(Dir(sPathFormatoResponsability)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathFormatoResponsability, vbCritical
           Me.MousePointer = vbDefault
           Exit Sub
    End If
    
    Set obj_excel = CreateObject("Excel.Application")
    obj_excel.DisplayAlerts = False
    Set Libro = obj_excel.Workbooks.Open(sPathFormatoResponsability)
    Set Hoja = Libro.ActiveSheet
    
    CargaData obj_excel
    
    sPathFormatoResponsability = App.path & "\Spooler\RepResponsability_" + Format(ldFecPeriodo, "yyyymmdd") + ".xlsx"
    If fs.FileExists(sPathFormatoResponsability) Then
        If ArchivoEstaAbierto(sPathFormatoResponsability) Then
            MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathFormatoResponsability)
        End If
        fs.DeleteFile sPathFormatoResponsability, True
    End If
    Hoja.SaveAs sPathFormatoResponsability
    Libro.Close
    obj_excel.Quit
    Set Hoja = Nothing
    Set Libro = Nothing
    Set obj_excel = Nothing
    Me.MousePointer = vbDefault
    
    Dim m_excel As New Excel.Application
    m_excel.Workbooks.Open (sPathFormatoResponsability)
    m_excel.Visible = True
    Exit Sub
error_sub:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        Me.MousePointer = vbDefault
End Sub
Private Sub CargaData(ByVal pobj_Excel As Excel.Application)
   Dim oDResponblt As DResponsability
   Set oDResponblt = New DResponsability
   
   Dim nfil As Integer
   Dim celdaCampo As Excel.Range
   Dim celdaValor As Excel.Range
   
   'Nombre de la Institucion
   nfil = 2
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = "CMAC Maynas"
   'Cifras al
   nfil = 3
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = DatePart("D", ldFecPeriodo) & "-" & Left(dameNombreMes(DatePart("M", ldFecPeriodo)), 3) & Right(DatePart("YYYY", ldFecPeriodo), 2)
   'Moneda Local
   nfil = 4
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = "PEN"
   'Número de prestatarios
   nfil = 5
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = oDResponblt.ObtieneValorNumeroPrestatario(ldFecPeriodo)
   'Número de ahorrantes (excl. ahorros forzosos)
    nfil = 6
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = oDResponblt.ObtieneValorNumeroAhorrantes(ldFecPeriodo)
   'Caja y Bancos
   nfil = 7
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 2), 0)
   'Inversiones Financieras
   nfil = 8
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 10), 0)
   'Cartera bruta (vigentes y de largo plazo)
   nfil = 10
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   Dim nValVig, nValRefinan, nValVenc, nValCobJudi As Currency
   nValVig = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 15), 0)
   nValRefinan = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 17), 0)
   nValVenc = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 18), 0)
   nValCobJudi = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 19), 0)
   celdaValor.value = CCur(Round(nValVig + nValRefinan + nValVenc + nValCobJudi, 0))
   'Reservas créditos vencidos
   nfil = 11
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(Round(Abs(ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 20), 0)), 0))
   'Otras cuentas corrientes
   nfil = 12
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   Dim nValCtaxCobrar, nValBienReal, nValorActivIntangible, nValorImpCte, nValorImpDif, nValOtroActiv As Currency
   nValCtaxCobrar = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 24), 0)
   nValBienReal = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 25), 0)
   nValorActivIntangible = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 31), 0)
   nValorImpCte = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 32), 0)
   nValorImpDif = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 33), 0)
   nValOtroActiv = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 35), 0)
   celdaValor.value = nValCtaxCobrar + nValBienReal + nValorActivIntangible + nValorImpCte + nValorImpDif + nValOtroActiv
   'Activos fijos netos(y otros activos no corrientes)
   nfil = 13
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   Dim nValorParticip, nValorInmueble As Currency
   nValorParticip = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 28), 0)
   nValorInmueble = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 29), 0)
   celdaValor.value = nValorParticip + nValorInmueble
   'Ahorros y Depositos a Plazo (excl. Ahorros forzosos)
   nfil = 16
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   Dim nValorOblPub, nValorDepESF As Currency
   nValorOblPub = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 37), 0)
   nValorDepESF = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 43), 0)
   celdaValor.value = nValorOblPub + nValorDepESF
   'Creditos en moneda local /de IFIs, bancos y otros
   nfil = 17
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Créditos en Moneda Local de IFIs, Bancos y otros:"))
   'Creditos en divisas (expresado en moneda local)
   nfil = 18
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Créditos en divisas (expresado en moneda local):"))
   'Otros pasivos
   nfil = 19
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = Abs(ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 62), 0))
   'Créditos subordinados en moneda local
   nfil = 20
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Créditos subordinados en moneda local:"))
   'Créditos subordinados en divisas (expresado en moneda local)
   nfil = 21
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur(oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Créditos subordinados en divisas (expresado en moneda local):"))
   'Patrimonio
   nfil = 22
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = Abs(ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 64), 0))
   'Monto total mantenido en bancos u otras instituciones financieras para financiacion (back to back)
   nfil = 24
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = 0
   'Activos Dado en garantia hacia los refinanciados
   nfil = 25
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = Abs(oDResponblt.ObtieneValorActivGarantia(ldFecPeriodo))
   'Cartera en administracion (Ver Hoja 'Definitions' Linea 26)
   nfil = 26
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = 0
   'Activos denominados o indexados en USD
   nfil = 28
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 1), 2)
   'Pasivos denominados o indexados en USD (no cubiertos)
   nfil = 29
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760108", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 63), 2)
   'Ratio de suficiencia de capital
   nfil = 30
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = CCur((oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Ratio de suficiencia del capital:"))) / 100
   'Cantidad de meses
   nfil = 31
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = DatePart("M", ldFecPeriodo)
   'Ingresos por intereses(cartera de créditos)
   nfil = 32
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 8), 0)
   'Ingresos por comisiones (cartera de crédito)
   nfil = 33
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
   celdaValor.value = 0
   'Gastos operativos
   nfil = 34
   Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
            'Gastos de personal
            Dim nValorGastoPers As Currency
            nValorGastoPers = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 51), 0)
            'Gastos Administrativos
            Dim nValorGastosAdmin As Currency
            nValorGastosAdmin = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 52), 0)
            'Otros Gastos Operativos
            Dim nValorOtrosGastosOpe, nValorPrimaFondo, nValorGastosDiv, nValorImpContri, nValorDeprecAmort, nValorValuaActivProv As Currency
            nValorPrimaFondo = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 37), 0)
            nValorGastosDiv = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 38), 0)
            nValorImpContri = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 53), 0)
            nValorDeprecAmort = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 54), 0)
            nValorValuaActivProv = ObtenerResultadoFormula(ldFecPeriodo, oDResponblt.ObtieneNIIFRepForResponsability("760109", Format(DatePart("YYYY", ldFecPeriodo), "0000"), DatePart("M", ldFecPeriodo), 56), 0)
            nValorOtrosGastosOpe = nValorPrimaFondo + nValorGastosDiv + nValorImpContri + nValorDeprecAmort + (nValorValuaActivProv * 2)
    celdaValor.value = nValorGastoPers + nValorGastosAdmin + nValorOtrosGastosOpe
    'Resultado neto del ejercicio despues de impuestos
    nfil = 35
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_ResNetoEjercicioByVal(ldFecPeriodo))
    
    ' PAR 1-30 Dias
    nfil = 36
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_PAR1and30Dias(ldFecPeriodo))
    'PAR > 30 dias
    nfil = 37
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_PARMayor30Dias(ldFecPeriodo))
    'Creditos reestructurados/reprogramados/refinanciados(no incluidos en PAR > 30)
    nfil = 38
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_CredReesctruc(ldFecPeriodo))
    'Total de Castigos
    nfil = 39
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = Abs(oDResponblt.ObtieneValorCalCartera_TotalCastigos(ldFecPeriodo))
    'Informe emtido/recibido en el periodo (Si/No)
    nfil = 40
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Informe emitido/recibido en el periodo (Si/No):"))
    'Agencia de Calificacion
    nfil = 41
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Agencia de Calificación:"))
    'Fecha del Reporte
    nfil = 42
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Fecha del Reporte:"))
    'Calificación:
    nfil = 43
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Calificación:"))
    'Esta su institucion incumpliendo algun covenant:
    nfil = 44
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "¿Está su institución rompiendo algún covenant y/o en default con algún acreedor (si/no)?"))
    'Si la respuesta es afirmativa, favor de explicar
    nfil = 45
    Set celdaValor = pobj_Excel.Range("REP!C" & nfil)
    celdaValor.value = (oDResponblt.ObtieneValorxRepResponsability(Format(DatePart("YYYY", ldFecPeriodo), "0000"), Format(DatePart("M", ldFecPeriodo), "00"), "Si la respuesta es afirmativa, favor explicar"))
End Sub
Private Function ObtenerResultadoFormula(ByVal pdFecha As Date, ByVal psFormula As String, ByVal pnMoneda As Integer, Optional psAgencia As String = "") As Currency
    Dim oBal As New DbalanceCont
    Dim oNBal As New NBalanceCont
    Dim oFormula As New NInterpreteFormula
    Dim lsFormula As String, lsTmp As String, lsTmp1 As String, lsCadFormula As String
    Dim MatDatos() As TCtaCont
    Dim i As Long, j As Long, nCtaCont As Long
    Dim sTempAD As String
    Dim nPosicion As Integer
    Dim signo As String
    Dim LsSigno As String
    lsFormula = Trim(psFormula)
    ReDim MatDatos(0)
    nCtaCont = 0
    lsTmp = ""
    lsFormula = Replace(lsFormula, "M", pnMoneda)
    sTempAD = ""
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                nCtaCont = nCtaCont + 1
                ReDim Preserve MatDatos(nCtaCont)
                
                MatDatos(nCtaCont).CuentaContable = lsTmp
                
                If MatDatos(nCtaCont).CuentaContable = "100" Or MatDatos(nCtaCont).CuentaContable = "1000" Then
                    MatDatos(nCtaCont).Saldo = MatDatos(nCtaCont).CuentaContable
                Else
                    If Trim(psAgencia) = "" Then
                        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
                    Else
                        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensualxAgencia(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True, psAgencia)
                    End If
                End If
                
                If nCtaCont > 1 Then
                    If Mid(Trim(lsFormula), i, 1) = ")" Then
                        nPosicion = 0
                    Else
                        nPosicion = i
                    End If
                End If
                    If sTempAD = "" Then
                        If nCtaCont = 1 Then
                            If ((i - Len(Trim(lsTmp))) - 3) > 1 Then
                                sTempAD = Mid(Trim(lsFormula), (i - Len(Trim(lsTmp))) - 3, 2)
                            Else
                                sTempAD = ""
                            End If
                        Else
                            sTempAD = Mid(Trim(lsFormula), (i - Len(MatDatos(nCtaCont).CuentaContable)) - 3, 2)
                        End If
                    End If
                
                If sTempAD = "SA" Or sTempAD = "SD" Then
                    MatDatos(nCtaCont).CuentaContable = DepuraSaldoAD(MatDatos(nCtaCont).CuentaContable)
                    If sTempAD = "SA" Then
                        MatDatos(nCtaCont).bSaldoA = True
                        MatDatos(nCtaCont).bSaldoD = False
                    Else
                        MatDatos(nCtaCont).bSaldoA = False
                        MatDatos(nCtaCont).bSaldoD = True
                    End If
                    Else
                        MatDatos(nCtaCont).bSaldoA = False
                        MatDatos(nCtaCont).bSaldoD = False
                End If
            End If
            If nPosicion = 0 Then
               sTempAD = ""
            End If
            lsTmp = ""
        End If
    Next i
    If Len(lsTmp) > 0 Then
        nCtaCont = nCtaCont + 1
        ReDim Preserve MatDatos(nCtaCont)
        MatDatos(nCtaCont).CuentaContable = lsTmp
        If MatDatos(nCtaCont).CuentaContable = "100" Or MatDatos(nCtaCont).CuentaContable = "1000" Then
            MatDatos(nCtaCont).Saldo = MatDatos(nCtaCont).CuentaContable
        Else
            If Trim(psAgencia) = "" Then
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
            Else
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensualxAgencia(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True, psAgencia)
            End If
        End If
    End If
    lsTmp = ""
    lsCadFormula = ""
    Dim nEncontrado As Integer
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                For j = 1 To nCtaCont
                    nEncontrado = 0
                    If MatDatos(j).CuentaContable = lsTmp Then
                            
                            If MatDatos(j).bSaldoA = True Or MatDatos(j).bSaldoD = True Then
                                MatDatos(j).Saldo = oNBal.CalculaSaldoBECuentaAD(MatDatos(j).CuentaContable, pnMoneda, MatDatos(j).bSaldoA, CStr(pnMoneda), Trim(psAgencia), Format(pdFecha, "YYYY"), Format(pdFecha, "MM"))
                                nEncontrado = 1
                            End If
                                If Left(Format(MatDatos(j).Saldo, "#0.00"), 1) = "-" And (Right(lsCadFormula, 1) = "-" Or Right(lsCadFormula, 1) = "+") Then
                                    
                                    If Right(Trim(lsCadFormula), 1) = "-" Or Right(Trim(lsCadFormula), 1) = "+" Then
                                        If Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo >= 0 Then
                                            LsSigno = "-"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo < 0 Then
                                            LsSigno = "+"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo >= 0 Then
                                            LsSigno = "+"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo < 0 Then
                                            LsSigno = "-"
                                        End If
                                    Else
                                        LsSigno = ""
                                    End If
                                    If LsSigno = "" Then
                                        lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                                    Else
                                        lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & LsSigno & Format(Abs(MatDatos(j).Saldo), "#0.00")
                                    End If
                                    nEncontrado = 1
                                Else
                                    lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                                    nEncontrado = 1
                                End If
                        Exit For
                    End If
                Next j
            End If
            lsTmp = ""
            If nEncontrado = 1 Or (Mid(Trim(lsFormula), i, 1) = "S" Or Mid(Trim(lsFormula), i, 1) = "A" Or Mid(Trim(lsFormula), i, 1) = "D") Then
            lsCadFormula = lsCadFormula & Mid(Trim(lsFormula), i, 1)
            Else
            lsCadFormula = lsCadFormula & "" & Mid(Trim(lsFormula), i, 1)
            End If
        End If
    Next
    If Len(lsTmp) > 0 Then
        For j = 1 To nCtaCont
           If MatDatos(j).CuentaContable = lsTmp Then
               If Left(Format(MatDatos(j).Saldo, "#0.00"), 1) = "-" And (Right(lsCadFormula, 1) = "-" Or Right(lsCadFormula, 1) = "+") Then
                    lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                    nEncontrado = 1
                Else
                    lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                    nEncontrado = 1
                End If
               Exit For
           End If
        Next j
    End If
    lsCadFormula = Replace(Replace(lsCadFormula, "SA", ""), "SD", "")
    ObtenerResultadoFormula = oFormula.ExprANum(lsCadFormula)
    Set oBal = Nothing
    Set oFormula = Nothing
End Function
Private Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
On Error GoTo HayErrores
Dim f As Integer
   f = FreeFile
   Open Ruta For Append As f
   Close f
   ArchivoEstaAbierto = False
   Exit Function
HayErrores:
   If Err.Number = 70 Then
      ArchivoEstaAbierto = True
   Else
      Err.Raise Err.Number
   End If
End Function
Private Function DepuraSaldoAD(ByVal sCta As String) As String
Dim i As Integer
Dim Cad As String
    Cad = ""
    For i = 1 To Len(sCta)
        If Mid(sCta, i, 1) >= "0" And Mid(sCta, i, 1) <= "9" Then
            Cad = Cad + Mid(sCta, i, 1)
        End If
    Next i
    DepuraSaldoAD = Cad
End Function
