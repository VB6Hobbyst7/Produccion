VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRHPrestamosAdm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmRHPrestamosAdm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSoloMarcados 
      Alignment       =   1  'Right Justify
      Caption         =   "Imp. Seleccionados"
      Height          =   195
      Left            =   3855
      TabIndex        =   30
      Top             =   6420
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7125
      TabIndex        =   27
      Top             =   6345
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   45
      TabIndex        =   26
      Top             =   6345
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8205
      TabIndex        =   25
      Top             =   6345
      Width           =   975
   End
   Begin VB.Frame fraCuentas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5100
      Left            =   45
      TabIndex        =   16
      Top             =   1140
      Width           =   9135
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   4635
         Width           =   855
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1800
         TabIndex        =   20
         Top             =   4647
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdAplicar 
         Height          =   375
         Left            =   4680
         Picture         =   "frmRHPrestamosAdm.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4635
         Width           =   375
      End
      Begin VB.TextBox txtMontoME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3360
         TabIndex        =   18
         Top             =   4647
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkCred 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   17
         Top             =   300
         Width           =   1245
      End
      Begin MSComctlLib.ListView lstCredAdm 
         Height          =   3915
         Left            =   120
         TabIndex        =   22
         Top             =   615
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6906
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Empleado"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Crédito"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "N°"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Fec.Venc."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Monto S/."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Monto $"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CodPers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "cEmpCod"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblSoles 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   24
         Top             =   4717
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblDolares 
         AutoSize        =   -1  'True
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3120
         TabIndex        =   23
         Top             =   4717
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Frame fraTipoBus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   4125
      TabIndex        =   11
      Top             =   45
      Width           =   3015
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Empleado"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   375
         Left            =   2400
         Picture         =   "frmRHPrestamosAdm.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtEmp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1305
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame fraPlanilla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Planilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   45
      TabIndex        =   7
      Top             =   45
      Width           =   4095
      Begin Sicmact.TxtBuscar txtPlanillas 
         Height          =   300
         Left            =   90
         TabIndex        =   28
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Appearance      =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.CommandButton cmdBuscaPla 
         Height          =   375
         Left            =   3645
         Picture         =   "frmRHPrestamosAdm.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   218
         Width           =   375
      End
      Begin MSComCtl2.DTPicker txtFecPla 
         Height          =   315
         Left            =   2085
         TabIndex        =   9
         Top             =   248
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   66387969
         CurrentDate     =   36963
      End
      Begin VB.Label lblPlanillaG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   90
         TabIndex        =   29
         Top             =   630
         Width           =   3930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.Frame fraTC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Tipo Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1095
      Left            =   7125
      TabIndex        =   2
      Top             =   45
      Width           =   2055
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TCC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblTCC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TCF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   630
         Width           =   420
      End
      Begin VB.Label lblTCF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6030
      TabIndex        =   1
      Top             =   6345
      Width           =   975
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   180
      Left            =   1215
      TabIndex        =   0
      Top             =   6420
      Visible         =   0   'False
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmRHPrestamosAdm.frx":0610
   End
End
Attribute VB_Name = "frmRHPrestamosAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sCodEmp As String

Private Sub BuscaDatosPlanilla()
    Dim rsEmp As ADODB.Recordset
    Dim sPla As String, sFec As String
    Dim L As ListItem
    Dim oCred As DRHPrestamosAdm
    Set oCred = New DRHPrestamosAdm
    
    sPla = Me.txtPlanillas.Text
    sFec = Format$(CDate(txtFecPla.value), "yyyymmdd")
    Set rsEmp = New ADODB.Recordset
    rsEmp.CursorLocation = adUseClient
    
    Set rsEmp = oCred.GetCreditosAdm(sPla, sFec)
    
    Set rsEmp.ActiveConnection = Nothing
    If Not (rsEmp.EOF And rsEmp.BOF) Then
        lstCredAdm.ListItems.Clear
        Do While Not rsEmp.EOF
            Set L = lstCredAdm.ListItems.Add(, , PstaNombre(rsEmp("cNomPers"), True))
            L.SubItems(1) = rsEmp("cCodCta")
            L.SubItems(2) = rsEmp("cNroCuota")
            L.SubItems(3) = Format(rsEmp("dFecVenc"), "dd/mm/yyyy")
            
            If gbBitCentral Then
                If Mid(L.SubItems(1), 9, 1) = Moneda.gMonedaNacional Then
                    L.SubItems(4) = Format$(CDbl(rsEmp("nMonto")), "#,##0.00")
                    L.SubItems(5) = "0.00"
                Else
                    L.SubItems(4) = Format$(CDbl(rsEmp("nMonto")) * CDbl(lblTCC), "#,##0.00")
                    L.SubItems(5) = Format$(CDbl(rsEmp("nMonto")), "#,##0.00")
                    L.ListSubItems(1).ForeColor = &H808000
                    L.ListSubItems(2).ForeColor = &H808000
                    L.ListSubItems(3).ForeColor = &H808000
                    L.ListSubItems(4).ForeColor = &H808000
                    L.ListSubItems(5).ForeColor = &H808000
                End If
            Else
                If Mid(L.SubItems(1), 6, 1) = Moneda.gMonedaNacional Then
                    L.SubItems(4) = Format$(CDbl(rsEmp("nMonto")), "#,##0.00")
                    L.SubItems(5) = "0.00"
                Else
                    L.SubItems(4) = Format$(CDbl(rsEmp("nMonto")) * CDbl(lblTCC), "#,##0.00")
                    L.SubItems(5) = Format$(CDbl(rsEmp("nMonto")), "#,##0.00")
                    L.ListSubItems(1).ForeColor = &H808000
                    L.ListSubItems(2).ForeColor = &H808000
                    L.ListSubItems(3).ForeColor = &H808000
                    L.ListSubItems(4).ForeColor = &H808000
                    L.ListSubItems(5).ForeColor = &H808000
                End If
            End If
            
            
            L.SubItems(6) = rsEmp("cPersCod")
            L.SubItems(7) = rsEmp("cEmpCod")
            L.Checked = True
            Set L = Nothing
            rsEmp.MoveNext
        Loop
        fraPlanilla.Enabled = False
    Else
       MsgBox "No se han registrado prestamos para esta planilla", vbInformation, "Aviso"
    End If
    fraTipoBus.Enabled = True
    optTipo(0).SetFocus
    rsEmp.Close
    Set rsEmp = Nothing
End Sub

Private Sub BuscaCodEmpleado()
    Dim L As ListItem
    Dim rsEmp As ADODB.Recordset
    Set rsEmp = New ADODB.Recordset
    rsEmp.CursorLocation = adUseClient
    'For Each L In lstCredAdm.ListItems
    '    If L.SubItems(7) = "" Then
    '        VSQL = "Select cEmpCod From Empleado Where cCodPers = '" & L.SubItems(6) & "'"
    '        rsEmp.Open VSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    '        Set rsEmp.ActiveConnection = Nothing
    '        If Not (rsEmp.EOF And rsEmp.BOF) Then
    '            L.SubItems(7) = rsEmp("cEmpCod")
    '        End If
    '        rsEmp.Close
    '    End If
    'Next
    'Set rsEmp = Nothing
End Sub

Private Sub HabilitaGrabar()
Dim L As ListItem
Dim bHab As Boolean
bHab = False
For Each L In lstCredAdm.ListItems
    If L.Checked Then
        bHab = True
        Exit For
    End If
Next
cmdGrabar.Enabled = bHab
End Sub

Private Sub cboPla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtFecPla.SetFocus
End If
End Sub

Private Sub chkCred_Click()
    Dim i As Integer
    For i = 1 To Me.lstCredAdm.ListItems.Count
        If Me.chkCred.value = 0 Then
            lstCredAdm.ListItems(i).Checked = False
        Else
            lstCredAdm.ListItems(i).Checked = True
        End If
 
    Next i
End Sub

Private Sub cmdAplicar_Click()
Dim L As ListItem
Set L = lstCredAdm.SelectedItem
If gbBitCentral Then
    If Mid(L.SubItems(1), 9, 1) = Moneda.gMonedaNacional Then
        L.SubItems(4) = txtMonto
    Else
        L.SubItems(4) = txtMonto
        L.SubItems(5) = txtMontoME
        txtMontoME = ""
        lblDolares.Visible = False
        txtMontoME.Visible = False
    End If
Else
    If Mid(L.SubItems(1), 6, 1) = Moneda.gMonedaNacional Then
        L.SubItems(4) = txtMonto
    Else
        L.SubItems(4) = txtMonto
        L.SubItems(5) = txtMontoME
        txtMontoME = ""
        lblDolares.Visible = False
        txtMontoME.Visible = False
    End If
End If
L.Checked = True
txtMonto = ""
lblSoles.Visible = False
txtMonto.Visible = False
cmdAplicar.Visible = False
cmdCancelar.Enabled = True
cmdGrabar.Enabled = True
lstCredAdm.SetFocus
End Sub

Private Sub cmdBuscaPla_Click()
    If Me.txtPlanillas.Text = "" Then
        MsgBox "Debe elegir una planilla.", vbInformation, "Aviso"
        Me.txtPlanillas.SetFocus
        Exit Sub
    End If

    BuscaDatosPlanilla
    fraPlanilla.Enabled = False
End Sub

Private Sub cmdBuscar_Click()
    Dim VSQL  As String
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH

    If optTipo(0).value Then
        If Trim(txtEmp) = "" Then
            MsgBox "Código de empleado no válido.", vbInformation, "Aviso"
            txtEmp.SetFocus
            Exit Sub
        End If
        Dim rsEmp As ADODB.Recordset
        Set rsEmp = New ADODB.Recordset
        rsEmp.CursorLocation = adUseClient
        
        sCodEmp = oRH.GetCodigoEmpleadoPers(Trim(txtEmp))
        If sCodEmp = "" Then
            MsgBox "Código de empleado no encontrado.", vbInformation, "Aviso"
            Set rsEmp = Nothing
            Exit Sub
        End If
    Else
        sCodEmp = ""
    End If
    frmRHPrestAdmAge.Show 1
    If lstCredAdm.ListItems.Count > 0 Then
        cmdEdit.Enabled = True
        lstCredAdm.SetFocus
        BuscaCodEmpleado
    End If
End Sub

Private Sub CmdCancelar_Click()
fraPlanilla.Enabled = True
fraTipoBus.Enabled = True
lstCredAdm.ListItems.Clear
txtEmp = ""
optTipo(0).value = True
fraTipoBus.Enabled = False
End Sub

Private Sub cmdedit_Click()
Dim L As ListItem
Set L = lstCredAdm.SelectedItem

lblSoles.Visible = True
txtMonto.Visible = True
txtMonto = L.SubItems(4)

If gbBitCentral Then
    If Mid(L.SubItems(1), 9, 1) = Moneda.gMonedaNacional Then
        cmdAplicar.Left = 3120
    Else
        lblDolares.Visible = True
        txtMontoME.Visible = True
        cmdAplicar.Left = 4680
        txtMontoME = L.SubItems(5)
    End If
Else
    If Mid(L.SubItems(1), 6, 1) = Moneda.gMonedaNacional Then
        cmdAplicar.Left = 3120
    Else
        lblDolares.Visible = True
        txtMontoME.Visible = True
        cmdAplicar.Left = 4680
        txtMontoME = L.SubItems(5)
    End If
End If

cmdAplicar.Visible = True
txtMonto.SetFocus
Set L = Nothing
cmdCancelar.Enabled = False
cmdGrabar.Enabled = False
End Sub

Private Sub cmdGrabar_Click()
    Dim L As ListItem
    Dim sPla As String, sFec As String, sFecAct As String
    Dim sEmp As String, sCta As String, sPers As String
    Dim nMonto As Double
    Dim sTCC As String, sCuota As String, sFecVenc As String
    Dim sqlEmpCon As String
    Dim rsEmpCon As ADODB.Recordset
    Set rsEmpCon = New ADODB.Recordset
    Dim oCred As DRHPrestamosAdm
    Set oCred = New DRHPrestamosAdm
    Dim oMov As DMov
    Set oMov = New DMov
    Dim lsMovNro As String
    
    lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        
    Me.MousePointer = 11
    sPla = Me.txtPlanillas.Text
    sFec = Format$(CDate(txtFecPla.value), gsFormatoMovFecha)
    sFecAct = FechaHora(gdFecSis)
    sTCC = lblTCC
     
    If MsgBox("Desea grabar la informacion??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    
    oCred.BeginTrans
    'Elimina todo lo existente en la tabla de conceptos por empleado y en la tabla de prestamos
    oCred.EliminaCreditosAdm Me.txtPlanillas.Text, Format(CDate(Me.txtFecPla), gsFormatoMovFecha)
    
    'Recorre la lista y graba todo lo seleccionado
    For Each L In lstCredAdm.ListItems
        If L.Checked Then
            sCta = L.SubItems(1)
            sCuota = L.SubItems(2)
            sEmp = L.SubItems(7)
            sPers = IIf(Len(L.SubItems(6)) = 10, "112" & L.SubItems(6), L.SubItems(6))
            sFecVenc = Format$(CDate(L.SubItems(3)), gsFormatoFecha)
            
            If gbBitCentral Then
                If Mid(sCta, 9, 1) = Moneda.gMonedaNacional Then
                    nMonto = CDbl(L.SubItems(4))
                    sTCC = "NULL"
                Else
                    nMonto = CDbl(L.SubItems(5))
                    sTCC = lblTCC
                End If
            Else
                If Mid(sCta, 6, 1) = Moneda.gMonedaNacional Then
                    nMonto = CDbl(L.SubItems(4))
                    sTCC = "NULL"
                Else
                    nMonto = CDbl(L.SubItems(5))
                    sTCC = lblTCC
                End If
            End If
            oCred.InsertaCreditosAdm sPers, sPla, sFec, sCta, sCuota, CDate(sFecVenc), CCur(nMonto), sTCC, "0", lsMovNro, Format(L.SubItems(4), "#.00")
            
        End If
    Next
    
    oCred.CommitTrans
    MsgBox "Grabación completa"
    Me.MousePointer = 0
    CmdCancelar_Click
End Sub

Private Sub cmdImprimir_Click()
    Dim i As Integer
    Dim lsCadena As String
    Dim lnTotal As Currency
    Dim lnTotalD As Currency
    Dim lnPagina As Long
    Dim lnItem As Long
    
    Dim lsNombre As String * 45
    Dim lsCredito As String * 19
    Dim lsMontoS As String * 15
    Dim lsMontoD As String * 15
    Dim lsFVen As String * 10
    
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    
    lsCadena = ""
    lnTotal = 0
    lnTotalD = 0
    lsCadena = lsCadena & CabeceraPagina("CREDITOS ADMINISTRATIVOS", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
    lsCadena = lsCadena & Encabezado("Nombre;10; ;26;Credito;20;Monto S/;25;Monto $;17; ;5;F.Venc;10; ;5;", lnItem)
    For i = 1 To Me.lstCredAdm.ListItems.Count
        If Me.lstCredAdm.ListItems(i).Checked = IIf(Me.chkSoloMarcados.value = 1, True, False) And Mid(Me.lstCredAdm.ListItems(i).ListSubItems(1), IIf(gbBitCentral, 9, 6), 1) = Moneda.gMonedaNacional Then
            
            lsNombre = Me.lstCredAdm.ListItems(i)
            lsCredito = Me.lstCredAdm.ListItems(i).ListSubItems(1)
            RSet lsMontoS = Me.lstCredAdm.ListItems(i).ListSubItems(4)
            RSet lsMontoD = Me.lstCredAdm.ListItems(i).ListSubItems(5)
            lsFVen = Me.lstCredAdm.ListItems(i).ListSubItems(3)
            
            lnTotal = lnTotal + CCur(Me.lstCredAdm.ListItems(i).ListSubItems(4))
            lnTotalD = lnTotalD + CCur(Me.lstCredAdm.ListItems(i).ListSubItems(5))
            
            lsCadena = lsCadena & lsNombre & "  " & lsCredito & "  " & lsMontoS & "  " & lsMontoD & "        " & lsFVen & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("CREDITOS ADMINISTRATIVOS", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
                lsCadena = lsCadena & Encabezado("Nombre;10; ;26;Credito;20;Monto S/;25;Monto $;17; ;5;F.Venc;10; ;5;", lnItem)
            End If
        End If
    Next i
    lsNombre = ""
    lsCredito = ""
    RSet lsMontoS = Format(lnTotal, "#,##0.00")
    RSet lsMontoD = Format(lnTotalD, "#,##0.00")
    lsCadena = lsCadena & String(116, "=") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsNombre & "  " & lsCredito & "  " & lsMontoS & "  " & lsMontoD & oImpresora.gPrnSaltoLinea
    
    lnTotal = 0
    lnTotalD = 0
    lsCadena = lsCadena & CabeceraPagina("CREDITOS ADMINISTRATIVOS", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "2")
    lsCadena = lsCadena & Encabezado("Nombre;10; ;26;Credito;20;Monto S/;25;Monto $;17; ;5;F.Venc;10; ;5;", lnItem)
    For i = 1 To Me.lstCredAdm.ListItems.Count
        If Me.lstCredAdm.ListItems(i).Checked = IIf(Me.chkSoloMarcados.value = 1, True, False) And Mid(Me.lstCredAdm.ListItems(i).ListSubItems(1), IIf(gbBitCentral, 9, 6), 1) = "2" Then
            
            lsNombre = Me.lstCredAdm.ListItems(i)
            lsCredito = Me.lstCredAdm.ListItems(i).ListSubItems(1)
            RSet lsMontoS = Me.lstCredAdm.ListItems(i).ListSubItems(4)
            RSet lsMontoD = Me.lstCredAdm.ListItems(i).ListSubItems(5)
            lsFVen = Me.lstCredAdm.ListItems(i).ListSubItems(3)
            
            lnTotal = lnTotal + CCur(Me.lstCredAdm.ListItems(i).ListSubItems(4))
            lnTotalD = lnTotalD + CCur(Me.lstCredAdm.ListItems(i).ListSubItems(5))
            
            lsCadena = lsCadena & lsNombre & "  " & lsCredito & "  " & lsMontoS & "  " & lsMontoD & "        " & lsFVen & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("CREDITOS ADMINISTRATIVOS", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
                lsCadena = lsCadena & Encabezado("Nombre;10; ;26;Credito;20;Monto S/;25;Monto $;17; ;5;F.Venc;10; ;5;", lnItem)
            End If
        End If
    Next i
    lsNombre = ""
    lsCredito = ""
    RSet lsMontoS = Format(lnTotal, "#,##0.00")
    RSet lsMontoD = Format(lnTotalD, "#,##0.00")
    lsCadena = lsCadena & String(116, "=") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsNombre & "  " & lsCredito & "  " & lsMontoS & "  " & lsMontoD & oImpresora.gPrnSaltoLinea
    
    oPrevio.Show lsCadena, Caption, True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Préstamos Administrativos"
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    GetTipCambio gdFecSis, Not gbBitCentral
    lblTCC = Format$(gnTipCambioV, "#0.0000")
    lblTCF = Format$(gnTipCambio, "#0.0000")
             
    optTipo(0).value = True
    
    Me.txtPlanillas.rs = oPla.GetPlanillas(, True)

    cmdGrabar.Enabled = False
    cmdAplicar.Visible = False
    txtMonto.Visible = False
    cmdEdit.Enabled = False
    lblSoles.Visible = False
    lblDolares.Visible = False
    txtMontoME.Visible = False
    txtFecPla.value = gdFecSis
    fraTipoBus.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub lstCredAdm_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCredAdm.SortKey = ColumnHeader.Index - 1
    lstCredAdm.Sorted = True
End Sub

Private Sub lstCredAdm_DblClick()
If lstCredAdm.ListItems.Count > 0 Then
    If gbBitCentral Then
        frmCredConsulta.ConsultaCliente Trim(lstCredAdm.SelectedItem.SubItems(1))
    Else
        frmRHPstaCred.frmIni Trim(lstCredAdm.SelectedItem.SubItems(1)), "VIGENTE"
    End If
End If
End Sub

Private Sub lstCredAdm_ItemCheck(ByVal Item As MSComctlLib.ListItem)
HabilitaGrabar
End Sub

Private Sub lstCredAdm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If lstCredAdm.ListItems.Count > 0 Then
        If gbBitCentral Then
            frmCredConsulta.ConsultaCliente Trim(lstCredAdm.SelectedItem.SubItems(1))
        Else
            frmRHPstaCred.frmIni Trim(lstCredAdm.SelectedItem.SubItems(1)), "VIGENTE"
        End If
   End If
End If
End Sub

Private Sub OptTipo_Click(Index As Integer)
Select Case Index
    Case 0
        txtEmp.Visible = True
    Case 1
        txtEmp.Visible = False
End Select
End Sub

Private Sub optTipo_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0
        If KeyAscii = 13 Then
            txtEmp.SetFocus
        End If
    Case 1
        If KeyAscii = 13 Then
            cmdBuscar.SetFocus
        End If
End Select
End Sub

Private Sub txtEmp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(txtEmp) <> "" Then
        txtEmp = Left(Trim(txtEmp), 1) & FillNum(Mid(Trim(txtEmp), 2), 5, "0")
    End If
    cmdBuscar.SetFocus
Else
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End If
End Sub


Private Sub txtMonto_GotFocus()
With txtMonto
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
Dim sCar As String * 1
If KeyAscii = 13 Then
    txtMonto = Format$(txtMonto, "#,##0.00")
    If txtMontoME.Visible Then
        txtMontoME.SetFocus
    Else
        cmdAplicar.SetFocus
    End If
Else
    sCar = Chr$(KeyAscii)
    If InStr(1, "0123456789.", sCar, vbTextCompare) = 0 And KeyAscii <> 8 Then
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtMonto_LostFocus()
txtMonto = Format$(CDbl(txtMonto), "#,##0.00")
If txtMonto = "" Then txtMonto = "0.00"
If txtMontoME.Visible Then
    txtMontoME = Format$(CDbl(txtMonto) / CDbl(lblTCC), "##0.00")
End If
End Sub

Private Sub txtMontoME_GotFocus()
With txtMontoME
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMontoME_KeyPress(KeyAscii As Integer)
Dim sCar As String * 1
If KeyAscii = 13 Then
    txtMontoME = Format$(txtMontoME, "#,##0.00")
    cmdAplicar.SetFocus
Else
    sCar = Chr$(KeyAscii)
    If InStr(1, "0123456789.", sCar, vbTextCompare) = 0 And KeyAscii <> 8 Then
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtMontoME_LostFocus()
txtMontoME = Format$(CDbl(txtMontoME), "#,##0.00")
If txtMontoME = "" Then
    txtMontoME = "0.00"
End If
txtMonto = Format$(CDbl(txtMontoME) * CDbl(lblTCC), "#,##0.00")
End Sub

Private Sub txtPlanillas_EmiteDatos()
    Me.lblPlanillaG.Caption = Me.txtPlanillas.psDescripcion
End Sub

