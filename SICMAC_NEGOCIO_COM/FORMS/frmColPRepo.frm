VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColPRepo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones Pignoraticio - Reportes"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   Icon            =   "frmColPRepo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEnviaExcel 
      Caption         =   "Enviar a Excel"
      Height          =   195
      Left            =   7080
      TabIndex        =   32
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Frame fraEstadoxDias 
      Height          =   975
      Left            =   8670
      TabIndex        =   28
      Top             =   3240
      Width           =   1575
      Begin VB.CheckBox chkEstadoxDias 
         Caption         =   "Para Remate"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkEstadoxDias 
         Caption         =   "Vencidos"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkEstadoxDias 
         Caption         =   "Normales"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame fraRenovado 
      Height          =   975
      Left            =   6990
      TabIndex        =   25
      Top             =   3240
      Width           =   1695
      Begin VB.CheckBox chkRenovado 
         Caption         =   "Renovados"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkRenovado 
         Caption         =   "No Renovados"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdBoveda 
      Caption         =   "&Bovedas ..."
      Height          =   375
      Left            =   8670
      TabIndex        =   24
      Top             =   4920
      Width           =   1125
   End
   Begin VB.CommandButton cmdAgencia 
      Caption         =   "A&gencias..."
      Height          =   345
      Left            =   8670
      TabIndex        =   11
      Top             =   4560
      Width           =   1140
   End
   Begin VB.Frame fraImpresion 
      Caption         =   "Impresión "
      Height          =   1095
      Left            =   6990
      TabIndex        =   23
      Top             =   4320
      Width           =   1545
      Begin VB.OptionButton optImpresion 
         Caption         =   "Pantalla"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optImpresion 
         Caption         =   "Impresora"
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   990
      End
      Begin VB.OptionButton optImpresion 
         Caption         =   "Archivo"
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   990
      End
   End
   Begin VB.Frame fraUsuario 
      Caption         =   "Usuario"
      Height          =   735
      Left            =   7680
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         ItemData        =   "frmColPRepo.frx":030A
         Left            =   120
         List            =   "frmColPRepo.frx":0311
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame fraDiasAtraso1 
      Caption         =   "Dias de Atraso"
      Height          =   660
      Left            =   7680
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   1995
      Begin VB.TextBox txtDiasAtraso1A 
         Height          =   255
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   9
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDiasAtraso1De 
         Height          =   255
         Left            =   480
         MaxLength       =   5
         TabIndex        =   8
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "A :"
         Height          =   225
         Left            =   1065
         TabIndex        =   21
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label4 
         Caption         =   "De :"
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fraPeriodo2 
      Caption         =   "Período"
      Height          =   675
      Left            =   7680
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   2025
      Begin MSMask.MaskEdBox mskPeriodo2Al 
         Height          =   315
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label26 
         Caption         =   "Al :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Período"
      Height          =   975
      Left            =   7680
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2010
      Begin MSMask.MaskEdBox mskPeriodo1Al 
         Height          =   330
         Left            =   555
         TabIndex        =   6
         Top             =   555
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPeriodo1Del 
         Height          =   315
         Left            =   555
         TabIndex        =   5
         Top             =   210
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Al :"
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   615
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Del :"
         Height          =   255
         Left            =   135
         TabIndex        =   15
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8430
      TabIndex        =   3
      Top             =   6015
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6990
      TabIndex        =   2
      Top             =   6015
      Width           =   1215
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Seleccione Operación"
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
      Height          =   6435
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   240
         Top             =   5640
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
               Picture         =   "frmColPRepo.frx":0324
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmColPRepo.frx":0676
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmColPRepo.frx":09C8
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmColPRepo.frx":0D1A
               Key             =   "Bebe"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwReporte 
         Height          =   6075
         Left            =   105
         TabIndex        =   1
         Top             =   240
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   10716
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmColPRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************
'* REPORTES DE PIGNORATICIO *
'Archivo:  frmColPRepo.frm
'LAYG   :  15/09/2001.
'Resumen:  Nos permite emitir los Reportes de Pignoraticio
Option Explicit

Private MatAgencias() As String ''*** PEAC 20090521

Dim fnRepoSelec As Long
Dim fsAgenciasSelect As String  ' Agencias Seleccionadas
Dim fsBovedasSelect As String  ' Bovedas seleccionadas
Dim fnTipoReporte As Integer ' 1 Diario / 2 Generales / 3 Estadisticas

Public Sub Inicio(ByVal pnTipoReporte As Integer)
    
    fnTipoReporte = pnTipoReporte
    CargaMenu
    Me.Show 1

End Sub

Private Sub CargaMenu()
Dim clsGen As DGeneral  'COMDConstSistema.DCOMGeneral ARCV 25-10-2006
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node
Dim lsTipREP As String

Select Case fnTipoReporte
    Case 1
        lsTipREP = "12800%' OR cOpeCod LIKE '12801%' OR cOpeCod LIKE '12802" '"1280[012]"
        Me.Caption = "Pignoraticio - Reportes "
    Case 2
        lsTipREP = "12803" '"1280[3]"
        Me.Caption = "Pignoraticio - Reportes "
    Case 3
        lsTipREP = "12804" '"1280[4]"
        Me.Caption = "Pignoraticio - Reportes "
End Select

Set clsGen = New DGeneral  'COMDConstSistema.DCOMGeneral

'ARCV 20-07-2006
'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
Set rsUsu = clsGen.GetOperacionesUsuario_NEW(lsTipREP, , gRsOpeRepo)

Set clsGen = Nothing
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = tvwReporte.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvwReporte.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = tvwReporte.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = tvwReporte.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
            nodOpe.Tag = sOpeCod
    End Select
    rsUsu.MoveNext
Loop
rsUsu.Close
Set rsUsu = Nothing
End Sub
Private Sub cmdAgencia_Click() 'Selecciona Agencias
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub cmdBoveda_Click() ' Selecciona Bovedas
    frmColPSelectBoveda.Inicio Me
    frmColPSelectBoveda.Show 1
End Sub

Private Sub cmdImprimir_Click()
Dim NodRep  As Node
Dim lsDesc As String

'*** PEAC 20171013
If Not (IsDate(Me.mskPeriodo2Al.Text)) Then
    MsgBox "Por favor, ingrese una fecha correcta.", vbInformation, "Atención"
    Exit Sub
End If

Set NodRep = tvwReporte.SelectedItem
    If NodRep Is Nothing Then
       Exit Sub
    End If
    lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
    fnRepoSelec = CLng(NodRep.Tag)
    ' Valida datos para Reporte
    VerificaDatosParaReporte CLng(NodRep.Tag)
    ' Ejecuta Reporte
    EjecutaReporte CLng(NodRep.Tag), lsDesc
Set NodRep = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me '*
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Me.mskPeriodo1Del = Format$(gdFecSis, "dd/mm/yyyy")
Me.mskPeriodo2Al = Format$(gdFecSis, "dd/mm/yyyy")
Me.mskPeriodo1Al = Format$(gdFecSis, "dd/mm/yyyy")

End Sub

Private Sub VerificaDatosParaReporte(ByVal pnOperacion As ColocPOperaciones)

fsAgenciasSelect = fgCargaAgenciasSelec() ' Carga Agencia a seleccionar
fsBovedasSelect = fgCargaBovedaSelec() ' Carga Bovedas a seleccionar
'Verifica que se haya seleccionado Agencia

    Select Case pnOperacion
        ' Reporte de Transacciones
        Case 128001  ' Reporte de Registro de Contratos
            Call ValidaControles(False, True, False, False, True)
        Case 128002  ' Reporte de Desembolsos
            Call ValidaControles(False, True, False, False, True)
        Case 128003  ' Reporte de Anulados
            Call ValidaControles(False, True, False, False, True)
        Case 128004  ' Reporte de Renovados
            Call ValidaControles(False, True, False, False, True)
        Case 128005  ' Reporte de Cancelados
            Call ValidaControles(False, True, False, False, True)
        Case 128006  '
            Call ValidaControles(False, True, False, False, True)
        Case 128007  ' Reporte de Pago Sobrantes
            Call ValidaControles(False, True, False, False, True)
        Case 128008  ' Reporte de Operaciones Extornadas
            Call ValidaControles(False, True, False, False, True)
        Case 128012  ' CROB20180612 ERS076-2017 Reporte de Contratos Rechazados
            Call ValidaControles(False, True, False, False, True) ' CROB20180612 ERS076-2017
    End Select

End Sub
Private Sub EjecutaReporte(ByVal pnOperacion As ColocPOperaciones, ByVal psDescOperacion As String)
Dim loRep As COMNColoCPig.NCOMColPRepo
Dim lsCadImp As String
Dim loPrevio As previo.clsprevio
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim X As Integer
Dim sCadAge As String
Dim sCadBove As String
Dim i As Integer
Dim lsmensaje As String

Dim nContAge As Integer, sAgenciasTemp As String, nContAgencias As Integer


''*** PEAC 20090521
    ReDim MatAgencias(0)
    nContAge = 0
    sAgenciasTemp = ""
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            nContAge = nContAge + 1
            nContAgencias = nContAgencias + 1
            ReDim Preserve MatAgencias(nContAge)
            MatAgencias(nContAge - 1) = Mid(frmSelectAgencias.List1.List(i), 1, 2)

            If Len(Trim(sCadAge)) = 0 Then
                sCadAge = "'" & Mid(frmSelectAgencias.List1.List(i), 1, 2) & "'"
                sAgenciasTemp = "" & Mid(frmSelectAgencias.List1.List(i), 1, 2) & ""
            Else
                sCadAge = sCadAge & ", '" & Mid(frmSelectAgencias.List1.List(i), 1, 2) & "'"
                sAgenciasTemp = sAgenciasTemp & ", " & Mid(frmSelectAgencias.List1.List(i), 1, 2) & ""
            End If

        End If
    Next i
    If nContAge = 0 Then
        ReDim MatAgencias(1)
        nContAgencias = 1
        MatAgencias(0) = gsCodAge
    End If
''****************************************


Set loRep = New COMNColoCPig.NCOMColPRepo
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    Select Case pnOperacion
        ' Reporte de Transacciones
        Case 128001  ' Reporte de Registro de Contratos
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128001_ContratoRegis(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"

        Case 128002  ' Reporte de Contratos Desembolsados
'            For X = 1 To frmSelectAgencias.List1.ListCount
'                If frmSelectAgencias.List1.Selected(X - 1) = True Then
'                    If Len(lscadimp) > 0 Then lscadimp = lscadimp & Chr(12)
'                    lscadimp = lscadimp & loRep.nRepo128002_ContratoDesemb(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
'                    If Trim(lsmensaje) <> "" Then
'                         MsgBox lsmensaje, vbInformation, "Aviso"
'                         Exit Sub
'                    End If
'                End If
'            Next X

'            If frmSelectAgencias.List1.Selected(X - 1) = True Then
                lsCadImp = lsCadImp & loRep.nRepo128002_ContratoDesemb(pnOperacion, fsAgenciasSelect, Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                If Trim(lsmensaje) <> "" Then
                    MsgBox lsmensaje, vbInformation, "Aviso"
                    Exit Sub
                End If
 '           End If
            lsDestino = "P"

        Case 128003  ' Reporte de Contratos Anulados
            Screen.MousePointer = 11
            If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then
                Call ReporteCredPigAnuladosExcel
            Else
                For X = 1 To frmSelectAgencias.List1.ListCount
                    If frmSelectAgencias.List1.Selected(X - 1) = True Then
                        If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                        lsCadImp = lsCadImp & loRep.nRepo128003_ContratoAnula(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
                    End If
                Next X
                lsDestino = "P"
            End If
            Screen.MousePointer = 0

        Case 128004  ' Reporte de Contratos Renovados
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128004_ContratoRenov(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        
        Case 128005  ' Reporte de Contratos Cancelados
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128005_ContratoCancel(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        
        
        Case 128006 ' REPORTE CREDITOS AMORTIZADOS PIGNORATICIO DIARIO
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128007_ContratoAmortiz(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
            
        Case 128007
            For X = 1 To frmSelectAgencias.List1.ListCount
               If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128007_ContratoRecup(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
            
        Case 128009
            For X = 1 To frmSelectAgencias.List1.ListCount
               If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128009_ContratoVend(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        
        
        Case 128008  ' Reporte de Pago de Sobrantes
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128008_PagoSobrantes(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        'RECO20141028 ERS125-2014****************
        Case 128011 'REPORTE SEGUIMIENTO CAMPAÑA CALL CENTER
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128011_SeguimientoCampCallCenter(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        'RECO FIN *******************************
        
        'CROB20180612 ERS076-2017 ****
        Case 128012  ' Reporte de Contratos Rechazados
'            Screen.MousePointer = 11
'            If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then
'                Call ReporteCredPigAnuladosExcel
'            Else
                For X = 1 To frmSelectAgencias.List1.ListCount
                    If frmSelectAgencias.List1.Selected(X - 1) = True Then
                        If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                        lsCadImp = lsCadImp & loRep.nRepo128012_ContratoRechazo(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
                    End If
                Next X
                lsDestino = "P"
'            End If
'            Screen.MousePointer = 0
        'CROB20180612 ERS076-2017 ****
        
        Case 128021, 128022, 128023, 128024, 128025 ' Reporte de Prendas Nuevas en Boveda // Rescatadas  // Nuevas Condicion Diferidas
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo12802_MovimientoJoyas(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         'Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        'RECO20140208 ERS002**********************************************
        Case 128026
             For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRS128026_PrendasModalidadAmpliados(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo1Del.Text, "yyyy/mm/dd"), Format(Me.mskPeriodo1Al.Text, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        'RECO FIN***********************************************************
        Case 128031  ' Listado Prestamos Vigentes
            Dim lsRenov As String
            Dim lsEstad As String
            lsRenov = lsRenov & IIf(Me.chkRenovado(0).value = 1, "1", "0")
            lsRenov = lsRenov & IIf(Me.chkRenovado(1).value = 1, "1", "0")
            lsEstad = lsEstad & IIf(Me.chkEstadoxDias(0).value = 1, "1", "0")
            lsEstad = lsEstad & IIf(Me.chkEstadoxDias(1).value = 1, "1", "0")
            lsEstad = lsEstad & IIf(Me.chkEstadoxDias(2).value = 1, "1", "0")
            
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128031_ListadoCredVigentes(mskPeriodo1Del.Text, mskPeriodo1Al.Text, pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), lsRenov, lsEstad, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        
        Case 128032  ' Listado Prestamos en Condicion de Diferidos
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128032_ListadoCredDiferidos(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
                   
        Case 128033  ' Listado Prestamos x Dias de Atraso (Interes Devengado)
            Dim lnContador As Integer   'RECO201408012 ERS064-2014
            Dim lnConAgeCheck As Integer
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    lnConAgeCheck = lnConAgeCheck + 1   'RECO201408012 ERS064-2014
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128033_ListadoCredxDiasAtraso(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), val(Me.txtDiasAtraso1De.Text), val(Me.txtDiasAtraso1A.Text), gdFecSis, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         lnContador = lnContador + 1    'RECO201408012 ERS064-2014
                         'MsgBox lsmensaje, vbInformation, "Aviso"  'RECO201408012 ERS064-2014
                         'Exit Sub                                  'RECO201408012 ERS064-2014
                    End If
                End If
            Next X
            'RECO201408012 ERS064-2014*****************************
                If lnContador = lnConAgeCheck Then
                    MsgBox "No existen datos para generar el reporte", vbInformation, "Aviso"
                End If
            'END RECO**********************************************
            lsDestino = "P"
        
        Case 128034  ' Listado Prestamos en Condicion de Adjudicados
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128034_ListadoCredAdjudicados(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"

        Case 128035  ' Listado Prestamos Vigentes DETALLADO POR LOTES ' peac 20071010
        Screen.MousePointer = 11
        If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then
            Call ReporteVigentesDetalladoPorLoteExcel
        Else
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                        lsCadImp = lsCadImp & loRep.nRepo128035_ListadoCredVigentes(mskPeriodo1Del.Text, mskPeriodo1Al.Text, pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), lsRenov, lsEstad, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                        MsgBox lsmensaje, vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        End If
        Screen.MousePointer = 0
        
        Case 128036
            Dim sTempo As String
            
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    sTempo = loRep.nRepo128036_Creditos_prendarios_fecha_determinada(Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), mskPeriodo2Al.Text, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         'Exit Sub
                    End If
                    If Len(Trim(sTempo)) > 0 Then
                        If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                        lsCadImp = lsCadImp & sTempo
                    End If
                End If
            Next X
            lsDestino = "P"
                
        Case 128037  '*** PEAC 20090312 - REPORTE DE SOBRANTE DE ADJUDICADO
        Screen.MousePointer = 11
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                        lsCadImp = lsCadImp & loRep.nRepo128037_ReporteSobranteAdjudicados(pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                        MsgBox lsmensaje, vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        Screen.MousePointer = 0
        
        Case 128038  '*** PEAC 20090412 - REPORTE DE CREDITOS ADJUDICADOS DETALLADO POR LOTE
        Screen.MousePointer = 11
        If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then
            Call ReporteAdjudicadosDetalladoPorLoteExcel
        Else
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                        lsCadImp = lsCadImp & loRep.nRepo128038_ListadoCredAdjdicadosXLote(mskPeriodo1Del.Text, mskPeriodo1Al.Text, pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), lsRenov, lsEstad, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                        MsgBox lsmensaje, vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        End If
        Screen.MousePointer = 0
        
        Case 128039  '*** PEAC 20090521 - REPORTE DE CREDITOS diferidos DETALLADO POR LOTE
        Screen.MousePointer = 11
        If Me.chkEnviaExcel.Visible And Me.chkEnviaExcel.value = 1 Then
            Call ReporteDiferidosDetalladoPorLoteExcel
        Else
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                        lsCadImp = lsCadImp & loRep.nRepo128039_ListadoCredDiferidosXLote(mskPeriodo1Del.Text, mskPeriodo1Al.Text, pnOperacion, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), lsRenov, lsEstad, fsBovedasSelect, False, lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                        MsgBox lsmensaje, vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
        End If
        Screen.MousePointer = 0
        
        Case 128041 ' Estadistica de Saldos
            lsCadImp = ""
            For X = 1 To frmSelectAgencias.List1.ListCount
                
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & loRep.nRepo128041_Estadistica_Saldos(mskPeriodo1Del.Text, mskPeriodo1Al.Text, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), gdFecSis, gsNomCmac, GetNombreAgencia(Mid(frmSelectAgencias.List1.List(X - 1), 1, 2)), lsmensaje, gImpresora)
                    If Trim(lsmensaje) <> "" Then
                         MsgBox lsmensaje, vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
            Next X
            lsDestino = "P"
            
         Case 128043 ' Estadistica de Saldos
            lsCadImp = ""
            For X = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(X - 1) = True Then
                    If Len(lsCadImp) > 0 Then lsCadImp = lsCadImp & Chr(12)
                    'lsCadImp = lsCadImp & loRep.nRepo128043_SaldosResumen(128043, Mid(frmSelectAgencias.List1.List(X - 1), 1, 2), Format(Me.mskPeriodo2Al, "yyyy/mm/dd"), Format(gdFecSis, "yyyy/mm/dd"), cboUsuario.Text, fsBovedasSelect)
                End If
            Next X
            lsDestino = "P"
            
            '*** PEAC 20170929
        Case 128044  'CLIENTES DE PIGNOS DEJADOS DE ATENDER
            Screen.MousePointer = 11
            Call ReporteClientesPignoDejadosDeAtenderExcel
            Screen.MousePointer = 0
            
            
          Case 128049 ' Estadistica de Oro
            lsCadImp = ""
            sCadAge = ""
            For i = 0 To frmSelectAgencias.List1.ListCount - 1
                If frmSelectAgencias.List1.Selected(i) = True Then
                    sCadAge = Left(frmSelectAgencias.List1.List(i), 2)
                    Exit For
                End If
            Next
            
            If Len(Trim(sCadAge)) = 0 Then
                MsgBox "Escoja  una agencia", vbInformation, "Aviso"
                Exit Sub
            End If
            lsCadImp = loRep.nRepo128049_Estadistica_Oro(sCadAge, Format(mskPeriodo1Del.Text, "yyyymmdd"), Format(mskPeriodo2Al.Text, "yyyymmdd"), gdFecSis, gsNomCmac, gsNomAge, gImpresora)
            lsDestino = "P"
            
    End Select
Set loRep = Nothing
    

If lsDestino = "P" Then
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
            loPrevio.Show lsCadImp, psDescOperacion, True
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
ElseIf lsDestino = "A" Then
'Else
'    dlgGrabar.CancelError = True
'    dlgGrabar.InitDir = App.Path
'    dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
'    dlgGrabar.ShowSave
'    If dlgGrabar.FileName <> "" Then
'       Open dlgGrabar.FileName For Output As #1
'        Print #1, vBuffer
'        Close #1
'    End If
End If
End Sub

'*** PEAC 20170330
Private Sub ReporteCredPigAnuladosExcel()
        
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer, lcAge As String, lcPersNom As String
Dim sCadAge As String
Dim lcCtaCod As String
Dim sBovedas As String
Dim conta As Integer

    For i = 0 To UBound(MatAgencias) - 1
        sCadAge = sCadAge & MatAgencias(i) & ","
        Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
       
    If Len(sCadAge) = 0 Then
        MsgBox "Por favor seleccione al menos una Agencia.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    sBovedas = Replace(Replace(Replace(fsBovedasSelect, "'", ""), "(", ""), ")", "")
    
    Dim oPig As COMNColoCPig.NCOMColPRepo
    Set oPig = New COMNColoCPig.NCOMColPRepo
        Set R = oPig.nObtieneCredPigAnulados(sCadAge, sBovedas, mskPeriodo1Del.Text, mskPeriodo1Al.Text)

    Set oPig = Nothing
       
    If (R.EOF And R.BOF) Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = gsNomCmac
    ApExcel.Cells(3, 2).Formula = gsNomAge  'UCase(R!CAgencia) 'gsNomAge
    ApExcel.Cells(2, 8).Formula = Date + Time()
    ApExcel.Cells(3, 8).Formula = gsCodUser
    ApExcel.Range("B2", "H8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("B8", "B8").HorizontalAlignment = xlLeft
    ApExcel.Range("H2", "H3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "H6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "H10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "H10").Borders.LineStyle = 1

    ApExcel.Cells(5, 2).Formula = "LISTADO DE CONTRATOS ANULADOS"
    ApExcel.Cells(6, 2).Formula = "Información del " & Format(mskPeriodo1Del.Text, "dd/MM/YYYY") & " al " & Format(mskPeriodo1Al.Text, "dd/MM/YYYY")
'    ApExcel.Cells(8, 2).Formula = "Tipo de Cambio : " & Trim(Str(pntipCam))
    
    ApExcel.Range("B5", "H5").MergeCells = True
    ApExcel.Range("B6", "H6").MergeCells = True

'    ApExcel.Range("B9", "C10").MergeCells = True
'    ApExcel.Range("D9", "E9").MergeCells = True
'    ApExcel.Range("F9", "G9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "ITEM"
    ApExcel.Cells(9, 3).Formula = "CONTRATO ANT."
    ApExcel.Cells(9, 4).Formula = "CONTRATO"
    ApExcel.Cells(9, 5).Formula = "CLIENTE"
    ApExcel.Cells(9, 6).Formula = "PRESTAMO"
    ApExcel.Cells(9, 7).Formula = "USUARIO"
    ApExcel.Cells(9, 8).Formula = "MOTIVO"
    
    ApExcel.Range("B9", "H10").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B9", "H10").Font.Bold = True
    ApExcel.Range("B9", "H10").HorizontalAlignment = 3

    i = 10
    conta = 0
    Do While Not R.EOF
    i = i + 1
           conta = conta + 1
            ApExcel.Cells(i, 2).Formula = conta
            ApExcel.Cells(i, 3).Formula = R!codigoant
            ApExcel.Cells(i, 4).Formula = R!cCtacod
            ApExcel.Cells(i, 5).Formula = R!nomcliente
            ApExcel.Cells(i, 6).Formula = R!nMontoCol
            ApExcel.Cells(i, 7).Formula = R!Usuario
            ApExcel.Cells(i, 8).Formula = R!cmovdesc
            
            ApExcel.Range("B" & Trim(str(i)) & ":" & "B" & Trim(str(i))).NumberFormat = "#,##0"
            ApExcel.Range("F" & Trim(str(i)) & ":" & "F" & Trim(str(i))).NumberFormat = "#,##0.00"
            ApExcel.Range("B" & Trim(str(i)) & ":" & "H" & Trim(str(i))).Borders.LineStyle = 1

            R.MoveNext
                
            If R.EOF Then
                Exit Do
            End If

    Loop

'    i = i + 2

'    ApExcel.Range("B" & Trim(Str(i)) & ":" & "H" & Trim(Str(i + 2))).Borders.LineStyle = 1
'    ApExcel.Range("B" & Trim(Str(i)) & ":" & "H" & Trim(Str(i + 2))).Font.Bold = True

    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing

End Sub

Private Sub mskPeriodo1Del_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskPeriodo1Al.SetFocus
End If
End Sub

Private Sub tvwReporte_NodeClick(ByVal Node As MSComctlLib.Node)

Dim NodRep  As Node
Dim lsDesc As String

Set NodRep = tvwReporte.SelectedItem

If NodRep Is Nothing Then
   Exit Sub
End If
lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
fnRepoSelec = CLng(NodRep.Tag)

Select Case fnRepoSelec
    Case 128001 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
      
    Case 128002 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
      
    Case 128003 ' reporte de creditos anulados
        Call HabilitaControles(True, True, True, True, False, False, True, False, False, True)
    
    Case 128004 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
      
    Case 128005 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
      
    Case 128006 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
    
    Case 128007 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
    
    Case 128009 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
    
    'RECO20141027 ERS125-2014
    Case 128011
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
        
    'CROB20180612 ERS076-2017
     Case 128012
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
    'CROB20180612 ERS076-2017
    Case 128008
    
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
    
    Case 128021 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
      
    Case 128022 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
      
    Case 128023 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
    Case 128024 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
    Case 128025 '
        Call HabilitaControles(True, True, True, True, False, False, True, False, False)
    Case 128026 ' RECO20140208 ERS002
        Call HabilitaControles(True, True, False, True, True, False, True, False, False)
    Case 128031 ' Listado de Creditos vigentes
        Call HabilitaControles(True, True, True, True, False, False, False, True, True)
      
    Case 128032 ' Listado de Creditos Diferidos
        Call HabilitaControles(True, True, True, False, False, False, False, False, False)
      
    Case 128033 ' Listado de Creditos x Dias Atraso
        Call HabilitaControles(True, True, True, False, False, True, False, False, False)
    
    Case 128034 ' Listado de Creditos Adjudicados
        Call HabilitaControles(True, True, True, False, False, False, False, False, False)
      
    'PEAC 20071018
    Case 128035 ' Listado de Creditos Adjudicados VIGENTES DETALLADO POR LOTES
        Call HabilitaControles(True, True, False, True, False, False, False, False, False, True)
      
    Case 128036 ' Creditos prendarios a una fecha determinada
        Call HabilitaControles(True, True, False, False, True, False, False, False, False)
      
      
    '*** PEAC 20090312
    Case 128037 ' REPORTE DE SOBRANTE DE ADJUDICADO
        Call HabilitaControles(True, True, False, False, False, False, False, False, False)
            
    Case 128038 ' REPORTE DE CREDITOS ADJUDICADOS DETALLADO POR LOTE
        Call HabilitaControles(True, True, False, True, False, False, False, False, False, True)
            
    Case 128039 ' REPORTE DE CREDITOS DIFERIDOS DETALLADO POR LOTE - PEAC 20090521
        Call HabilitaControles(True, True, False, False, False, False, False, False, False, True)
            
    Case 128041 ' Estadistica de Saldos
        Call HabilitaControles(True, True, True, True, False, False, False, False, False)
      
    Case 128042 ' Estadistica de Lotes
         Call HabilitaControles(True, True, True, True, False, False, False, False, False)
    
    Case 128043 'Estadistica de saldos
        Call HabilitaControles(True, True, True, False, True, False, True, False, False)
        
    '*** PEAC 20170929
    Case 128044 'clientes de pignos que se dejaron de atender
        Call HabilitaControles(True, True, False, False, True, False, False, False, False)
        
        fraPeriodo2.Caption = "Antiguedad:"
        
    Case 128049 'Estadistica de Oro
        Call HabilitaControles(True, True, False, True, False, False, False, False, False)

    
    Case Else
        Call HabilitaControles(False, False, False, False, False, False, False, False, False)
End Select

Set NodRep = Nothing
  
End Sub

Private Sub HabilitaControles(ByVal pbcmdImprimir As Boolean, ByVal pbCmdAgencia As Boolean, ByVal pbCmdBoveda As Boolean, _
        ByVal pbFraPeriodo1 As Boolean, _
        ByVal pbFraPeriodo2 As Boolean, ByVal pbFraDiasAtraso As Boolean, ByVal pbFraUsuario As Boolean, _
        ByVal pbFraRenovado As Boolean, ByVal pbFraEstadoxDias As Boolean, Optional ByVal pbEnviaExcel As Boolean = False)

    Me.cboUsuario.ListIndex = 0
    Me.cmdImprimir.Visible = pbcmdImprimir
    Me.cmdAgencia.Visible = pbCmdAgencia
    Me.cmdBoveda.Visible = pbCmdBoveda
    Me.fraPeriodo1.Visible = pbFraPeriodo1
    Me.fraPeriodo2.Visible = pbFraPeriodo2
    Me.fraDiasAtraso1.Visible = pbFraDiasAtraso
    Me.fraUsuario.Visible = pbFraUsuario
    
    Me.fraRenovado.Visible = pbFraRenovado
    Me.fraEstadoxDias.Visible = pbFraEstadoxDias

    Me.chkEnviaExcel.Visible = pbEnviaExcel ''*** PEAC 20090521

End Sub

Private Sub ValidaControles(ByVal pbcmdImprimir As Boolean, ByVal pbFraPeriodo1 As Boolean, _
        ByVal pbFraPeriodo2 As Boolean, ByVal pbFraDiasAtraso1 As Boolean, ByVal pbFraUsuario As Boolean)
Dim lbOk As Boolean

lbOk = True

    If pbcmdImprimir = True Then
    End If
    If pbFraPeriodo1 = True Then
        If Not (IsDate(Me.mskPeriodo1Del.Text) And IsDate(Me.mskPeriodo1Al.Text)) Then
            lbOk = False
            Me.mskPeriodo1Del.SetFocus
        End If
        If Format(Me.mskPeriodo1Del.Text, "dd/mm/yyyy") > Format(Me.mskPeriodo1Al.Text, "dd/mm/yyyy") Then
            lbOk = False
            Me.mskPeriodo1Del.SetFocus
        End If
    End If
    If pbFraPeriodo2 = True Then
        If Not (IsDate(Me.mskPeriodo2Al.Text)) Then
            lbOk = False
            Me.mskPeriodo2Al.SetFocus
        End If
    End If
    
    If pbFraDiasAtraso1 = True Then
    
    End If

End Sub

'*** PEAC 20170929
Private Sub ReporteClientesPignoDejadosDeAtenderExcel()
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer, lcAge As String, lcPersNom As String
Dim sCadAge As String
Dim lcCtaCod As String

    For i = 0 To UBound(MatAgencias) - 1
        sCadAge = sCadAge & MatAgencias(i) & ","
        Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
       
    If Len(sCadAge) = 0 Then
        MsgBox "Por favor seleccione al menos una Agencia.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Dim oPig As COMNColoCPig.NCOMColPRepo
    Set oPig = New COMNColoCPig.NCOMColPRepo
        Set R = oPig.nObtieneClientesPignoDejadosDeAtender(mskPeriodo2Al, sCadAge)
    Set oPig = Nothing
       
    If (R.EOF And R.BOF) Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = gsNomCmac
    ApExcel.Cells(3, 2).Formula = UCase(R!cAgeDescripcion) 'gsNomAge
    ApExcel.Cells(2, 15).Formula = Date + Time()
    ApExcel.Cells(3, 15).Formula = gsCodUser
    ApExcel.Range("B2", "O8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("B8", "B8").HorizontalAlignment = xlLeft
    ApExcel.Range("F2", "F3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "O6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "O10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "O10").Borders.LineStyle = 1

    ApExcel.Cells(5, 2).Formula = "CLIENTES DE PIGNORATICIOS QUE FUERON ATENDIDOS HASTA ANTES DEL " & Format(mskPeriodo2Al, "dd/MM/yyyy")
'    ApExcel.Cells(6, 2).Formula = "Información al " & Format(RFec!Ult_Fec_RCC, "dd/MM/YYYY")
'    ApExcel.Cells(8, 2).Formula = "Tipo de Cambio : " & Trim(Str(pntipCam))
    
    ApExcel.Range("B5", "F5").MergeCells = True
'    ApExcel.Range("B6", "F6").MergeCells = True

'    ApExcel.Range("B9", "C10").MergeCells = True
'    ApExcel.Range("D9", "E9").MergeCells = True
'    ApExcel.Range("F9", "G9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "CLIENTE"
    ApExcel.Cells(9, 3).Formula = "DNI"
    ApExcel.Cells(9, 4).Formula = "ULTIMO CREDITO"
    ApExcel.Cells(9, 5).Formula = "FECHA ULT. OPE."
    ApExcel.Cells(9, 6).Formula = "AGENCIA"
    ApExcel.Cells(9, 7).Formula = "CALIFICACION"
    ApExcel.Cells(9, 8).Formula = "NUM. ADJUDICADOS"
    ApExcel.Cells(9, 9).Formula = "ENTID. FINANCIERAS"
    ApExcel.Cells(9, 10).Formula = "DEPARTAMENTO"
    ApExcel.Cells(9, 11).Formula = "PROVINCIA"
    ApExcel.Cells(9, 12).Formula = "DISTRITO"
    ApExcel.Cells(9, 13).Formula = "ZONA"
    ApExcel.Cells(9, 14).Formula = "DIRECCION"
    ApExcel.Cells(9, 15).Formula = "TELEFONOS"
    
    ApExcel.Range("B9", "O10").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B9", "O10").Font.Bold = True
    ApExcel.Range("B9", "O10").HorizontalAlignment = 3

    i = 10
    
    Do While Not R.EOF
    i = i + 1
            
        lcAge = R!cAgeDescripcion
        
        ApExcel.Cells(i, 2).Formula = "AGENCIA :" & UCase(R!cAgeDescripcion)
        ApExcel.Cells(i, 2).Font.Bold = True
                   
        Do While R!cAgeDescripcion = lcAge

            i = i + 1

            ApExcel.Cells(i, 2).Formula = R!cPersNombre
            ApExcel.Cells(i, 3).Formula = R!cPersIDNro
            ApExcel.Cells(i, 4).Formula = R!cCtaCodUlt
            ApExcel.Cells(i, 5).Formula = R!UltMov
            ApExcel.Cells(i, 6).Formula = R!cAgeDescripcion
            ApExcel.Cells(i, 7).Formula = R!Calif
            ApExcel.Cells(i, 8).Formula = R!nNumAdju
            ApExcel.Cells(i, 9).Formula = R!EntiFinan
            ApExcel.Cells(i, 10).Formula = R!DPTO
            ApExcel.Cells(i, 11).Formula = R!Provincia
            ApExcel.Cells(i, 12).Formula = R!distrito
            ApExcel.Cells(i, 13).Formula = R!Zona
            ApExcel.Cells(i, 14).Formula = R!cPersDireccDomicilio
            ApExcel.Cells(i, 15).Formula = R!Telf

            'ApExcel.Range("C" & Trim(Str(i)) & ":" & "F" & Trim(Str(i))).NumberFormat = "#,##0.00"
            ApExcel.Range("B" & Trim(str(i)) & ":" & "O" & Trim(str(i))).Borders.LineStyle = 1
                                          
            R.MoveNext
        
            If R.EOF Then
                Exit Do
            End If
        Loop
        If R.EOF Then
            Exit Do
        End If
    Loop
    i = i + 2
    
'    ApExcel.Range("B" & Trim(Str(i)) & ":" & "F" & Trim(Str(i + 2))).Borders.LineStyle = 1
'    ApExcel.Range("B" & Trim(Str(i)) & ":" & "F" & Trim(Str(i + 2))).Font.Bold = True

    R.Close
    Set R = Nothing

    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 38#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing

End Sub


'*** PEAC 20090521
Private Sub ReporteDiferidosDetalladoPorLoteExcel()
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer, lcAge As String, lcPersNom As String
Dim sCadAge As String
Dim lcCtaCod As String

    For i = 0 To UBound(MatAgencias) - 1
        sCadAge = sCadAge & MatAgencias(i) & ","
        Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
       
    If Len(sCadAge) = 0 Then
        MsgBox "Por favor seleccione al menos una Agencia.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Dim oPig As COMNColoCPig.NCOMColPRecGar
    Set oPig = New COMNColoCPig.NCOMColPRecGar
        Set R = oPig.nObtieneDiferidosPorLote(sCadAge)
    Set oPig = Nothing
       
    If (R.EOF And R.BOF) Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = gsNomCmac
    ApExcel.Cells(3, 2).Formula = UCase(R!cAgeDescripcion) 'gsNomAge
    ApExcel.Cells(2, 6).Formula = Date + Time()
    ApExcel.Cells(3, 6).Formula = gsCodUser
    ApExcel.Range("B2", "F8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("B8", "B8").HorizontalAlignment = xlLeft
    ApExcel.Range("F2", "F3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "F6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "F10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "F10").Borders.LineStyle = 1

    ApExcel.Cells(5, 2).Formula = "CONTRATOS DIFERIDOS DETALLADO POR LOTE"
'    ApExcel.Cells(6, 2).Formula = "Información al " & Format(RFec!Ult_Fec_RCC, "dd/MM/YYYY")
'    ApExcel.Cells(8, 2).Formula = "Tipo de Cambio : " & Trim(Str(pntipCam))
    
    ApExcel.Range("B5", "F5").MergeCells = True
    ApExcel.Range("B6", "F6").MergeCells = True

'    ApExcel.Range("B9", "C10").MergeCells = True
'    ApExcel.Range("D9", "E9").MergeCells = True
'    ApExcel.Range("F9", "G9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "DESCRIPCION"
    ApExcel.Cells(9, 3).Formula = "PIEZAS"
    ApExcel.Cells(9, 4).Formula = "KILATAJE"
    ApExcel.Cells(9, 5).Formula = "PESO BRUTO"
    ApExcel.Cells(9, 6).Formula = "PESO NETO"
    
    ApExcel.Range("B9", "F10").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B9", "F10").Font.Bold = True
    ApExcel.Range("B9", "F10").HorizontalAlignment = 3

    i = 10
    
    Do While Not R.EOF
    i = i + 1
            
        lcAge = R!cAgeDescripcion
        
        ApExcel.Cells(i, 2).Formula = "AGENCIA :" & UCase(R!cAgeDescripcion)
        ApExcel.Cells(i, 2).Font.Bold = True
                   
        Do While R!cAgeDescripcion = lcAge

            i = i + 1
                
            lcPersNom = R!cPersNombre
            
            ApExcel.Cells(i, 2).Formula = "CUENTA : " & R!cCtacod & " CLIENTE : " & Trim(R!cPersNombre)
            ApExcel.Cells(i, 2).Font.Bold = True
            
                Do While R!cPersNombre = lcPersNom And R!cAgeDescripcion = lcAge
                    
                    i = i + 1
                
                    ApExcel.Cells(i, 2).Formula = R!cDescrip
                    ApExcel.Cells(i, 3).Formula = R!nPiezas
                    ApExcel.Cells(i, 4).Formula = R!cKilataje
                    ApExcel.Cells(i, 5).Formula = R!nPesoBruto
                    ApExcel.Cells(i, 6).Formula = R!nPesoNeto
                    
                    ApExcel.Range("C" & Trim(str(i)) & ":" & "F" & Trim(str(i))).NumberFormat = "#,##0.00"
                    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i))).Borders.LineStyle = 1
                                                  
                    R.MoveNext
                
                    If R.EOF Then
                        Exit Do
                    End If
                Loop
                If R.EOF Then
                    Exit Do
                End If
        Loop
        i = i + 1
    Loop
    
    i = i + 2
    
    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i + 2))).Borders.LineStyle = 1
    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i + 2))).Font.Bold = True

    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 63#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing

End Sub


'*** PEAC 20090521
Private Sub ReporteAdjudicadosDetalladoPorLoteExcel()
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer, lcAge As String, lcPersNom As String
Dim sCadAge As String
Dim lcCtaCod As String

    For i = 0 To UBound(MatAgencias) - 1
        sCadAge = sCadAge & MatAgencias(i) & ","
        Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
       
    If Len(sCadAge) = 0 Then
        MsgBox "Por favor seleccione al menos una Agencia.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Dim oPig As COMNColoCPig.NCOMColPRecGar
    Set oPig = New COMNColoCPig.NCOMColPRecGar
        Set R = oPig.nObtieneAdjudicadosPorLote(sCadAge, mskPeriodo1Del.Text, mskPeriodo1Al.Text)
    Set oPig = Nothing
       
    If (R.EOF And R.BOF) Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = gsNomCmac
    ApExcel.Cells(3, 2).Formula = UCase(R!cAgeDescripcion) 'gsNomAge
    ApExcel.Cells(2, 6).Formula = Date + Time()
    ApExcel.Cells(3, 6).Formula = gsCodUser
    ApExcel.Range("B2", "F8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("B8", "B8").HorizontalAlignment = xlLeft
    ApExcel.Range("F2", "F3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "F6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "F10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "F10").Borders.LineStyle = 1

    ApExcel.Cells(5, 2).Formula = "CONTRATOS ADJUDICADOS DETALLADO POR LOTE"
'    ApExcel.Cells(6, 2).Formula = "Información al " & Format(RFec!Ult_Fec_RCC, "dd/MM/YYYY")
'    ApExcel.Cells(8, 2).Formula = "Tipo de Cambio : " & Trim(Str(pntipCam))
    
    ApExcel.Range("B5", "F5").MergeCells = True
    ApExcel.Range("B6", "F6").MergeCells = True

'    ApExcel.Range("B9", "C10").MergeCells = True
'    ApExcel.Range("D9", "E9").MergeCells = True
'    ApExcel.Range("F9", "G9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "DESCRIPCION"
    ApExcel.Cells(9, 3).Formula = "PIEZAS"
    ApExcel.Cells(9, 4).Formula = "KILATAJE"
    ApExcel.Cells(9, 5).Formula = "PESO BRUTO"
    ApExcel.Cells(9, 6).Formula = "PESO NETO"
    
    ApExcel.Range("B9", "F10").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B9", "F10").Font.Bold = True
    ApExcel.Range("B9", "F10").HorizontalAlignment = 3

    i = 10
    
    Do While Not R.EOF
    i = i + 1
            
        lcAge = R!cAgeDescripcion
        
        ApExcel.Cells(i, 2).Formula = "AGENCIA :" & UCase(R!cAgeDescripcion)
        ApExcel.Cells(i, 2).Font.Bold = True
                   
        Do While R!cAgeDescripcion = lcAge

            i = i + 1
                
            lcPersNom = R!cPersNombre
            lcCtaCod = R!cCtacod '**********RECO 20131017
            ApExcel.Cells(i, 2).Formula = "CUENTA : " & R!cCtacod & " CLIENTE : " & Trim(R!cPersNombre) & " ADJUDICADO :" & Format(R!dProceso, "dd/mm/yyyy")
            ApExcel.Cells(i, 2).Font.Bold = True
            
                'Do While R!cPersNombre = lcPersNom And R!cAgeDescripcion = lcAge
                Do While R!cPersNombre = lcPersNom And R!cAgeDescripcion = lcAge And lcCtaCod = R!cCtacod '******RECO 20131017******
                    
                    i = i + 1
                
                    ApExcel.Cells(i, 2).Formula = R!cDescrip
                    ApExcel.Cells(i, 3).Formula = R!nPiezas
                    ApExcel.Cells(i, 4).Formula = R!cKilataje
                    ApExcel.Cells(i, 5).Formula = R!nPesoBruto
                    ApExcel.Cells(i, 6).Formula = R!nPesoNeto
                    
                    ApExcel.Range("C" & Trim(str(i)) & ":" & "F" & Trim(str(i))).NumberFormat = "#,##0.00"
                    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i))).Borders.LineStyle = 1
                                                  
                    R.MoveNext
                
                    If R.EOF Then
                        Exit Do
                    End If
                Loop
                If R.EOF Then
                    Exit Do
                End If
        Loop
        i = i + 1
    Loop
    
    i = i + 2
    
    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i + 2))).Borders.LineStyle = 1
    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i + 2))).Font.Bold = True

    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 63#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing

End Sub


'*** PEAC 20090521
Private Sub ReporteVigentesDetalladoPorLoteExcel()
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer, lcAge As String, lcPersNom As String
Dim sCadAge As String
Dim lcCtaCod As String

    For i = 0 To UBound(MatAgencias) - 1
        sCadAge = sCadAge & MatAgencias(i) & ","
        Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
       
    If Len(sCadAge) = 0 Then
        MsgBox "Por favor seleccione al menos una Agencia.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Dim oPig As COMNColoCPig.NCOMColPRecGar
    Set oPig = New COMNColoCPig.NCOMColPRecGar
        Set R = oPig.nObtieneVigentesPorLote(sCadAge, mskPeriodo1Del.Text, mskPeriodo1Al.Text)
    Set oPig = Nothing

    If (R.EOF And R.BOF) Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = gsNomCmac
    ApExcel.Cells(3, 2).Formula = UCase(R!CAgencia) 'gsNomAge
    ApExcel.Cells(2, 6).Formula = Date + Time()
    ApExcel.Cells(3, 6).Formula = gsCodUser
    ApExcel.Range("B2", "F8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("B8", "B8").HorizontalAlignment = xlLeft
    ApExcel.Range("F2", "F3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "F6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "F10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "F10").Borders.LineStyle = 1

    ApExcel.Cells(5, 2).Formula = "CONTRATOS VIGENTES DETALLADO POR LOTE"
'    ApExcel.Cells(6, 2).Formula = "Información al " & Format(RFec!Ult_Fec_RCC, "dd/MM/YYYY")
'    ApExcel.Cells(8, 2).Formula = "Tipo de Cambio : " & Trim(Str(pntipCam))
    
    ApExcel.Range("B5", "F5").MergeCells = True
    ApExcel.Range("B6", "F6").MergeCells = True

'    ApExcel.Range("B9", "C10").MergeCells = True
'    ApExcel.Range("D9", "E9").MergeCells = True
'    ApExcel.Range("F9", "G9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "DESCRIPCION"
    ApExcel.Cells(9, 3).Formula = "PIEZAS"
    ApExcel.Cells(9, 4).Formula = "KILATAJE"
    ApExcel.Cells(9, 5).Formula = "PESO BRUTO"
    ApExcel.Cells(9, 6).Formula = "PESO NETO"
    
    ApExcel.Range("B9", "F10").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B9", "F10").Font.Bold = True
    ApExcel.Range("B9", "F10").HorizontalAlignment = 3

    i = 10
    
    Do While Not R.EOF
    i = i + 1
            
        lcAge = R!CAgencia
        
        ApExcel.Cells(i, 2).Formula = "AGENCIA :" & UCase(R!CAgencia)
        ApExcel.Cells(i, 2).Font.Bold = True
                   
        Do While R!CAgencia = lcAge

            i = i + 1
                
            lcPersNom = R!cPersNombre
            lcCtaCod = R!cCtacod 'EAAS 20170908
            ApExcel.Cells(i, 2).Formula = "CUENTA : " & R!cCtacod & " CLIENTE :" & Trim(R!cPersNombre) & " DESEMBOLSO : " & Format(R!dVigencia, "dd/mm/yyyy") & " CONDICION : " & Trim(R!Condicion)
            ApExcel.Cells(i, 2).Font.Bold = True
            
                Do While R!cPersNombre = lcPersNom And R!CAgencia = lcAge And lcCtaCod = R!cCtacod 'EAAS 20170908
                    
                    i = i + 1
                
                    ApExcel.Cells(i, 2).Formula = R!cDescrip
                    ApExcel.Cells(i, 3).Formula = R!nPiezas
                    ApExcel.Cells(i, 4).Formula = R!cKilataje
                    ApExcel.Cells(i, 5).Formula = R!nPesoBruto
                    ApExcel.Cells(i, 6).Formula = R!nPesoNeto
                    
                    ApExcel.Range("C" & Trim(str(i)) & ":" & "F" & Trim(str(i))).NumberFormat = "#,##0.00"
                    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i))).Borders.LineStyle = 1
                                                  
                    R.MoveNext
                
                    If R.EOF Then
                        Exit Do
                    End If
                Loop
                If R.EOF Then
                    Exit Do
                End If
        Loop
        i = i + 1
    Loop
    
    i = i + 2
    
    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i + 2))).Borders.LineStyle = 1
    ApExcel.Range("B" & Trim(str(i)) & ":" & "F" & Trim(str(i + 2))).Font.Bold = True

    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 63#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing

End Sub


