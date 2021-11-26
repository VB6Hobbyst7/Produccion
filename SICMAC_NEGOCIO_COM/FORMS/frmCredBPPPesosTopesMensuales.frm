VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredBPPPesosTopesMensuales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Pesos y Topes Mensuales"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   Icon            =   "frmCredBPPPesosTopesMensuales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Pesos y Topes"
      TabPicture(0)   =   "frmCredBPPPesosTopesMensuales.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSTab2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "uspAnio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdMostrar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbAgencias"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbMeses"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.ComboBox cmbMeses 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   810
         Width           =   1815
      End
      Begin VB.ComboBox cmbAgencias 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   810
         Width           =   2175
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   7440
         TabIndex        =   1
         Top             =   780
         Width           =   1095
      End
      Begin Spinner.uSpinner uspAnio 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   810
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Max             =   9999
         Min             =   1900
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5741
         _Version        =   393216
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "Nivel I"
         TabPicture(0)   =   "frmCredBPPPesosTopesMensuales.frx":0326
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Shape1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Shape2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label5"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Shape3"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label6"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Shape4"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label7"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdGuardarNivelI"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdCancelarNivelI"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Nivel II"
         TabPicture(1)   =   "frmCredBPPPesosTopesMensuales.frx":0342
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Shape5"
         Tab(1).Control(1)=   "Label8"
         Tab(1).Control(2)=   "Shape6"
         Tab(1).Control(3)=   "Label9"
         Tab(1).Control(4)=   "Shape7"
         Tab(1).Control(5)=   "Label10"
         Tab(1).Control(6)=   "Shape8"
         Tab(1).Control(7)=   "Label11"
         Tab(1).Control(8)=   "cmdCancelarNivelII"
         Tab(1).Control(9)=   "cmdGuardarNivelII"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Nivel III"
         TabPicture(2)   =   "frmCredBPPPesosTopesMensuales.frx":035E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Shape9"
         Tab(2).Control(1)=   "Label12"
         Tab(2).Control(2)=   "Shape10"
         Tab(2).Control(3)=   "Label13"
         Tab(2).Control(4)=   "Shape11"
         Tab(2).Control(5)=   "Label14"
         Tab(2).Control(6)=   "Shape12"
         Tab(2).Control(7)=   "Label15"
         Tab(2).Control(8)=   "cmdCancelarNivelIII"
         Tab(2).Control(9)=   "cmdGuardarNivelIII"
         Tab(2).ControlCount=   10
         Begin VB.CommandButton cmdCancelarNivelI 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8880
            TabIndex        =   14
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdGuardarNivelI 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   7680
            TabIndex        =   13
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdGuardarNivelII 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   -67320
            TabIndex        =   12
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelarNivelII 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   -66120
            TabIndex        =   11
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdGuardarNivelIII 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   -67320
            TabIndex        =   10
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelarNivelIII 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   -66120
            TabIndex        =   9
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -66060
            TabIndex        =   15
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape12 
            Height          =   285
            Left            =   -66840
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -67740
            TabIndex        =   20
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape11 
            Height          =   285
            Left            =   -68550
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -69435
            TabIndex        =   25
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape10 
            Height          =   285
            Left            =   -70260
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -71160
            TabIndex        =   26
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape9 
            Height          =   285
            Left            =   -71970
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -66060
            TabIndex        =   24
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape8 
            Height          =   285
            Left            =   -66840
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -67740
            TabIndex        =   23
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape7 
            Height          =   285
            Left            =   -68550
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -69435
            TabIndex        =   22
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape6 
            Height          =   285
            Left            =   -70260
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -71160
            TabIndex        =   21
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape5 
            Height          =   285
            Left            =   -71970
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8940
            TabIndex        =   19
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape4 
            Height          =   285
            Left            =   8160
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7260
            TabIndex        =   18
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape3 
            Height          =   285
            Left            =   6450
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5565
            TabIndex        =   17
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape2 
            Height          =   285
            Left            =   4740
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   16
            Top             =   765
            Width           =   135
         End
         Begin VB.Shape Shape1 
            Height          =   285
            Left            =   3030
            Top             =   720
            Width           =   1725
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Filtro de Registro"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes - Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   870
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         Height          =   195
         Left            =   4320
         TabIndex        =   5
         Top             =   870
         Width           =   675
      End
   End
   Begin VB.Line Line1 
      X1              =   4920
      X2              =   6120
      Y1              =   2400
      Y2              =   2880
   End
End
Attribute VB_Name = "frmCredBPPPesosTopesMensuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim lnNivel As Integer
'Dim lnMes As Integer
'Dim lnAnio As Integer
'Dim lsCodAge As String
'Dim lsFecha As String
'Dim lnFila As Integer
'Dim lsCodFact As String
'Dim lnTopeCatA As Double
'Dim lnPesoCatA As Double
'Dim lnTopeCatB As Double
'Dim lnPesoCatB As Double
'Dim lnTopeCatC As Double
'Dim lnPesoCatC As Double
'Dim lnTopeCatD As Double
'Dim lnPesoCatD As Double
'Dim i As Integer, j As Integer
'
'Private Sub cmdCancelarNivelI_Click()
'    LimpiaFlex flxNivelI
'    CargaFactoresNivelI
'End Sub
'
'Private Sub cmdCancelarNivelII_Click()
'    LimpiaFlex flxNivelII
'    CargaFactoresNivelII
'End Sub
'
'Private Sub cmdCancelarNivelIII_Click()
'    LimpiaFlex flxNivelIII
'    CargaFactoresNivelIII
'End Sub
'
'Private Sub cmdGuardarNivelI_Click()
'On Error GoTo Error
'If ValidaDatos(1) Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim oBPP As COMNCredito.NCOMBPPR
'
'    lnNivel = 1
'    lnMes = Right(cmbMeses.Text, 2)
'    lnAnio = uspAnio.valor
'    lsCodAge = Right(cmbAgencias.Text, 2)
'    lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'    lnFila = flxNivelI.Rows - 1
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'    Call oBPP.EliminaPesoyTopes(lnNivel, lnMes, lnAnio, lsCodAge)
'    For i = 1 To lnFila
'        lsCodFact = flxNivelI.TextMatrix(i, 10)
'        lnTopeCatA = flxNivelI.TextMatrix(i, 9)
'        lnPesoCatA = flxNivelI.TextMatrix(i, 8)
'        lnTopeCatB = flxNivelI.TextMatrix(i, 7)
'        lnPesoCatB = flxNivelI.TextMatrix(i, 6)
'        lnTopeCatC = flxNivelI.TextMatrix(i, 5)
'        lnPesoCatC = flxNivelI.TextMatrix(i, 4)
'        lnTopeCatD = flxNivelI.TextMatrix(i, 3)
'        lnPesoCatD = flxNivelI.TextMatrix(i, 2)
'
'        oBPP.InsertaPesoyTopes lnNivel, lsCodFact, lnPesoCatA, lnTopeCatA, lnPesoCatB, lnTopeCatB, lnPesoCatC, lnTopeCatC, lnPesoCatD, lnTopeCatD, _
'                               lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha
'
'    Next
'    MsgBox "Datos Ingresados Correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarNivelII_Click()
'On Error GoTo Error
'If ValidaDatos(2) Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim oBPP As COMNCredito.NCOMBPPR
'
'        lnNivel = 2
'        lnMes = Right(cmbMeses.Text, 2)
'        lnAnio = uspAnio.valor
'        lsCodAge = Right(cmbAgencias.Text, 2)
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        lnFila = flxNivelII.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        Call oBPP.EliminaPesoyTopes(lnNivel, lnMes, lnAnio, lsCodAge)
'        For i = 1 To lnFila
'            lsCodFact = flxNivelII.TextMatrix(i, 10)
'            lnTopeCatA = flxNivelII.TextMatrix(i, 9)
'            lnPesoCatA = flxNivelII.TextMatrix(i, 8)
'            lnTopeCatB = flxNivelII.TextMatrix(i, 7)
'            lnPesoCatB = flxNivelII.TextMatrix(i, 6)
'            lnTopeCatC = flxNivelII.TextMatrix(i, 5)
'            lnPesoCatC = flxNivelII.TextMatrix(i, 4)
'            lnTopeCatD = flxNivelII.TextMatrix(i, 3)
'            lnPesoCatD = flxNivelII.TextMatrix(i, 2)
'
'            oBPP.InsertaPesoyTopes lnNivel, lsCodFact, lnPesoCatA, lnTopeCatA, lnPesoCatB, lnTopeCatB, lnPesoCatC, lnTopeCatC, lnPesoCatD, lnTopeCatD, _
'                                   lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha
'
'        Next
'        MsgBox "Datos Ingresados Correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarNivelIII_Click()
'On Error GoTo Error
'If ValidaDatos(3) Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim oBPP As COMNCredito.NCOMBPPR
'
'    lnNivel = 3
'    lnMes = Right(cmbMeses.Text, 2)
'    lnAnio = uspAnio.valor
'    lsCodAge = Right(cmbAgencias.Text, 2)
'    lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'    lnFila = flxNivelIII.Rows - 1
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'    Call oBPP.EliminaPesoyTopes(lnNivel, lnMes, lnAnio, lsCodAge)
'
'    For i = 1 To lnFila
'        lsCodFact = flxNivelIII.TextMatrix(i, 10)
'        lnTopeCatA = flxNivelIII.TextMatrix(i, 9)
'        lnPesoCatA = flxNivelIII.TextMatrix(i, 8)
'        lnTopeCatB = flxNivelIII.TextMatrix(i, 7)
'        lnPesoCatB = flxNivelIII.TextMatrix(i, 6)
'        lnTopeCatC = flxNivelIII.TextMatrix(i, 5)
'        lnPesoCatC = flxNivelIII.TextMatrix(i, 4)
'        lnTopeCatD = flxNivelIII.TextMatrix(i, 3)
'        lnPesoCatD = flxNivelIII.TextMatrix(i, 2)
'
'        oBPP.InsertaPesoyTopes lnNivel, lsCodFact, lnPesoCatA, lnTopeCatA, lnPesoCatB, lnTopeCatB, lnPesoCatC, lnTopeCatC, lnPesoCatD, lnTopeCatD, _
'                               lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha
'
'    Next
'    MsgBox "Datos Ingresados Correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdMostrar_Click()
'If ValidaDatos Then
'    CargaGridNivelI 1, Right(cmbMeses.Text, 2), uspAnio.valor, Right(cmbAgencias.Text, 2)
'    CargaGridNivelII 2, Right(cmbMeses.Text, 2), uspAnio.valor, Right(cmbAgencias.Text, 2)
'    CargaGridNivelIII 3, Right(cmbMeses.Text, 2), uspAnio.valor, Right(cmbAgencias.Text, 2)
'End If
'End Sub
'
'Private Sub Form_Load()
'    CargaCombos
'    uspAnio.valor = Year(gdFecSis)
'End Sub
'
'Private Sub CargaCombos()
'    CargaComboMeses cmbMeses
'    CargaComboAgencias cmbAgencias
'    CargaFactoresNivelI
'    CargaFactoresNivelII
'    CargaFactoresNivelIII
'End Sub
'
'Private Sub CargaFactoresNivelI()
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverFactoresCumplimiento(1, 1)
'
'    i = 1
'
'    Do While Not rs.EOF
'        flxNivelI.AdicionaFila
'
'        flxNivelI.TextMatrix(i, 1) = rs!cDescFactor
'        flxNivelI.TextMatrix(i, 10) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'
'    Set rs = Nothing
'
'End Sub
'
'Private Sub CargaFactoresNivelII()
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverFactoresCumplimiento(2, 1)
'
'    i = 1
'
'    Do While Not rs.EOF
'        flxNivelII.AdicionaFila
'
'        flxNivelII.TextMatrix(i, 1) = rs!cDescFactor
'        flxNivelII.TextMatrix(i, 10) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'
'    Set rs = Nothing
'End Sub
'
'Private Sub CargaFactoresNivelIII()
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverFactoresCumplimiento(3, 1)
'    i = 1
'
'    Do While Not rs.EOF
'        flxNivelIII.AdicionaFila
'
'        flxNivelIII.TextMatrix(i, 1) = rs!cDescFactor
'        flxNivelIII.TextMatrix(i, 10) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'End Sub
'
'Private Sub CargaGridNivelI(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverPesoyTopes(pnNivel, pnMes, pnAnio, psCodAge)
'
'    LimpiaFlex flxNivelI
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            flxNivelI.AdicionaFila
'
'            flxNivelI.TextMatrix(i, 1) = rs!cDescFactor
'            flxNivelI.TextMatrix(i, 2) = Format(rs!nPesoCatD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 3) = Format(rs!nTopesCatD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 4) = Format(rs!nPesoCatC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 5) = Format(rs!nTopesCatC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 6) = Format(rs!nPesoCatB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 7) = Format(rs!nTopesCatB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 8) = Format(rs!nPesoCatA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 9) = Format(rs!nTopesCatA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelI.TextMatrix(i, 10) = rs!cCodFactor
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    Else
'        CargaFactoresNivelI
'    End If
'
'    Set rs = Nothing
'End Sub
'
'Private Sub CargaGridNivelII(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverPesoyTopes(pnNivel, pnMes, pnAnio, psCodAge)
'
'    LimpiaFlex flxNivelII
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            flxNivelII.AdicionaFila
'
'            flxNivelII.TextMatrix(i, 1) = rs!cDescFactor
'            flxNivelII.TextMatrix(i, 2) = Format(rs!nPesoCatD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 3) = Format(rs!nTopesCatD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 4) = Format(rs!nPesoCatC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 5) = Format(rs!nTopesCatC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 6) = Format(rs!nPesoCatB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 7) = Format(rs!nTopesCatB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 8) = Format(rs!nPesoCatA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 9) = Format(rs!nTopesCatA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelII.TextMatrix(i, 10) = rs!cCodFactor
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    Else
'        CargaFactoresNivelII
'    End If
'
'    Set rs = Nothing
'End Sub
'
'Private Sub CargaGridNivelIII(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverPesoyTopes(pnNivel, pnMes, pnAnio, psCodAge)
'
'    LimpiaFlex flxNivelIII
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            flxNivelIII.AdicionaFila
'
'            flxNivelIII.TextMatrix(i, 1) = rs!cDescFactor
'            flxNivelIII.TextMatrix(i, 2) = Format(rs!nPesoCatD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 3) = Format(rs!nTopesCatD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 4) = Format(rs!nPesoCatC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 5) = Format(rs!nTopesCatC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 6) = Format(rs!nPesoCatB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 7) = Format(rs!nTopesCatB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 8) = Format(rs!nPesoCatA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 9) = Format(rs!nTopesCatA, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxNivelIII.TextMatrix(i, 10) = rs!cCodFactor
'
'            rs.MoveNext
'            i = i + 1
'        Loop
'    Else
'        CargaFactoresNivelIII
'    End If
'
'    Set rs = Nothing
'End Sub
'
'
'Private Function ValidaDatos(Optional ByVal pnNivel As Integer = 4) As Boolean
' If Trim(cmbMeses.Text) = "" Then
'    MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'
'If Trim(uspAnio.valor) = "" Or CDbl(uspAnio.valor) = 0 Then
'    MsgBox "Ingrese el Año", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'
'If Trim(cmbAgencias.Text) = "" Then
'    MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'
'If pnNivel = 1 Then
'    For i = 0 To flxNivelI.Rows - 2
'        For j = 2 To 9
'            If Trim(flxNivelI.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If IsNumeric(Trim(flxNivelI.TextMatrix(i + 1, j))) Then
'                If CDbl(Trim(flxNivelI.TextMatrix(i + 1, j))) < 0 Then
'                    MsgBox "Ingrese correctamente los valores del Factor ''" & Trim(flxNivelI.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'        Next j
'    Next i
'ElseIf pnNivel = 2 Then
'    For i = 0 To flxNivelII.Rows - 2
'         For j = 2 To 9
'            If Trim(flxNivelII.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If IsNumeric(Trim(flxNivelII.TextMatrix(i + 1, j))) Then
'                If CDbl(Trim(flxNivelII.TextMatrix(i + 1, j))) < 0 Then
'                    MsgBox "Ingrese correctamente los valores del Factor ''" & Trim(flxNivelII.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'        Next j
'    Next i
'ElseIf pnNivel = 3 Then
'    For i = 0 To flxNivelIII.Rows - 2
'        For j = 2 To 9
'            If Trim(flxNivelIII.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If IsNumeric(Trim(flxNivelIII.TextMatrix(i + 1, j))) Then
'                If CDbl(Trim(flxNivelIII.TextMatrix(i + 1, j))) < 0 Then
'                    MsgBox "Ingrese correctamente los valores del Factor ''" & Trim(flxNivelIII.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'
'        Next j
'    Next i
'End If
'
'ValidaDatos = True
'End Function
