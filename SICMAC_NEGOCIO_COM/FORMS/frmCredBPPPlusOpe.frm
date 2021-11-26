VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPPlusOpe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Parámetros Plus y % Operaciones"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   Icon            =   "frmCredBPPPlusOpe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Factor cálculo Plus y % Operaciones"
      TabPicture(0)   =   "frmCredBPPPlusOpe.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPorcQuinc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraBonoPlus"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fraBonoPlus 
         Caption         =   "Bono Plus"
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
         Height          =   4335
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   10095
         Begin TabDlg.SSTab SSTab2 
            Height          =   3375
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   5953
            _Version        =   393216
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Nivel I"
            TabPicture(0)   =   "frmCredBPPPlusOpe.frx":0326
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "cmdGuardarN1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdCancelarN1"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Nivel II"
            TabPicture(1)   =   "frmCredBPPPlusOpe.frx":0342
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cmdGuardarN2"
            Tab(1).Control(1)=   "cmdCancelarN2"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Nivel III"
            TabPicture(2)   =   "frmCredBPPPlusOpe.frx":035E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "cmdGuardarN3"
            Tab(2).Control(1)=   "cmdCancelarN3"
            Tab(2).ControlCount=   2
            Begin VB.CommandButton cmdCancelarN3 
               Caption         =   "Cancelar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -66480
               TabIndex        =   9
               Top             =   2760
               Width           =   1170
            End
            Begin VB.CommandButton cmdGuardarN3 
               Caption         =   "Guardar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -67800
               TabIndex        =   14
               Top             =   2760
               Width           =   1170
            End
            Begin VB.CommandButton cmdCancelarN2 
               Caption         =   "Cancelar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -66480
               TabIndex        =   18
               Top             =   2760
               Width           =   1170
            End
            Begin VB.CommandButton cmdGuardarN2 
               Caption         =   "Guardar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -67800
               TabIndex        =   23
               Top             =   2760
               Width           =   1170
            End
            Begin VB.CommandButton cmdCancelarN1 
               Caption         =   "Cancelar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   8520
               TabIndex        =   22
               Top             =   2760
               Width           =   1170
            End
            Begin VB.CommandButton cmdGuardarN1 
               Caption         =   "Guardar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7200
               TabIndex        =   21
               Top             =   2760
               Width           =   1170
            End
         End
         Begin VB.CommandButton cmdMostrarBP 
            Caption         =   "Mostrar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   16
            Top             =   320
            Width           =   1170
         End
         Begin VB.ComboBox cmbAgencias 
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   360
            Width           =   3015
         End
         Begin VB.ComboBox cmbMesAnioBP 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Mes - Año:"
            Height          =   195
            Left            =   360
            TabIndex        =   7
            Top             =   400
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
            Height          =   195
            Left            =   4920
            TabIndex        =   6
            Top             =   400
            Width           =   630
         End
      End
      Begin VB.Frame fraPorcQuinc 
         Caption         =   "Porcentaje quincenal Factor Operaciones"
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
         Height          =   1335
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   10095
         Begin VB.ComboBox cmbAgenciasFO 
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdGuardarFO 
            Caption         =   "Guardar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   12
            Top             =   840
            Width           =   1170
         End
         Begin VB.CommandButton cmdCancelarFO 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   11
            Top             =   840
            Width           =   1170
         End
         Begin VB.CommandButton cmdMostrarFO 
            Caption         =   "Mostrar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   10
            Top             =   300
            Width           =   1170
         End
         Begin VB.ComboBox cmbMesAnioFO 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   9720
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agencia :"
            Height          =   195
            Left            =   4800
            TabIndex        =   20
            Top             =   390
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes - Año:"
            Height          =   195
            Left            =   360
            TabIndex        =   4
            Top             =   390
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%2º Quincena:"
            Height          =   195
            Left            =   2160
            TabIndex        =   3
            Top             =   855
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%1º Quincena:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   860
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPPlusOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim i As Integer, j As Integer
'Dim lnNivel As Integer
'Dim lnMes As Integer
'Dim lnAnio As Integer
'Dim lsCodAge As String
'Dim lsCodFact As String
'Dim lnCatA As Double
'Dim lnCatB As Double
'Dim lnCatC As Double
'Dim lnCatD As Double
'Dim lsFecha As String
'Dim lnFila As Integer
'
'Private Sub cmdCancelarFO_Click()
'txtQuincena1.Text = 0
'txtQuincena2.Text = 0
'End Sub
'
'Private Sub cmdCancelarN1_Click()
'    LimpiaFlex feFactorN1
'    CargaFactoresNivelI
'End Sub
'
'Private Sub cmdCancelarN2_Click()
'    LimpiaFlex feFactorN2
'    CargaFactoresNivelI
'End Sub
'
'Private Sub cmdCancelarN3_Click()
'    LimpiaFlex feFactorN3
'    CargaFactoresNivelI
'End Sub
'
'Private Sub cmdGuardarFO_Click()
'On Error GoTo Error
'If ValidaDatosOpe(1) Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Dim lsFecha As String
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'    lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'    Call oBPP.EliminaFactorOpe(CInt(Right(cmbMesAnioFO.Text, 2)), spAnioFO.valor, Trim(Right(cmbAgenciasFO.Text, 4)))
'    Call oBPP.InsertaFactorOpe(CInt(Right(cmbMesAnioFO.Text, 2)), spAnioFO.valor, CDbl(txtQuincena1.Text), CDbl(txtQuincena2.Text), gsCodUser, lsFecha, Trim(Right(cmbAgenciasFO.Text, 4)))
'
'    MsgBox "Se registraron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarN1_Click()
'On Error GoTo Error
'If ValidaDatos(1) Then
'    Dim oBPP As COMNCredito.NCOMBPPR
'        If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'        lnNivel = 1
'        lnMes = Right(cmbMesAnioBP.Text, 2)
'        lnAnio = spAnioBP.valor
'        lsCodAge = Right(cmbAgencias.Text, 2)
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        lnFila = feFactorN1.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'
'        Call oBPP.EliminaParametrosCumplimiento(lnNivel, lnMes, lnAnio, lsCodAge, 3)
'
'        For i = 1 To lnFila
'            lsCodFact = feFactorN1.TextMatrix(i, 2)
'            lnCatA = feFactorN1.TextMatrix(i, 6)
'            lnCatB = feFactorN1.TextMatrix(i, 5)
'            lnCatC = feFactorN1.TextMatrix(i, 4)
'            lnCatD = feFactorN1.TextMatrix(i, 3)
'
'            oBPP.InsertaParametroCumplimiento lnNivel, lsCodFact, lnCatA, lnCatB, lnCatC, lnCatD, lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha, 3
'
'        Next
'        MsgBox "Se guardaron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarN2_Click()
'On Error GoTo Error
'If ValidaDatos(2) Then
'    Dim oBPP As COMNCredito.NCOMBPPR
'        If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'        lnNivel = 2
'        lnMes = Right(cmbMesAnioBP.Text, 2)
'        lnAnio = spAnioBP.valor
'        lsCodAge = Right(cmbAgencias.Text, 2)
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        lnFila = feFactorN2.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        Call oBPP.EliminaParametrosCumplimiento(lnNivel, lnMes, lnAnio, lsCodAge, 3)
'        For i = 1 To lnFila
'            lsCodFact = feFactorN2.TextMatrix(i, 2)
'            lnCatA = feFactorN2.TextMatrix(i, 6)
'            lnCatB = feFactorN2.TextMatrix(i, 5)
'            lnCatC = feFactorN2.TextMatrix(i, 4)
'            lnCatD = feFactorN2.TextMatrix(i, 3)
'
'            oBPP.InsertaParametroCumplimiento lnNivel, lsCodFact, lnCatA, lnCatB, lnCatC, lnCatD, lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha, 3
'
'        Next
'        MsgBox "Se guardaron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarN3_Click()
'On Error GoTo Error
'If ValidaDatos(3) Then
'    Dim oBPP As COMNCredito.NCOMBPPR
'        If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'
'        lnNivel = 3
'        lnMes = Right(cmbMesAnioBP.Text, 2)
'        lnAnio = spAnioBP.valor
'        lsCodAge = Right(cmbAgencias.Text, 2)
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        lnFila = feFactorN3.Rows - 1
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        Call oBPP.EliminaParametrosCumplimiento(lnNivel, lnMes, lnAnio, lsCodAge, 3)
'        For i = 1 To lnFila
'            lsCodFact = feFactorN3.TextMatrix(i, 2)
'            lnCatA = feFactorN3.TextMatrix(i, 6)
'            lnCatB = feFactorN3.TextMatrix(i, 5)
'            lnCatC = feFactorN3.TextMatrix(i, 4)
'            lnCatD = feFactorN3.TextMatrix(i, 3)
'
'            oBPP.InsertaParametroCumplimiento lnNivel, lsCodFact, lnCatA, lnCatB, lnCatC, lnCatD, lnMes, lnAnio, lsCodAge, gsCodUser, lsFecha, 3
'
'        Next
'        MsgBox "Se guardaron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdMostrarBP_Click()
'If ValidaDatos Then
'    CargaGridNivelI 1, Right(cmbMesAnioBP.Text, 2), spAnioBP.valor, Right(cmbAgencias.Text, 2)
'    CargaGridNivelII 2, Right(cmbMesAnioBP.Text, 2), spAnioBP.valor, Right(cmbAgencias.Text, 2)
'    CargaGridNivelIII 3, Right(cmbMesAnioBP.Text, 2), spAnioBP.valor, Right(cmbAgencias.Text, 2)
'End If
'End Sub
'Private Sub CargaGridNivelI(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverParametrosCumplimiento(pnNivel, pnMes, pnAnio, psCodAge, 3)
'
'    LimpiaFlex feFactorN1
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            feFactorN1.AdicionaFila
'
'            feFactorN1.TextMatrix(i, 1) = rs!cDescFactor
'            feFactorN1.TextMatrix(i, 2) = rs!cCodFactor
'            feFactorN1.TextMatrix(i, 3) = Format(rs!nCategoriaD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN1.TextMatrix(i, 4) = Format(rs!nCategoriaC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN1.TextMatrix(i, 5) = Format(rs!nCategoriaB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN1.TextMatrix(i, 6) = Format(rs!nCategoriaA, "###," & String(15, "#") & "#0." & String(2, "0"))
'
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
'Private Sub CargaGridNivelII(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverParametrosCumplimiento(pnNivel, pnMes, pnAnio, psCodAge, 3)
'
'    LimpiaFlex feFactorN2
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            feFactorN2.AdicionaFila
'
'            feFactorN2.TextMatrix(i, 1) = rs!cDescFactor
'            feFactorN2.TextMatrix(i, 2) = rs!cCodFactor
'            feFactorN2.TextMatrix(i, 3) = Format(rs!nCategoriaD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN2.TextMatrix(i, 4) = Format(rs!nCategoriaC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN2.TextMatrix(i, 5) = Format(rs!nCategoriaB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN2.TextMatrix(i, 6) = Format(rs!nCategoriaA, "###," & String(15, "#") & "#0." & String(2, "0"))
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
'Private Sub CargaGridNivelIII(ByVal pnNivel As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim i As Integer
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverParametrosCumplimiento(pnNivel, pnMes, pnAnio, psCodAge, 3)
'
'    LimpiaFlex feFactorN3
'
'    If rs.RecordCount > 0 Then
'        i = 1
'
'        Do While Not rs.EOF
'            feFactorN3.AdicionaFila
'
'            feFactorN3.TextMatrix(i, 1) = rs!cDescFactor
'            feFactorN3.TextMatrix(i, 2) = rs!cCodFactor
'            feFactorN3.TextMatrix(i, 3) = Format(rs!nCategoriaD, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN3.TextMatrix(i, 4) = Format(rs!nCategoriaC, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN3.TextMatrix(i, 5) = Format(rs!nCategoriaB, "###," & String(15, "#") & "#0." & String(2, "0"))
'            feFactorN3.TextMatrix(i, 6) = Format(rs!nCategoriaA, "###," & String(15, "#") & "#0." & String(2, "0"))
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
'Private Sub cmdMostrarFO_Click()
'If ValidaDatosOpe Then
'    Dim rs As ADODB.Recordset
'    Dim oBPP As COMNCredito.NCOMBPPR
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    Set rs = oBPP.DevolverFactorOpe(CInt(Right(cmbMesAnioFO.Text, 2)), spAnioFO.valor, Trim(Right(cmbAgenciasFO.Text, 4)))
'
'    If Not (rs.EOF And rs.BOF) Then
'        txtQuincena1.Text = Format(rs!nQuincena1, "###," & String(15, "#") & "#0." & String(2, "0"))
'        txtQuincena2.Text = Format(rs!nQuincena2, "###," & String(15, "#") & "#0." & String(2, "0"))
'    Else
'        txtQuincena1.Text = 0
'        txtQuincena2.Text = 0
'    End If
'
'End If
'End Sub
'
'Private Sub Form_Load()
'CargaControles
'End Sub
'Private Sub CargaControles()
'CargaComboMeses cmbMesAnioFO
'CargaComboMeses cmbMesAnioBP
'CargaComboAgencias cmbAgencias
'CargaComboAgencias cmbAgenciasFO
'
'spAnioFO.valor = Year(gdFecSis)
'spAnioBP.valor = Year(gdFecSis)
'
'CargaFactoresNivelI
'CargaFactoresNivelII
'CargaFactoresNivelIII
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
'        feFactorN1.AdicionaFila
'
'        feFactorN1.TextMatrix(i, 1) = rs!cDescFactor
'        feFactorN1.TextMatrix(i, 2) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'
'    Set rs = Nothing
'
'End Sub
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
'        feFactorN2.AdicionaFila
'        feFactorN2.TextMatrix(i, 1) = rs!cDescFactor
'        feFactorN2.TextMatrix(i, 2) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'
'    Set rs = Nothing
'
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
'
'    i = 1
'
'    Do While Not rs.EOF
'        feFactorN3.AdicionaFila
'        feFactorN3.TextMatrix(i, 1) = rs!cDescFactor
'        feFactorN3.TextMatrix(i, 2) = rs!cCodFactor
'
'        rs.MoveNext
'        i = i + 1
'    Loop
'
'    Set rs = Nothing
'
'End Sub
'
'Private Function ValidaDatos(Optional ByVal pnNivel As Integer = 4)
'If Trim(cmbMesAnioBP.Text) = "" Then
'    MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'
'If Trim(spAnioBP.valor) = "" Or CDbl(spAnioBP.valor) = 0 Then
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
'    For i = 0 To feFactorN1.Rows - 2
'        For j = 3 To 6
'            If Trim(feFactorN1.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If IsNumeric(Trim(feFactorN1.TextMatrix(i + 1, j))) Then
'                If CDbl(Trim(feFactorN1.TextMatrix(i + 1, j))) < 0 Or CDbl(Trim(feFactorN1.TextMatrix(i + 1, j))) > 100 Then
'                    MsgBox "Ingrese correctamente los valores (0.00% - 100.00%) del Factor ''" & Trim(feFactorN1.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'        Next j
'    Next i
'ElseIf pnNivel = 2 Then
'    For i = 0 To feFactorN2.Rows - 2
'        For j = 3 To 6
'            If Trim(feFactorN2.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If IsNumeric(Trim(feFactorN2.TextMatrix(i + 1, j))) Then
'                If CDbl(Trim(feFactorN2.TextMatrix(i + 1, j))) < 0 Or CDbl(Trim(feFactorN2.TextMatrix(i + 1, j))) > 100 Then
'                    MsgBox "Ingrese correctamente los valores (0.00% - 100.00%) del Factor ''" & Trim(feFactorN2.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'        Next j
'    Next i
'ElseIf pnNivel = 3 Then
'    For i = 0 To feFactorN3.Rows - 2
'        For j = 3 To 6
'            If Trim(feFactorN3.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese correctamente los parámetros", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If IsNumeric(Trim(feFactorN3.TextMatrix(i + 1, j))) Then
'                If CDbl(Trim(feFactorN3.TextMatrix(i + 1, j))) < 0 Or CDbl(Trim(feFactorN3.TextMatrix(i + 1, j))) > 100 Then
'                    MsgBox "Ingrese correctamente los valores (0.00% - 100.00%) del Factor ''" & Trim(feFactorN3.TextMatrix(i + 1, 1)) & "''", vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'        Next j
'    Next i
'End If
'
'ValidaDatos = True
'End Function
'
'Private Function ValidaDatosOpe(Optional ByVal pnTipo As Integer = 0)
'    If Trim(cmbMesAnioFO.Text) = "" Then
'        MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'        ValidaDatosOpe = False
'        Exit Function
'    End If
'
'    If Trim(spAnioFO.valor) = "" Or CDbl(spAnioBP.valor) = 0 Then
'        MsgBox "Ingrese el Año", vbInformation, "Aviso"
'        ValidaDatosOpe = False
'        Exit Function
'    End If
'
'    If Trim(cmbAgenciasFO.Text) = "" Then
'        MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'        ValidaDatosOpe = False
'        Exit Function
'    End If
'If pnTipo = 1 Then
'    If txtQuincena1.Text = "" Or txtQuincena2.Text = "" Then
'        MsgBox "Ingrese los Datos Completos", vbInformation, "Aviso"
'        ValidaDatosOpe = False
'        Exit Function
'    End If
'
'    If (CCur(txtQuincena1.Text) + CCur(txtQuincena2.Text)) <> 100 Then
'        MsgBox "Las suma de las 2 quincenas debe ser 100%", vbInformation, "Aviso"
'        ValidaDatosOpe = False
'        Exit Function
'    End If
'End If
'
'ValidaDatosOpe = True
'End Function
'
