VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredBPPTopeMetaCoordJefes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Metas y Topes Coordinadores, Jefes de Agencia y Jefes Territoriales"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   Icon            =   "frmCredBPPTopeMetaCoordJefes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Coordinadores y Jefes de Agencia"
      TabPicture(0)   =   "frmCredBPPTopeMetaCoordJefes.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTopes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Jefes Territoriales"
      TabPicture(1)   =   "frmCredBPPTopeMetaCoordJefes.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Metas y Topes"
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
         Height          =   3975
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   9495
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   3030
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Zonas"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1200
               TabIndex        =   16
               Top             =   315
               Width           =   510
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   7335
            TabIndex        =   25
            Top             =   840
            Width           =   1510
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Tope"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   480
               TabIndex        =   17
               Top             =   315
               Width           =   420
            End
         End
         Begin VB.CommandButton cmdMostrarJT 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   4560
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbMesesJT 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdGuardarJT 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   8280
            TabIndex        =   20
            Top             =   3000
            Width           =   1095
         End
         Begin Spinner.uSpinner uspAnioJT 
            Height          =   315
            Left            =   3480
            TabIndex        =   22
            Top             =   360
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
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Meta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3135
            TabIndex        =   18
            Top             =   930
            Width           =   4215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Mes - Año :"
            Height          =   195
            Left            =   720
            TabIndex        =   23
            Top             =   420
            Width           =   810
         End
      End
      Begin VB.Frame fraTopes 
         Caption         =   "Metas y Topes"
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
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9615
         Begin VB.CommandButton cmdGuardarCJA 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   8400
            TabIndex        =   19
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   7450
            TabIndex        =   14
            Top             =   920
            Width           =   1510
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tope"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   480
               TabIndex        =   15
               Top             =   315
               Width           =   420
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   240
            TabIndex        =   11
            Top             =   920
            Width           =   3030
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Comités"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1200
               TabIndex        =   12
               Top             =   315
               Width           =   690
            End
         End
         Begin VB.ComboBox cmbMesesCJA 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdMostrarCJA 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   6960
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbAgenciasCJA 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   2175
         End
         Begin Spinner.uSpinner uspAnioCJA 
            Height          =   315
            Left            =   2880
            TabIndex        =   5
            Top             =   360
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
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Meta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3255
            TabIndex        =   13
            Top             =   1000
            Width           =   4210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Jefe de Agencia :"
            Height          =   195
            Left            =   1920
            TabIndex        =   10
            Top             =   3050
            Width           =   1245
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Mes - Año :"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   420
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
            Height          =   195
            Left            =   3960
            TabIndex        =   6
            Top             =   360
            Width           =   630
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPTopeMetaCoordJefes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private i As Integer
'Private j As Integer
'
'Private Sub LimpiaDatos(Optional pnPestana As Integer = 0)
'If pnPestana = 0 Then
'    LimpiaFlex feParemetrosCJA
'    txtSaldoJA.Text = "0.00"
'    txtSaldoVencJudJA.Text = "0.00"
'    txtTopeJA.Text = "0.00"
'    cmdGuardarCJA.Enabled = False
'    cmdMostrarCJA.Enabled = True
'Else
'    LimpiaFlex feZona
'    cmdGuardarJT.Enabled = False
'    cmdMostrarJT.Enabled = True
'End If
'End Sub
'
'Private Sub cmbAgenciasCJA_Click()
'LimpiaDatos
'End Sub
'
'
'Private Sub cmbMesesCJA_Click()
'LimpiaDatos
'End Sub
'
'
'Private Sub cmbMesesJT_Click()
'LimpiaDatos 1
'End Sub
'
'Private Sub cmdGuardarCJA_Click()
'If ValidaDatos(1) Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim oBPP As COMNCredito.NCOMBPPR
'        Dim lsAgeCod As String
'        Dim lnMes As Integer
'        Dim lnAnio As Integer
'        Dim lsFecha As String
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        lsAgeCod = Trim(Right(cmbAgenciasCJA.Text, 4))
'        lnMes = CInt(Trim(Right(cmbMesesCJA.Text, 4)))
'        lnAnio = CInt(uspAnioCJA.valor)
'
'        For i = 1 To feParemetrosCJA.Rows - 1
'            Call oBPP.OpeMetaTopeCoord(lsAgeCod, CInt(feParemetrosCJA.TextMatrix(i, 5)), lnMes, lnAnio, CDbl(feParemetrosCJA.TextMatrix(i, 2)), _
'            CDbl(feParemetrosCJA.TextMatrix(i, 3)), CDbl(feParemetrosCJA.TextMatrix(i, 4)), gsCodUser, lsFecha)
'        Next i
'
'        Call oBPP.OpeMetaTopeJefA(lsAgeCod, lnMes, lnAnio, CDbl(txtSaldoJA.Text), CDbl(txtSaldoVencJudJA.Text), CDbl(txtTopeJA.Text), gsCodUser, lsFecha)
'
'        Set oBPP = Nothing
'        cmdMostrarCJA.Enabled = True
'        cmdGuardarCJA.Enabled = False
'        MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'    End If
'End If
'End Sub
'
'Private Sub cmdGuardarJT_Click()
'If ValidaDatos(1, 1) Then
'    If MsgBox("Estas Seguro de guardar los datos", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim oBPP As COMNCredito.NCOMBPPR
'        Dim lnMes As Integer
'        Dim lnAnio As Integer
'        Dim lsFecha As String
'        lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'        Set oBPP = New COMNCredito.NCOMBPPR
'        lnMes = CInt(Trim(Right(cmbMesesJT.Text, 4)))
'        lnAnio = CInt(uspAnioJT.valor)
'
'        For i = 1 To feZona.Rows - 1
'            Call oBPP.OpeMetaTopeJefTerritorial(lnMes, lnAnio, CDbl(feZona.TextMatrix(i, 2)), CDbl(feZona.TextMatrix(i, 3)), CDbl(feZona.TextMatrix(i, 4)), gsCodUser, lsFecha, CInt(feZona.TextMatrix(i, 5)))
'        Next i
'
'        Set oBPP = Nothing
'
'        cmdMostrarJT.Enabled = True
'        cmdGuardarJT.Enabled = False
'        MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'    End If
'End If
'End Sub
'
'Private Sub cmdMostrarCJA_Click()
'If ValidaDatos Then
'    Dim lsAgeCod As String
'    Dim lnMes As Integer
'    Dim lnAnio As Integer
'
'    lsAgeCod = Trim(Right(cmbAgenciasCJA.Text, 4))
'    lnMes = CInt(Trim(Right(cmbMesesCJA.Text, 4)))
'    lnAnio = CInt(uspAnioCJA.valor)
'
'    Call CargaDatos(lnMes, lnAnio, lsAgeCod)
'
'    cmdGuardarCJA.Enabled = True
'    cmdMostrarCJA.Enabled = False
'End If
'End Sub
'
'Private Sub cmdMostrarJT_Click()
'If ValidaDatos(0, 1) Then
'    Dim lnMes As Integer
'    Dim lnAnio As Integer
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Dim rsBPP As ADODB.Recordset
'
'    lnMes = CInt(Trim(Right(cmbMesesJT.Text, 4)))
'    lnAnio = CInt(uspAnioJT.valor)
'
'    Call CargaDatosJT(lnMes, lnAnio)
'
'    cmdMostrarJT.Enabled = False
'    cmdGuardarJT.Enabled = True
'End If
'End Sub
'
'Private Sub Form_Load()
'uspAnioCJA.valor = Year(gdFecSis)
'uspAnioJT.valor = Year(gdFecSis)
'CargaComboMeses cmbMesesCJA
'CargaComboMeses cmbMesesJT
'CargaComboAgencias cmbAgenciasCJA
'cmdGuardarCJA.Enabled = False
'cmdGuardarJT.Enabled = False
'End Sub
'
'
'Private Function ValidaDatos(Optional ByVal pnTipo As Integer = 0, Optional ByVal pnPestana As Integer = 0) As Boolean
'ValidaDatos = True
'If pnPestana = 0 Then
'    If Trim(cmbMesesCJA.Text) = "" Then
'        MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'        ValidaDatos = False
'        cmbMesesCJA.SetFocus
'        Exit Function
'    End If
'
'    If Trim(cmbAgenciasCJA.Text) = "" Then
'        MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'        ValidaDatos = False
'        cmbAgenciasCJA.SetFocus
'        Exit Function
'    End If
'
'    If pnTipo = 1 Then
'        For i = 1 To feParemetrosCJA.Rows - 1
'            For j = 2 To 5
'                If Trim(feParemetrosCJA.TextMatrix(i, j)) = "" Then
'                    MsgBox "Ingrese los datos correctamente en los parametros del " & Trim(feParemetrosCJA.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    feParemetrosCJA.SetFocus
'                    Exit Function
'                End If
'
'                If IsNumeric(Trim(feParemetrosCJA.TextMatrix(i, j))) Then
'                    If CDbl(Trim(feParemetrosCJA.TextMatrix(i, j))) <= 0 Then
'                        MsgBox "Favor de Ingresar Valores Mayores a 0 en los parametros del " & Trim(feParemetrosCJA.TextMatrix(i, 1)), vbInformation, "Aviso"
'                        ValidaDatos = False
'                        feParemetrosCJA.SetFocus
'                        Exit Function
'                    End If
'               End If
'
'               If j = 3 Then
'                If IsNumeric(Trim(feParemetrosCJA.TextMatrix(i, j))) Then
'                     If CDbl(Trim(feParemetrosCJA.TextMatrix(i, j))) > 100 Then
'                         MsgBox "El Valor de la Meta Saldo de Cartera Vencido y Judicial(%) del " & Trim(feParemetrosCJA.TextMatrix(i, 1)) & " no debe ser mayor a 100.", vbInformation, "Aviso"
'                         ValidaDatos = False
'                         feParemetrosCJA.SetFocus
'                         Exit Function
'                     End If
'                End If
'               End If
'            Next j
'        Next i
'
'        If Trim(txtSaldoJA.Text) = "" Then
'            MsgBox "Ingrese correctamente el valor de la Meta Saldo de Cartera del Jefe de Agencia", vbInformation, "Aviso"
'            ValidaDatos = False
'            txtSaldoJA.SetFocus
'            Exit Function
'        End If
'
'        If Trim(txtSaldoVencJudJA.Text) = "" Then
'            MsgBox "Ingrese correctamente el valor de la Meta  Saldo de Cartera Vencido y Judicial del Jefe de Agencia", vbInformation, "Aviso"
'            ValidaDatos = False
'            txtSaldoVencJudJA.SetFocus
'            Exit Function
'        End If
'
'        If Trim(txtTopeJA.Text) = "" Then
'            MsgBox "Ingrese correctamente el valor del Tope a Pagar del Jefe de Agencia", vbInformation, "Aviso"
'            ValidaDatos = False
'            txtTopeJA.SetFocus
'            Exit Function
'        End If
'
'        If IsNumeric(Trim(txtSaldoJA.Text)) Then
'            If CDbl(Trim(txtSaldoJA.Text)) <= 0 Then
'                MsgBox "Ingrese un valor Mayor a 0 de la Meta Saldo de Cartera del Jefe de Agencia", vbInformation, "Aviso"
'                ValidaDatos = False
'                txtSaldoJA.SetFocus
'                Exit Function
'            End If
'        End If
'
'        If IsNumeric(Trim(txtSaldoVencJudJA.Text)) Then
'            If CDbl(Trim(txtSaldoVencJudJA.Text)) <= 0 Then
'                MsgBox "Ingrese un valor Mayor a 0 de la Meta  Saldo de Cartera Vencido y Judicial del Jefe de Agencia", vbInformation, "Aviso"
'                ValidaDatos = False
'                txtSaldoVencJudJA.SetFocus
'                Exit Function
'            End If
'        End If
'
'        If IsNumeric(Trim(txtTopeJA.Text)) Then
'            If CDbl(Trim(txtTopeJA.Text)) <= 0 Then
'                MsgBox "Ingrese un valor Mayor a 0 del Tope a Pagar del Jefe de Agencia", vbInformation, "Aviso"
'                ValidaDatos = False
'                txtTopeJA.SetFocus
'                Exit Function
'            End If
'        End If
'    End If
'Else
'
'    If Trim(cmbMesesJT.Text) = "" Then
'        MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'        ValidaDatos = False
'        cmbMesesJT.SetFocus
'        Exit Function
'    End If
'
'    If pnTipo = 1 Then
'        For i = 1 To feZona.Rows - 1
'            For j = 2 To 5
'                If Trim(feZona.TextMatrix(i, j)) = "" Then
'                    MsgBox "Ingrese los datos correctamente en los parametros de la " & Trim(feZona.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    feZona.SetFocus
'                    Exit Function
'                End If
'
'                If IsNumeric(Trim(feZona.TextMatrix(i, j))) Then
'                    If CDbl(Trim(feZona.TextMatrix(i, j))) <= 0 Then
'                        MsgBox "Favor de Ingresar Valores Mayores a 0 en los parametros de la " & Trim(feZona.TextMatrix(i, 1)), vbInformation, "Aviso"
'                        ValidaDatos = False
'                        feZona.SetFocus
'                        Exit Function
'                    End If
'               End If
'
'               If j = 3 Then
'                If IsNumeric(Trim(feZona.TextMatrix(i, j))) Then
'                     If CDbl(Trim(feZona.TextMatrix(i, j))) > 100 Then
'                         MsgBox "El Valor de la Meta Saldo de Cartera Vencido y Judicial(%) de la " & Trim(feZona.TextMatrix(i, 1)) & " no debe ser mayor a 100.", vbInformation, "Aviso"
'                         ValidaDatos = False
'                         feZona.SetFocus
'                         Exit Function
'                     End If
'                End If
'               End If
'            Next j
'        Next i
'    End If
'End If
'End Function
'
'Private Sub txtSaldoVencJudJA_Change()
'  If Trim(txtSaldoVencJudJA.Text) <> "." Then
'        If CDbl(txtSaldoVencJudJA.Text) > 100 Then
'            txtSaldoVencJudJA.Text = "100.00"
'        End If
'
'         If CDbl(txtSaldoVencJudJA.Text) < 0 Then
'            txtSaldoVencJudJA.Text = "0.00"
'        End If
'    Else
'        txtSaldoVencJudJA.Text = "0.00"
'    End If
'End Sub
'
'Private Sub uspAnioCJA_Change()
'LimpiaDatos
'End Sub
'
'Private Sub CargaDatos(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psAgeCod As String)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.ObtenerMetaTopeCoord(pnMes, pnAnio, psAgeCod)
'
'LimpiaFlex feParemetrosCJA
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    For i = 1 To rsBPP.RecordCount
'        feParemetrosCJA.AdicionaFila
'        feParemetrosCJA.TextMatrix(i, 1) = Trim(rsBPP!DescComite)
'        feParemetrosCJA.TextMatrix(i, 2) = Format(Trim(rsBPP!nSaldo), "###," & String(15, "#") & "#0.00")
'        feParemetrosCJA.TextMatrix(i, 3) = Format(Trim(rsBPP!nSaldoVenJud), "###," & String(15, "#") & "#0.00")
'        feParemetrosCJA.TextMatrix(i, 4) = Format(Trim(rsBPP!nTope), "###," & String(15, "#") & "#0.00")
'        feParemetrosCJA.TextMatrix(i, 5) = Trim(rsBPP!nComite)
'        rsBPP.MoveNext
'    Next i
'End If
'feParemetrosCJA.TopRow = 1
'Set rsBPP = Nothing
'
'Set rsBPP = oBPP.ObtenerMetaTopeJefA(pnMes, pnAnio, psAgeCod)
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    txtSaldoJA.Text = Format(Trim(rsBPP!nSaldo), "###," & String(15, "#") & "#0.00")
'    txtSaldoVencJudJA.Text = Format(Trim(rsBPP!nSaldoVenJud), "###," & String(15, "#") & "#0.00")
'    txtTopeJA.Text = Format(Trim(rsBPP!nTope), "###," & String(15, "#") & "#0.00")
'End If
'
'Set rsBPP = Nothing
'Set oBPP = Nothing
'End Sub
'
'Private Sub uspAnioJT_Change()
'LimpiaDatos 1
'End Sub
'
'Private Sub CargaDatosJT(ByVal pnMes As Integer, ByVal pnAnio As Integer)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'Set oBPP = New COMNCredito.NCOMBPPR
'
'Set rsBPP = oBPP.ObtenerMetaTopeJefTerritorial(pnMes, pnAnio)
'
'LimpiaFlex feZona
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    For i = 1 To rsBPP.RecordCount
'        feZona.AdicionaFila
'        feZona.TextMatrix(i, 1) = Trim(rsBPP!DescZona)
'        feZona.TextMatrix(i, 2) = Format(Trim(rsBPP!nSaldo), "###," & String(15, "#") & "#0.00")
'        feZona.TextMatrix(i, 3) = Format(Trim(rsBPP!nSaldoVenJud), "###," & String(15, "#") & "#0.00")
'        feZona.TextMatrix(i, 4) = Format(Trim(rsBPP!nTope), "###," & String(15, "#") & "#0.00")
'        feZona.TextMatrix(i, 5) = Trim(rsBPP!nZona)
'        rsBPP.MoveNext
'    Next i
'End If
'feZona.TopRow = 1
'
'Set rsBPP = Nothing
'Set oBPP = Nothing
'End Sub
