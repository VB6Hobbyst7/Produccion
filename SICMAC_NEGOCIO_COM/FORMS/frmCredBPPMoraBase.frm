VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPMoraBase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Mora Base"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14820
   Icon            =   "frmCredBPPMoraBase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTBusqueda 
      Height          =   7290
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   12859
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Filtro de Búsquedas"
      TabPicture(0)   =   "frmCredBPPMoraBase.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBusqueda"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCerrar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdVerDetalle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCambiarBase"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CommandButton cmdCambiarBase 
         Caption         =   "Cambiar Base"
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
         Left            =   1440
         TabIndex        =   8
         Top             =   6720
         Width           =   1170
      End
      Begin VB.CommandButton cmdVerDetalle 
         Caption         =   "Ver Detalle"
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
         Left            =   120
         TabIndex        =   7
         Top             =   6720
         Width           =   1170
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
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
         Left            =   13080
         TabIndex        =   6
         Top             =   6720
         Width           =   1170
      End
      Begin VB.Frame fraBusqueda 
         Caption         =   "Filtro"
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
         Height          =   6255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   14175
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "Mostrar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5280
            TabIndex        =   9
            Top             =   550
            Width           =   1170
         End
         Begin VB.OptionButton optUsuario 
            Caption         =   "Usuario"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optAgencia 
            Caption         =   "Agencia"
            Height          =   195
            Left            =   1440
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox txtUsuario 
            Height          =   320
            Left            =   240
            MaxLength       =   4
            TabIndex        =   3
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cmbAgencias 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   600
            Width           =   3735
         End
         Begin SICMACT.FlexEdit feAnalistas 
            Height          =   4695
            Left            =   240
            TabIndex        =   11
            Top             =   1320
            Width           =   13815
            _extentx        =   24368
            _extenty        =   8281
            cols0           =   17
            highlight       =   1
            encabezadosnombres=   "#-Usuario-Base-Dic. A.-Enero-Febrero-Marzo-Abril-Mayo-Junio-Julio-Agosto-Septiembre-Octubre-Noviembre-Diciembre-PersCod"
            encabezadosanchos=   "0-1000-1000-1000-1000-1000-1000-1000-1000-1000-1000-1000-1100-1000-1000-1000-0"
            font            =   "frmCredBPPMoraBase.frx":0326
            font            =   "frmCredBPPMoraBase.frx":034E
            font            =   "frmCredBPPMoraBase.frx":0376
            font            =   "frmCredBPPMoraBase.frx":039E
            font            =   "frmCredBPPMoraBase.frx":03C6
            fontfixed       =   "frmCredBPPMoraBase.frx":03EE
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-R-R-R-R-R-C-R-R-R-R-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-2-2-2-3-2-2-2-2-2-2-2-2-3"
            cantentero      =   15
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label lblFiltroSelect 
            AutoSize        =   -1  'True
            Caption         =   "Filtro Seleccionado: "
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   1440
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPMoraBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************************
''***     Rutina         :   frmCredBPPMoraBase
''***     Descripcion    :   Configurar al Mora Base del BPP
''***     Creado por     :   WIOR
''***     Maquina        :   TIF-1-19
''***     Fecha-Creación :   24/05/2013 08:20:00 AM
''*****************************************************************************************
'Option Explicit
'Dim i As Integer
'Dim nIndex As Integer
'Dim fMatMoraBase() As MoraBase
'
'Private Sub cmdCambiarBase_Click()
'If fMatMoraBase(0).Usuario <> "" Then
'    Call BuscaIndexMatriz(Trim(feAnalistas.TextMatrix(feAnalistas.row, 16)), nIndex)
'    Call frmCredBPPMoraBaseCamb.Inicio(fMatMoraBase(nIndex))
'End If
'End Sub
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Function ValidaDatosMostrar(Optional ByRef psSelect As String = "") As Boolean
'psSelect = ""
'If optUsuario.value Then
'    If Trim(txtUsuario.Text) = "" Then
'        MsgBox "Ingrese el Usuario.", vbInformation, "Aviso"
'        ValidaDatosMostrar = False
'        txtUsuario.SetFocus
'        Exit Function
'    End If
'
'    If Len(Trim(txtUsuario.Text)) <> 4 Then
'        MsgBox "Ingrese el Usuario Completo.", vbInformation, "Aviso"
'        ValidaDatosMostrar = False
'        txtUsuario.SetFocus
'        Exit Function
'    End If
'    psSelect = Trim(txtUsuario.Text)
'Else
'    If Trim(cmbAgencias.Text) = "" Then
'        MsgBox "Seleccione la Agencia.", vbInformation, "Aviso"
'        ValidaDatosMostrar = False
'        cmbAgencias.SetFocus
'        Exit Function
'    End If
'    psSelect = Trim(cmbAgencias.Text)
'End If
'
'ValidaDatosMostrar = True
'End Function
'
'Private Sub cmdMostrar_Click()
'On Error GoTo Error
'Dim sSel As String
'lblFiltroSelect.Caption = "Filtro Seleccionado: "
'
'If ValidaDatosMostrar(sSel) Then
'    Dim sUser As String
'    Dim sAgencia As String
'
'    sUser = IIf(optUsuario.value, sSel, "%")
'    sAgencia = IIf(optAgencia.value, Trim(Right(sSel, 5)), "%")
'
'    Call CargaDatos(sAgencia, sUser)
'
'    lblFiltroSelect.Caption = "Filtro Seleccionado: " & Trim(Left(sSel, 80))
'
'    LimpiaFlex feAnalistas
'
'    If fMatMoraBase(0).Usuario <> "" Then
'        For i = 0 To UBound(fMatMoraBase)
'            feAnalistas.AdicionaFila
'            feAnalistas.TextMatrix(i + 1, 1) = fMatMoraBase(i).Usuario
'            feAnalistas.TextMatrix(i + 1, 16) = fMatMoraBase(i).codigo
'            feAnalistas.TextMatrix(i + 1, 2) = Format(Round(fMatMoraBase(i).PorcMoraBase * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 3) = Format(Round(fMatMoraBase(i).PorcMoraBaseDicAnt * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 4) = Format(Round(fMatMoraBase(i).PorcMoraBaseEne * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 5) = Format(Round(fMatMoraBase(i).PorcMoraBaseFeb * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 6) = Format(Round(fMatMoraBase(i).PorcMoraBaseMar * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 7) = Format(Round(fMatMoraBase(i).PorcMoraBaseAbr * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 8) = Format(Round(fMatMoraBase(i).PorcMoraBaseMay * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 9) = Format(Round(fMatMoraBase(i).PorcMoraBaseJun * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 10) = Format(Round(fMatMoraBase(i).PorcMoraBaseJul * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 11) = Format(Round(fMatMoraBase(i).PorcMoraBaseAgo * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 12) = Format(Round(fMatMoraBase(i).PorcMoraBaseSep * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 13) = Format(Round(fMatMoraBase(i).PorcMoraBaseOct * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 14) = Format(Round(fMatMoraBase(i).PorcMoraBaseNov * 100, 2), "##0.00")
'            feAnalistas.TextMatrix(i + 1, 15) = Format(Round(fMatMoraBase(i).PorcMoraBaseDic * 100, 2), "##0.00")
'        Next i
'        feAnalistas.TopRow = 1
'    Else
'        MsgBox "No hay Datos", vbInformation, "Aviso"
'    End If
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdVerDetalle_Click()
'If fMatMoraBase(0).Usuario <> "" Then
'    Call BuscaIndexMatriz(Trim(feAnalistas.TextMatrix(feAnalistas.row, 16)), nIndex)
'    Call frmCredBPPMoraBaseDet.Inicio(fMatMoraBase(nIndex))
'End If
'End Sub
'
'Private Sub feAnalistas_DblClick()
'cmdVerDetalle_Click
'End Sub
'
'Private Sub Form_Load()
'Call CargaControles
'End Sub
'
'Private Sub CargaControles()
'CargaComboAgencias cmbAgencias
'End Sub
'
'Private Sub optAgencia_Click()
'If optAgencia.value Then
'    txtUsuario.Text = ""
'    txtUsuario.Enabled = False
'    cmbAgencias.Enabled = True
'    cmbAgencias.SetFocus
'End If
'End Sub
'
'Private Sub optUsuario_Click()
'If optUsuario.value Then
'    cmbAgencias.ListIndex = -1
'    txtUsuario.Text = ""
'    txtUsuario.Enabled = True
'    cmbAgencias.Enabled = False
'    txtUsuario.SetFocus
'End If
'End Sub
'
'Private Sub txtUsuario_Change()
'If txtUsuario.SelStart > 0 Then
'    i = Len(Mid(txtUsuario.Text, 1, txtUsuario.SelStart))
'End If
'
'txtUsuario.Text = UCase(txtUsuario.Text)
'txtUsuario.SelStart = i
'End Sub
'
'Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cmdMostrar.SetFocus
'End If
'End Sub
'
'Private Sub CargaDatos(ByVal psAgencia As String, ByVal psUser As String)
'Dim oDBPP As COMDCredito.DCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'Set oDBPP = New COMDCredito.DCOMBPPR
'
'Set rsBPP = oDBPP.ObtenerAnalistaXAgencia(psAgencia, psUser)
'
'    If Not (rsBPP.EOF And rsBPP.BOF) Then
'        ReDim fMatMoraBase(rsBPP.RecordCount - 1)
'        For i = 0 To rsBPP.RecordCount - 1
'
'            fMatMoraBase(i).Usuario = Trim(rsBPP!Usuario)
'            fMatMoraBase(i).codigo = Trim(rsBPP!cPersCod)
'            fMatMoraBase(i).Nombre = Trim(rsBPP!cPersNombre)
'
'            fMatMoraBase(i).Fecha = CDate(rsBPP!Fecha)
'            fMatMoraBase(i).FecMoraBase = CDate(rsBPP!FecMoraBase)
'            fMatMoraBase(i).PorcMoraBase = CDbl(rsBPP!PorcMoraBase)
'            fMatMoraBase(i).SaldoHer = CDbl(rsBPP!SaldoHer)
'            fMatMoraBase(i).MoraHer = CDbl(rsBPP!MoraHer)
'            fMatMoraBase(i).SaldoSal = CDbl(rsBPP!SaldoSal)
'            fMatMoraBase(i).MoraSal = CDbl(rsBPP!MoraSal)
'            fMatMoraBase(i).SaldoPro = CDbl(rsBPP!SaldoPro)
'            fMatMoraBase(i).MoraPro = CDbl(rsBPP!MoraPro)
'            fMatMoraBase(i).Saldo = CDbl(rsBPP!Saldo)
'            fMatMoraBase(i).Mora = CDbl(rsBPP!Mora)
'            fMatMoraBase(i).PorcMora = CDbl(rsBPP!PorcMora)
'
'            fMatMoraBase(i).FechaDicAnt = CDate(rsBPP!FechaDicAnt)
'            fMatMoraBase(i).FecMoraBaseDicAnt = CDate(rsBPP!FecMoraBaseDicAnt)
'            fMatMoraBase(i).PorcMoraBaseDicAnt = CDbl(rsBPP!PorcMoraBaseDicAnt)
'            fMatMoraBase(i).SaldoHerDicAnt = CDbl(rsBPP!SaldoHerDicAnt)
'            fMatMoraBase(i).MoraHerDicAnt = CDbl(rsBPP!MoraHerDicAnt)
'            fMatMoraBase(i).SaldoSalDicAnt = CDbl(rsBPP!SaldoSalDicAnt)
'            fMatMoraBase(i).MoraSalDicAnt = CDbl(rsBPP!MoraSalDicAnt)
'            fMatMoraBase(i).SaldoProDicAnt = CDbl(rsBPP!SaldoProDicAnt)
'            fMatMoraBase(i).MoraProDicAnt = CDbl(rsBPP!MoraProDicAnt)
'            fMatMoraBase(i).SaldoDicAnt = CDbl(rsBPP!SaldoDicAnt)
'            fMatMoraBase(i).MoraDicAnt = CDbl(rsBPP!MoraDicAnt)
'            fMatMoraBase(i).PorcMoraDicAnt = CDbl(rsBPP!PorcMoraDicAnt)
'
'            fMatMoraBase(i).FechaEne = CDate(rsBPP!FechaEne)
'            fMatMoraBase(i).FecMoraBaseEne = CDate(rsBPP!FecMoraBaseEne)
'            fMatMoraBase(i).PorcMoraBaseEne = CDbl(rsBPP!PorcMoraBaseEne)
'            fMatMoraBase(i).SaldoHerEne = CDbl(rsBPP!SaldoHerEne)
'            fMatMoraBase(i).MoraHerEne = CDbl(rsBPP!MoraHerEne)
'            fMatMoraBase(i).SaldoSalEne = CDbl(rsBPP!SaldoSalEne)
'            fMatMoraBase(i).MoraSalEne = CDbl(rsBPP!MoraSalEne)
'            fMatMoraBase(i).SaldoProEne = CDbl(rsBPP!SaldoProEne)
'            fMatMoraBase(i).MoraProEne = CDbl(rsBPP!MoraProEne)
'            fMatMoraBase(i).SaldoEne = CDbl(rsBPP!SaldoEne)
'            fMatMoraBase(i).MoraEne = CDbl(rsBPP!MoraEne)
'            fMatMoraBase(i).PorcMoraEne = CDbl(rsBPP!PorcMoraEne)
'
'            fMatMoraBase(i).FechaFeb = CDate(rsBPP!FechaFeb)
'            fMatMoraBase(i).FecMoraBaseFeb = CDate(rsBPP!FecMoraBaseFeb)
'            fMatMoraBase(i).PorcMoraBaseFeb = CDbl(rsBPP!PorcMoraBaseFeb)
'            fMatMoraBase(i).SaldoHerFeb = CDbl(rsBPP!SaldoHerFeb)
'            fMatMoraBase(i).MoraHerFeb = CDbl(rsBPP!MoraHerFeb)
'            fMatMoraBase(i).SaldoSalFeb = CDbl(rsBPP!SaldoSalFeb)
'            fMatMoraBase(i).MoraSalFeb = CDbl(rsBPP!MoraSalFeb)
'            fMatMoraBase(i).SaldoProFeb = CDbl(rsBPP!SaldoProFeb)
'            fMatMoraBase(i).MoraProFeb = CDbl(rsBPP!MoraProFeb)
'            fMatMoraBase(i).SaldoFeb = CDbl(rsBPP!SaldoFeb)
'            fMatMoraBase(i).MoraFeb = CDbl(rsBPP!MoraFeb)
'            fMatMoraBase(i).PorcMoraFeb = CDbl(rsBPP!PorcMoraFeb)
'
'            fMatMoraBase(i).FechaMar = CDate(rsBPP!FechaMar)
'            fMatMoraBase(i).FecMoraBaseMar = CDate(rsBPP!FecMoraBaseMar)
'            fMatMoraBase(i).PorcMoraBaseMar = CDbl(rsBPP!PorcMoraBaseMar)
'            fMatMoraBase(i).SaldoHerMar = CDbl(rsBPP!SaldoHerMar)
'            fMatMoraBase(i).MoraHerMar = CDbl(rsBPP!MoraHerMar)
'            fMatMoraBase(i).SaldoSalMar = CDbl(rsBPP!SaldoSalMar)
'            fMatMoraBase(i).MoraSalMar = CDbl(rsBPP!MoraSalMar)
'            fMatMoraBase(i).SaldoProMar = CDbl(rsBPP!SaldoProMar)
'            fMatMoraBase(i).MoraProMar = CDbl(rsBPP!MoraProMar)
'            fMatMoraBase(i).SaldoMar = CDbl(rsBPP!SaldoMar)
'            fMatMoraBase(i).MoraMar = CDbl(rsBPP!MoraMar)
'            fMatMoraBase(i).PorcMoraMar = CDbl(rsBPP!PorcMoraMar)
'
'            fMatMoraBase(i).FechaAbr = CDate(rsBPP!FechaAbr)
'            fMatMoraBase(i).FecMoraBaseAbr = CDate(rsBPP!FecMoraBaseAbr)
'            fMatMoraBase(i).PorcMoraBaseAbr = CDbl(rsBPP!PorcMoraBaseAbr)
'            fMatMoraBase(i).SaldoHerAbr = CDbl(rsBPP!SaldoHerAbr)
'            fMatMoraBase(i).MoraHerAbr = CDbl(rsBPP!MoraHerAbr)
'            fMatMoraBase(i).SaldoSalAbr = CDbl(rsBPP!SaldoSalAbr)
'            fMatMoraBase(i).MoraSalAbr = CDbl(rsBPP!MoraSalAbr)
'            fMatMoraBase(i).SaldoProAbr = CDbl(rsBPP!SaldoProAbr)
'            fMatMoraBase(i).MoraProAbr = CDbl(rsBPP!MoraProAbr)
'            fMatMoraBase(i).SaldoAbr = CDbl(rsBPP!SaldoAbr)
'            fMatMoraBase(i).MoraAbr = CDbl(rsBPP!MoraAbr)
'            fMatMoraBase(i).PorcMoraAbr = CDbl(rsBPP!PorcMoraAbr)
'
'            fMatMoraBase(i).FechaMay = CDate(rsBPP!FechaMay)
'            fMatMoraBase(i).FecMoraBaseMay = CDate(rsBPP!FecMoraBaseMay)
'            fMatMoraBase(i).PorcMoraBaseMay = CDbl(rsBPP!PorcMoraBaseMay)
'            fMatMoraBase(i).SaldoHerMay = CDbl(rsBPP!SaldoHerMay)
'            fMatMoraBase(i).MoraHerMay = CDbl(rsBPP!MoraHerMay)
'            fMatMoraBase(i).SaldoSalMay = CDbl(rsBPP!SaldoSalMay)
'            fMatMoraBase(i).MoraSalMay = CDbl(rsBPP!MoraSalMay)
'            fMatMoraBase(i).SaldoProMay = CDbl(rsBPP!SaldoProMay)
'            fMatMoraBase(i).MoraProMay = CDbl(rsBPP!MoraProMay)
'            fMatMoraBase(i).SaldoMay = CDbl(rsBPP!SaldoMay)
'            fMatMoraBase(i).MoraMay = CDbl(rsBPP!MoraMay)
'            fMatMoraBase(i).PorcMoraMay = CDbl(rsBPP!PorcMoraMay)
'
'            fMatMoraBase(i).FechaJun = CDate(rsBPP!FechaJun)
'            fMatMoraBase(i).FecMoraBaseJun = CDate(rsBPP!FecMoraBaseJun)
'            fMatMoraBase(i).PorcMoraBaseJun = CDbl(rsBPP!PorcMoraBaseJun)
'            fMatMoraBase(i).SaldoHerJun = CDbl(rsBPP!SaldoHerJun)
'            fMatMoraBase(i).MoraHerJun = CDbl(rsBPP!MoraHerJun)
'            fMatMoraBase(i).SaldoSalJun = CDbl(rsBPP!SaldoSalJun)
'            fMatMoraBase(i).MoraSalJun = CDbl(rsBPP!MoraSalJun)
'            fMatMoraBase(i).SaldoProJun = CDbl(rsBPP!SaldoProJun)
'            fMatMoraBase(i).MoraProJun = CDbl(rsBPP!MoraProJun)
'            fMatMoraBase(i).SaldoJun = CDbl(rsBPP!SaldoJun)
'            fMatMoraBase(i).MoraJun = CDbl(rsBPP!MoraJun)
'            fMatMoraBase(i).PorcMoraJun = CDbl(rsBPP!PorcMoraJun)
'
'            fMatMoraBase(i).FechaJul = CDate(rsBPP!FechaJul)
'            fMatMoraBase(i).FecMoraBaseJul = CDate(rsBPP!FecMoraBaseJul)
'            fMatMoraBase(i).PorcMoraBaseJul = CDbl(rsBPP!PorcMoraBaseJul)
'            fMatMoraBase(i).SaldoHerJul = CDbl(rsBPP!SaldoHerJul)
'            fMatMoraBase(i).MoraHerJul = CDbl(rsBPP!MoraHerJul)
'            fMatMoraBase(i).SaldoSalJul = CDbl(rsBPP!SaldoSalJul)
'            fMatMoraBase(i).MoraSalJul = CDbl(rsBPP!MoraSalJul)
'            fMatMoraBase(i).SaldoProJul = CDbl(rsBPP!SaldoProJul)
'            fMatMoraBase(i).MoraProJul = CDbl(rsBPP!MoraProJul)
'            fMatMoraBase(i).SaldoJul = CDbl(rsBPP!SaldoJul)
'            fMatMoraBase(i).MoraJul = CDbl(rsBPP!MoraJul)
'            fMatMoraBase(i).PorcMoraJul = CDbl(rsBPP!PorcMoraJul)
'
'            fMatMoraBase(i).FechaAgo = CDate(rsBPP!FechaAgo)
'            fMatMoraBase(i).FecMoraBaseAgo = CDate(rsBPP!FecMoraBaseAgo)
'            fMatMoraBase(i).PorcMoraBaseAgo = CDbl(rsBPP!PorcMoraBaseAgo)
'            fMatMoraBase(i).SaldoHerAgo = CDbl(rsBPP!SaldoHerAgo)
'            fMatMoraBase(i).MoraHerAgo = CDbl(rsBPP!MoraHerAgo)
'            fMatMoraBase(i).SaldoSalAgo = CDbl(rsBPP!SaldoSalAgo)
'            fMatMoraBase(i).MoraSalAgo = CDbl(rsBPP!MoraSalAgo)
'            fMatMoraBase(i).SaldoProAgo = CDbl(rsBPP!SaldoProAgo)
'            fMatMoraBase(i).MoraProAgo = CDbl(rsBPP!MoraProAgo)
'            fMatMoraBase(i).SaldoAgo = CDbl(rsBPP!SaldoAgo)
'            fMatMoraBase(i).MoraAgo = CDbl(rsBPP!MoraAgo)
'            fMatMoraBase(i).PorcMoraAgo = CDbl(rsBPP!PorcMoraAgo)
'
'            fMatMoraBase(i).FechaSep = CDate(rsBPP!FechaSep)
'            fMatMoraBase(i).FecMoraBaseSep = CDate(rsBPP!FecMoraBaseSep)
'            fMatMoraBase(i).PorcMoraBaseSep = CDbl(rsBPP!PorcMoraBaseSep)
'            fMatMoraBase(i).SaldoHerSep = CDbl(rsBPP!SaldoHerSep)
'            fMatMoraBase(i).MoraHerSep = CDbl(rsBPP!MoraHerSep)
'            fMatMoraBase(i).SaldoSalSep = CDbl(rsBPP!SaldoSalSep)
'            fMatMoraBase(i).MoraSalSep = CDbl(rsBPP!MoraSalSep)
'            fMatMoraBase(i).SaldoProSep = CDbl(rsBPP!SaldoProSep)
'            fMatMoraBase(i).MoraProSep = CDbl(rsBPP!MoraProSep)
'            fMatMoraBase(i).SaldoSep = CDbl(rsBPP!SaldoSep)
'            fMatMoraBase(i).MoraSep = CDbl(rsBPP!MoraSep)
'            fMatMoraBase(i).PorcMoraSep = CDbl(rsBPP!PorcMoraSep)
'
'            fMatMoraBase(i).FechaOct = CDate(rsBPP!FechaOct)
'            fMatMoraBase(i).FecMoraBaseOct = CDate(rsBPP!FecMoraBaseOct)
'            fMatMoraBase(i).PorcMoraBaseOct = CDbl(rsBPP!PorcMoraBaseOct)
'            fMatMoraBase(i).SaldoHerOct = CDbl(rsBPP!SaldoHerOct)
'            fMatMoraBase(i).MoraHerOct = CDbl(rsBPP!MoraHerOct)
'            fMatMoraBase(i).SaldoSalOct = CDbl(rsBPP!SaldoSalOct)
'            fMatMoraBase(i).MoraSalOct = CDbl(rsBPP!MoraSalOct)
'            fMatMoraBase(i).SaldoProOct = CDbl(rsBPP!SaldoProOct)
'            fMatMoraBase(i).MoraProOct = CDbl(rsBPP!MoraProOct)
'            fMatMoraBase(i).SaldoOct = CDbl(rsBPP!SaldoOct)
'            fMatMoraBase(i).MoraOct = CDbl(rsBPP!MoraOct)
'            fMatMoraBase(i).PorcMoraOct = CDbl(rsBPP!PorcMoraOct)
'
'            fMatMoraBase(i).FechaNov = CDate(rsBPP!FechaNov)
'            fMatMoraBase(i).FecMoraBaseNov = CDate(rsBPP!FecMoraBaseNov)
'            fMatMoraBase(i).PorcMoraBaseNov = CDbl(rsBPP!PorcMoraBaseNov)
'            fMatMoraBase(i).SaldoHerNov = CDbl(rsBPP!SaldoHerNov)
'            fMatMoraBase(i).MoraHerNov = CDbl(rsBPP!MoraHerNov)
'            fMatMoraBase(i).SaldoSalNov = CDbl(rsBPP!SaldoSalNov)
'            fMatMoraBase(i).MoraSalNov = CDbl(rsBPP!MoraSalNov)
'            fMatMoraBase(i).SaldoProNov = CDbl(rsBPP!SaldoProNov)
'            fMatMoraBase(i).MoraProNov = CDbl(rsBPP!MoraProNov)
'            fMatMoraBase(i).SaldoNov = CDbl(rsBPP!SaldoNov)
'            fMatMoraBase(i).MoraNov = CDbl(rsBPP!MoraNov)
'            fMatMoraBase(i).PorcMoraNov = CDbl(rsBPP!PorcMoraNov)
'
'            fMatMoraBase(i).FechaDic = CDate(rsBPP!FechaDic)
'            fMatMoraBase(i).FecMoraBaseDic = CDate(rsBPP!FecMoraBaseDic)
'            fMatMoraBase(i).PorcMoraBaseDic = CDbl(rsBPP!PorcMoraBaseDic)
'            fMatMoraBase(i).SaldoHerDic = CDbl(rsBPP!SaldoHerDic)
'            fMatMoraBase(i).MoraHerDic = CDbl(rsBPP!MoraHerDic)
'            fMatMoraBase(i).SaldoSalDic = CDbl(rsBPP!SaldoSalDic)
'            fMatMoraBase(i).MoraSalDic = CDbl(rsBPP!MoraSalDic)
'            fMatMoraBase(i).SaldoProDic = CDbl(rsBPP!SaldoProDic)
'            fMatMoraBase(i).MoraProDic = CDbl(rsBPP!MoraProDic)
'            fMatMoraBase(i).SaldoDic = CDbl(rsBPP!SaldoDic)
'            fMatMoraBase(i).MoraDic = CDbl(rsBPP!MoraDic)
'            fMatMoraBase(i).PorcMoraDic = CDbl(rsBPP!PorcMoraDic)
'
'
'            rsBPP.MoveNext
'        Next i
'    Else
'        ReDim fMatMoraBase(0)
'    End If
'Set rsBPP = Nothing
'Set oDBPP = Nothing
'End Sub
'
'Private Sub BuscaIndexMatriz(ByVal psPersCod As String, ByRef pnIndex As Integer)
'pnIndex = -1
'For i = 0 To UBound(fMatMoraBase)
'    If Trim(fMatMoraBase(i).codigo) = Trim(psPersCod) Then
'        pnIndex = i
'        Exit For
'    End If
'Next i
'End Sub
'
'Public Sub LimpiaDatos()
'    LimpiaFlex feAnalistas
'End Sub
