VERSION 5.00
Begin VB.Form frmCredAutorizar 
   Caption         =   "Autorizar Aprobacion de Credito"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   Icon            =   "frmCredAutorizar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin SICMACT.FlexEdit FEAutoCred 
      Height          =   1815
      Left            =   90
      TabIndex        =   23
      Top             =   3405
      Width           =   7110
      _extentx        =   12541
      _extenty        =   3201
      cols0           =   7
      fixedcols       =   0
      scrollbars      =   1
      highlight       =   1
      rowsizingmode   =   1
      encabezadosnombres=   "Item-Cargo-Usuario-VoBo-CodCargo-Coment.-ComDes"
      encabezadosanchos=   "400-3600-1200-700-0-700-0"
      font            =   "frmCredAutorizar.frx":030A
      font            =   "frmCredAutorizar.frx":0336
      font            =   "frmCredAutorizar.frx":0362
      font            =   "frmCredAutorizar.frx":038E
      font            =   "frmCredAutorizar.frx":03BA
      fontfixed       =   "frmCredAutorizar.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-3-X-X-X"
      listacontroles  =   "0-0-0-4-0-0-0"
      encabezadosalineacion=   "C-C-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0"
      textarray0      =   "Item"
      selectionmode   =   1
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1380
      TabIndex        =   22
      Top             =   5385
      Width           =   1275
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   5850
      TabIndex        =   3
      Top             =   5385
      Width           =   1275
   End
   Begin VB.CommandButton CmdAutorizar 
      Caption         =   "&Autorizar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   75
      TabIndex        =   2
      Top             =   5385
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Credito"
      Height          =   2610
      Left            =   405
      TabIndex        =   1
      Top             =   660
      Width           =   6420
      Begin VB.Label LblDiaVen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1095
         TabIndex        =   21
         Top             =   2145
         Width           =   990
      End
      Begin VB.Label Label12 
         Caption         =   "Dia Venc :"
         Height          =   285
         Left            =   150
         TabIndex        =   20
         Top             =   2145
         Width           =   750
      End
      Begin VB.Label LblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4725
         TabIndex        =   19
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label10 
         Caption         =   "Moneda  :"
         Height          =   285
         Left            =   3780
         TabIndex        =   18
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label LblPlazo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2880
         TabIndex        =   17
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label8 
         Caption         =   "Plazo   :"
         Height          =   285
         Left            =   2250
         TabIndex        =   16
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label LblNCuo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1095
         TabIndex        =   15
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label6 
         Caption         =   "Nro Cuotas:"
         Height          =   285
         Left            =   150
         TabIndex        =   14
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label LblAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1095
         TabIndex        =   13
         Top             =   1410
         Width           =   2550
      End
      Begin VB.Label Label3 
         Caption         =   "Agencia  :"
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   1425
         Width           =   750
      End
      Begin VB.Label LblPrestamo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5025
         TabIndex        =   11
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label4 
         Caption         =   "Prestamo :"
         Height          =   285
         Left            =   4080
         TabIndex        =   10
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label LblProd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1110
         TabIndex        =   9
         Top             =   1035
         Width           =   2550
      End
      Begin VB.Label Label2 
         Caption         =   "Producto :"
         Height          =   285
         Left            =   165
         TabIndex        =   8
         Top             =   1050
         Width           =   750
      End
      Begin VB.Label LblCli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1125
         TabIndex        =   7
         Top             =   660
         Width           =   4305
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente    :"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   750
      End
      Begin VB.Label LblAna 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1125
         TabIndex        =   5
         Top             =   315
         Width           =   4305
      End
      Begin VB.Label LblAnalista 
         Caption         =   "Analista   :"
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   315
         Width           =   750
      End
   End
   Begin SICMACT.ActXCodCta ActCredito 
      Height          =   435
      Left            =   405
      TabIndex        =   0
      Top             =   210
      Width           =   3720
      _extentx        =   6562
      _extenty        =   767
      texto           =   "Credito : "
      enabledcta      =   -1
      enabledprod     =   -1
      enabledage      =   -1
   End
End
Attribute VB_Name = "frmCredAutorizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim bAprobacion As Boolean
'Dim objPista As COMManejador.Pista
'
'
'Private Sub LimpiaData()
'
'        LblAna.Caption = ""
'        LblCli.Caption = ""
'        LblProd.Caption = ""
'        LblMoneda.Caption = ""
'        LblAge.Caption = ""
'        LblNCuo.Caption = ""
'        LblPlazo.Caption = ""
'        LblPrestamo.Caption = ""
'        LblDiaVen.Caption = ""
'        CmdAutorizar.Enabled = False
'        FEAutoCred.Clear
'        FEAutoCred.Rows = 2
'        FEAutoCred.FormaCabecera
'        ActCredito.NroCuenta = ""
'        ActCredito.CMAC = gsCodCMAC
'        ActCredito.Age = gsCodAge
'
'End Sub
'
'Private Sub ActCredito_KeyPress(KeyAscii As Integer)
'Dim d As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim rsDat As ADODB.Recordset
'    If KeyAscii = 13 Then
'        Set d = New COMDCredito.DCOMCredito
'        If d.CargaDatosAutorizacionCredito(ActCredito.NroCuenta, rsDat, R) = "" Then
'            MsgBox "No se pudo encontrar el Credito, o el Credito no esta sugerido por el Analista", vbInformation, "Aviso"
'            Call LimpiaData
'        Else
'            LblAna.Caption = rsDat!cAnalista
'            LblCli.Caption = rsDat!cCliente
'            LblProd.Caption = rsDat!cProd
'            LblMoneda.Caption = rsDat!cmoneda
'            LblAge.Caption = rsDat!cAgeDescripcion
'            LblNCuo.Caption = Trim(str(rsDat!nCuotas))
'            LblPlazo.Caption = Trim(str(rsDat!nPlazo))
'            Me.LblPrestamo.Caption = Trim(str(rsDat!nMonto))
'            LblDiaVen.Caption = Trim(str(rsDat!nPeriodoFechaFija))
'
'            FEAutoCred.Clear
'            FEAutoCred.Rows = 2
'            FEAutoCred.FormaCabecera
'
'            CmdAutorizar.Enabled = True
'
'            '05-05-2005
'            'If rsDat!cCodAnalista = gsCodPersUser Then
'            '    FEAutoCred.AdicionaFila
'            '    FEAutoCred.TextMatrix(R.Bookmark, 1) = "Analista del Credito"
'            '    FEAutoCred.TextMatrix(R.Bookmark, 2) = gsCodUser
'            '    FEAutoCred.TextMatrix(R.Bookmark, 3) = "1"
'            '    FEAutoCred.TextMatrix(R.Bookmark, 4) = "001"
'            '    FEAutoCred.TextMatrix(R.Bookmark, 5) = "NO"
'            '    FEAutoCred.TextMatrix(R.Bookmark, 6) = ""
'            'End If
'            '*****************************
'            Do While Not R.EOF
'                '06-05-2005
'                'If R!cAgeCod = "" Or Mid(ActCredito.NroCuenta, 4, 2) = R!cAgeCod Then
'                ''****************
'                    FEAutoCred.AdicionaFila
'                    FEAutoCred.TextMatrix(R.Bookmark, 1) = R!cNomCargo
'                    FEAutoCred.TextMatrix(R.Bookmark, 2) = IIf(IsNull(R!cCodUsu), "", R!cCodUsu)
'                    FEAutoCred.TextMatrix(R.Bookmark, 3) = IIf(IsNull(R!cCodUsu), "0", "1")
'                    FEAutoCred.TextMatrix(R.Bookmark, 4) = Trim(R!cCodCargo)
'                    FEAutoCred.TextMatrix(R.Bookmark, 5) = IIf(IsNull(R!bComen), "NO", IIf(R!bComen, "SI", "NO"))
'                    FEAutoCred.TextMatrix(R.Bookmark, 6) = IIf(IsNull(R!cComen), "", Trim(R!cComen))
'                'End If
'                R.MoveNext
'            Loop
'
'            If Not VerificaSoloUnaAutorizacion Then
'                CmdAutorizar.Enabled = False
'            End If
'        End If
'        Set d = Nothing
'    End If
'
'    If Not bAprobacion Then
'        CmdNuevo.Enabled = True
'    End If
'End Sub
'
'Private Function VerificaSoloUnaAutorizacion() As Boolean
'Dim C As Integer
'Dim nCont As Integer
'
'    VerificaSoloUnaAutorizacion = True
'
''08-05-2006
''Esta restriccion queda pendiente
''    nCont = 0
''    For c = 1 To FEAutoCred.Rows - 1
''        If Trim(FEAutoCred.TextMatrix(c, 2)) = gsCodUser Then
''            nCont = nCont + 1
''            If nCont = 1 Then
''                VerificaSoloUnaAutorizacion = False
''                Exit Function
''            End If
''        End If
''    Next c
''****************************
'
'End Function
'
'Private Function Autorizar(ByVal pbBusca As Boolean) As Boolean
'Dim C As Integer
'Dim d As COMDCredito.DCOMCredito
'Dim sComen As String
'    Autorizar = False
'
'    If Not VerificaSoloUnaAutorizacion Then
'        Autorizar = False
'        MsgBox "Ya autorizo este Credito", vbInformation, "Aviso"
'        Exit Function
'    End If
'    Set d = New COMDCredito.DCOMCredito
'    For C = 1 To FEAutoCred.Rows - 1
'        '06-05-2005
'        If d.UsuarioPerteneceACargo(FEAutoCred.TextMatrix(C, 4), gsCodUser, gsCodAge) Then  'Or FEAutoCred.TextMatrix(c, 4) = "001" Then
'            If pbBusca Then
'                Autorizar = True
'                Exit Function
'            Else
'                'Verifica si se debe ingresar comentario
'                If Me.FEAutoCred.TextMatrix(C, 5) = "SI" Then
'                    Call frmCredAutorizaComen.AgregarComentario(sComen)
'                Else
'                    sComen = ""
'                End If
'
'                Call d.NuevaAutorizacion(ActCredito.NroCuenta, FEAutoCred.TextMatrix(C, 4), gsCodUser, _
'                           gsCodPersUser, gdFecSis, sComen)
'
'                ''*** PEAC 20090220
'                objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActCredito.NroCuenta, gCodigoCuenta
'
'                Autorizar = True
'                '07-05-2006
'                'Autorizar para todos los cargos definidos
'                'Exit Function
'                '*************
'            End If
'        End If
'    Next C
'    Set d = Nothing
'    If pbBusca Then
'         MsgBox "Usted No esta Asignado para Autorizar este Credito", vbInformation, "Aviso"
'    End If
'
'
'End Function
'
'Private Sub CmdAutorizar_Click()
'Dim d As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim rsDat As ADODB.Recordset
'    If Autorizar(False) Then
'        Call ActCredito_KeyPress(13)
'        CmdAutorizar.Enabled = False
'    End If
'
'End Sub
'
'Public Function Aprobacion(ByVal psCtaCod As String) As Boolean
'Dim i As Integer
'
'    bAprobacion = True
'    ActCredito.NroCuenta = psCtaCod
'    Call ActCredito_KeyPress(13)
'    Me.Show 1
'    CmdNuevo.Enabled = False
'
'    ActCredito.NroCuenta = psCtaCod
'    Call ActCredito_KeyPress(13)
'
'    Aprobacion = True
'    For i = 1 To FEAutoCred.Rows - 1
'        If FEAutoCred.TextMatrix(i, 3) <> "." Then
'            Aprobacion = False
'        End If
'    Next i
'
'End Function
'Private Sub cmdNuevo_Click()
'    Call LimpiaData
'End Sub
'
'Private Sub cmdsalir_Click()
'    Unload Me
'End Sub
'
'Private Sub FEAutoCred_DblClick()
'
'    If FEAutoCred.TextMatrix(FEAutoCred.row, 5) = "SI" Then
'        Call frmCredAutorizaComen.MostarComentario(FEAutoCred.TextMatrix(FEAutoCred.row, 6))
'    End If
'
'End Sub
'
'
'Private Sub Form_Load()
'    ActCredito.CMAC = gsCodCMAC
'    ActCredito.Age = gsCodAge
'    CentraForm Me
'    Set objPista = New COMManejador.Pista
'    gsOpeCod = gCredAutorizarAprobacion
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    Set objPista = Nothing
'End Sub
