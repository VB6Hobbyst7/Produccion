VERSION 5.00
Begin VB.Form frmVinculaCta 
   Caption         =   "Vinculación de Cuentas"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   720
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   10875
      Begin VB.CommandButton CmdRegistrar 
         Caption         =   "Registrar"
         Height          =   360
         Left            =   165
         TabIndex        =   7
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   9360
         TabIndex        =   6
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuentas"
      Height          =   2835
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   10875
      Begin SICMACT.FlexEdit flxCuentas 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   10695
         _ExtentX        =   17965
         _ExtentY        =   4419
         Cols0           =   8
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cuenta-Tipo Cuenta-N° Titular-Registrada-Tipo Producto-Prioridad-Afilia"
         EncabezadosAnchos=   "400-1900-1700-1100-1100-2000-1000-900"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-6-7"
         ListaControles  =   "0-0-0-0-0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7250
      Begin VB.Label TxtNumTarj 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Nombre :"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label txtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   6375
   End
End
Attribute VB_Name = "frmVinculaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CodPers As String
Dim nNumFila As Integer
Dim lsNumTarj As String
Dim lsCuenta As String
Dim lsFecha As String
Dim lnPrior As Integer
Dim lnRelac As Integer
Dim lnConsulta As Integer
Dim lnRetiro As Integer
Dim i As Integer

Public Sub Inicia(ByVal sPersCod As String)
    CodPers = sPersCod
    Call CargaDatos(CodPers)
    Me.Show 1
End Sub

Private Sub CargaDatos(ByVal pPersCod As String)
Dim oCaptaAN As New COMNCaptaGenerales.NCOMCaptaGenerales
Dim R As ADODB.Recordset
 Set R = New ADODB.Recordset
    Set R = oCaptaAN.ObtieneDatosCuentas(pPersCod)
    Set Me.flxCuentas.Recordset = R
    Set R = Nothing
        
     Set R = New ADODB.Recordset
     Set R = oCaptaAN.DatosPersCuenta(pPersCod)
     Me.TxtNumTarj.Caption = R!cNumTarjeta
     Me.txtCliente.Caption = R!cPersNombre
     Set R = Nothing
End Sub

Private Sub cmdRegistrar_Click()
Dim oCaptaAN As New COMNCaptaGenerales.NCOMCaptaGenerales
Dim sResp As String
Dim sTramaResp As String

    nNumFila = flxCuentas.Rows - 1
    lsNumTarj = TxtNumTarj.Caption
    lsFecha = Format(Now(), "YYYY/MM/DD HH:MM:SS")
    lnConsulta = 1
    lnRetiro = 1
    
    If ValidaPrior() Then
        For i = 1 To nNumFila
            lsCuenta = flxCuentas.TextMatrix(i, 1)
            lnPrior = flxCuentas.TextMatrix(i, 6)
            lnRelac = IIf(flxCuentas.TextMatrix(i, 7) = ".", 1, 0)
                
            If lnRelac = 1 Then
                If flxCuentas.TextMatrix(i, 4) = "NO" Then
                  Call oCaptaAN.RegistraTarjetaCuenta(lsNumTarj, lsCuenta, lsFecha, lnPrior, lnRelac, lnConsulta, lnRetiro)
                Else
                  Call oCaptaAN.ActualizaTarjetaCuenta(lsNumTarj, lsCuenta, lnRelac, lnPrior, lnConsulta, lnRetiro)
                End If

                Call oCaptaAN.RegistraHistorPrioridad(lsNumTarj, lsCuenta, lsFecha, lnPrior)
            Else
                If flxCuentas.TextMatrix(i, 4) = "SI" Then
                    Call oCaptaAN.DesafiliarTarjetaCuenta(lsNumTarj, lsCuenta)
                    
                   Call oCaptaAN.RegistraHistorPrioridad(lsNumTarj, lsCuenta, lsFecha, 0)
                End If
            End If
        Next i
        
        MsgBox "Cuentas Registradas Correctamente"
        MsgBox "Se va a proceder a generar el documento de condiciones generales de uso de tarjeta de débito.", vbInformation + vbOKOnly, "Aviso"
        Unload Me
        
        Call GenerarPDF(lsNumTarj)
        
        Call CargaDatos(CodPers)
   End If
End Sub

Private Function ValidaPrior() As Boolean
Dim lnMonCta As Integer
Dim lsTipoCta As String
'Variables para validar la prioriad
Dim PS1, PS2, PS3, PS4, PS5, PS6, PS7, PS8, PS9, PS10 As Integer
Dim PD1, PD2, PD3, PD4, PD5, PD6, PD7, PD8, PD9, PD10 As Integer
Dim nValAct As Integer

    nNumFila = flxCuentas.Rows - 1
    
    For i = 1 To nNumFila
        lnMonCta = Mid(flxCuentas.TextMatrix(i, 1), 9, 1)
        lnPrior = flxCuentas.TextMatrix(i, 6)
        lsTipoCta = flxCuentas.TextMatrix(i, 5)
        
        If flxCuentas.TextMatrix(i, 7) = "." Then
            If lnPrior = 0 Then
                ValidaPrior = False
                MsgBox "Tiene que ingresar la prioridad de la cuenta que va afiliar", vbInformation, "MENSAJE DEL SISTEMA"
                Exit Function
            End If
            If lnMonCta = 1 Then
                If PS1 = 1 And lnPrior = 1 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS2 = 1 And lnPrior = 2 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS3 = 1 And lnPrior = 3 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS4 = 1 And lnPrior = 4 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS5 = 1 And lnPrior = 5 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS6 = 1 And lnPrior = 6 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS7 = 1 And lnPrior = 7 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS8 = 1 And lnPrior = 8 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS9 = 1 And lnPrior = 9 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                ElseIf PS10 = 1 And lnPrior = 10 Then
                    ValidaPrior = False
                    MsgBox "No puede vincular dos cuentas de la misma moneda con la misma prioridad", vbInformation, "MENSAJE DEL SISTEMA"
                    Exit Function
                End If
                
                If lnPrior = 1 Then
                    PS1 = 1
                ElseIf lnPrior = 2 Then
                    PS2 = 1
                ElseIf lnPrior = 3 Then
                    PS3 = 1
                ElseIf lnPrior = 4 Then
                    PS4 = 1
                ElseIf lnPrior = 5 Then
                    PS5 = 1
                ElseIf lnPrior = 6 Then
                    PS6 = 1
                ElseIf lnPrior = 7 Then
                    PS7 = 1
                ElseIf lnPrior = 8 Then
                    PS8 = 1
                ElseIf lnPrior = 9 Then
                    PS9 = 1
                Else
                    PS10 = 1
                End If
                
            Else
                If PD1 = 1 And lnPrior = 1 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD2 = 1 And lnPrior = 2 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD3 = 1 And lnPrior = 3 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD4 = 1 And lnPrior = 4 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD5 = 1 And lnPrior = 5 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD6 = 1 And lnPrior = 6 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD7 = 1 And lnPrior = 7 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD8 = 1 And lnPrior = 8 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD9 = 1 And lnPrior = 9 Then
                    ValidaPrior = False
                    Exit Function
                ElseIf PD10 = 1 And lnPrior = 10 Then
                    ValidaPrior = False
                    Exit Function
                End If
                
                If lnPrior = 1 Then
                    PD1 = 1
                ElseIf lnPrior = 2 Then
                    PD2 = 1
                ElseIf lnPrior = 3 Then
                    PD3 = 1
                ElseIf lnPrior = 4 Then
                    PD4 = 1
                ElseIf lnPrior = 5 Then
                    PD5 = 1
                ElseIf lnPrior = 6 Then
                    PD6 = 1
                ElseIf lnPrior = 7 Then
                    PD7 = 1
                ElseIf lnPrior = 8 Then
                    PD8 = 1
                ElseIf lnPrior = 9 Then
                    PD9 = 1
                Else
                    PD10 = 1
                End If
                
            End If
            nValAct = nValAct + 1
            
        End If
    Next i
    
    If nValAct = 0 Then
        ValidaPrior = False
        MsgBox "Tiene que seleccionar por lo menos una cuenta", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PS1 = 0 And PS2 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 1 en Soles", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PD1 = 0 And PD2 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 1 en Dolares", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PS2 = 0 And PS3 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 2 en Soles", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PD2 = 0 And PD3 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 2 en Dolares", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PS3 = 0 And PS4 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 3 en Soles", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PD3 = 0 And PD4 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 3 en Dolares", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PS4 = 0 And PS5 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 4 en Soles", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PD4 = 0 And PD5 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 4 en Dolares", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PS5 = 0 And PS6 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 5 en Soles", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    If PD6 = 0 And PD6 > 0 Then
        ValidaPrior = False
        MsgBox "Tiene que ingresar una cuenta con prioridad 5 en Dolares", vbInformation, "MENSAJE DEL SISTEMA"
        Exit Function
    End If
    
    ValidaPrior = True
    
End Function

Public Sub GenerarPDF(ByVal pcNumTarjeta As String)
    Dim oPDF As New cPDF
    Dim cNombreArchivo As String
    Dim rDatosAfiliacion As ADODB.Recordset
    Dim oCaptaAN As New COMNCaptaGenerales.NCOMCaptaGenerales
   
    'cargando cuentas afiliadas a la tarjeta
    Set rDatosAfiliacion = oCaptaAN.DatosPDF(pcNumTarjeta)

    If Not (rDatosAfiliacion.BOF And rDatosAfiliacion.EOF) Then

        cNombreArchivo = "ActTarjeta_" & rDatosAfiliacion!cPersCod & "_" & Format(gdFecSis, "DDMMYYYY") & Format(Time, "hhmmss") & ".pdf"
        If Not oPDF.PDFCreate(App.Path & "\Spooler\" & cNombreArchivo) Then
            Exit Sub
        End If
        
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Set R = oCaptaAN.DOCUMENTOPDF()
        
        oPDF.Author = gsCodUser
        oPDF.Creator = "SICMACT - Tarjeta"
        oPDF.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
        oPDF.Title = "Activación de tarjeta N° " & pcNumTarjeta

        oPDF.Fonts.Add "F1", "Arial Narrow", TrueType, Normal, WinAnsiEncoding
        oPDF.Fonts.Add "F2", "Arial Narrow", TrueType, Bold, WinAnsiEncoding
        oPDF.NewPage A4_Vertical

         Dim nTopX As Integer
         Dim nLefY As Integer
         Dim inFila As Integer
         Dim nHeight As Double
         Dim nWidth As Integer
         Dim nVinetas As Integer
         Dim space As Integer
         Dim nCentrar As Long
         Dim nTamLet As Integer

         nTopX = 70
         nLefY = 60
         inFila = 0
         nHeight = 11.25
         nWidth = 490
         nVinetas = 0
         space = 0


        If R.RecordCount > 0 Then
            Do While Not R.EOF
            If R!cTipo = "TITULO" Then
            oPDF.WTextBox 50, nLefY, nHeight, nWidth, R!cDescripcion, "F2", 11, hCenter, , , , , , 3 'titulo
            inFila = inFila + 1
            nCentrar = nCentrar + 60 + 12

            Else

                Dim filas As Integer
                Dim Residuo As Integer

                If Mid(R!cTipo, 1, 1) = "P" Then
                    filas = Len(R!cDescripcion) / 183
                    Residuo = Len(R!cDescripcion) Mod 183
                    If Residuo > 0 Then
                    filas = filas + 1
                    End If
                    oPDF.WTextBox nCentrar, nLefY, filas * nHeight, nWidth, Replace(R!cDescripcion, "#", String(1.5, vbTab)), "F1", 11, hLeft, , , , , , 1 'titulo
                    inFila = inFila + filas
                   nCentrar = 3 + nCentrar + filas * nHeight
                End If

                If Mid(R!cTipo, 1, 1) = "#" Then
                    nVinetas = nVinetas + 1
                    filas = Len(R!cDescripcion) / 110
                    Residuo = Len(R!cDescripcion) Mod 110
                    nTamLet = filas Mod 2
                    If Residuo > 0 And Residuo < 100 And (Len(R!cDescripcion) / 110) > 2 Then
                    filas = filas + 1
                    End If
                    If nVinetas = 14 Or nVinetas = 1 Or nVinetas = 4 Or nVinetas = 11 Or (nVinetas > 15 And nVinetas < 19) Then
                    filas = filas - 1
                    End If

                    oPDF.WTextBox nCentrar, nLefY, nHeight, 18, nVinetas & ".", "F1", 11, hLeft, , , , , , 1 'titulo
                    oPDF.WTextBox nCentrar, nLefY + 15, filas * nHeight, nWidth - 15, Replace(R!cDescripcion, "#", String(1.5, vbTab)), "F1", 11, hLeft, , , , , , 1 'titulo
                    nCentrar = 5 + nCentrar + filas * nHeight

                End If
                If nCentrar > 760 Then
                oPDF.NewPage A4_Vertical
                 nCentrar = 40
                 nLefY = 60
                 inFila = 0
                 nHeight = 11
                 nWidth = 495
                End If

            End If
            R.MoveNext
        Loop

        Else

        End If

        oPDF.NewPage A4_Vertical
        nTopX = 70
        nLefY = 60
        inFila = 0
        nHeight = 20
        nWidth = 490

        oPDF.WTextBox 50, 50, 15, 500, "SOLICITUD DE ACTIVACIÓN DE TARJETA DE DÉBITO Y AFILIACIÓN DE CUENTAS", "F2", 10, hCenter, vMiddle, , , , , 3
        oPDF.WTextBox nTopX + nHeight * 1, 80, 15, 100, "Agencia:", "F1", 8, hLeft, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 1, 185, 15, 100, gsNomAge, "F1", 8, hLeft, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 1, 290, 15, 100, "Fecha:", "F1", 8, hLeft, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 1, 395, 15, 100, gdFecSis, "F1", 8, hLeft, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 2, 80, 15, 100, "Nombres y apellidos:", "F1", 8, hLeft, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 2, 185, 15, 310, rDatosAfiliacion!Nombre, "F1", 8, hLeft, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 3, 80, 15, 100, "Tarjeta de Débito N°:", "F1", 8, hLeft, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 3, 185, 15, 310, pcNumTarjeta, "F1", 8, hLeft, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 5, 80, 15, 415, "Cuentas afiliadas a la tarjeta", "F2", 10, hLeft, vMiddle, , , , , 3

        'tabla
        oPDF.WTextBox nTopX + nHeight * 7, 50, 15, 100, "N° de Cuenta", "F1", 9, hCenter, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 7, 152, 15, 140, "Tipo", "F1", 9, hCenter, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 7, 294, 15, 60, "Moneda", "F1", 9, hCenter, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 7, 356, 15, 100, "Orden de Prioridad", "F1", 9, hCenter, vMiddle, , 1, , , 3
        oPDF.WTextBox nTopX + nHeight * 7, 458, 15, 100, "Tipo de cuenta", "F1", 9, hCenter, vMiddle, , 1, , , 3

        Dim nTopIniFila, nFilasPrimerHoja, nFilaXHoja, i, nFila As Integer
        Dim bPrimeraHoja As Boolean
        bPrimeraHoja = False
        nTopIniFila = nTopX + nHeight * 8 '
        nFilasPrimerHoja = 20
        nFilaXHoja = 40
        nFila = 1
        Do While Not rDatosAfiliacion.EOF

            oPDF.WTextBox nTopIniFila, 50, 15, 100, rDatosAfiliacion!cCtaCod, "F1", 9, hCenter, vMiddle, , , , , 3
            oPDF.WTextBox nTopIniFila, 152, 15, 140, rDatosAfiliacion!TProducto, "F1", 9, hCenter, vMiddle, , , , , 3
            oPDF.WTextBox nTopIniFila, 294, 15, 60, rDatosAfiliacion!Moneda, "F1", 9, hCenter, vMiddle, , , , , 3
            oPDF.WTextBox nTopIniFila, 356, 15, 100, rDatosAfiliacion!Prioridad, "F1", 9, hCenter, vMiddle, , , , , 3
            oPDF.WTextBox nTopIniFila, 458, 15, 100, rDatosAfiliacion!TCuenta, "F1", 9, hCenter, vMiddle, , , , , 3

            If nTopIniFila > 760 Then
                oPDF.NewPage A4_Vertical
                nTopIniFila = 50
                nFila = 0
            Else
              nTopIniFila = nTopIniFila + 20
            End If

            nFila = nFila + 1
            rDatosAfiliacion.MoveNext
        Loop

        If nTopIniFila > 560 Then
                oPDF.NewPage A4_Vertical
                nTopIniFila = 50
                nFila = 0
        End If

        oPDF.WTextBox nTopIniFila + 30, 50, 50, 500, "En la fecha confirmo haber recibido en sobre cerrado LA TARJETA, y haber establecido mi clave secreta, asimismo declaro aceptar las " _
                                                        & "condiciones generales de uso de la tarjeta de Débito y haber recibido las instrucciones para el uso de la misma", "F1", 10, hjustify, , , , , , 3
        oPDF.WTextBox nTopIniFila + 85, 50, 15, 500, "Debo señalar que queda sin efecto cualquier solicitud realizada antes de esta fecha.", "F1", 10, hjustify, , , , , , 3
        'firma
        oPDF.WTextBox nTopIniFila + 200, 50, 15, 150, "----------------------------------------", "F1", 10, hCenter, , , , , , 3
        oPDF.WTextBox nTopIniFila + 220, 50, 15, 150, "Firma del Cliente.", "F1", 10, hCenter, , , , , , 3
        'tipo y n° de doc
        oPDF.WTextBox nTopIniFila + 200, 205, 15, 150, "---------------------------------------", "F1", 10, hCenter, , , , , , 3
        oPDF.WTextBox nTopIniFila + 220, 205, 15, 150, "Tipo y N° de Doc.", "F1", 10, hCenter, , , , , , 3
        'visto bueno
        oPDF.WTextBox nTopIniFila + 200, 360, 15, 150, "---------------------------------------", "F1", 10, hCenter, , , , , , 3
        oPDF.WTextBox nTopIniFila + 220, 360, 15, 150, "V° B°.", "F1", 10, hCenter, , , , , , 3


        oPDF.PDFClose
        oPDF.Show
        Set oPDF = Nothing
    Else
        MsgBox "No se pudo generar el contrato de afiliación, comuníquese a TI.", vbInformation + vbError, "Error"
    End If
End Sub
'
Private Sub cmdSalir_Click()
    Unload Me
End Sub
