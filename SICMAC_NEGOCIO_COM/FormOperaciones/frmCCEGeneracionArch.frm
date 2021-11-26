VERSION 5.00
Begin VB.Form frmCCEGeneracionArch 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Archivos para la CCE"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6600
   Icon            =   "frmCCEGeneracionArch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNulo 
      Caption         =   "Generar Nulo"
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   5280
      Width           =   975
   End
   Begin VB.Frame fraResultado 
      BackColor       =   &H80000016&
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2895
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   6255
      Begin VB.TextBox txtResultado 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.CommandButton cmdGenerarArchivo 
      Caption         =   "&Generar Archivo"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame FraSesion 
      BackColor       =   &H80000016&
      Caption         =   "Sesión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   6255
      Begin VB.ComboBox cboSesion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblCargaSeCheque 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblSeCheque 
         BackColor       =   &H80000016&
         Caption         =   "Cheques "
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
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCargaAplicacion 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3720
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblaplicacion 
         BackColor       =   &H80000016&
         Caption         =   "Aplicación "
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
         Left            =   2520
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblsesion 
         BackColor       =   &H80000016&
         Caption         =   "Sesión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   5280
      Width           =   975
   End
   Begin VB.Frame FraInstrumento 
      BackColor       =   &H80000016&
      Caption         =   " Instrumentos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Timer Timer1 
         Left            =   3960
         Top             =   120
      End
      Begin VB.ComboBox cboInstrumento 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblCargaHora 
         BackColor       =   &H80000016&
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCargaFecha 
         BackColor       =   &H80000016&
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblfecha 
         BackColor       =   &H80000016&
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblTipoSesion 
         BackColor       =   &H80000016&
         Caption         =   "Instrumento "
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdProTra 
      Caption         =   "Procesar Transferencia "
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdProChe 
      Caption         =   " Procesar Cheque "
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   5280
      Width           =   1815
   End
End
Attribute VB_Name = "frmCCEGeneracionArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmCCEGeneracionArch
'** Descripción : Para la Generación de Archivos de Cheques , Proyecto: Implementacion del Servicio de Compensaciòn Electrónica Diferido de Instrumentos Compensables CCE
'** Creación : VAPA, 20160722
'**********************************************************************
Option Explicit
Dim oCCE As COMNCajaGeneral.NCOMCCE
'Dim oCamArchivo As COMNCajaGeneral.NCOMCCE
'Dim oCamPTra As COMNCajaGeneral.NCOMCCE
'Dim oCamPChe As COMNCajaGeneral.NCOMCCE
'Dim oCamNul As COMNCajaGeneral.NCOMCCE
Dim bLogicoDev As Boolean
Dim bLogicoCon As Boolean
Private Sub cboInstrumento_Click()
Dim rs As ADODB.Recordset
Dim rsAj As ADODB.Recordset
    
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    cboSesion.Clear
    lblCargaAplicacion.Caption = ""
    
   'cboSesion.Enabled = IIf(Left(cboInstrumento.Text, 1) = gCCECheque, False, True)
    
   cboSesion.Enabled = True
   
    If Left(cboInstrumento.Text, 1) = gCCECheque Then
        'Set rs = oCCE.CCE_ObtieneAplicacionSesionCheques(Left(cboInstrumento.Text, 1))
        Set rsAj = oCCE.CCE_ObtieneSesionChe(Left(cboInstrumento.Text, 1))
         Do While Not rsAj.EOF
            cboSesion.AddItem rsAj!cDesSesion & Space(200) & rsAj!nIdSesion
            rsAj.MoveNext
        Loop
    End If
    If Left(cboInstrumento.Text, 1) = gCCETransferencia Then
        Set rs = oCCE.CCE_ObtieneSesionxIdInstrumento(Left(cboInstrumento.Text, 1))
'        If rs.EOF And rs.BOF Then
'            cboSesion.Enabled = False
'            Exit Sub
'        End If
        Do While Not rs.EOF
            lblCargaSeCheque.Caption = ""
            cboSesion.AddItem rs!cDesSesion & Space(200) & rs!nIdSesion
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub LimpiaDatos()
    cboInstrumento.ListIndex = -1
    cboSesion.ListIndex = -1
    lblCargaSeCheque.Caption = ""
    lblCargaAplicacion.Caption = ""
    cboSesion.Enabled = False
End Sub
Private Sub cboSesion_Click()
    Dim rsAplicacion As ADODB.Recordset
    Dim rsAChe As ADODB.Recordset
    Set rsAplicacion = oCCE.CCE_ObtieneAplicacionTransferencia(Right(Me.cboSesion.Text, 1))
    Set rsAChe = oCCE.CCE_ObtieneAplicacionSesionCheques(Left(cboInstrumento.Text, 1))
'
'    'If Right(cboSesion.Text, 1) = 2 And bLogicoDev = False Then
'
'            MsgBox "No hay Transaciones para Generar Archivos de Devolución", vbInformation, "Aviso"
'            cboSesion.SetFocus
'            Exit Sub
'    End If
'    If Right(cboSesion.Text, 1) = 3 And bLogicoCon = False Then
'            MsgBox "No hay Transaciones para Generar Archivos de Confirmacion", vbInformation, "Aviso"
'            cboSesion.SetFocus
'            Exit Sub
'    End If
    
    If Not (rsAplicacion.EOF And rsAplicacion.BOF) Then
       lblCargaAplicacion.Caption = rsAplicacion!cCodAplicacion
       lblCargaAplicacion.BackColor = &H80000005
    End If
    
    If (rsAplicacion.EOF And rsAplicacion.BOF) Then
                                If (rsAChe.EOF And rsAChe.BOF) Then
                                        MsgBox "Debe esperar la Inicialización de la apertura de la Ventana Horaria para realizar la Generación del Archivo  ", vbInformation, "Aviso"
                                        cboSesion.Enabled = False
                                        lblCargaAplicacion.BackColor = &H80000003
                                        lblCargaAplicacion.Caption = ""
                                        cmdGenerarArchivo.Enabled = False
                                        Call CleanAll 'VAPA20170328
                                  End If
    End If
     
     If Right(cboSesion.Text, 1) = 4 Then
     
                    If Not (rsAChe.EOF And rsAChe.BOF) Then
                            lblCargaAplicacion.Caption = rsAChe!cCodAplicacion
                            lblCargaSeCheque.Caption = rsAChe!cDesSesion
                                If Left(lblCargaSeCheque.Caption, 1) = "R" Then
                                    MsgBox "Debe esperar la Inicialización de la apertura de la Ventana Horaria para realizar la Generación del Archivo de Presentados   ", vbInformation, "Aviso"
                                    cboSesion.Enabled = False
                                    lblCargaAplicacion.BackColor = &H80000003
                                    lblCargaAplicacion.Caption = ""
                                    cmdGenerarArchivo.Enabled = False
                                    Call CleanAll 'VAPA20170328
                                End If
                    End If
        End If
     If Right(cboSesion.Text, 1) = 5 Or Right(cboSesion.Text, 1) = 7 Then
                   If Not (rsAChe.EOF And rsAChe.BOF) Then
                            lblCargaAplicacion.Caption = rsAChe!cCodAplicacion
                            lblCargaSeCheque.Caption = rsAChe!cDesSesion
                            If Left(lblCargaSeCheque.Caption, 1) = "P" Then
                                    MsgBox "Debe esperar la Inicialización de la apertura de la Ventana Horaria para realizar la Generación del Archivo de Presentados   ", vbInformation, "Aviso"
                                    cboSesion.Enabled = False
                                    lblCargaAplicacion.BackColor = &H80000003
                                    lblCargaAplicacion.Caption = ""
                                    cmdGenerarArchivo.Enabled = False
                                    Call CleanAll 'VAPA20170328
                             End If
                    End If
    End If
'    If Right(cboSesion.Text, 1) = 3 Then
'     '  cmdGenerarArchivo.Enabled = True
'       lblCargaAplicacion.Caption = rsAplicacion!cCodAplicacion
'       lblCargaAplicacion.BackColor = &H80000005
'
'    End If
    
    
    
    If Right(cboSesion.Text, 1) = 6 Then
                   If Not (rsAChe.EOF And rsAChe.BOF) Then
                            lblCargaAplicacion.Caption = rsAChe!cCodAplicacion
                            lblCargaSeCheque.Caption = rsAChe!cDesSesion
                    End If
    End If
    
    
End Sub
Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    If cboInstrumento.ListIndex = -1 Then
        MsgBox "Debe Elejir un instrumento para la Generación de Archivo ", vbInformation, "Aviso"
        cboInstrumento.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If cboSesion.ListIndex = -1 And cboInstrumento.ListIndex = 0 Then
        MsgBox "Debe Elejir una Sesión de TransFerencias para la Generación de Archivo ", vbInformation, "Aviso"
        cboSesion.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'VAPA20170328
    If cboSesion.ListIndex = -1 And cboInstrumento.ListIndex = 1 Then
        MsgBox "Debe Elejir una Sesión de Cheques para la Generación de Archivo ", vbInformation, "Aviso"
        cboSesion.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'END VAPA
End Function
Private Sub cmdGenerarArchivo_Click()
    Dim oCamArchivo As COMNCajaGeneral.NCOMCCE
    Dim rsTrama As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim lsCargaCheque As Integer
    Dim fs As Scripting.FileSystemObject
    Dim tsArchivo As TextStream
    Dim lsArchivo As String
    Dim lsSesionChe As Integer
    Dim Carp As String
    Dim bLogicoTrama As Boolean
    Dim lbArch As Boolean
    
    Set oCamArchivo = New COMNCajaGeneral.NCOMCCE
    
    If Not ValidaDatos Then Exit Sub
    
    On Error GoTo ErrorGenArch
    oCamArchivo.BeginTrans
    lbArch = True
        
        bLogicoTrama = False
              
                    If Left(cboInstrumento.Text, 1) = gCCETransferencia Then
                            Select Case Right(cboSesion.Text, 1)
                                Case TRAPresentados
                                    Set rsTrama = oCamArchivo.CCE_GeneraTrama_TRAPRE(Right(cboInstrumento.Text, 3), Left(cboInstrumento.Text, 1), Right(cboSesion.Text, 1), Trim(lblCargaAplicacion.Caption), gdFecSis, gsCodUser)
                                    lsArchivo = App.Path & "\spooler\" & Replace("ECXXXPE.txt", "XXX", Trim(lblCargaAplicacion.Caption))
                                    bLogicoTrama = True
                                Case TRADevoluciones
                                    Set rsTrama = oCamArchivo.CCE_GeneraTrama_Devolucion(Right(cboInstrumento.Text, 3), Left(cboInstrumento.Text, 1), Right(cboSesion.Text, 1), Trim(lblCargaAplicacion.Caption), gdFecSis, gsCodUser)
                                    lsArchivo = App.Path & "\spooler\" & Replace("ECXXXDE.txt", "XXX", Trim(lblCargaAplicacion.Caption))
                                    bLogicoTrama = True
                                Case TRAAbono
                                    Set rsTrama = oCamArchivo.CCE_GeneraTrama_Confirmacion(Right(cboInstrumento.Text, 3), Left(cboInstrumento.Text, 1), 7, Trim(lblCargaAplicacion.Caption), gdFecSis, gsCodUser)
                                    lsArchivo = App.Path & "\spooler\" & Replace("ECXXXAE.txt", "XXX", Trim(lblCargaAplicacion.Caption))
                                    bLogicoTrama = True
                            End Select
                    End If
                
                    If Left(cboInstrumento.Text, 1) = gCCECheque Then
                        Set rs = oCamArchivo.CCE_ObtieneAplicacionSesionCheques(Left(cboInstrumento.Text, 1))
                        lsCargaCheque = rs!nTipoSesion
                            Select Case lsCargaCheque
                                Case CHEPresentados
                                     Set rsTrama = oCamArchivo.CCE_GeneraTrama_CHEPRE(Right(cboInstrumento.Text, 3), Left(cboInstrumento.Text, 1), lsCargaCheque, Trim(lblCargaAplicacion.Caption), gdFecSis, gsCodUser)
                                     lsArchivo = App.Path & "\spooler\" & Replace("ECCXXXE.txt", "XXX", "PRE")
                                     bLogicoTrama = True
                                Case CHERechazados
                                     Set rsTrama = oCamArchivo.CCE_GeneraTrama_CHEPRE(Right(cboInstrumento.Text, 3), Left(cboInstrumento.Text, 1), lsCargaCheque, Trim(lblCargaAplicacion.Caption), gdFecSis, gsCodUser)
                                     lsArchivo = App.Path & "\spooler\" & Replace("ECCXXXE.txt", "XXX", "REC")
                                     bLogicoTrama = True
                                Case 3
                                    'aca hay ajustes
                            End Select
                    End If
                    
                'If RSVacio(rsTrama) Then 'comentado por vapa 20161019
                If Not bLogicoTrama Then
                    MsgBox "No existe información para generar el archivo de envío a la Cámara.", vbInformation, "Aviso"
                    oCamArchivo.CommitTrans 'VAPA20170328
                    Exit Sub
                End If
                Set fs = New Scripting.FileSystemObject
                Set tsArchivo = fs.CreateTextFile(lsArchivo, True)
                If Not rsTrama.EOF Then 'VAPA20170328
                Do While Not rsTrama.EOF
                        tsArchivo.WriteLine rsTrama!cDato
                        rsTrama.MoveNext
                Loop
                Else
                    MsgBox "No existe Registros para generar el archivo de envío a la Cámara.", vbInformation, "Aviso"
                    oCamArchivo.CommitTrans 'VAPA20170328
                    Exit Sub
                End If 'VAPA20170328
                'tsArchivo.WriteLine ""
                'tsArchivo.WriteBlankLines (1)
                tsArchivo.Close
                If MsgBox("Proceso terminado correctamente. El archivo ha sido generado exitosamente" & Chr(13) & " en " & App.Path & "\SPOOLER\" & Chr(13) & "¿Desea abrir el archivo?", vbQuestion + vbYesNo + vbDefaultButton1, "Aviso!!!") = vbNo Then
                    oCamArchivo.CommitTrans 'VAPA20170328
                    Exit Sub
                Else
                    Carp = "notepad " & lsArchivo
                    Shell Carp, vbMaximizedFocus
                End If
                
                
    oCamArchivo.CommitTrans
    lbArch = False
    LimpiaDatos
    Set oCamArchivo = Nothing
    Exit Sub

ErrorGenArch:
    If lbArch Then
        oCamArchivo.RollbackTrans
        Set oCamArchivo = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub cmdLimpiar_Click()
    txtResultado.Text = ""
    cboInstrumento.ListIndex = -1
    cboSesion.ListIndex = -1
    lblCargaSeCheque.Caption = ""
    lblCargaAplicacion.Caption = ""
    cboSesion.Enabled = False
    cmdGenerarArchivo.Enabled = True 'VAPA20170328
End Sub
Private Sub CleanAll()
txtResultado.Text = ""
    cboInstrumento.ListIndex = -1
    cboSesion.ListIndex = -1
    lblCargaSeCheque.Caption = ""
    lblCargaAplicacion.Caption = ""
    cboSesion.Enabled = False
End Sub

Private Sub cmdNulo_Click()
Dim oCamNul As COMNCajaGeneral.NCOMCCE
Dim rsTrama As ADODB.Recordset
Dim rsValidacion As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim lsCargaCheque As Integer
    Dim fs As Scripting.FileSystemObject
    Dim tsArchivo As TextStream
    Dim lsArchivo As String
    Dim lsSesionChe As Integer
    Dim Carp As String
    Dim bLogicoTrama As Boolean
    Dim lsCabecera As String
    Dim lsNul As String
    Dim nResult As Integer
    Dim lsSesion As String
    Dim lbArch As Boolean
    Dim lsCodAplica As String
    

    Set oCamNul = New COMNCajaGeneral.NCOMCCE
    
    If Not ValidaDatos Then Exit Sub
    

On Error GoTo ErrorGenNul
 oCamNul.BeginTrans

    lbArch = True
    bLogicoTrama = False
    lsNul = "NUL"
    lsCodAplica = Trim(lblCargaAplicacion.Caption)
    
    
                                    If Left(cboInstrumento.Text, 1) = gCCETransferencia Then
                                                lsSesion = Right(cboSesion.Text, 1)
                                                 Set rsValidacion = oCamNul.CCE_ValidaNulo_Envio_Tra(lsSesion, lsNul, gdFecSis, Trim(lblCargaAplicacion.Caption))
                                                       nResult = rsValidacion!nResult
                                                 If nResult = 0 Then
                                                Select Case lsSesion
                                                                Case "1"
                                                                   Set rsTrama = oCamNul.CCE_GeneraTrama_Nulo_Tra(Right(cboSesion.Text, 1), gdFecSis, Trim(lblCargaAplicacion.Caption), "TRA", "2", gsCodUser)
                                                                   lsArchivo = App.Path & "\spooler\" & Replace("ECXXXPE.txt", "XXX", Trim(lblCargaAplicacion.Caption))
                                                                   bLogicoTrama = True
                                                                Case "2"
                                                                   Set rsTrama = oCamNul.CCE_GeneraTrama_Nulo_Tra(Right(cboSesion.Text, 1), gdFecSis, Trim(lblCargaAplicacion.Caption), "TRA", "2", gsCodUser)
                                                                   lsArchivo = App.Path & "\spooler\" & Replace("ECXXXDE.txt", "XXX", Trim(lblCargaAplicacion.Caption))
                                                                   bLogicoTrama = True
                                                                  Case "3"
                                                                   Set rsTrama = oCamNul.CCE_GeneraTrama_Nulo_Tra(Right(cboSesion.Text, 1), gdFecSis, Trim(lblCargaAplicacion.Caption), "TRA", "2", gsCodUser)
                                                                   lsArchivo = App.Path & "\spooler\" & Replace("ECXXXAE.txt", "XXX", Trim(lblCargaAplicacion.Caption))
                                                                   bLogicoTrama = True
                                                End Select
                                              Else
                                                    MsgBox "Ya se Genero el Archivo Nulo de Transferencia .", vbInformation, "Aviso"
                                                    oCamNul.CommitTrans 'VAPA20170328
                                                    Exit Sub
                                              End If
                                    End If
                        
                                    If Left(cboInstrumento.Text, 1) = gCCECheque Then
                                    Set rs = oCCE.CCE_ObtieneAplicacionSesionCheques(Left(cboInstrumento.Text, 1))
                                    lsCargaCheque = rs!nTipoSesion
                                    
                                                'Set rsValidacion = oCamNul.CCE_ValidaNulo_Envio_Che(lsCargaCheque, lsNul, gdFecSis)
                                                Set rsValidacion = oCamNul.CCE_ValidaNulo_Envio_Che(lsCodAplica, gdFecSis)
                                                
                                                nResult = rsValidacion!nResult
                                                If nResult = 0 Then
                                                        Set rsTrama = oCamNul.CCE_GeneraTrama_Nulo_Che(lsCargaCheque, gdFecSis, Trim(lblCargaAplicacion.Caption), "2", gsCodUser)
                                                Else
                                                        MsgBox "Ya se Genero el Archivo Nulo de Cheques .", vbInformation, "Aviso"
                                                        oCamNul.CommitTrans 'VAPA20170328
                                                        Exit Sub
                                                End If
                                                If Not lsCargaCheque = 2 Then
                                                        lsArchivo = App.Path & "\spooler\" & Replace("ECCXXXE.txt", "XXX", "PRE")
                                                        bLogicoTrama = True
                                                Else
                                                        lsArchivo = App.Path & "\spooler\" & Replace("ECCXXXE.txt", "XXX", "REC")
                                                        bLogicoTrama = True
                                                End If
                                        
                                    End If
                            
                                    If Not bLogicoTrama Then
                                                MsgBox "No existe información para generar el archivo de envío a la Cámara.", vbInformation, "Aviso"
                                                oCamNul.CommitTrans 'VAPA20170328
                                                Exit Sub
                                    End If
                                    Set fs = New Scripting.FileSystemObject
                                    Set tsArchivo = fs.CreateTextFile(lsArchivo, True)
                                    Do While Not rsTrama.EOF
                                            tsArchivo.WriteLine rsTrama!cDato
                                            rsTrama.MoveNext
                                    Loop
                                    tsArchivo.Close
                
                        If MsgBox("Proceso terminado correctamente. El archivo Nulo ha sido generado exitosamente" & Chr(13) & " en " & App.Path & "\SPOOLER\" & Chr(13) & "¿Desea abrir el archivo?", vbQuestion + vbYesNo + vbDefaultButton1, "Aviso!!!") = vbNo Then
                                oCamNul.CommitTrans 'VAPA20170328
                                Exit Sub
                        Else
                            Carp = "notepad " & lsArchivo
                            Shell Carp, vbMaximizedFocus
                        End If
                        
oCamNul.CommitTrans
lbArch = False
LimpiaDatos
Set oCamNul = Nothing
Exit Sub


ErrorGenNul:
    If lbArch Then
        oCamNul.RollbackTrans
        Set oCamNul = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Sub cmdProChe_Click()
    Dim oCamPChe As COMNCajaGeneral.NCOMCCE
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME, lsSecUnivoca As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivocaS, lsSecUnivocaD, lsTipTra, lsCabNullMN, lsCabNullME   As String
    Dim lbArch As Boolean
    Dim lsCtaDebitarE As String
    Dim lsNroChequeE As String
    Dim lsIfiCodE As String
    Dim lsCtaDebitarC As String
    Dim lsNroChequeC As String
    Dim lsIfiCodC As String
    Dim lsCtaDebitarR As String
    Dim lsNroChequeR As String
    Dim lsImporte As Double
    Dim lsImporteComision As Double
    Dim lsMismoTitular As Boolean
    Dim nSubTpo As Integer
    Dim rsChePre As ADODB.Recordset
    Dim rsCheRec As ADODB.Recordset
    Dim rsConfirmacion As ADODB.Recordset
    Dim rsVal As ADODB.Recordset
    Dim nEstado As Integer
    Dim nContDev As Integer
    Dim nContCon As Integer
    Dim sFechaAnterior As Date
    Dim contCon As Integer
    Dim contRec As Integer
    Dim bLogicoBlog As Boolean
    Dim lsResult As String
    'fecha = DateAdd(DateInterval.Day, -1, fecha)
    'Dim lbArch As Boolean
    Dim lnPresentados As Integer 'VAPA20170328
    Dim lnRechazados As Integer 'VAPA20170328
    
     lnPresentados = 4
     lnRechazados = 7 'VAPA20170707 SE CAMBIO A 7
     If Not ValidaDatos Then Exit Sub 'VAPA20170328
    Set oCamPChe = New COMNCajaGeneral.NCOMCCE
    
 On Error GoTo ErrorProChe
    oCamPChe.BeginTrans
    lbArch = True
    'VAPA COMBO
    If Left(cboInstrumento.Text, 1) = gCCETransferencia Then
           Select Case Right(cboSesion.Text, 1)
                 Case TRAAbono
                    
                                If Trim(lblCargaAplicacion.Caption) = "TRT" Then
                                        txtResultado.Text = ""
                                        MsgBox "No se puede procesar  Transeferencia en Cheques ", vbInformation, "Aviso"
                                        oCamPChe.CommitTrans
                                        Exit Sub
                                 End If
                                 If Trim(lblCargaAplicacion.Caption) = "TRI" Or Trim(lblCargaAplicacion.Caption) = "TRM" Then
                                        txtResultado.Text = ""
                                        MsgBox "No se puede procesar  Transeferencia en Cheques ", vbInformation, "Aviso"
                                        oCamPChe.CommitTrans
                                        Exit Sub
                                 End If
                     Case TRADevoluciones
                                 MsgBox "No se puede procesar  Transeferencia en Cheques ", vbInformation, "Aviso"
                                 oCamPChe.CommitTrans
                                  Exit Sub
                     Case TRAPresentados
                                  MsgBox "No se puede procesar Transeferencias en Cheques ", vbInformation, "Aviso"
                                  oCamPChe.CommitTrans
                                  Exit Sub
             End Select
        End If
                                     
     
         If Left(cboInstrumento.Text, 1) = gCCECheque Then
          Select Case Right(cboSesion.Text, 1)
                    Case lnPresentados
                    MsgBox "No Existe Procesamiento de Cheques en la sesión de Presentados  ", vbInformation, "Aviso"
                    oCamPChe.CommitTrans
                    Exit Sub
                    Case lnRechazados
                            txtResultado.Text = ""
                            sFechaAnterior = DateAdd("d", -1, gdFecSis)
            
            
                            Set rsChePre = oCamPChe.CCE_SeleccionaPresentados(sFechaAnterior)
                            Set rsCheRec = oCamPChe.CCE_SeleccionaRechazados(gdFecSis)
                            Set rsConfirmacion = oCamPChe.CCE_SeleccionaChequesConfirmados(gdFecSis)
                            contCon = 0
                            contRec = 0
               End Select
                    
        End If
     
    'END VAPA
    
    
'    txtResultado.Text = ""
'    sFechaAnterior = DateAdd("d", -1, gdFecSis)
'
'
'    Set rsChePre = oCamPChe.CCE_SeleccionaPresentados(sFechaAnterior)
'    Set rsCheRec = oCamPChe.CCE_SeleccionaRechazados(gdFecSis)
'    Set rsConfirmacion = oCamPChe.CCE_SeleccionaChequesConfirmados(gdFecSis)
'    contCon = 0
'    contRec = 0
   
    
'Rechazos
    If Not rsCheRec.EOF Then
        Do While Not rsCheRec.EOF
                lsCtaDebitarR = rsCheRec!cCtaDebitar
                lsNroChequeR = rsCheRec!NroCheque
                rsChePre.MoveFirst
                Do While Not rsChePre.EOF
                           lsCtaDebitarE = rsChePre!cCtaDebitar
                           lsNroChequeE = rsChePre!NroCheque
                           lsIfiCodE = rsChePre!IfiCod
                           If lsCtaDebitarE = lsCtaDebitarR And lsNroChequeE = lsNroChequeR Then
                           txtResultado.Text = txtResultado.Text & "Se Rechazo la Cuenta  " & lsCtaDebitarR & "   con Numero de Cheque " & lsNroChequeR & vbCrLf
                           contRec = contRec + 1
                            End If
                        rsChePre.MoveNext
                Loop
                rsCheRec.MoveNext
        Loop
    Else
                If Not rsChePre.EOF Then
                          rsChePre.MoveFirst
                                     Do While Not rsChePre.EOF
                                                         lsCtaDebitarE = rsChePre!cCtaDebitar
                                                         lsNroChequeE = rsChePre!NroCheque
                                                         lsIfiCodE = rsChePre!IfiCod
                                                         
                                                         Set rsVal = oCamPChe.CCE_ValidaProCheque(lsCtaDebitarE, lsIfiCodE, lsNroChequeE, gdFecSis)
                                                         lsResult = rsVal!cVal
                                                            If lsResult = "no" Then
                                                                oCamPChe.CCE_InsertaLogProCheque gdFecSis, gsCodAge, lsIfiCodE, lsCtaDebitarE, lsNroChequeE
                                                                oCamPChe.CCE_Confirma_CHE gdFecSis, gsCodAge, lsIfiCodE, lsCtaDebitarE, lsNroChequeE
                                                                'oCamPChe.CCE_InsertaLogProCheque gdFecSis, gsCodAge, lsIfiCodE, lsCtaDebitarE, lsNroChequeE
                                                                contCon = contCon + 1
                                                                rsChePre.MoveNext
                                                            Else
                                                                     MsgBox "Ya se realizo el Proceso de Aplicacion del Cheque   " & lsNroChequeE & " ""   de la Cuenta :  " & lsCtaDebitarE & " ", vbInformation, "Aviso"
                                                                     rsChePre.MoveNext
                                                                     
                                                            End If
                                                         
'                                                         oCamPChe.CCE_Confirma_CHE gdFecSis, gsCodAge, lsIfiCodE, lsCtaDebitarE, lsNroChequeE
'                                                         oCamPChe.CCE_InsertaLogProCheque gdFecSis, gsCodAge, lsIfiCodE, lsCtaDebitarE, lsNroChequeE
'                                                          rsChePre.MoveNext
                                     Loop
                Else
                MsgBox "No se encontraron Cheques por Procesar ", vbInformation, "Aviso"
                oCamPChe.CommitTrans
                Exit Sub
           End If
    End If
    
'        'confirmados
    If contRec > 0 Then
   
                Do While Not rsConfirmacion.EOF
                                lsCtaDebitarC = rsConfirmacion!cCtaDebitar
                                lsNroChequeC = rsConfirmacion!NroCheque
                                lsIfiCodC = rsConfirmacion!IfiCod
                                'LogProceso
                                 Set rsVal = oCamPChe.CCE_ValidaProCheque(lsCtaDebitarC, lsIfiCodC, lsNroChequeC, gdFecSis)
                                 lsResult = rsVal!cVal
                                
                                If lsResult = "no" Then
                                            oCamPChe.CCE_InsertaLogProCheque gdFecSis, gsCodAge, lsIfiCodC, lsCtaDebitarC, lsNroChequeC
                                            oCamPChe.CCE_Confirma_CHE gdFecSis, gsCodAge, lsIfiCodC, lsCtaDebitarC, lsNroChequeC
                                            'oCamPChe.CCE_InsertaLogProCheque gdFecSis, gsCodAge, lsIfiCodC, lsCtaDebitarC, lsNroChequeC
                                            contCon = contCon + 1
                                            rsConfirmacion.MoveNext
                                Else
                                            MsgBox "Ya se realizo el Proceso de Aplicacion del Cheque   " & lsNroChequeC & " ""   de la Cuenta :  " & lsCtaDebitarC & " ", vbInformation, "Aviso"
                                            rsConfirmacion.MoveNext
                                End If
'                                oCamPChe.CCE_Confirma_CHE gdFecSis, gsCodAge, lsIfiCodC, lsCtaDebitarC, lsNroChequeC
'                                oCamPChe.CCE_InsertaLogProCheque gdFecSis, gsCodAge, lsIfiCodC, lsCtaDebitarC, lsNroChequeC
                            
                                
'                                rsConfirmacion.MoveNext
'                                contCon = contCon + 1
                Loop
    End If
'    'end confirmados
    

                                txtResultado.Text = txtResultado.Text & "RECHAZADOS  " & contRec & "   con Numero de Cheque " & lsNroChequeE & vbCrLf
                                txtResultado.Text = txtResultado.Text & "CONFIRMADOS  " & contCon & "   con Numero de Cheque " & lsNroChequeE & vbCrLf
                                                        
    
oCamPChe.CommitTrans
lbArch = False
Set oCamPChe = Nothing
Exit Sub

    
ErrorProChe:
    If lbArch Then
        oCamPChe.RollbackTrans
        Set oCamPChe = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub cmdProTra_Click()
    Dim oCamPTra As COMNCajaGeneral.NCOMCCE
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME, lsSecUnivoca As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivocaS, lsSecUnivocaD, lsTipTra, lsCabNullMN, lsCabNullME   As String
    Dim lbArch As Boolean
    Dim lsCtaCodCci As String
    Dim lsImporte As Double
    Dim lsImporteComision As Double
    Dim lsMismoTitular As Boolean
    Dim nSubTpo As Integer
    Dim rsAplicaTra As ADODB.Recordset
    Dim nEstado As Integer
    Dim nContDev As Integer
    Dim nContCon As Integer
    Dim lsLDeposito As Long
    'Dim lsId As Integer
    Dim lsId As Long
    Dim lbArchP As Boolean
    Dim sFechaAnterior As Date
    sFechaAnterior = DateAdd("d", -1, gdFecSis)
    
    
        
    If Not ValidaDatos Then Exit Sub 'VAPA20170328
    Set oCamPTra = New COMNCajaGeneral.NCOMCCE
    
On Error GoTo ErrorProTra
oCamPTra.BeginTrans
lbArchP = True
  
    
       If Left(cboInstrumento.Text, 1) = gCCETransferencia Then
           Select Case Right(cboSesion.Text, 1)
                 Case 6 'TRAAbono
                    'Trim(lblCargaAplicacion.Caption), gdFecSis, gsCodUser)
                                If Trim(lblCargaAplicacion.Caption) = "TRT" Then
                                        txtResultado.Text = ""
                                        Set rsAplicaTra = oCamPTra.CCE_AplicaTra(Trim(lblCargaAplicacion.Caption), sFechaAnterior)
                                         nContDev = 0
                                         nContCon = 0
                                 End If
                                 If Trim(lblCargaAplicacion.Caption) = "TRI" Or Trim(lblCargaAplicacion.Caption) = "TRM" Then
                                        txtResultado.Text = ""
                                        Set rsAplicaTra = oCamPTra.CCE_AplicaTra(Trim(lblCargaAplicacion.Caption), gdFecSis)
                                         nContDev = 0
                                         nContCon = 0
                                 End If
                     Case TRADevoluciones
                                 MsgBox "No se puede procesar  Transeferencia en Devoluciones ", vbInformation, "Aviso"
                                 oCamPTra.CommitTrans
                                  Exit Sub
                     Case TRAPresentados
                                  MsgBox "No se puede procesar Transeferencias en Presentados ", vbInformation, "Aviso"
                                  oCamPTra.CommitTrans
                                  Exit Sub
             End Select
        End If
                                     
     'V
         If Left(cboInstrumento.Text, 1) = gCCECheque Then
                    MsgBox "No se puede procesar Cheques en el Procesamiento de transeferencias ", vbInformation, "Aviso"
                    oCamPTra.CommitTrans
                    Exit Sub
        End If
     'V
     
 If Not rsAplicaTra.EOF Then
        Do While Not rsAplicaTra.EOF
     
                                lsTipTra = rsAplicaTra!cTipTrans
                                nEstado = rsAplicaTra!nEstado
                                lsCtaCodCci = rsAplicaTra!cCtaCCi
                                lsImporte = rsAplicaTra!Importe
                                sMoneda = rsAplicaTra!Moneda
                                lsImporteComision = rsAplicaTra!ImporteComision
                                lsMismoTitular = rsAplicaTra!MismoTitular
                                lsSecUnivoca = rsAplicaTra!SecUnivoca
                                lsId = rsAplicaTra!IdVal
                   If nEstado = 23 Then
                                Select Case lsTipTra
                                                  Case "220"
                                                          
                                                              lsLDeposito = oCamPTra.CCE_DepositoOrd(gdFecSis, lsCtaCodCci, lsImporte, sMoneda, lsImporteComision, lsMismoTitular, lsSecUnivoca)
                                                              bLogicoCon = True
                                                              nContCon = nContCon + 1
                                                              If Not lsLDeposito = 0 Then
                                                                oCamPTra.CCE_CondicionTra lsId, 24
                                                              Else
                                                                oCamPTra.CCE_CondicionTra lsId, 23
                                                              End If
                                                  Case "221"
                                                              oCamPTra.CCE_PagoHaberes gdFecSis, lsCtaCodCci, lsImporte, sMoneda, lsImporteComision, IIf(Left(Data, 1) = "6", nSubTpo = Mid(Data, 177, 1), 0), lsSecUnivoca
                                                  Case "222"
                                                              oCamPTra.CCE_PagoProveedores gdFecSis, lsCtaCodCci, lsImporte, sMoneda, lsImporteComision, lsSecUnivoca
                                                  Case "223"
                                                            'oCamara.CCE_DepositoCTS gdFecSis, lsCtaCodCci, lsImporte, sMoneda, lsImporteComision, lsSecUnivoca
                                          End Select
    
                   Else
                                            nContDev = nContDev + 1
                                            bLogicoDev = True
                   End If
            rsAplicaTra.MoveNext
        Loop
    Else
    MsgBox "No se encontraron Transferencias por Procesar ", vbInformation, "Aviso"
    oCamPTra.CommitTrans
    Exit Sub
    End If
    
    txtResultado.Text = txtResultado.Text & "Se encontraron  " & nContDev & "   Transferencias Rechazados" & vbCrLf
    txtResultado.Text = txtResultado.Text & "Se encontraron  " & nContCon & "   Transferencias Confirmadas" & vbCrLf
    
oCamPTra.CommitTrans
lbArchP = False
Set oCamPTra = Nothing
Exit Sub

    
ErrorProTra:
    If lbArchP Then
        oCamPTra.RollbackTrans
        Set oCamPTra = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Dim rs As ADODB.Recordset
bLogicoDev = False
bLogicoCon = False
    lblCargaFecha.Caption = DateValue(Now)
    lblCargaHora.Caption = TimeValue(Now)
    Timer1.Interval = 1000
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    Set rs = oCCE.CCE_ObtieneInstrumento
    Do While Not rs.EOF
        cboInstrumento.AddItem Format(rs!nIdInstrumento, "0") & " " & Mid(rs!cNomInstrumento & Space(100), 1, 200) & rs!cCodInstrumento
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub Timer1_Timer()
Dim xseg As Integer
    xseg = xseg + 1
    lblCargaFecha.Caption = DateValue(Now)
    lblCargaHora.Caption = TimeValue(Now)
End Sub

