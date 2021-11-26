VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChequeNegocioValorizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VALORIZACIÓN DE CHEQUES"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
   Icon            =   "frmChequeNegocioValorizacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5120
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   9022
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cheques para Valorizar"
      TabPicture(0)   =   "frmChequeNegocioValorizacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feCheque"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkTodos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCerrar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRechazar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdValorizar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdActualizar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   345
         Left            =   8775
         TabIndex        =   6
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdValorizar 
         Caption         =   "&Valorizar"
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "&Rechazar"
         Height          =   345
         Left            =   1230
         TabIndex        =   4
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   345
         Left            =   9885
         TabIndex        =   3
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   435
         Width           =   855
      End
      Begin Sicmact.FlexEdit feCheque 
         Height          =   3930
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   10860
         _ExtentX        =   19156
         _ExtentY        =   6932
         Cols0           =   15
         FixedCols       =   0
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmChequeNegocioValorizacion.frx":0326
         EncabezadosAnchos=   "0-0-400-1800-2000-2300-1800-1200-1800-1800-1200-1200-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-L-L-L-L-C-L-L-L-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmChequeNegocioValorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'** Nombre : frmChequeNegocioValorizacion
'** Descripción : Para valorización de Cheques creado segun TI-ERS126-2013
'** Creación : EJVG, 20130125 12:00:00 AM
'*************************************************************************
Option Explicit
Dim fsopecod As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsopecod = psOpeCod
    Caption = UCase(psOpeDesc)
    Show 1
End Sub
Private Sub chkTodos_Click()
    Dim row As Long
    Dim lsCheck As String
    If feCheque.TextMatrix(1, 0) = "" Then
        chkTodos.value = 0
        Exit Sub
    End If
    If chkTodos.value = vbChecked Then
        lsCheck = "1"
    End If
    For row = 1 To feCheque.Rows - 1
        feCheque.TextMatrix(row, 2) = lsCheck
    Next
End Sub
Private Sub cmdActualizar_Click()
    carga_cheques
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Function carga_cheques() As Boolean
    Dim oDR As New DDocRec
    Dim oRs As New ADODB.Recordset
    Dim row As Long
    
    On Error GoTo ErrCarga
    Screen.MousePointer = 11
    Set oRs = oDR.ListaChequeNegxValorizacion(CInt(Mid(fsopecod, 3, 1)))
    FormateaFlex feCheque
    If Not oRs.EOF Then
        Do While Not oRs.EOF
            feCheque.AdicionaFila
            row = feCheque.row
            feCheque.TextMatrix(row, 1) = oRs!nId
            feCheque.TextMatrix(row, 3) = oRs!cIFiReceptorNombre
            feCheque.TextMatrix(row, 4) = oRs!cCtaIFDesc
            feCheque.TextMatrix(row, 5) = oRs!cNroDoc
            feCheque.TextMatrix(row, 6) = oRs!cAgeDescripcion
            feCheque.TextMatrix(row, 7) = oRs!dRegistro
            feCheque.TextMatrix(row, 8) = oRs!cGiradorNombre
            feCheque.TextMatrix(row, 9) = oRs!cIFiEmisorNombre
            feCheque.TextMatrix(row, 10) = oRs!cMoneda
            feCheque.TextMatrix(row, 11) = Format(oRs!nMonto, gsFormatoNumeroView)
            feCheque.TextMatrix(row, 12) = oRs!cDepIFTpo
            feCheque.TextMatrix(row, 13) = oRs!cDepPersCod
            feCheque.TextMatrix(row, 14) = oRs!cDepIFCta
            oRs.MoveNext
        Loop
        carga_cheques = True
        
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
    Else
        carga_cheques = False
        MsgBox "No se encontraron datos para realizar la operación", vbInformation, "Aviso"
    End If
    
    Set oRs = Nothing
    Set oDR = Nothing
    Screen.MousePointer = 0
    Exit Function
ErrCarga:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Private Sub cmdRechazar_Click()
    Dim oDR As NDocRec
    Dim oImp As NContImprimir
    Dim oOpe As DOperacion
    Dim bExito As Boolean
    Dim I As Long, row As Long
    Dim lsGlosa As String
    Dim MatDatos() As Variant
    Dim lsCtaContCodD As String, lsCtaContCodH As String
    Dim lsOpeCod As String
    Dim lsImpre As String, lsMovNroImpre As String
    Dim MovNro() As String
    Dim lsIFTpo As String, lsPersCod As String, lsCtaIFCod As String
    
    On Error GoTo ErrRechazar
    If Not Validar Then Exit Sub
    
    lsOpeCod = IIf(Mid(fsopecod, 3, 1) = "1", OpeCGOtrosOpeChequeRechazo, OpeCGOtrosOpeChequeRechazoME)
    Set oOpe = New DOperacion
    lsCtaContCodD = oOpe.EmiteOpeCta(lsOpeCod, "D", "0")
    If Not oOpe.ValidaCtaCont(lsCtaContCodD) Then
        MsgBox "La cuenta contable " & lsCtaContCodD & " para el DEBE no está a último nivel, consultar con el Dpto. de Contabilidad", vbInformation, "Aviso"
        Set oOpe = Nothing
        Exit Sub
    End If
    ReDim MatDatos(1 To 6, 0 To 0)
    For I = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(I, 2) = "." Then
            row = UBound(MatDatos, 2) + 1
            ReDim Preserve MatDatos(1 To 6, 0 To row)
            MatDatos(1, row) = CLng(feCheque.TextMatrix(I, 1))
            lsIFTpo = feCheque.TextMatrix(I, 12)
            lsPersCod = feCheque.TextMatrix(I, 13)
            lsCtaIFCod = feCheque.TextMatrix(I, 14)
            lsCtaContCodH = oOpe.EmiteOpeCta(lsOpeCod, "H", "0", lsIFTpo & "." & lsPersCod & "." & lsCtaIFCod, ObjEntidadesFinancieras)
            If Not oOpe.ValidaCtaCont(lsCtaContCodH) Then
                MsgBox "Al cheque Nro. " & feCheque.TextMatrix(I, 5) & " le falta Cuenta Contable HABER de rechazo del Cheque, coordinar con el Dpto. de TI", vbInformation, "Aviso"
                feCheque.TopRow = I
                Set oOpe = Nothing
                Exit Sub
            End If
            MatDatos(2, row) = lsCtaContCodD
            MatDatos(3, row) = lsCtaContCodH
            MatDatos(4, row) = lsIFTpo
            MatDatos(5, row) = lsPersCod
            MatDatos(6, row) = lsCtaIFCod
        End If
    Next
    Set oOpe = Nothing
    Select Case UBound(MatDatos, 2)
        Case 0
            MsgBox "Ud. debe seleccionar primero el Cheque para continuar", vbInformation, "Aviso"
            Exit Sub
        Case Is > 1
            MsgBox "Solo se puede rechazar un cheque a la vez, no se puede continuar", vbInformation, "Aviso"
            Exit Sub
    End Select
    
    lsGlosa = Trim(InputBox("Ingrese una Glosa para rechazar el cheque", "Rechazo de Cheque"))
    If Len(lsGlosa) = 0 Then Exit Sub
    
    If MsgBox("¿Esta seguro de realizar el rechazo del cheque?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    cmdRechazar.Enabled = False
    Set oDR = New NDocRec
    bExito = oDR.RechazaChequeNegocio(MatDatos, lsOpeCod, gdFecSis, Right(gsCodAge, 2), gsCodUser, lsGlosa, lsMovNroImpre)
    If bExito Then
        MsgBox "Se ha rechazado con éxito los cheques seleccionados", vbInformation, "Aviso"
        Set oImp = New NContImprimir
        MovNro = Split(lsMovNroImpre, ",")
        For I = 0 To UBound(MovNro)
            lsImpre = lsImpre & oImp.ImprimeAsientoContable(MovNro(I), gnLinPage, gnColPage, "RECHAZO DE CHEQUES DE NEGOCIO") & oImpresora.gPrnSaltoPagina
        Next
        EnviaPrevio lsImpre, "RECHAZO CHEQUES DE NEGOCIO", gnLinPage, False
        Set oImp = Nothing
        cmdActualizar_Click
    Else
        MsgBox "Ha sucedido un error al rechazar los cheques seleccionados, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Screen.MousePointer = 0
    cmdRechazar.Enabled = True
    Set oDR = Nothing
    Exit Sub
ErrRechazar:
    Screen.MousePointer = 0
    cmdRechazar.Enabled = True
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdValorizar_Click()
    Dim oDR As NDocRec
    Dim bExito As Boolean
    Dim I As Long, row As Long
    Dim MatDatos() As Variant

    On Error GoTo ErrValorizar
    If Not Validar Then Exit Sub

    ReDim MatDatos(1 To 1, 0 To 0)
    For I = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(I, 2) = "." Then
            row = UBound(MatDatos, 2) + 1
            ReDim Preserve MatDatos(1 To 1, 0 To row)
            MatDatos(1, row) = CLng(feCheque.TextMatrix(I, 1))
        End If
    Next
    If UBound(MatDatos, 2) = 0 Then
        MsgBox "Ud. debe seleccionar primero al menos un Cheque para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de valorizar los cheques seleccionados?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    cmdValorizar.Enabled = False
    Set oDR = New NDocRec
    bExito = oDR.ValorizaChequeNegocio(MatDatos, fsopecod, gdFecSis, Right(gsCodAge, 2), gsCodUser)
    If bExito Then
        MsgBox "Se ha valorizado con éxito los cheques seleccionados", vbInformation, "Aviso"
        cmdActualizar_Click
    Else
        MsgBox "Ha sucedido un error al valorizar los cheques seleccionados, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Screen.MousePointer = 0
    cmdValorizar.Enabled = True
    Set oDR = Nothing
    Exit Sub
ErrValorizar:
    Screen.MousePointer = 0
    cmdValorizar.Enabled = True
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feCheque_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(feCheque.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    cmdActualizar_Click
End Sub
Private Function Validar() As Boolean
    Dim row As Long
    Dim bSelecciona As Boolean
    Validar = True
    If feCheque.TextMatrix(1, 0) = "" Then
        MsgBox "No existen cheques para realizar la operación", vbInformation, "Aviso"
        Validar = False
        Exit Function
    End If
    For row = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(row, 2) = "." Then
            bSelecciona = True
            Exit For
        End If
    Next
    If Not bSelecciona Then
        MsgBox "Ud. debe seleccionar al menos un cheque para realizar la operación", vbInformation, "Aviso"
        Validar = False
        Exit Function
    End If
End Function
