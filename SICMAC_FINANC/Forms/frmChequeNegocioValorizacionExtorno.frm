VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChequeNegocioValorizacionExtorno 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11235
   Icon            =   "frmChequeNegocioValorizacionExtorno.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4980
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   8784
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cheques"
      TabPicture(0)   =   "frmChequeNegocioValorizacionExtorno.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feCheque"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExtornar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCerrar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdActualizar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   345
         Left            =   7685
         TabIndex        =   4
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   345
         Left            =   9885
         TabIndex        =   2
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         Height          =   345
         Left            =   8783
         TabIndex        =   1
         Top             =   4560
         Width           =   1095
      End
      Begin Sicmact.FlexEdit feCheque 
         Height          =   4050
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10860
         _ExtentX        =   19156
         _ExtentY        =   7144
         Cols0           =   9
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-nID-Fecha Operación-Banco-N° Cheque-Moneda-Monto-Agencia-Monto Utilizado"
         EncabezadosAnchos=   "400-0-1800-2500-3000-1200-1200-2000-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-L-L-C-R-L-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmChequeNegocioValorizacionExtorno"
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
Dim fsOpeDesc As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsopecod = psOpeCod
    fsOpeDesc = UCase(psOpeDesc)
    Caption = fsOpeDesc
    Show 1
End Sub
Private Sub Form_Load()
    cargar_datos
End Sub
Private Sub cmdActualizar_Click()
    cargar_datos
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cargar_datos()
    Dim oDR As New DDocRec
    Dim oRs As New ADODB.Recordset
    Dim row As Long
    
    On Error GoTo errcargar
    Screen.MousePointer = 11
    FormateaFlex feCheque
    If fsopecod = OpeCGOtrosOpeChequeExtValorizacion Or fsopecod = OpeCGOtrosOpeChequeExtValorizacionME Then
        Set oRs = oDR.ListaChequeNegxExtornoValorizacion(CInt(Mid(fsopecod, 3, 1)), gdFecSis)
    ElseIf fsopecod = OpeCGOtrosOpeChequeExtRechazo Or fsopecod = OpeCGOtrosOpeChequeExtRechazoME Then
        Set oRs = oDR.ListaChequeNegxExtornoRechazo(CInt(Mid(fsopecod, 3, 1)), gdFecSis)
    End If
    If Not oRs.EOF Then
        Do While Not oRs.EOF
            feCheque.AdicionaFila
            row = feCheque.row
            feCheque.TextMatrix(row, 1) = oRs!nId
            feCheque.TextMatrix(row, 2) = Format(oRs!dFecha, gsFormatoFechaHoraView)
            feCheque.TextMatrix(row, 3) = oRs!cIFiNombre
            feCheque.TextMatrix(row, 4) = oRs!cDocumento
            feCheque.TextMatrix(row, 5) = oRs!cMoneda
            feCheque.TextMatrix(row, 6) = Format(oRs!nMonto, gsFormatoNumeroView)
            feCheque.TextMatrix(row, 7) = oRs!cAgeDescripcion
            feCheque.TextMatrix(row, 8) = oRs!nMontoUtilizado
            oRs.MoveNext
        Loop
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Actulizado la Operación "
        Set objPista = Nothing
        '****
    Else
        MsgBox "No se encontraron datos para realizar la operación", vbInformation, "Aviso"
    End If
    Screen.MousePointer = 0
    Set oDR = Nothing
    Exit Sub
errcargar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdExtornar_Click()
    Dim oDR As NDocRec
    Dim oFun As NContFunciones
    Dim lsGlosa As String
    Dim row As Long
    Dim lnID As Long
    Dim bExito As Boolean
    Dim lsMovNro As String
    Dim lnMontoUtilizado As Double
    
    On Error GoTo ErrExtornar
    If feCheque.TextMatrix(1, 0) = "" Then
        MsgBox "No existen cheques para realizar la operación", vbInformation, "Aviso"
        Exit Sub
    End If
    row = feCheque.row
    lnID = CLng(feCheque.TextMatrix(row, 1))
    lnMontoUtilizado = CDbl(feCheque.TextMatrix(row, 8))

    If lnMontoUtilizado > 0 Then
        MsgBox "No se puede continuar porque el cheque ya fue utilizado en Operaciones, favor de verificar..", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    lsGlosa = Trim(InputBox("Ingrese una Glosa para realizar el extorno respectivo", fsOpeDesc))
    If Len(lsGlosa) = 0 Then Exit Sub
    
    If MsgBox("¿Esta seguro de realizar la operación de extorno?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    cmdExtornar.Enabled = False
    Set oDR = New NDocRec
    Set oFun = New NContFunciones
    lsMovNro = oFun.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    If fsopecod = OpeCGOtrosOpeChequeExtValorizacion Or fsopecod = OpeCGOtrosOpeChequeExtValorizacionME Then
        bExito = oDR.ExtornaValorizaChequeNegocio(lnID, lsMovNro, fsopecod, lsGlosa)
    ElseIf fsopecod = OpeCGOtrosOpeChequeExtRechazo Or fsopecod = OpeCGOtrosOpeChequeExtRechazoME Then
        bExito = oDR.ExtornarRechazoCheque(lnID, lsMovNro, fsopecod, lsGlosa)
    End If
    If bExito Then
        MsgBox "Se ha extornado con éxito la operación", vbInformation, "Aviso"
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Extornado la Operación "
        Set objPista = Nothing
        '****
        feCheque.EliminaFila row
    Else
        MsgBox "Ha sucedido un error al realizar la operación, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    cmdExtornar.Enabled = True
    Screen.MousePointer = 0
    Set oDR = Nothing
    Set oFun = Nothing
    Exit Sub
ErrExtornar:
    cmdExtornar.Enabled = True
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdExtornar.Visible And cmdExtornar.Enabled Then cmdExtornar.SetFocus
    End If
End Sub
