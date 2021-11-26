VERSION 5.00
Begin VB.Form frmCapOrdPagProceso 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmCapOrdPagProceso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5250
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefrescar 
      Caption         =   "&Refrescar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5235
      Width           =   900
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9285
      TabIndex        =   4
      Top             =   5250
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar y Enviar"
      Height          =   375
      Left            =   7665
      TabIndex        =   3
      Top             =   5250
      Width           =   1515
   End
   Begin VB.Frame fraOrdPag 
      Caption         =   "Orden Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4980
      Left            =   90
      TabIndex        =   6
      Top             =   75
      Width           =   10050
      Begin VB.OptionButton optSeleccion 
         Caption         =   "&Ninguno"
         Height          =   255
         Index           =   1
         Left            =   1455
         TabIndex        =   1
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton optSeleccion 
         Caption         =   "&Todos"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   270
         Width           =   1020
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "&Cambio Estado"
         Height          =   375
         Left            =   8430
         TabIndex        =   2
         Top             =   180
         Width           =   1485
      End
      Begin SICMACT.FlexEdit grdOrdPag 
         Height          =   4230
         Left            =   105
         TabIndex        =   7
         Top             =   675
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   7461
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Act-Cuenta-Inicio-Fin-Fecha-# Tal-Titular-nTipo-cMovNro"
         EncabezadosAnchos=   "300-500-1700-800-800-1000-600-3700-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapOrdPagProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nEstadoOrden As COMDConstantes.CapOrdPagTalEstado
Dim fsCadImp As String

Private Sub GridClear()
grdOrdPag.Rows = 2
grdOrdPag.Clear
grdOrdPag.FormaCabecera
End Sub

Private Function ExistenItemsMarcados() As Boolean
Dim I As Long
For I = 1 To grdOrdPag.Rows - 1
    If grdOrdPag.TextMatrix(I, 1) = "." Then
        ExistenItemsMarcados = True
        Exit Function
    End If
Next I
ExistenItemsMarcados = False
End Function

Public Sub Inicia(ByVal nEst As CapOrdPagTalEstado)
nEstadoOrden = nEst
If nEstadoOrden = gCapTalOrdPagEstSolicitado Then
    Me.Caption = "Orden Pago - Consolidación y Envío"
    cmdGenerar.Caption = "&Generar y Enviar"
ElseIf nEstadoOrden = gCapTalOrdPagEstEnviado Then
    Me.Caption = "Orden Pago - Recepción"
    cmdGenerar.Caption = "&Grabar"
ElseIf nEstadoOrden = gCapTalOrdPagEstRecepcionado Then
    Me.Caption = "Orden Pago - Entrega al Cliente"
    cmdGenerar.Caption = "&Grabar"
    optSeleccion(0).value = False
    optSeleccion(1).value = True
End If
cmdRefrescar.Enabled = False
cmdGenerar.Enabled = False
ObtieneDatosOrdenPago
Me.Show 1
End Sub

Private Sub ObtieneDatosOrdenPago()
Dim rsOrden As New ADODB.Recordset
Dim I As Long
Dim oCapMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento

GridClear
Set oCapMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsOrden = oCapMant.GetCapOrdPagEmision(nEstadoOrden)
If Not (rsOrden.EOF And rsOrden.BOF) Then
    I = 1
    Do While Not rsOrden.EOF
        If I >= grdOrdPag.Rows Then grdOrdPag.AdicionaFila
        grdOrdPag.TextMatrix(I, 0) = Trim(I)
        grdOrdPag.TextMatrix(I, 1) = "1"
        grdOrdPag.TextMatrix(I, 2) = rsOrden("cCtaCod")
        grdOrdPag.TextMatrix(I, 3) = rsOrden("nInicio")
        grdOrdPag.TextMatrix(I, 4) = rsOrden("nFin")
        grdOrdPag.TextMatrix(I, 5) = Format$(rsOrden("dFecha"), "dd/mm/yyyy")
        grdOrdPag.TextMatrix(I, 6) = rsOrden("nNumTal")
        grdOrdPag.TextMatrix(I, 7) = PstaNombre(rsOrden("cPersNombre"), False)
        grdOrdPag.TextMatrix(I, 8) = rsOrden("nTipo")
        grdOrdPag.TextMatrix(I, 9) = rsOrden("cMovNro")
        I = I + 1
        rsOrden.MoveNext
    Loop
    cmdRefrescar.Enabled = True
    cmdGenerar.Enabled = True
    cmdEstado.Enabled = True
    optSeleccion(0).Enabled = True
    optSeleccion(1).Enabled = True
    
    If nEstadoOrden = gCapTalOrdPagEstRecepcionado Then
       Call optSeleccion_Click(1)
    End If
    
Else
    cmdRefrescar.Enabled = True
    cmdGenerar.Enabled = False
    cmdEstado.Enabled = False
    optSeleccion(0).Enabled = False
    optSeleccion(1).Enabled = False
    MsgBox "No Existen Ordenes de Pago para este proceso", vbInformation, "Aviso"
End If
Set oCapMant = Nothing
End Sub

Private Sub cmdEstado_Click()
Dim nFila As Long, nInicio As Long
Dim nEstadoNuevo As COMDConstantes.CapOrdPagTalEstado
Dim sCuenta As String
Dim sMovNro As String, sMovNroAnt As String
Dim oMov As COMNContabilidad.NCOMContFunciones
Dim oCapMant As COMNCaptaGenerales.NCOMCaptaGenerales
If MsgBox("Desea cambiar el estado al siguiente item?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If

Set oMov = New COMNContabilidad.NCOMContFunciones
sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set oMov = Nothing

Select Case nEstadoOrden
    Case gCapTalOrdPagEstSolicitado
        nEstadoNuevo = gCapTalOrdPagEstEnviado
    Case gCapTalOrdPagEstEnviado
        nEstadoNuevo = gCapTalOrdPagEstRecepcionado
    Case gCapTalOrdPagEstRecepcionado
        nEstadoNuevo = gCapTalOrdPagEstEntregado
    Case gCapTalOrdPagEstEntregado
        nEstadoNuevo = gCapTalOrdPagEstExtornado
End Select

nFila = grdOrdPag.Row
sCuenta = grdOrdPag.TextMatrix(nFila, 2)
sMovNroAnt = grdOrdPag.TextMatrix(nFila, 9)
nInicio = CLng(grdOrdPag.TextMatrix(nFila, 3))
Set oCapMant = New COMNCaptaGenerales.NCOMCaptaGenerales
oCapMant.ActualizaCapOrdPagEmision sCuenta, sMovNro, nEstadoNuevo, sMovNroAnt, nEstadoOrden, nInicio
Set oCapMant = Nothing
ObtieneDatosOrdenPago
End Sub

Private Sub CmdImprimir_Click()
    Dim Prev As previo.clsPrevio
    If fsCadImp = "" Then
        MsgBox "No existen Ordenes para Imprimir", vbInformation, "Aviso"
    End If
    Set Prev = New clsPrevio
    Prev.Show fsCadImp, "", False, , gImpresora
    Set Prev = Nothing
    
End Sub

Private Sub cmdRefrescar_Click()
ObtieneDatosOrdenPago
End Sub

Private Sub cmdGenerar_Click()
Dim L As ListItem
Dim sCuenta As String, sMovNroAnt As String
Dim sMovNro As String
Dim oMov As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim oMant As COMNCaptaGenerales.NCOMCaptaGenerales  'NCapMantenimiento
Dim rsOrd As New ADODB.Recordset


Dim clsPrevio As previo.clsPrevio
Set clsPrevio = New previo.clsPrevio


If Not ExistenItemsMarcados() Then
    MsgBox "Debe seleccionar algún item para este proceso.", vbInformation, "Aviso"
    grdOrdPag.SetFocus
    Exit Sub
End If

Set oMov = New COMNContabilidad.NCOMContFunciones
sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set oMov = Nothing
Select Case nEstadoOrden
    Case gCapTalOrdPagEstSolicitado
        Dim sRuta As String
        Dim oHra As COMDConstSistema.DCOMGeneral
        Dim lscadimp As String
        Set rsOrd = grdOrdPag.GetRsNew()
        Set oMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Dim fs As Scripting.FileSystemObject
        Set fs = New Scripting.FileSystemObject
        Set oHra = New COMDConstSistema.DCOMGeneral
        'sRuta = App.path & "\SPOOLER\" & "CICA" & IIf(Len(Day(gdFecSis)) < 2, "0", "") & CStr(Day(dFecSis)) & IIf(Len(Month(dFecSis) < 2), "0", "") & CStr(Month(dFecSis)) & CStr(Year(dFecSis)) & Left(oHra.GetHoraServer(), 2) & Mid(GetHoraServer(), 4, 2) & Right(GetHoraServer(), 2) & ".TXT"
        sRuta = App.path & "\SPOOLER\" & "CUSCO" & IIf(Len(Day(gdFecSis)) < 2, "0", "") & CStr(Day(gdFecSis)) & IIf(Len(Month(gdFecSis) < 2), "0", "") & CStr(Month(gdFecSis)) & CStr(Year(gdFecSis)) & ".TXT"
        If fs.FileExists(sRuta) Then
            fs.DeleteFile sRuta, True
        End If
        
        Open sRuta For Output As #1
             lscadimp = oMant.GeneraArchivoEnvioOrdPagTal("", gdFecSis, rsOrd, sMovNro)
             Print #1, lscadimp
        Close #1
        Set oHra = Nothing
        Set oMant = Nothing
        If sRuta <> "" Then
            MsgBox "Archivo Creado : " & sRuta, vbInformation, "Aviso"
            fsCadImp = ImprimirOrdenesPago
            ObtieneDatosOrdenPago
        End If
    Case gCapTalOrdPagEstEnviado
        Set rsOrd = grdOrdPag.GetRsNew()
        Set oMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        oMant.ActualizaCapOrdPagEstado gdFecSis, rsOrd, sMovNro, gCapTalOrdPagEstRecepcionado, gCapTalOrdPagEstEnviado
        Set oMant = Nothing
        fsCadImp = ImprimirOrdenesPago
        ObtieneDatosOrdenPago
        
    Case gCapTalOrdPagEstRecepcionado
    
        Set rsOrd = grdOrdPag.GetRsNew()
        Set oMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        oMant.ActualizaCapOrdPagEstado gdFecSis, rsOrd, sMovNro, gCapTalOrdPagEstEntregado, gCapTalOrdPagEstRecepcionado
        oMant.ImprimeConevenioOP rsOrd  'ppoa
        Set oMant = Nothing
        fsCadImp = ImprimirOrdenesPago
        ObtieneDatosOrdenPago
                
End Select
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim sRuta As String
Dim sCabecera As String
Dim I As Long
Dim sCuerpo As String
Dim sCuenta As String
Dim sNro As String
Dim sNombre As String
Dim O As Long
Dim Conta As Integer

sRuta = App.path & "\SPOOLER\ordenP.TXT"
sCabecera = "'Cuenta'" & vbTab & "'Nombre1'" & vbTab & "'Nombre2'" & vbTab & "'Nombre3'" & vbTab & "'Numero'"
sCabecera = Replace(sCabecera, "'", Chr(34))
'Open sRuta For Output As #1
'    Print #1, sCabecera
'sCuerpo = oImpresora.gPrnCondensadaOFF & oImpresora.gPrnTamLetra15CPIDef

     For I = 1 To Me.grdOrdPag.Rows - 1
        If grdOrdPag.TextMatrix(I, 1) = "." Then
            Printer.Font.Size = 8
            Conta = 1
            For O = grdOrdPag.TextMatrix(I, 3) To grdOrdPag.TextMatrix(I, 3) + 3 'grdOrdPag.TextMatrix(i, 4)
                sCuenta = grdOrdPag.TextMatrix(I, 2)
                sNombre = Trim(grdOrdPag.TextMatrix(I, 7))
                sNro = Format(O, "00000000")
                Printer.Print Space(20) & sNro
                Printer.Print " "
                Printer.Print " "
                If Conta = 1 Then Printer.Print " "
                Printer.Font.Size = 12
                Printer.Print Space(110) & sNro
                Printer.Font.Size = 8
                Printer.Print " "
                Printer.Print " "
                'If Conta <> 2 Then Printer.Print " "
                If Conta <> 3 Then Printer.Print " "
                If Conta <> 3 Then Printer.Print " "
                If Conta <> 4 Then Printer.Print " "
                Printer.Print " "
                Printer.Print " "
                Printer.Print " "
                If Conta = 3 Then Printer.Print " "
                If Conta = 3 Then Printer.Print " "
                If Conta = 4 Then Printer.Print " "
                Printer.Print Space(80) & "AL BANCO WISSE SUDAMERICS Cta. CtE. M.N. 780-0136700"
                Printer.Print Space(80) & sCuenta
                Printer.Print Space(80) & sNombre
                If Conta <> 4 Then
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                End If
                
                Conta = Conta + 1
                
'                sCuerpo = Chr(34) & sCuenta & Chr(34) & vbTab
'                sCuerpo = sCuerpo & Chr(34) & sNombre & Chr(34) & vbTab
'                sCuerpo = sCuerpo & Chr(34) & "    " & Chr(34) & vbTab
'                sCuerpo = sCuerpo & Chr(34) & "    " & Chr(34) & vbTab
'                sCuerpo = sCuerpo & Chr(34) & sNro & Chr(34)
'
'                Print #1, sCuerpo
            Next O
        End If
     Next I
'Close #1
 

 Printer.EndDoc
 
'    Dim oWord As Word.Application
'    Dim oDoc As Word.Document
'    Dim oRange As Word.Range
'
'    Set oWord = CreateObject("Word.Application")
'    oWord.Visible = True
'
'    Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\OrdenP.doc")
'
'
'  oWord.ActiveDocument.MailMerge.OpenDataSource sRuta
'  oWord.ActiveDocument.MailMerge.Destination = wdSendToNewDocument
'  oWord.ActiveDocument.MailMerge.SuppressBlankLines = True
'  oWord.ActiveDocument.MailMerge.Execute
'  oDoc.SaveAs App.path & "\SPOOLER\ORDENES.DOC"
'  oDoc.Close
'
'  Set oWord = Nothing
'  Set oDoc = Nothing
'        '"C:\SICMAC_NEGOCIO_COM\spooler\ordenP.TXT", ConfirmConversions:=False, _
'        'ReadOnly:=False, LinkToSource:=True, AddToRecentFiles:=False, _
'        'PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
'        'WritePasswordTemplate:="", Revert:=False, Format:=wdOpenFormatAuto, _
'        'Connection:="", SQLStatement:="", SQLStatement1:="", SubType:= _

End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub optSeleccion_Click(Index As Integer)
Dim I As Long
Select Case Index
    Case 0
        For I = 1 To grdOrdPag.Rows - 1
            grdOrdPag.TextMatrix(I, 1) = "1"
        Next
    Case 1
        For I = 1 To grdOrdPag.Rows - 1
            grdOrdPag.TextMatrix(I, 1) = ""
        Next
End Select
End Sub

Public Function ImprimirOrdenesPago() As String
    Dim sCadImp As String
    Dim I As Integer
    
    If grdOrdPag.TextMatrix(1, 2) = "" Then
        MsgBox "No existen Ordenes para Imprimir", vbInformation, "Aviso"
        Exit Function
    End If
    
    sCadImp = sCadImp & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    sCadImp = sCadImp & Space(7) & "Agencia: " & gsNomAge & Space(50) & "Fecha :" & gdFecSis & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    sCadImp = sCadImp & Space(50) & "L I S T A D O   D E   O R D E N E S  D E  P A G O " & oImpresora.gPrnSaltoLinea
    sCadImp = sCadImp & Space(50) & "------------------------------------------------- " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    sCadImp = sCadImp & Space(7) & ImpreFormat("CUENTA", 20) & ImpreFormat("INICIO", 10) & ImpreFormat("FIN", 10) & ImpreFormat("FECHA", 15) & ImpreFormat("TALONARIO", 10) & ImpreFormat("TITULAR", 50) & oImpresora.gPrnSaltoLinea
    sCadImp = sCadImp & Space(7) & String(110, "-") & oImpresora.gPrnSaltoLinea
    For I = 1 To grdOrdPag.Rows - 1
        If grdOrdPag.TextMatrix(I, 1) = "." Then
           sCadImp = sCadImp & Space(7) & ImpreFormat(grdOrdPag.TextMatrix(I, 2), 20) & ImpreFormat(grdOrdPag.TextMatrix(I, 3), 10) & ImpreFormat(grdOrdPag.TextMatrix(I, 4), 10) & ImpreFormat(grdOrdPag.TextMatrix(I, 5), 15) & ImpreFormat(grdOrdPag.TextMatrix(I, 6), 10) & ImpreFormat(grdOrdPag.TextMatrix(I, 7), 50) & oImpresora.gPrnSaltoLinea
        End If
    Next I
    ImprimirOrdenesPago = sCadImp
    
End Function

