VERSION 5.00
Begin VB.Form frmCFImpresion 
   Caption         =   "Imprimir Carta Fianza"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblNroFolio 
      AutoSize        =   -1  'True
      Caption         =   "@folio"
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nº Folio:"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   600
   End
End
Attribute VB_Name = "frmCFImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCFImpresion
'***     Descripcion:      Ventana Para imprimir y ver una vista previa de la carta Fianza
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     12/03/2012 06:00:00 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Dim fsCtaCod As String
Dim fnTipo As Integer
Dim fsNumFolio As String
Dim fbAvalado As Boolean
Dim nNunImpresiones As Integer 'WIOR 20120523
Dim bCerrar As Boolean 'WIOR 20120523
Dim nControlar As Integer 'WIOR 20120526

Private Sub cmdImprimir_Click()
nControlar = 0 'WIOR 20120526
If fnTipo = 1 Then
    Call ImpreDoc(fsCtaCod)
Else
    Call ImpreDocRenovado(fsCtaCod)
End If
End Sub

Private Sub cmdVistaPrevia_Click()
    If fnTipo = 1 Then
    Call VistaPreviaDocNormal(fsCtaCod)
    Else
    Call VistaPreviaDocRenovado(fsCtaCod)
    End If
End Sub

Public Sub Inicio(ByVal psCtaCod As String, ByVal pbAvalado As Boolean, ByVal psNumFolio As String, ByVal pnTipo As Integer)
fsCtaCod = psCtaCod
fnTipo = pnTipo
fsNumFolio = psNumFolio
Me.lblNroFolio.Caption = Format(fsNumFolio, "0000000")
fbAvalado = pbAvalado
Me.Show 1
'WIOR 2012052******************
If Not bCerrar Then
    Call Inicio(fsCtaCod, fbAvalado, fsNumFolio, fnTipo)
End If
'WIOR FIN *****************
End Sub
Sub ImpreDoc(ByVal psCtaCod As String)
Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim rsCF As ADODB.Recordset
Set oCF = New COMDCartaFianza.DCOMCartaFianza
Set rsCF = oCF.RecuperaCartaFianzaDetalle(psCtaCod)
    On Error GoTo ErrorImpresion 'WIOR 20120522
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    
    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    
    Dim nDias As Integer
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim lsFechas As String
    
    Dim nPoliza As Long
    Dim cDirecAgencia As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim nCFPoliza As Long
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCtaCod)
    Set lrDataT = loRs.RecuperaDatosT(psCtaCod)
    Set lrDataCR = loRs.RecuperaDatosAcreedor(psCtaCod)
    

    cDirecAgencia = loRs.Get_Agencia_CF(psCtaCod)
    nCFPoliza = fsNumFolio
    Set loRs = Nothing
    
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = False
    
    If Not fbAvalado Then
        Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CFMaynas.doc")
        nControlar = 1
    Else
        Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CFMaynasGar.doc")
        nControlar = 1
    End If

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgenciaDir As String
    Dim lnPosicion As Integer
    
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
            lnPosicion = InStr(lsAgencia, "(")
            If lnPosicion > 0 Then
                lsAgencia = Left(lsAgencia, lnPosicion - 1)
            End If
            
        End If
    Set loAge = Nothing
    
    'Agencia
    With oWord.Selection.Find
        .Text = "sAgencia"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Numero de Cuenta
    With oWord.Selection.Find
        .Text = "<<CRED>>"
        .Replacement.Text = Left(psCtaCod, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'AVAL
    If fbAvalado Then
    With oWord.Selection.Find
        .Text = "<<AVAL>>"
        .Replacement.Text = PstaNombre(rsCF!cAvalNombre, True)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    End If
    
    'Direccion Agencia
    lsAgenciaDir = cDirecAgencia
    lnPosicion = InStr(lsAgenciaDir, "(")
    cDirecAgencia = Left(lsAgenciaDir, lnPosicion - 2)
    With oWord.Selection.Find
        .Text = "<<DIRECCION>>"
        .Replacement.Text = cDirecAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Numero Folio
    With oWord.Selection.Find
        .Text = "<<FOLIO>>"
        .Replacement.Text = Format(nCFPoliza, "0000000")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    dfechafin = CDate(lrDataCF!Vence)
    lsFechas = Format(dfechafin, "dd") & " de " & Format(dfechafin, "mmmm") & " del " & Format(dfechafin, "yyyy")
    'Fecha Vencimineto
    With oWord.Selection.Find
        .Text = "<<VENCIMIENTO>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha Actual
    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    With oWord.Selection.Find
        .Text = "<<FECHA>>"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'ACREEDOR
    With oWord.Selection.Find
        .Text = "<<SEÑORES>>"
        .Replacement.Text = PstaNombre(rsCF!cPersNomAcre, True)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'CLIENTE
    With oWord.Selection.Find
        .Text = "<<SOLICITANTE>>"
        .Replacement.Text = PstaNombre(rsCF!cPersNombre, True)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Monto
    With oWord.Selection.Find
        .Text = "<<MONTO>>"
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/. ", "$. ") & Format(lrDataT!nSaldo, "#,###0.00") & " " & "(" & UCase(NumLet(lrDataT!nSaldo)) & IIf(Mid(psCtaCod, 9, 1) = "2", "", " Y " & IIf(InStr(1, lrDataT!nSaldo, ".") = 0, "00", Mid(lrDataT!nSaldo, InStr(1, lrDataT!nSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(psCtaCod, 9, 1) = "1", "NUEVOS SOLES)", "US DOLARES)")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Finalidad
    With oWord.Selection.Find
        .Text = "<<Finalidad>>"
        .Replacement.Text = rsCF!cfinalidad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Modalidad
    With oWord.Selection.Find
        .Text = "<<Modalidad>>"
        .Replacement.Text = rsCF!MODALIDAD
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    nControlar = 0 'WIOR 20120526
    oDoc.SaveAs App.path & "\SPOOLER\" & psCtaCod & ".doc"
    nControlar = 1 'WIOR 20120526
    
    'Imprimir-INICIO
    Dim X, sImpresora As String
    Dim Prt As Printer
    Dim xbol As Boolean
    Dim Pred As String
    xbol = False
    sImpresora = Printer.DeviceName
    X = App.path & "\SPOOLER\" & psCtaCod & ".doc"

    frmImpresora.Show 1
    

    If sImpresora <> sLpt Or sImpresora <> "" Then
        'WIOR 20120522
        If ((Trim(Mid(sLpt, 1, 3)) = "LPT" Or Trim(Mid(sLpt, 1, 3)) = "COM") And Len(sLpt) = 4) Or Trim(Mid(sLpt, 1, 2)) = "Ne" Or Trim(Mid(sLpt, 1, 3)) = "USB" Or Len(sLpt) = 4 Then
            MsgBox "Impresora no encontrada", vbExclamation, "Aviso"
            oWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
            oWord.Quit
            Kill App.path & "\SPOOLER\" & psCtaCod & ".doc"
            Exit Sub
        Else
            oWord.Application.ActivePrinter = sLpt
            xbol = True
        End If
    End If

    If oWord.Application.ActivePrinter = "" Then
    Else
        oWord.PrintOut Filename:=X, Range:=wdPrintAllDocument, iTem:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
        nNunImpresiones = nNunImpresiones + 1 'WIOR 20120523
        bCerrar = True 'WIOR 20120523
    End If

    Do While MsgBox("Desea Reimprimir la Carta Fianza?", vbInformation + vbYesNo, "Aviso") = vbYes
    If oWord.Application.ActivePrinter <> "" Then
        oWord.PrintOut Filename:=X, Range:=wdPrintAllDocument, iTem:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
        nNunImpresiones = nNunImpresiones + 1 'WIOR 20120523
    End If
    Loop
    oWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges

    Kill App.path & "\SPOOLER\" & psCtaCod & ".doc"

    If xbol = True Then
         oWord.Application.ActivePrinter = sImpresora
    End If
    oWord.Quit
    'FIN
    nControlar = 0
    'WIOR 20120522
    Exit Sub
ErrorImpresion:
    MsgBox Err.Description, vbExclamation, "Aviso"
    'WIOR 20120526**************************************
    If nControlar = 1 Then
        oWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
        oWord.Quit
        Kill App.path & "\SPOOLER\" & psCtaCod & ".doc"
        Exit Sub
    End If
    'WIOR - FIN *****************************************
End Sub


Sub VistaPreviaDocNormal(ByVal psCtaCod As String)
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    On Error GoTo ErrorVistaPrevia 'WIOR 20120522
    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    
    Dim nDias As Integer
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim lsFechas As String
    
    Dim nPoliza As Long
    Dim cDirecAgencia As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim nCFPoliza As Long
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCtaCod)
    Set lrDataT = loRs.RecuperaDatosT(psCtaCod)
    Set lrDataCR = loRs.RecuperaDatosAcreedor(psCtaCod)

    cDirecAgencia = loRs.Get_Agencia_CF(psCtaCod)
    Set loRs = Nothing

    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgenciaDir As String
    Dim lnPosicion As Integer


    lsAgenciaDir = cDirecAgencia
    lnPosicion = InStr(lsAgenciaDir, "(")
    cDirecAgencia = Left(lsAgenciaDir, lnPosicion - 2)

    'Vista Previa-INICIO
    Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza
    Dim loImprime As COMNCartaFianza.NCOMCartaFianzaReporte
    Dim oPrevio As PrevioCredito.clsPrevioCredito
    Dim lsCadImprimir As String
    Dim lsmensaje As String
        Set loImprime = New COMNCartaFianza.NCOMCartaFianzaReporte
            loImprime.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsCadImprimir = loImprime.EmisionCF(psCtaCod, fbAvalado, lsmensaje, gdFecSis, cDirecAgencia)
            If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
            End If
        Set loImprime = Nothing
        Set oPrevio = New PrevioCredito.clsPrevioCredito
        oPrevio.Show lsCadImprimir, "EMISION DE CARTA FIANZA", True
        Set oPrevio = Nothing
    'FIN
    'WIOR 20120523
    Exit Sub
ErrorVistaPrevia:
    MsgBox Err.Description, vbExclamation, "Aviso"
End Sub

Sub VistaPreviaDocRenovado(ByVal psCtaCod As String)
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    On Error GoTo ErrorVistaPrevia 'WIOR 20120522
    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    
    Dim nDias As Integer
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim lsFechas As String
    
    Dim nPoliza As Long
    Dim cDirecAgencia As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim nCFPoliza As Long
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCtaCod)
    Set lrDataT = loRs.RecuperaDatosT(psCtaCod)
    Set lrDataCR = loRs.RecuperaDatosAcreedor(psCtaCod)

    cDirecAgencia = loRs.Get_Agencia_CF(psCtaCod)
    
    Set loRs = Nothing
    
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsAgenciaDir As String
    Dim lnPosicion As Integer
    
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))

            lnPosicion = InStr(lsAgencia, "(")
            If lnPosicion > 0 Then
                lsAgencia = Left(lsAgencia, lnPosicion - 1)
            End If
        End If
    Set loAge = Nothing

    lsAgenciaDir = cDirecAgencia
    lnPosicion = InStr(lsAgenciaDir, "(")
    cDirecAgencia = Left(lsAgenciaDir, lnPosicion - 2)

    'Vista Previa-INICIO
    Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza
    Dim loImprime As COMNCartaFianza.NCOMCartaFianzaReporte
    Dim oPrevio As PrevioCredito.clsPrevioCredito
    Dim lsCadImprimir As String
    Dim lsmensaje As String
        Set loImprime = New COMNCartaFianza.NCOMCartaFianzaReporte
            loImprime.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsCadImprimir = loImprime.EmisionCFRenova(fsCtaCod, fbAvalado, lsmensaje, gdFecSis, cDirecAgencia, lsAgencia)
            If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
            End If
        Set loImprime = Nothing
        Set oPrevio = New PrevioCredito.clsPrevioCredito
        oPrevio.Show lsCadImprimir, "EMISION DE CARTA FIANZA", True
        Set oPrevio = Nothing
    'FIN
    'WIOR 20120523
    Exit Sub
ErrorVistaPrevia:
    MsgBox Err.Description, vbExclamation, "Aviso"
End Sub
Sub ImpreDocRenovado(ByVal psCtaCod As String)

    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    
    Dim nDias As Integer
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim lsFechas As String
    
    Dim nPoliza As Long
    Dim cDirecAgencia As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    On Error GoTo ErrorImpresion 'WIOR 20120522
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
   
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCtaCod)
    Set lrDataT = loRs.RecuperaDatosT(psCtaCod)
    Set lrDataCR = loRs.RecuperaDatosAcreedor(psCtaCod)
      
    Dim oCF As COMDCartaFianza.DCOMCartaFianza
    Dim rsCF As ADODB.Recordset
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set rsCF = oCF.RecuperaCartaFianzaDetalle(psCtaCod)
    
   cDirecAgencia = loRs.Get_Agencia_CF(psCtaCod)
    nPoliza = CLng(loRs.GetCF_Poliza(psCtaCod))
     
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = False
    
    If fbAvalado Then
        Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CartaFianzaRenovacionGar.doc")
    Else
        Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CartaFianzaRenovacion.doc")
    End If
    
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lnPosicion As Integer
    
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))

            lnPosicion = InStr(lsAgencia, "(")
            If lnPosicion > 0 Then
                lsAgencia = Left(lsAgencia, lnPosicion - 1)
            End If
        End If
    Set loAge = Nothing
    
    'Agencia
    With oWord.Selection.Find
        .Text = "sAgencia"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Direccion Agencia
    Dim lsAgenciaDir As String
    lsAgenciaDir = cDirecAgencia
    lnPosicion = InStr(lsAgenciaDir, "(")
    cDirecAgencia = Left(lsAgenciaDir, lnPosicion - 2)
    With oWord.Selection.Find
        .Text = "sDireccion"
        .Replacement.Text = cDirecAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'Fecha
    lsFechas = Format(lrDataCF!F_Asignacion, "dd") & " de " & Format(lrDataCF!F_Asignacion, "mmmm") & " del " & Format(lrDataCF!F_Asignacion, "yyyy")
    With oWord.Selection.Find
        .Text = "dFecha"
        .Replacement.Text = lsFechas
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Titular
    With oWord.Selection.Find
        .Text = "cTitular"
        .Replacement.Text = PstaNombre(rsCF!cPersNombre, True)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'AVALADO
    If fbAvalado Then
        With oWord.Selection.Find
            .Text = "AVAL"
            .Replacement.Text = PstaNombre(rsCF!cAvalNombre, True)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    
    'Cuenta
    With oWord.Selection.Find
        .Text = "cCtaCod"
        .Replacement.Text = fsCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Numero Renovacion
    With oWord.Selection.Find
        .Text = "NroRen"
        .Replacement.Text = rsCF!nRenovacion
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
     
    'Fecha de creacion
    dfechaini = lrDataCF!dPrdEstado
    With oWord.Selection.Find
        .Text = "dFecCrea"
        .Replacement.Text = Format(dfechaini, "DD/MM/YYYY")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Monto
    With oWord.Selection.Find
        .Text = "nMonto"
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "$.") & Format(rsCF!nSaldo, "#,###0.00") '& " " & UnNumero(lblMontoSol)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Fecha Vencimiento
    dfechafin = CDate(lrDataCF!dVenc)
    With oWord.Selection.Find
        .Text = "dFecVenA"
        .Replacement.Text = Format(dfechafin, "DD/MM/YYYY")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    'ACREEDOR
    With oWord.Selection.Find
        .Text = "dAcreedor"
        .Replacement.Text = PstaNombre(rsCF!cPersNomAcre, True)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Finalidad
    With oWord.Selection.Find
        .Text = "cFinalidad"
        .Replacement.Text = lrDataCF!Finalidad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Nueva Fecha
    With oWord.Selection.Find
        .Text = "dFecVenN"
        .Replacement.Text = rsCF!dVencimiento
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    'Numero Folio
    With oWord.Selection.Find
        .Text = "FOLIO"
        .Replacement.Text = Format(nPoliza, "0000000")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    nControlar = 0 'WIOR 20120526
    oDoc.SaveAs App.path & "\SPOOLER\Renovacion" & psCtaCod & ".doc"
    nControlar = 1 'WIOR 20120526
    
    'Imprimir - INICIO
    Dim X, sImpresora As String
    Dim Prt As Printer
    Dim xbol As Boolean

    xbol = False
    sImpresora = Printer.DeviceName

    X = App.path & "\SPOOLER\Renovacion" & psCtaCod & ".doc"

    frmImpresora.Show 1

    If sImpresora <> sLpt Or sImpresora <> "" Then
        'WIOR 20120522
        If ((Trim(Mid(sLpt, 1, 3)) = "LPT" Or Trim(Mid(sLpt, 1, 3)) = "COM") And Len(sLpt) = 4) Or Trim(Mid(sLpt, 1, 2)) = "Ne" Or Trim(Mid(sLpt, 1, 3)) = "USB" Or Len(sLpt) = 4 Then
            MsgBox "Impresora no encontrada", vbExclamation, "Aviso"
            oWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
            oWord.Quit
            Kill App.path & "\SPOOLER\Renovacion" & psCtaCod & ".doc"
            Exit Sub
        Else
            oWord.Application.ActivePrinter = sLpt
            xbol = True
        End If
    End If

    If oWord.Application.ActivePrinter = "" Then
    Else
        oWord.PrintOut Filename:=X, Range:=wdPrintAllDocument, iTem:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
        nNunImpresiones = nNunImpresiones + 1 'WIOR 20120523
        bCerrar = True 'WIOR 20120523
        
    End If

    Do While MsgBox("Desea Reimprimir la Carta Fianza?", vbInformation + vbYesNo, "Aviso") = vbYes
         If oWord.Application.ActivePrinter <> "" Then
            nNunImpresiones = nNunImpresiones + 1 'WIOR 20120523
             oWord.PrintOut Filename:=X, Range:=wdPrintAllDocument, iTem:= _
             wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
             ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
             False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
             PrintZoomPaperHeight:=0
        End If
    Loop
    oWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges

    Kill App.path & "\SPOOLER\Renovacion" & psCtaCod & ".doc"

    If xbol = True Then
        oWord.Application.ActivePrinter = sImpresora
    End If

    oWord.Quit
    'FIN
     nControlar = 0
    'WIOR 20120522
    Exit Sub
ErrorImpresion:
    MsgBox Err.Description, vbExclamation, "Aviso"
    'WIOR 20120526**************************************
    'On Error GoTo ErrorAlCerrar
    If nControlar = 1 Then
        oWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
        oWord.Quit
        Kill App.path & "\SPOOLER\Renovacion" & psCtaCod & ".doc"
        Exit Sub
    End If
    'ErrorAlCerrar:
    'MsgBox Err.Description, vbExclamation, "Aviso"
    'WIOR - FIN *****************************************
End Sub


Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
nNunImpresiones = 0 'WIOR 20120523
nControlar = 0 'WIOR 20120526
bCerrar = False 'WIOR 20120523
End Sub
'WIOR 20120523
Private Sub Form_Unload(Cancel As Integer)
    If nNunImpresiones < 1 Then
        MsgBox "Debe imprimir por lo menos 1 vez.", vbInformation, "Aviso"
        bCerrar = False
        Exit Sub
    Else
        Unload Me
    End If
End Sub
'WIOR - FIN
