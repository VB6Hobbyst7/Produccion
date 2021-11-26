Attribute VB_Name = "gColPigFunciones"
'**  Funciones Generales de Colocaciones-Pignoraticio.
'**
Option Explicit
 Public vcodper As String 'MADM 2001010
'Función verifica el texto de un TextBox
Public Function fgTextoNum(ByVal vTexto As TextBox) As Boolean
    If Len(Trim(vTexto)) > 0 And Val(vTexto) > 0 Then
        fgTextoNum = True
    Else
        fgTextoNum = False
    End If
End Function

Public Function fgEliminaEnters(lsTexto As String) As String
Dim i As Integer
Dim lsAux As String
Dim pos As Integer
    lsAux = ""
    pos = 0
    For i = Len(lsTexto) To 1 Step -1
        If Mid(lsTexto, i, 1) <> Chr(13) And Mid(lsTexto, i, 1) <> Chr(10) Then
            pos = i
            Exit For
        End If
    Next
    lsAux = lsAux + Mid(lsTexto, 1, pos)
    fgEliminaEnters = lsAux
End Function

'*********************************************************************
'RUTINA QUE VALIDA EL INGRESO DE TEXTO EN MAYUSCULAS EN UN TEXTCONTROL
'**********************************************************************
Public Function fgIntfMayusculas(intTecla As Integer) As Integer
 If Chr(intTecla) >= "a" And Chr(intTecla) <= "z" Then
    intTecla = intTecla - 32
 End If
 If intTecla = 39 Then
    intTecla = 0
 End If
 If intTecla = 209 Or intTecla = 241 Or intTecla = 8 Or intTecla = 32 Then
    fgIntfMayusculas = Asc(UCase(Chr(intTecla)))
     Exit Function
 End If
 fgIntfMayusculas = intTecla
End Function

'*******************************************************
'RUTINA VALIDA EL INGRESO DE UN NUMERO MAXIMO DE LINEAS
'*******************************************************
'FECHA CREACION : 24/06/99  -   MAVF
'MODIFICACION:
'**********************************************
Public Function fgIntfLineas(cCadena As String, intTecla As Integer, intLinea As Integer) As Integer
Dim vLineas As Byte
Dim X As Byte
    If intTecla = 13 Then
        For X = 1 To Len(cCadena)
            If Mid(cCadena, X, 1) = Chr(13) Then
                vLineas = vLineas + 1
            End If
        Next X
        If vLineas >= intLinea Then
            MsgBox " No se permite mas lineas ", vbInformation, " Aviso "
            intTecla = 0
            Beep
        End If
    End If
    fgIntfLineas = intTecla
End Function

Public Function fgGetCodigoPersonaListaRsNew(pLista As ListView) As ADODB.Recordset
Dim rsAux As ADODB.Recordset
Dim lnFila As Integer

Set rsAux = New ADODB.Recordset
    rsAux.Fields.Append "cPersCod", adVarChar, 13
    rsAux.Fields.Append "cPersApellido", adVarChar, 50
    rsAux.Fields.Append "cPersNombre", adVarChar, 50
    rsAux.Fields.Append "cPersDireccDomicilio", adVarChar, 50
    rsAux.Fields.Append "cPersTelefono", adVarChar, 50
    rsAux.Fields.Append "NroDNI", adVarChar, 50
    rsAux.Fields.Append "NroRUC", adVarChar, 50
    rsAux.Fields.Append "Zona", adVarChar, 50
    rsAux.Fields.Append "Prov", adVarChar, 50
    rsAux.Fields.Append "Dpto", adVarChar, 50
    
    rsAux.Open
    For lnFila = 1 To pLista.ListItems.count
        rsAux.AddNew
        rsAux.Fields("cPersCod") = pLista.ListItems.Item(lnFila).Text
        Dim ssql As String, rstemp As ADODB.Recordset, Cn As DConecta
        Set rstemp = New ADODB.Recordset
        Dim oPers As COMDPersona.DCOMPersonas
        
        Set oPers = New COMDPersona.DCOMPersonas
            Set rstemp = oPers.ObtenerCodPersPig(Trim(pLista.ListItems.Item(lnFila).Text))
        Set oPers = Nothing
        
        If rstemp.RecordCount > 0 Then
           rsAux.Fields("cPersApellido") = rstemp!cpersapellido
           rsAux.Fields("cPersNombre") = rstemp!cPersNombre
        End If
               
        'rsAux.Fields("cPersApellido") = pLista.ListItems.Item(lnFila).ListSubItems.Item(1)
        'rsAux.Fields("cPersNombre") = pLista.ListItems.Item(lnFila).ListSubItems.Item(1)
        rsAux.Fields("cPersDireccDomicilio") = Trim(Left(pLista.ListItems.Item(lnFila).ListSubItems.Item(2), 50))
        rsAux.Fields("cPersTelefono") = Trim(Left(pLista.ListItems.Item(lnFila).ListSubItems.Item(3), 50))
        rsAux.Fields("NroDNI") = Trim(Left(pLista.ListItems.Item(lnFila).ListSubItems.Item(7), 50))
        rsAux.Fields("NroRUC") = Trim(Left(pLista.ListItems.Item(lnFila).ListSubItems.Item(9), 50))
        rsAux.Fields("Zona") = Trim(Left(pLista.ListItems.Item(lnFila).ListSubItems.Item(4), 50))
        'rsAux.Fields("Prov") = pLista.ListItems.Item(lnFila).ListSubItems.Item(11)
        'rsAux.Fields("Dpto") = pLista.ListItems.Item(lnFila).ListSubItems.Item(12)
        
        rsAux.Update
    Next
    rsAux.MoveFirst
    Set fgGetCodigoPersonaListaRsNew = rsAux
    
Set rsAux = Nothing
End Function



Public Function fgIniciaAxCuentaPignoraticio() As String
    'MAVM 20100609 BAS II
    fgIniciaAxCuentaPignoraticio = gsCodCMAC & gsCodAge '& "305"
    fgIniciaAxCuentaPignoraticio = gsCodCMAC & gsCodAge & "705"
End Function

Public Function fgMostrarClientes(lstCliente As ListView, ByVal prPers As ADODB.Recordset) As Boolean

Dim lstTmpCliente As ListItem

    If prPers.BOF And prPers.EOF Then
        MsgBox " Error al mostrar datos del cliente ", vbCritical, " Aviso "
        fgMostrarClientes = False
    Else
        lstCliente.ListItems.Clear
        Do While Not prPers.EOF
            ' Verifica si Documento de Identidad esta correcto
            'If fgVerIdentificacionPers(RegPersona!cTipPers, RegPersona!cTidoci, IIf(IsNull(RegPersona!cNudoci), "null", RegPersona!cNudoci)) = False Then
            '    MsgBox "Por favor, Actualizar el Documento de Identidad de " & vbCr & Trim(RegPersona!cNomPers) & vbCr & "Consulte al Administrador o Asistente", vbInformation, "Aviso"
            '    MostrarCliente = False
            '    Exit Function
            'End If
            
            'JUEZ 20130717 *********************************************
            Dim oPREDA As COMDPersona.DCOMPersonas
            Set oPREDA = New COMDPersona.DCOMPersonas
            If oPREDA.VerificarPersonaPREDA(Trim(prPers!cperscod), 1) Then
                MsgBox "El cliente " & prPers!cpersapellido & ", " & prPers!cPersNombre & " es un cliente PREDA no sujeto de Crédito, consultar a Coordinador de Producto Agropecuario", vbInformation, "Aviso"
                fgMostrarClientes = False
                Exit Function
            End If
            Set oPREDA = Nothing
            'END JUEZ **************************************************
            
            Set lstTmpCliente = lstCliente.ListItems.Add(, , Trim(prPers!cperscod))
                lstTmpCliente.SubItems(1) = prPers!cpersapellido & " , " & prPers!cPersNombre  'Trim(PstaNombre(prPers!cpersnombre, False))
                lstTmpCliente.SubItems(2) = Trim(prPers!cPersDireccDomicilio)
                lstTmpCliente.SubItems(3) = Trim(prPers!cPersTelefono & "")
                'lstTmpCliente.SubItems(4) = Trim(ClienteZona(RegPersona!cCodZon))
                'lstTmpCliente.SubItems(5) = Trim(ClienteCiudad(RegPersona!cCodZon))
                'lstTmpCliente.SubItems(6) = TipoDoCi(RegPersona!cTidoci & "")
                lstTmpCliente.SubItems(7) = Trim(IIf(IsNull(prPers!NroDNI), "", prPers!NroDNI))
                'lstTmpCliente.SubItems(8) = TipoDoTr(RegPersona!cTidotr & "")
                lstTmpCliente.SubItems(9) = Trim(IIf(IsNull(prPers!NroRuc), "", prPers!NroRuc))
            prPers.MoveNext
        Loop
        fgMostrarClientes = True
    End If
End Function
'Permite Muestra el Credito Pignoraticio en el AXDesCon
Public Function fgMuestraCredPig_AXDesCon(ByVal psNroContrato As String, pAXDesCon As ActXColPDesCon, Optional pbHabilitaDescrLote As Boolean = False, Optional ByVal pnTipoBusqueda As Integer = 0) As Boolean
'On Error GoTo ControlError
Dim lrCredPig As ADODB.Recordset
Dim lrCredPigJoyas As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnJoyasDet As Integer
Dim lnSolicitudApr As Integer 'MACM 20210319
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
fgMuestraCredPig_AXDesCon = True
    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
    
        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psNroContrato)
        Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psNroContrato)
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psNroContrato)
              
    Set loMuestraContrato = Nothing
        
    If lrCredPig.BOF And lrCredPig.EOF Then
        lrCredPig.Close
        Set lrCredPig = Nothing
        Set lrCredPigJoyas = Nothing
        Set lrCredPigPersonas = Nothing
        MsgBox " No se encuentra el Credito Pignoraticio " & psNroContrato, vbInformation, " Aviso "
        fgMuestraCredPig_AXDesCon = False
        Exit Function
    Else
        'MACM 20210319 VALIDACION SOL TASA PREFERENCIAL
        Dim objSolicitud As COMDCredito.DCOMNivelAprobacion
        Set objSolicitud = New COMDCredito.DCOMNivelAprobacion
        lnSolicitudApr = objSolicitud.GetCargarColocacPermisoAprobacionCuenta(psNroContrato, 1)
        Set objSolicitud = Nothing
        fgMuestraCredPig_AXDesCon = True
        If lnSolicitudApr > 0 Then
            MsgBox "El crédito tiene una solicitud de tasa preferencial pendiente", vbInformation, " Aviso "
            fgMuestraCredPig_AXDesCon = False
            Exit Function
        Else
            
            'variable para ser utilizadas por la firma
            vcodper = lrCredPigPersonas.Fields(0)
            
            pAXDesCon.Limpiar
            pAXDesCon.OroBruto = lrCredPig!nOroBruto
            pAXDesCon.OroNeto = lrCredPig!nOroNeto
            'pAXDesCon.TipoCuenta = lrCredPig!cTipCta 'RECO20140410 ERS044-2014
            pAXDesCon.Piezas = lrCredPig!nPiezas
            pAXDesCon.ValTasa = lrCredPig!nTasacion
            pAXDesCon.prestamo = lrCredPig!nMontoCol
            pAXDesCon.SaldoCapital = lrCredPig!nSaldo
            pAXDesCon.CodEstadoCred = lrCredPig!nPrdEstado
            pAXDesCon.EstadoCred = lrCredPig!nPrdEstado
            pAXDesCon.FechaPrestamo = Format(lrCredPig!dVigencia, "dd/mm/yyyy")
            pAXDesCon.FechaVencimiento = Format(lrCredPig!dVenc, "dd/mm/yyyy")
            pAXDesCon.DescLote = fgEliminaEnters(lrCredPig!cLote)
            'pAXDesCon.TasaEfectivaMensual = Format(lrCredPig!nTasaInteres, "0.00") 'RECO20140410 RES044-2014'JOEP20210916 comento campana prendario
            pAXDesCon.TasaEfectivaMensual = Format(lrCredPig!nTasaInteres, "0.00000") 'RECO20140410 RES044-2014'JOEP20210916 campana prendario
            lrCredPig.Close
            Set lrCredPig = Nothing
    
            ' Kilatajes
            
            pAXDesCon.Oro14 = IIf(IsNull(lrCredPigJoyas!nK14), "0.00", lrCredPigJoyas!nK14)
            pAXDesCon.Oro16 = IIf(IsNull(lrCredPigJoyas!nK16), "0.00", lrCredPigJoyas!nK16)
            pAXDesCon.Oro18 = IIf(IsNull(lrCredPigJoyas!nK18), "0.00", lrCredPigJoyas!nK18)
            pAXDesCon.Oro21 = IIf(IsNull(lrCredPigJoyas!nK21), "0.00", lrCredPigJoyas!nK21)
            
            lrCredPigJoyas.Close
            Set lrCredPigJoyas = Nothing
    
            ' Mostrar Clientes
            If fgMostrarClientes(pAXDesCon.listaClientes, lrCredPigPersonas) = False Then
                MsgBox " No se encuentra Datos de Clientes de Contrato " & psNroContrato, vbInformation, " Aviso "
                fgMuestraCredPig_AXDesCon = False
                Exit Function
            End If
            
           
            lrCredPigPersonas.Close
            Set lrCredPigPersonas = Nothing
            
            'Para mostrar el Detalle de las joyas
            Set loConstSis = New COMDConstSistema.NCOMConstSistema
                lnJoyasDet = loConstSis.LeeConstSistema(109)
            Set loConstSis = Nothing
            If lnJoyasDet = 1 Then
                pAXDesCon.listaJoyasDet.Visible = True
                If fgMostrarJoyasDet(pAXDesCon.listaJoyasDet, lrCredPigJoyasDet) = False Then
                    MsgBox " No se encuentra Datos de Joyas de Contrato " & psNroContrato, vbInformation, " Aviso "
                    fgMuestraCredPig_AXDesCon = False
                    Exit Function
                End If
            End If
            If pnTipoBusqueda = 0 Then
                If pbHabilitaDescrLote = True Then
                    pAXDesCon.EnabledDescLot = True
                    pAXDesCon.SetFocusDesLot
                Else
                    pAXDesCon.EnabledDescLot = False
                End If
            End If
        End If
    End If
Exit Function

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Function


Public Function fgMostrarJoyasDet(lstJoyasDet As ListView, ByVal prJoyas As ADODB.Recordset) As Boolean
Dim lstTmpCliente As ListItem

    If prJoyas.BOF And prJoyas.EOF Then
        MsgBox " Error al mostrar datos del cliente ", vbCritical, " Aviso "
        fgMostrarJoyasDet = False
    Else
        lstJoyasDet.ListItems.Clear
        Do While Not prJoyas.EOF
            Set lstTmpCliente = lstJoyasDet.ListItems.Add(, , Trim(prJoyas!nItem))
                lstTmpCliente.SubItems(1) = ImpreFormat(prJoyas!nPiezas, 3, 0)
                lstTmpCliente.SubItems(2) = ImpreFormat(prJoyas!nPesoBruto, 4, 2)
                lstTmpCliente.SubItems(3) = ImpreFormat(prJoyas!nPesoNeto, 4, 2)
                lstTmpCliente.SubItems(4) = ImpreFormat(prJoyas!nValTasac, 4, 2) 'APRI 20170417
                lstTmpCliente.SubItems(5) = Trim(prJoyas!cdescrip)
            prJoyas.MoveNext
        Loop
        fgMostrarJoyasDet = True
    End If
End Function


Public Function fgEstadoCredPigDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColPEstRegis
            lsDesc = "Registrado"
        Case gColPEstDesem
            lsDesc = "Desembolsado"
        Case gColPEstCance
            lsDesc = "Cancelado"
        Case gColPEstDifer
            lsDesc = "Entrega Joya"
        Case gColPEstVenci
            lsDesc = "Vencido"
        Case gColPEstRenov
            lsDesc = "Renovado"
        Case gColPEstRemat
            lsDesc = "Remate"
        Case gColPEstPRema
            lsDesc = "Para Remate"
        Case gColPEstAdjud
            lsDesc = "Adjudicado"
        Case gColPEstAnula
            lsDesc = "Anulado"
    End Select
    fgEstadoCredPigDesc = lsDesc
End Function

Public Function fgFormatoContratro(psContrato As String) As String
    fgFormatoContratro = Mid(psContrato, 4, 2) & "-" & Mid(psContrato, 6, 4) & "-" & Mid(psContrato, 10, 8) & "-" & Mid(psContrato, 18, 1)
End Function

Public Function fgDameNombreMes(ByVal pNroMes As Integer)
fgDameNombreMes = ""
fgDameNombreMes = Choose(pNroMes, "Enero", "Febrero", "Marzo", "Abril", _
                                    "Mayo", "Junio", "Julio", "Agosto", _
                                    "Setiembre", "Octubre", "Noviembre", "Diciembre")

End Function

Public Function fgDevuelveDatosQuery(ByVal psSql As String) As String
Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As ADODB.Recordset
    
    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(psSql)
    Set loValida = Nothing
    If Not (lrs.BOF And lrs.EOF) Then
        fgDevuelveDatosQuery = lrs!Campo
    Else
        fgDevuelveDatosQuery = ""
    End If
End Function
Public Function CalculaTasaEfectivaAnual(ByVal pnTasaMensual As Double) As Double
Dim lnTasa As Double
    lnTasa = (pnTasaMensual + 1) ^ 12 - 1
CalculaTasaEfectivaAnual = lnTasa
End Function

'*** PEAC 20080412
Public Function PigAdjudica_CartaNotarialAdju(ByVal pdFecSis As Date, ByVal pdDiasVenc As Integer, ByVal pdFecha As Date, ByVal psUser As String, ByVal psMaquina As String, _
                                        ByVal psCodAge As String, Optional psNumRemate As String) As ADODB.Recordset
    Dim OCon As COMConecta.DCOMConecta
    Dim lsSQL As String
    Dim lsSQL1 As String
    Dim pdFechaLey  As Date
    Dim ldFecAviso As Date
    Dim pnDiasVctoParaRemate As Integer
    Dim oConecta As COMConecta.DCOMConecta

   Dim loParam As COMDColocPig.DCOMColPCalculos
   Set loParam = New COMDColocPig.DCOMColPCalculos
   pnDiasVctoParaRemate = Int(loParam.dObtieneColocParametro(gConsColPDiasAtrasoParaRemate))
   Set loParam = Nothing

   ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, pdFecha)
   Set OCon = New COMConecta.DCOMConecta

   pdFechaLey = "01/06/2006"

    '*** PEAC 20080305
    'lsSQL = "exec stp_sel_ObtieneDatosCartaNotarialAdju '" & Format(pdFecSis, "yyyymmdd") & "','" & psCodAge & "'," & pdDiasVenc & ",'" & Format(pdFechaLey, "yyyymmdd") & "'"
    lsSQL = "exec stp_sel_ObtieneDatosCartaNotarialAdju '" & psCodAge & "','" & psNumRemate & "'"
        
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set PigAdjudica_CartaNotarialAdju = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing



'   If oCon.AbreConexion = True Then
'      oCon.Ejecutar lsSQL
'   End If
   'Set oCon = Nothing
   
End Function


Public Function PigRemate_CartaNotarial(ByVal pdFecha As Date, ByVal psUser As String, ByVal psMaquina As String, _
                                        ByVal psCodAge As String) As ADODB.Recordset
                                                        
   Dim OCon As COMConecta.DCOMConecta
   Dim lsSQL As String
   Dim lsSQL1 As String
   Dim pdFechaLey  As Date
   Dim ldFecAviso As Date
   Dim pnDiasVctoParaRemate As Integer
   
   Dim loParam As COMDColocPig.DCOMColPCalculos
   Set loParam = New COMDColocPig.DCOMColPCalculos
   pnDiasVctoParaRemate = Int(loParam.dObtieneColocParametro(gConsColPDiasAtrasoParaRemate))
   Set loParam = Nothing
   
   ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, pdFecha)
   Set OCon = New COMConecta.DCOMConecta
   
   pdFechaLey = "01/06/2006"
   
    '*** PEAC 20080305
   'lsSQL = "INSERT INTO PigRep_CartaNotarial  " _
           & "  SELECT  '" & psUser & "','" & psMaquina & "',P.cCtaCod,cPersNombre,cPersDireccDomicilio, P.nSaldo,nNotificacion=(SELECT nParamValor FROM ColocParametro WHERE nParamVar = 3103)," _
           & " nCostoRemate =round(CP.nTasacion*(SELECT nParamValor FROM ColocParametro WHERE nParamVar = 3005),2)," _
           & " CP.nTasacion,C.dVenc," _
           & " nTasaIntMora = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT" _
           & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = 3 )," _
           & " nTasaIntVenc = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT" _
           & "                 WHERE LCT.cLineaCred = C.cLineaCred and LCT.nColocLinCredTasaTpo = 6 )," _
           & " dVigencia,P.nPrdEstado,C.nPlazo " _
           & " FROM Producto P " _
           & " INNER JOIN ProductoPersona PP ON PP.cCtaCod= P.cCtaCod and nPrdPersRelac=20" _
           & " INNER JOIN Persona PE ON PE.cPersCod=PP.cPersCod" _
           & " INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod" _
           & " INNER JOIN ColocPignoraticio CP ON C.cCtaCod = CP.cCtaCod" _
           & " WHERE nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ")  and dVigencia >= '" & Format(pdFechaLey, "yyyymmdd") & "' " _
           & " and substring(P.cCtaCod,4,2)='" & psCodAge & "' and datediff(day,dVenc,'" & Format(ldFecAviso, "yyyymmdd") & "') >=  0  "
           
    '*** PEAC 20080305
    ' Obtiene los contratos para marcar con el costo de notificacion
    'lsSQL = "exec stp_upd_MarcaCredPigNotificadosParaRemate '" & Format(pdFechaLey, "yyyymmdd") & "','" & psCodAge & "','" & Format(ldFecAviso, "yyyymmdd") & "'"
    'oCon.Ejecutar lsSQL
    
    ' Obtiene los contratos para emitir las cartas notariales
    lsSQL = "exec stp_ins_CartaNotarial '" & psUser & "','" & Trim(psMaquina) & "','" & Format(pdFechaLey, "yyyymmdd") & "','" & psCodAge & "','" & Format(ldFecAviso, "yyyymmdd") & "'"
    '*** FIN PEAC 20080305
    
   If OCon.AbreConexion = True Then
       'ARCV 22-06-2007 : Se modifico el Orden
      OCon.Ejecutar lsSQL
      
       '*** PEAC 20080305
      'lsSQL = "SELECT * FROM PigRep_CartaNotarial where cCodUser ='" & psUser & "' and cCodMaquina ='" & psMaquina & "' "
      
      lsSQL = "exec stp_sel_SelDelCartaNotarial '" & psUser & "','" & Trim(psMaquina) & "','1'" '*** SELECT
      
      Set PigRemate_CartaNotarial = OCon.CargaRecordSet(lsSQL)
      
      'lsSQL1 = "DELETE PigRep_CartaNotarial where cCodUser ='" & psUser & "' and cCodMaquina ='" & psMaquina & "' "
      lsSQL1 = "exec stp_sel_SelDelCartaNotarial '" & psUser & "','" & Trim(psMaquina) & "','2'"  '*** DELETE
      
      '*** END PEAC
      
      OCon.Ejecutar lsSQL1
        
   End If
   Set OCon = Nothing
   
End Function

'*** Comentado por PEAC 20080305
'Public Sub Pig_CartasNotariales(ByVal dFechaCorte As Date)
'
'    Dim lrPig As ADODB.Recordset
'    Dim rs1 As ADODB.Recordset
'    Dim lsAgencia As String
'    Dim lsFecha As String
'    Dim lnDiasAtraso As Integer
'    Dim lnIntVencido As Double
'    Dim lnIntMoratorio As Double
'    Dim lnDeuda As Double
'    Dim nPag As Integer
'    Dim nDoc As Integer
'    Dim lsArchivo As String
'    Dim loAge As COMDConstantes.DCOMAgencias
'    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
'    Dim lsModeloPlantilla As String
'    Dim lsNomMaq As String
'
'    Set loAge = New COMDConstantes.DCOMAgencias
'    Set rs1 = New ADODB.Recordset
'        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
'        If Not (rs1.EOF And rs1.BOF) Then
'            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
'        End If
'    Set loAge = Nothing
'
'    lsModeloPlantilla = App.path & "\FormatoCarta\CARTANOTARIAL.doc"
'
'    lsNomMaq = GetMaquinaUsuario
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
'
'    Set lrPig = New ADODB.Recordset
'    Set lrPig = PigRemate_CartaNotarial(dFechaCorte, gsCodUser, lsNomMaq, gsCodAge)
'
'    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
'    If Not (lrPig.EOF And lrPig.BOF) Then
'        Do Until lrPig.EOF
'
'
'            wApp.Application.Selection.TypeParagraph
'            wApp.Application.Selection.Paste
'            wApp.Application.Selection.InsertBreak
'            wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
'            wApp.Selection.MoveEnd
'
'            lnDiasAtraso = DateDiff("d", lrPig!dVencimiento, gdFecSis)
'            lnIntMoratorio = loCalculos.nCalculaInteresMoratorio(lrPig!nSaldo, lrPig!nTasaMora, lnDiasAtraso)
'            lnIntVencido = loCalculos.nCalculaInteresMoratorio(lrPig!nSaldo, lrPig!nTasaIntven, lnDiasAtraso)
'
'            'Ciudad
'            With wApp.Selection.Find
'                .Text = "<<Ciudad>>"
'                .Replacement.Text = lsAgencia
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Cliente
'            With wApp.Selection.Find
'                .Text = "<<Cliente>>"
'                .Replacement.Text = Trim(lrPig!cpersnombre)
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Direccion
'            With wApp.Selection.Find
'                .Text = "<<Direccion>>"
'                .Replacement.Text = Trim(lrPig!cPersDireccion)
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Fecha Actual
'            lsFecha = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
'            With wApp.Selection.Find
'                .Text = "<<FechaActC>>"
'                .Replacement.Text = lsFecha
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Saldo Capital
'            With wApp.Selection.Find
'                .Text = "<<nCapital>>"
'                .Replacement.Text = "S/. " & Format(lrPig!nSaldo, "0.00")
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Interes Moratorio
'            With wApp.Selection.Find
'                .Text = "<<nIntVencido>>"
'                .Replacement.Text = "S/. " & Format(lnIntVencido, "0.00")
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Interes Vencido
'            With wApp.Selection.Find
'                .Text = "<<nIntMora>>"
'                .Replacement.Text = "S/. " & Format(lnIntMoratorio, "0.00")
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'
'            'Notificacion
'            With wApp.Selection.Find
'                .Text = "<<nNotificacion>>"
'                .Replacement.Text = "S/. " & Format(lrPig!nCostoNotificacion, "0.00")
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'
'            'Costo de Remate
'            '**Comentado por DAOR 20070713
'            'With wApp.Selection.Find
'            '    .Text = "<<nCostoRemate>>"
'            '    .Replacement.Text = "S/. " & Format(lrPig!nCostoRemate, "0.00")
'            '    .Forward = True
'            '    .Wrap = wdFindContinue
'            '    .Format = False
'            '    .Execute Replace:=wdReplaceAll
'            'End With
'
'            'deuda
'            'lnDeuda = lrPig!nSaldo + lnIntMoratorio + lnIntVencido + lrPig!nCostoNotificacion + lrPig!nCostoRemate
'            lnDeuda = lrPig!nSaldo + lnIntMoratorio + lnIntVencido + lrPig!nCostoNotificacion
'            With wApp.Selection.Find
'                .Text = "<<nTotalDeuda>>"
'                .Replacement.Text = "S/. " & Format(lnDeuda, "0.00")
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Tasacion
'            With wApp.Selection.Find
'                .Text = "<<nTasacion>>"
'                .Replacement.Text = "S/. " & Format(lrPig!nTasacion, "0.00")
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Fecha de Vencimiento
'            With wApp.Selection.Find
'                .Text = "<<FechaVen>>"
'                .Replacement.Text = Format(lrPig!dVencimiento, "dd/mm/yyyy")
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'            'Fecha de Vigencia
'            With wApp.Selection.Find
'                .Text = "<<FechaDes>>"
'                .Replacement.Text = Format(lrPig!dVigencia, "dd/mm/yyyy")
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
'
'
'            'para prueba peac 20080222 *******************
'            With wApp.Selection.Find
''                .Text = "<<CMAC MAYNAS S.A.>>"
''
''                Clipboard.Clear ' limpia el portapapeles
''                Clipboard.SetData Image1.Picture
''                ObjWord.Selection.Paste
''                Clipboard.Clear
'
''                .Forward = True
''                .Wrap = wdFindContinue
''                .Format = False
''                .Execute Replace:=wdReplaceAll
'            End With
'
'
'            'Cuenta
'            With wApp.Selection.Find
'                .Text = "<<Cuenta>>"
'                .Replacement.Text = lrPig!cCtaCod
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .Execute Replace:=wdReplaceAll
'            End With
''            nPag = nPag + 1
''            If nPag = 50 Then
''                nDoc = nDoc + 1
''                lsArchivo = "CartaNotarial_" & nDoc
''                'wAppSource.Documents.Save App.Path & "\SPOOLER\" & lsArchivo & ".doc"
''                wAppSource.ActiveDocument.SaveAs App.path & "\SPOOLER\" & lsArchivo & ".doc"
''                wAppSource.ActiveDocument.Close
''                wApp.Visible = True
''
''                wAppSource.Documents.Open FileName:=lsModeloPlantilla
''                Set RangeSource = wAppSource.ActiveDocument.Content
''                'Lo carga en Memoria
''                wAppSource.ActiveDocument.Content.Copy
''                'Crea Nuevo Documento
''                wApp.Documents.Add
''                nPag = 0
''            End If
'            lrPig.MoveNext
'        Loop
'    End If
'    Set loCalculos = Nothing
' 'wAppSource.ActiveDocument.Save
' wAppSource.ActiveDocument.Close
' wApp.Visible = True
'
'End Sub

Public Sub Pig_CartasNotariales(ByVal dFechaCorte As Date, ByVal psNombre As String, _
                                    ByVal psCargo As String, ByVal pImagen As Variant)

    Dim lrPig As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsFecha As String
    Dim lnDiasAtraso As Integer
    Dim lnIntVencido As Double
    Dim lnIntMoratorio As Double
    Dim lnDeuda As Double
    Dim nPag As Integer
    Dim nDoc As Integer
    Dim lsArchivo As String
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
    Dim lsModeloPlantilla As String
    Dim lsNomMaq As String
    
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
        End If
    Set loAge = Nothing

    lsModeloPlantilla = App.Path & "\FormatoCarta\CARTANOTARIAL.doc"

    lsNomMaq = GetMaquinaUsuario
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

   'Crea Nuevo Documento
    wApp.Documents.Add
 
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla

    Set RangeSource = wAppSource.ActiveDocument.Content
    
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
   
    Set lrPig = New ADODB.Recordset
    Set lrPig = PigRemate_CartaNotarial(dFechaCorte, gsCodUser, lsNomMaq, gsCodAge)
    'Set lrPig = PigAdjudica_CartaNotarial(dFechaCorte, gsCodUser, lsNomMaq, gsCodAge)

    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    If Not (lrPig.EOF And lrPig.BOF) Then
        Do Until lrPig.EOF
            wApp.Application.Selection.TypeParagraph
            wApp.Application.Selection.Paste
            wApp.Application.Selection.InsertBreak
            wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
            wApp.Selection.MoveEnd

            lnDiasAtraso = DateDiff("d", lrPig!dVencimiento, gdFecSis)
            lnIntMoratorio = loCalculos.nCalculaInteresMoratorio(lrPig!nSaldo, lrPig!nTasaMora, lnDiasAtraso)
            lnIntVencido = loCalculos.nCalculaInteresMoratorio(lrPig!nSaldo, lrPig!nTasaIntven, lnDiasAtraso)

            'Ciudad
            With wApp.Selection.Find
                .Text = "<<Ciudad>>"
                .Replacement.Text = lsAgencia
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With

            'Cliente
            With wApp.Selection.Find
                .Text = "<<Cliente>>"
                .Replacement.Text = Trim(lrPig!cPersNombre)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Direccion
            With wApp.Selection.Find
                .Text = "<<Direccion>>"
                .Replacement.Text = Trim(lrPig!cPersDireccion)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Fecha Actual
            lsFecha = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
            With wApp.Selection.Find
                .Text = "<<FechaActC>>"
                .Replacement.Text = lsFecha
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Saldo Capital
            With wApp.Selection.Find
                .Text = "<<nCapital>>"
                .Replacement.Text = "S/. " & Format(lrPig!nSaldo, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Interes Moratorio
            With wApp.Selection.Find
                .Text = "<<nIntVencido>>"
                .Replacement.Text = "S/. " & Format(lnIntVencido, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Interes Vencido
            With wApp.Selection.Find
                .Text = "<<nIntMora>>"
                .Replacement.Text = "S/. " & Format(lnIntMoratorio, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            
            'Notificacion
            With wApp.Selection.Find
                .Text = "<<nNotificacion>>"
                .Replacement.Text = "S/. " & Format(lrPig!nCostoNotificacion, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
               
            'Costo de Remate
            '**Comentado por DAOR 20070713
            'With wApp.Selection.Find
            '    .Text = "<<nCostoRemate>>"
            '    .Replacement.Text = "S/. " & Format(lrPig!nCostoRemate, "0.00")
            '    .Forward = True
            '    .Wrap = wdFindContinue
            '    .Format = False
            '    .Execute Replace:=wdReplaceAll
            'End With
            
            'deuda
            'lnDeuda = lrPig!nSaldo + lnIntMoratorio + lnIntVencido + lrPig!nCostoNotificacion + lrPig!nCostoRemate
            lnDeuda = lrPig!nSaldo + lnIntMoratorio + lnIntVencido + lrPig!nCostoNotificacion
            With wApp.Selection.Find
                .Text = "<<nTotalDeuda>>"
                .Replacement.Text = "S/. " & Format(lnDeuda, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Tasacion
            With wApp.Selection.Find
                .Text = "<<nTasacion>>"
                .Replacement.Text = "S/. " & Format(lrPig!nTasacion, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Fecha de Vencimiento
            With wApp.Selection.Find
                .Text = "<<FechaVen>>"
                .Replacement.Text = Format(lrPig!dVencimiento, "dd/mm/yyyy")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Fecha de Vigencia
            With wApp.Selection.Find
                .Text = "<<FechaDes>>"
                .Replacement.Text = Format(lrPig!dVigencia, "dd/mm/yyyy")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With



            'Cuenta
            With wApp.Selection.Find
                .Text = "<<Cuenta>>"
                .Replacement.Text = lrPig!cCtaCod
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
'            nPag = nPag + 1
'            If nPag = 50 Then
'                nDoc = nDoc + 1
'                lsArchivo = "CartaNotarial_" & nDoc
'                'wAppSource.Documents.Save App.Path & "\SPOOLER\" & lsArchivo & ".doc"
'                wAppSource.ActiveDocument.SaveAs App.path & "\SPOOLER\" & lsArchivo & ".doc"
'                wAppSource.ActiveDocument.Close
'                wApp.Visible = True
'
'                wAppSource.Documents.Open FileName:=lsModeloPlantilla
'                Set RangeSource = wAppSource.ActiveDocument.Content
'                'Lo carga en Memoria
'                wAppSource.ActiveDocument.Content.Copy
'                'Crea Nuevo Documento
'                wApp.Documents.Add
'                nPag = 0
'            End If

'wAppSource.ActiveDocument.Close
'wApp.Visible = True
'Exit Sub

            lrPig.MoveNext
        Loop
    End If
    Set loCalculos = Nothing
 'wAppSource.ActiveDocument.Save
 
'*x*x*x*x*x*x*x*x PEAC 20080305
    'crea el pie de pagina en la cual pone la firma nombre y cargo del jefe de agencia
    If wApp.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        wApp.ActiveWindow.Panes(2).Close
    End If
    If wApp.ActiveWindow.ActivePane.View.Type = wdNormalView Or wApp.ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        wApp.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    If wApp.Selection.HeaderFooter.IsHeader = True Then
        wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Else
        wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    End If

    'por algun motivo no tiene firma el jefe solo se muestra nombre y cargo
    If CStr(pImagen) <> 0 Then
        Clipboard.Clear ' limpia el portapapeles
        Clipboard.SetData pImagen ''Image1.Picture
        wApp.Selection.Paste
        Clipboard.Clear ' limpia el portapapeles
    End If

    wApp.Selection.TypeParagraph

    wApp.Selection.TypeText Text:=Trim(psNombre) '"cpersnombre"
    wApp.Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    wApp.Selection.Font.Bold = wdToggle
    wApp.Selection.Range.Case = wdTitleWord
    wApp.Selection.EndKey Unit:=wdLine
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeText Text:=Trim(psCargo) '"cargo"
    wApp.Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    wApp.Selection.Font.Size = 10
    wApp.Selection.EndKey Unit:=wdLine
    wApp.Selection.Font.Bold = wdToggle
    wApp.Selection.Font.Size = 12
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeText Text:="En caso de haber cancelado, hacer caso omiso de este aviso"
    wApp.Selection.TypeParagraph
    wApp.Selection.WholeStory
    wApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
'*x*x*x*x*x*x*x*x
 
 wAppSource.ActiveDocument.Close
 wApp.Visible = True

Set wAppSource = Nothing
Set wApp = Nothing

Clipboard.Clear

End Sub

'*** PEAC 20080419
Public Sub Pig_CartasNotarialesAdju(ByVal dFecSis As Date, ByVal nDiasVenc As Integer, ByVal dFechaCorte As Date, ByVal psNombre As String, _
                                    ByVal psCargo As String, ByVal pImagen As Variant, Optional psNumRemate As String)

    Dim lrPig As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsFecha As String
    Dim lnDiasAtraso As Integer
    Dim lnIntVencido As Double
    Dim lnIntMoratorio As Double
    Dim lnIntAdelantado As Double '*** PEAC 20080710
    Dim lnDeuda As Double
    Dim nPag As Integer
    Dim nDoc As Integer
    Dim lsArchivo As String
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
    Dim lsModeloPlantilla As String
    Dim lsNomMaq As String
        
    Set lrPig = New ADODB.Recordset
    'Set lrPig = PigAdjudica_CartaNotarial(dFechaCorte, gsCodUser, lsNomMaq, gsCodAge)
    Set lrPig = PigAdjudica_CartaNotarialAdju(dFecSis, nDiasVenc, dFechaCorte, gsCodUser, lsNomMaq, gsCodAge, psNumRemate)

    If lrPig.BOF And lrPig.EOF Then
        MsgBox "No existe créditos procesados para Notificar.", vbOKOnly, "Atención"
        Exit Sub
    End If
    
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
        End If
    Set loAge = Nothing

    lsModeloPlantilla = App.Path & "\FormatoCarta\CARTANOTARIAL.doc"

    lsNomMaq = GetMaquinaUsuario
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

   'Crea Nuevo Documento
    wApp.Documents.Add
 
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla

    Set RangeSource = wAppSource.ActiveDocument.Content
    
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
   
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    If Not (lrPig.EOF And lrPig.BOF) Then
        Do Until lrPig.EOF
        
        If lrPig!nPrdEstado <> 2102 Then '*** PEAC 20190419 - processa crédito que no sean diferidos
        
            wApp.Application.Selection.TypeParagraph
            'wApp.Application.Selection.Paste ' RIRO 20150414, Segun INC1504130017
            wApp.Application.Selection.PasteAndFormat (wdPasteDefault) ' RIRO 20150414, Segun INC1504130017
            wApp.Application.Selection.InsertBreak
            wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
            wApp.Selection.MoveEnd

            lnDiasAtraso = DateDiff("d", lrPig!dVenc, gdFecSis)
            
            '*** PEAC 20080806
            '*** PEAC 20080710 ***********************************
            'lnIntAdelantado = loCalculos.nCalculaInteresAdelantado(lrPig!nSaldo, lrPig!nTasaIntVenc, 30)
             lnIntAdelantado = loCalculos.nCalculaInteresAlVencimiento(lrPig!nSaldo, lrPig!nTasaIntVenc, 30)
            '*** FIN PEAC ****************************************
            
            lnIntMoratorio = loCalculos.nCalculaInteresMoratorio(lrPig!nSaldo, lrPig!nTasaIntMora, lnDiasAtraso)
            'lnIntVencido = loCalculos.nCalculaInteresMoratorio(lrPig!nSaldo, lrPig!nTasaIntVenc, lnDiasAtraso)
            lnIntVencido = loCalculos.nCalculaInteresMoratorio(lrPig!nSaldo, lrPig!nTasaIntVenc, lnDiasAtraso, lnIntAdelantado) '*** PEAC SE AGREGÓ "lnIntAdelantado"
            
            
            'ByVal pnSaldoCapital As Currency, ByVal pnTasaInteres As Double, _
        ByVal pnPlazo As Integer
            
            'Ciudad
            With wApp.Selection.Find
                .Text = "<<Ciudad>>"
                .Replacement.Text = lsAgencia
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With

            'Cliente
            With wApp.Selection.Find
                .Text = "<<Cliente>>"
                .Replacement.Text = Trim(lrPig!cPersNombre)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Direccion
            With wApp.Selection.Find
                .Text = "<<Direccion>>"
                .Replacement.Text = Trim(lrPig!cPersDireccDomicilio)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Fecha Actual
            lsFecha = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
            With wApp.Selection.Find
                .Text = "<<FechaActC>>"
                .Replacement.Text = lsFecha
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Saldo Capital
            With wApp.Selection.Find
                .Text = "<<nCapital>>"
                .Replacement.Text = "S/. " & Format(lrPig!nSaldo, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            '*** PEAC 20080710
            'Interes Compensatorio
            With wApp.Selection.Find
                .Text = "<<nInteres>>"
                .Replacement.Text = "S/. " & Format(lnIntAdelantado, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Interes Moratorio
            With wApp.Selection.Find
                .Text = "<<nIntVencido>>"
                .Replacement.Text = "S/. " & Format(lnIntVencido, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Interes Vencido
            With wApp.Selection.Find
                .Text = "<<nIntMora>>"
                .Replacement.Text = "S/. " & Format(lnIntMoratorio, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            
            'Notificacion
            With wApp.Selection.Find
                .Text = "<<nNotificacion>>"
                .Replacement.Text = "S/. " & Format(lrPig!nNotificacion, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
                           
            lnDeuda = lrPig!nSaldo + Round(lnIntMoratorio, 2) + Round(lnIntVencido, 2) + Round(lnIntAdelantado, 2) + lrPig!nNotificacion
            lnDeuda = Round(lnDeuda, 2)
            With wApp.Selection.Find
                .Text = "<<nTotalDeuda>>"
                .Replacement.Text = "S/. " & Format(lnDeuda, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Tasacion
            With wApp.Selection.Find
                .Text = "<<nTasacion>>"
                .Replacement.Text = "S/. " & Format(lrPig!nTasacion, "0.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Fecha de Vencimiento
            With wApp.Selection.Find
                .Text = "<<FechaVen>>"
                .Replacement.Text = Format(lrPig!dVenc, "dd/mm/yyyy")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            'Fecha de Vigencia
            With wApp.Selection.Find
                .Text = "<<FechaDes>>"
                .Replacement.Text = Format(lrPig!dVigencia, "dd/mm/yyyy")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With

            'Cuenta
            With wApp.Selection.Find
                .Text = "<<Cuenta>>"
                .Replacement.Text = lrPig!cCtaCod
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With

            End If

            lrPig.MoveNext
            
        Loop
    End If
    Set loCalculos = Nothing
 
'*x*x*x*x*x*x*x*x - pie de pagina - PEAC 20080305
    'crea el pie de pagina en la cual pone la firma nombre y cargo del jefe de agencia
    If wApp.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        wApp.ActiveWindow.Panes(2).Close
    End If
    If wApp.ActiveWindow.ActivePane.View.Type = wdNormalView Or wApp.ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        wApp.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    If wApp.Selection.HeaderFooter.IsHeader = True Then
        wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Else
        wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    End If

    'por algun motivo no tiene firma el jefe solo se muestra nombre y cargo
    If CStr(pImagen) <> 0 Then
        Clipboard.Clear ' limpia el portapapeles
        Clipboard.SetData pImagen
        wApp.Selection.Paste
        Clipboard.Clear ' limpia el portapapeles
    End If

    wApp.Selection.TypeParagraph

    wApp.Selection.TypeText Text:=Trim(psNombre) '"cpersnombre"
    wApp.Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    wApp.Selection.Font.Bold = wdToggle
    wApp.Selection.Range.Case = wdTitleWord
    wApp.Selection.EndKey Unit:=wdLine
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeText Text:=Trim(psCargo) '"cargo"
    wApp.Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    wApp.Selection.Font.Size = 10
    wApp.Selection.EndKey Unit:=wdLine
    wApp.Selection.Font.Bold = wdToggle
    wApp.Selection.Font.Size = 12
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeText Text:="En caso de haber cancelado, hacer caso omiso de este aviso"
    wApp.Selection.TypeParagraph
    wApp.Selection.WholeStory
    wApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
'*x*x*x*x*x*x*x*x
 
 wAppSource.ActiveDocument.Close
 wApp.Visible = True

Set wAppSource = Nothing
Set wApp = Nothing

Clipboard.Clear

End Sub


'Public Function GetMaquinaUsuario() As String  'Para obtener la Maquina del Usuario
'    Dim buffMaq As String
'    Dim lSizeMaq As Long
'    buffMaq = Space(255)
'    lSizeMaq = Len(buffMaq)
'
'    GetMaquinaUsuario = GetComputerName()
'End Function
'

Public Sub ImprimeComprobanteAdj(ByVal pbReimpresion As Boolean, Optional ByVal psCtaCod As String = "", Optional ByVal psClienteNombre As String = "", Optional ByVal psClienteCod As String = "" _
                            , Optional ByVal poActXColPDesCon As ActXColPDesCon, Optional ByVal pnMonto As Currency = 0, Optional ByVal pnMovNro As Long = 0 _
                            , Optional ByVal nNroElemento As Integer = 0, Optional ByVal pbReimpEmisor As Boolean = False, Optional ByVal pbReimpUsuario As Boolean = False _
                            , Optional ByVal pbReimpContabilidad As Boolean = False, Optional ByVal psPersDirec As String = "", _
                            Optional ByVal pnDsctoInt As Double = 0) 'NAGL ERS012-2017

    'PEAC 20200729 - se agregó dscto int por pnDsctoInt

  Dim oFun As New COMFunciones.FCOMImpresion
    Dim oCons As New COMDConstantes.DCOMAgencias
    Dim oColP As New COMNColoCPig.NCOMColPContrato
    Dim rsAge As New ADODB.Recordset
    Dim rsColP As New ADODB.Recordset
    Dim rsDatosReimp As New ADODB.Recordset
    Dim nTamLet As Integer
    Dim oDoc  As New cPDF
    
    Dim sNroComprobante As String
    Dim sNroComprAge As String
    Dim nIndex As Integer
    Dim contador As Integer
    Dim nCentrar As Integer
    Dim nIndexJoya As Integer
    Dim nContInterno As Integer
    Dim nMontoReimp As Currency
    Dim sCiudadEmite As String
    Dim sCiudadMinMayu As String 'NAGL ERS012-2017
    Dim nIGV As Double 'NAGL ERS012-2017
    Dim cDescripcionReg As String 'NAGL ERS012-2017
    Dim cDocNro As String 'NAGL ERS012-2017
    Dim gOpeTpo As String 'NAGL ERS012-2017
    Dim gsOpeCodReg As String 'NAGL ERS012-2017
    Dim gnDocTpo As Long 'NAGL ERS012-2017
    Dim pdFecha As Date
    Dim oMov As New DMov
    Dim sMovNroRegV As String
    Dim nMovNroRegV As Long
    
    On Error GoTo ErrorImprimirPDF
    
    
    nTamLet = 6
    nIGV = 0
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "COMPROBANTE DE PAGO "
    oDoc.Title = "COMPROBANTE DE PAGO "
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\_RecupJoya" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then Exit Sub
    
    oDoc.Fonts.Add "F1", "Arial Narrow", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial Narrow", TrueType, Bold, WinAnsiEncoding
    
    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo"
    

    contador = 0: nIndexJoya = 0: nCentrar = 80
    gsSimbolo = IIf(Mid(psCtaCod, 9, 1) = 1, "S/.", "")
    Set rsColP = oColP.PignoDevuelveNroComprobanteAdj(gsCodAge)
    
    If Not (rsColP.EOF And rsColP.BOF) Then
        sNroComprobante = CInt(rsColP!nNroComp)
        sNroComprAge = rsColP!cNroComp
        Select Case Len(sNroComprobante)
            Case 1
                sNroComprobante = "000000" & sNroComprobante
            Case 2
                sNroComprobante = "00000" & sNroComprobante
            Case 3
                sNroComprobante = "0000" & sNroComprobante
            Case 4
                sNroComprobante = "000" & sNroComprobante
            Case 5
                sNroComprobante = "00" & sNroComprobante
            Case 6
                sNroComprobante = "0" & sNroComprobante
            Case 7
                sNroComprobante = sNroComprobante
        End Select
    End If
    
    If Not pbReimpresion Then
        For nIndex = 1 To 3
            If nIndex = 1 Or nIndex = 3 Then
               oDoc.NewPage A4_Vertical
            End If
        
            oDoc.WImage 95 + contador, 95, 50, 100, "Logo"
            oDoc.WTextBox 45 + contador, 185, 10, 200, "CAJA MUNICIPAL DE AHORRO Y CREDITO MAYNAS S.A.", "F1", nTamLet, hCenter
            
            'oDoc.WTextBox 70 + contador, 237, 10, 200, gsNomAge, "F1", nTamLet, hCenter
            oDoc.WTextBox 55 + contador, 160 + nCentrar, 3, 200, "AGENCIA:", "F2", nTamLet, hLeft
            Set rsAge = oCons.ObtieneDatosAgencia(gsCodAge)
            If Not (rsAge.EOF And rsAge.BOF) Then
                oDoc.WTextBox 55 + contador, 200 + nCentrar, 3, 200, rsAge!cAgeDireccion, "F1", nTamLet, hLeft
                oDoc.WTextBox 61 + contador, 200 + nCentrar, 3, 200, rsAge!DI, "F1", nTamLet, hLeft
                oDoc.WTextBox 67 + contador, 200 + nCentrar, 3, 200, "Telf.:" & rsAge!cAgeTelefono, "F1", nTamLet, hLeft
                oDoc.WTextBox 73 + contador, 200 + nCentrar, 3, 200, rsAge!DI & "-" & rsAge!pro & "-" & rsAge!DEP, "F1", nTamLet, hLeft
            End If
            
            'Oficina Principal
            oDoc.WTextBox 81 + contador, 160 + nCentrar, 3, 200, "OFIC. PRINC.:", "F2", nTamLet, hLeft
            Set rsAge = oCons.ObtieneDatosAgencia("01")
            If Not (rsAge.EOF And rsAge.BOF) Then
                oDoc.WTextBox 81 + contador, 200 + nCentrar, 3, 200, rsAge!cAgeDireccion, "F1", nTamLet, hLeft
                oDoc.WTextBox 87 + contador, 200 + nCentrar, 3, 200, "Telf.:" & rsAge!cAgeTelefono, "F1", nTamLet, hLeft
                oDoc.WTextBox 93 + contador, 200 + nCentrar, 3, 200, rsAge!DI & "-" & rsAge!pro & "-" & rsAge!DEP, "F1", nTamLet, hLeft
            End If
            
            
            oDoc.WTextBox 41 + contador, 375, 25, 120, "RUC N° 20103845328", "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
            oDoc.WTextBox 66 + contador, 375, 20, 120, "", "F1", 10, hCenter, vMiddle, vbWhite, 1, vbBlack
            oDoc.WTextBox 66 + contador, 375, 20, 120, "COMPROBANTE DE PAGO", "F1", 10, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 86 + contador, 375, 25, 120, Replace(sNroComprAge, "PA", "") & " - " & "N° " & sNroComprobante, "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
             
            oDoc.WTextBox 106 + contador, 40 + nCentrar, 10, 200, "Señor(es):", "F2", nTamLet, hLeft
            oDoc.WTextBox 106 + contador, 80 + nCentrar, 10, 200, psClienteNombre, "F1", nTamLet, hLeft
            
            oDoc.WTextBox 116 + contador, 40 + nCentrar, 10, 200, "Dirección:", "F2", nTamLet, hLeft
            oDoc.WTextBox 116 + contador, 80 + nCentrar, 10, 200, psPersDirec, "F1", nTamLet, hLeft
            
            oDoc.WTextBox 126 + contador, 40 + nCentrar, 10, 200, "R.U.C. Nº:", "F2", nTamLet, hLeft
            oDoc.WTextBox 126 + contador, 80 + nCentrar, 10, 200, poActXColPDesCon.listaClientes.ListItems(1).SubItems(4), "F1", nTamLet, hLeft
            
            oDoc.WTextBox 126 + contador, 170 + nCentrar, 10, 200, "D.N.I. Nº:", "F2", nTamLet, hLeft
            oDoc.WTextBox 126 + contador, 210 + nCentrar, 10, 200, poActXColPDesCon.listaClientes.ListItems(1).SubItems(7), "F1", nTamLet, hLeft
            
            
            sCiudadEmite = Trim(Mid(rsAge!UB, 1, InStr(1, rsAge!UB, "(") - 1))
            oDoc.WTextBox 126 + contador, 215 + nCentrar, 10, 150, sCiudadEmite & "," & ArmaFecha(gdFecSis), "F1", 7, hRight
            
            oDoc.WTextBox 140 + contador, 20 + nCentrar, 15, 45, "CANTIDAD", "F1", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 140 + contador, 20 + nCentrar, 130, 45, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 140 + contador, 65 + nCentrar, 15, 245, " DESCRIPCION", "F1", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 140 + contador, 65 + nCentrar, 130, 245, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 140 + contador, 310 + nCentrar, 15, 50, "P.UNITARIO", "F1", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 140 + contador, 310 + nCentrar, 130, 50, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 140 + contador, 360 + nCentrar, 15, 50, " IMPORTE", "F1", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 140 + contador, 360 + nCentrar, 130, 50, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            
            For nIndexJoya = 1 To poActXColPDesCon.listaJoyasDet.ListItems.count
                oDoc.WTextBox 155 + nContInterno + contador, 20 + nCentrar, 130, 45, poActXColPDesCon.listaJoyasDet.ListItems(nIndexJoya).SubItems(1), "F1", nTamLet, hCenter
                oDoc.WTextBox 155 + nContInterno + contador, 70 + nCentrar, 130, 200, UCase(poActXColPDesCon.listaJoyasDet.ListItems(nIndexJoya).SubItems(5)), "F1", nTamLet, hLeft
                nContInterno = nContInterno + 8
            Next
            
            'PEAC 20200811 - SE SUMA A LA VARIABLE pnDsctoInt
            oDoc.WTextBox 155 + IIf(nContInterno = 8, 0, (nContInterno / 2)) + contador, 310 + nCentrar, 100, 50, Format(pnMonto + pnDsctoInt, gsFormatoNumeroView), "F1", nTamLet, hCenter
            oDoc.WTextBox 155 + IIf(nContInterno = 8, 0, (nContInterno / 2)) + contador, 365 + nCentrar, 100, 50, Format(pnMonto + pnDsctoInt, gsFormatoNumeroView), "F1", nTamLet, hCenter
            
            oDoc.WTextBox 155 + nContInterno + contador, 70 + nCentrar, 130, 245, "Nº CONTRATO: " & psCtaCod, "F2", nTamLet, hLeft
            
            'PEAC 20200729
            'oDoc.WTextBox 260 + contador, 70 + nCentrar, 130, 245, " SON: " & ConvNumLet(pnMonto), "F1", nTamLet, hLeft
            oDoc.WTextBox 260 + contador, 70 + nCentrar, 130, 245, " SON: " & ConvNumLet(pnMonto), "F1", nTamLet, hLeft
            
            nContInterno = 0
            
            'oDoc.WTextBox 275 + contador, 20 + nCentrar, 8, 140, "NOTA: NO HAY DERECHO A DEVOLUCIÓN", "F2", 6, hLeft, vMiddle, vbBlack, 1, vbBlack
            'oDoc.WTextBox 285 + contador, 20 + nCentrar, 10, 310, "BIENES TRANSFERIDOS / SERVICIOS PRESTADOS EN LA REGION SELVA PARA SER CONSUMIDOS EN LA MISMA", "F1", 6, hLeft
            
            '*******************Begin NAGL*********************************************'
             oDoc.WTextBox 310 + contador, 20 + nCentrar, 8, 140, "NOTA: RECIBO DE INGRESO AUTORIZADO", "F1", 6, hLeft ', vMiddle, vbBlack, 1, vbBlack
             oDoc.WTextBox 318 + contador, 20 + nCentrar, 10, 310, "EMITIDO EN APLICACIÓN DE LA RESOLUCIÓN DE SUPERINTENDENCIA N° 007-99 SUNAT (24.01.99 - NUMERAL 6 - INCISO B)", "F1", 6, hLeft
             'oDoc.WTextBox 320 + contador, 20 + nCentrar, 10, 400, "NOTA: RECIBO DE INGRESO AUTORIZADO EMITIDO EN APLICACION DE LA RESOLUCIÓN DE SUPERINTENDENCIA N°007-99SUNAT (24.01.99-NUMERAL 6-INCISOB)", "F1", 6, hLeft
            '**********Agregado by NAGL 20170918 según INC1709150018*******************'
        
            sCiudadMinMayu = UCase(Mid(Trim(rsAge!UB), 1, 1)) & LCase(Mid(Trim(rsAge!UB), 2, Len(rsAge!UB))) 'Para que la ciudad sea minuscula y mayuscula
            sCiudadEmite = Trim(Mid(sCiudadMinMayu, 1, InStr(1, sCiudadMinMayu, "(") - 1))
            oDoc.WTextBox 290 + contador, 150 + nCentrar, 10, 185, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", "F2", 6, hLeft
            oDoc.WTextBox 297 + contador, 170 + nCentrar, 10, 55, "C A N C E L A D O", "F2", 7, hCenter
            oDoc.WTextBox 305 + contador, 122 + nCentrar, 10, 150, sCiudadEmite & "," & ArmaFecha(gdFecSis), "F1", nTamLet, hCenter
            
            oDoc.WTextBox 277 + contador, 330 + nCentrar, 5, 50, "SUBTOTAL", "F2", nTamLet, hLeft
            oDoc.WTextBox 277 + contador + 10, 330 + nCentrar, 5, 50, "DSCTO ", "F2", nTamLet, hLeft
            oDoc.WTextBox 287 + contador + 10, 330 + nCentrar, 5, 50, "IGV", "F2", nTamLet, hLeft
            oDoc.WTextBox 297 + contador + 10, 330 + nCentrar, 5, 50, "TOTAL", "F2", nTamLet, hLeft
            
            oDoc.WTextBox 275 + contador, 440, 10, 50, Format(pnMonto + pnDsctoInt, gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack, 1, vbBlack
            oDoc.WTextBox 275 + contador + 10, 440, 10, 50, "( " + Format(pnDsctoInt, gsFormatoNumeroView) + " )", "F1", nTamLet, hCenter, vMiddle, vbBlack, 1, vbBlack
            oDoc.WTextBox 285 + contador + 10, 440, 10, 50, Format(nIGV, gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack, 1, vbBlack
            oDoc.WTextBox 295 + contador + 10, 440, 10, 50, Format(pnMonto, gsFormatoNumeroView), "F2", nTamLet, hCenter, vMiddle, vbBlack, 1, vbBlack
            
            oDoc.WTextBox 310 + contador + 10, 360 + nCentrar, 10, 50, IIf(nIndex = 1, "EMISOR", IIf(nIndex = 2, "USUARIO", "SUNAT")), "F2", 6, hCenter
            
            If Not nIndex = 2 Then
                oDoc.WTextBox 360 + contador + 10, 0, 10, 700, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ", "F2", 6, hLeft
            End If
            If nIndex = 1 Then
                contador = contador + 380
            Else
                contador = 0
            End If
        Next
        Call oColP.PignoRegistraNroCompAdj(sNroComprAge, sNroComprobante, psCtaCod, pnMovNro)
        
        sMovNroRegV = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        gsOpeCodReg = "760210"
        oMov.InsertaMov sMovNroRegV, gsOpeCodReg, "REGISTRO DE VENTAS", gMovEstContabNoContable, gMovFlagVigente
        nMovNroRegV = oMov.GetnMovNro(sMovNroRegV)
        gOpeTpo = "8" 'Venta de Adjudicado
        gnDocTpo = 13
        cDocNro = Replace(sNroComprAge, "PA", "") & sNroComprobante
        Set rsColP = oColP.PignoDescripcionJoyaAdj(psCtaCod)
        Call oColP.RegistroVentaContratoPigno(gOpeTpo, gnDocTpo, cDocNro, gdFecSis, psClienteCod, psCtaCod, Trim(rsColP!cdescrip), pnMonto, nIGV, pnMonto, nMovNroRegV, 1, , , , gsCodArea & gsCodAge)  'NAGL ERS012-2017
        
    Else
        For nIndex = 1 To nNroElemento
        Set rsDatosReimp = oColP.PignoReimpresionCompAdj(psCtaCod)
        If Not (rsDatosReimp.EOF And rsDatosReimp.BOF) Then
            If nIndex = 1 Or nIndex = 3 Then
               oDoc.NewPage A4_Vertical
            End If
            Set rsDatosReimp = oColP.PignoReimpresionCompAdj(psCtaCod)
            oDoc.WImage 95 + contador, 95, 50, 100, "Logo"
            oDoc.WTextBox 45 + contador, 185, 10, 200, "CAJA MUNICIPAL DE AHORRO Y CREDITO MAYNAS S.A.", "F1", nTamLet, hCenter
            
            'oDoc.WTextBox 70 + contador, 237, 10, 200, gsNomAge, "F1", nTamLet, hCenter
            oDoc.WTextBox 55 + contador, 160 + nCentrar, 3, 200, "AGENCIA:", "F2", nTamLet, hLeft
            Set rsAge = oCons.ObtieneDatosAgencia(gsCodAge)
            If Not (rsAge.EOF And rsAge.BOF) Then
                oDoc.WTextBox 55 + contador, 200 + nCentrar, 3, 200, rsAge!cAgeDireccion, "F1", nTamLet, hLeft
                oDoc.WTextBox 61 + contador, 200 + nCentrar, 3, 200, rsAge!DI, "F1", nTamLet, hLeft
                oDoc.WTextBox 67 + contador, 200 + nCentrar, 3, 200, "Telf.:" & rsAge!cAgeTelefono, "F1", nTamLet, hLeft
                oDoc.WTextBox 73 + contador, 200 + nCentrar, 3, 200, rsAge!DI & "-" & rsAge!pro & "-" & rsAge!DEP, "F1", nTamLet, hLeft
            End If
            
            'Oficina Principal
            oDoc.WTextBox 81 + contador, 160 + nCentrar, 3, 200, "OFIC. PRINC.:", "F2", nTamLet, hLeft
            Set rsAge = oCons.ObtieneDatosAgencia("01")
            If Not (rsAge.EOF And rsAge.BOF) Then
                oDoc.WTextBox 81 + contador, 200 + nCentrar, 3, 200, rsAge!cAgeDireccion, "F1", nTamLet, hLeft
                oDoc.WTextBox 87 + contador, 200 + nCentrar, 3, 200, "Telf.:" & rsAge!cAgeTelefono, "F1", nTamLet, hLeft
                oDoc.WTextBox 93 + contador, 200 + nCentrar, 3, 200, rsAge!DI & "-" & rsAge!pro & "-" & rsAge!DEP, "F1", nTamLet, hLeft
            End If
            
            
            oDoc.WTextBox 41 + contador, 375, 25, 120, "RUC N° 20103845328", "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
            oDoc.WTextBox 66 + contador, 375, 20, 120, "", "F1", 10, hCenter, vMiddle, vbWhite, 1, vbBlack
            oDoc.WTextBox 66 + contador, 375, 20, 120, "COMPROBANTE DE PAGO", "F1", 10, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 86 + contador, 375, 25, 120, Replace(sNroComprAge, "PA", "") & " - " & "N° " & sNroComprobante, "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
             
            Set rsDatosReimp = oColP.PignoReimpresionCompAdj(psCtaCod)
            oDoc.WTextBox 106 + contador, 40 + nCentrar, 10, 200, "Señor(es):", "F2", nTamLet, hLeft
            oDoc.WTextBox 106 + contador, 80 + nCentrar, 10, 200, rsDatosReimp!cPersNombre, "F1", nTamLet, hLeft
            
            oDoc.WTextBox 116 + contador, 40 + nCentrar, 10, 200, "Dirección:", "F2", nTamLet, hLeft
            oDoc.WTextBox 116 + contador, 80 + nCentrar, 10, 200, rsDatosReimp!cPersDireccDomicilio, "F1", nTamLet, hLeft
            
            oDoc.WTextBox 126 + contador, 40 + nCentrar, 10, 200, "R.U.C. Nº:", "F2", nTamLet, hLeft
            oDoc.WTextBox 126 + contador, 80 + nCentrar, 10, 200, rsDatosReimp!Ruc, "F1", nTamLet, hLeft
            
            oDoc.WTextBox 126 + contador, 170 + nCentrar, 10, 200, "D.N.I. Nº:", "F2", nTamLet, hLeft
            oDoc.WTextBox 126 + contador, 210 + nCentrar, 10, 200, rsDatosReimp!DNI, "F1", nTamLet, hLeft
            
            
            sCiudadEmite = Trim(Mid(rsAge!UB, 1, InStr(1, rsAge!UB, "(") - 1))
            oDoc.WTextBox 126 + contador, 215 + nCentrar, 10, 150, sCiudadEmite & "," & ArmaFecha(gdFecSis), "F1", 7, hRight
            
            oDoc.WTextBox 140 + contador, 20 + nCentrar, 15, 45, "CANTIDAD", "F1", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 140 + contador, 20 + nCentrar, 130, 45, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 140 + contador, 65 + nCentrar, 15, 245, " DESCRIPCION", "F1", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 140 + contador, 65 + nCentrar, 130, 245, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 140 + contador, 310 + nCentrar, 15, 50, "P.UNITARIO", "F1", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 140 + contador, 310 + nCentrar, 130, 50, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            oDoc.WTextBox 140 + contador, 360 + nCentrar, 15, 50, " IMPORTE", "F1", nTamLet, hCenter, vMiddle, vbWhite, 1, vbBlack, True
            oDoc.WTextBox 140 + contador, 360 + nCentrar, 130, 50, "", "F1", nTamLet, hCenter, , , 1, vbBlack
            
            For nIndexJoya = 1 To rsDatosReimp.RecordCount
                oDoc.WTextBox 155 + nContInterno + contador, 20 + nCentrar, 130, 45, rsDatosReimp!nPiezas, "F1", nTamLet, hCenter
                oDoc.WTextBox 155 + nContInterno + contador, 70 + nCentrar, 130, 200, rsDatosReimp!cJoyaDesc, "F1", nTamLet, hLeft
                nContInterno = nContInterno + 8
            Next
            
            oDoc.WTextBox 155 + IIf(nContInterno = 8, 0, (nContInterno / 2)) + contador, 310 + nCentrar, 100, 50, Format(nMontoReimp + pnDsctoInt, gsFormatoNumeroView), "F1", nTamLet, hCenter
            oDoc.WTextBox 155 + IIf(nContInterno = 8, 0, (nContInterno / 2)) + contador, 365 + nCentrar, 100, 50, Format(nMontoReimp + pnDsctoInt, gsFormatoNumeroView), "F1", nTamLet, hCenter
            
            oDoc.WTextBox 155 + nContInterno + contador, 70 + nCentrar, 130, 245, "Nº CONTRATO: " & psCtaCod, "F2", nTamLet, hLeft
            oDoc.WTextBox 260 + contador, 70 + nCentrar, 130, 245, " SON: " & ConvNumLet(pnMonto), "F1", nTamLet, hLeft
            nContInterno = 0
            
            'oDoc.WTextBox 275 + contador, 20 + nCentrar, 8, 140, "NOTA: NO HAY DERECHO A DEVOLUCIÓN", "F2", 6, hLeft, vMiddle, vbBlack, 1, vbBlack
            'oDoc.WTextBox 285 + contador, 20 + nCentrar, 10, 310, "BIENES TRANSFERIDOS / SERVICIOS PRESTADOS EN LA REGION SELVA PARA SER CONSUMIDOS EN LA MISMA", "F1", 6, hLeft
            
            '*******************Begin NAGL*********************************************'
             oDoc.WTextBox 310 + contador, 20 + nCentrar, 8, 140, "NOTA: RECIBO DE INGRESO AUTORIZADO", "F1", 6, hLeft ', vMiddle, vbBlack, 1, vbBlack
             oDoc.WTextBox 318 + contador, 20 + nCentrar, 10, 310, "EMITIDO EN APLICACIÓN DE LA RESOLUCIÓN DE SUPERINTENDENCIA N° 007-99 SUNAT (24.01.99 - NUMERAL 6 - INCISO B)", "F1", 6, hLeft
             'oDoc.WTextBox 320 + contador, 20 + nCentrar, 10, 400, "NOTA: RECIBO DE INGRESO AUTORIZADO EMITIDO EN APLICACION DE LA RESOLUCIÓN DE SUPERINTENDENCIA N°007-99SUNAT (24.01.99-NUMERAL 6-INCISOB)", "F1", 6, hLeft
            '**********Agregado by NAGL 20170918 según INC1709150018*******************'
            
            sCiudadMinMayu = UCase(Mid(Trim(rsAge!UB), 1, 1)) & LCase(Mid(Trim(rsAge!UB), 2, Len(rsAge!UB))) 'Para que la ciudad sea minuscula y mayuscula
            sCiudadEmite = Trim(Mid(sCiudadMinMayu, 1, InStr(1, sCiudadMinMayu, "(") - 1))
            oDoc.WTextBox 290 + contador, 150 + nCentrar, 10, 185, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", "F2", 6, hLeft
            oDoc.WTextBox 297 + contador, 170 + nCentrar, 10, 55, "C A N C E L A D O", "F2", 7, hCenter
            oDoc.WTextBox 305 + contador, 122 + nCentrar, 10, 150, sCiudadEmite & "," & ArmaFecha(gdFecSis), "F1", nTamLet, hCenter
            
            oDoc.WTextBox 277 + contador, 330 + nCentrar, 5, 50, "SUBTOTAL", "F2", nTamLet, hLeft
            oDoc.WTextBox 277 + contador + 10, 330 + nCentrar, 5, 50, "DSCTO ", "F2", nTamLet, hLeft
            oDoc.WTextBox 287 + contador + 10, 330 + nCentrar, 5, 50, "IGV", "F2", nTamLet, hLeft
            oDoc.WTextBox 297 + contador + 10, 330 + nCentrar, 5, 50, "TOTAL", "F2", nTamLet, hLeft
            
            oDoc.WTextBox 275 + contador + 10, 440, 10, 50, Format(pnMonto + pnDsctoInt, gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack, 1, vbBlack
            oDoc.WTextBox 275 + contador + 10, 440, 10, 50, "( " + Format(pnDsctoInt, gsFormatoNumeroView) + " )", "F1", nTamLet, hCenter, vMiddle, vbBlack, 1, vbBlack
            oDoc.WTextBox 285 + contador + 10, 440, 10, 50, Format(nIGV, gsFormatoNumeroView), "F1", nTamLet, hCenter, vMiddle, vbBlack, 1, vbBlack
            oDoc.WTextBox 295 + contador + 10, 440, 10, 50, Format(pnMonto, gsFormatoNumeroView), "F2", nTamLet, hCenter, vMiddle, vbBlack, 1, vbBlack
            
            oDoc.WTextBox 310 + contador + 10, 360 + nCentrar, 10, 50, IIf(nIndex = 1, "EMISOR", IIf(nIndex = 2, "USUARIO", "SUNAT")), "F2", 6, hCenter
            
            If Not nIndex = 2 Then
                oDoc.WTextBox 360 + contador + 10, 0, 10, 700, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ", "F2", 6, hLeft
            End If
            If nIndex = 1 Then
                contador = contador + 380
            Else
                contador = 0
            End If
        End If
     Next
    End If
    
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox err.Description, vbInformation, "Aviso"
    
End Sub 'CAMBIO TOTAL ******************************************NAGL ERS 012-2017
'INICIO EAAS SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
Public Function CargaHojaResumenPignoPDF( _
                    ByVal rsPig As ADODB.Recordset, _
                    ByVal rsPigJoyas As ADODB.Recordset, _
                    ByVal rsPigPers As ADODB.Recordset, _
                    ByVal rsPigCostos As ADODB.Recordset, _
                    ByVal rsPigDet As ADODB.Recordset, _
                    ByVal rsPigTasas As ADODB.Recordset, _
                    ByVal rsPigCosNot As ADODB.Recordset, _
                    Optional nCampana As Integer = 0 _
                    ) As Boolean

    ' RIRO 20210922 ADD nCampana
    CargaHojaResumenPignoPDF = False
    
    Dim oPDF As New cPDF
    Dim obj As New COMNCajaGeneral.NCOMDocRec
    Dim R As New ADODB.Recordset
    Dim lsLetras As String
    Dim lnMonto As Currency
    Dim lnLinea As Integer
    Dim nFila, i As Integer
    Dim nValor As Double
    Dim nHeight As Integer
    Dim nFila2 As Integer
    Dim nFila3 As Integer
    Dim nFila4 As Integer
    Dim nFila5 As Integer
    Dim nFilaSalto As Integer
    Dim nFilaTope As Integer
    Dim nFilaAncho As Integer
    Dim nMargIzq1, nMargIzq2 As Integer
    Dim lcMoneda As String
    Dim nTotPiezas As Integer
    Dim nTotPezoBruto As Double
    Dim nTotPesoNeto As Double
    Dim nTotTasacion As Double
    Dim nContador As Integer
    Dim cCodHojaPie As String
    Dim nAnchoTex As Integer
    Dim nFilaBucle As Integer
    Dim nCortePagina As Integer
    
    Dim loAge1 As COMDConstantes.DCOMAgencias
    Set loAge1 = New COMDConstantes.DCOMAgencias
    Dim lsAgencia1 As String
    Dim rs11 As ADODB.Recordset
    Set rs11 = New ADODB.Recordset
        Set rs11 = loAge1.RecuperaAgencias(gsCodAge)
        If Not (rs11.EOF And rs11.BOF) Then
            lsAgencia1 = Trim(rs11("Dist"))
        End If
    Set loAge1 = Nothing
    
'    Set R = obj.ChequexImpresion(pnId)
    If rsPig.EOF Then
        CargaHojaResumenPignoPDF = False
        Exit Function
    End If

    oPDF.Author = gsCodUser
    oPDF.Creator = "SICMACT - Negocio"
    oPDF.Producer = gsNomCmac
    oPDF.Subject = "HOJA RESUMEN PIGNORATICIO" ' & R!cNroCheque
    oPDF.Title = oPDF.Subject

    If Not oPDF.PDFCreate(App.Path & "\Spooler\HojResCredConGarMobOro_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Function
    End If

    oPDF.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oPDF.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    oPDF.Fonts.Add "F3", "Arial", TrueType, Bold, WinAnsiEncoding
    oPDF.Fonts.Add "F4", "Arial Narrow", TrueType, Bold, WinAnsiEncoding
    oPDF.Fonts.Add "F5", "Arial Narrow", TrueType, Normal, WinAnsiEncoding

    oPDF.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
    oPDF.NewPage A4_Vertical

    'establece medidas
    nMargIzq1 = 20: nMargIzq2 = 25: nFilaSalto = 23: nFilaTope = 800
    nFilaAncho = 12: nAnchoTex = 0
    nTotPiezas = 0: nTotPezoBruto = 0: nTotPesoNeto = 0: nTotTasacion = 0
    'nContador = 0: cCodHojaPie = "HRP.122016"
    'nContador = 0: cCodHojaPie = "HRP.022017"
    'nContador = 0: cCodHojaPie = "HR-COL-CRE-GMJOV01-2019" ' comentado by GEMO 13/03/2020
    nContador = 0: cCodHojaPie = "HR-COL-CRE-GMJOV01-2020"  ' ADD by GEMO 13/03/2020

    '830, 40, 12, 550,
    'oPDF.WImage 55, nMargIzq1, 35, 85, "Logo"
    oPDF.WImage 45, 498, 35, 73, "Logo"
    oPDF.WTextBox 23, 20, 15, 600, "HOJA RESUMEN CRÉDITO CON GARANTIA MOBILIARIA DE ORO", "F4", 12, hCenter
    'INICIO EAAS SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
    oPDF.WTextBox 43, 20, 15, 600, "Crédito N°" & rsPig!cCtaCod & "", "F4", 11, hCenter
    'END EAAS SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
    oPDF.WTextBox 77, nMargIzq1, 12, 600, "El presente documento forma parte integrante del contrato de crédito suscrito entre las partes, EL CLIENTE declara haber sido", "F5", 12, hjustify, , , 0
    oPDF.WTextBox 87, nMargIzq1, 12, 600, "informado sobre las condiciones del crédito que ha solicitado a LA CAJA conforme al siguiente detalle:", "F5", 12, hjustify, , , 0
    oPDF.WTextBox 107, nMargIzq1, 12, 600, "DATOS DEL CRÉDITO:", "F4", 12, hLeft, , , 0
    
    '' DATOS DEL CREDITO
    '1
    oPDF.WTextBox 121, nMargIzq1, 14, 310, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 121, 330, 14, 253, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 121, 330, 14, 253, Format(rsPigCostos!TEAComp, "#,#0.00") & " % (TEA Fija)", "F5", 12, hCenter, , , 0
    '3 TCEA
    oPDF.WTextBox 135, nMargIzq1, 14, 310, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 135, 330, 14, 253, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 135, 330, 14, 253, Format(rsPigCostos!TCEA, "#,#0.00") & " %", "F5", 12, hCenter, , , 0
    '4 tasa de interes morat efectiva anual
    oPDF.WTextBox 149, nMargIzq1, 14, 310, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 149, 330, 14, 253, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 149, 330, 14, 253, Format(rsPigCostos!TEAMora, "#,#0.00") & " % (TEA Fija)", "F5", 12, hCenter, , , 0
    '5 monto del prestamo
    oPDF.WTextBox 163, nMargIzq1, 14, 310, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 163, 330, 14, 253, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 163, 330, 14, 253, rsPigCostos!cMonedaSimbol & " " & Format(rsPig!nMontoCol, "#,#0.00") & " " & rsPigCostos!cMonedaNombre, "F5", 12, hCenter, , , 0
    '6 fecha de desembolso
    oPDF.WTextBox 177, nMargIzq1, 14, 310, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 177, 330, 14, 253, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 177, 330, 14, 253, Format(rsPig!dVigencia, "dd/mm/yyyy"), "F5", 12, hCenter, , , 0
    '7 PLazo
    oPDF.WTextBox 191, nMargIzq1, 14, 310, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 191, 330, 14, 253, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 191, 330, 14, 253, Format(rsPig!nPlazo, "#0") & " Días", "F5", 12, hCenter, , , 0
    '8 fecha de vencimiento
    oPDF.WTextBox 205, nMargIzq1, 14, 310, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 205, 330, 14, 253, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 205, 330, 14, 253, Format(rsPig!dVenc, "dd/mm/yyyy"), "F5", 12, hCenter, , , 0
    '9 monto tot de inter compensa
    oPDF.WTextBox 219, nMargIzq1, 14, 310, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 219, 330, 14, 253, "", "F5", 12, hLeft, , , 1, RGB(0, 0, 0), False
    oPDF.WTextBox 219, 330, 14, 253, rsPigCostos!cMonedaSimbol & " " & Format(rsPigCostos!nInteres, "#,#0.00") & " " & rsPigCostos!cMonedaNombre, "F5", 12, hCenter, , , 0

    'oPDF.WTextBox 107, nMargIzq2, 14, 310, "1.  N° Crédito:", "F5", 12, hLeft, , , 0 COMENTO EAAS SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
    oPDF.WTextBox 121, nMargIzq2, 14, 310, "1.  Tasa de Interés Compensatoria Efectiva Anual(TEA)(¹), 360 días:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox 135, nMargIzq2, 14, 310, "2.  Tasa de Costo Efectivo Anual (TCEA):", "F5", 12, hLeft, , , 0
    oPDF.WTextBox 149, nMargIzq2, 14, 310, "3.  Tasa de interés Moratorio Efectiva Anual, 360 días:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox 163, nMargIzq2, 14, 310, "4.  Monto y moneda del préstamo:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox 177, nMargIzq2, 14, 310, "5.  Fecha de desembolso:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox 191, nMargIzq2, 14, 310, "6.  Plazo:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox 205, nMargIzq2, 14, 310, "7.  Fecha de vencimiento:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox 219, nMargIzq2, 14, 310, "8.  Monto total de intereses compensatorios a pagar:", "F5", 12, hLeft, , , 0
    
    Dim nFilaTop As Integer
    Dim nRazonTexto As Integer
    Dim sMensaje As String
    
    nFilaTop = 219 'Inicio de top
    nRazonTexto = 14 ' razón que aplica al texto libre
    
    If nCampana > 0 Then
        If nCampana = 1 Then
            nFilaTop = nFilaTop + nRazonTexto + 7
            sMensaje = "(¹) Para acceder a la campaña del descuento del 20% y 30% de TEA, deberá realizar la primera renovación de su crédito en un"
            oPDF.WTextBox nFilaTop, nMargIzq1, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto
            sMensaje = " plazo mayor a 07 días después del desembolso y hasta 03 días posteriores a la fecha de vencimiento de acuerdo a las"
            oPDF.WTextBox nFilaTop, nMargIzq1, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto
            sMensaje = " condiciones pactadas en la presente hoja resumen."
            oPDF.WTextBox nFilaTop, nMargIzq1, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto
            sMensaje = "  -   El descuento del 20%, aplica solo en la segunda renovación, " _
                        & "para ello deberá realizar el pago en la fecha de vencimiento " _
                        & "señalada en su voucher, hasta con un máximo de 03 días de atraso."
            oPDF.WTextBox nFilaTop, nMargIzq1 + 5, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto + nRazonTexto
            sMensaje = "  -   El descuento del 30%, aplica solo en la tercera renovación, " _
                        & "para ello deberá realizar el pago en la fecha de vencimiento señalada " _
                        & "en su voucher, hasta con un máximo de 03 días de atraso."
            oPDF.WTextBox nFilaTop, nMargIzq1 + 5, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto + nRazonTexto
            sMensaje = "-   Las siguientes renovaciones serán calculadas a la TEA pactada en " _
                                & "la presente Hoja Resumen."
            oPDF.WTextBox nFilaTop, nMargIzq1 + 5, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto + 7
            sMensaje = "Restricciones: "
            oPDF.WTextBox nFilaTop, nMargIzq1, 30, 534, sMensaje, "F3", 11, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto
            sMensaje = "Si se realiza la renovación, cancelación o amortización de capital antes de las fechas de vencimientos pactadas, se perderán los beneficios de los descuentos de la presente campaña."
            oPDF.WTextBox nFilaTop, nMargIzq1 + 5, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto
            
        ElseIf nCampana = 2 Then
            nFilaTop = nFilaTop + nRazonTexto + 7
            sMensaje = "(¹) La TEA preferencial otorgada por la campaña ""El primer mes no pagas"" solo aplica para la primera renovación del crédito. La renovación deberá ser realizada en la fecha de vencimiento pactada en la presente hoja resumen, hasta con un máximo de 03 días de atraso; las posteriores renovaciones serán calculadas a la TEA del producto según tarifario vigente al momento del desembolso."
            oPDF.WTextBox nFilaTop, nMargIzq1, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto + nRazonTexto + nRazonTexto + nRazonTexto
                       
            sMensaje = "Restricciones: "
            oPDF.WTextBox nFilaTop, nMargIzq1, 30, 534, sMensaje, "F3", 11, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto
            sMensaje = "Si se realiza la renovación, cancelación o amortización de capital antes de la fecha de vencimiento de la primera renovación señalada en numeral 7 de la presente hoja resumen, se perderá el beneficio de la TEA preferencial otorgada, aplicándose la TEA establecida para el producto según el tarifario vigente al momento del desembolso."
            oPDF.WTextBox nFilaTop, nMargIzq1 + 5, 30, 534, sMensaje, "F5", 12, hjustify, hjustify, , 0
            nFilaTop = nFilaTop + nRazonTexto + nRazonTexto
            
        End If
    End If
    
    nFilaTop = nFilaTop + nRazonTexto + 7
    oPDF.WTextBox nFilaTop, nMargIzq1, 12, 250, "9. Tasación:", "F5", 12, hLeft, , , 0
    nFilaTop = nFilaTop + nRazonTexto
    oPDF.WTextBox nFilaTop, nMargIzq1 + 17, 12, 600, "LA CAJA, declara recibir en garantía del préstamo de dinero otorgado, los bienes que a continuación se detallan:", "F5", 12, hLeft, , , 0
    nFilaTop = nFilaTop + nRazonTexto + 7
    'cabecera del primer cuadro
    'pintamos de color gris
    oPDF.WTextBox nFilaTop, nMargIzq1, 30, 57, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57, 30, 273, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273, 30, 45, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45, 30, 58, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45 + 58, 30, 58, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45 + 58 + 58, 30, 72, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    'hacemos los ciadros
    oPDF.WTextBox nFilaTop, nMargIzq1, 30, 57, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57, 30, 273, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273, 30, 45, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45, 30, 58, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45 + 58, 30, 53, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45 + 58 + 53, 30, 77, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    'ponemos el texto
    oPDF.WTextBox nFilaTop, nMargIzq1, 24, 57, "Pieza(s)", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57, 24, 273, "Descripción del bien(es)", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273, 24, 45, "Kilates", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45, 24, 58, "Peso Bruto (g)", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45 + 58, 24, 55, "Peso Oro Neto (g)", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFilaTop, nMargIzq1 + 57 + 273 + 45 + 58 + 53, 24, 72, "Gravamen/Valor Tasación S/", "F4", 12, hCenter, hCenter, , 0
    nFilaTop = nFilaTop + nRazonTexto + 4

    'nFila = 281 Comentado RIRO 20210921
    nFila = nFilaTop
    nFila5 = nFila
'    nValor = 90
    nCortePagina = 0
    
    Dim nnfila As Integer
    rsPigDet.MoveFirst
    Do While Not rsPigDet.EOF
        'nContador = nContador + 1
        'nFila = nFila + 12

        nFila = IIf((nFila + nFilaAncho) >= nFilaTope, nFilaSalto, nFila + nFilaAncho)
        If nFila = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
        If nFila = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
        If nFila = nFilaSalto Then
            nFilaBucle = nFila5 + nFilaAncho
            ' se amplia el cuadro de detalle de acuerdo al numero de registros
            oPDF.WTextBox nFilaBucle, nMargIzq1, 12 + (nContador * 12) + 24, 57, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
            oPDF.WTextBox nFilaBucle, nMargIzq1 + 57, 12 + (nContador * 12) + 24, 273, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
            oPDF.WTextBox nFilaBucle, nMargIzq1 + 57 + 273, 12 + (nContador * 12) + 24, 45, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
            oPDF.WTextBox nFilaBucle, nMargIzq1 + 57 + 273 + 45, 12 + (nContador * 12) + 24, 58, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
            oPDF.WTextBox nFilaBucle, nMargIzq1 + 57 + 273 + 45 + 58, 12 + (nContador * 12) + 24, 53, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
            oPDF.WTextBox nFilaBucle, nMargIzq1 + 57 + 273 + 45 + 58 + 53, 12 + (nContador * 12) + 24, 77, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
            nCortePagina = 1
        End If
        If nFila = nFilaSalto Then oPDF.NewPage A4_Vertical

        'nAnchoTex = IIf((Round((Len(Trim(rsPigDet!cdescrip))) / 55, 0) * 12) > 12, (Round((Len(Trim(rsPigDet!cdescrip))) / 55, 0) * 12), 0)
        nAnchoTex = IIf((Round(Len(Trim(rsPigDet!cdescrip)) / 55 * 12, 0)) > 12, (Round(Len(Trim(rsPigDet!cdescrip)) / 55 * 12, 0)), 0)

        oPDF.WTextBox nFila, nMargIzq1, 12 + nAnchoTex, 67, Format(rsPigDet!nPiezas, "#0"), "F5", 12, hCenter, hCenter, , 0
        oPDF.WTextBox nFila, nMargIzq1 + 67, 12 + nAnchoTex, 250, Trim(Trim(rsPigDet!cdescrip)), "F5", 12, hLeft, hCenter, , 0
        oPDF.WTextBox nFila, nMargIzq1 + 67 + 268, 12 + nAnchoTex, 40, Format(rsPigDet!cKilataje, "#0"), "F5", 12, hCenter, hCenter, , 0
        oPDF.WTextBox nFila, nMargIzq1 + 67 + 268 + 40, 12 + nAnchoTex, 53, Format(rsPigDet!nPesoBruto, "#,#0.00"), "F5", 12, hRight, hCenter, , 0
        oPDF.WTextBox nFila, nMargIzq1 + 67 + 268 + 40 + 53, 12 + nAnchoTex, 53, Format(rsPigDet!nPesoNeto, "#,#0.00"), "F5", 12, hRight, hCenter, , 0
        oPDF.WTextBox nFila, nMargIzq1 + 67 + 268 + 40 + 53 + 53 + 20, 12 + nAnchoTex, 57, Format(rsPigDet!nValTasac, "#,#0.00"), "F5", 12, hRight, hCenter, , 0

        nTotPiezas = nTotPiezas + rsPigDet!nPiezas
        nTotPezoBruto = nTotPezoBruto + rsPigDet!nPesoBruto
        nTotPesoNeto = nTotPesoNeto + rsPigDet!nPesoNeto
        nTotTasacion = nTotTasacion + rsPigDet!nValTasac

        nFila = nFila + nAnchoTex
        'nContador = nContador + IIf(Round((Len(Trim(rsPigDet!cdescrip))) / 55, 0) = 0, 1, Round((Len(Trim(rsPigDet!cdescrip))) / 55, 0))
        nContador = nContador + IIf((Round(Len(Trim(rsPigDet!cdescrip)) / 55, 0)) = 0, 1, (Round(Len(Trim(rsPigDet!cdescrip)) / 55, 0)))

        rsPigDet.MoveNext

    Loop

    nFila2 = nFila

    If nCortePagina = 0 Then
        nFila = nFila5 + nFilaAncho
        ' se amplia el cuadro de detalle de acuerdo al numero de registros
        oPDF.WTextBox nFila, nMargIzq1, 12 + (nContador * 12) + 24, 57, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
        oPDF.WTextBox nFila, nMargIzq1 + 57, 12 + (nContador * 12) + 24, 273, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
        oPDF.WTextBox nFila, nMargIzq1 + 57 + 273, 12 + (nContador * 12) + 24, 45, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
        oPDF.WTextBox nFila, nMargIzq1 + 57 + 273 + 45, 12 + (nContador * 12) + 24, 58, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
        oPDF.WTextBox nFila, nMargIzq1 + 57 + 273 + 45 + 58, 12 + (nContador * 12) + 24, 53, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
        oPDF.WTextBox nFila, nMargIzq1 + 57 + 273 + 45 + 58 + 53, 12 + (nContador * 12) + 24, 77, "", "F5", 14, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    End If
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'sub titulo de totales

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'pintamos de color gris
    oPDF.WTextBox nFila2, nMargIzq1, 30, 57, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFila2, nMargIzq1 + 57, 30, 273, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273, 30, 45, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45, 30, 58, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45 + 58, 30, 58, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45 + 58 + 58, 30, 72, "", "F5", 12, hCenter, hCenter, , 1, RGB(234, 234, 234), True

    'cuadro vacio
    oPDF.WTextBox nFila2, nMargIzq1, 30, 57, "", "F5", 10, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57, 30, 273, "", "F5", 10, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273, 30, 45, "", "F5", 10, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45, 30, 58, "", "F5", 10, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45 + 58, 30, 53, "", "F5", 10, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45 + 58 + 53, 30, 77, "", "F5", 10, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    '' texto para llenar el cudro
    oPDF.WTextBox nFila2, nMargIzq1, 30, 55, "Total Piezas", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 55, 30, 273, "-", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 55 + 273, 30, 45, "-", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 56 + 273 + 45, 30, 55, "Total Peso Bruto (g)", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 56 + 273 + 45 + 55, 30, 55, "Total Peso Neto (g)", "F4", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 56 + 273 + 45 + 55 + 55 + 10, 30, 57, "Total Tasación S/", "F4", 12, hCenter, hCenter, , 0

    'espacio vacios para totales
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 6))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
        
    oPDF.WTextBox nFila2, nMargIzq1, 20, 57, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57, 20, 273, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273, 20, 45, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45, 20, 58, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45 + 58, 20, 53, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45 + 58 + 53, 20, 77, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False

    'rellena los espacios para totales
    oPDF.WTextBox nFila2, nMargIzq1, 12, 57, Format(nTotPiezas, "#0"), "F5", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 57, 12, 273, "-", "F5", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273, 12, 45, "-", "F5", 12, hCenter, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45, 12, 53, Format(nTotPezoBruto, "#,#0.00"), "F5", 12, hRight, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45 + 53, 12, 53, Format(nTotPesoNeto, "#,#0.00"), "F5", 12, hRight, hCenter, , 0
    oPDF.WTextBox nFila2, nMargIzq1 + 57 + 273 + 45 + 53 + 53 + 20, 12, 57, Format(nTotTasacion, "#,#0.00"), "F5", 12, hRight, hCenter, , 0

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'Punto 11
    
    oPDF.WTextBox nFila2, nMargIzq1, 12, 250, "10. Comisiones aplicables:", "F5", 12, hLeft, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 3))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    '1ra fila
    'rellena los espacios de color gris
    oPDF.WTextBox nFila2, nMargIzq1, 14, 300, "", "F4", 12, hCenter, 3, vbWhite, 1, vbRed, True
    oPDF.WTextBox nFila2, nMargIzq1 + 300, 14, 90, "", "F4", 12, hCenter, 3, vbWhite, 1, vbRed, True
    oPDF.WTextBox nFila2, nMargIzq1 + 300 + 90, 14, 173, "", "F4", 12, hCenter, 3, vbWhite, 1, vbRed, True
    'espacios vacios con margenes de lineas
    oPDF.WTextBox nFila2, nMargIzq1, 14, 300, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 300, 14, 90, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 300 + 90, 14, 173, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    'textos
    oPDF.WTextBox nFila2, nMargIzq1, 12, 300, "Comisión", "F4", 12, hCenter, , vbWhite, 0
    oPDF.WTextBox nFila2, nMargIzq1 + 300, 12, 90, "Importe MN", "F4", 12, hCenter, , vbWhite, 0
    oPDF.WTextBox nFila2, nMargIzq1 + 300 + 90, 12, 173, "Oportunidad de Cobro", "F4", 12, hCenter, , vbWhite, 0
       
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 2))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    '2da fila
    'rellena los espacios con el color gris
    oPDF.WTextBox nFila2, nMargIzq1, 14, 563, "", "F4", 9, hLeft, 3, RGB(0, 0, 0), 1, RGB(234, 234, 234), True
    
    'cuadro con espacios vacios
    oPDF.WTextBox nFila2, nMargIzq1, 14, 563, "", "F5", 12, , , , 1, RGB(0, 0, 0), False
    
    'Llena texto
    oPDF.WTextBox nFila2, nMargIzq2, 12, 563, "Servicios brindados a solicitud del cliente: Constancias", "F4", 12, hLeft, , RGB(0, 0, 0), 0
        
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 2))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    '3ra fila
    'rellena los espacios
    oPDF.WTextBox nFila2, nMargIzq2, 12, 300, "Duplicado de contrato u hoja resumen del crédito", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, nMargIzq2 + 300, 12, 90, "S/ " & Format(rsPigCostos!nDupli_contrato, "#,#0.00"), "F5", 12, hCenter, , , 0
    oPDF.WTextBox nFila2, nMargIzq2 + 300 + 90, 12, 173, Trim(rsPigCostos!cOportuCosto1), "F5", 12, hCenter, , , 0
    'espacios vacios
    oPDF.WTextBox nFila2, nMargIzq1, 14, 300, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 300, 14, 90, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 300 + 90, 14, 173, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 2))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    '4ta fila
    'rellena los espacios con el color gris
    oPDF.WTextBox nFila2, nMargIzq1, 14, 563, "", "F4", 9, hLeft, 3, RGB(0, 0, 0), 1, RGB(234, 234, 234), True
    'espacios vacios
    oPDF.WTextBox nFila2, nMargIzq1, 14, 563, "", "F5", 12, , , , 1, RGB(0, 0, 0), False
    'rellena los espacios
    oPDF.WTextBox nFila2, nMargIzq2, 12, 563, "Uso de canales: Uso de módulo electrónico", "F4", 12, hLeft, , RGB(0, 0, 0), 0

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 2))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

' Inicio GEMO
    '3ra fila
    'rellena los espacios
    oPDF.WTextBox nFila2, nMargIzq2, 12, 400, "Caja Maynas Online: Renovación o cancelación de créditos pignoraticios", "F5", 11, hLeft, , , 0
    oPDF.WTextBox nFila2, nMargIzq2 + 300, 12, 90, "Sin costo", "F5", 12, hCenter, , , 0
    oPDF.WTextBox nFila2, nMargIzq2 + 300 + 90, 12, 173, "Al efectuar la operación", "F5", 12, hCenter, , , 0
    
    'espacios vacios
    oPDF.WTextBox nFila2, nMargIzq1, 14, 300, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 300, 14, 90, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 300 + 90, 14, 173, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 2))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
'end gemo
    
    '4ta fila
    'rellena los espacios con el color gris
    oPDF.WTextBox nFila2, nMargIzq1, 14, 563, "", "F4", 9, hLeft, 3, RGB(0, 0, 0), 1, RGB(234, 234, 234), True
    'espacios vacios
    oPDF.WTextBox nFila2, nMargIzq1, 14, 563, "", "F5", 12, , , , 1, RGB(0, 0, 0), False
    'rellena los espacios
    oPDF.WTextBox nFila2, nMargIzq2, 12, 563, "Custodia de Joyas", "F4", 12, hLeft, , RGB(0, 0, 0), 0

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 2))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical


    '5ta fila
    'rellena los espacios
    oPDF.WTextBox nFila2, nMargIzq2, 12, 300, "Comision por custodia de joyas ", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, nMargIzq2 + 300, 12, 90, Format(rsPigCostos!nComi_custodia, "#,#0.00") & " % mensual", "F5", 12, hCenter, , , 0
    oPDF.WTextBox nFila2, nMargIzq2 + 300 + 90, 12, 173, Trim(rsPigCostos!cOportuCosto2), "F5", 12, hCenter, , , 0
    'espacios vacios
    oPDF.WTextBox nFila2, nMargIzq1, 14, 300, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 300, 14, 90, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
    oPDF.WTextBox nFila2, nMargIzq1 + 300 + 90, 14, 173, "", "F5", 12, hCenter, 3, , 1, RGB(0, 0, 0), False
       
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + 3))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'inicio gemo

    oPDF.WTextBox nFila2, nMargIzq1, 12, 551, "Además de las comisiones señaladas en la presente Hoja Resumen, LA CAJA cuenta con comisiones adicionales por Servicios", "F5", 12, hjustify, , , 0
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    
    oPDF.WTextBox nFila2, nMargIzq1, 12, 551, "Transversales cuyo tarifario se encuentra disponible en nuestra red de agencias y en el portal web www.cajamaynas.pe", "F5", 12, hjustify, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    'fin gemo
    
'-----------------------
  'Punto 11
    oPDF.WText nFila2, nMargIzq1, "11.", "F5", 12
    oPDF.WText nFila2, nMargIzq1 + 17, "Ante el incumplimiento de pago según las condiciones pactadas, se procederá a realizar el reporte correspondiente a la", "F5", 12

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "Central de Riesgo de la SBS, con arreglo a las disposiciones del reglamento para la evaluación y clasificación del deudor", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "y la exigencia de provisiones vigentes.", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

'-----------------

    'Punto 12
    oPDF.WText nFila2, nMargIzq1, "12.", "F5", 12
    oPDF.WText nFila2, nMargIzq1 + 17, "EL CLIENTE  tiene  el derecho de efectuar el pago anticipado de la cuota del crédito en  forma total con  la consiguiente", "F5", 12

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "reducción de intereses al día de pago; asimismo, las comisiones y gastos derivados de las cláusulas contractuales pactadas", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "entre las partes. LA CAJA admite  el pago anticipado del préstamo  sin pago de  penalidad, ni  de comisión  por  cancelación", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "anticipada, ni gastos por el mismo concepto.", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'Punto 13
    oPDF.WText nFila2, nMargIzq1, "13.", "F5", 12
    oPDF.WText nFila2, nMargIzq1 + 17, "Al llegar  la fecha de vencimiento  del pago del crédito o antes de  que ocurra el mismo, EL CLIENTE podrá ampliar  el  plazo", "F5", 12

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "del crédito, por igual periodo de 30 días calendario, a partir de la fecha que se realiza el pago,siempre que cumpla con pagar", "F5", 12

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "como " & String(0.55, vbTab) & " mínimo, el " & Trim(CStr(rsPig!nPorceMinK * 100)) & "% del " & String(0.55, vbTab) & " capital, más los " & String(0.55, vbTab) & " gastos, intereses " & String(0.55, vbTab) & " compensatorios e " & String(0.55, vbTab) & " intereses " & String(0.55, vbTab) & " moratorios, correspondientes   ", "F5", 12

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "al plazo pactado y al periodo vencido de ser  el caso, incluyendo los impuestos  correspondientes. Asimismo, cuando la fecha  ", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "de vencimiento recae sobre un día no laborable(domingo o feriado) el pago se realizará el primer día útil siguiente. La  nueva ", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "fecha de  vencimiento constará  en  el voucher  de pago de  la  última  renovación  realizada por EL CLIENTE y emitido por", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    'INICIO EAAS
    oPDF.WText nFila2, nMargIzq1 + 17, "LA CAJA. El plazo total máximo de las renovaciones que puede realizar EL CLIENTE no debe exceder los " & Trim(CStr(rsPig!nVigenciaMesesPigno)) & " meses", "F5", 12
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    oPDF.WText nFila2, nMargIzq1 + 17, "contados a partir de la fecha de desembolso del crédito.", "F5", 12
    'FIN EAAS
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    'Punto 16
    oPDF.WText nFila2, nMargIzq1, "14.", "F5", 12
    oPDF.WText nFila2, nMargIzq1 + 17, "Una  vez  cancelada  la  obligación,  y  EL(LOS) BIEN(ES) recibido(s)  en garantía,  no sea(n) recogido(s)  por  EL  CLIENTE,", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "quedará(n) en poder de LA CAJA en condición de lote diferido, si transcurren treinta(30) días calendario y EL(LOS) BIEN(ES)", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "continua(n) sin ser recogido(s), el lote generará una comisión de custodia, a partir del día 31.", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'Punto 17
    oPDF.WText nFila2, nMargIzq1, "15.", "F5", 12
    oPDF.WText nFila2, nMargIzq1 + 17, "El(los) bien(es) otorgado(s) en garantía y cuya(s) obligación(es) que garantizaba(n), ya se encuentra(n) cancelada(s), será(n)", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "entregado(s) a partir del día hábil siguiente. El  recojo  de dicho(s) bien(es) dejado(s) en  garantía, se efectuará en la agencia", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "u oficina de LA CAJA  donde fue otorgado el crédito  y se hará sólo previa identificación  de  EL CLIENTE  con el Documento", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "Oficial  de Identidad - DOI (DNI, Pasaporte, Carnet  de  Extranjería), y  presentación  del contrato  y  Hoja  Resumen  original", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "respectivo; Si EL CLIENTE no puede recoger personalmente EL(LOS) BIEN(ES) que fueron otorgados en garantía,  lo podrá", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "realizar a través de un representante legal acreditado.", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'Punto 18
    oPDF.WText nFila2, nMargIzq1, "16.", "F5", 12
    oPDF.WText nFila2, nMargIzq1 + 17, "Todas las operaciones se encuentran gravados  por el impuesto a  las transacciones financieras (ITF) excepto en  el caso de", "F5", 12

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WText nFila2, nMargIzq1 + 17, "cuentas exoneradas u operaciones inafectas de acuerdo a Ley. La tasa vigente del ITF es de " & CStr(rsPigCostos!nTasaITF) & " %.", "F5", 12
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    nFila2 = nFila2 - 15
'-----------------------------------------------------

    'Despedida donde se declara  la responsabilidad
    oPDF.WTextBox nFila2, nMargIzq1, 12, 551, "Declaro haber  recibido y leído plenamente el contrato y la Hoja  Resumen del crédito con Garantía Mobiliaria de Oro, por lo que", "F5", 12, hjustify, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, nMargIzq1, 12, 551, "suscribo el presente documento en duplicado y con conocimiento pleno de las condiciones establecidas en dichos documentos.", "F5", 12, hjustify, , , 0
    
    '*** inicio filas vacias
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    '*** fin filas vacias
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    nFila2 = nFila2 - 20
    oPDF.WTextBox nFila2, nMargIzq1, 12, 551, "Lugar y fecha: " & lsAgencia1 & ", " & Day(rsPig!dVigencia) & " de " & rsPig!cMesContrato & " de " & Year(rsPig!dVigencia) & ".", "F5", 12, hjustify, , , 0
    
    '*** inicio filas vacias
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
       
    '*** fin filas vacias

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    nFila3 = IIf((nFila2 + 80) >= nFilaTope, nFila2 - 40, nFila2)
    'cuadro para primera huella digital
    oPDF.WTextBox nFila3, 430, 80, 80, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    nFila2 = nFila2 - 20
    oPDF.WTextBox nFila2, nMargIzq1, 12, 551, "Apellidos y Nombres:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12, 500, rsPigPers!cpersapellido & " " & rsPigPers!cPersNombre, "F5", 12, hLeft, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)  '12
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, nMargIzq1, 12, 551, "D.O.I.:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12, 500, rsPigPers!NroDNI, "F5", 12, hLeft, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)  '12
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    nAnchoTex = IIf((Round((Len(Trim(rsPigPers!cPersDireccDomicilio))) / 55, 0) * 12) > 12, (Round((Len(Trim(rsPigPers!cPersDireccDomicilio))) / 55, 0) * 12), 0)
    
    oPDF.WTextBox nFila2, nMargIzq1, 12, 500, "Dirección:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12 + nAnchoTex, 300, rsPigPers!cPersDireccDomicilio, "F5", 12, hLeft, , , 0
    nFila2 = nFila2 + 12 + nAnchoTex
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)  '12
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, nMargIzq1, 12, 500, "Teléfono:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12, 500, IIf(rsPigPers!cPersTelefono = "", "   --   ", rsPigPers!cPersTelefono) & "        Celular: " & IIf(rsPigPers!cPersCelular = "", "  --  ", rsPigPers!cPersCelular), "F5", 12, hLeft, , , 0

    '----inicio filas vacias
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    '----fin filas vacias
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, nMargIzq1, 12, 500, "Firma:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12, 500, ".................................................................................", "F5", 12, hLeft, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, 150, 12, 500, " ", "F5", 12, hLeft, , , 0

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    '--- inicio filas vacias
    oPDF.WTextBox nFila2, 150, 12, 500, " ", "F5", 12, hLeft, , , 0

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    '--- fin filas vacias
    
    oPDF.WTextBox nFila2, nMargIzq1, 12, 500, "Por la CMAC MAYNAS S.A.:    " & rsPigPers!Agencia, "F5", 12, hLeft, , , 0

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'verifica si el cuadro se mostrará todo si no va a la siguiente hoja
    nFila2 = IIf((nFila2 + nFilaAncho) >= 647, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    'cuadro para llenar cuando se cancela
    oPDF.WTextBox nFila2, nMargIzq1, 180, 553, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, nMargIzq2, 12, 553, "(Solo para ser llenado una vez cancelada la obligación)", "F5", 12, hLeft, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, nMargIzq2, 12, 553, "Declaro recibir la(s) joya(s) dadas en garantía, a mi entera satisfacción y conformidad, al haber cancelado la obligación:", "F5", 12, hLeft, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    nFila4 = nFila2 'para la segunda huella digital
    
    oPDF.WTextBox nFila2, nMargIzq2, 12, 500, "Apellidos y Nombres:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12, 500, ".................................................................................", "F5", 12, hLeft, , , 0

    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, nMargIzq2, 12, 500, "D.O.I.:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12, 500, ".................................................................................", "F5", 12, hLeft, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical

    oPDF.WTextBox nFila2, nMargIzq2, 12, 500, "Fecha:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12, 500, ".................................................................................", "F5", 12, hLeft, , , 0
    
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, nFila2 + nFilaAncho)
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    nFila2 = IIf((nFila2 + nFilaAncho) >= nFilaTope, nFilaSalto, IIf(nFila2 = nFilaSalto, nFila2, nFila2 + nFilaAncho))
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    If nFila2 = nFilaSalto Then oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight
    If nFila2 = nFilaSalto Then oPDF.NewPage A4_Vertical
    
    oPDF.WTextBox nFila2, nMargIzq2, 12, 500, "Firma:", "F5", 12, hLeft, , , 0
    oPDF.WTextBox nFila2, 150, 12, 500, ".................................................................................", "F5", 12, hLeft, , , 0
    
    'cuadro para la huella digital
    oPDF.WTextBox nFila4, 430, 80, 80, "", "F5", 12, hCenter, hCenter, , 1, RGB(0, 0, 0), False
    'pie de pagina de la ultima hoja
    oPDF.WTextBox 830, 40, 12, 550, oPDF.PageCount, "F5", 7, hCenter
    oPDF.WTextBox 830, 40, 12, 550, cCodHojaPie, "F5", 7, hRight

    oPDF.PDFClose
    oPDF.Show
    
    Set R = Nothing
    Set oPDF = Nothing
    Set obj = Nothing

CargaHojaResumenPignoPDF = True
End Function

Public Sub Pig_ContratosAutomaticos(ByRef rsPigPers As ADODB.Recordset, ByRef rsPig As ADODB.Recordset, ByVal lsContrato As String)

    Dim lrPig As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsFecha As String
    Dim lnDiasAtraso As Integer
    Dim lnIntVencido As Double
    Dim lnIntMoratorio As Double
    Dim lnIntAdelantado As Double
    Dim lnDeuda As Double
    Dim nPag As Integer
    Dim nDoc As Integer
    Dim lsArchivo As String
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
    Dim lsModeloPlantilla As String
    Dim lsNomMaq As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim sArchivo As String
        
    Set lrPig = New ADODB.Recordset
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
        lsAgencia = Trim(rs1("Dist"))
    End If
    Set loAge = Nothing

    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\contratopignoraticio" & gsCodAge & ".doc")

    sArchivo = App.Path & "\FormatoCarta\ICP_" & lsContrato & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)
    With oWord.Selection.Find
        .Text = "<<cTitular>>"
        .Replacement.Text = rsPigPers!cpersapellido & "/" & rsPigPers!cPersNombre
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<nNroDNI>>"
        .Replacement.Text = rsPigPers!NroDNI
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<cPersDireccDomicilio>>"
        .Replacement.Text = rsPigPers!cPersDireccDomicilio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oWord.Selection.Find
        .Text = "<<Zona>>"
        .Replacement.Text = lsAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fDay>>"
        .Replacement.Text = Day(rsPig!dVigencia)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False

        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fMes>>"
        .Replacement.Text = rsPig!cMesContrato
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<fAnio>>"
        .Replacement.Text = Year(rsPig!dVigencia)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    oDoc.Close
    Set oDoc = Nothing

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing
End Sub
'END EAAS SEGUN Memorándum Nº 756-2019-GM-DI/CMACM

'*** PEAC 20190408
Public Sub Pig_CartasNotarialesCustodAdju(ByVal dFecSis As Date, ByVal nDiasVenc As Integer, ByVal dFechaCorte As Date, ByVal psNombre As String, _
                                    ByVal psCargo As String, ByVal pImagen As Variant, Optional psNumRemate As String)

    Dim lrPig As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim lsAgencia As String
    Dim lsFecha As String
    Dim lnDiasAtraso As Integer
    Dim lnIntVencido As Double
    Dim lnIntMoratorio As Double
    Dim lnIntAdelantado As Double
    Dim lnDeuda As Double
    Dim nPag As Integer
    Dim nDoc As Integer
    Dim lsArchivo As String
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
    Dim lsModeloPlantilla As String
    Dim lsNomMaq As String
    Dim cFecCanCred As String
    Dim cMone As String
        
    Set lrPig = New ADODB.Recordset
    Set lrPig = PigAdjudica_CartaNotarialAdju(dFecSis, nDiasVenc, dFechaCorte, gsCodUser, lsNomMaq, gsCodAge, psNumRemate)

    If lrPig.BOF And lrPig.EOF Then
        MsgBox "No existe créditos procesados para Notificar.", vbOKOnly, "Atención"
        Exit Sub
    End If
    
    Set loAge = New COMDConstantes.DCOMAgencias
    Set rs1 = New ADODB.Recordset
        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
        If Not (rs1.EOF And rs1.BOF) Then
            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
        End If
    Set loAge = Nothing

    lsModeloPlantilla = App.Path & "\FormatoCarta\CARTANOTADJCUS.doc"

    lsNomMaq = GetMaquinaUsuario
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

   'Crea Nuevo Documento
    wApp.Documents.Add
 
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla

    Set RangeSource = wAppSource.ActiveDocument.Content
        
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
   
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    If Not (lrPig.EOF And lrPig.BOF) Then
        Do Until lrPig.EOF
        
            cMone = IIf(Mid(lrPig!cCtaCod, 9, 1) = "1", "S/", "US$.")
        
            If lrPig!nPrdEstado = 2102 Then
            
                wApp.Application.Selection.TypeParagraph
                wApp.Application.Selection.PasteAndFormat (wdPasteDefault)
                wApp.Application.Selection.InsertBreak
                wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
                wApp.Selection.MoveEnd
    
                lnDiasAtraso = DateDiff("d", lrPig!dVenc, gdFecSis)
                            
                'Ciudad
                With wApp.Selection.Find
                    .Text = "<<Ciudad>>"
                    .Replacement.Text = lsAgencia
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
    
                'Cliente
                With wApp.Selection.Find
                    .Text = "<<Cliente>>"
                    .Replacement.Text = Trim(lrPig!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                'Direccion
                With wApp.Selection.Find
                    .Text = "<<Direccion>>"
                    .Replacement.Text = Trim(lrPig!cPersDireccDomicilio)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                'Fecha Actual
                lsFecha = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
                With wApp.Selection.Find
                    .Text = "<<FechaActC>>"
                    .Replacement.Text = lsFecha
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With

                'Fecha Cancelacion
                cFecCanCred = Mid(lrPig!fec_cancel_difer, 1, 4) + "-" + Mid(lrPig!fec_cancel_difer, 5, 2) + "-" + Mid(lrPig!fec_cancel_difer, 7, 2)
                With wApp.Selection.Find
                    .Text = "<<cFecCancelCred>>"
                    .Replacement.Text = Format(CDate(cFecCanCred), "dd/mm/yyyy")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                'Cuenta
                With wApp.Selection.Find
                    .Text = "<<cNumCred>>"
                    .Replacement.Text = lrPig!cCtaCod
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                'Fecha de Vigencia
                With wApp.Selection.Find
                    .Text = "<<cFecDesemb>>"
                    .Replacement.Text = Format(lrPig!dVigencia, "dd/mm/yyyy")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With

                'Fecha de sistema
                lsFecha = Format(gdFecSis, "dd") & "/" & Format(gdFecSis, "mmmm") & "/" & Format(gdFecSis, "yyyy")
                With wApp.Selection.Find
                    .Text = "<<cFecSistema>>"
                    .Replacement.Text = lsFecha
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                'monto custodia diferida
                With wApp.Selection.Find
                    .Text = "<<cCostoCustod>>"
                    .Replacement.Text = cMone & " " & Format(lrPig!monto_custo, "0.00")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                lnDeuda = Round(lrPig!monto_custo, 2)
                lnDeuda = Round(lnDeuda, 2)
                With wApp.Selection.Find
                    .Text = "<<cDeudaFecha>>"
                    .Replacement.Text = cMone & " " & Format(lnDeuda, "0.00")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                'Tasacion
                With wApp.Selection.Find
                    .Text = "<<cValorBien>>"
                    .Replacement.Text = cMone & " " & Format(lrPig!nTasacion, "0.00")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
    
    
            End If
            lrPig.MoveNext
        Loop
    End If
    Set loCalculos = Nothing
 
'*x*x*x*x*x*x*x*x - pie de pagina - PEAC 20080305
    'crea el pie de pagina en la cual pone la firma nombre y cargo del jefe de agencia
    If wApp.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        wApp.ActiveWindow.Panes(2).Close
    End If
    If wApp.ActiveWindow.ActivePane.View.Type = wdNormalView Or wApp.ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        wApp.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    If wApp.Selection.HeaderFooter.IsHeader = True Then
        wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Else
        wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    End If

    'si por algun motivo no tiene firma el jefe, solo se muestra nombre y cargo
    If CStr(pImagen) <> 0 Then
        Clipboard.Clear ' limpia el portapapeles
        Clipboard.SetData pImagen
        wApp.Selection.Paste
        
        wApp.Selection.InlineShapes.Application.Height = 20 'original 20
        'wApp.Selection.InlineShapes.Application.Application.Application.
        
        'wApp.Selection.InlineShapes.Height = 10 '(80 * CCur(wApp.Selection.InlineShapes.Height)) / 100
        'wApp.Selection.InlineShapes.Width = 10 '(80 * CCur(wApp.Selection.InlineShapes.Width)) / 100
        'Var.Height = (80 * CCur(Var.Height)) / 100
                
        'wApp.Selection.InlineShapes(0).ScaleHeight = 50
        'wApp.Selection.InlineShapes(0).ScaleWidth = 50
        
        'wApp.Selection.InlineShapes(0).ScaleHeight = 50
        'wApp.Selection.InlineShapes(0).ScaleWidth = 50
                
        'wApp.Selection.ShapeRange.LockAspectRatio = False
        'wApp.Selection.ShapeRange.Height = InchesToPoints(0.78)
        'wApp.Selection.ShapeRange.Width = InchesToPoints(0.78)
        
        'wApp.Selection.ShapeRange.ScaleHeight Factor:=0.7, RelativeToOriginalSize:=True     'Factor:=(70 / 100), RelativeToOriginalSize:=True
        'wApp.Selection.ShapeRange.ScaleWidth Factor:=0.7, RelativeToOriginalSize:=True   'Factor:=(70 / 100), RelativeToOriginalSize:=True
        
        Clipboard.Clear ' limpia el portapapeles
    End If

    wApp.Selection.TypeParagraph

    wApp.Selection.TypeText Text:=Trim(psNombre) '"cpersnombre"
    wApp.Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    wApp.Selection.Font.Bold = wdToggle
    wApp.Selection.Range.Case = wdTitleWord
    wApp.Selection.EndKey Unit:=wdLine
    wApp.Selection.TypeParagraph
    
    wApp.Selection.TypeText Text:=Trim(psCargo) '"cargo"
    wApp.Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    wApp.Selection.Font.Size = 10
    wApp.Selection.EndKey Unit:=wdLine
    wApp.Selection.Font.Bold = wdToggle
    wApp.Selection.Font.Size = 12
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeParagraph
    
    wApp.Selection.Font.Size = 6
    wApp.Selection.TypeText Text:="En caso no pueda recoger personalmente el(los) bien(es) que fueron otorgados en garantía, lo " & _
                                "podrá realizar a través de un representante mediante Carta Poder con firma legalizada notarialmente " & _
                                "cuando el(los) bien(es) tengan un valor de tasación que no exceda de 1/4 UIT, vigente al momento del recojo; " & _
                                "o Poder fuera de Registro cuando el(los) bien(es) tenga(n) un valor mayor de 1/4 UIT y que no exceda de 3 UIT, " & _
                                "vigente al momento del recojo. En caso de residir en el extranjero, EL CLIENTE podrá autorizar a un tercero a recoger " & _
                                "el(los) bien(es) mediante poder consular de acuerdo a las formalidades establecidas en el párrafo anterior."
    wApp.Selection.TypeParagraph
    wApp.Selection.TypeText Text:="En el caso de fallecimiento de EL CLIENTE, el(los) bien(es) dejados en garantía podrán ser recogidos por sus herederos, " & _
                                "presentando el testamento o la declaración de sucesión intestada, debidamente inscrito en el registro de testamentos o en el registro " & _
                                "de sucesiones intestadas respectivamente."
    wApp.Selection.WholeStory
    wApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    wApp.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    
    'Modifica margenes de pagina
    wApp.Selection.PageSetup.TopMargin = CentimetersToPoints(1.25)
    wApp.Selection.PageSetup.LeftMargin = CentimetersToPoints(2)
    wApp.Selection.PageSetup.RightMargin = CentimetersToPoints(0.75)
    
'*x*x*x*x*x*x*x*x
 
 wAppSource.ActiveDocument.Close
 wApp.Visible = True

Set wAppSource = Nothing
Set wApp = Nothing

Clipboard.Clear

End Sub
