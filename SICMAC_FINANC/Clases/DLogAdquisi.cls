VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogAdquisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSql As String

Public Enum TpoCargaAdq
    AdqTodosGnral = 5
    AdqTodosEstado = 6
End Enum
Public Enum TpoCargaAdqDet
    AdqDetUnRegistro = 0
    AdqDetUnRegCoti = 1
End Enum

Public Enum TpoCargaSel
    SelUnRegistro = 1
    SelTodosGnral = 5
    SelTodosEstado = 6
End Enum

Public Enum TpoCargaSelCotDet
    SelCotDetUnRegistro = 1
    SelCotDetUnRegEvalua = 5
    SelCotDetUnRegEvaluaTodos = 6
End Enum

Private Sub Class_Initialize()
    Dim oini As ClasIni
    
    Set oini = New ClasIni
        vsConexion = oini.CadenaConexion
        vsCentralPer = oini.BasePersonas
        vsCentralCom = oini.BaseComunes
        vsCentralImg = oini.BaseImagenes
    Set oini = Nothing
End Sub

'Devuelve un RecordSet con LogAdquisicion
Public Function CargaAdquisicion(ByVal psTpoCarga As TpoCargaAdq, ByVal psAdqTpo As String, _
ByVal psPeriodo As String, Optional ByVal psEstado As LogAdqEstado) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    If oConec.AbreConexion(vsConexion) Then
        If psTpoCarga = AdqTodosGnral Then
            sSql = "SELECT a.cLogAdqNro, c.cConsCod, c.cConsDescripcion " & _
                " FROM LogAdquisicion A LEFT JOIN Constante C ON a.cLogAdqEstado = c.cConsValor " & _
                "           AND c.cConsCod = '" & gLogAdqEstado & "' " & _
                " WHERE a.cLogAdqTpo = '" & psAdqTpo & "' AND a.cLogAdqPeriodo = '" & psPeriodo & "' "
        ElseIf psTpoCarga = AdqTodosEstado Then
            If psAdqTpo = "" And psPeriodo = "" Then
                sSql = "SELECT a.cLogReqNro Código, c.cConsDescripcion Estado, a.cLogReqPeriodo Periodo, c2.cConsDescripcion [Tipo Adquisición]" & _
                    " FROM LogRequerimiento A LEFT JOIN Constante C ON '1' = c.cConsValor " & _
                    "           AND c.cConsCod = '" & gLogAdqEstado & "' " & _
                    "       LEFT JOIN Constante C2 ON a.cLogReqTpo = c2.cConsValor " & _
                    "           AND c2.cConsCod = '" & gLogReqTipo & "' " & _
                    " "
            Else
                sSql = "SELECT cLogAdqNro " & _
                    " FROM LogAdquisicion " & _
                    " WHERE cLogAdqTpo = '" & psAdqTpo & "' AND cLogAdqPeriodo = '" & psPeriodo & "' " & _
                    "       AND cLogAdqEstado = '" & psEstado & "'"
            End If
        Else
            sSql = ""
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaAdquisicion = Rs
End Function

'Devuelve un RecordSet con LogAdqDetalle
Public Function CargaAdqDetalle(ByVal psTpoCarga As TpoCargaAdqDet, ByVal psAdqNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If psTpoCarga = AdqDetUnRegistro Then
            sSql = "SELECT ad.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       15, ad.nLogReqDetrefPrecio, " & _
                "       (15 * ad.nLogReqDetrefPrecio) SubTotal" & _
                " FROM LogReqDetalle AD LEFT JOIN BienesServicios BS ON ad.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE ad.cLogReqNro IN ('" & psAdqNro & "') " & _
                " ORDER BY ad.cBSCod "
        ElseIf psTpoCarga = AdqDetUnRegCoti Then
            sSql = "SELECT ad.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       20, 100 " & _
                " FROM LogreqDetalle AD LEFT JOIN BienesServicios BS ON ad.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE ad.cLogReqNro IN ('" & psAdqNro & "') " & _
                " ORDER BY ad.cBSCod "
        Else
            sSql = ""
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaAdqDetalle = Rs
End Function


'Devuelve un RecordSet con LogSeleccion
Public Function CargaSeleccion(ByVal psTpoCarga As TpoCargaSel, Optional ByVal psSelNro As String, _
Optional ByVal psEstado As LogSelEstado, Optional ByVal psEstSec As LogSelEstado) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If psTpoCarga = SelUnRegistro Then
            sSql = "SELECT s.cLogSelNro, s.dLogSelRes, s.cLogSelResNro, s.cLogSelEstado, " & _
                "       s.cAreaCod, a.cAreaDescripcion, s.cPersCod, p.cPersNombre, " & _
                "       s.nLogSelCostoBase, IsNull(s.cLogSelNro,'') cLogAdqNro, " & _
                "       isnull(s.cLogSelSisAdj,'') cLogSelSisAdj, IsNull(c.cConsDescripcion,'') cConsDescripcion, " & _
                "       IsNull(s.cLogSelCotNro,'') cLogSelCotNro " & _
                " FROM LogSeleccion S LEFT JOIN Areas A ON s.cAreaCod = a.cAreaCod " & _
                "       LEFT JOIN " & vsCentralPer & "Persona P ON s.cPersCod = p.cPersCod " & _
                "       LEFT JOIN Constante C ON s.cLogSelSisAdj = c.cConsValor " & _
                "           AND c.cConsCod = '" & gLogSelSisAdj & "' " & _
                " WHERE s.cLogSelNro = '" & psSelNro & "'"
        ElseIf psTpoCarga = SelTodosGnral Then
            sSql = "SELECT s.cLogSelNro Código, IsNull(s.cLogSelResNro,'') Resolución, c.cConsDescripcion Estado " & _
                " FROM LogSeleccion S LEFT JOIN Constante C ON s.cLogSelEstado = c.cConsValor " & _
                "           AND c.cConsCod = '" & gLogSelEstado & "' " & _
                " WHERE NOT c.cConsValor IN ('" & psEstado & "','" & psEstSec & "')"
        ElseIf psTpoCarga = SelTodosEstado Then
            sSql = "SELECT s.cLogSelNro Código, s.cLogSelResNro Resolución, c.cConsDescripcion Estado " & _
                " FROM LogSeleccion S LEFT JOIN Constante C ON s.cLogSelEstado = c.cConsValor " & _
                "           AND c.cConsCod = '" & gLogSelEstado & "' " & _
                " WHERE cLogSelEstado IN ('" & psEstado & "','" & psEstSec & "')"
        Else
            sSql = ""
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSeleccion = Rs
End Function

'Devuelve un RecordSet con LogSelComite
Public Function CargaSelComite(ByVal psSelNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT sc.cAreaCod, a.cAreaDescripcion, " & _
            "       sc.cPersCod, p.cPersNombre " & _
            " FROM LogSelComite SC LEFT JOIN Areas A ON sc.cAreaCod = a.cAreaCod " & _
            "       LEFT JOIN " & vsCentralPer & "Persona P ON sc.cPersCod = p.cPersCod " & _
            " WHERE sc.cLogSelNro = '" & psSelNro & "'"
        
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelComite = Rs
End Function

'Devuelve un RecordSet con LogSelParametro
Public Function CargaSelParametro(ByVal psSelNro As String, ByVal psSelParTpo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT sp.nLogSelParNro, c.cConsDescripcion, sp.nLogSelParPuntaje " & _
            " FROM LogSelParametro SP LEFT JOIN Constante C ON sp.nLogSelParNro = c.cConsValor " & _
            "           AND c.cConsCod = '" & IIf(psSelParTpo = "1", gLogSelParTec, gLogSelParEco) & "' " & _
            " WHERE sp.cLogSelNro = '" & psSelNro & "' AND sp.cLogSelParTpo = '" & psSelParTpo & "' "
        
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelParametro = Rs
End Function

'Devuelve un RecordSet con LogSelPostor
Public Function CargaSelPostor(ByVal psSelNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        sSql = "SELECT sp.cPersCod, p.cPersNombre, isnull(sp.cLogSelPosObserva,'') cLogSelPosObserva " & _
            " FROM LogSelPostor SP LEFT JOIN " & vsCentralPer & "Persona P ON sp.cPersCod = p.cPersCod " & _
            " WHERE sp.cLogSelNro = '" & psSelNro & "' "
        
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelPostor = Rs
End Function

'Devuelve un RecordSet con LogSelPublica
Public Function CargaSelPublica(ByVal psSelNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then

        sSql = "SELECT sp.cPersCod, p.cPersNombre, Convert(varchar(12),sp.dLogSelPubIni,103), Convert(varchar(12),sp.dLogSelPubFin,103) " & _
            " FROM LogSelPublica SP LEFT JOIN " & vsCentralPer & "Persona P ON sp.cPersCod = p.cPersCod " & _
            " WHERE sp.cLogSelNro = '" & psSelNro & "' "
        
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelPublica = Rs
End Function

'Devuelve un RecordSet con LogSelCotiza
Public Function CargaSelCotiza(ByVal psSelNro As String, Optional ByVal psSelCotNro As String = "") As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If psSelCotNro = "" Then
            sSql = "SELECT sc.cLogSelCotNro, sc.cPersCod, p.cPersNombre " & _
                " FROM LogSelCotiza SC LEFT JOIN " & vsCentralPer & "Persona P ON sc.cPersCod = p.cPersCod " & _
                " WHERE sc.cLogSelNro = '" & psSelNro & "' "
        Else
            sSql = "SELECT sc.cLogSelCotNro, sc.cPersCod, p.cPersNombre " & _
                " FROM LogSelCotiza SC LEFT JOIN " & vsCentralPer & "Persona P ON sc.cPersCod = p.cPersCod " & _
                " WHERE sc.cLogSelNro = '" & psSelNro & "' AND sc.cLogSelCotNro = '" & psSelCotNro & "'"
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelCotiza = Rs
End Function

'Devuelve un RecordSet con LogSelCotDetalle
Public Function CargaSelCotDetalle(ByVal psTpoCarga As TpoCargaSelCotDet, _
ByVal psSelCotNro As String, Optional ByVal psSelNro As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim Rs As ADODB.Recordset
    Set oConec = New DConecta
    Set Rs = New ADODB.Recordset
    
    If oConec.AbreConexion(vsConexion) Then
        If psTpoCarga = SelCotDetUnRegistro Then
            sSql = "SELECT scd.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       ad.nLogAdqDetCantidad, ad.nLogAdqDetPrecio, " & _
                "       (ad.nLogAdqDetCantidad * ad.nLogAdqDetPrecio) Total, " & _
                "       scd.nLogSelCotDetCantidad, scd.nLogSelCotDetPrecio, " & _
                "       (scd.nLogSelCotDetCantidad * scd.nLogSelCotDetPrecio) TotalProv" & _
                " FROM LogSelCotDetalle SCD LEFT JOIN BienesServicios BS ON scd.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                "       JOIN LogSelCotiza SC ON scd.cLogSelCotNro = sc.cLogSelCotNro " & _
                "       JOIN LogSeleccion S ON sc.cLogSelNro = s.cLogSelNro " & _
                "       JOIN LogAdqDetalle AD ON s.cLogAdqNro = ad.cLogAdqNro " & _
                " WHERE scd.cLogSelCotNro = '" & psSelCotNro & "' " & _
                "       AND ad.cBSCod = scd.cBSCod "
        ElseIf psTpoCarga = SelCotDetUnRegEvalua Then
            sSql = "SELECT ad.cBSCod, bs.cBSDescripcion, " & _
                "       IsNull((ltrim(cu.cConsDescripcion) + space(40) + Convert(char(1),bs.nBSUnidad)),'') cConsUnidad, " & _
                "       ad.nLogAdqDetCantidad, ad.nLogAdqDetPrecio, " & _
                "       (ad.nLogAdqDetCantidad * ad.nLogAdqDetPrecio) Total " & _
                " FROM LogSeleccion S JOIN LogAdqDetalle AD ON s.cLogAdqNro = ad.cLogAdqNro " & _
                "       LEFT JOIN BienesServicios BS ON ad.cBSCod = bs.cBSCod " & _
                "       LEFT JOIN Constante CU ON bs.nBSUnidad = cu.cConsValor And cu.cConsCod = " & gUnidadMedida & " " & _
                " WHERE s.cLogSelNro = '" & psSelNro & "' " & _
                " ORDER BY ad.cBSCod "
        ElseIf psTpoCarga = SelCotDetUnRegEvaluaTodos Then
            sSql = "SELECT sc.cLogSelCotNro, scd.cBSCod, " & _
                " scd.nLogSelCotDetCantidad, scd.nLogSelCotDetPrecio, " & _
                " (scd.nLogSelCotDetCantidad * scd.nLogSelCotDetPrecio) Total" & _
                " FROM LogSelCotiza SC JOIN LogSelCotDetalle SCD ON sc.cLogSelCotNro = scd.cLogSelCotNro " & _
                " WHERE sc.cLogSelNro = '" & psSelNro & "'" & _
                " ORDER BY sc.cLogSelNro, sc.cLogSelCotNro, scd.cBSCod  "
        End If
        Set Rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    Set CargaSelCotDetalle = Rs
End Function

