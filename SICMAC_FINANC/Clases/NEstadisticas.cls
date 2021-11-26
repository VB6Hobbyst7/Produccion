VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset


'Public Function GetEstadisticaAhorro(pnBitCentral As Boolean, pdFecIni As Date, pdFecFin As Date, psMoneda As String, Optional psAgeCod As String = "") As ADODB.Recordset
'
'Dim lsWhere1 As String
'Dim lsWhere2 As String
'
'If pnBitCentral = True Then
'
'    lsWhere1 = " nMoneda=" & Val(psMoneda) & " and nProducto=" & Producto.gCapAhorros & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
'    lsWhere2 = " Cs.nMoneda=" & Val(psMoneda) & " and CS.nProducto=" & Producto.gCapAhorros & IIf(psAgeCod = "", "", " and CS.cCodAge ='" & psAgeCod & "'")
'
'
'    If pdFecIni = pdFecFin Then
'
'        psSql = "select convert(datetime, CI.dEstad,101) as dEstadAC, "
'        psSql = psSql & " isnull(C1.nSaldoAC,0) as nSaldoSinCMAC, isnull(C2.nMonChqAC,0) as nMonChqAC, "
'        psSql = psSql & " isnull(C3.nSaldoCMAC,0) AS nSaldoCMAC, isnull(C4.nChqCMAC,0) as nChqCMAC, "
'        psSql = psSql & " isnull(C5.nMonChqVal,0) as nMonChqVal, isnull(C6.nSaldCRAC,0) as nSaldCRAC, "
'        psSql = psSql & " isnull(C7.nChqCRAC,0) as nChqCRAC, "
'        psSql = psSql & " isnull(C9.nSaldCOOP,0) as nSaldCOOP, "
'        psSql = psSql & " isnull(C10.nSaldEdiPyme,0) as nSaldEdiPyme "
'
'        'Cabecera
'        psSql = psSql & " From ( select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
'                " WHERE " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
'
'        'Saldo de Ahorro de Todos menos CMACS - CRACS - COOPERATIVAS si debe ir  and nPersoneria<>" & gPersonaJurCFLCooperativa & "
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoAC From CapEstadSaldo "
'        psSql = psSql & " where (nPersoneria Not In (" & gPersonaJurCFLCMAC & "," & gPersonaJurCFLCRAC & ",6)) AND " & lsWhere1 & " "
'        psSql = psSql & " GROUP BY Convert(Varchar(8),dEstad,112)) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
'
'        '        " where nPersoneria IN(" & gPersonaNat & ", " & gPersonaJurSFL & ", " & gPersonaJurCFL & ") AND " & lsWhere1 & " "
'
'        'Cheques Ingresados en el dia de todas las personerias
'        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nMonChqAC From CapEstadSaldo CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where " & lsWhere2 & " " & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112)) C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) LEFT JOIN "
'
'        'Saldo de CMACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112)) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) LEFT JOIN "
'
'        'Cheques de CMACS ingresados en el dia
'        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC From CapEstadSaldo  CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) LEFT JOIN "
'
'        'Cheques en Valorizacion de toda las personerias
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal From CapEstadSaldo " & _
'                " where " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) LEFT JOIN "
'
'        'Saldo de CRACS
'        psSql = psSql & " ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC From CapEstadSaldo " & _
'                " where nPersoneria In (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C6 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C6.dEstad,112) "
'
'        'Saldo de Cooperativas
'        psSql = psSql & " LEFT JOIN  ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCOOP From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLCooperativa & " and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C9 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C9.dEstad,112) "
'
'        'Saldo de edypymes
'        psSql = psSql & " LEFT JOIN ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldEdiPyme From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLEdpyme & " and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C10 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C10.dEstad,112) "
'
'        'Cheques de CRACS ingresados en el dia
'        psSql = psSql & " LEFT JOIN " & _
'                " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where CS.nPersoneria In (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere2 & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) "
'
'        psSql = psSql & " WHERE CI.dEstad=( SELECT CONVERT(VARCHAR(8),MAX(CS.dEstad),112)  " & _
'                " FROM CapEstadSaldo CS " & _
'                " WHERE datediff(d, CS.dEstad, '" & Format(pdFecIni, gsFormatoFecha) & "') >=0  AND " & lsWhere2 & ")"
'    Else
'        psSql = "select convert(datetime, CI.dEstad,101) as dEstadAC, "
'        psSql = psSql & " isnull(C1.nSaldoAC,0) as nSaldoSinCMAC, "
'        psSql = psSql & " isnull(C2.nMonChqAC,0) as nMonChqAC, isnull(C3.nSaldoCMAC,0) AS nSaldoCMAC, "
'        psSql = psSql & " isnull(C4.nChqCMAC,0) as nChqCMAC, isnull(C5.nMonChqVal,0) as nMonChqVal, "
'        psSql = psSql & " isnull(C6.nSaldCRAC,0) as nSaldCRAC, isnull(C7.nChqCRAC,0) as nChqCRAC, "
'        psSql = psSql & " isnull(C9.nSaldCOOP,0) as nSaldCOOP, "
'        psSql = psSql & " isnull(C10.nSaldEdiPyme,0) as nSaldEdiPyme "
'
'        psSql = psSql & " From ( select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
'                " WHERE " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
'
'        'Ahorro de Todos menos CMACS - CRACS  - COOPERATIVAS si debe ir  and nPersoneria<>" & gPersonaJurCFLCooperativa & "
'
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoAC From CapEstadSaldo " & _
'                " where ( nPersoneria Not In (" & gPersonaJurCFLCMAC & "," & gPersonaJurCFLCRAC & ",6)) AND " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112)) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
'
'        'Cheques Ingresados en el dia de todas las personerias
'        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nMonChqAC From CapEstadSaldo CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where " & lsWhere2 & " " & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112)) C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) LEFT JOIN "
'
'        'Saldo de CMACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112)) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) LEFT JOIN "
'
'        'Cheques de CMACS ingresados en el dia
'        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC From CapEstadSaldo  CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) LEFT JOIN "
'
'        'Cheques en Valorizacion de toda las personerias
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal From CapEstadSaldo " & _
'                " where " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) LEFT JOIN "
'
'        'Saldo de CRACS
'        psSql = psSql & " ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC From CapEstadSaldo " & _
'                " where nPersoneria  in (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C6 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C6.dEstad,112) LEFT JOIN "
'
'        'Saldo de Cooperativas
'        psSql = psSql & " ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCOOP From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLCooperativa & " and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C9 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C9.dEstad,112) LEFT JOIN "
'
'        'Saldo de edypymes
'        psSql = psSql & " ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldEdiPyme From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLEdpyme & " and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C10 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C10.dEstad,112) "
'
'        'Cheques de CRACS ingresados en el dia
'        psSql = psSql & " LEFT JOIN " & _
'                " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where CS.nPersoneria In (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere2 & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) "
'
'        psSql = psSql & " WHERE convert(varchar(8),CI.dEstad,112) BETWEEN '" & Format(pdFecIni, gsFormatoMovFecha) & "' AND " & _
'                " '" & Format(pdFecFin, gsFormatoMovFecha) & "' Order By dEstadAC "
'    End If
'Else
'    'Para Leer del Distribuido Negocio
'    If pdFecIni = pdFecFin Then
'        psSql = " SELECT convert(datetime,convert(varchar(10),dEstadAC,102)) dEstadAC, ISNULL(SUM(nSaldoAC),0) nSaldoAC, ISNULL(SUM(nMonChqAC),0) nMonChqAC, ISNULL(SUM(nSaldCMAC),0) nSaldoCMAC, ISNULL(SUM(nChqCMAC),0) nChqCMAC, ISNULL(SUM(nMonChqVal),0) nMonChqVal, ISNULL(SUM(nSaldCRAC),0) nSaldCRAC  FROM EstadDiaACConsol " _
'          & " WHERE CONVERT(varchar(10),dEstadAC,101) = ( SELECT  CONVERT(VARCHAR(10),MAX(AH1.dEstadAC),101) " _
'          & "                                             FROM    ESTADDIAACconsol AH1 " _
'          & "                                             WHERE   datediff(d,Ah1.dEstadAC,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " _
'          & "                                                     AND AH1.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and AH1.cCodAge='" & psAgeCod & "'") & ") " _
'          & " AND cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and cCodage='" & psAgeCod & "'") & " GROUP BY convert(datetime,convert(varchar(10),dEstadAC,102)) "
'    Else
'        psSql = "SELECT convert(datetime,convert(varchar(10),dEstadAC,102)) dEstadAC, nSaldoAC, nMonChqAC, nSaldCMAC nSaldoCMAC, nChqCMAC, nMonChqVal, ISNULL(nSaldCRAC,0) nSaldCRAC " _
'             & "FROM   EstadDiaACConsol " _
'             & "WHERE  dEstadAC BETWEEN '" & Format(pdFecIni, gsFormatoFecha) & "' AND " _
'             & "       '" & Format(pdFecFin + 1, gsFormatoFecha) & "' AND cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
'    End If
'End If
'
'Set GetEstadisticaAhorro = dbConec.CargaRecordSet(psSql)
'
'End Function

Public Function GetEstadisticaAhorro(pnBitCentral As Boolean, pdFecIni As Date, pdFecFin As Date, psMoneda As String, Optional psAgeCod As String = "") As ADODB.Recordset

Dim lsWhere1 As String
Dim lsWhere2 As String

If pnBitCentral = True Then

    lsWhere1 = " nMoneda=" & Val(psMoneda) & " and nProducto=" & Producto.gCapAhorros & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
    lsWhere2 = " Cs.nMoneda=" & Val(psMoneda) & " and CS.nProducto=" & Producto.gCapAhorros & IIf(psAgeCod = "", "", " and CS.cCodAge ='" & psAgeCod & "'")
    
 
    If pdFecIni = pdFecFin Then
        
        psSql = "select convert(datetime, CI.dEstad,101) as dEstadAC, "
        psSql = psSql & " isnull(C1.nSaldoAC,0) as nSaldoSinCMAC, isnull(C2.nMonChqAC,0) as nMonChqAC, "
        psSql = psSql & " isnull(C3.nSaldoCMAC,0) AS nSaldoCMAC, isnull(C4.nChqCMAC,0) as nChqCMAC, "
        psSql = psSql & " isnull(C5.nMonChqVal,0) as nMonChqVal, isnull(C6.nSaldCRAC,0) as nSaldCRAC, "
        psSql = psSql & " isnull(C7.nChqCRAC,0) as nChqCRAC, "
        psSql = psSql & " isnull(C9.nSaldCOOP,0) as nSaldCOOP, "
        psSql = psSql & " isnull(C10.nSaldEdiPyme,0) as nSaldEdiPyme "
        
        'Cabecera
        psSql = psSql & " From ( select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
                " WHERE " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
        
        'Saldo de Ahorro de Todos menos CMACS - CRACS - COOPERATIVAS si debe ir  and nPersoneria<>" & gPersonaJurCFLCooperativa & "
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoAC From CapEstadSaldo "
        psSql = psSql & " where (nPersoneria Not In (" & gPersonaJurCFLCMAC & "," & gPersonaJurCFLCRAC & ",6,7)) AND " & lsWhere1 & " "
        psSql = psSql & " GROUP BY Convert(Varchar(8),dEstad,112)) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
        
        '        " where nPersoneria IN(" & gPersonaNat & ", " & gPersonaJurSFL & ", " & gPersonaJurCFL & ") AND " & lsWhere1 & " "
                
        'Cheques Ingresados en el dia de todas las personerias
        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nMonChqAC From CapEstadSaldo CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where " & lsWhere2 & " " & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112)) C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) LEFT JOIN "
                
        'Saldo de CMACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112)) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) LEFT JOIN "
                
        'Cheques de CMACS ingresados en el dia
        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC From CapEstadSaldo  CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) LEFT JOIN "
                
        'Cheques en Valorizacion de toda las personerias
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal From CapEstadSaldo " & _
                " where " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) LEFT JOIN "

        'Saldo de CRACS
        psSql = psSql & " ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC From CapEstadSaldo " & _
                " where nPersoneria In (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C6 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C6.dEstad,112) "
        
        'Saldo de Cooperativas
        psSql = psSql & " LEFT JOIN  ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCOOP From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLCooperativa & " and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C9 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C9.dEstad,112) "
        
        'Saldo de edypymes
        psSql = psSql & " LEFT JOIN ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldEdiPyme From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLEdpyme & " and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C10 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C10.dEstad,112) "
                
        'Cheques de CRACS ingresados en el dia
        psSql = psSql & " LEFT JOIN " & _
                " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where CS.nPersoneria In (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere2 & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) "
        
        psSql = psSql & " WHERE CI.dEstad=( SELECT CONVERT(VARCHAR(8),MAX(CS.dEstad),112)  " & _
                " FROM CapEstadSaldo CS " & _
                " WHERE datediff(d, CS.dEstad, '" & Format(pdFecIni, gsFormatoFecha) & "') >=0  AND " & lsWhere2 & ")"
    Else
        psSql = "select convert(datetime, CI.dEstad,101) as dEstadAC, "
        psSql = psSql & " isnull(C1.nSaldoAC,0) as nSaldoSinCMAC, "
        psSql = psSql & " isnull(C2.nMonChqAC,0) as nMonChqAC, isnull(C3.nSaldoCMAC,0) AS nSaldoCMAC, "
        psSql = psSql & " isnull(C4.nChqCMAC,0) as nChqCMAC, isnull(C5.nMonChqVal,0) as nMonChqVal, "
        psSql = psSql & " isnull(C6.nSaldCRAC,0) as nSaldCRAC, isnull(C7.nChqCRAC,0) as nChqCRAC, "
        psSql = psSql & " isnull(C9.nSaldCOOP,0) as nSaldCOOP, "
        psSql = psSql & " isnull(C10.nSaldEdiPyme,0) as nSaldEdiPyme "
        
        psSql = psSql & " From ( select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
                " WHERE " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
                
        'Ahorro de Todos menos CMACS - CRACS  - COOPERATIVAS si debe ir  and nPersoneria<>" & gPersonaJurCFLCooperativa & "
        
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoAC From CapEstadSaldo " & _
                " where ( nPersoneria Not In (" & gPersonaJurCFLCMAC & "," & gPersonaJurCFLCRAC & ",6,7)) AND " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112)) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
                
        'Cheques Ingresados en el dia de todas las personerias
        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nMonChqAC From CapEstadSaldo CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where " & lsWhere2 & " " & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112)) C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) LEFT JOIN "
                
        'Saldo de CMACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112)) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) LEFT JOIN "
                
        'Cheques de CMACS ingresados en el dia
        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC From CapEstadSaldo  CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) LEFT JOIN "
                
        'Cheques en Valorizacion de toda las personerias
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal From CapEstadSaldo " & _
                " where " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) LEFT JOIN "

        'Saldo de CRACS
        psSql = psSql & " ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC From CapEstadSaldo " & _
                " where nPersoneria  in (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C6 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C6.dEstad,112) LEFT JOIN "
        
        'Saldo de Cooperativas
        psSql = psSql & " ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCOOP From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLCooperativa & " and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C9 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C9.dEstad,112) LEFT JOIN "
        
        'Saldo de edypymes
        psSql = psSql & " ( select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldEdiPyme From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLEdpyme & " and " & lsWhere1 & " GROUP BY Convert(Varchar(8),dEstad,112) )C10 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C10.dEstad,112) "
        
        'Cheques de CRACS ingresados en el dia
        psSql = psSql & " LEFT JOIN " & _
                " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where CS.nPersoneria In (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere2 & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) "
                
        psSql = psSql & " WHERE convert(varchar(8),CI.dEstad,112) BETWEEN '" & Format(pdFecIni, gsFormatoMovFecha) & "' AND " & _
                " '" & Format(pdFecFin, gsFormatoMovFecha) & "' Order By dEstadAC "
    End If
Else
    'Para Leer del Distribuido Negocio
    If pdFecIni = pdFecFin Then
        psSql = " SELECT convert(datetime,convert(varchar(10),dEstadAC,102)) dEstadAC, ISNULL(SUM(nSaldoAC),0) nSaldoAC, ISNULL(SUM(nMonChqAC),0) nMonChqAC, ISNULL(SUM(nSaldCMAC),0) nSaldoCMAC, ISNULL(SUM(nChqCMAC),0) nChqCMAC, ISNULL(SUM(nMonChqVal),0) nMonChqVal, ISNULL(SUM(nSaldCRAC),0) nSaldCRAC  FROM EstadDiaACConsol " _
          & " WHERE CONVERT(varchar(10),dEstadAC,101) = ( SELECT  CONVERT(VARCHAR(10),MAX(AH1.dEstadAC),101) " _
          & "                                             FROM    ESTADDIAACconsol AH1 " _
          & "                                             WHERE   datediff(d,Ah1.dEstadAC,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " _
          & "                                                     AND AH1.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and AH1.cCodAge='" & psAgeCod & "'") & ") " _
          & " AND cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and cCodage='" & psAgeCod & "'") & " GROUP BY convert(datetime,convert(varchar(10),dEstadAC,102)) "
    Else
        psSql = "SELECT convert(datetime,convert(varchar(10),dEstadAC,102)) dEstadAC, nSaldoAC, nMonChqAC, nSaldCMAC nSaldoCMAC, nChqCMAC, nMonChqVal, ISNULL(nSaldCRAC,0) nSaldCRAC " _
             & "FROM   EstadDiaACConsol " _
             & "WHERE  dEstadAC BETWEEN '" & Format(pdFecIni, gsFormatoFecha) & "' AND " _
             & "       '" & Format(pdFecFin + 1, gsFormatoFecha) & "' AND cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
    End If
End If

Set GetEstadisticaAhorro = dbConec.CargaRecordSet(psSql)

End Function


'Public Function GetEstadisticaPlazoFijo(pnBitCentral As Boolean, pdFecIni As Date, pdFecFin As Date, psMoneda As String, Optional psAgeCod As String = "") As ADODB.Recordset
'
'Dim lsWhere1 As String
'Dim lsWhere2 As String
'
'lsWhere1 = " nMoneda=" & Val(psMoneda) & " and nProducto=" & Producto.gCapPlazoFijo & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
'lsWhere2 = " Cs.nMoneda=" & Val(psMoneda) & " and CS.nProducto=" & Producto.gCapPlazoFijo & IIf(psAgeCod = "", "", " and CS.cCodAge ='" & psAgeCod & "'")
'
'If pnBitCentral = True Then
'    If pdFecIni = pdFecFin Then
'
'        psSql = " Select convert(datetime, CI.dEstad,101) as dEstadPF, ISNULL(C1.nSaldoPF,0) as nSaldoSinCMAC, ISNULL(C4.nMonChqVal,0) as nMonChqVal, ISNULL(C8.nChqCMAC,0) AS nChqCMAC, iSNULL(C3.nSaldoCMAC,0) AS nSaldoCMAC, ISNULL(C5.nSaldCRAC,0) AS nSaldCRAC, ISNULL(C7.nChqCRAC,0) AS nChqCRAC " & _
'                " From ( select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
'                " WHERE " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) ) CI LEFT JOIN "
'
'        'PF de todos menos CMACS - CRACS '- COOPERATIVAS & " and nPersoneria<>" & gPersonaJurCFLCooperativa & "
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoPF From CapEstadSaldo " & _
'                " where (nPersoneria <>" & gPersonaJurCFLCMAC & " and nPersoneria<>" & gPersonaJurCFLCRAC & ") and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
'
'        'Saldo de CMACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) LEFT JOIN "
'
'        'Cheques en Valorizacion de toda las personerias
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal From CapEstadSaldo " & _
'                " where " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) )C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) LEFT JOIN "
'
'        'Saldo de CRACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) LEFT JOIN "
'
'        'Cheques de CRACS ingresados en el dia
'        psSql = psSql & " " & _
'                " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where CS.nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere2 & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) LEFT JOIN "
'
'        'Cheques de CRACS ingresados en el dia
'        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC From CapEstadSaldo  CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C8 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C8.dEstad,112) "
'
'        psSql = psSql & " WHERE CI.dEstad=( SELECT  CONVERT(VARCHAR(8),MAX(CS.dEstad),112) FROM CapEstadSaldo CS " & _
'                "WHERE datediff(d, CS.dEstad,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " & _
'                " AND " & lsWhere2 & " )"
'
'    Else
'        psSql = " Select convert(datetime, CI.dEstad,101) as dEstadPF, dbo.GetSaldoCtaAcumulado(convert(datetime, CI.dEstad,101),(Select top 1 cCtaContCod from OpeCta where cOpeCod = '46" & psMoneda & "411' and cOpeCtaOrden = 5) + '%' + '" & psAgeCod & "',1) nIntDev , ISNULL(C1.nSaldoPF,0) as nSaldoSinCMAC, ISNULL(C4.nMonChqVal,0) as nMonChqVal, ISNULL(C8.nChqCMAC,0) AS nChqCMAC, ISNULL(C3.nSaldoCMAC,0) AS nSaldoCMAC, ISNULL(C5.nSaldCRAC,0) AS nSaldCRAC, ISNULL(C7.nChqCRAC,0) AS nChqCRAC " & _
'                " From ( select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
'                " WHERE " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) ) CI LEFT JOIN "
'
'        'PF de todos menos CMACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoPF From CapEstadSaldo " & _
'                " where nPersoneria Not In (" & gPersonaJurCFLCMAC & ", " & gPersonaJurCFLCRAC & ",6) And " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
'
'        'Saldo de CMACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) LEFT JOIN "
'
'        'Cheques en Valorizacion de toda las personerias
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal From CapEstadSaldo " & _
'                " where " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) )C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) LEFT JOIN "
'
'        'Saldo de CRACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC From CapEstadSaldo " & _
'                " where nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) LEFT JOIN "
'
'        'Cheques de CRACS ingresados en el dia
'        psSql = psSql & " " & _
'                " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where CS.nPersoneria  in (" & gPersonaJurCFLCRAC & ",6) and " & lsWhere2 & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) LEFT JOIN "
'
'        'Cheques de CRACS ingresados en el dia
'        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC From CapEstadSaldo  CS " & _
'                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
'                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
'                " where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & _
'                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C8 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C8.dEstad,112) "
'
'        psSql = psSql & " WHERE Convert(varchar(8), CI.dEstad,112) BETWEEN '" & Format(pdFecIni, gsFormatoMovFecha) & "' " & _
'                " AND '" & Format(pdFecFin, gsFormatoMovFecha) & "' Order by dEstadPF "
'
'    End If
'Else
'    'Para Leer del Distribuido Negocio
'    If pdFecIni = pdFecFin Then
'        psSql = " SELECT Convert(datetime,Convert(varchar(10),dEstadPF,102)) dEstadPF, ISNULL(SUM(PF.nSaldoPF),0) AS nSaldoPF, ISNULL(SUM(PF.nMonChqVal),0) as nMonChqVal, 0 nChqCMAC, ISNULL(SUM(nNumCMAC),0) AS nNumCMAC , ISNULL(SUM(nSaldCMAC),0) AS nSaldoCMAC, ISNULL(SUM(nSaldCRAC),0) nSaldCRAC " _
'            & " FROM     ESTADDIAPFCONSOL PF " _
'            & " WHERE    CONVERT(varchar(10),dEstadPF,101) = (  SELECT  CONVERT(VARCHAR(10),MAX(PF1.dEstadPF),101) " _
'            & "                                                 FROM    ESTADDIAPFCONSOL PF1 " _
'            & "                                                 WHERE   datediff(d,PF1.dEstadPF,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " _
'            & "                                                         AND PF1.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and PF1.cCodAge='" & psAgeCod & "'") & ") " _
'            & "         AND PF.cMoneda='" & psMoneda & "' " & IIf(psAgeCod = "", "", " and cCodage='" & psAgeCod & "'") & " GROUP BY Convert(datetime,Convert(varchar(10),dEstadPF,102)) "
'    Else
'        psSql = " SELECT Convert(datetime,Convert(varchar(10),dEstadPF,102)) dEstadPF, nSaldoPF, nMonChqVal , nNumCMAC, Isnull(nSaldCMAC,0) as nSaldoCMAC, Isnull(nSaldCRAC,0) as nSaldCRAC " _
'                & "FROM  EstadDiaPFConsol   " _
'                & " WHERE  dEstadPF  BETWEEN '" & Format(pdFecIni, gsFormatoFecha) & "'  " _
'                & "       AND '" & Format(pdFecFin + 1, gsFormatoFecha) & "' AND cMoneda = '" & psMoneda & "' and cCodAge ='" & psAgeCod & "'"
'    End If
'End If
'Set GetEstadisticaPlazoFijo = dbConec.CargaRecordSet(psSql)
'End Function

Public Function GetEstadisticaPlazoFijo(pnBitCentral As Boolean, pdFecIni As Date, pdFecFin As Date, psMoneda As String, Optional psAgeCod As String = "") As ADODB.Recordset
 
Dim lsWhere1 As String
Dim lsWhere2 As String

lsWhere1 = " nMoneda=" & Val(psMoneda) & " and nProducto=" & Producto.gCapPlazoFijo & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
lsWhere2 = " Cs.nMoneda=" & Val(psMoneda) & " and CS.nProducto=" & Producto.gCapPlazoFijo & IIf(psAgeCod = "", "", " and CS.cCodAge ='" & psAgeCod & "'")

If pnBitCentral = True Then
    If pdFecIni = pdFecFin Then
        
        psSql = " Select convert(datetime, CI.dEstad,101) as dEstadPF, dbo.GetSaldoCtaAcumulado(convert(datetime, CI.dEstad,101),(Select top 1 cCtaContCod from OpeCta where cOpeCod = '46" & psMoneda & "411' and cOpeCtaOrden = 5) + '%' + '" & psAgeCod & "', 1)  /  Case When " & psMoneda & " = 1 Then 1 Else dbo.GetTipoCambioFinMes(1,CI.dEstad) End nIntDev,   ISNULL(C1.nSaldoPF,0) as nSaldoSinCMAC, ISNULL(C4.nMonChqVal,0) as nMonChqVal, ISNULL(C8.nChqCMAC,0) AS nChqCMAC, iSNULL(C3.nSaldoCMAC,0) AS nSaldoCMAC, ISNULL(C5.nSaldCRAC,0) AS nSaldCRAC, ISNULL(C7.nChqCRAC,0) AS nChqCRAC " & _
                " From ( select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
                " WHERE " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) ) CI LEFT JOIN "
                
        'PF de todos menos CMACS - CRACS '- COOPERATIVAS & " and nPersoneria<>" & gPersonaJurCFLCooperativa & "
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoPF From CapEstadSaldo " & _
                " where nPersoneria Not In (" & gPersonaJurCFLCMAC & "," & gPersonaJurCFLCRAC & ",6,7) and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
                
        'Saldo de CMACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) LEFT JOIN "
            
        'Cheques en Valorizacion de toda las personerias
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal From CapEstadSaldo " & _
                " where " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) )C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) LEFT JOIN "
        
        'Saldo de CRACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) LEFT JOIN "
                
        'Cheques de CRACS ingresados en el dia
        psSql = psSql & " " & _
                " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where CS.nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere2 & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) LEFT JOIN "
                
        'Cheques de CRACS ingresados en el dia
        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC From CapEstadSaldo  CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C8 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C8.dEstad,112) "
                
        psSql = psSql & " WHERE CI.dEstad=( SELECT  CONVERT(VARCHAR(8),MAX(CS.dEstad),112) FROM CapEstadSaldo CS " & _
                "WHERE datediff(d, CS.dEstad,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " & _
                " AND " & lsWhere2 & " )"
        
    Else
        psSql = " Select convert(datetime, CI.dEstad,101) as dEstadPF, dbo.GetSaldoCtaAcumulado(convert(datetime, CI.dEstad,101),(Select top 1 cCtaContCod from OpeCta where cOpeCod = '46" & psMoneda & "411' and cOpeCtaOrden = 5) + '%' + '" & psAgeCod & "', 1)  /  Case When " & psMoneda & " = 1 Then 1 Else dbo.GetTipoCambioFinMes(1,CI.dEstad) End nIntDev , ISNULL(C1.nSaldoPF,0) as nSaldoSinCMAC, ISNULL(C4.nMonChqVal,0) as nMonChqVal, ISNULL(C8.nChqCMAC,0) AS nChqCMAC, ISNULL(C3.nSaldoCMAC,0) AS nSaldoCMAC, ISNULL(C5.nSaldCRAC,0) AS nSaldCRAC, ISNULL(C7.nChqCRAC,0) AS nChqCRAC " & _
                " From ( select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
                " WHERE " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) ) CI LEFT JOIN "
                
        'PF de todos menos CMACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoPF From CapEstadSaldo " & _
                " where nPersoneria Not In (" & gPersonaJurCFLCMAC & "," & gPersonaJurCFLCRAC & ",6,7) and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
                
        'Saldo de CMACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) LEFT JOIN "
            
        'Cheques en Valorizacion de toda las personerias
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal From CapEstadSaldo " & _
                " where " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) )C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) LEFT JOIN "
        
        'Saldo de CRACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC From CapEstadSaldo " & _
                " where nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) LEFT JOIN "
                
        'Cheques de CRACS ingresados en el dia
        psSql = psSql & " " & _
                " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where CS.nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere2 & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) LEFT JOIN "
                
        'Cheques de CRACS ingresados en el dia
        psSql = psSql & " ( Select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC From CapEstadSaldo  CS " & _
                " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND " & _
                " Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda " & _
                " where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & _
                " GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C8 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C8.dEstad,112) "
                
        psSql = psSql & " WHERE Convert(varchar(8), CI.dEstad,112) BETWEEN '" & Format(pdFecIni, gsFormatoMovFecha) & "' " & _
                " AND '" & Format(pdFecFin, gsFormatoMovFecha) & "' Order by dEstadPF "
                
    End If
Else
    'Para Leer del Distribuido Negocio
    If pdFecIni = pdFecFin Then
        psSql = " SELECT Convert(datetime,Convert(varchar(10),dEstadPF,102)) dEstadPF, ISNULL(SUM(PF.nSaldoPF),0) AS nSaldoPF, ISNULL(SUM(PF.nMonChqVal),0) as nMonChqVal, 0 nChqCMAC, ISNULL(SUM(nNumCMAC),0) AS nNumCMAC , ISNULL(SUM(nSaldCMAC),0) AS nSaldoCMAC, ISNULL(SUM(nSaldCRAC),0) nSaldCRAC " _
            & " FROM     ESTADDIAPFCONSOL PF " _
            & " WHERE    CONVERT(varchar(10),dEstadPF,101) = (  SELECT  CONVERT(VARCHAR(10),MAX(PF1.dEstadPF),101) " _
            & "                                                 FROM    ESTADDIAPFCONSOL PF1 " _
            & "                                                 WHERE   datediff(d,PF1.dEstadPF,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " _
            & "                                                         AND PF1.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and PF1.cCodAge='" & psAgeCod & "'") & ") " _
            & "         AND PF.cMoneda='" & psMoneda & "' " & IIf(psAgeCod = "", "", " and cCodage='" & psAgeCod & "'") & " GROUP BY Convert(datetime,Convert(varchar(10),dEstadPF,102)) "
    Else
        psSql = " SELECT Convert(datetime,Convert(varchar(10),dEstadPF,102)) dEstadPF, nSaldoPF, nMonChqVal , nNumCMAC, Isnull(nSaldCMAC,0) as nSaldoCMAC, Isnull(nSaldCRAC,0) as nSaldCRAC " _
                & "FROM  EstadDiaPFConsol   " _
                & " WHERE  dEstadPF  BETWEEN '" & Format(pdFecIni, gsFormatoFecha) & "'  " _
                & "       AND '" & Format(pdFecFin + 1, gsFormatoFecha) & "' AND cMoneda = '" & psMoneda & "' and cCodAge ='" & psAgeCod & "'"
    End If
End If
Set GetEstadisticaPlazoFijo = dbConec.CargaRecordSet(psSql)
End Function

'Public Function GetEstadisticaCTS(pnBitCentral As Boolean, pdFecIni As Date, pdFecFin As Date, psMoneda As String, Optional psAgeCod As String = "") As ADODB.Recordset
'
'Dim lsWhere1 As String
'Dim lsWhere2 As String
'
'lsWhere1 = " nMoneda=" & Val(psMoneda) & " and nProducto=" & Producto.gCapCTS & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
'lsWhere2 = " Cs.nMoneda=" & Val(psMoneda) & " and CS.nProducto=" & Producto.gCapCTS & IIf(psAgeCod = "", "", " and CS.cCodAge ='" & psAgeCod & "'")
'
'If pnBitCentral = True Then
'    If pdFecIni = pdFecFin Then
'
'        psSql = " select convert(datetime, CI.dEstad,101) as dEstadCTS, ISNULL(C1.nSaldo,0) as nSaldoSinCMAC, ISNULL(C2.MontoChq,0) AS MontoChq, 0 As nNumCMAC, 0 as nSaldoCMAC, 0 as nSaldCRAC " & _
'                " From ( Select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
'                " WHERE " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
'
'        'Saldos de CTS de todo menos CMACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldo " & _
'                " From CapEstadSaldo " & _
'                " where nPersoneria <>" & gPersonaJurCFLCMAC & " and nPersoneria<>" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
'
'        'Cheques en Valorizacion de toda las personerias
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) MontoChq From CapEstadSaldo " & _
'                " where " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112))C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) "
'
'        psSql = psSql & " WHERE CI.dEstad=( SELECT  CONVERT(VARCHAR(8),MAX(CS.dEstad),112) " & _
'                " FROM CapEstadSaldo CS " & _
'                " WHERE datediff(d, CS.dEstad,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " & _
'                " AND " & lsWhere2 & " )"
'
'    Else
'        'Probado Pepe
'        psSql = " select convert(datetime, CI.dEstad,101) as dEstadCTS, ISNULL(C1.nSaldo,0) as nSaldoSinCMAC, ISNULL(C2.MontoChq,0) AS MontoChq, 0 As nNumCMAC, 0 as nSaldoCMAC, 0 as nSaldCRAC " & _
'                " From ( Select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
'                " WHERE " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
'
'        'Saldos de CTS de todo menos CMACS
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldo " & _
'                " From CapEstadSaldo " & _
'                " where nPersoneria <>" & gPersonaJurCFLCMAC & " and nPersoneria<>" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
'
'        'Cheques en Valorizacion de toda las personerias
'        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) MontoChq From CapEstadSaldo " & _
'                " where " & lsWhere1 & " " & _
'                " GROUP BY Convert(Varchar(8),dEstad,112))C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) "
'
'        psSql = psSql & " WHERE Convert(varchar(8), CI.dEstad, 112) BETWEEN '" & Format(pdFecIni, gsFormatoMovFecha) & "'  " & _
'                " AND '" & Format(pdFecFin, gsFormatoMovFecha) & "' Order by dEstadCTS "
'
'    End If
'Else
'    'Para Leer del Distribuido Negocio
'    If pdFecIni = pdFecFin Then
'        psSql = " SELECT  Convert(datetime,Convert(varchar(10),dEstadCTS,102)) dEstadCTS, ISNULL(SUM(CTS.nSaldoCTS),0) AS nSaldo, ISNULL(SUM(CTS.nMonChqVal),0) as MontoChq , 0 AS nNumCMAC , 0 AS nSaldoCMAC, 0 nSaldCRAC " _
'            & " FROM     ESTADDIACTSCONSOL CTS " _
'            & " WHERE    CONVERT(VarChar(10),dEstadCTS,101) =(   SELECT  CONVERT(VARCHAR(10),MAX(CTS1.dEstadCTS),101) " _
'            & "                                                 FROM    ESTADDIACTSCONSOL CTS1 " _
'            & "                                                 WHERE   Datediff(d,CTS1.dEstadCTS,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " _
'            & "                                                         AND CTS1.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and CTS1.cCodAge='" & psAgeCod & "'") & " ) " _
'            & "         AND CTS.cMoneda = '" & psMoneda & "'" & IIf(psAgeCod = "", "", " and cCodage='" & psAgeCod & "'") & " GROUP BY Convert(datetime,Convert(varchar(10),dEstadCTS,102)) "
'    Else
'        psSql = "SELECT Convert(datetime,Convert(varchar(10),dEstadCTS,102)) dEstadCTS, (nSaldoCTS) AS nSaldo, nMonChqVal FROM  EstadDiaCTSConsol  " _
'             & "WHERE dEstadCTS BETWEEN '" & Format(pdFecIni, gsFormatoFecha) & "'  " _
'             & "AND '" & Format(pdFecFin + 1, gsFormatoFecha) & "' AND cMoneda = '" & psMoneda & "' and cCodAge ='" & psAgeCod & "'"
'    End If
'End If
'Set GetEstadisticaCTS = dbConec.CargaRecordSet(psSql)
'End Function

Public Function GetEstadisticaCTS(pnBitCentral As Boolean, pdFecIni As Date, pdFecFin As Date, psMoneda As String, Optional psAgeCod As String = "") As ADODB.Recordset

Dim lsWhere1 As String
Dim lsWhere2 As String

lsWhere1 = " nMoneda=" & Val(psMoneda) & " and nProducto=" & Producto.gCapCTS & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
lsWhere2 = " Cs.nMoneda=" & Val(psMoneda) & " and CS.nProducto=" & Producto.gCapCTS & IIf(psAgeCod = "", "", " and CS.cCodAge ='" & psAgeCod & "'")

If pnBitCentral = True Then
    If pdFecIni = pdFecFin Then
    
        psSql = " select convert(datetime, CI.dEstad,101) as dEstadCTS, ISNULL(C1.nSaldo,0) as nSaldoSinCMAC, ISNULL(C2.MontoChq,0) AS MontoChq, 0 As nNumCMAC, 0 as nSaldoCMAC, 0 as nSaldCRAC " & _
                " From ( Select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
                " WHERE " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
                
        'Saldos de CTS de todo menos CMACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldo " & _
                " From CapEstadSaldo " & _
                " where nPersoneria <>" & gPersonaJurCFLCMAC & " and nPersoneria<>" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
                
        'Cheques en Valorizacion de toda las personerias
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) MontoChq From CapEstadSaldo " & _
                " where " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112))C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) "
                
        psSql = psSql & " WHERE CI.dEstad=( SELECT  CONVERT(VARCHAR(8),MAX(CS.dEstad),112) " & _
                " FROM CapEstadSaldo CS " & _
                " WHERE datediff(d, CS.dEstad,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " & _
                " AND " & lsWhere2 & " )"

    Else
        'Probado Pepe
        psSql = " select convert(datetime, CI.dEstad,101) as dEstadCTS, ISNULL(C1.nSaldo,0) as nSaldoSinCMAC, ISNULL(C2.MontoChq,0) AS MontoChq, 0 As nNumCMAC, 0 as nSaldoCMAC, 0 as nSaldCRAC " & _
                " From ( Select Convert(Varchar(8),dEstad,112) as dEstad from dbo.capestadsaldo " & _
                " WHERE " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
                
        'Saldos de CTS de todo menos CMACS
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldo " & _
                " From CapEstadSaldo " & _
                " where nPersoneria <>" & gPersonaJurCFLCMAC & " and nPersoneria<>" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) LEFT JOIN "
        
        'Cheques en Valorizacion de toda las personerias
        psSql = psSql & " ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) MontoChq From CapEstadSaldo " & _
                " where " & lsWhere1 & " " & _
                " GROUP BY Convert(Varchar(8),dEstad,112))C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) "

        psSql = psSql & " WHERE Convert(varchar(8), CI.dEstad, 112) BETWEEN '" & Format(pdFecIni, gsFormatoMovFecha) & "'  " & _
                " AND '" & Format(pdFecFin, gsFormatoMovFecha) & "' Order by dEstadCTS "

    End If
Else
    'Para Leer del Distribuido Negocio
    If pdFecIni = pdFecFin Then
        psSql = " SELECT  Convert(datetime,Convert(varchar(10),dEstadCTS,102)) dEstadCTS, ISNULL(SUM(CTS.nSaldoCTS),0) AS nSaldo, ISNULL(SUM(CTS.nMonChqVal),0) as MontoChq , 0 AS nNumCMAC , 0 AS nSaldoCMAC, 0 nSaldCRAC " _
            & " FROM     ESTADDIACTSCONSOL CTS " _
            & " WHERE    CONVERT(VarChar(10),dEstadCTS,101) =(   SELECT  CONVERT(VARCHAR(10),MAX(CTS1.dEstadCTS),101) " _
            & "                                                 FROM    ESTADDIACTSCONSOL CTS1 " _
            & "                                                 WHERE   Datediff(d,CTS1.dEstadCTS,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " _
            & "                                                         AND CTS1.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and CTS1.cCodAge='" & psAgeCod & "'") & " ) " _
            & "         AND CTS.cMoneda = '" & psMoneda & "'" & IIf(psAgeCod = "", "", " and cCodage='" & psAgeCod & "'") & " GROUP BY Convert(datetime,Convert(varchar(10),dEstadCTS,102)) "
    Else
        psSql = "SELECT Convert(datetime,Convert(varchar(10),dEstadCTS,102)) dEstadCTS, (nSaldoCTS) AS nSaldo, nMonChqVal FROM  EstadDiaCTSConsol  " _
             & "WHERE dEstadCTS BETWEEN '" & Format(pdFecIni, gsFormatoFecha) & "'  " _
             & "AND '" & Format(pdFecFin + 1, gsFormatoFecha) & "' AND cMoneda = '" & psMoneda & "' and cCodAge ='" & psAgeCod & "'"
    End If
End If
Set GetEstadisticaCTS = dbConec.CargaRecordSet(psSql)
End Function

Public Function GetEstadisticaPlazoFijoCts(pnBitCentral As Boolean, pdFecIni As Date, pdFecFin As Date, psMoneda As String, Optional psAgeCod As String = "") As ADODB.Recordset
 
Dim lsWhere1 As String
Dim lsWhere2 As String

If pnBitCentral = True Then
    lsWhere1 = " nMoneda=" & Val(psMoneda) & "  " & IIf(psAgeCod = "", "", " and cCodAge ='" & psAgeCod & "'")
    lsWhere2 = " Cs.nMoneda=" & Val(psMoneda) & "  " & IIf(psAgeCod = "", "", " and CS.cCodAge ='" & psAgeCod & "'")
   
   ' and nPersoneria<>" & gPersonaJurCFLCooperativa & "
     'Probado por Pepe OK
    psSql = " select  isnull(sum(C.Saldo),0) as Saldo, isnull(Sum(C.nChqCMAC),0) as nChqCMAC, isnull(sum(nchqCRAC),0) as nChqCRAC, isnull(sum(nSaldEdiPyme),0) as nSaldEdiPyme,"
    psSql = psSql & " isnull(SUM(C.nMonChqVal),0) as nMonChqVal, "
    psSql = psSql & " isnull(SUM(C.nNumCMAC),0) as nNumCMAC, isnull(SUM(C.nSaldoCMAC),0) as nSaldoCMAC, "
    psSql = psSql & " isnull(SUM(C.nSaldCRAC),0) as nSaldCRAC, ISNULL(SUM(C.nSaldCoop),0) AS nSaldCoop "
    psSql = psSql & " From ( select C1.nSaldoPF AS Saldo, C2.nChqCMAC, C4.nMonChqVal,  C6.nNumCMAC, C3.nSaldoCMAC, C5.nSaldCRAC, C7.nChqCRAC, C8.nSaldCoop, CE.nSaldEdiPyme "
    psSql = psSql & "         From (  select Convert(Varchar(8),dEstad,112) as dEstad "
    psSql = psSql & "                 from dbo.capestadsaldo "
    psSql = psSql & "                 WHERE " & lsWhere1 & " and nProducto=" & Producto.gCapPlazoFijo
    psSql = psSql & "                 GROUP BY Convert(Varchar(8),dEstad,112) ) CI "
    psSql = psSql & "                 LEFT JOIN "
    psSql = psSql & "               ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoPF "
    psSql = psSql & "                 From   CapEstadSaldo "
    psSql = psSql & "                 where  (nPersoneria <>" & gPersonaJurCFLCMAC & " and nPersoneria<>" & gPersonaJurCFLCRAC & ") and " & lsWhere1 & " "
    psSql = psSql & "                        and (nProducto=" & Producto.gCapPlazoFijo & ") "
    psSql = psSql & "                 GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C1.dEstad,112) "
    psSql = psSql & "                 LEFT JOIN "
    psSql = psSql & "                 ( select Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCMAC "
    psSql = psSql & "                   From  CapEstadSaldo CS "
    psSql = psSql & "                   inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND "
    psSql = psSql & "                         Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda "
    psSql = psSql & "                   where CS.nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere2 & " and (CS.nProducto=" & Producto.gCapPlazoFijo & ") "
    psSql = psSql & "                   GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) "
    psSql = psSql & "                  LEFT JOIN "
    psSql = psSql & "                 ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldoCMAC "
    psSql = psSql & "                   From CapEstadSaldo "
    psSql = psSql & "                   where nPersoneria =" & gPersonaJurCFLCMAC & " and " & lsWhere1 & " and (nProducto=" & Producto.gCapPlazoFijo & ") "
    psSql = psSql & "                   GROUP BY Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) "
    psSql = psSql & "                   LEFT JOIN "
    psSql = psSql & "                 ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) nMonChqVal "
    psSql = psSql & "                   From CapEstadSaldo "
    psSql = psSql & "                   where " & lsWhere1 & " and nProducto=" & Producto.gCapPlazoFijo
    psSql = psSql & "                   GROUP BY Convert(Varchar(8),dEstad,112)) C4 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C4.dEstad,112) "
    psSql = psSql & "                   LEFT JOIN "
    psSql = psSql & "                 ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCRAC "
    psSql = psSql & "                   From CapEstadSaldo "
    psSql = psSql & "                   where nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " and (nProducto=" & Producto.gCapPlazoFijo & ") "
    psSql = psSql & "                   GROUP BY Convert(Varchar(8),dEstad,112) )C5 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C5.dEstad,112) "
    psSql = psSql & "                   LEFT JOIN  "
    psSql = psSql & "                 ( Select  Convert(Varchar(8),dEstad,112) as dEstad, sum(nNumCtas) nNumCMAC "
    psSql = psSql & "                   From    CapEstadSaldo "
    psSql = psSql & "                   Where   nPersoneria=" & gPersonaJurCFLCMAC & "  and " & lsWhere1 & " "
    psSql = psSql & "                           and (nProducto=" & Producto.gCapPlazoFijo & ") "
    psSql = psSql & "                   GROUP BY Convert(Varchar(8),dEstad,112) ) C6 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C6.dEstad,112) "
    psSql = psSql & "                   LEFT JOIN "
    psSql = psSql & "                 ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldCoop "
    psSql = psSql & "                   From   CapEstadSaldo "
    psSql = psSql & "                   where  nPersoneria =" & gPersonaJurCFLCooperativa & " and " & lsWhere1 & " and (nProducto=" & Producto.gCapPlazoFijo & ") "
    psSql = psSql & "                   GROUP  BY Convert(Varchar(8),dEstad,112) ) C8 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C8.dEstad,112) "
    psSql = psSql & "                   LEFT JOIN "
    
    psSql = psSql & "                 ( Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) nSaldEdiPyme "
    psSql = psSql & "                   From   CapEstadSaldo "
    psSql = psSql & "                   where  nPersoneria =" & gPersonaJurCFLEdpyme & " and " & lsWhere1 & " and (nProducto=" & Producto.gCapPlazoFijo & ") "
    psSql = psSql & "                   GROUP  BY Convert(Varchar(8),dEstad,112) ) CE ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),CE.dEstad,112) "
    psSql = psSql & "                   LEFT JOIN "
    
    psSql = psSql & "                 ( Select  Convert(Varchar(8),CS.dEstad,112) as dEstad, sum(CM.nMonChq) nChqCRAC From CapEstadSaldo  CS "
    psSql = psSql & "                           inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge AND "
    psSql = psSql & "                           Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112) And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda "
    psSql = psSql & "                   where   CS.nPersoneria =" & gPersonaJurCFLCRAC & " and " & lsWhere2 & " and (CS.nProducto=" & Producto.gCapPlazoFijo & ") "
    psSql = psSql & "                   GROUP BY Convert(Varchar(8),CS.dEstad,112) ) C7 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C7.dEstad,112) "
    psSql = psSql & "                   WHERE CI.dEstad=(   SELECT CONVERT(VARCHAR(8),MAX(CS.dEstad),112) "
    psSql = psSql & "                                       FROM CapEstadSaldo CS "
    psSql = psSql & "                                       WHERE datediff(d, CS.dEstad, '" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 "
    psSql = psSql & "                                       AND " & lsWhere2 & " AND CS.nProducto=" & Producto.gCapPlazoFijo & ") "
    
    '" WHERE datediff(d,Convert(Varchar(8), CS.dEstad,112),'" & Format(pdFecIni, "yyyymmdd") & "') >= 0 " & _

    'CTS
    psSql = psSql & "       Union All "
    psSql = psSql & "       select  C1.Saldo, isnull(C2.NMonChqVaL,0) as nMonChqVal, 0 As nNumCMAC, 0 as nSaldoCMAC, 0 as nChqCMAC, 0 as nSaldCRAC, 0 as nChqCRAC, 0 as nSaldCoop, 0 as nSaldEdiPyme "
    psSql = psSql & "       From (  Select Convert(Varchar(8),dEstad,112) as dEstad "
    psSql = psSql & "               from dbo.capestadsaldo "
    psSql = psSql & "               WHERE " & lsWhere1 & " and nProducto=" & Producto.gCapCTS
    psSql = psSql & "               GROUP BY Convert(Varchar(8),dEstad,112)) CI LEFT JOIN "
    psSql = psSql & "               (   Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nSaldo) Saldo "
    psSql = psSql & "                   From CapEstadSaldo "
    psSql = psSql & "                   where nPersoneria <>" & gPersonaJurCFLCMAC & " and nPersoneria<>" & gPersonaJurCFLCRAC & " and " & lsWhere1 & " and nProducto=" & Producto.gCapCTS & " "
    psSql = psSql & "                   GROUP BY Convert(Varchar(8),dEstad,112) ) C1 ON Convert(Varchar(8),CI.dEstad,112) = Convert(Varchar(8),C1.dEstad,112) "
    psSql = psSql & "                   LEFT JOIN "
    psSql = psSql & "               (   Select Convert(Varchar(8),dEstad,112) as dEstad, sum(nMonChqVal) NMonChqVal "
    psSql = psSql & "                   From CapEstadSaldo "
    psSql = psSql & "                   where " & lsWhere1 & " and nProducto=" & Producto.gCapCTS
    psSql = psSql & "                   GROUP BY Convert(Varchar(8),dEstad,112) )C2 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C2.dEstad,112) "
    psSql = psSql & "                   WHERE CI.dEstad =(  SELECT CONVERT(VARCHAR(8),MAX(CS.dEstad),112) "
    psSql = psSql & "                                       FROM   CapEstadSaldo CS "
    psSql = psSql & "                                       WHERE  datediff(d, CS.dEstad,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 "
    psSql = psSql & "                                       AND " & lsWhere2 & " and CS.nProducto=" & Producto.gCapCTS & " )) C "
            
Else
    'Para Leer del Distribuido Negocio
    psSql = "SELECT ISNULL(SUM(PF.Saldo),0) AS Saldo, ISNULL(SUM(PF.nMonChqVal),0) nMonChqVal, ISNULL(SUM(nNumCMAC),0) AS nNumCMAC , ISNULL(SUM(nSaldCMAC),0) AS nSaldoCMAC, SUM(nChqCMAC) nChqCMAC, ISNULL(SUM(nSaldCRAC),0) nSaldCRAC FROM " & _
          "( SELECT   ISNULL(SUM(PF.nSaldoPF),0) AS Saldo, ISNULL(SUM(PF.nMonChqVal),0) nMonChqVal, ISNULL(SUM(nNumCMAC),0) AS nNumCMAC , ISNULL(SUM(nSaldCMAC),0) AS nSaldCMAC, 0 nChqCMAC, ISNULL(SUM(nSaldCRAC),0) nSaldCRAC " _
        & " FROM     ESTADDIAPFCONSOL PF " _
        & " WHERE    CONVERT(varchar(10),dEstadPF,101) = (   SELECT  CONVERT(VARCHAR(10),MAX(PF1.dEstadPF),101) " _
        & "                                                 FROM    ESTADDIAPFCONSOL PF1 " _
        & "                                                 WHERE   datediff(d,PF1.dEstadPF,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " _
        & "                                                         AND PF1.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", "and PF1.cCodAge='" & psAgeCod & "'") & " ) " _
        & "         AND PF.cMoneda='" & psMoneda & "' " & IIf(psAgeCod = "", "", " and cCodage='" & psAgeCod & "'") _
        & " UNION ALL " _
        & " SELECT   ISNULL(SUM(CTS.nSaldoCTS),0) AS Saldo, ISNULL(SUM(CTS.nMonChqVal),0) nMonChqVal, 0 AS nNumCMAC , 0 AS nSaldoCMAC, 0 nChqCMAC, 0 nSaldCRAC " _
        & " FROM     ESTADDIACTSCONSOL CTS " _
        & " WHERE    CONVERT(VarChar(10),dEstadCTS,101) =(   SELECT  CONVERT(VARCHAR(10),MAX(CTS1.dEstadCTS),101) " _
        & "                                                 FROM    ESTADDIACTSCONSOL CTS1 " _
        & "                                                 WHERE   Datediff(d,CTS1.dEstadCTS,'" & Format(pdFecIni, gsFormatoFecha) & "') >= 0 " _
        & "                                                         AND CTS1.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and CTS1.cCodAge='" & psAgeCod & "'") & ") " _
        & "         AND CTS.cMoneda = '" & psMoneda & "' " & IIf(psAgeCod = "", "", " and cCodage='" & psAgeCod & "'") & " ) PF "
End If

Set GetEstadisticaPlazoFijoCts = dbConec.CargaRecordSet(psSql)
End Function

'Nuevo
Public Function GetEstadSaldoCheques(ldFecha As Date, lsProducto As String, lsMoneda As String, Optional ByVal pnBitCentral As Boolean = False) As Currency
Dim TotalCheque As Currency
Dim prs As ADODB.Recordset
TotalCheque = 0

If pnBitCentral = True Then
 
    Select Case lsProducto
    
    Case "Plazo"
        psSql = "SELECT sum(nMonChqVal) as Monto "
        psSql = psSql & " FROM CAPESTADSALDO CS "
        psSql = psSql & " WHERE Convert(Varchar(8),CS.dEstad,112) LIKE '" & Format(ldFecha, "YYYYMMdd") & "%' "
        psSql = psSql & " AND NMONEDA=" & lsMoneda & " AND NPRODUCTO in (" & Producto.gCapPlazoFijo & ", " & Producto.gCapCTS & " ) "
        psSql = psSql & " and (nPersoneria <>4 and nPersoneria<>5 ) "
    
    Case "Ahorros"
        psSql = "SELECT sum(nMonChqVal) as Monto "
        psSql = psSql & " FROM CAPESTADSALDO CS "
        psSql = psSql & " WHERE Convert(Varchar(8),CS.dEstad,112) LIKE '" & Format(ldFecha, "YYYYMMdd") & "%' "
        psSql = psSql & " AND NMONEDA=" & lsMoneda & " AND NPRODUCTO in(" & Producto.gCapAhorros & ") "
        psSql = psSql & " and (nPersoneria <>4 and nPersoneria<>5 ) "
        
    End Select
    
Else

   Select Case lsProducto
       Case "Plazo"
               psSql = "SELECT SUM(nMonChqVal) AS Monto " _
                   & "FROM ESTADDIAPFCONSOL WHERE DATEDIFF(DAY,dEstadPF,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' " _
                   & "Union " _
                   & "SELECT SUM(nMonChqVal) AS Monto " _
                   & "FROM ESTADDIACTSCONSOL WHERE DATEDIFF(DAY,dEstadcts,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' "
           
       Case "Ahorros"
               psSql = "SELECT dEstadAC, SUM(nMonChqVal+nChqCMAC) as Monto " _
                   & "FROM ESTADDIAACCONSOL WHERE DATEDIFF(DAY,dEstadAC,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' " _
                   & "GROUP BY  dEstadAC"
   End Select
End If
Set prs = dbConec.CargaRecordSet(psSql)
Do While Not prs.EOF
   TotalCheque = TotalCheque + IIf(IsNull(prs!Monto), 0, prs!Monto)
   prs.MoveNext
Loop
prs.Close
GetEstadSaldoCheques = TotalCheque
End Function

Public Function EXGetEstadSaldoCheques(ldFecha As Date, lsProducto As String, lsMoneda As String, Optional ByVal pnBitCentral As Boolean = False) As Currency
Dim TotalCheque As Currency
Dim prs As ADODB.Recordset
TotalCheque = 0
Select Case lsProducto
    Case "Plazo"
        'OK
        If pnBitCentral = True Then
            psSql = " Select Sum(nMonChqVal) as Monto From CapEstadSaldo " & _
                    " Where nMoneda=" & Val(lsMoneda) & " And nProducto=" & Producto.gCapPlazoFijo & " AND Convert(Varchar(8),dEstad,112)='" & Format(ldFecha, "YYYYMMdd") & "' " & _
                    " GROUP BY Convert(Varchar(8),dEstad,112) " & _
                    "   " & _
                    " Union " & _
                    " Select Sum(nMonChqVal) as Monto From CapEstadSaldo " & _
                    " Where nMoneda=" & Val(lsMoneda) & " And nProducto=" & Producto.gCapCTS & " AND Convert(Varchar(8),dEstad,112)='" & Format(ldFecha, "YYYYMMdd") & "' " & _
                    " GROUP BY Convert(Varchar(8),dEstad,112) " & _
                    " "
        Else
            psSql = "SELECT SUM(nMonChqVal) AS Monto " _
                & "FROM ESTADDIAPFCONSOL WHERE DATEDIFF(DAY,dEstadPF,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' " _
                & "Union " _
                & "SELECT SUM(nMonChqVal) AS Monto " _
                & "FROM ESTADDIACTSCONSOL WHERE DATEDIFF(DAY,dEstadcts,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' "
        End If
        
    Case "Ahorros"
        If pnBitCentral = True Then
            psSql = "Select sum(CS.nMonChqVal) + sum(CM.nMonChq) as Monto "
            psSql = psSql & " From CapEstadSaldo CS "
            psSql = psSql & " inner join capestadmovimiento CM ON CS.cCodAge=CM.cCodAge"
            psSql = psSql & " AND Convert(Varchar(8),CS.dEstad,112) = Convert(Varchar(8),CM.dEstad,112)"
            psSql = psSql & " And CS.nProducto = CM.nProducto And CS.nMoneda = CM.nMoneda"
            psSql = psSql & " Where CS.nPersoneria=" & gPersonaJurCFLCMAC & " And CS.nMoneda=" & Val(lsMoneda) & " And CS.nProducto=" & Producto.gCapAhorros & " "
            psSql = psSql & " and Convert(Varchar(8),CS.dEstad,112)='" & Format(ldFecha, "YYYYMMdd") & "' "
            psSql = psSql & " GROUP BY Convert(Varchar(8),CS.dEstad,112)"
        Else
            psSql = "SELECT dEstadAC, SUM(nMonChqVal+nChqCMAC) as Monto " _
                & "FROM ESTADDIAACCONSOL WHERE DATEDIFF(DAY,dEstadAC,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' " _
                & "GROUP BY  dEstadAC"
        End If
End Select
Set prs = dbConec.CargaRecordSet(psSql)
Do While Not prs.EOF
   TotalCheque = TotalCheque + IIf(IsNull(prs!Monto), 0, prs!Monto)
   prs.MoveNext
Loop
prs.Close
'GetEstadSaldoCheques = TotalCheque
End Function

Public Function GetEstadSaldoMiViv(pnBitCentral As Boolean, ldFecha As Date, lsMoneda As String) As Currency
Dim TotalMiViv As Currency
Dim prs1 As ADODB.Recordset
TotalMiViv = 0
If pnBitCentral = True Then
    psSql = " select sum(nSaldoCap) as nSaldo " & _
            " from colocestaddiacred where substring(cLineaCred, 7, 3)='423' and substring(cLineaCred, 1, 2) <> '01' " & _
            " and convert(varchar(8), dEstad, 112)='" & Format(ldFecha, "YYYYMMdd") & "' AND substring(cLineaCred, 5, 1)='" & Trim(lsMoneda) & "' " & _
            " Group By substring(cLineaCred, 5, 1), convert(varchar(8), dEstad, 112) "
            ' order by substring(cLineaCred, 5, 1), convert(varchar(8), dEstad, 112)
ElseIf pnBitCentral = False Then
    psSql = " Select sum(nSaldoCap) as nSaldo " & _
            " From estaddiacredconsol where substring(ccodlincred, 1, 3)='423' and substring(ccodlincred, 6,1) <> '1' " & _
            " And convert(varchar(8), dfecha, 112)='" & Format(ldFecha, "YYYYMMdd") & "' AND substring(ccodlincred, 4, 1)='" & Trim(lsMoneda) & "' " & _
            " Group by substring(ccodlincred, 4, 1), convert(varchar(8), dfecha, 112) "
            ' " Order by substring(ccodlincred, 4, 1), convert(varchar(8), dfecha, 112) "
End If
 
Set prs1 = dbConec.CargaRecordSet(psSql)
If prs1.BOF Then
    TotalMiViv = 0
Else
    TotalMiViv = prs1!nSaldo
End If
prs1.Close
GetEstadSaldoMiViv = TotalMiViv
End Function
Public Function GetCajaAnterior(ldFechaFinPasado As String, lsCodOpe As String, lnColumna As Integer, Optional pnBitCentral As Boolean = False) As Currency
Dim TotalCajaAnterior As Currency
Dim prs1 As ADODB.Recordset
Dim ldFechaIniPasado As String

ldFechaIniPasado = "01/" & Mid(ldFechaFinPasado, 4, 7)

TotalCajaAnterior = 0
    'If pnBitCentral = True Then Es igual para el distribuido y el centralizado
        'psSql = "  "
    'ElseIf pnBitCentral = False Then
    psSql = " Select ISNULL(sum(total) / count(*),0) as nSaldoCaja from (" & _
            " SELECT fec.dFecha, SUM(nCtaSaldoImporte) as Total, SUM(nCtaSaldoImporteME) AS TotalME " & _
            " FROM fechatmp('" & Format(ldFechaFinPasado, "YYYY/MM/dd") & "') fec, CTASALDO CS JOIN RepColumnaCta OC ON cs.cCtaContCod LIKE oc.cCtaContCod + '%' " & _
            " WHERE   OC.nNroCol=" & lnColumna & " AND OC.cOpeCod = '" & lsCodOpe & "' " & _
            " AND CS.dCtaSaldoFecha IN ( SELECT MAX(CS1.dCtaSaldoFecha) FROM CTASALDO CS1 " & _
            " Where CS1.cCtaContCod = CS.cCtaContCod AND CS1.dCtaSaldoFecha <=fec.dFecha) " & _
            " and Fec.dFecha between '" & Format(ldFechaIniPasado, "YYYY/MM/dd") & "' and '" & Format(ldFechaFinPasado, "YYYY/MM/dd") & "' " & _
            " group by Fec.dFecha) a"
    'End If
 
Set prs1 = dbConec.CargaRecordSet(psSql)
If prs1.BOF Then
    psSql = "Select nsaldo nSaldoCaja from EncajeCajaSaldoAnt where dfecha = '" & Format(ldFechaFinPasado, "YYYY/MM/dd") & "'"
    Set prs1 = dbConec.CargaRecordSet(psSql)
    
    If prs1.BOF Then
        TotalCajaAnterior = 0
    Else
        TotalCajaAnterior = prs1!nSaldoCaja
    End If
Else
    If prs1!nSaldoCaja = 0 Then
        psSql = "Select nsaldo nSaldoCaja from EncajeCajaSaldoAnt where dfecha = '" & Format(ldFechaFinPasado, "YYYY/MM/dd") & "'"
        Set prs1 = dbConec.CargaRecordSet(psSql)
        
        If prs1.BOF Then
            TotalCajaAnterior = 0
        Else
            TotalCajaAnterior = prs1!nSaldoCaja
        End If
    Else
        TotalCajaAnterior = prs1!nSaldoCaja
    End If
End If
prs1.Close
GetCajaAnterior = TotalCajaAnterior
End Function
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    Dim oIni As New ClasIni
    Set oIni = Nothing
   
    Set dbConec = New DConecta
    dbConec.AbreConexion

End Sub


Private Sub Class_Terminate()
   dbConec.CierraConexion
   Set dbConec = Nothing
End Sub

Public Function GetEstadisticaPrendario(pnBitCentral As Boolean, pdFecha As Date, psMoneda As String, Optional psAgeCod As String = "") As ADODB.Recordset
    
    If pnBitCentral = True Then
        'Cambiado Pepe con nueva tabla y eliminando moneda
        psSql = "SELECT nOroVig,nNumCredVig,nCapVig,nOroAdj,nNumCredAdj,nCapAdj" _
               & " FROM  ColocEstadDiaPrenda " _
               & " WHERE substring(cLineaCred,5,1) = '" & psMoneda & "' and convert(varchar(10),dEstad,101) =  (SELECT convert(varchar(10),MAX(dEstad),101) FROM ColocEstadDiaPrenda E WHERE E.cCodAge='" & psAgeCod & "' and datediff(d,E.dEstad,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 )  " _
               & "       and cCodAge='" & psAgeCod & "'"
    Else
        'No se ha modificado del original
         psSql = "SELECT nOroVig,nNumCredVig,nCapVig,nOroAdj,nNumCredAdj,nCapAdj" _
                & " FROM  EstadDiaPrendaConsol " _
                & " WHERE cMoneda = '" & psMoneda & "' and convert(varchar(10),dFecha,101) =  (SELECT convert(varchar(10),MAX(dFecha),101) FROM EstadDiaPrendaConSol E WHERE E.cAgeBoveda='" & psAgeCod & "' and datediff(d,E.dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 )  " _
                & "       and cAgeBoveda='" & psAgeCod & "'"
    End If
 Set GetEstadisticaPrendario = dbConec.CargaRecordSet(psSql)
End Function

Public Function GetEstadisticaCreditos(pnBitCentral As Boolean, pdFecha As Date, psMoneda As String, Optional psAgeCod As String = "", Optional pnProd As Producto = -1, Optional psFiltro As String = "") As ADODB.Recordset
Dim sFiltro As String
sFiltro = ""
If Not pnProd = -1 Then
    If pnBitCentral = True Then
        sFiltro = " and SUBSTRING(cLineaCred,7,3) = '" & pnProd & "' "
    Else
        sFiltro = " and LEFT(cCodLinCred,3) = '" & pnProd & "' "
    End If
End If
If Not psAgeCod = "" Then
   sFiltro = sFiltro & " and cCodAge = '" & psAgeCod & "' "
End If
If Not psFiltro = "" Then
   sFiltro = sFiltro & psFiltro
End If

    If pnBitCentral = True Then
        'Modificado por Pepe
        psSql = "SELECT IsNull(SUM(nSaldoCap),0) as Saldo, IsNull(SUM(nNumSaldos),0) AS TotalCreditos " _
            & "FROM   ColocEstadDiaCred E " _
            & "WHERE  DATEDIFF(d,dEstad, '" & Format(pdFecha, gsFormatoFecha) & "') = 0 " _
            & "   and SubString(cLineaCred,5,1)='" & psMoneda & "' " & sFiltro
        
    Else
        psSql = "SELECT IsNull(SUM(nSaldoCap),0) as Saldo, IsNull(SUM(nNumSaldos),0) AS TotalCreditos " _
            & "FROM   EstadDiaCredConsol  E " _
            & "WHERE  DATEDIFF(d,dFecha, '" & Format(pdFecha, gsFormatoFecha) & "') = 0 " _
            & "   and SubString(cCodLinCred,4,1)='" & psMoneda & "' " & sFiltro
    End If

Set GetEstadisticaCreditos = dbConec.CargaRecordSet(psSql)
End Function

Public Function InsertaEstadAnexos(pdFecha As Date, psCodigo As String, psMoneda As String, psRuta As String)
psSql = "INSERT EstadDiaAnexos (dFecha, cCodigo, nMoneda, cRuta) VALUES ('" & Format(pdFecha, gsFormatoFecha) & "', '" & psCodigo & "', " & psMoneda & ", '" & psRuta & "')"
dbConec.Ejecutar psSql
End Function

Public Function ActualizaEstadAnexos(pdFecha As Date, psCodigo As String, psMoneda As String, psRuta As String)
psSql = "UPDATE EstadDiaAnexos SET cRuta = " & psRuta & " WHERE DatedIff(d,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cCodigo = '" & psCodigo & "' and nMoneda = " & psMoneda
dbConec.Ejecutar psSql
End Function

Public Function EliminaEstadAnexos(pdFecha As Date, psCodigo As String, psMoneda As String)
psSql = "DELETE EstadDiaAnexos WHERE DatedIff(d,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cCodigo = '" & psCodigo & "' and nMoneda = " & psMoneda
dbConec.Ejecutar psSql
End Function

Public Function GetImporteEstadAnexos(pdFecha As Date, psCodigo As String, psMoneda As String) As String
GetImporteEstadAnexos = 0
psSql = "SELECT cRuta FROM EstadDiaAnexos WHERE DatedIff(d,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cCodigo = '" & psCodigo & "' and nMoneda = " & psMoneda
Set prs = dbConec.CargaRecordSet(psSql)
If Not prs.EOF Then
   GetImporteEstadAnexos = prs!cRuta
End If
RSClose prs
End Function

Public Function GetImporteEstadAnexosMax(pdFecha As Date, psCodigo As String, psMoneda As String) As String
GetImporteEstadAnexosMax = 0

psSql = "SELECT CRUTA FROM EstadDiaAnexos WHERE CONVERT(VARCHAR(8), DFECHA, 112)=" & _
        " (SELECT MAX(CONVERT(VARCHAR(8),DFECHA,112)) FROM EstadDiaAnexos WHERE " & _
                " CONVERT(VARCHAR(8),DFECHA,112)<='" & Format(pdFecha, "YYYYMMDD") & "' " & _
                " AND cCodigo = '" & psCodigo & "' and nMoneda = " & psMoneda & _
        " ) AND cCodigo = '" & psCodigo & "' and nMoneda = " & psMoneda
 
Set prs = dbConec.CargaRecordSet(psSql)
If Not prs.EOF Then
   GetImporteEstadAnexosMax = prs!cRuta
End If
RSClose prs
End Function

Public Function GetFondosMutuos(psMoneda As String) As Currency
GetFondosMutuos = 0
Dim psCodigo As String

psCodigo = "13" & psMoneda & "30705%"
        
psSql = "select cctacontcod,nctasaldoimporte,nctasaldoimporteME " & _
    " From ctasaldo where dctasaldofecha = (select Max(dctasaldofecha)" & _
    "From ctasaldo where cctacontcod like '" & psCodigo & "') and cctacontcod like '" & psCodigo & "'"
 
Set prs = dbConec.CargaRecordSet(psSql)
If Not prs.EOF Then
    If psMoneda = "1" Then
        GetFondosMutuos = prs!nCtaSaldoImporte
    Else
        GetFondosMutuos = prs!nCtasaldoImporteME
    End If
End If
RSClose prs
End Function



