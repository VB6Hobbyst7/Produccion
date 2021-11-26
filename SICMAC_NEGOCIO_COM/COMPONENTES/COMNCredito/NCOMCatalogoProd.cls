VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCatalogoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GrabarCheckList(ByVal pnCodCatg As Long, ByVal pnCodProd As Long, ByVal pMatDocumento As Variant, ByVal pMatDetalle As Variant, ByVal pMatCondicion As Variant, ByVal cNrMov As String, ByVal nEstado As Integer)
Dim objChekList As COMDCredito.DCOMCatalogoProd
Dim i As Integer
Dim nCantConf As Integer
Dim cItem As String
On Error GoTo ErrorGrabarCheckList
    Set objChekList = New COMDCredito.DCOMCatalogoProd
    
    If IsArray(pMatDocumento) Then
        If UBound(pMatDocumento) > 0 Then
        Call objChekList.EliminaDatosCheckList(1, pnCodCatg, pnCodProd, pMatDocumento(1, 1), pMatDocumento(1, 3))
            
            If pMatDocumento(1, 1) = "" Then
                nCantConf = -1
                cItem = "-1"
            Else
                nCantConf = pMatDocumento(1, 3)
                cItem = pMatDocumento(1, 1)
            End If
            
            For i = 1 To UBound(pMatDocumento)
                Call objChekList.GrabarDatosCheckList(1, pnCodCatg, pnCodProd, IIf(pMatDocumento(i, 1) = "", "-1", CStr(pMatDocumento(i, 1))), IIf(pMatDocumento(i, 2) = "", "-1", CStr(pMatDocumento(i, 2))), nCantConf, IIf(pMatDocumento(i, 4) = "", -1, pMatDocumento(i, 4)), IIf(pMatDocumento(i, 5) = "", -1, pMatDocumento(i, 5)), , , , , , , , , cNrMov, nEstado)
            Next i
        End If
    End If
    
    If IsArray(pMatDetalle) Then
         If UBound(pMatDetalle) > 0 Then
        Call objChekList.EliminaDatosCheckList(2, pnCodCatg, pnCodProd, pMatDetalle(1, 1), nCantConf)
            For i = 1 To UBound(pMatDetalle)
                Call objChekList.GrabarDatosCheckList(2, pnCodCatg, pnCodProd, IIf(CStr(pMatDetalle(i, 1)) = "", "-1", CStr(pMatDetalle(i, 1))), pMatDetalle(i, 2), nCantConf, , , pMatDetalle(i, 3), , , , , , , , cNrMov, nEstado)
            Next i
        End If
    End If
        
    If IsArray(pMatCondicion) Then
        If UBound(pMatCondicion) > 0 Then
        Call objChekList.EliminaDatosCheckList(3, pnCodCatg, pnCodProd, cItem, nCantConf)
            For i = 1 To UBound(pMatCondicion)
                Call objChekList.GrabarDatosCheckList(3, pnCodCatg, pnCodProd, cItem, , nCantConf, , , , IIf(pMatCondicion(i, 1) = "", -1, pMatCondicion(i, 1)), IIf(pMatCondicion(i, 2) = "", -1, pMatCondicion(i, 2)), IIf(pMatCondicion(i, 3) = "", -1, pMatCondicion(i, 3)), IIf(pMatCondicion(i, 4) = "", "-1", pMatCondicion(i, 4)), IIf(pMatCondicion(i, 5) = "", -1, pMatCondicion(i, 5)), IIf(pMatCondicion(i, 6) = "", "-1", pMatCondicion(i, 6)), IIf(pMatCondicion(i, 7) = "", -1, pMatCondicion(i, 7)), cNrMov, nEstado)
            Next i
        End If
    End If
    
    Set objChekList = Nothing
    GrabarCheckList = True
    Exit Function
ErrorGrabarCheckList:
    GrabarCheckList = False
    Err.Raise Err.Number, "Error Registro de CheckList", Err.Description
End Function
