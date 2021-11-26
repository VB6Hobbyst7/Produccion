VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Public Function NroEnLetras(ByVal curNumero As Double, Optional blnO_Final As Boolean = True) As String
'Devuelve un número expresado en letras.
'El parámetro blnO_Final se utiliza en la recursión para saber si se debe colocar
'la "O" final cuando la palabra es UN(O)
    Dim dblCentavos As Double
    Dim lngContDec As Long
    Dim lngContCent As Long
    Dim lngContMil As Long
    Dim lngContMillon As Long
    Dim strNumLetras As String
    Dim strNumero As Variant
    Dim strDecenas As Variant
    Dim strCentenas As Variant
    Dim blnNegativo As Boolean
    Dim blnPlural As Boolean
    
    If Int(curNumero) = 0# Then
        strNumLetras = "CERO"
    End If
    
    strNumero = Array(vbNullString, "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", _
                   "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", _
                   "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", _
                   "VEINTE")

    strDecenas = Array(vbNullString, vbNullString, "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", _
                    "SETENTA", "OCHENTA", "NOVENTA", "CIEN")

    strCentenas = Array(vbNullString, "CIENTO", "DOSCIENTOS", "TRESCIENTOS", _
                     "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", _
                     "OCHOCIENTOS", "NOVECIENTOS")

    If curNumero < 0# Then
        blnNegativo = True
        curNumero = Abs(curNumero)
    End If

    If Int(curNumero) <> curNumero Then
        dblCentavos = Abs(curNumero - Int(curNumero))
        curNumero = Int(curNumero)
    End If

    Do While curNumero >= 1000000#
        lngContMillon = lngContMillon + 1
        curNumero = curNumero - 1000000#
    Loop

    Do While curNumero >= 1000#
        lngContMil = lngContMil + 1
        curNumero = curNumero - 1000#
    Loop
    
    Do While curNumero >= 100#
        lngContCent = lngContCent + 1
        curNumero = curNumero - 100#
    Loop
    
    If Not (curNumero > 10# And curNumero <= 20#) Then
        Do While curNumero >= 10#
            lngContDec = lngContDec + 1
            curNumero = curNumero - 10#
        Loop
    End If
    
    If lngContMillon > 0 Then
        If lngContMillon >= 1 Then   'si el número es >1000000 usa recursividad
            strNumLetras = NroEnLetras(lngContMillon, False)
            If Not blnPlural Then blnPlural = (lngContMillon > 1)
            lngContMillon = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMillon) & " MILLON" & _
                                                                    IIf(blnPlural, "ES ", " ")
    End If
    
    If lngContMil > 0 Then
        If lngContMil >= 1 Then   'si el número es >100000 usa recursividad
            strNumLetras = strNumLetras & NroEnLetras(lngContMil, False)
            lngContMil = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMil) & " MIL "
    End If
    
    If lngContCent > 0 Then
        If lngContCent = 1 And lngContDec = 0 And curNumero = 0# Then
            strNumLetras = strNumLetras & "CIEN"
        Else
            strNumLetras = strNumLetras & strCentenas(lngContCent) & " "
        End If
    End If
    
    If lngContDec >= 1 Then
        If lngContDec = 1 Then
            strNumLetras = strNumLetras & strNumero(10)
        Else
            strNumLetras = strNumLetras & strDecenas(lngContDec)
        End If
        
        If lngContDec >= 3 And curNumero > 0# Then
            strNumLetras = strNumLetras & " Y "
        End If
    Else
        If curNumero >= 0# And curNumero <= 20# Then
            strNumLetras = strNumLetras & strNumero(curNumero)
            If curNumero = 1# And blnO_Final Then
                strNumLetras = strNumLetras & "O"
            End If
            If dblCentavos > 0# Then
                strNumLetras = Trim(strNumLetras) & " CON " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
            End If
            NroEnLetras = strNumLetras
            Exit Function
        End If
    End If
    
    If curNumero > 0# Then
        strNumLetras = strNumLetras & strNumero(curNumero)
        If curNumero = 1# And blnO_Final Then
            strNumLetras = strNumLetras & "O"
        End If
    End If
    
    If dblCentavos > 0# Then
        strNumLetras = strNumLetras & " CON " + Format$(CInt(dblCentavos * 100#), "00") & "/100"
    End If
    
    NroEnLetras = IIf(blnNegativo, "(" & strNumLetras & ")", strNumLetras)
End Function

