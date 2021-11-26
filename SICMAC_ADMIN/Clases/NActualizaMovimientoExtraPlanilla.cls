VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaMovimientoExtraPla"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9701ED03AB"
'Actualiza datos extra planillas
Option Base 0
Option Explicit

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D80320
    Private mlClassDebugID As Long
#End If

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'##ModelId=3AB902D803C0
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D803AC
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega un nuevo movimiento extra planilla
'##ModelId=3A9AF8830289
Public Function AgregaExtPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, prRS As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaExtPlanillaErr
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla
    
    oExtra.AgregaExtPlanilla psRRHHPeriodo, psPlanillaTpoCod, prRS, psUltimaActualizacion
    
    Set oExtra = Nothing
    
    Exit Function
AgregaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaMovimientoExtraPlanilla:AgregaExtPlanilla Method")
End Function

'Elimina un movimiento extra planilla
'##ModelId=3A9AF8830316
Public Function EliminaExtPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, psRHPersCod As String, psCodCta As String) As Boolean
    On Error GoTo EliminaExtPlanillaErr

    'your code goes here...

    Exit Function
EliminaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaMovimientoExtraPlanilla:EliminaExtPlanilla Method")
End Function

'Actualiza un movimiento extra planilla
'##ModelId=3A9AF883035C
Public Function ModificaExtPlanilla(psRRHHPeriodo As String, psPlanillaTpoCod As String, prRS As ADODB.Recordset, psUltimaActualizacion As String, psMonto As String, psComentario As String, psTipoOpe As String) As Boolean
    On Error GoTo ModificaExtPlanillaErr
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla

    oExtra.ModificaExtPlanilla psRRHHPeriodo, psPlanillaTpoCod, prRS, psUltimaActualizacion, psMonto, psComentario, psTipoOpe
    
    Exit Function
ModificaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaMovimientoExtraPlanilla:ModificaExtPlanilla Method")
End Function

'Actualiza un movimiento extra planilla
'##ModelId=3A9AF883035C
Public Function GetReporte(psRRHHPeriodo As String, psPlanillaTpoCod As String, psNombrePlanilla As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo ModificaExtPlanillaErr
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsCodigo As String * 13
    Dim lsNombre As String * 30
    Dim lsMonto As String * 8
    Dim lsComentario As String * 30
    Dim lsDocumentos As String * 10
    Dim lsCuenta As String * 18
    Dim lsCuentaRef As String * 18
    
    Dim lnAcumCob As Currency
    Dim lnAcumNoCob As Currency
    
    Set rsE = oExtra.GetExtPlanilla(psRRHHPeriodo, psPlanillaTpoCod, True)
    lsCadena = ""
    
    If Not (rsE.EOF And rsE.BOF) Then
        lsCadena = lsCadena & CabeceraPagina(psNombrePlanilla & " - " & psRRHHPeriodo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;12; ;2;Nombre;20; ;12;Cuenta;10; ;8;Monto;8; ;2;Comentario;20; ;15;Cta_Ref;10;Docu;18; ;2;", lnItem)
        
        lnAcumCob = 0
        lnAcumNoCob = 0
        
        While Not rsE.EOF
            
            lsCodigo = rsE.Fields(0)
            lsNombre = rsE.Fields(1)
            lsCuenta = rsE.Fields(4)
            lsCuentaRef = rsE.Fields(10)
            RSet lsMonto = Format(rsE.Fields(2), "#,##0.00  ")
            lsComentario = rsE.Fields(3)
            lsDocumentos = rsE.Fields(11) & ""
            lsCadena = lsCadena & lsCodigo & Space(2) & lsNombre & Space(2) & lsCuenta & Space(2) & lsMonto & Space(2) & lsComentario & Space(2) & lsCuentaRef & Space(2) & lsDocumentos & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            
            lnAcumCob = lnAcumCob + rsE.Fields(2)
            If Trim(lsDocumentos) = "" Then lnAcumNoCob = lnAcumNoCob + rsE.Fields(2)
            
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina(psNombrePlanilla & " - " & psRRHHPeriodo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;12; ;2;Nombre;20; ;12;Cuenta;10; ;8;Monto;8; ;2;Comentario;20; ;15;Cta_Ref;10;Docu;18; ;2;", lnItem)
            End If
            
            rsE.MoveNext
        Wend
    End If
    
   lsCadena = lsCadena & String(140, "=") & oImpresora.gPrnSaltoLinea
   lsCadena = lsCadena & Space(10) & "Resumen    :        Monto Total  = " & Format(lnAcumCob, "#,##0.00") & "     Monto Cobrado " & Format(lnAcumCob - lnAcumNoCob, "#,##0.00") & "         Monto No Cobrado " & Format(lnAcumNoCob, "#,##0.00") & oImpresora.gPrnSaltoLinea
    
    GetReporte = lsCadena
    
    Exit Function
ModificaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaMovimientoExtraPlanilla:ModificaExtPlanilla Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902D80334
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function GetValida(psRRHHPeriodo As String, psPlanillaTpoCod As String, psNombrePlanilla As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo ModificaExtPlanillaErr
    Dim oExtra As DActualizaMovExtraPlanilla
    Set oExtra = New DActualizaMovExtraPlanilla
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsCodigo As String * 13
    Dim lsNombre As String * 30
    Dim lsMonto As String * 8
    
    Dim lnAcumCob As Currency
    
    Set rsE = oExtra.GetRHExtraPlanillaValida(psPlanillaTpoCod, psRRHHPeriodo, True)
    lsCadena = ""
    
    If Not (rsE.EOF And rsE.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("VALIDA EXTRA PLANILLA X CUENTAS - " & psNombrePlanilla & " - " & psRRHHPeriodo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;12; ;2;Referencia;20; ;12;Monto;8; ;5;", lnItem)
        
        lnAcumCob = 0
        
        While Not rsE.EOF
            lsCodigo = rsE.Fields(1)
            lsNombre = rsE.Fields(2)
            RSet lsMonto = Format(rsE.Fields(0), "#,##0.00")
            lsCadena = lsCadena & lsCodigo & Space(2) & lsNombre & Space(2) & lsMonto & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            
            lnAcumCob = lnAcumCob + rsE.Fields(0)
            
            rsE.MoveNext
        Wend
    End If
    
   lsCadena = lsCadena & String(60, "=") & oImpresora.gPrnSaltoLinea
   lsCadena = lsCadena & Space(10) & "Resumen    :          Monto Total  = " & Format(lnAcumCob, "#,##0.00") & oImpresora.gPrnSaltoLinea
    
    Set rsE = oExtra.GetRHExtraPlanillaValida(psPlanillaTpoCod, psRRHHPeriodo, False)
    If Not (rsE.EOF And rsE.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("VALIDA EXTRA PLANILLA X DESTINO - " & psNombrePlanilla & " - " & psRRHHPeriodo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;12; ;2;Referencia;20; ;12;Monto;8; ;5;", lnItem)
        
        lnAcumCob = 0
        
        While Not rsE.EOF
            lsCodigo = rsE.Fields(1)
            lsNombre = rsE.Fields(2)
            RSet lsMonto = Format(rsE.Fields(0), "#,##0.00")
            lsCadena = lsCadena & lsCodigo & Space(2) & lsNombre & Space(2) & lsMonto & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            
            lnAcumCob = lnAcumCob + rsE.Fields(0)
            
            rsE.MoveNext
        Wend
    End If
    
   lsCadena = lsCadena & String(60, "=") & oImpresora.gPrnSaltoLinea
   lsCadena = lsCadena & Space(10) & "Resumen    :          Monto Total  = " & Format(lnAcumCob, "#,##0.00") & oImpresora.gPrnSaltoLinea
    
    GetValida = lsCadena
    
    Exit Function
ModificaExtPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaMovimientoExtraPlanilla:ModificaExtPlanilla Method")
End Function

