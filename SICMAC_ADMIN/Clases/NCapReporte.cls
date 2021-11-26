VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCapReporte"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7C4C23012E"
'Genera todos los reportes de Captaciones
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3A835D6302BB
    Private mlClassDebugID As Long
#End If

'##ModelId=3A7CC7DE030A
'Public theNCapMovimientos As NCapMovimientos

'##ModelId=3A7CC9FD0076
'Public theNCapMovimientosUser As NCapMovimientosUser

'##ModelId=3A80BDB60283
Public theNImprimir As NImprimir

'##ModelId=3A7CC7FA01A2
Public theNCaptaImpresion As NCaptaImpresion

'##ModelId=3A7CC6970089
'Public theDReportes As DReporte

'Devuelve los reportes de captaciones, puede ser en un rango de fechas, para un
'usuario especifico.
'
'Se le indica que un indice para indicar el reporte :
'
'Diarios
'001     Listado de Cheques Recibidos Hoy
'002     Listado de Cuentas Movidas
'003     Protocolo del Dia Soles
'004     Protocolo del Dia Dolares
'005     Estadistica Diaria Ahorro
'006     Estadistica Diaria PlazoFijo
'007     Estadistica Diaria CTS
'008     Movimiento Diario Total Consolidado
'009     Informe de Saldos por Tipo de Cuenta
'010     Estratificación de las Cuentas
'011     Listado de Cobranza de Servicios
'012     Compra Venta de Moneda Extranjera
'013     Informe para Caja General
'014     Ahorro para Efectos de Encaje
'015     Saldos Diarios de Moneda Extranjera
'016     Ordenes de Pago en Canje
'017     Listado de Sobrantes y/o Faltantes
'018     Listado de Notas de Abono y/o Cargo
'019     Listado de Otras Operaciones
'020     Operaciones con CMACs
'
'Mensuales
'101     Saldo de Cuentas de Ahorro - Personeria Natural
'102     Saldo de Cuentas de Ahorro - Personeria Juridica SF             Lucro
'103     Saldo de Cuentas de Ahorro - Personeria Juridica CF            Lucro
'104     Saldo de Cuentas de Ahorro - Personeria CMAC
'105     Saldo de Cuentas de Ahorro - Instituciones Financieras
'106     Saldo de Cuentas de Plazo Fijo - Personería Natural
'107     Saldo de Cuentas de Plazo Fijo - Personería Jurídica            SF
'Lucro
'108     Saldo de Cuentas de Plazo Fijo - Personería Jurídica
'     CF Lucro
'109     Saldo de Cuentas de Plazo Fijo - Personería Otras
'  CMAC's
'110     Saldo de Cuentas de Plazo Fijo - Instituciones
'Financieras
'111     Saldo de Cuentas de CTS
'112     Listado General de Cuentas
'113     Depósitos Según Escala de Montos - Soles
'114     Depósitos Según Escala de Montos - Dolares
'115     Depósitos Según Escala de Montos - Consolidado
'116     Rango de Saldo de Depósitos - Soles
'117     Rango de Saldo de Depósitos - Dolares
'118     Movimiento de los depósitos según Monto y Número
' de Cuentas
'119     Plazos Fijos por Vencimiento
'120     Listado de Ahorros para el F.S.D.
'
'Varios
'201     Listado de Cheques Recibidos
'202     Listado de Cheques Valorizados
'203     Listado de Cheques en Valorización
'204     Listado de Cheques Anulados
'205     Listado de Cheques Observados
'206     Reporte de Extornos
'207     Cuentas Inactivas
'208     Consolidado de Inactivas
'209     Cartas de Inactivas
'210     Cuentas Aperturadas
'211     Cuentas Canceladas
'212     Cuentas Inactivas a la Fecha
'213     Riesgo de Tasas de Interés
'214     Ahorros Sin Orden de Pago
'215     Ahorros Con Orden de Pago
'216     Plazo Fijo
'217     CTS
'218     Listado Total de Cuentas de Ahorros
'219     Listado Total de Cuentas de Plazo Fijo
'220     Listado Total de Cuentas de CTS
'221     Situación de Captaciones
'222     Reporte de Operaciones con Entidades Convenio
'
'Reportes de Cajero:
'301    Billetaje
'302    Detalle de Operaciones
'304    Habilitaciones y Devoluciones.
'305    Protocolo por Usuario
'
'
'##ModelId=3A7C504F0365
Public Function GetReporte(psItem As String, psUser As String, pdInicio As Date, pdFecFin As Date) As String
    On Error GoTo GetReporteErr

    'your code goes here...

    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NCapReporte:GetReporte Method")
End Function

'Obtiene el calculo de ingresos y egresos, por usuario, y en una fecha
'determinada
'##ModelId=3A7C54510099
Public Function GetResumenIngresosEgresos(psUser As String, pdFecha As Date) As String
    On Error GoTo GetResumenIngresosEgresosErr

    'your code goes here...

    Exit Function
GetResumenIngresosEgresosErr:
    Call RaiseError(MyUnhandledError, "NCapReporte:GetResumenIngresosEgresos Method")
End Function

'Declara el Sobrante faltante de un usuario que mueve caja
'##ModelId=3A82FECC029D
Public Function SobranteFaltante(psUser As String) As Boolean
    On Error GoTo SobranteFaltanteErr

    'your code goes here...

    Exit Function
SobranteFaltanteErr:
    Call RaiseError(MyUnhandledError, "NCapReporte:SobranteFaltante Method")
End Function

'##ModelId=3A835D640078
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

'##ModelId=3A835D6400D2
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3A835D63033D
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
