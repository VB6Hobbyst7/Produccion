Attribute VB_Name = "gSugerenciaProducto"
'****APRI 20163010
Public Sub BusquedaProducto(ByVal cPersCod As String, ByVal cCtaCod As String, ByVal nOpe As COMDConstantes.CaptacOperacion)

   Dim ClsPersonas As COMDPersona.DCOMPersonas
   Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Set ClsPersonas = New COMDPersona.DCOMPersonas
    Set R = ClsPersonas.SugerenciaProductoAhorros(cPersCod, cCtaCod)
    
    frmSugerenciaProductos.Inicio R

End Sub

