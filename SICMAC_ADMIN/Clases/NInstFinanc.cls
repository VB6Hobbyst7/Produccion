VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NInstFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ActualizaInstitucion(ByVal psPersCod As String, ByVal psTipoInst As String, ByVal psSubCta As String, ByVal psBusqRef As String) As String
Dim oInstFin As DInstFinanc
    On Error GoTo ErrorNuevaInstitucion
    ActualizaInstitucion = ""
    Set oInstFin = New DInstFinanc
    
    'Valida Existencia de Registro
    If Right("00" & Trim(psTipoInst), 2) <> Right("00" & Trim(psBusqRef), 2) Then
        If oInstFin.ExisteRegistro(psPersCod, psTipoInst, gInstFinExistenciaRegistro) Then
            ActualizaInstitucion = "Registro ya Existe"
            Set oInstFin = Nothing
            Exit Function
        End If
    End If
    
    'Valida Existencia de Referencia en Tabla CtaIF
    If Right("00" & psTipoInst, 2) <> Right("00" & Trim(psBusqRef), 2) Then
        If oInstFin.ExisteRegistro(psPersCod, psBusqRef, gInstFinExistenciaReferencia) Then
            ActualizaInstitucion = "No se Puede Realizar el proceso porque la Institucion posee Cuentas de Bancos"
            Set oInstFin = Nothing
            Exit Function
        End If
    End If
    'Ingreso de Datos
    Call oInstFin.ActualizaInstitucion(psPersCod, psTipoInst, psSubCta, psBusqRef)
    Set oInstFin = Nothing
    Exit Function

ErrorNuevaInstitucion:
        Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function NuevaInstitucion(ByVal psPersCod As String, ByVal psTipoInst As String, ByVal psSubCta As String) As String
Dim oInstFin As DInstFinanc
    On Error GoTo ErrorNuevaInstitucion
    NuevaInstitucion = ""
    Set oInstFin = New DInstFinanc
    
    'Valida Existencoia de Registro
    If oInstFin.ExisteRegistro(psPersCod, psTipoInst, gInstFinExistenciaRegistro) Then
        NuevaInstitucion = "Registro ya Existe"
        Set oInstFin = Nothing
        Exit Function
    End If
    
    'Ingreso de Datos
    Call oInstFin.InsertaNuevaInstitucion(psPersCod, psTipoInst, psSubCta)
    Set oInstFin = Nothing
    Exit Function

ErrorNuevaInstitucion:
        Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function EliminaInstitucion(ByVal psPersCod As String, ByVal psTpoInst As String) As String
Dim oInstFin As DInstFinanc
    
    On Error GoTo ErrorNuevaInstitucion
    EliminaInstitucion = ""
    Set oInstFin = New DInstFinanc
            
    'Valida Existencia de Referencia en Tabla CtaIF
    If oInstFin.ExisteRegistro(psPersCod, psTpoInst, gInstFinExistenciaReferencia) Then
        EliminaInstitucion = "No se Puede Realizar el proceso porque la Institucion posee Cuentas de Bancos"
        Set oInstFin = Nothing
        Exit Function
    End If
    
    'Ingreso de Datos
    Call oInstFin.EliminarInstitucion(psPersCod, psTpoInst)
    Set oInstFin = Nothing
    Exit Function

ErrorNuevaInstitucion:
        Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
