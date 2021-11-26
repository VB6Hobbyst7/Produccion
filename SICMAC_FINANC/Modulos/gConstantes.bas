Attribute VB_Name = "gConstantes"
Option Explicit
'Constantes de Contabilidad
Global Const gsContDebe = "D"
Global Const gsContHaber = "H"
Global Const gsContDebeDesc = "Debe"
Global Const gsContHaberDesc = "Haber"

Global Const gsOpeCtaCaracterObligaDesc = "OBLIGATORIO"
Global Const gsOpeCtaCaracterOpcionDesc = "OPCIONAL"
Global Const gsOpeAnalCtaHisto = "82_02"

Global Const gsSI = "SI"
Global Const gsNO = "NO"

Global Const gTpoCtaIFCtaPFOverNight = 4 'EJVG20120801
'SOLO HASTA CREAR EN TABLA
'Global Const gMovEstContabSustPendRendir = 14

'ARLO 20170208***************************************
Public glsMovNro As String
'****************************************************

'ARLO ARLO20170208**************************************
Public Enum LogOperacionesPistas
    gIngresarSalirSistema = 700100
    LogPistaTipoCambio = 900021
    LogPistaParaMantViatico = 790300
    LogPistaMantDocumento = 790500
    LogPistaMantImpuesto = 790700
    LogPistaMantClasifOperacion = 790900
    LogPistaMantObjetos = 791900
    LogPistaMantCctaCont = 791600
    LogMantSaldoProducto = 559400
    LogPistaRegistoAsientCont = 742121
    LogPistaAjusteTpoCambio = 701201
    LogPistaReclasifiCartera = 701221
    LogPistaManCtaPend = 792700
    LogPistaOpePend = 760400
    LogPistaActuaReferencia = 792800
    LogPistaCierreDiarioCont = 795400
    LogPistaAsignaIGVCreditoNoFiscal = 760499
    LogPistaExtornoCierreMensual = 701900
    LogPistaActSaldos = 793500
    
    gPersonaRegistro = 900110 'LUCV20181220, Anexo01 de Acta 199-2018
    gPersonaMantenimiento = 900111 'LUCV20181220, Anexo01 de Acta 199-2018
    gPersonaConsulta = 900112 'LUCV20181220, Anexo01 de Acta 199-2018

End Enum
