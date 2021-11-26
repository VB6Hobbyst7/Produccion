Attribute VB_Name = "gConstCaptaciones"
Option Explicit
'Operacion de Inicio de Dia
Global Const gCapInicioDia = "905004"


'Constantes para el cálculo de la planilla
Global Const gsCajHabDeBove = "06501001"
Global Const gsCajDevABove = "06502001"
Global Const gsCajDevBilletaje = "06503001"

Global Const gsCajHabCajGen = "06001001"
Global Const gsCajDevCajGen = "06002001"

Global Const gsCajSobCaja = "07001001"
Global Const gsCajFaltCaja = "07002001"

Global Const cPlantillaCartaRenPF = "\FormatoCarta\CartaAvisoRenovacionPF.txt"

'RIRO20161103 **********************
Public Type Seleccion
    dFechaVencimiento As Date
    nOrdenFechaVenc As Integer
End Type
'END RIRO **************************
