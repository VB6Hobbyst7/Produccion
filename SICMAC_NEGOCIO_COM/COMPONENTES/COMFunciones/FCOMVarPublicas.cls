VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCOMVarPublicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'************************ Constantes Públicas *******************
'*******************************************************************

Public Property Get gnNroDigitosDNI() As Variant
    gnNroDigitosDNI = 8
End Property

Public Property Get gnNroDigitosRUC() As Variant
    gnNroDigitosRUC = 11
End Property

Public Property Get gsFormatoFecha() As Variant
    gsFormatoFecha = "mm/dd/yyyy"
End Property

Public Property Get gsFormatoFechaHora() As Variant
    gsFormatoFechaHora = "mm/dd/yyyy hh:mm:ss"
End Property

Public Property Get gsFormatoFechaHoraView() As Variant
    gsFormatoFechaHoraView = "dd/mm/yyyy hh:mm:ss"
End Property


Public Property Get gsFormatoFechaView() As Variant
    gsFormatoFechaView = "dd/mm/yyyy"
End Property

Public Property Get gcFormatoFechaHoraView() As Variant
     gcFormatoFechaHoraView = "dd/mm/yyyy hh:mm:ss"
End Property

Public Property Get gcFormatoFechaHora() As Variant
     gcFormatoFechaHora = "mm/dd/yyyy hh:mm:ss"
End Property

Public Property Get gcFormatoFechaView() As Variant
     gcFormatoFechaView = "dd/mm/yyyy"
End Property

Public Property Get gsFormatoNumeroView() As Variant
    gsFormatoNumeroView = "##,###,##0.00##"
End Property

Public Property Get gcFormView() As Variant
    gcFormView = "##,###,##0.00##"
End Property

Public Property Get gsFormatoNumeroDato() As Variant
    gsFormatoNumeroDato = "#######0.00##"
End Property

Public Property Get gcFormDato() As Variant
    gcFormDato = "#######0.00##"
End Property

Public Property Get gcFormatoTC() As Variant
    gcFormatoTC = "#0.00##"
End Property

Public Property Get gnColPage() As Variant
    gnColPage = 79
End Property

Public Property Get gnLinVert() As Variant
    gnLinVert = 66
End Property

Public Property Get gnLinHori() As Variant
    gnLinHori = 46
End Property

Public Property Get gcFormatoMov() As Variant
    gcFormatoMov = "yyyymmdd"
End Property

Public Property Get gbComunRemoto() As Variant
    gbComunRemoto = False
End Property

Public Property Get IDPlantillaOP() As Variant
    IDPlantillaOP = "OPBatch"
End Property

Public Property Get IDPlantillaVOP() As Variant
    IDPlantillaVOP = "OPVEBatch"
End Property

Public Property Get gsUsuarioBOVEDA() As Variant
    gsUsuarioBOVEDA = "BOVE"
End Property

Public Property Get gsContDebe() As Variant
    gsContDebe = "D"
End Property


Public Property Get gsContHaber() As Variant
    gsContHaber = "H"
End Property


Public Property Get gsContDebeDesc() As Variant
   gsContDebeDesc = "Debe"
End Property

Public Property Get gsContHaberDesc() As Variant
  gsContHaberDesc = "Haber"
End Property

Public Property Get gsOpeCtaCaracterObligaDesc() As Variant
   gsOpeCtaCaracterObligaDesc = "OBLIGATORIO"
End Property

Public Property Get gsOpeCtaCaracterOpcionDesc() As Variant
 gsOpeCtaCaracterOpcionDesc = "OPCIONAL"
End Property


Public Property Get gsSI() As Variant
    gsSI = "SI"
End Property

Public Property Get gsNO() As Variant
   gsNO = "NO"
End Property

Public Property Get gsFormatoMovFecha() As Variant
   gsFormatoMovFecha = "yyyymmdd"
End Property

Public Property Get gsFormatoMovFechaHora() As Variant
   gsFormatoMovFechaHora = "yyyymmddhhmmss"
End Property

Public Property Get gsMenuAplicac() As Variant
   gsMenuAplicac = "1"
End Property

Public Property Get CB_SETDROPPEDWIDTH() As Variant
   CB_SETDROPPEDWIDTH = &H160
End Property

Public Property Get CB_FINDSTRING() As Variant
   CB_FINDSTRING = &H14C
End Property

Public Property Get gsConnServDBF() As Variant
   gsConnServDBF = "DSN=DSNCmactServ"
End Property

Public Property Get gsCajHabDeBove() As Variant
   gsCajHabDeBove = "06501001"
End Property

Public Property Get gsCajDevABove() As Variant
   gsCajDevABove = "06502001"
End Property

Public Property Get gsCajDevBilletaje() As Variant
   gsCajDevBilletaje = "06503001"
End Property

Public Property Get gsCajHabCajGen() As Variant
   gsCajHabCajGen = "06001001"
End Property

Public Property Get gsCajDevCajGen() As Variant
   gsCajDevCajGen = "06002001"
End Property

Public Property Get gsCajSobCaja() As Variant
   gsCajSobCaja = "07001001"
End Property

Public Property Get gsCajFaltCaja() As Variant
   gsCajFaltCaja = "07002001"
End Property
