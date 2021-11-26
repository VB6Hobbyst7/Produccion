VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCOMVarImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Impresoras
    gEPSON = 1
    gHEWLETT_PACKARD = 2
    gIBM = 3
End Enum

Public gImpDfecto As Impresoras
Public gImpFinal As Impresoras

'Se encontraba en el Modulo
'Public lmImpresoras(2) As String
Public lmImpresorasEPSON As String
Public lmImpresorasgHP As String


'************************************************************************
'DECLARACIONES PARA JUEGO DE CARACTERES DE IMPRESION x DEFECTO
'************************************************************************
Dim gPrnMargenIzqDef As String
Dim gPrnMargenIzq As String

Public gPrnInicializaDef As String
Public gPrnNegritaONDef As String
Public gPrnNegritaOFFDef As String
Public gPrnBoldONDef As String
Public gPrnBoldOFFDef As String
Public gPrnSaltoLineaDef As String
Public gPrnSaltoPaginaDef As String
Public gPrnEspaLineaValorDef As String
Public gPrnEspaLineaNDef As String
Public gPrnTamPaginaCabDef As String
Public gPrnTamPagina16Def As String
Public gPrnTamPagina17Def As String
Public gPrnTamPagina18Def As String
Public gPrnTamPagina22Def As String
Public gPrnTamPagina65Def As String
Public gPrnTamPagina66Def As String
Public gPrnTamPagina70Def As String
Public gPrnTamLetra10CPIDef As String
Public gPrnTamLetra12CPIDef As String
Public gPrnTamLetra15CPIDef As String
Public gPrnTpoLetraRomanDef As String
Public gPrnTpoLetraSansSerifDef As String
Public gPrnTpoLetraCurierDef As String

Public gPrnMargenIzqCabDef As String
Public gPrnMargenIzq00Def As String
Public gPrnMargenIzq01Def As String
Public gPrnMargenIzq02Def As String
Public gPrnMargenIzq06Def As String
Public gPrnMargenIzq44Def As String
Public gPrnMargenDerCabDef As String

Public gPrnTpoLetraRoman1PDef As String
Public gPrnTpoLetraSansSerif1PDef As String
Public gPrnTpoLetraCurier1PDef As String

Public gPrnTpoLetraRomanE2Def As String
Public gPrnTpoLetraSansSerifE2Def As String
Public gPrnTpoLetraCurierE2Def As String
Public gPrnTpoLetraCurierE240Def As String

Public gPrnCondensadaOFFDef As String
Public gPrnCondensadaONDef As String
Public gPrnUnderLineONOFFDef As String
Public gPrnItalicONDef As String
Public gPrnItalicOFFDef As String
Public gPrnDblAnchoONDef As String
Public gPrnDblAnchoOFFDef As String
Public gPrnUnoMedioEspacioDef As String
Public gPrnSuperIdxOnDef As String
Public gPrnSuperIdxOFFDef As String


'************************************************************************
'DECLARACIONES PARA JUEGO DE CARACTERES DE IMPRESION
'************************************************************************
Public gPrnInicializa As String
Public gPrnNegritaON As String
Public gPrnNegritaOFF As String
Public gPrnBoldON As String
Public gPrnBoldOFF As String
Public gPrnSaltoLinea As String
Public gPrnSaltoPagina As String
Public gPrnEspaLineaValor As String
Public gPrnEspaLineaN As String
Public gPrnTamPaginaCab As String
Public gPrnTamPagina16 As String
Public gPrnTamPagina17 As String
Public gPrnTamPagina18 As String
Public gPrnTamPagina22 As String
Public gPrnTamPagina65 As String
Public gPrnTamPagina66 As String
Public gPrnTamPagina70 As String
Public gPrnTamLetra10CPI As String
Public gPrnTamLetra12CPI As String
Public gPrnTamLetra15CPI As String
Public gPrnTpoLetraRoman As String
Public gPrnTpoLetraSansSerif As String
Public gPrnTpoLetraCurier As String
  
Public gPrnMargenIzqCab As String
Public gPrnMargenIzq00 As String
Public gPrnMargenIzq01 As String
Public gPrnMargenIzq02 As String
Public gPrnMargenIzq06 As String
Public gPrnMargenIzq44 As String
Public gPrnMargenDerCab As String
  
Public gPrnTpoLetraRoman1P As String
Public gPrnTpoLetraSansSerif1P As String
Public gPrnTpoLetraCurier1P As String
  
Public gPrnTpoLetraRomanE2 As String
Public gPrnTpoLetraSansSerifE2 As String
Public gPrnTpoLetraCurierE2 As String
Public gPrnTpoLetraCurierE240 As String
 
Public gPrnCondensadaOFF As String
Public gPrnCondensadaON As String
Public gPrnUnderLineONOFF As String
Public gPrnItalicON As String
Public gPrnItalicOFF As String
Public gPrnDblAnchoON As String
Public gPrnDblAnchoOFF As String
Public gPrnUnoMedioEspacio As String
Public gPrnSuperIdxOn As String
Public gPrnSuperIdxOFF As String


Public Sub Inicia(Optional pnTipoImp As Impresoras = gEPSON)
    gImpDfecto = gEPSON
    gImpFinal = gHEWLETT_PACKARD
    
    'DEFINICION DE IMPRESORAS
    lmImpresorasEPSON = "EPSON"
    lmImpresorasgHP = "HEWLETT PACKARD"
    
    gPrnInicializa = ""
    gPrnNegritaON = ""
    gPrnNegritaOFF = ""
    gPrnBoldON = ""
    gPrnBoldOFF = ""
    gPrnSaltoLinea = ""
    gPrnSaltoPagina = ""
    gPrnEspaLineaValor = ""
    gPrnEspaLineaN = ""
    gPrnTamPaginaCab = ""
    gPrnTamPagina16 = ""
    gPrnTamPagina17 = ""
    gPrnTamPagina18 = ""
    gPrnTamPagina22 = ""
    gPrnTamPagina65 = ""
    gPrnTamPagina66 = ""
    gPrnTamPagina70 = ""
    gPrnTamLetra10CPI = ""
    gPrnTamLetra12CPI = ""
    gPrnTamLetra15CPI = ""
    gPrnTpoLetraRoman = ""
    gPrnTpoLetraSansSerif = ""
    gPrnTpoLetraCurier = ""
           
    gPrnMargenIzqCab = ""
    gPrnMargenIzq00 = ""
    gPrnMargenIzq01 = ""
    gPrnMargenIzq02 = ""
    gPrnMargenIzq06 = ""
    gPrnMargenIzq44 = ""
    gPrnMargenDerCab = ""
            
    gPrnTpoLetraRoman1P = ""
    gPrnTpoLetraSansSerif1P = ""
    gPrnTpoLetraCurier1P = ""
            
    gPrnTpoLetraRomanE2 = ""
    gPrnTpoLetraSansSerifE2 = ""
    gPrnTpoLetraCurierE2 = ""
    gPrnTpoLetraCurierE240 = ""
       
    gPrnCondensadaOFF = ""
    gPrnCondensadaON = ""
    gPrnUnderLineONOFF = ""
    gPrnItalicON = ""
    gPrnItalicOFF = ""
    gPrnDblAnchoON = ""
    gPrnDblAnchoOFF = ""
    gPrnUnoMedioEspacio = ""
    gPrnSuperIdxOn = ""
    gPrnSuperIdxOFF = ""

    gPrnBoldOFFDef = Chr$(27) & Chr$(70)
    gPrnBoldONDef = Chr$(27) & Chr$(69)
    gPrnCondensadaOFFDef = Chr$(27) & Chr$(18)
    gPrnCondensadaONDef = Chr$(27) & Chr$(15)
    'gPrnCondensadaOFFDef = Chr$(18)
    'gPrnCondensadaONDef = Chr$(15)
    gPrnDblAnchoOFFDef = Chr$(27) & Chr$(20)
    gPrnDblAnchoONDef = Chr$(27) & Chr$(87)
    gPrnEspaLineaNDef = Chr$(27) & Chr$(50)
    gPrnEspaLineaValorDef = Chr$(27) & Chr$(65)
    gPrnInicializaDef = Chr$(27) & Chr$(64)
    gPrnItalicOFFDef = Chr$(27) & Chr$(53)
    gPrnItalicONDef = Chr$(27) & Chr$(52)
    gPrnMargenDerCabDef = Chr$(27) & Chr$(81)
    gPrnMargenIzq00Def = Chr$(27) & Chr$(108) & Chr$(0)
    gPrnMargenIzq01Def = Chr$(27) & Chr$(108) & Chr$(1)
    gPrnMargenIzqDef = Chr$(27) & Chr$(108) & Chr$(2)
    gPrnMargenIzq06Def = Chr$(27) & Chr$(108) & Chr$(6)
    gPrnMargenIzq44Def = Chr$(27) & Chr$(108) & Chr$(44)
    gPrnMargenIzqCabDef = Chr$(27) & Chr$(108)
    gPrnNegritaOFFDef = Chr$(27) & Chr$(72)
    gPrnNegritaONDef = Chr$(27) & Chr$(71)
    gPrnSaltoLineaDef = Chr$(10)
    gPrnSaltoPaginaDef = Chr$(12)
    gPrnSuperIdxOFFDef = Chr$(27) & Chr$(84) & Chr$(1)
    gPrnSuperIdxOnDef = Chr$(27) & Chr$(83) & Chr$(1)
    gPrnTamLetra10CPIDef = Chr$(27) & Chr$(77)
    gPrnTamLetra12CPIDef = Chr$(27) & Chr$(80)
    gPrnTamLetra15CPIDef = Chr$(27) & Chr$(103)
    gPrnTamPagina16Def = Chr$(27) & Chr$(67) & Chr$(16)
    gPrnTamPagina17Def = Chr$(27) & Chr$(67) & Chr$(17)
    gPrnTamPagina18Def = Chr$(27) & Chr$(67) & Chr$(18)
    gPrnTamPagina22Def = Chr$(27) & Chr$(67) & Chr$(22)
    gPrnTamPagina65Def = Chr$(27) & Chr$(67) & Chr$(65)
    gPrnTamPagina66Def = Chr$(27) & Chr$(67) & Chr$(66)
    gPrnTamPagina70Def = Chr$(27) & Chr$(67) & Chr$(70)
    gPrnTamPaginaCabDef = Chr$(27) & Chr$(67)
    gPrnTpoLetraCurierDef = Chr$(27) & Chr$(107) & Chr$(2)
    gPrnTpoLetraCurier1PDef = Chr$(27) & Chr$(120) & Chr$(2)
    gPrnTpoLetraCurierE2Def = Chr$(27) & Chr$(33) & Chr$(2)
    gPrnTpoLetraCurierE240Def = Chr$(27) & Chr$(33) & Chr$(40)
    gPrnTpoLetraRomanDef = Chr$(27) & Chr$(107) & Chr$(0)
    gPrnTpoLetraRoman1PDef = Chr$(27) & Chr$(120) & Chr$(0)
    gPrnTpoLetraRomanE2Def = Chr$(27) & Chr$(33) & Chr$(0)
    gPrnTpoLetraSansSerifDef = Chr$(27) & Chr$(107) & Chr$(1)
    gPrnTpoLetraSansSerif1PDef = Chr$(27) & Chr$(120) & Chr$(1)
    gPrnTpoLetraSansSerifE2Def = Chr$(27) & Chr$(33) & Chr$(1)
    gPrnUnderLineONOFFDef = Chr$(27) & Chr$(45)
    gPrnUnoMedioEspacioDef = Chr$(27) & Chr$(48)

    If pnTipoImp = gEPSON Then
        gPrnBoldOFF = Chr$(27) & Chr$(70)
        gPrnBoldON = Chr$(27) & Chr$(69)
        gPrnCondensadaOFF = Chr$(27) & Chr$(18)
        gPrnCondensadaON = Chr$(27) & Chr$(15)
        'gPrnCondensadaOFF = Chr$(18)
        'gPrnCondensadaON = Chr$(15)
        gPrnDblAnchoOFF = Chr$(27) & Chr$(20)
        gPrnDblAnchoON = Chr$(27) & Chr$(87)
        gPrnEspaLineaN = Chr$(27) & Chr$(50)
        gPrnEspaLineaValor = Chr$(27) & Chr$(65)
        gPrnInicializa = Chr$(27) & Chr$(64)
        gPrnItalicOFF = Chr$(27) & Chr$(53)
        gPrnItalicON = Chr$(27) & Chr$(52)
        gPrnMargenDerCab = Chr$(27) & Chr$(81)
        gPrnMargenIzq00 = Chr$(27) & Chr$(108) & Chr$(0)
        gPrnMargenIzq01 = Chr$(27) & Chr$(108) & Chr$(1)
        gPrnMargenIzq = Chr$(27) & Chr$(108) & Chr$(2)
        gPrnMargenIzq06 = Chr$(27) & Chr$(108) & Chr$(6)
        gPrnMargenIzq44 = Chr$(27) & Chr$(108) & Chr$(44)
        gPrnMargenIzqCab = Chr$(27) & Chr$(108)
        gPrnNegritaOFF = Chr$(27) & Chr$(72)
        gPrnNegritaON = Chr$(27) & Chr$(71)
        gPrnSaltoLinea = Chr$(10)
        gPrnSaltoPagina = Chr$(12)
        gPrnSuperIdxOFF = Chr$(27) & Chr$(84) & Chr$(1)
        gPrnSuperIdxOn = Chr$(27) & Chr$(83) & Chr$(1)
        gPrnTamLetra10CPI = Chr$(27) & Chr$(77)
        gPrnTamLetra12CPI = Chr$(27) & Chr$(80)
        gPrnTamLetra15CPI = Chr$(27) & Chr$(103)
        gPrnTamPagina16 = Chr$(27) & Chr$(67) & Chr$(16)
        gPrnTamPagina17 = Chr$(27) & Chr$(67) & Chr$(17)
        gPrnTamPagina18 = Chr$(27) & Chr$(67) & Chr$(18)
        gPrnTamPagina22 = Chr$(27) & Chr$(67) & Chr$(22)
        gPrnTamPagina65 = Chr$(27) & Chr$(67) & Chr$(65)
        gPrnTamPagina66 = Chr$(27) & Chr$(67) & Chr$(66)
        gPrnTamPagina70 = Chr$(27) & Chr$(67) & Chr$(70)
        gPrnTamPaginaCab = Chr$(27) & Chr$(67)
        gPrnTpoLetraCurier = Chr$(27) & Chr$(107) & Chr$(2)
        gPrnTpoLetraCurier1P = Chr$(27) & Chr$(120) & Chr$(2)
        gPrnTpoLetraCurierE2 = Chr$(27) & Chr$(33) & Chr$(2)
        gPrnTpoLetraCurierE240 = Chr$(27) & Chr$(33) & Chr$(40)
        gPrnTpoLetraRoman = Chr$(27) & Chr$(107) & Chr$(0)
        gPrnTpoLetraRoman1P = Chr$(27) & Chr$(120) & Chr$(0)
        gPrnTpoLetraRomanE2 = Chr$(27) & Chr$(33) & Chr$(0)
        gPrnTpoLetraSansSerif = Chr$(27) & Chr$(107) & Chr$(1)
        gPrnTpoLetraSansSerif1P = Chr$(27) & Chr$(120) & Chr$(1)
        gPrnTpoLetraSansSerifE2 = Chr$(27) & Chr$(33) & Chr$(1)
        gPrnUnderLineONOFF = Chr$(27) & Chr$(45)
        gPrnUnoMedioEspacio = Chr$(27) & Chr$(48)
    ElseIf pnTipoImp = gHEWLETT_PACKARD Or pnTipoImp = gIBM Then
        gPrnBoldOFF = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(48) & Chr$(66)
        gPrnBoldON = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(51) & Chr$(66)
        gPrnCondensadaOFF = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(49) & Chr$(54) & Chr$(72)
        gPrnCondensadaON = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(50) & Chr$(48) & Chr$(72)
        'gPrnCondensadaOFF = Chr$(40) & Chr$(115) & Chr$(49) & Chr$(54) & Chr$(72)
        'gPrnCondensadaON = Chr$(40) & Chr$(115) & Chr$(50) & Chr$(48) & Chr$(72)
        gPrnEspaLineaN = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(54) & Chr$(68)
        gPrnEspaLineaValor = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(54) & Chr$(68)
        gPrnInicializa = Chr$(27) & Chr$(64)
        gPrnInicializaDef = Chr$(27) & Chr$(64)
        gPrnItalicOFF = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(48) & Chr$(66)
        gPrnItalicON = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(49) & Chr$(83)
        gPrnNegritaOFF = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(48) & Chr$(66)
        gPrnNegritaON = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(51) & Chr$(66)
        If pnTipoImp = gHEWLETT_PACKARD Then
           gPrnSaltoLinea = Chr$(10)
        ElseIf pnTipoImp = gIBM Then
           gPrnSaltoLinea = Chr$(10) & Chr$(13)
        End If
        gPrnSaltoPagina = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(48) & Chr$(72)
        gPrnTamLetra10CPI = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(49) & Chr$(54) & Chr$(72)
        gPrnTamLetra12CPI = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(49) & Chr$(54) & Chr$(72)
        gPrnTamLetra15CPI = Chr$(72) & Chr$(27) & Chr$(40) & Chr$(115) & Chr$(49) & Chr$(54)
        gPrnTamPagina16 = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(49) & Chr$(54) & Chr$(80)
        gPrnTamPagina17 = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(49) & Chr$(55) & Chr$(80)
        gPrnTamPagina18 = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(49) & Chr$(56) & Chr$(80)
        gPrnTamPagina22 = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(50) & Chr$(50) & Chr$(80)
        gPrnTamPagina65 = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(54) & Chr$(53) & Chr$(80)
        gPrnTamPagina66 = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(54) & Chr$(54) & Chr$(80)
        gPrnTamPagina70 = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(55) & Chr$(48) & Chr$(80)
        gPrnTamPaginaCab = Chr$(27) & Chr$(38) & Chr$(108) & Chr$(55) & Chr$(48) & Chr$(80)
        gPrnTpoLetraCurier = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(52) & Chr$(49) & Chr$(48) & Chr$(50) & Chr$(84)
        gPrnTpoLetraCurier1P = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(52) & Chr$(49) & Chr$(48) & Chr$(50) & Chr$(84)
        gPrnTpoLetraCurierE2 = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(52) & Chr$(49) & Chr$(48) & Chr$(50) & Chr$(84)
        gPrnTpoLetraCurierE240 = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(52) & Chr$(49) & Chr$(48) & Chr$(50) & Chr$(84)
        gPrnTpoLetraRoman = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(49) & Chr$(54) & Chr$(57) & Chr$(48) & Chr$(49) & Chr$(84)
        gPrnTpoLetraRoman1P = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(49) & Chr$(54) & Chr$(57) & Chr$(48) & Chr$(49) & Chr$(84)
        gPrnTpoLetraRomanE2 = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(49) & Chr$(54) & Chr$(57) & Chr$(48) & Chr$(49) & Chr$(84)
        gPrnTpoLetraSansSerif = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(52) & Chr$(49) & Chr$(48) & Chr$(50) & Chr$(84)
        gPrnTpoLetraSansSerif1P = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(52) & Chr$(49) & Chr$(48) & Chr$(50) & Chr$(84)
        gPrnTpoLetraSansSerifE2 = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(52) & Chr$(49) & Chr$(48) & Chr$(50) & Chr$(84)
        gPrnUnderLineONOFF = Chr$(27) & Chr$(38) & Chr$(100) & Chr$(48) & Chr$(68)
    End If
End Sub

Public Function PrnSet(sCaracterImpresora As String, nValor As Integer) As String
PrnSet = sCaracterImpresora & Chr$(nValor)
End Function

