VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocPago"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvarvsGlosa As String 'copia local
Private mvarvdFechaDoc As Date 'copia local
Private mvarvbOk As Boolean 'copia local
Private mvarvsTpoDoc As String 'copia local
Private mvarvsNroDoc As String 'copia local
Private mvarvnImporte As Currency 'copia local
Private mvarvsPersNombre As String 'copia local
Private mvarvsNroVoucher As String 'copia local

Private mvarvnMargDer As Integer 'Margen Derecho
Private mvarvnMargIzq As Integer 'Margen Izquiero
Private mvarvnMargSup As Integer 'Margen Superior

'variables locales para almacenar los valores de las propiedades
Private mvarvsDocumento As String 'copia local
Dim psTipoDocVoucher As String
'variables locales para almacenar los valores de las propiedades
Private mvarvsFile As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarvsFormaDoc As String 'copia local


Public Property Let vsFormaDoc(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vsFormaDoc = 5
    mvarvsFormaDoc = vData
End Property
Public Property Get vsFormaDoc() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vsFormaDoc
    vsFormaDoc = mvarvsFormaDoc
End Property
Public Property Let vsFile(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vsFile = 5
    mvarvsFile = vData
End Property
Public Property Get vsFile() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vsFile
    vsFile = mvarvsFile
End Property
Public Property Let vsDocumento(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vsDocumento = 5
    mvarvsDocumento = vData
End Property
Public Property Get vsDocumento() As String
Attribute vsDocumento.VB_UserMemId = 0
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vsDocumento
    vsDocumento = mvarvsDocumento
End Property
Public Property Let vsNroVoucher(ByVal vData As String)
    'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
    'Syntax: X.vsNroVoucher = 5
    mvarvsNroVoucher = vData
End Property
Public Property Get vsNroVoucher() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vsNroVoucher
    vsNroVoucher = mvarvsNroVoucher
End Property
Public Property Let vsPersNombre(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vsPersNombre = 5
    mvarvsPersNombre = vData
End Property
Public Property Get vsPersNombre() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vsPersNombre
    vsPersNombre = mvarvsPersNombre
End Property
Public Property Let vnImporte(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vnImporte = 5
    mvarvnImporte = vData
End Property
Public Property Get vnImporte() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vnImporte
    vnImporte = mvarvnImporte
End Property
Public Property Let vsNroDoc(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vsNroDoc = 5
    mvarvsNroDoc = vData
End Property
Public Property Get vsNroDoc() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vsNroDoc
    vsNroDoc = mvarvsNroDoc
End Property
Public Property Let vsTpoDoc(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vsTpoDoc = 5
    mvarvsTpoDoc = vData
End Property
Public Property Get vsTpoDoc() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vsTpoDoc
    vsTpoDoc = mvarvsTpoDoc
End Property
Public Property Let vbOk(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vbOk = 5
    mvarvbOk = vData
End Property
Public Property Get vbOk() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vbOk
    vbOk = mvarvbOk
End Property
Public Property Let vdFechaDoc(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vdFechaDoc = 5
    mvarvdFechaDoc = vData
End Property
Public Property Get vdFechaDoc() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vdFechaDoc
    vdFechaDoc = mvarvdFechaDoc
End Property
Public Property Let vsGlosa(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.vsGlosa = 5
    mvarvsGlosa = vData
End Property
Public Property Get vsGlosa() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.vsGlosa
    vsGlosa = mvarvsGlosa
End Property

Public Property Let vnMargDer(ByVal vData As Integer)
    mvarvnMargDer = vData
End Property
Public Property Get vnMargDer() As Integer
    vnMargDer = mvarvnMargDer
End Property

Public Property Let vnMargIzq(ByVal vData As Integer)
    mvarvnMargIzq = vData
End Property
Public Property Get vnMargIzq() As Integer
    vnMargIzq = mvarvnMargIzq
End Property

Public Property Let vnMargSup(ByVal vData As Integer)
    mvarvnMargSup = vData
End Property
Public Property Get vnMargSup() As Integer
    vnMargSup = mvarvnMargSup
End Property

Public Sub InicioOrdenPago(ByVal psNroDoc As String, ByVal pbIngresoPers As Boolean, ByVal psPersCod As String, ByVal psOpeCod As String, ByVal psPersNombre As String, ByVal psOpeDesc As String, ByVal psGlosa As String, ByVal pnImporte As Currency, ByVal pdFechaSist As Date, Optional ByVal psDocNroVoucher As String = "", Optional ByVal pbIngPers As Boolean = False, Optional psCodAge As String = "", Optional pbModificaMonto As Boolean = False)
    Dim oContFunc As NContFunciones
    Set oContFunc = New NContFunciones
    frmDocPago.vsNroDoc = psNroDoc
    frmDocPago.vsPersCod = psPersCod
    frmDocPago.lbIngresoPers = pbIngresoPers
    If psDocNroVoucher = "" Then
        vsNroVoucher = oContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(psOpeCod, 3, 1), Right(psCodAge, 2))
    Else
        vsNroVoucher = psDocNroVoucher
    End If
    frmDocPago.InicioOrdenPago psOpeCod, psOpeDesc, psPersNombre, psGlosa, pnImporte, pdFechaSist, vsNroVoucher, pbIngPers, psCodAge, pbModificaMonto

    DoEvents
    CargaDatos
    Unload frmDocPago
    Set frmDocPago = Nothing
    Set oContFunc = Nothing
End Sub

'Public Sub InicioCheque(ByVal psNroDoc As String, ByVal pbIngresoPers As Boolean, ByVal psPersCod As String, ByVal psOpeCod As String, ByVal psPersNombre As String, ByVal psOpeDesc As String, ByVal psGlosa As String, ByVal pnImporte As Currency, ByVal pdFechaSist As Date, ByVal psEmpresaRuc As String, ByVal psSubCtaIF As String, ByVal psEntidadOrig As String, ByVal psCtaEntidadOrig As String, Optional ByVal psDocNroVoucher As String = "", Optional ByVal pbIngPers As Boolean = False, Optional ByVal psCodAge As String = "", Optional ByVal psCtaBanco As String = "", Optional pbModificaMonto As Boolean = False)
Public Sub InicioCheque(ByVal psNroDoc As String, ByVal pbIngresoPers As Boolean, ByVal psPersCod As String, ByVal psOpeCod As String, ByVal psPersNombre As String, ByVal psOpeDesc As String, ByVal psGlosa As String, ByVal pnImporte As Currency, ByVal pdFechaSist As Date, ByVal psEmpresaRuc As String, ByVal psSubCtaIF As String, ByVal psEntidadOrig As String, ByVal psCtaEntidadOrig As String, Optional ByVal psDocNroVoucher As String = "", Optional ByVal pbIngPers As Boolean = False, Optional ByVal psCodAge As String = "", Optional ByVal psCtaBanco As String = "", Optional pbModificaMonto As Boolean = False, Optional ByVal psIFTpo As String = "", Optional ByVal psPersCodIF As String = "", Optional ByVal psCtaIFCod As String = "") 'EJVG20121129
    Dim oContFunc As NContFunciones
    Set oContFunc = New NContFunciones
    frmDocPago.vsNroDoc = psNroDoc
    frmDocPago.vsPersCod = psPersCod
    frmDocPago.lbIngresoPers = pbIngresoPers
    If psDocNroVoucher = "" Then
        vsNroVoucher = oContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(psOpeCod, 3, 1), Right(psCodAge, 2))
    Else
        vsNroVoucher = psDocNroVoucher
    End If
    If Not psCtaBanco = "" Then
       If Not Val(Left(psCtaBanco, 2)) = gTpoCtaIFCtaCte Then
         MsgBox "Solo se puede emitir Cheques en Cuentas Corrientes", vbInformation, "¡Aviso!"
         Exit Sub
       End If
    End If
    'frmDocPago.InicioCheque psOpeCod, psOpeDesc, psPersNombre, psGlosa, pnImporte, pdFechaSist, psEmpresaRuc, psSubCtaIF, psEntidadOrig, psCtaEntidadOrig, vsNroVoucher, pbIngPers, psCodAge, pbModificaMonto
    frmDocPago.InicioCheque psOpeCod, psOpeDesc, psPersNombre, psGlosa, pnImporte, pdFechaSist, psEmpresaRuc, psSubCtaIF, psEntidadOrig, psCtaEntidadOrig, vsNroVoucher, pbIngPers, psCodAge, pbModificaMonto, psIFTpo, psPersCodIF, psCtaIFCod 'EJVG20121129

    DoEvents
    CargaDatos
    Unload frmDocPago
    Set frmDocPago = Nothing
    Set oContFunc = Nothing
End Sub

Public Sub InicioCarta(ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psOpeCod As String, ByVal psOpeDesc As String, ByVal psGlosa As String, ByVal psFile As String, _
                        ByVal pnImporte As Currency, ByVal pdFechaSist As Date, _
                        ByVal psEntiOrig As String, ByVal psCtaEntiOrig As String, _
                        ByVal psEntiDest As String, ByVal psCtaEntiDest As String, _
                        ByVal psMovNro As String, Optional pnMargDer As Integer, Optional pnMargIzq As Integer, Optional pnMargSup As Integer)
                        
    frmDocPago.vsNroDoc = psNroDoc
    frmDocPago.vsPersCod = psPersCod
    vnMargDer = pnMargDer
    vnMargIzq = pnMargIzq
    vnMargSup = pnMargSup
    frmDocPago.InicioCarta psOpeCod, psOpeDesc, psGlosa, psFile, pnImporte, pdFechaSist, psEntiOrig, psCtaEntiOrig, _
                    psEntiDest, psCtaEntiDest, psMovNro
    DoEvents
    vnMargDer = frmDocPago.lnMagDer
    vnMargIzq = frmDocPago.lnMagIzq
    vnMargSup = frmDocPago.lnMagSup
    
    CargaDatos
    Unload frmDocPago
    Set frmDocPago = Nothing
End Sub

Private Sub CargaDatos()
    vsGlosa = frmDocPago.vsGlosa
    vdFechaDoc = frmDocPago.vdFechaDoc
    vbOk = frmDocPago.vbOk
    vsTpoDoc = frmDocPago.vsTpoDoc
    vsNroDoc = frmDocPago.vsNroDoc
    vnImporte = frmDocPago.vnImporte
    vsPersNombre = frmDocPago.vsPersNombre
    vsNroVoucher = frmDocPago.vsNroVoucher
    vsFile = frmDocPago.vsFileCarta
    psTipoDocVoucher = frmDocPago.vsTipoDocVoucher
    vsFormaDoc = frmDocPago.vsFormaDoc
    mvarvsDocumento = frmDocPago.vsFormaDoc
    
End Sub

Public Function ProcesaPlantilla(rtf As String, lImprime As Boolean, psMovNro As String, pdFecha As Date, psEntiOrig As String, _
                                 psEntiDest As String, pnImporte As Currency, psSimbolo As String, psCtaEntiOrig As String, psCtaEntiDest As String, _
                                 psDocNro, pnColPage As Integer, pnMgIzq As Integer, pnMgDer As Integer, psMoneda As String, Optional pnMgSup As Integer = 0) As String
Dim sTemp As String, sTexto As String

    sTemp = rtf
    sTemp = Replace(sTemp, "<<gcMovNro>>", oImpresora.gPrnNegritaON & psMovNro & oImpresora.gPrnNegritaOFF, , 1, vbTextCompare)
    sTemp = Replace(sTemp, "<<gdFecha>>", oImpresora.gPrnNegritaON & GetFecha(pdFecha) & oImpresora.gPrnNegritaOFF, , 1, vbTextCompare)
    sTemp = Replace(sTemp, "<<gcEntiOrig>>", oImpresora.gPrnNegritaON & psEntiOrig & oImpresora.gPrnNegritaOFF, , 1, vbTextCompare)
    sTemp = Replace(sTemp, "<<gcEntiDest>>", oImpresora.gPrnNegritaON & psEntiDest & oImpresora.gPrnNegritaOFF, , 1, vbTextCompare)
    sTemp = Replace(sTemp, "<<gnImporte>>", oImpresora.gPrnNegritaON & psSimbolo & Format(pnImporte, "#,#0.00") & " ( " & ConvNumLet(pnImporte, psMoneda) & " )" & oImpresora.gPrnNegritaOFF, , 1, vbTextCompare)
    sTemp = Replace(sTemp, "<<gcCtaEntiOrig>>", oImpresora.gPrnNegritaON & psCtaEntiOrig & oImpresora.gPrnNegritaOFF, , 1, vbTextCompare)
    sTemp = Replace(sTemp, "<<gcCtaEntiDest>>", oImpresora.gPrnNegritaON & psCtaEntiDest & oImpresora.gPrnNegritaOFF, , 1, vbTextCompare)
    sTemp = Replace(sTemp, "<<gcDocNro>>", oImpresora.gPrnNegritaON & psDocNro & oImpresora.gPrnNegritaOFF, , 1, vbTextCompare)
    ProcesaPlantilla = ProcesaPlantilla & oImpresora.gPrnMargenIzqCab & Chr(pnMgIzq)
    ProcesaPlantilla = ProcesaPlantilla & String(pnMgSup, oImpresora.gPrnSaltoLinea)
    ProcesaPlantilla = ProcesaPlantilla & gGeneral.JustificaTexto(sTemp, pnColPage - pnMgIzq - pnMgDer)
    ProcesaPlantilla = ProcesaPlantilla & oImpresora.gPrnMargenIzqCab & Chr(0)
    
End Function


Private Sub Class_Initialize()
Dim oImp As DImpresoras
    Set oImp = New DImpresoras
oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    
End Sub
'ALPA 20090320*********************
Public Sub InicioPenalidad(ByVal psNroDoc As String, ByVal pbIngresoPers As Boolean, ByVal psPersCod As String, ByVal psOpeCod As String, ByVal psPersNombre As String, ByVal psOpeDesc As String, ByVal psGlosa As String, ByVal pnImporte As Currency, ByVal pdFechaSist As Date, Optional ByVal psDocNroVoucher As String = "", Optional ByVal pbIngPers As Boolean = False, Optional psCodAge As String = "", Optional pbModificaMonto As Boolean = False)
    Dim oContFunc As NContFunciones
    Set oContFunc = New NContFunciones
    frmDocPago.vsNroDoc = psNroDoc
    frmDocPago.vsPersCod = psPersCod
    frmDocPago.lbIngresoPers = pbIngresoPers
    If psDocNroVoucher = "" Then
        vsNroVoucher = oContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(psOpeCod, 3, 1), Right(psCodAge, 2))
    Else
        vsNroVoucher = psDocNroVoucher
    End If
    frmDocPago.InicioPenalidad psOpeCod, psOpeDesc, psPersNombre, psGlosa, pnImporte, pdFechaSist, vsNroVoucher, pbIngPers, psCodAge, pbModificaMonto
    DoEvents
    CargaDatos
    Unload frmDocPago
    Set frmDocPago = Nothing
    Set oContFunc = Nothing
End Sub
'**********************************
