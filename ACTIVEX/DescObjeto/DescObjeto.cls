VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDescObjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public gsSelecCod As String
Public gsSelecDesc As String
Public lbOk As Boolean
Public gnOrdenBusca As Integer
'variables locales para almacenar los valores de las propiedades
Private mvargsSeleCtasDH As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarlbUltNivel As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarColCod As Long 'copia local
Private mvarColDesc As Long 'copia local
Public Property Let ColDesc(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ColDesc = 5
    mvarColDesc = vData
End Property


Public Property Get ColDesc() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ColDesc
    ColDesc = mvarColDesc
End Property



Public Property Let ColCod(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ColCod = 5
    mvarColCod = vData
End Property


Public Property Get ColCod() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ColCod
    ColCod = mvarColCod
End Property



Public Property Let lbUltNivel(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.lbUltNivel = 5
    mvarlbUltNivel = vData
End Property


Public Property Get lbUltNivel() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.lbUltNivel
    lbUltNivel = mvarlbUltNivel
End Property
Public Property Let gsSeleCtasDH(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.gsSeleCtasDH = 5
    mvargsSeleCtasDH = vData
End Property


Public Property Get gsSeleCtasDH() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.gsSeleCtasDH
    gsSeleCtasDH = mvargsSeleCtasDH
End Property
Public Sub Show(rs As ADODB.Recordset, psObjCod As String, Optional psRaiz As String = "")
    frmDescObjeto.ColCod = mvarColCod
    frmDescObjeto.ColDesc = IIf(mvarColDesc = 0, 1, mvarColDesc)
    frmDescObjeto.vbUltNiv = mvarlbUltNivel
    frmDescObjeto.inicio rs, psObjCod, psRaiz
    gsSelecCod = frmDescObjeto.psDatoCod
    gsSelecDesc = Trim(frmDescObjeto.psDatoDesc)
    lbOk = frmDescObjeto.lOk
    Set frmDescObjeto = Nothing
End Sub
Public Sub ShowGrid(prs As ADODB.Recordset, Optional psTitulo As String = "Búsqueda de Datos")
   frmBusqDiversa.lvUltNivel = mvarlbUltNivel
   frmBusqDiversa.inicio prs, psTitulo
   gsSelecCod = frmBusqDiversa.pCod
   gsSelecDesc = Trim(frmBusqDiversa.pDesc)
   lbOk = frmBusqDiversa.lOk
   Set frmBusqDiversa = Nothing
End Sub
Public Sub BuscarDato(rs As ADODB.Recordset, pnOrden As Integer, psTitulo As String, Optional pnColCod As Integer = 0, Optional pnColDes As Integer = 1)
   frmBuscaDatoGrd.Inicia rs, pnOrden, psTitulo, pnColCod, pnColDes
   gnOrdenBusca = frmBuscaDatoGrd.nOrden
   lbOk = frmBuscaDatoGrd.lOk
   If lbOk Then
      gsSelecCod = rs(pnColCod)
      gsSelecDesc = rs(pnColDes)
   End If
   Set frmBuscaDatoGrd = Nothing
End Sub
Public Sub ShowSeleCuentas(rsDebe As ADODB.Recordset, rsHaber As ADODB.Recordset, psTitulo As String)
    frmSeleCuentas.inicio rsDebe, rsHaber
    lbOk = frmSeleCuentas.lOk
    gsSelecCod = frmSeleCuentas.sCtaCod
    gsSelecDesc = Trim(frmSeleCuentas.sCtaDesc)
    gsSeleCtasDH = frmSeleCuentas.sDH
    Set frmSeleCuentas = Nothing
End Sub


