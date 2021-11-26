VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum eCaptionStyle
    eCap_None = 0
    eCap_CaptionOnly = 1
    eCap_PercentOnly = 2
    eCap_CaptionPercent = 3
End Enum

Dim oFormPadre As Object
'variables locales para almacenar los valores de las propiedades
Private mvarMax As Variant 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarColor1 As ColorConstants 'copia local
Private mvarColor2 As ColorConstants 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarCaptionSyle As eCaptionStyle 'copia local
Public Property Let CaptionSyle(ByVal vData As eCaptionStyle)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.CaptionSyle = Form1
    mvarCaptionSyle = vData
End Property
Public Property Get CaptionSyle() As eCaptionStyle
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CaptionSyle
    CaptionSyle = mvarCaptionSyle
End Property

Public Property Let Color2(ByVal vData As ColorConstants)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Color2 = Form1
    mvarColor2 = vData
End Property
Public Property Get Color2() As ColorConstants
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Color2
    Color2 = mvarColor2
End Property
Public Property Let Color1(ByVal vData As ColorConstants)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Color1 = Form1
    mvarColor1 = vData
End Property
Public Property Get Color1() As ColorConstants
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Color1
    Color1 = mvarColor1
End Property
Public Property Let Max(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Max = 5
    mvarMax = vData
End Property
Public Property Set Max(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Max = Form1
    Set mvarMax = vData
End Property
Public Property Get Max() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Max
    If IsObject(mvarMax) Then
        Set Max = mvarMax
    Else
        Max = mvarMax
    End If
End Property
Public Sub ShowForm(oFormPadre As Object, Optional isModal As Boolean = False)
oFormPadre.Enabled = False
'frmProgress.Show ', oFormPadre
'frmProgress.Inicio oFormPadre

'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
frmProgress.Inicio oFormPadre, isModal

mvarColor1 = vbWhite
mvarColor2 = vbBlue
mvarCaptionSyle = eCap_PercentOnly
End Sub
Public Sub CloseForm(oFormPadre As Object)  'ByVal OwnerForm As Variant
Unload frmProgress
Set frmProgress = Nothing
oFormPadre.Enabled = True
End Sub
Public Function Progress(value, Optional psTitulo As String, Optional psSubTitulo As String = "", Optional psTituloBarra As String, Optional ColorLetras As ColorConstants)
frmProgress.Color1 = mvarColor1
frmProgress.Color2 = mvarColor2
frmProgress.CaptionStyleBarra = mvarCaptionSyle
frmProgress.Max = mvarMax
frmProgress.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Function
'Public Property Get FormPadre() As Object
'Set FormPadre = oFormPadre
'End Property

'Public Property Set FormPadre(ByVal vNewValue As Object)
'Set oFormPadre = vNewValue
'End Property
