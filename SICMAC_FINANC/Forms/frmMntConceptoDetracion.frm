VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMntConceptoDetracion 
   Caption         =   "Detracción"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   Icon            =   "frmMntConceptoDetracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   14970
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabDet 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Conceptos de Detracción"
      TabPicture(0)   =   "frmMntConceptoDetracion.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FEDetraccion"
      Tab(0).Control(1)=   "cmdNuevoDetCab"
      Tab(0).Control(2)=   "cmdModificarDetCab"
      Tab(0).Control(3)=   "cmdEliminarDetCab"
      Tab(0).Control(4)=   "cmdCancelarDetCab"
      Tab(0).Control(5)=   "cmdAceptarDetCab"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Cuentas Contables con Detracción"
      TabPicture(1)   =   "frmMntConceptoDetracion.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FECtaDet"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdNuevoDetDet"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdModificarDetDet"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdEliminarDetDet"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdCancelarDetDet"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdAceptarDetDet"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdAceptarDetDet 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   13440
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarDetDet 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   13440
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminarDetDet 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   13440
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdModificarDetDet 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   13440
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuevoDetDet 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   13440
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin Sicmact.FlexEdit FECtaDet 
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   5106
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Concepto-Item-Cta. Con.-nDetraCod-cCtaContCod"
         EncabezadosAnchos=   "500-6900-500-4900-1-1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-L-R-R"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "Nro"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   5
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdAceptarDetCab 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   -61560
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarDetCab 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   -61560
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminarDetCab 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   -61560
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdModificarDetCab 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   -61560
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuevoDetCab 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   -61560
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin Sicmact.FlexEdit FEDetraccion 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   5106
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Concepto-Porcentaje-Rango Inicio-Rango Final-Documento-Codigo"
         EncabezadosAnchos=   "500-6900-1000-1200-1200-2000-1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-X"
         ListaControles  =   "0-0-0-0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-L-R"
         FormatosEdit    =   "0-0-2-4-4-0-0"
         TextArray0      =   "Nro"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmMntConceptoDetracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmConceptoDetracion
'***Descripción:    Formulario para mantenimiento de los conceptos de detracción
'***Creación:       ELRO el 20111222 según Acta 328-2011/TI-D
'************************************************************
Option Explicit

Private Enum Accion
gValorDefectoAccion = 0
gNuevoRegistro = 1
gEditarRegistro = 2
gEliminarRegistro = 3
End Enum

Private fnAccion, fnFilaNoEditar, fnFilaNoEditar2 As Integer

Private fnCodigoDetraccion As Currency
Private fsConcepto, fsDocumento As String
Private fnPorcentaje, fnRangoInicio, fnRangoFinal As Currency

Private Sub CargarFEDetraccion()
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim rsConceptoDetracion As ADODB.Recordset
    Set rsConceptoDetracion = New ADODB.Recordset
    Dim i As Integer
    fnAccion = gValorDefectoAccion
    fnFilaNoEditar = -1

    Set rsConceptoDetracion = oNContFunciones.obtenerConceptoDetracciones
   
    Call LimpiaFlex(FEDetraccion)
    
    FEDetraccion.lbEditarFlex = True
    
   
    If Not rsConceptoDetracion.BOF And Not rsConceptoDetracion.EOF Then
    i = 1
        Do While Not rsConceptoDetracion.EOF
            FEDetraccion.AdicionaFila
            FEDetraccion.TextMatrix(i, 1) = UCase(rsConceptoDetracion!cDetraDes)
            FEDetraccion.TextMatrix(i, 2) = rsConceptoDetracion!nDetraPorc
            FEDetraccion.TextMatrix(i, 3) = Format(rsConceptoDetracion!Rango1, "#,#0.00")
            FEDetraccion.TextMatrix(i, 4) = Format(rsConceptoDetracion!Rango2, "#,#0.00")
            FEDetraccion.TextMatrix(i, 5) = rsConceptoDetracion!cDocumento
            FEDetraccion.TextMatrix(i, 6) = rsConceptoDetracion!nDetraCod
            i = i + 1
            rsConceptoDetracion.MoveNext
        Loop
    Else
           MsgBox "No existe conceptos de Detracción registrados", vbInformation, "Aviso"
    End If
    FEDetraccion.lbEditarFlex = False
End Sub

Private Sub CargarFECtaDet()
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim rsFECtaDet As ADODB.Recordset
    Set rsFECtaDet = New ADODB.Recordset
    Dim i As Integer
    fnAccion = gValorDefectoAccion
    fnFilaNoEditar2 = -1

    Set rsFECtaDet = oNContFunciones.obtenerCtaDetraccionesDet(CInt(FEDetraccion.TextMatrix(FEDetraccion.Row, 6)))
   
    Call LimpiaFlex(FECtaDet)
    
    FECtaDet.lbEditarFlex = True
       
    If Not rsFECtaDet.BOF And Not rsFECtaDet.EOF Then
        i = 1
        Do While Not rsFECtaDet.EOF
            FECtaDet.AdicionaFila
            FECtaDet.TextMatrix(i, 1) = UCase(rsFECtaDet!cDetraDes)
            FECtaDet.TextMatrix(i, 2) = rsFECtaDet!nItem
            FECtaDet.TextMatrix(i, 3) = UCase(rsFECtaDet!CtaConDes)
            FECtaDet.TextMatrix(i, 4) = rsFECtaDet!nDetraCod
            FECtaDet.TextMatrix(i, 5) = rsFECtaDet!cCtaContCod
            rsFECtaDet.MoveNext
        Loop
        SSTabDet.Tab = 1
        FECtaDet.SetFocus
    Else
           MsgBox "No existe Cuentas Contables que aplican Detracción ", vbInformation, "Aviso"
    End If
    FECtaDet.lbEditarFlex = False
End Sub

Private Sub cargarDocumento()
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim rsDocumentos As ADODB.Recordset
    Set rsDocumentos = New ADODB.Recordset
    
    Set rsDocumentos = oNContFunciones.obtenerDocumentos
        
    FEDetraccion.CargaCombo rsDocumentos
   
    Set rsDocumentos = Nothing
    Set oNContFunciones = Nothing
    
End Sub

Private Function validarDatosConcepto() As Boolean

validarDatosConcepto = False

    If Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 1)) = "" Then
        MsgBox "Falta ingresar el concepto", vbInformation, "Aviso"
        FEDetraccion.SetFocus
        Exit Function
    End If
    
    If Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 2)) = "" Then
        MsgBox "Falta ingresar el porcentaje", vbInformation, "Aviso"
        FEDetraccion.SetFocus
        Exit Function
    End If
    
    If CCur(Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 2))) = 0# Then
        MsgBox "El porcentaje debe ser mayor que cero", vbInformation, "Aviso"
        FEDetraccion.SetFocus
        Exit Function
    End If
    
    If Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 3)) = "" Then
        MsgBox "Falta ingresar el rango inicio", vbInformation, "Aviso"
        FEDetraccion.SetFocus
        Exit Function
    End If
    
    If CCur(Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 3))) = 0# Then
        MsgBox "El rango inicio debe ser mayor que cero", vbInformation, "Aviso"
        FEDetraccion.SetFocus
        Exit Function
    End If
    
    If Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 4)) = "" Then
        MsgBox "Falta ingresar el rango final", vbInformation, "Aviso"
        FEDetraccion.SetFocus
        Exit Function
    End If
    
    If CCur(Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 4))) = 0# Then
        MsgBox "El rango final debe ser mayor que cero", vbInformation, "Aviso"
        FEDetraccion.SetFocus
        Exit Function
    End If

    If Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 5)) = "" Then
        MsgBox "Falta elegir el documento", vbInformation, "Aviso"
        FEDetraccion.SetFocus
        Exit Function
    End If

validarDatosConcepto = True
End Function

Private Function validarDatosCtaDetraccion() As Boolean
Dim nRegistro As Integer

validarDatosCtaDetraccion = False

If Trim(FECtaDet.TextMatrix(FECtaDet.Row, 3)) = "" Then
    MsgBox "Ingrese la cuenta contable", vbInformation, "Aviso"
    FECtaDet.SetFocus
    Exit Function
End If

nRegistro = validarCuentaContableConsolidada(Left(Format(gsMesCerrado, "yyyyMMdd"), 4), Mid(Format(gsMesCerrado, "yyyyMMdd"), 5, 2), Trim(FECtaDet.TextMatrix(FECtaDet.Row, 3)))
If nRegistro = 0 Then
    MsgBox "Solo puede ingresar cuentas contables consilidadas", vbInformation, "Aviso"
    FECtaDet.SetFocus
    Exit Function
End If

validarDatosCtaDetraccion = True
End Function

Private Sub cmdAceptarDetCab_Click()
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim lsMovNro As String

If validarDatosConcepto = False Then Exit Sub

lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

If fnAccion = gNuevoRegistro Then
 
    
    Call oNContFunciones.registrarConceptoDetracion(FEDetraccion.TextMatrix(FEDetraccion.Row, 1), _
                                                    FEDetraccion.TextMatrix(FEDetraccion.Row, 2), _
                                                    CInt(Right(FEDetraccion.TextMatrix(FEDetraccion.Row, 5), 4)), _
                                                    FEDetraccion.TextMatrix(FEDetraccion.Row, 3), _
                                                    FEDetraccion.TextMatrix(FEDetraccion.Row, 4), _
                                                    lsMovNro)
    

    
End If

If fnAccion = gEditarRegistro Then

    
    Call oNContFunciones.actualizarConceptoDetraccion(FEDetraccion.TextMatrix(FEDetraccion.Row, 6), _
                                                      FEDetraccion.TextMatrix(FEDetraccion.Row, 1), _
                                                      FEDetraccion.TextMatrix(FEDetraccion.Row, 2), _
                                                      CInt(Trim(Right(FEDetraccion.TextMatrix(FEDetraccion.Row, 5), 4))), _
                                                      FEDetraccion.TextMatrix(FEDetraccion.Row, 3), _
                                                      FEDetraccion.TextMatrix(FEDetraccion.Row, 4), _
                                                      lsMovNro)

End If

Set oNContFunciones = Nothing
cmdNuevoDetCab.Visible = True
cmdModificarDetCab.Visible = True
cmdEliminarDetCab.Visible = True
cmdAceptarDetCab.Visible = False
cmdCancelarDetCab.Visible = False
cmdNuevoDetDet.Visible = True
cmdModificarDetDet.Visible = True
cmdEliminarDetDet.Visible = True
FEDetraccion.lbEditarFlex = False
fnAccion = gValorDefectoAccion
CargarFEDetraccion '***Agregado por ELRO el 20130605, según SATI INC1306040011
End Sub

Private Sub cmdAceptarDetDet_Click()
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim lsMovNro As String
Dim lnItem As Integer

'***Modificado por ELRO el 20130605, según SATI INC1306040011****
'If validarDatosConcepto = False Then Exit Sub
If validarDatosCtaDetraccion = False Then Exit Sub
'***Fin Modificado por ELRO el 20130605, según SATI INC1306040011

lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

If fnAccion = gNuevoRegistro Then
 
    Call oNContFunciones.registrarCtaDetraccionesDet(CInt(FEDetraccion.TextMatrix(FEDetraccion.Row, 6)), _
                                                     CInt(FECtaDet.TextMatrix(FECtaDet.Row, 2)), _
                                                     FECtaDet.TextMatrix(FECtaDet.Row, 3), _
                                                     lsMovNro)
    

    
End If

If fnAccion = gEditarRegistro Then

    
    Call oNContFunciones.actualizarCtaDetraccionesDet(CInt(FECtaDet.TextMatrix(FECtaDet.Row, 4)), _
                                                      CInt(FECtaDet.TextMatrix(FECtaDet.Row, 2)), _
                                                      FECtaDet.TextMatrix(FECtaDet.Row, 3), _
                                                      lsMovNro)

End If

Set oNContFunciones = Nothing

cmdNuevoDetCab.Visible = True
cmdModificarDetCab.Visible = True
cmdEliminarDetCab.Visible = True
cmdAceptarDetCab.Visible = False
cmdCancelarDetCab.Visible = False
cmdNuevoDetDet.Visible = True
cmdModificarDetDet.Visible = True
cmdEliminarDetDet.Visible = True
cmdAceptarDetDet.Visible = False
cmdCancelarDetDet.Visible = False
FECtaDet.lbEditarFlex = False
fnAccion = gValorDefectoAccion
CargarFECtaDet '***Agregado por ELRO el 20130605, según SATI INC1306040011
SSTabDet.TabEnabled(0) = True '***Agregado por ELRO el 20130605, según SATI INC1306040011
End Sub

Private Sub cmdCancelarDetCab_Click()
Call cargarDocumento
Call CargarFEDetraccion
Call CargarFECtaDet
cmdNuevoDetCab.Visible = True
cmdModificarDetCab.Visible = True
cmdEliminarDetCab.Visible = True
cmdAceptarDetCab.Visible = False
cmdCancelarDetCab.Visible = False
cmdNuevoDetDet.Visible = True
cmdModificarDetDet.Visible = True
cmdEliminarDetDet.Visible = True
cmdAceptarDetDet.Visible = False
cmdCancelarDetDet.Visible = False
FEDetraccion.lbEditarFlex = False
fnAccion = gValorDefectoAccion
End Sub

Private Sub cmdCancelarDetDet_Click()
Call CargarFECtaDet
cmdNuevoDetCab.Visible = True
cmdModificarDetCab.Visible = True
cmdEliminarDetCab.Visible = True
cmdAceptarDetCab.Visible = False
cmdCancelarDetCab.Visible = False
cmdNuevoDetDet.Visible = True
cmdModificarDetDet.Visible = True
cmdEliminarDetDet.Visible = True
cmdAceptarDetDet.Visible = False
cmdCancelarDetDet.Visible = False
FECtaDet.lbEditarFlex = False
fnAccion = gValorDefectoAccion
SSTabDet.TabEnabled(0) = True '***Agregado por ELRO el 20130605, según SATI INC1306040011
End Sub

Private Sub cmdModificarDetCab_Click()
cmdNuevoDetCab.Visible = False
cmdModificarDetCab.Visible = False
cmdEliminarDetCab.Visible = False
cmdAceptarDetCab.Visible = True
cmdCancelarDetCab.Visible = True
cmdNuevoDetDet.Visible = False
cmdModificarDetDet.Visible = False
cmdEliminarDetDet.Visible = False
cmdAceptarDetDet.Visible = False
cmdCancelarDetDet.Visible = False
fnAccion = gEditarRegistro
FEDetraccion.lbEditarFlex = True
fnFilaNoEditar = FEDetraccion.Row
End Sub

Private Sub cmdModificarDetDet_Click()
cmdNuevoDetCab.Visible = False
cmdModificarDetCab.Visible = False
cmdEliminarDetCab.Visible = False
cmdAceptarDetCab.Visible = False
cmdCancelarDetCab.Visible = False
cmdNuevoDetDet.Visible = False
cmdModificarDetDet.Visible = False
cmdEliminarDetDet.Visible = False
cmdAceptarDetDet.Visible = True
cmdCancelarDetDet.Visible = True
fnAccion = gEditarRegistro
FECtaDet.lbEditarFlex = True
fnFilaNoEditar2 = FEDetraccion.Row
FECtaDet.TextMatrix(FECtaDet.Row, 3) = FECtaDet.TextMatrix(FECtaDet.Row, 5)
SSTabDet.TabEnabled(0) = False '***Agregado por ELRO el 20130605, según SATI INC1306040011
End Sub

Private Sub cmdEliminarDetCab_Click()
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim lsMovNro As String

If Trim(FECtaDet.TextMatrix(1, 1)) <> "" Then
    MsgBox "No se puede eliminar el concepto porque tiene detalle"
    FEDetraccion.SetFocus
    Exit Sub
End If

If MsgBox("¿Esta seguro que desea eliminar el Concepto " & Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 1)) & "?", vbYesNo, "Aviso") = vbYes Then
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call oNContFunciones.eliminarConceptoDetraccion(CInt(Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 6))), lsMovNro)
    Call cargarDocumento
    Call CargarFEDetraccion
    Call CargarFECtaDet
End If

End Sub

Private Sub cmdEliminarDetDet_Click()
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim lsMovNro As String

If MsgBox("¿Esta seguro que desea eliminar la Cuenta Contable aplicable a detracción " & Trim(FECtaDet.TextMatrix(FECtaDet.Row, 3)) & "?", vbYesNo, "Aviso") = vbYes Then
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call oNContFunciones.eliminarCtaDetraccionesDet(CInt(Trim(FEDetraccion.TextMatrix(FEDetraccion.Row, 6))), CInt(FECtaDet.TextMatrix(FECtaDet.Row, 2)), lsMovNro)
    Call cargarDocumento
    Call CargarFEDetraccion
    Call CargarFECtaDet
End If

End Sub

Private Sub FEDetraccion_OnCellChange(pnRow As Long, pnCol As Long)
    If fnFilaNoEditar > -1 Then
        If validarDatosConcepto() = False Then
            Exit Sub
        End If
    End If
End Sub

Private Sub FECtaDet_OnCellChange(pnRow As Long, pnCol As Long)
    If fnFilaNoEditar2 > -1 Then
        If validarDatosCtaDetraccion() = False Then
            Exit Sub
        End If
    End If
End Sub

Private Sub FEDetraccion_OnRowChange(pnRow As Long, pnCol As Long)
If fnFilaNoEditar = -1 Then
    Call CargarFECtaDet
End If
End Sub

'***Comentado por ELRO el 20130605, según SATI INC1306040011****
'Private Sub Form_Activate()
'    FEDetraccion.SetFocus
'End Sub
'***Fin Comentado por ELRO el 20130605, según SATI INC1306040011

Private Sub Form_Load()
    Call cargarDocumento
    Call CargarFEDetraccion
End Sub

Private Sub cmdNuevoDetCab_Click()
cmdNuevoDetCab.Visible = False
cmdModificarDetCab.Visible = False
cmdEliminarDetCab.Visible = False
cmdAceptarDetCab.Visible = True
cmdCancelarDetCab.Visible = True
cmdNuevoDetDet.Visible = False
cmdModificarDetDet.Visible = False
cmdEliminarDetDet.Visible = False
cmdAceptarDetDet.Visible = False
cmdCancelarDetDet.Visible = False
fnAccion = gNuevoRegistro
FEDetraccion.lbEditarFlex = True
FEDetraccion.AdicionaFila
fnFilaNoEditar = FEDetraccion.Rows - 1
Call LimpiaFlex(FECtaDet) '***Agregado por ELRO el 20130605, según SATI INC1306040011
End Sub

Private Sub cmdNuevoDetDet_Click()
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
cmdNuevoDetCab.Visible = False
cmdModificarDetCab.Visible = False
cmdEliminarDetCab.Visible = False
cmdAceptarDetCab.Visible = False
cmdCancelarDetCab.Visible = False
cmdNuevoDetDet.Visible = False
cmdModificarDetDet.Visible = False
cmdEliminarDetDet.Visible = False
cmdAceptarDetDet.Visible = True
cmdCancelarDetDet.Visible = True
fnAccion = gNuevoRegistro
FECtaDet.lbEditarFlex = True
FECtaDet.AdicionaFila
fnFilaNoEditar2 = FEDetraccion.Rows - 1
FECtaDet.TextMatrix(FECtaDet.Row, 1) = FEDetraccion.TextMatrix(FEDetraccion.Row, 1)
FECtaDet.TextMatrix(FECtaDet.Row, 2) = oNContFunciones.obtenerNumeroItemCtaDetraccionesDet(CInt(FEDetraccion.TextMatrix(FEDetraccion.Row, 6)))
SSTabDet.TabEnabled(0) = False '***Agregado por ELRO el 20130605, según SATI INC1306040011
Set oNContFunciones = Nothing
End Sub

Private Sub FEDetraccion_RowColChange()
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim rsMeses As ADODB.Recordset
    Set rsMeses = New ADODB.Recordset
   
    If FEDetraccion.lbEditarFlex Then
        If fnFilaNoEditar <> -1 Then
            FEDetraccion.Row = fnFilaNoEditar
        End If
        Set rsMeses = oNContFunciones.obtenerDocumentos
        Select Case FEDetraccion.Col
           Case 3
                FEDetraccion.CargaCombo rsMeses
        End Select
    End If
     Set rsMeses = Nothing
    Set oNContFunciones = Nothing
End Sub

Private Sub FECtaDet_RowColChange()
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim rsMeses As ADODB.Recordset
    Set rsMeses = New ADODB.Recordset
    Set rsMeses = Nothing
    Set oNContFunciones = Nothing
End Sub



