VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogRptImpresionContratos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Contratos de Personal - Proceso de Selección  de Personal"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmLogRptImpresionContratos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7515
      TabIndex        =   12
      Top             =   4275
      Width           =   975
   End
   Begin TabDlg.SSTab TabContratos 
      Height          =   4245
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   7488
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Listado                                                                                  "
      TabPicture(0)   =   "frmLogRptImpresionContratos.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraEmpleados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraContratos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Registro                                                                       "
      TabPicture(1)   =   "frmLogRptImpresionContratos.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblProceso"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FeListado"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtCodigo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtDescriProceso"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraRepresentantes"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame fraContratos 
         Caption         =   "Contratos"
         Height          =   1635
         Left            =   -74955
         TabIndex        =   16
         Top             =   2520
         Width           =   8340
         Begin Sicmact.FlexEdit FeContratos 
            Height          =   1335
            Left            =   135
            TabIndex        =   17
            Top             =   225
            Width           =   8130
            _ExtentX        =   14340
            _ExtentY        =   2355
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Numero-Fecha-Inicio-Fin-cPersCodResp1-cPersCodResp2-cNumero"
            EncabezadosAnchos=   "500-0-2000-1800-1800-1800-0-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbRsLoad        =   -1  'True
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   300
            CellForeColor   =   -2147483627
         End
      End
      Begin VB.Frame fraEmpleados 
         Caption         =   "Empleados"
         Height          =   2040
         Left            =   -74955
         TabIndex        =   14
         Top             =   450
         Width           =   8340
         Begin Sicmact.FlexEdit FeEmpleados 
            Height          =   1605
            Left            =   90
            TabIndex        =   15
            Top             =   270
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2831
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Cargo-cRHCargoCod"
            EncabezadosAnchos=   "500-0-4000-3300-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-C"
            FormatosEdit    =   "0-0-0-0-0"
            TextArray0      =   "#"
            lbRsLoad        =   -1  'True
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   300
            CellForeColor   =   -2147483627
         End
      End
      Begin VB.Frame fraRepresentantes 
         Caption         =   "Representantes del Empleado"
         Height          =   1095
         Left            =   45
         TabIndex        =   5
         Top             =   2970
         Width           =   8340
         Begin Sicmact.TxtBuscar txtRepresentante1 
            Height          =   345
            Left            =   1545
            TabIndex        =   6
            Top             =   270
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin Sicmact.TxtBuscar txtRepresentante2 
            Height          =   345
            Left            =   1545
            TabIndex        =   7
            Top             =   630
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin VB.Label lblExpositor 
            Caption         =   "1er Representante:"
            Height          =   255
            Left            =   90
            TabIndex        =   11
            Top             =   315
            Width           =   1410
         End
         Begin VB.Label lblRepresentante1 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   3195
            TabIndex        =   10
            Top             =   270
            Width           =   5025
         End
         Begin VB.Label Label1 
            Caption         =   "2do Representante:"
            Height          =   255
            Left            =   90
            TabIndex        =   9
            Top             =   675
            Width           =   1410
         End
         Begin VB.Label lblRepresentante2 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   3195
            TabIndex        =   8
            Top             =   630
            Width           =   5025
         End
      End
      Begin VB.TextBox txtDescriProceso 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   345
         Left            =   2370
         TabIndex        =   1
         Top             =   540
         Width           =   5970
      End
      Begin Sicmact.TxtBuscar txtCodigo 
         Height          =   360
         Left            =   810
         TabIndex        =   2
         Top             =   540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin Sicmact.FlexEdit FeListado 
         Height          =   1830
         Left            =   90
         TabIndex        =   3
         Top             =   990
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   3228
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Codigo-Nombre-Inicio-Fin-Sueldo-Cargo-Fecha-cRHCargoCod-cRHContratoNro"
         EncabezadosAnchos=   "500-0-3000-1200-1200-1200-2600-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-4-5-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-2-2-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-R-R-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0"
         TextArray0      =   "#"
         lbRsLoad        =   -1  'True
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   495
         RowHeight0      =   300
         CellForeColor   =   -2147483627
      End
      Begin VB.Label lblProceso 
         Caption         =   "Proceso:"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   585
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   45
      TabIndex        =   13
      Top             =   4275
      Width           =   975
   End
End
Attribute VB_Name = "frmLogRptImpresionContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub Grabar_Impresion()
Dim oContratacion As New DContratacionProceso
Dim sNumero As String

With FeListado
    sNumero = oContratacion.Obtener_Numero_Impresion
    Call oContratacion.InsertaRHContratoDet_Impresion(.TextMatrix(.row, 1), .TextMatrix(.row, 9), CDate(.TextMatrix(.row, 7)), txtRepresentante1.Text, txtRepresentante2.Text, sNumero)
End With
Set oContratacion = Nothing
End Sub

Private Sub cmdImprimir_Click()

'If ValidaDatosImpresion = False Then Exit Sub

If TabContratos.Tab = 0 Then
    If FeContratos.TextMatrix(FeContratos.row, 2) = "" Then
        MsgBox "Debe indicar el Contrato a imprimir", vbInformation, "Mensaje"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Call Imprime_Contrato_Historico
    Screen.MousePointer = vbDefault
Else
    Screen.MousePointer = vbHourglass
    If ValidaDatosImpresion = False Then Exit Sub
    If MsgBox("Esta seguro de realizar la Operacion?", vbQuestion + vbYesNo, "Confirmación") = vbNo Then Exit Sub
    Call Imprime_Contrato_en_Linea
    Call Grabar_Impresion
    Screen.MousePointer = vbDefault
End If
End Sub

Sub Imprime_Contrato_Historico()

Dim oContratacion As New DContratacionProceso
Dim rs As New ADODB.Recordset
Dim rsFBasicas As New ADODB.Recordset
Dim rsFEspecif As New ADODB.Recordset
Dim rsSubFunc As New ADODB.Recordset

Dim sNumero As String
Dim sRepresentante1 As String
Dim sRepresentante2 As String
Dim sEmpleado As String
Dim sDocumento As String
Dim sDomicilio As String
Dim sTrabajador As String
Dim sCargo As String
Dim sInicio As String
Dim sFin As String
Dim sSueldo As String
Dim sFecha As String
Dim sFuncionesBasicas As String
Dim sFuncionesEspecificas As String
Dim sCodPersona As String
Dim sCodCargo As String
Dim nFila As Integer
Dim sNumContrato As String
Dim dFechaContrato As Date

With FeContratos
    sCodPersona = .TextMatrix(.row, 1)
    sNumContrato = .TextMatrix(.row, 2)
    dFechaContrato = CDate(.TextMatrix(.row, 3))
    sNumero = .TextMatrix(.row, 8)
    sRepresentante1 = RTrim(oContratacion.Obtener_Datos_Representante_Contrato(.TextMatrix(.row, 6)))
    sRepresentante2 = RTrim(oContratacion.Obtener_Datos_Representante_Contrato(.TextMatrix(.row, 7)))
End With
With FeEmpleados
    sCodCargo = .TextMatrix(.row, 4)
End With

Set rs = oContratacion.Obtener_Datos_Para_ImpresionContrato(sCodPersona, sNumContrato, dFechaContrato)
If Not rs.EOF Then
    sEmpleado = PstaNombre(rs("Empleado"), True)
    sDocumento = rs("Documento")
    sDomicilio = RTrim(rs("Domicilio"))
    sTrabajador = rs("Trabajador")
    sCargo = rs("Cargo")
    sInicio = Format(rs("Inicio"), gsFormatoFechaView)
    sFin = Format(rs("Fin"), gsFormatoFechaView)
    sSueldo = rs("Sueldo")
    sFecha = rs("Fecha")
End If
nFila = 0
sFuncionesBasicas = ""
Set rsFBasicas = oContratacion.Obtener_Funciones_x_Cargo(sCodCargo, 1)
While Not rsFBasicas.EOF
    nFila = nFila + 1
    sFuncionesBasicas = sFuncionesBasicas & CStr(nFila) & ". " & rsFBasicas("cFuncionDescripcion") & Chr(13) 'vbCrLf
    rsFBasicas.MoveNext
Wend

nFila = 0
sFuncionesEspecificas = ""
Set rsFEspecif = oContratacion.Obtener_Funciones_x_Cargo(sCodCargo, 0)
While Not rsFEspecif.EOF
    nFila = nFila + 1
    sFuncionesEspecificas = sFuncionesEspecificas & CStr(nFila) & ". " & rsFEspecif("cFuncionDescripcion") & Chr(13) 'vbCrLf
    Set rsSubFunc = oContratacion.Obtener_SubFunciones_x_Funcion(rsFEspecif("nFuncionCod"))
    While Not rsSubFunc.EOF
        sFuncionesEspecificas = sFuncionesEspecificas & Space(4) & "- " & rsSubFunc("cSubFuncionDescripcion") & Chr(13) 'vbCrLf
        rsSubFunc.MoveNext
    Wend
    rsFEspecif.MoveNext
Wend

Call ImpPlantillaContrato(sNumero, sRepresentante1, sRepresentante2, sEmpleado, sDocumento, sDomicilio, sTrabajador, sCargo, sInicio, sFin, sSueldo, sFecha, sFuncionesBasicas, sFuncionesEspecificas)

Set rs = Nothing
Set oContratacion = Nothing

End Sub

Sub Imprime_Contrato_en_Linea()

Dim oContratacion As New DContratacionProceso
Dim rs As New ADODB.Recordset
Dim rsFBasicas As New ADODB.Recordset
Dim rsFEspecif As New ADODB.Recordset
Dim rsSubFunc As New ADODB.Recordset

Dim sNumero As String
Dim sRepresentante1 As String
Dim sRepresentante2 As String
Dim sEmpleado As String
Dim sDocumento As String
Dim sDomicilio As String
Dim sTrabajador As String
Dim sCargo As String
Dim sInicio As String
Dim sFin As String
Dim sSueldo As String
Dim sFecha As String
Dim sFuncionesBasicas As String
Dim sFuncionesEspecificas As String
Dim sCodPersona As String
Dim sCodCargo As String
Dim nFila As Integer

With FeListado
    sCodPersona = .TextMatrix(.row, 1)
    sCodCargo = .TextMatrix(.row, 8)
End With

sNumero = oContratacion.Obtener_Numero_Impresion

sRepresentante1 = RTrim(oContratacion.Obtener_Datos_Representante_Contrato(txtRepresentante1.Text))
sRepresentante2 = RTrim(oContratacion.Obtener_Datos_Representante_Contrato(txtRepresentante2.Text))

Set rs = oContratacion.Obtener_Datos_Para_ImpresionContrato(sCodPersona)
If Not rs.EOF Then
    sEmpleado = PstaNombre(rs("Empleado"), True)
    sDocumento = rs("Documento")
    sDomicilio = RTrim(rs("Domicilio"))
    sTrabajador = rs("Trabajador")
    sCargo = rs("Cargo")
    sInicio = Format(rs("Inicio"), gsFormatoFechaView)
    sFin = Format(rs("Fin"), gsFormatoFechaView)
    sSueldo = rs("Sueldo")
    sFecha = rs("Fecha")
End If
nFila = 0
sFuncionesBasicas = ""
Set rsFBasicas = oContratacion.Obtener_Funciones_x_Cargo(sCodCargo, 1)
While Not rsFBasicas.EOF
    nFila = nFila + 1
    sFuncionesBasicas = sFuncionesBasicas & CStr(nFila) & ". " & rsFBasicas("cFuncionDescripcion") & Chr(13) 'vbCrLf
    rsFBasicas.MoveNext
Wend

nFila = 0
sFuncionesEspecificas = ""
Set rsFEspecif = oContratacion.Obtener_Funciones_x_Cargo(sCodCargo, 0)
While Not rsFEspecif.EOF
    nFila = nFila + 1
    sFuncionesEspecificas = sFuncionesEspecificas & CStr(nFila) & ". " & rsFEspecif("cFuncionDescripcion") & Chr(13)
    Set rsSubFunc = oContratacion.Obtener_SubFunciones_x_Funcion(rsFEspecif("nFuncionCod"))
    While Not rsSubFunc.EOF
        sFuncionesEspecificas = sFuncionesEspecificas & Space(4) & "- " & rsSubFunc("cSubFuncionDescripcion") & Chr(13) 'vbCrLf
        rsSubFunc.MoveNext
    Wend
    rsFEspecif.MoveNext
Wend

Call ImpPlantillaContrato(sNumero, sRepresentante1, sRepresentante2, sEmpleado, sDocumento, sDomicilio, sTrabajador, sCargo, sInicio, sFin, sSueldo, sFecha, sFuncionesBasicas, sFuncionesEspecificas)

Set rs = Nothing
Set oContratacion = Nothing
End Sub

Function ImpPlantillaContrato(psNumero As String, psRepresentante1 As String, psRepresentante2 As String, _
                            psEmpleado As String, psDocumento As String, psDomicilio As String, psTrabajador As String, _
                            psCargo As String, psInicio As String, psFin As String, psSueldo As String, psFecha As String, _
                            psFuncionesBasicas As String, psFuncionesEspecificas As String) As Boolean

Dim lsModeloPlantilla As String
Dim sTemporal As String
Dim nNumCadenas As Integer
Dim nTempo As Integer

On Error GoTo ErrorReemplazo
'******** NUEVAS VARIABLES ********

lsModeloPlantilla = App.path & "\SPOOLER\" & "Plantilla_Contratos.doc"
'lsModeloPlantilla = "C:\PROYECTO\Sicmac_Admin\spooler\Plantilla_Contratos.doc"
'CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))

    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
    
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
    'wAppSource.ActiveDocument
    
    'Crea Nuevo Documento
    wApp.Documents.Add

        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.Paste
        wApp.Application.Selection.InsertBreak
        
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
        
        With wApp.Selection.Find
            .Text = "<<NUMERO>>"
            .Replacement.Text = psNumero
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll

        With wApp.Selection.Find
            .Text = "<<REPRESENTANTE1>>"
            .Replacement.Text = psRepresentante1
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<REPRESENTANTE2>>"
            .Replacement.Text = psRepresentante2
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<EMPLEADO>>"
            .Replacement.Text = psEmpleado
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<DOCUMENTO>>"
            .Replacement.Text = psDocumento
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With

        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<DOMICILIO>>"
            .Replacement.Text = psDomicilio
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<TRABAJADOR>>"
            .Replacement.Text = psTrabajador
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<CARGO>>"
            .Replacement.Text = psCargo
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<INICIO>>"
            .Replacement.Text = psInicio
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<FIN>>"
            .Replacement.Text = psFin
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<SUELDO>>"
            .Replacement.Text = psSueldo & "(" & UCase(NumLet(psSueldo))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<<FECHA>>"
            .Replacement.Text = psFecha
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        'With wApp.Selection.Find
        '    .Text = "<<FUNCION BASICA>>"
        '    .Replacement.Text = "Reemplaza Funciones Basicas" & vbCrLf & vbCrLf & vbCrLf
        '    .Forward = True
        '    .Wrap = wdFindContinue
        '    .Format = False
        'End With
        'wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        nNumCadenas = Len(psFuncionesBasicas) / 100
        If (Len(psFuncionesBasicas) Mod 100) > 0 Then
            nNumCadenas = nNumCadenas + 1
        End If
        'sTemporal = "<<FUNCION BASICA>>"
        For nTempo = 1 To nNumCadenas
        With wApp.Selection.Find
            .Text = "<<FUNCIONES BASICAS>>"
            .Replacement.Text = Mid(psFuncionesBasicas, 1 + (100 * (nTempo - 1)), 100) & IIf(nTempo <> nNumCadenas, "<<FUNCIONES BASICAS>>", "")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        Next

        nNumCadenas = Len(psFuncionesEspecificas) / 100
        If (Len(psFuncionesEspecificas) Mod 100) > 0 Then
            nNumCadenas = nNumCadenas + 1
        End If

        For nTempo = 1 To nNumCadenas
        With wApp.Selection.Find
            .Text = "<<FUNCIONES ESPECIFICAS>>"
            .Replacement.Text = Mid(psFuncionesEspecificas, 1 + (100 * (nTempo - 1)), 100) & IIf(nTempo <> nNumCadenas, "<<FUNCIONES ESPECIFICAS>>", "")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        Next

wAppSource.ActiveDocument.Close
wApp.Visible = True
Exit Function

ErrorReemplazo: MsgBox "Problemas en el documento de Word,intente de nuevo mas tarde", vbInformation, "Mensaje"
                wAppSource.ActiveDocument.Close
                'wApp.Visible = True

End Function

Function ValidaDatosImpresion() As Boolean

ValidaDatosImpresion = True
If txtRepresentante1.Text = "" Or txtRepresentante2.Text = "" Then
    MsgBox "Debe indicar los Representantes del Empleado", vbInformation, "Mensaje"
    ValidaDatosImpresion = False
    Exit Function
End If

With FeListado
    If .TextMatrix(.row, 1) = "" Then
        MsgBox "Debe seleccionar el Empleado a generar su Contrato", vbInformation, "Mensaje"
        ValidaDatosImpresion = False
        Exit Function
    End If
End With
End Function

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub FeEmpleados_Click()
Dim oContratacion As New DContratacionProceso
Dim rs As New ADODB.Recordset

FeContratos.Clear
FeContratos.Rows = 2
FeContratos.FormaCabecera

With FeEmpleados
    If .TextMatrix(.row, 1) = "" Then Exit Sub
    Set rs = oContratacion.Obtener_Contratos_x_Empleado(.TextMatrix(.row, 1))
End With
FeContratos.rsFlex = rs

Set rs = Nothing
Set oContratacion = Nothing

End Sub

Private Sub Form_Load()
    txtCodigo.rs = Obtener_Procesos_para_Contrato
    Call Llenar_Lista_Empleados
End Sub

Sub Llenar_Lista_Empleados()
Dim oContratacion As New DContratacionProceso
Dim rs As New ADODB.Recordset

Set rs = oContratacion.Obtener_Personal_ProcesosSeleccion
FeEmpleados.rsFlex = rs

Set rs = Nothing
Set oContratacion = Nothing
End Sub

Function Obtener_Procesos_para_Contrato() As ADODB.Recordset
Dim oselproceso As New NSeleccionProceso

Set Obtener_Procesos_para_Contrato = oselproceso.Obtener_DatosProceso_x_Estado(7)

Set oselproceso = Nothing
End Function

Private Sub TabContratos_Click(PreviousTab As Integer)

If TabContratos.Tab = 1 Then
    With FeListado
        .Clear
        .Rows = 2
        .FormaCabecera
    End With

    txtCodigo.Text = ""
    txtDescriProceso.Text = ""
    txtRepresentante1.Text = ""
    txtRepresentante2.Text = ""
    lblRepresentante1.Caption = ""
    lblRepresentante2.Caption = ""
End If
End Sub

Private Sub txtCodigo_EmiteDatos()
If txtCodigo.Text = "" Then Exit Sub
Dim oContratacion As New DContratacionProceso
Dim rs As New ADODB.Recordset

Set rs = oContratacion.Obtener_Personal_Ingresante(txtCodigo.Text)
txtDescriProceso.Text = txtCodigo.psDescripcion
FeListado.rsFlex = rs

Set rs = Nothing
Set oContratacion = Nothing
End Sub

Private Sub txtRepresentante1_EmiteDatos()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio(True)
If X Is Nothing Then Exit Sub

If Len(Trim(X.sPersNombre)) > 0 Then
    txtRepresentante1.Text = X.sPersCod
    lblRepresentante1.Caption = X.sPersNombre
End If

End Sub

Private Sub txtRepresentante2_EmiteDatos()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio(True)
If X Is Nothing Then Exit Sub

If Len(Trim(X.sPersNombre)) > 0 Then
    txtRepresentante2.Text = X.sPersCod
    lblRepresentante2.Caption = X.sPersNombre
End If

End Sub
