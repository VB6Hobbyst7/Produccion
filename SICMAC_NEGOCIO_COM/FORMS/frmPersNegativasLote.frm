VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmPersNegativasLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista Negativa: Carga Lote"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   Icon            =   "frmPersNegativasLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Selección de Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   7815
      Begin VB.TextBox txtNomArchivo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   280
         Width           =   3735
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar Formato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         Top             =   240
         Width           =   1530
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton cmdBuscarArchivo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Información"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3855
      Begin VB.OptionButton optPersJur 
         Caption         =   "Personas Jurídicas"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton opPersNat 
         Caption         =   "Personas Naturales"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame fraBonoPlus 
      Caption         =   "Lista de Personas a Registrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11775
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   1170
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   8
         Top             =   4560
         Width           =   1170
      End
      Begin SICMACT.FlexEdit fePersNegativas 
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   11520
         _extentx        =   20320
         _extenty        =   7223
         cols0           =   23
         highlight       =   1
         encabezadosnombres=   $"frmPersNegativasLote.frx":030A
         encabezadosanchos=   "500-0-1800-2000-1600-1600-1600-3000-2000-2000-2000-2000-2000-2000-3000-2000-1000-0-0-0-0-1200-1200"
         font            =   "frmPersNegativasLote.frx":03F1
         fontfixed       =   "frmPersNegativasLote.frx":0419
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         encabezadosalineacion=   "C-C-C-L-L-L-L-L-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1  'True
         lbultimainstancia=   -1  'True
         tipobusqueda    =   3
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   495
         rowheight0      =   300
      End
      Begin MSComctlLib.ProgressBar BarraProgreso 
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   4680
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   661
      Filtro          =   "Contratos Digital (*.pdf)|*.pdf"
      Altura          =   0
   End
   Begin ComctlLib.StatusBar EstadoBarra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   6405
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersNegativasLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmPersNegativasLote
'***     Descripcion:      Opcion para realizar la carga de Personas Negativas en Lote
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     12/07/2013 10:52:16 AM
'***     Ultima Modificacion: Fecha de Crecion
'*****************************************************************************************
Option Explicit
Private fsPathFile As String
Private fsRuta As String
Private fsNomFile As String
Private fnTipoPersona As Integer
Private nExcel As Integer
Private nFlex As Integer
Private i As Integer
Private Type PersNegativa
    NumDoc As String
    NombreRazon As String
    ApePat As String
    ApeMat As String
    ApeCas As String
    Delito As String
    Juzgado As String
    OFMultiple As String
    OFJuzgado As String
    Expediente As String
    Departamento As String
    CodDepartamento As String
    tipo As String
    CodTipo As String
    Comentario As String
    Condicion As String
    CodCondicion As String
    existe As String
    CadError As String
    cMovNro As String
    Institucion As String 'marg
    Cargo As String 'marg
    
End Type

Dim fmPersNegativa() As PersNegativa

Private Sub cmdBuscarArchivo_Click()
LimpiaFlex fePersNegativas
Dim i As Integer
CdlgFile.nHwd = Me.hwnd
CdlgFile.Filtro = "Archivos Excel (*.xls)|*.xls"
CdlgFile.altura = 300
CdlgFile.Show

fsPathFile = CdlgFile.Ruta
fsRuta = fsPathFile
        If fsPathFile <> Empty Then
            For i = Len(fsPathFile) - 1 To 1 Step -1
                    If Mid(fsPathFile, i, 1) = "\" Then
                        fsPathFile = Mid(CdlgFile.Ruta, 1, i)
                        fsNomFile = Mid(CdlgFile.Ruta, i + 1, Len(CdlgFile.Ruta) - i)
                        Exit For
                    End If
             Next i
          Screen.MousePointer = 11
          txtNomArchivo.Text = fsNomFile
        Else
           MsgBox "No se Selecciono Ningun Archivo", vbInformation, "Aviso"
           txtNomArchivo.Text = ""
           LimpiaFlex fePersNegativas
           Exit Sub
        End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdCargar_Click()
If Not ValidaDatos Then Exit Sub
If MsgBox("Estas seguro de cargar el archivo adjuntado?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
On Error GoTo ErrorCargaArchivo

LimpiaFlex Me.fePersNegativas
cmdProcesar.Enabled = False

Dim oConstante As COMDConstantes.DCOMConstantes
Dim oPersonas As COMDPersona.DCOMPersonas
Dim rsDepartamento As ADODB.Recordset
Dim rsTipo As ADODB.Recordset
Dim rsCondicion As ADODB.Recordset
Dim rsPersona As ADODB.Recordset

Set oPersonas = New COMDPersona.DCOMPersonas
Set oConstante = New COMDConstantes.DCOMConstantes


Set rsDepartamento = oPersonas.CargarUbicacionesGeograficas(True, 5, "04028")
Set rsTipo = oConstante.RecuperaConstantes(9991)
Set rsCondicion = oConstante.RecuperaConstantes(9072)

Set oPersonas = Nothing
Set oConstante = Nothing



'Variables para los Datos a Cargar
Dim lsNumDoc As String
Dim lsNombreRazon As String
Dim lsApePat As String
Dim lsApeMat As String
Dim lsApeCas As String
Dim lsDelito As String
Dim lsJuzgado As String
Dim lsOFMultiple As String
Dim lsOFJuzgado As String
Dim lsExpediente As String
Dim lsDepartamento As String
Dim lsCodDepartamento As String
Dim lsTipo As String
Dim lsCodTipo As String
Dim lsComentario As String
Dim lsCondicion As String
Dim lsCodCondicion As String
Dim lsCodCondicionBD As String
Dim lcMovNro As String
'marg------------
Dim lsInstitucion As String
Dim lsCargo As String
'--------------
Dim lnTipoPersBD As String

Dim lsExiste As String
Dim lsCadError As String

Dim lbHayDatos As Boolean
Dim lbError As Boolean
Dim lbExisteBD As Boolean
Dim lbJustificacion As Boolean

Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim lbExisteHoja As Boolean
Dim lsHoja As String

If fnTipoPersona = 1 Then
    lsHoja = "PersonasNaturales"
Else
    lsHoja = "PersonasJuridicas"
End If
    

Set xlsAplicacion = New Excel.Application
    
Set xlsLibro = xlsAplicacion.Workbooks.Open(fsRuta)

'Activa la hoja correspondiente
For Each xlHoja In xlsLibro.Worksheets
   If UCase(Trim(xlHoja.Name)) = UCase(Trim(lsHoja)) Then
        xlHoja.Activate
        lbExisteHoja = True
    Exit For
   End If
Next
ReDim fmPersNegativa(0)

If lbExisteHoja = False Then
    MsgBox "El Nombre de la Hoja debe ser ''" & lsHoja & "''", vbCritical, "Aviso"
    xlsAplicacion.Quit
    xlsAplicacion.Visible = False
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja = Nothing
   
    Exit Sub
End If


LimpiaFlex fePersNegativas

nExcel = 5
nFlex = 1

lbHayDatos = True
lbError = False

Do While lbHayDatos
    lbJustificacion = False
    lcMovNro = ""
    lsCadError = ""
    lsCodCondicionBD = 0
    
    lsNumDoc = Trim(xlHoja.Cells(nExcel, 2))
    lsNombreRazon = UCase(Trim(xlHoja.Cells(nExcel, 3)))
    
    lsApePat = ""
    lsApeMat = ""
    lsApeCas = ""
    
    If fnTipoPersona = 1 Then
        lsApePat = UCase(Trim(xlHoja.Cells(nExcel, 4)))
        lsApeMat = UCase(Trim(xlHoja.Cells(nExcel, 5)))
        lsApeCas = UCase(Trim(xlHoja.Cells(nExcel, 6)))
        lsDelito = Trim(xlHoja.Cells(nExcel, 7))
        lsJuzgado = Trim(xlHoja.Cells(nExcel, 8))
        lsOFMultiple = Trim(xlHoja.Cells(nExcel, 9))
        lsOFJuzgado = Trim(xlHoja.Cells(nExcel, 10))
        lsExpediente = Trim(xlHoja.Cells(nExcel, 11))
        lsDepartamento = Trim(xlHoja.Cells(nExcel, 12))
        lsTipo = Trim(xlHoja.Cells(nExcel, 13))
        lsComentario = Trim(xlHoja.Cells(nExcel, 14))
        lsCondicion = Trim(xlHoja.Cells(nExcel, 15))
        lsInstitucion = Trim(xlHoja.Cells(nExcel, 16)) 'marg
        lsCargo = Trim(xlHoja.Cells(nExcel, 17)) 'marg
        lsExiste = Trim(xlHoja.Cells(nExcel, 18))
    Else
        lsDelito = Trim(xlHoja.Cells(nExcel, 4))
        lsJuzgado = Trim(xlHoja.Cells(nExcel, 5))
        lsOFMultiple = Trim(xlHoja.Cells(nExcel, 6))
        lsOFJuzgado = Trim(xlHoja.Cells(nExcel, 7))
        lsExpediente = Trim(xlHoja.Cells(nExcel, 8))
        lsDepartamento = Trim(xlHoja.Cells(nExcel, 9))
        lsTipo = Trim(xlHoja.Cells(nExcel, 10))
        lsComentario = Trim(xlHoja.Cells(nExcel, 11))
        lsCondicion = Trim(xlHoja.Cells(nExcel, 12))
        lsInstitucion = Trim(xlHoja.Cells(nExcel, 13)) 'marg
        lsCargo = Trim(xlHoja.Cells(nExcel, 14)) 'marg
        lsExiste = Trim(xlHoja.Cells(nExcel, 15))
    End If
    
    If Trim(lsNumDoc) = "" And Trim(lsNombreRazon) = "" And Trim(lsApePat) = "" And Trim(lsApeMat) = "" _
        And Trim(lsApeCas) = "" And Trim(lsDelito) = "" And Trim(lsJuzgado) = "" And Trim(lsOFMultiple) = "" _
        And Trim(lsOFJuzgado) = "" And Trim(lsExpediente) = "" And Trim(lsDepartamento) = "" And Trim(lsTipo) = "" _
        And Trim(lsComentario) = "" And Trim(lsCondicion) = "" And Trim(lsExiste) = "" Then
            lbHayDatos = False
            Exit Do
    End If
    
    
    If fnTipoPersona = 1 Then
        If Trim(lsNombreRazon) = "" Then
            lbError = True
            lsCadError = lsCadError & " - (Columna Nombres - Dato no valido)"
        End If
        
        If Trim(lsApePat) = "" And Trim(lsApeMat) = "" And Trim(lsApeCas) = "" Then
            lbError = True
            lsCadError = lsCadError & " - (Columna Apellido Paterno, Materno y/o Casada - Debe consignar por lo menos un dato en uno de los campos)"
        End If
    Else
        If Trim(lsNombreRazon) = "" Then
            lbError = True
            lsCadError = lsCadError & " - (Columna Razon Social - Dato no valido)"
        End If
    End If
    
   
    lsCodDepartamento = ""
    lsCodTipo = ""
    lsCodCondicion = ""
    
    If Trim(lsDelito) = "" And Trim(lsJuzgado) = "" And Trim(lsOFMultiple) = "" And Trim(lsOFJuzgado) = "" And _
           Trim(lsExpediente) = "" And Trim(lsDepartamento) = "" And Trim(lsTipo) = "" And Trim(lsComentario) = "" Then
        lbJustificacion = False
    Else
        lbJustificacion = True
    End If
    
    If lbJustificacion Then
        If Trim(lsDelito) = "" And Trim(lsOFMultiple) = "" Then
            lbError = True
            lsCadError = lsCadError & " - (Columna Delito - Dato no valido)"
            lsCadError = lsCadError & " - (Columna Oficio Múltiple - Dato no valido)"
        End If
    
        If Trim(lsDelito) = "" Then
            lsDelito = "."
        End If
    End If
        
    lbError = ValidaRegistros(1, rsDepartamento, lsDepartamento, lsCodDepartamento, lsCadError, lbError)
    lbError = ValidaRegistros(2, rsTipo, lsTipo, lsCodTipo, lsCadError, lbError)
    lbError = ValidaRegistros(3, rsCondicion, lsCondicion, lsCodCondicion, lsCadError, lbError)
    
    
    
    If Trim(lsCodTipo) = "" Then
        lsCodTipo = "0"
    End If
    
    If Trim(lsCodCondicion) = "" Then
        lbError = True
        lsCadError = lsCadError & " - (Columna Condición-Dato no valido)"
    End If
    
    If Trim(lsCodCondicion) = "1" Then
        If lbJustificacion = False Then
            If (Trim(lsDelito) = "" And Trim(lsOFMultiple) = "") Then
                lbError = True
                lsCadError = lsCadError & " - (Columna Delito-Dato no valido)"
                lsCadError = lsCadError & " - (Columna Oficio Múltiple-Dato no valido)"
            End If
            
             If Trim(lsDelito) = "" Then
                    lsDelito = "."
                End If
        End If
    End If
    
    If Not (Trim(UCase(lsExiste)) = "SI" Or Trim(UCase(lsExiste)) = "NO") Then
        lbError = True
        lsCadError = lsCadError & " - (Columna Existe-Dato no valido)"
    End If
    
    Set oPersonas = New COMDPersona.DCOMPersonas
    Set rsPersona = oPersonas.ObtenerDatosBasicosPersNegativa(lsNumDoc, lsNombreRazon, lsApePat, lsApeMat, lsApeCas)
    lbExisteBD = False
    If Not (rsPersona.EOF And rsPersona.BOF) Then
        lcMovNro = Trim(rsPersona!cMovNro)
        lsCodCondicionBD = Trim(rsPersona!nCondicion)
        lnTipoPersBD = CInt(Trim(rsPersona!nTipoPers))
        
        lbExisteBD = True
        If Trim(UCase(lsExiste)) = "NO" Then
            lbError = True
            lsCadError = lsCadError & " - (Persona Existe en la Base de Datos )"
        End If
    Else
        lnTipoPersBD = fnTipoPersona
        lbExisteBD = False
        lcMovNro = ""
        lsCodCondicionBD = ""
        
        If Trim(UCase(lsExiste)) = "SI" Then
            lbError = True
            lsCadError = lsCadError & " - (Persona No Existe en la Base de Datos)"
        End If
    End If
    Set rsPersona = Nothing
    
     If Trim(lsExiste) = "NO" Then
        If Trim(lsNumDoc) <> "" Then
            Set rsPersona = oPersonas.ObtenerDatosBasicosPersNegativa(Trim(lsNumDoc), , , , , 1)
            If Not (rsPersona.EOF And rsPersona.BOF) Then
                lbError = True
                lsCadError = lsCadError & " - (Numero de Documento ya existe en la base de datos)"
            End If
        End If
    End If
    
    
    If lbExisteBD Then
        If Trim(lsCodCondicion) <> Trim(lsCodCondicionBD) Then
            lbError = True
            lsCadError = lsCadError & " - (Condición Diferente al dato registrado en la Base de Datos)"
        End If
    End If
    
    If lnTipoPersBD <> fnTipoPersona Then
        lbError = True
        lsCadError = lsCadError & " - (No es una Persona " & IIf(fnTipoPersona = 1, "Natural", "Juridica") & ")"
    End If
    
   
    
    ReDim Preserve fmPersNegativa(nFlex)
    
    fmPersNegativa(nFlex).NumDoc = lsNumDoc
    fmPersNegativa(nFlex).NombreRazon = lsNombreRazon
    fmPersNegativa(nFlex).ApePat = lsApePat
    fmPersNegativa(nFlex).ApeMat = lsApeMat
    fmPersNegativa(nFlex).ApeCas = lsApeCas
    fmPersNegativa(nFlex).Delito = lsDelito
    fmPersNegativa(nFlex).Juzgado = lsJuzgado
    fmPersNegativa(nFlex).OFMultiple = lsOFMultiple
    fmPersNegativa(nFlex).OFJuzgado = lsOFJuzgado
    fmPersNegativa(nFlex).Expediente = lsExpediente
    fmPersNegativa(nFlex).Departamento = lsDepartamento
    fmPersNegativa(nFlex).tipo = lsTipo
    fmPersNegativa(nFlex).Comentario = lsComentario
    fmPersNegativa(nFlex).Condicion = lsCondicion
    fmPersNegativa(nFlex).Institucion = lsInstitucion 'marg
    fmPersNegativa(nFlex).Cargo = lsCargo 'marg
    fmPersNegativa(nFlex).existe = lsExiste
    
    fmPersNegativa(nFlex).CodDepartamento = lsCodDepartamento
    fmPersNegativa(nFlex).CodTipo = lsCodTipo
    fmPersNegativa(nFlex).CodCondicion = lsCodCondicion
    fmPersNegativa(nFlex).CadError = lsCadError
    
    fmPersNegativa(nFlex).cMovNro = lcMovNro
    lbHayDatos = True
    
    nExcel = nExcel + 1
    nFlex = nFlex + 1
    Set oPersonas = Nothing
    Set rsPersona = Nothing
Loop


lbError = ValidaDatosGeneral(lbError)
If lbError Then
    If MsgBox("Existen errores en el Archivo, Desea Exportar la Lista de Errores?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        Call ExportarExcelError(fmPersNegativa())
    End If
    cmdProcesar.Enabled = False
Else
    For i = 1 To UBound(fmPersNegativa)
        fePersNegativas.AdicionaFila
        fePersNegativas.TextMatrix(i, 1) = 1
        fePersNegativas.TextMatrix(i, 2) = fmPersNegativa(i).NumDoc
        fePersNegativas.TextMatrix(i, 3) = fmPersNegativa(i).NombreRazon
        fePersNegativas.TextMatrix(i, 4) = fmPersNegativa(i).ApePat
        fePersNegativas.TextMatrix(i, 5) = fmPersNegativa(i).ApeMat
        fePersNegativas.TextMatrix(i, 6) = fmPersNegativa(i).ApeCas
        fePersNegativas.TextMatrix(i, 7) = fmPersNegativa(i).Delito
        fePersNegativas.TextMatrix(i, 8) = fmPersNegativa(i).Juzgado
        fePersNegativas.TextMatrix(i, 9) = fmPersNegativa(i).OFMultiple
        fePersNegativas.TextMatrix(i, 10) = fmPersNegativa(i).OFJuzgado
        fePersNegativas.TextMatrix(i, 11) = fmPersNegativa(i).Expediente
        fePersNegativas.TextMatrix(i, 12) = fmPersNegativa(i).Departamento
        fePersNegativas.TextMatrix(i, 13) = fmPersNegativa(i).tipo
        fePersNegativas.TextMatrix(i, 14) = fmPersNegativa(i).Comentario
        fePersNegativas.TextMatrix(i, 15) = fmPersNegativa(i).Condicion
        fePersNegativas.TextMatrix(i, 16) = fmPersNegativa(i).existe
        
        fePersNegativas.TextMatrix(i, 17) = fmPersNegativa(i).CodDepartamento
        fePersNegativas.TextMatrix(i, 18) = fmPersNegativa(i).CodTipo
        fePersNegativas.TextMatrix(i, 19) = fmPersNegativa(i).CodCondicion
        fePersNegativas.TextMatrix(i, 20) = fmPersNegativa(i).cMovNro
        fePersNegativas.TextMatrix(i, 21) = fmPersNegativa(i).Institucion 'marg
        fePersNegativas.TextMatrix(i, 22) = fmPersNegativa(i).Cargo 'marg
    Next
    MsgBox "Datos Cargados Correctamente", vbInformation, "Aviso"
    cmdProcesar.Enabled = True
End If

fePersNegativas.TopRow = 1
xlsAplicacion.Quit
    
xlsAplicacion.Visible = False
Set xlsAplicacion = Nothing
Set xlsLibro = Nothing
Set xlHoja = Nothing

Exit Sub

ErrorCargaArchivo:
xlsAplicacion.Quit
Set xlsAplicacion = Nothing
Set xlsLibro = Nothing
Set xlHoja = Nothing

MsgBox Err.Description & " - Al cargar el Archivo verifique la Columnas que encuentren en la columnas correctas", vbCritical, "Aviso"

End Sub

Public Function ValidaDatosGeneral(ByVal pbValida As Boolean) As Boolean
Dim j As Integer
Dim bEsError As Boolean
Dim nContError As Long
bEsError = False
nContError = 0



For i = 1 To UBound(fmPersNegativa)
    For j = 1 To UBound(fmPersNegativa)
        If i <> j Then
            If fmPersNegativa(i).existe = "NO" Then
                
                
                If fmPersNegativa(i).NumDoc = fmPersNegativa(j).NumDoc And fmPersNegativa(i).NombreRazon = fmPersNegativa(j).NombreRazon And _
                    fmPersNegativa(i).ApePat = fmPersNegativa(j).ApePat And fmPersNegativa(i).ApeMat = fmPersNegativa(j).ApeMat And _
                    fmPersNegativa(i).ApeCas = fmPersNegativa(j).ApeCas Then
                    
                    If fmPersNegativa(j).CodCondicion <> fmPersNegativa(i).CodCondicion Then
                        nContError = nContError + 1
                        fmPersNegativa(i).CadError = fmPersNegativa(i).CadError & " - (Condición Incorrecta con respecto a la fila " & j & ")"
                    End If
                Else
                    If fmPersNegativa(i).NumDoc = fmPersNegativa(j).NumDoc Then
                        If fmPersNegativa(i).NumDoc <> "" Then
                            nContError = nContError + 1
                            fmPersNegativa(i).CadError = fmPersNegativa(i).CadError & " - (Nº de Documento duplicado con la fila " & j & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next j
Next i

If nContError > 0 Then
    bEsError = True
End If

If Not pbValida Then
    If bEsError Then
        ValidaDatosGeneral = True
    Else
        ValidaDatosGeneral = False
    End If
Else
    ValidaDatosGeneral = True
End If
End Function

Private Sub cmdCerrar_Click()
     Unload Me
End Sub

Private Sub cmdGenerar_Click()

Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String
Dim lsArchivoMostrar As String

Dim lsRuta As String
Dim lsHoja As String
Dim lsHojaDescarte As String

Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim xlHojaDel As Excel.Worksheet
Dim lbExisteHoja As Boolean

    ' marg
    lsArchivo = "PersNegativasLote2"
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    lsRuta = App.Path & "\FormatoCarta\" & lsArchivo & ".xls"
    
    If fnTipoPersona = 1 Then
        lsHoja = "PersonasNaturales"
        lsHojaDescarte = "PersonasJuridicas"
    Else
        lsHoja = "PersonasJuridicas"
        lsHojaDescarte = "PersonasNaturales"
    End If
    
    lsArchivoMostrar = "\Spooler\" & lsHoja & "_" & gsCodUser & "_" & gsCodAge & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(lsRuta) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(lsRuta)
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
  
    'Activa la hoja correspondiente
    For Each xlHoja In xlsLibro.Worksheets
       If UCase(Trim(xlHoja.Name)) = UCase(Trim(lsHoja)) Then
            xlHoja.Activate
            lbExisteHoja = True
        Exit For
       End If
    Next
    
    
    If lbExisteHoja = False Then
        Set xlHoja = xlsLibro.Worksheets
        xlHoja.Name = lsHoja
    End If
    
    'Elimina lo que no corresponde
    For Each xlHojaDel In xlsLibro.Worksheets
        If UCase(Trim(xlHojaDel.Name)) = UCase(Trim(lsHojaDescarte)) Then
            xlHojaDel.Visible = xlSheetHidden
        End If
    Next

    lsArchivoMostrar = App.Path & lsArchivoMostrar
    
    xlHoja.SaveAs lsArchivoMostrar
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja = Nothing
    
End Sub

Private Sub cmdProcesar_Click()
If MsgBox("Estas seguro de Procesar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
On Error GoTo ErrorProcesarArchivo

Dim oPersonas As COMDPersona.DCOMPersonas
Dim rsPer As ADODB.Recordset
Set oPersonas = New COMDPersona.DCOMPersonas


Dim Hora As String
Hora = Time


cmdProcesar.Enabled = False
BarraProgreso.Visible = True

BarraProgreso.value = 0
BarraProgreso.Min = 0
BarraProgreso.value = 0

EstadoBarra.Panels(1).Visible = True
EstadoBarra.Panels(1) = "Generando..."
    
BarraProgreso.Max = UBound(fmPersNegativa)

EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
For i = 1 To UBound(fmPersNegativa)
    Set oPersonas = New COMDPersona.DCOMPersonas
    If Trim(UCase(fmPersNegativa(i).existe)) = "NO" Then
        Set rsPer = oPersonas.ObtenerDatosBasicosPersNegativa(fmPersNegativa(i).NumDoc, fmPersNegativa(i).NombreRazon, fmPersNegativa(i).ApePat, fmPersNegativa(i).ApeMat, fmPersNegativa(i).ApeCas)
    
        If Not (rsPer.EOF And rsPer.BOF) Then
            fmPersNegativa(i).cMovNro = Trim(rsPer!cMovNro)
        Else
            fmPersNegativa(i).cMovNro = Mid(gdFecSis, 7, 4) & Mid(gdFecSis, 4, 2) & Mid(gdFecSis, 1, 2) & Format(DateAdd("s", i, Hora), "hhmmss") & "109" & gsCodAge & "00" & gsCodUser

            Call oPersonas.RegistraPersNegativo(fnTipoPersona, fnTipoPersona, fmPersNegativa(i).NumDoc, _
                            fmPersNegativa(i).NombreRazon, fmPersNegativa(i).ApePat, fmPersNegativa(i).ApeMat, "", "", fmPersNegativa(i).Institucion, fmPersNegativa(i).Cargo, _
                            fmPersNegativa(i).cMovNro, CInt(fmPersNegativa(i).CodCondicion), , fmPersNegativa(i).ApeCas) 'marg
        End If
    End If
    
    If Not (Trim(fmPersNegativa(i).Delito) = "" And Trim(fmPersNegativa(i).Juzgado) = "" And Trim(fmPersNegativa(i).OFMultiple) = "" And Trim(fmPersNegativa(i).OFJuzgado) = "" _
            And Trim(fmPersNegativa(i).Expediente) = "" And Trim(fmPersNegativa(i).CodDepartamento) = "" And Trim(fmPersNegativa(i).CodTipo) = "0" _
            And Trim(fmPersNegativa(i).Comentario) = "") Then
                    
        Call oPersonas.ModificaPersNegativoJustifica(fmPersNegativa(i).cMovNro, fmPersNegativa(i).Delito, fmPersNegativa(i).Juzgado, _
                        fmPersNegativa(i).OFMultiple, fmPersNegativa(i).OFJuzgado, fmPersNegativa(i).Expediente, fmPersNegativa(i).CodDepartamento, _
                        CInt(fmPersNegativa(i).CodTipo), fmPersNegativa(i).Comentario)
    End If
    BarraProgreso.value = i
    EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
    
    Set oPersonas = Nothing
Next

Set rsPer = Nothing
EstadoBarra.Panels(1) = "Proceso Finalizado"
Call ExportarExcelArchivoProcesado(fmPersNegativa)
MsgBox "Datos Procesados Satisfactoriamente.", vbInformation, "Aviso"
EstadoBarra.Panels(1).Visible = False
BarraProgreso.Visible = False
LimpiaFlex fePersNegativas

ReDim fmPersNegativa(0)
ReDim fmPersNegativaIns(0)

Exit Sub
ErrorProcesarArchivo:
MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub fePersNegativas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim sColumnas() As String
sColumnas = Split(fePersNegativas.ColumnasAEditar, "-")
If sColumnas(pnCol) = "X" Then
   Cancel = False
   SendKeys "{Tab}", True
   Exit Sub
End If
End Sub

Private Sub Form_Load()
ReDim fmPersNegativa(0)
ReDim fmPersNegativaIns(0)
fnTipoPersona = 1
cmdProcesar.Enabled = False
BarraProgreso.Visible = False
EstadoBarra.Panels(1).Visible = False
End Sub


Private Sub opPersNat_Click()
txtNomArchivo.Text = ""
fsRuta = ""
LimpiaFlex fePersNegativas
If opPersNat.value Then
    fnTipoPersona = 1
    fePersNegativas.EncabezadosAnchos = "500-0-1800-2000-1600-1600-1600-3000-2000-2000-2000-2000-2000-2000-3000-2000-1000-0-0-0-0-1200-1200"
    fePersNegativas.EncabezadosNombres = "#-Aux-Nº Documento-Nombre-Apellido Paterno-Apellido Materno-Apellido Casada-Delito-Juzgado-Oficio Múltiple-Oficio Juzgado-Expediente-Departamento-Tipo-Comentario-Condición-Existe-CodDep-CodTipo-CodCond-cMovNro-Institucion-Cargo" 'marg
Else
    fnTipoPersona = 2
    fePersNegativas.EncabezadosAnchos = "500-0-1800-5000-0-0-0-3000-2000-2000-2000-2000-2000-2000-3000-2000-1000-0-0-0-0-1200-1200"
    fePersNegativas.EncabezadosNombres = "#-Aux-Nº Documento-Razon Social-Apellido Paterno-Apellido Materno-Apellido Casada-Delito-Juzgado-Oficio Múltiple-Oficio Juzgado-Expediente-Departamento-Tipo-Comentario-Condición-Existe-CodDep-CodTipo-CodCond-cMovNro-Institucion-Cargo" 'marg
End If
End Sub

Private Sub optPersJur_Click()
txtNomArchivo.Text = ""
fsRuta = ""
LimpiaFlex fePersNegativas
If optPersJur.value Then
    fnTipoPersona = 2
    fePersNegativas.EncabezadosAnchos = "500-0-1800-5000-0-0-0-3000-2000-2000-2000-2000-2000-2000-3000-2000-1000-0-0-0-0-1200-1200"
    fePersNegativas.EncabezadosNombres = "#-Aux-Nº Documento-Razon Social-Apellido Paterno-Apellido Materno-Apellido Casada-Delito-Juzgado-Oficio Múltiple-Oficio Juzgado-Expediente-Departamento-Tipo-Comentario-Condición-Existe-CodDep-CodTipo-CodCond-cMovNro-Institucion-Cargo" 'marg
Else
    fnTipoPersona = 1
    fePersNegativas.EncabezadosAnchos = "500-0-1800-2000-1600-1600-1600-3000-2000-2000-2000-2000-2000-2000-3000-2000-1000-0-0-0-0"
    fePersNegativas.EncabezadosNombres = "#-Aux-Nº Documento-Nombre-Apellido Paterno-Apellido Materno-Apellido Casada-Delito-Juzgado-Oficio Múltiple-Oficio Juzgado-Expediente-Departamento-Tipo-Comentario-Condición-Existe-CodDep-CodTipo-CodCond-cMovNro-Institucion-Cargo" 'marg
End If
End Sub

Private Function ValidaDatos() As Boolean
If Trim(txtNomArchivo.Text) = "" Then
    MsgBox "Seleccione el Archivo a cargar", vbInformation, "Aviso"
    Exit Function
    ValidaDatos = False
End If

ValidaDatos = True
End Function

Private Function ValidaRegistros(ByVal pnTipo As Integer, ByVal prs As ADODB.Recordset, ByVal psNombre As String, ByRef psCodigo As String, ByRef psCadError As String, ByVal pbValida As Boolean) As Boolean
Dim bEncontro As Boolean
bEncontro = True
prs.MoveFirst
If psNombre <> "" Then
    Select Case pnTipo
        Case 1:
                If Not (prs.EOF And prs.BOF) Then
                    For i = 1 To prs.RecordCount
                        If Trim(UCase(prs!cUbiGeoDescripcion)) = Trim(UCase(psNombre)) Then
                            psCodigo = Trim(prs!cUbiGeoCod)
                            bEncontro = True
                            Exit For
                        Else
                            bEncontro = False
                        End If
                        prs.MoveNext
                    Next i
                End If
        Case 2, 3:
                If Not (prs.EOF And prs.BOF) Then
                     For i = 1 To prs.RecordCount
                        If Trim(UCase(prs!cConsDescripcion)) = Trim(UCase(psNombre)) Then
                            psCodigo = Trim(prs!nConsValor)
                            bEncontro = True
                            Exit For
                        Else
                            bEncontro = False
                        End If
                        prs.MoveNext
                    Next i
                End If
    End Select
End If

If Not bEncontro Then
    psCodigo = ""
    Select Case pnTipo
        Case 1: psCadError = psCadError + " - (Columna Departamento - Dato no valido)"
        Case 2: psCadError = psCadError + " - (Columna Tipo - Dato no valido)"
        Case 3: psCadError = psCadError + " - (Columna Condicion - Dato no valido)"
    End Select
End If

If Not pbValida Then
    If Not bEncontro Then
        ValidaRegistros = True
    Else
        ValidaRegistros = False
    End If
Else
    ValidaRegistros = True
End If

End Function

Private Sub ExportarExcelError(ByRef pmPersNegativos() As PersNegativa)
Dim xlsAplicacion As New Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlsHoja As Excel.Worksheet
Dim xlHoja1 As Excel.Worksheet
Dim ldFecha As Date
Dim lsArchivo As String
Dim lnExcel As Integer

On Error GoTo ErrExportarExcelError
    
lsArchivo = "\spooler\PersNegativasLoteErrores" & "_" & gsCodUser & "_" & gsCodAge & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"

Set xlsLibro = xlsAplicacion.Workbooks.Add
Set xlsHoja = xlsLibro.Worksheets.Add

xlsHoja.Name = "Errores"
xlsHoja.Cells.Font.Name = "Arial"
xlsHoja.Cells.Font.Size = 9
    
xlsHoja.Columns("A:A").ColumnWidth = 2
xlsHoja.Cells(3, 3) = IIf(fnTipoPersona = 1, "PERSONAS NATURALES", "PERSONAS JURIDICAS") & " - ERRORES"
xlsHoja.Cells(3, 3).Font.Size = 13
xlsHoja.Cells(3, 3).Font.Bold = 1
xlsHoja.Range("B3", IIf(fnTipoPersona = 1, "R3", "O3")).MergeCells = True
xlsHoja.Cells(3, 3).HorizontalAlignment = 3

lnExcel = 5

xlsHoja.Cells(lnExcel, 2) = "Nº Fila"
If fnTipoPersona = 1 Then
    xlsHoja.Cells(lnExcel, 3) = "Nº Documento"
    xlsHoja.Cells(lnExcel, 4) = "Nombres"
    xlsHoja.Cells(lnExcel, 5) = "Apellido Paterno"
    xlsHoja.Cells(lnExcel, 6) = "Apellido Materno"
    xlsHoja.Cells(lnExcel, 7) = "Apellido Casada"
    xlsHoja.Cells(lnExcel, 8) = "Delito"
    xlsHoja.Cells(lnExcel, 9) = "Juzgado"
    xlsHoja.Cells(lnExcel, 10) = "Oficio Múltiple"
    xlsHoja.Cells(lnExcel, 11) = "Oficio Juzgado"
    xlsHoja.Cells(lnExcel, 12) = "Expediente"
    xlsHoja.Cells(lnExcel, 13) = "Departamento"
    xlsHoja.Cells(lnExcel, 14) = "Tipo"
    xlsHoja.Cells(lnExcel, 15) = "Comentario"
    xlsHoja.Cells(lnExcel, 16) = "Condición"
    xlsHoja.Cells(lnExcel, 17) = "Institución" 'marg
    xlsHoja.Cells(lnExcel, 18) = "Cargo" 'marg
    xlsHoja.Cells(lnExcel, 19) = "Existe"
    xlsHoja.Cells(lnExcel, 20) = "Tipo Error"
Else
    xlsHoja.Cells(lnExcel, 3) = "Nº Documento"
    xlsHoja.Cells(lnExcel, 4) = "Razon Social"
    xlsHoja.Cells(lnExcel, 5) = "Delito"
    xlsHoja.Cells(lnExcel, 6) = "Juzgado"
    xlsHoja.Cells(lnExcel, 7) = "Oficio Múltiple"
    xlsHoja.Cells(lnExcel, 8) = "Oficio Juzgado"
    xlsHoja.Cells(lnExcel, 9) = "Expediente"
    xlsHoja.Cells(lnExcel, 10) = "Departamento"
    xlsHoja.Cells(lnExcel, 11) = "Tipo"
    xlsHoja.Cells(lnExcel, 12) = "Comentario"
    xlsHoja.Cells(lnExcel, 13) = "Condición"
    xlsHoja.Cells(lnExcel, 14) = "Institución" 'marg
    xlsHoja.Cells(lnExcel, 15) = "Cargo" 'marg
    xlsHoja.Cells(lnExcel, 16) = "Existe"
    xlsHoja.Cells(lnExcel, 17) = "Tipo Error"
End If
xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).Borders.LineStyle = 1
xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).Font.Bold = 1
xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).HorizontalAlignment = 3
xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).Interior.Color = RGB(200, 200, 200)

lnExcel = lnExcel + 1
For i = 1 To UBound(pmPersNegativos)
    If Trim(pmPersNegativos(i).CadError) <> "" Then
            xlsHoja.Cells(lnExcel, 2) = i
            xlsHoja.Cells(lnExcel, 2).HorizontalAlignment = 3
            
            If fnTipoPersona = 1 Then
                xlsHoja.Cells(lnExcel, 3).NumberFormat = "@" 'WIOR 20130911
                xlsHoja.Cells(lnExcel, 3) = pmPersNegativos(i).NumDoc
                xlsHoja.Cells(lnExcel, 4) = pmPersNegativos(i).NombreRazon
                xlsHoja.Cells(lnExcel, 5) = pmPersNegativos(i).ApePat
                xlsHoja.Cells(lnExcel, 6) = pmPersNegativos(i).ApeMat
                xlsHoja.Cells(lnExcel, 7) = pmPersNegativos(i).ApeCas
                xlsHoja.Cells(lnExcel, 8) = pmPersNegativos(i).Delito
                xlsHoja.Cells(lnExcel, 9) = pmPersNegativos(i).Juzgado
                xlsHoja.Cells(lnExcel, 10) = pmPersNegativos(i).OFMultiple
                xlsHoja.Cells(lnExcel, 11) = pmPersNegativos(i).OFJuzgado
                xlsHoja.Cells(lnExcel, 12) = pmPersNegativos(i).Expediente
                xlsHoja.Cells(lnExcel, 13) = pmPersNegativos(i).Departamento
                xlsHoja.Cells(lnExcel, 14) = pmPersNegativos(i).tipo
                xlsHoja.Cells(lnExcel, 15) = pmPersNegativos(i).Comentario
                xlsHoja.Cells(lnExcel, 16) = pmPersNegativos(i).Condicion
                xlsHoja.Cells(lnExcel, 17) = pmPersNegativos(i).Institucion 'marg
                xlsHoja.Cells(lnExcel, 18) = pmPersNegativos(i).Cargo 'marg
                xlsHoja.Cells(lnExcel, 19) = pmPersNegativos(i).existe
                xlsHoja.Cells(lnExcel, 20) = pmPersNegativos(i).CadError
                xlsHoja.Range(xlsHoja.Cells(lnExcel, 2), xlsHoja.Cells(lnExcel, 18)).Borders.LineStyle = 1
            Else
                xlsHoja.Cells(lnExcel, 3).NumberFormat = "@" 'WIOR 20130911
                xlsHoja.Cells(lnExcel, 3) = pmPersNegativos(i).NumDoc
                xlsHoja.Cells(lnExcel, 4) = pmPersNegativos(i).NombreRazon
                xlsHoja.Cells(lnExcel, 5) = pmPersNegativos(i).Delito
                xlsHoja.Cells(lnExcel, 6) = pmPersNegativos(i).Juzgado
                xlsHoja.Cells(lnExcel, 7) = pmPersNegativos(i).OFMultiple
                xlsHoja.Cells(lnExcel, 8) = pmPersNegativos(i).OFJuzgado
                xlsHoja.Cells(lnExcel, 9) = pmPersNegativos(i).Expediente
                xlsHoja.Cells(lnExcel, 10) = pmPersNegativos(i).Departamento
                xlsHoja.Cells(lnExcel, 11) = pmPersNegativos(i).tipo
                xlsHoja.Cells(lnExcel, 12) = pmPersNegativos(i).Comentario
                xlsHoja.Cells(lnExcel, 13) = pmPersNegativos(i).Condicion
                xlsHoja.Cells(lnExcel, 14) = pmPersNegativos(i).Institucion 'marg
                xlsHoja.Cells(lnExcel, 15) = pmPersNegativos(i).Cargo 'marg
                xlsHoja.Cells(lnExcel, 16) = pmPersNegativos(i).existe
                xlsHoja.Cells(lnExcel, 17) = pmPersNegativos(i).CadError
                xlsHoja.Range(xlsHoja.Cells(lnExcel, 2), xlsHoja.Cells(lnExcel, 15)).Borders.LineStyle = 1
            End If
        lnExcel = lnExcel + 1
    End If
Next i

xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).EntireColumn.AutoFit

For Each xlHoja1 In xlsLibro.Worksheets
    If UCase(xlHoja1.Name) = "HOJA1" Or UCase(xlHoja1.Name) = "HOJA2" Or UCase(xlHoja1.Name) = "HOJA3" Then
        xlHoja1.Delete
    End If
Next
    
xlsHoja.SaveAs App.Path & lsArchivo
xlsAplicacion.Visible = True
xlsAplicacion.Windows(1).Visible = True
    
MsgBox "Archivos de Errores Generados", vbInformation, "Aviso"
    
Set xlsAplicacion = Nothing
Set xlsLibro = Nothing
Set xlsHoja = Nothing
    

Exit Sub
ErrExportarExcelError:
    MsgBox Err.Description, vbCritical, "Aviso"
    Exit Sub
End Sub

Private Sub ExportarExcelArchivoProcesado(ByRef pmPersNegativos() As PersNegativa)
Dim xlsAplicacion As New Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlsHoja As Excel.Worksheet
Dim xlHoja1 As Excel.Worksheet
Dim ldFecha As Date
Dim lsArchivo As String
Dim lnExcel As Integer
Dim oPersonas As COMDPersona.DCOMPersonas

On Error GoTo ErrExportarExcelError
    
lsArchivo = "\spooler\PersNegativasLoteFinal" & "_" & gsCodUser & "_" & gsCodAge & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"

Set xlsLibro = xlsAplicacion.Workbooks.Add
Set xlsHoja = xlsLibro.Worksheets.Add

xlsHoja.Name = "PersonasRegistradas"
xlsHoja.Cells.Font.Name = "Arial"
xlsHoja.Cells.Font.Size = 9
    
xlsHoja.Columns("A:A").ColumnWidth = 2
xlsHoja.Cells(3, 3) = IIf(fnTipoPersona = 1, "PERSONAS NATURALES", "PERSONAS JURIDICAS") & " - REGISTRADAS"
xlsHoja.Cells(3, 3).Font.Size = 13
xlsHoja.Cells(3, 3).Font.Bold = 1
xlsHoja.Range("B3", IIf(fnTipoPersona = 1, "R3", "O3")).MergeCells = True
xlsHoja.Cells(3, 3).HorizontalAlignment = 3

lnExcel = 5

xlsHoja.Cells(lnExcel, 2) = "Nº Fila"
If fnTipoPersona = 1 Then
    xlsHoja.Cells(lnExcel, 3) = "Nº Documento"
    xlsHoja.Cells(lnExcel, 4) = "Nombres"
    xlsHoja.Cells(lnExcel, 5) = "Apellido Paterno"
    xlsHoja.Cells(lnExcel, 6) = "Apellido Materno"
    xlsHoja.Cells(lnExcel, 7) = "Apellido Casada"
    xlsHoja.Cells(lnExcel, 8) = "Delito"
    xlsHoja.Cells(lnExcel, 9) = "Juzgado"
    xlsHoja.Cells(lnExcel, 10) = "Oficio Múltiple"
    xlsHoja.Cells(lnExcel, 11) = "Oficio Juzgado"
    xlsHoja.Cells(lnExcel, 12) = "Expediente"
    xlsHoja.Cells(lnExcel, 13) = "Departamento"
    xlsHoja.Cells(lnExcel, 14) = "Tipo"
    xlsHoja.Cells(lnExcel, 15) = "Comentario"
    xlsHoja.Cells(lnExcel, 16) = "Condición" 'marg
    xlsHoja.Cells(lnExcel, 17) = "Institución" 'marg
    xlsHoja.Cells(lnExcel, 18) = "Cargo"
    xlsHoja.Cells(lnExcel, 19) = "Existe"
    xlsHoja.Cells(lnExcel, 20) = "Cliente"
Else
    xlsHoja.Cells(lnExcel, 3) = "Nº Documento"
    xlsHoja.Cells(lnExcel, 4) = "Razon Social"
    xlsHoja.Cells(lnExcel, 5) = "Delito"
    xlsHoja.Cells(lnExcel, 6) = "Juzgado"
    xlsHoja.Cells(lnExcel, 7) = "Oficio Múltiple"
    xlsHoja.Cells(lnExcel, 8) = "Oficio Juzgado"
    xlsHoja.Cells(lnExcel, 9) = "Expediente"
    xlsHoja.Cells(lnExcel, 10) = "Departamento"
    xlsHoja.Cells(lnExcel, 11) = "Tipo"
    xlsHoja.Cells(lnExcel, 12) = "Comentario"
    xlsHoja.Cells(lnExcel, 13) = "Condición" 'marg
    xlsHoja.Cells(lnExcel, 14) = "Institución" 'marg
    xlsHoja.Cells(lnExcel, 15) = "Cargo"
    xlsHoja.Cells(lnExcel, 16) = "Existe"
    xlsHoja.Cells(lnExcel, 17) = "Cliente"
End If
xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).Borders.LineStyle = 1
xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).Font.Bold = 1
xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).HorizontalAlignment = 3
xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).Interior.Color = RGB(200, 200, 200)

lnExcel = lnExcel + 1
For i = 1 To UBound(pmPersNegativos)
    Set oPersonas = New COMDPersona.DCOMPersonas
    pmPersNegativos(i).CadError = "NO"
    
    If Trim(pmPersNegativos(i).NumDoc) <> "" Then
        pmPersNegativos(i).CadError = oPersonas.PersNegativaExiste(pmPersNegativos(i).NumDoc, , , , , 2, fnTipoPersona)
    End If
    
    If fnTipoPersona = 1 Then
        If pmPersNegativos(i).CadError = "NO" Then
            pmPersNegativos(i).CadError = oPersonas.PersNegativaExiste(, pmPersNegativos(i).NombreRazon, pmPersNegativos(i).ApePat, pmPersNegativos(i).ApeMat, pmPersNegativos(i).ApeCas, 1, fnTipoPersona)
        End If
    ElseIf fnTipoPersona = 2 Then
        If pmPersNegativos(i).CadError = "NO" Then
            pmPersNegativos(i).CadError = oPersonas.PersNegativaExiste(, pmPersNegativos(i).NombreRazon, , , , 1, fnTipoPersona)
        End If
    End If

    xlsHoja.Cells(lnExcel, 2) = i
    xlsHoja.Cells(lnExcel, 2).HorizontalAlignment = 3
    
    xlsHoja.Cells(lnExcel, 3).NumberFormat = "@" 'WIOR 20130911
    xlsHoja.Cells(lnExcel, 3) = pmPersNegativos(i).NumDoc
    xlsHoja.Cells(lnExcel, 4) = pmPersNegativos(i).NombreRazon

    If fnTipoPersona = 1 Then
        xlsHoja.Cells(lnExcel, 5) = pmPersNegativos(i).ApePat
        xlsHoja.Cells(lnExcel, 6) = pmPersNegativos(i).ApeMat
        xlsHoja.Cells(lnExcel, 7) = pmPersNegativos(i).ApeCas
        xlsHoja.Cells(lnExcel, 8) = pmPersNegativos(i).Delito
        xlsHoja.Cells(lnExcel, 9) = pmPersNegativos(i).Juzgado
        xlsHoja.Cells(lnExcel, 10) = pmPersNegativos(i).OFMultiple
        xlsHoja.Cells(lnExcel, 11) = pmPersNegativos(i).OFJuzgado
        xlsHoja.Cells(lnExcel, 12) = pmPersNegativos(i).Expediente
        xlsHoja.Cells(lnExcel, 13) = pmPersNegativos(i).Departamento
        xlsHoja.Cells(lnExcel, 14) = pmPersNegativos(i).tipo
        xlsHoja.Cells(lnExcel, 15) = pmPersNegativos(i).Comentario
        xlsHoja.Cells(lnExcel, 16) = pmPersNegativos(i).Condicion
        xlsHoja.Cells(lnExcel, 17) = pmPersNegativos(i).Institucion 'marg
        xlsHoja.Cells(lnExcel, 18) = pmPersNegativos(i).Cargo 'marg
        xlsHoja.Cells(lnExcel, 19) = pmPersNegativos(i).existe
        xlsHoja.Cells(lnExcel, 20) = pmPersNegativos(i).CadError
        xlsHoja.Range(xlsHoja.Cells(lnExcel, 2), xlsHoja.Cells(lnExcel, 18)).Borders.LineStyle = 1
    Else
        xlsHoja.Cells(lnExcel, 5) = pmPersNegativos(i).Delito
        xlsHoja.Cells(lnExcel, 6) = pmPersNegativos(i).Juzgado
        xlsHoja.Cells(lnExcel, 7) = pmPersNegativos(i).OFMultiple
        xlsHoja.Cells(lnExcel, 8) = pmPersNegativos(i).OFJuzgado
        xlsHoja.Cells(lnExcel, 9) = pmPersNegativos(i).Expediente
        xlsHoja.Cells(lnExcel, 10) = pmPersNegativos(i).Departamento
        xlsHoja.Cells(lnExcel, 11) = pmPersNegativos(i).tipo
        xlsHoja.Cells(lnExcel, 12) = pmPersNegativos(i).Comentario
        xlsHoja.Cells(lnExcel, 13) = pmPersNegativos(i).Condicion
        xlsHoja.Cells(lnExcel, 14) = pmPersNegativos(i).Institucion
        xlsHoja.Cells(lnExcel, 15) = pmPersNegativos(i).Cargo
        xlsHoja.Cells(lnExcel, 16) = pmPersNegativos(i).existe
        xlsHoja.Cells(lnExcel, 17) = pmPersNegativos(i).CadError
        xlsHoja.Range(xlsHoja.Cells(lnExcel, 2), xlsHoja.Cells(lnExcel, 15)).Borders.LineStyle = 1
    End If
    lnExcel = lnExcel + 1
    Set oPersonas = Nothing
Next i

xlsHoja.Range("B5", IIf(fnTipoPersona = 1, "R5", "O5")).EntireColumn.AutoFit

For Each xlHoja1 In xlsLibro.Worksheets
    If UCase(xlHoja1.Name) = "HOJA1" Or UCase(xlHoja1.Name) = "HOJA2" Or UCase(xlHoja1.Name) = "HOJA3" Then
        xlHoja1.Delete
    End If
Next
    
xlsHoja.SaveAs App.Path & lsArchivo
xlsAplicacion.Visible = True
xlsAplicacion.Windows(1).Visible = True
    
Set xlsAplicacion = Nothing
Set xlsLibro = Nothing
Set xlsHoja = Nothing
    

Exit Sub
ErrExportarExcelError:
    MsgBox Err.Description, vbCritical, "Aviso"
    Exit Sub
End Sub






