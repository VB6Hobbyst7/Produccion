VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapGeneraArchivoRecaudo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicio de Recaudo - Generación de Retorno"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11085
   Icon            =   "frmCapGeneraArchivoRecaudo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   8505
      Top             =   3885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6120
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   10795
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Selección del Convenio"
      TabPicture(0)   =   "frmCapGeneraArchivoRecaudo.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtTotalRegistros"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboTipoArchivo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdGenerar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCerrar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "grdDetallePago"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin SICMACT.FlexEdit grdDetallePago 
         Height          =   3015
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   5318
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-ID-CODIGO-T.DOI-DOI-NOMBRE-CONCEPTO-IMPORTE-FECHA-NumeroCobro"
         EncabezadosAnchos=   "0-500-800-500-1000-3000-3000-1200-1200-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   9450
         TabIndex        =   12
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
         Height          =   375
         Left            =   8190
         TabIndex        =   11
         Top             =   5520
         Width           =   1095
      End
      Begin VB.ComboBox cboTipoArchivo 
         Height          =   315
         ItemData        =   "frmCapGeneraArchivoRecaudo.frx":0326
         Left            =   6825
         List            =   "frmCapGeneraArchivoRecaudo.frx":0333
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   5570
         Width           =   1095
      End
      Begin VB.TextBox txtTotalRegistros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   5520
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Búsqueda de convenio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   10410
         Begin VB.CommandButton cmdBuscarConvenio 
            Caption         =   "..."
            Height          =   375
            Left            =   3255
            TabIndex        =   0
            Top             =   315
            Width           =   400
         End
         Begin VB.TextBox txtCodigoConvenio 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   20
            Top             =   315
            Width           =   2175
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   375
            Left            =   8610
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker dHasta 
            Height          =   375
            Left            =   7035
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   75628545
            CurrentDate     =   41324
         End
         Begin MSComCtl2.DTPicker dDesde 
            Height          =   375
            Left            =   4830
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   75628545
            CurrentDate     =   41324
         End
         Begin VB.TextBox txtCodigoEmpresa 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox txtNombreEmpresa 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3330
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1320
            Width           =   3765
         End
         Begin VB.TextBox txtDescripcionConvenio 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   840
            Width           =   6015
         End
         Begin VB.Label lblHasta 
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6405
            TabIndex        =   3
            Top             =   420
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblDesde 
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4095
            TabIndex        =   1
            Top             =   420
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Código: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   400
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Empresa: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Convenio: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Total Registros: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5250
         TabIndex        =   19
         Top             =   5625
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "Total Registros: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5570
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCapGeneraArchivoRecaudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sCliente As String
Private sCodConvenio As String
Private sTipoValidacion As String
Private nRegistros As Double
Private dFechaGeneracion As String
Private sRucEmpresa As String
Dim oBuscaPersonas As COMDPersona.DCOMPersonas
Dim rsEmpresa As ADODB.Recordset

Private Sub CmdBuscar_Click()
    
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Dim rsRecaudoDetalle As Recordset
    Dim cTipoConvenio As String
    Dim objMovimiento As COMNContabilidad.NCOMContFunciones
    Dim nContadorRegistros As Integer
    
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    Set objMovimiento = New COMNContabilidad.NCOMContFunciones
    
    Set rsRecaudoDetalle = ClsServicioRecaudo.getListaPagosRecaudoSV(Trim(txtCodigoConvenio.Text), _
                                                                    dDesde.value, dHasta.value)

    If Not (rsRecaudoDetalle.BOF Or rsRecaudoDetalle.EOF) Then

        If grdDetallePago.Rows > 2 Then
              
           grdDetallePago.Clear
           grdDetallePago.Rows = 2
           grdDetallePago.FormaCabecera
              
        End If

        sCliente = rsRecaudoDetalle!cNomCliente
        sCodConvenio = txtCodigoConvenio.Text
        cTipoConvenio = Mid(sCodConvenio, 14, 2)
        sTipoValidacion = IIf(cTipoConvenio = "VC", "VALIDACION COMPLETA", _
                          IIf(cTipoConvenio = "VI", "VALIDACION INCOMPLETA", _
                          IIf(cTipoConvenio = "SV", "SIN VALIDACION", "VALIDACION POR IMPORTES")))
        nRegistros = CDbl(rsRecaudoDetalle.RecordCount)
        dFechaGeneracion = Format$(gdFecSis, "yyyy/mm/dd")
        Set oBuscaPersonas = New COMDPersona.DCOMPersonas
        Set rsEmpresa = oBuscaPersonas.BuscaCliente(txtCodigoEmpresa.Text, BusquedaCodigo)
        sRucEmpresa = rsEmpresa!cPersIDnroRUC
        
        Set oBuscaPersonas = Nothing
        Set rsEmpresa = Nothing
        
    Else
        If grdDetallePago.Rows > 2 Then
        
           grdDetallePago.Clear
           grdDetallePago.Rows = 2
           grdDetallePago.FormaCabecera
              
        End If
        
        Exit Sub
    
    End If
    
    Do While Not (rsRecaudoDetalle.BOF Or rsRecaudoDetalle.EOF)
        grdDetallePago.AdicionaFila , , True
        
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 1) = rsRecaudoDetalle!cId
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 2) = rsRecaudoDetalle!cCodCliente
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 3) = IIf(rsRecaudoDetalle!nTipoDOI = 1, "DNI", "RUC")
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 4) = rsRecaudoDetalle!cDOI
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 5) = rsRecaudoDetalle!cNomCliente
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 6) = rsRecaudoDetalle!cConcepto
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 7) = rsRecaudoDetalle!nImporte
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 8) = objMovimiento.ObtieneFechaMov(rsRecaudoDetalle!cMovNro, True)
        grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 9) = rsRecaudoDetalle!nNumeroCobro
        
        rsRecaudoDetalle.MoveNext
        
        nContadorRegistros = nContadorRegistros + 1
    Loop
    txtTotalRegistros.Text = nContadorRegistros

End Sub

Private Sub cmdBuscarConvenio_Click()
  Dim rsRecaudo As Recordset
    Set rsRecaudo = New Recordset
    Set rsRecaudo = frmBuscarConvenio.Inicio
    Dim rsRecaudoDetalle As Recordset
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Dim objMovimiento As COMNContabilidad.NCOMContFunciones
    Dim nContadorRegistros As Integer
    Dim cTipoConvenio As String
    
    If Not rsRecaudo Is Nothing Then
        
        If Not (rsRecaudo.EOF And rsRecaudo.BOF) Then
            
            txtCodigoConvenio.Text = rsRecaudo!cCodConvenio
            txtDescripcionConvenio.Text = rsRecaudo!cNombreConvenio
            txtCodigoEmpresa.Text = rsRecaudo!cPersCod
            txtNombreEmpresa.Text = rsRecaudo!cPersNombre
            txtCodigoConvenio.Locked = True
            
            If Mid(rsRecaudo!cCodConvenio, 14, 2) = "SV" Then
                               
                If grdDetallePago.Rows > 2 Then
                    
                    grdDetallePago.Clear
                    grdDetallePago.Rows = 2
                    grdDetallePago.FormaCabecera
                    
                End If
                
                dDesde.Visible = True
                dHasta.Visible = True
                lblDesde.Visible = True
                lblHasta.Visible = True
                cmdBuscar.Visible = True
                
                Exit Sub
            
            End If
            
            dDesde.Visible = False
            dHasta.Visible = False
            lblDesde.Visible = False
            lblHasta.Visible = False
            cmdBuscar.Visible = False
                
            'Set rsRecaudoDetalle = New Recordset
            Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
            Set objMovimiento = New COMNContabilidad.NCOMContFunciones
            
            Set rsRecaudoDetalle = ClsServicioRecaudo.getListaPagosRecaudo(Mid(Trim(txtCodigoConvenio.Text), 14, 2), _
                                                                                   Trim(txtCodigoConvenio.Text))
            
            grdDetallePago.Clear
            grdDetallePago.Rows = 2
            grdDetallePago.FormaCabecera
            
            If Not (rsRecaudoDetalle.BOF Or rsRecaudoDetalle.EOF) Then

                sCliente = rsRecaudoDetalle!cNomCliente
                sCodConvenio = txtCodigoConvenio.Text
                cTipoConvenio = Mid(sCodConvenio, 14, 2)
                sTipoValidacion = IIf(cTipoConvenio = "VC", "VALIDACION COMPLETA", _
                                  IIf(cTipoConvenio = "VI", "VALIDACION INCOMPLETA", _
                                  IIf(cTipoConvenio = "SV", "SIN VALIDACION", "VALIDACION POR IMPORTES")))
                nRegistros = CDbl(rsRecaudoDetalle.RecordCount)
                dFechaGeneracion = Format$(gdFecSis, "yyyy/mm/dd")
                Set oBuscaPersonas = New COMDPersona.DCOMPersonas
                Set rsEmpresa = oBuscaPersonas.BuscaCliente(txtCodigoEmpresa.Text, BusquedaCodigo)
                sRucEmpresa = rsEmpresa!cPersIDnroRUC
                
                Set oBuscaPersonas = Nothing
                Set rsEmpresa = Nothing
                
            Else
                Exit Sub
            
            End If
            
            Do While Not (rsRecaudoDetalle.BOF Or rsRecaudoDetalle.EOF)
                grdDetallePago.AdicionaFila , , True
                
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 1) = rsRecaudoDetalle!cId
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 2) = rsRecaudoDetalle!cCodCliente
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 3) = IIf(rsRecaudoDetalle!nTipoDOI = 1, "DNI", "RUC")
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 4) = rsRecaudoDetalle!cDOI
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 5) = rsRecaudoDetalle!cNomCliente
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 6) = rsRecaudoDetalle!cConcepto
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 7) = Format(rsRecaudoDetalle!nImporte, "#,##0.00")
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 8) = objMovimiento.ObtieneFechaMov(rsRecaudoDetalle!cMovNro, True)
                grdDetallePago.TextMatrix(grdDetallePago.Rows - 1, 9) = rsRecaudoDetalle!nNumeroCobro
                
                rsRecaudoDetalle.MoveNext
                
                nContadorRegistros = nContadorRegistros + 1
            Loop
            
            txtTotalRegistros.Text = nContadorRegistros
        End If
    Else
         MsgBox "Usted no selecciono ninguna Empresa", vbInformation, "Aviso"
    
    End If
End Sub

Private Sub cmdGenerar_Click()
    
    If validarFormulario = False Then
        MsgBox "Completar todos los datos del formulario", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Dim oConecta As COMConecta.DCOMConecta
    
    Set oConecta = New COMConecta.DCOMConecta
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    
    dlgArchivo.Filename = Empty
    dlgArchivo.Filter = "Archivo *." & Trim(cboTipoArchivo.Text) & "|." & IIf(Trim(cboTipoArchivo.Text) = "txt", "txt", _
                                                                          IIf(Trim(cboTipoArchivo.Text) = "xls", "xls", _
                                                                       "xlsx"))
    
    'dlgArchivo.FileTitle = Trim(txtCodigoConvenio.Text) & Year(gdFecSis) & Month(gdFecSis) & Day(gdFecSis)
    dlgArchivo.Filename = Trim(txtCodigoConvenio.Text) & Year(gdFecSis) & Month(gdFecSis) & Day(gdFecSis) & "." & cboTipoArchivo.Text
    
    On Error GoTo error_handler
    
    dlgArchivo.ShowSave
                
    If dlgArchivo.FileTitle = "" Then Exit Sub
    If dlgArchivo.Filename = "" Then Exit Sub
    
    If cboTipoArchivo.Text = "txt" Then
    
            Dim oSys As Scripting.FileSystemObject
            Dim oText As Scripting.TextStream
                
            Set oSys = New Scripting.FileSystemObject
            
            dlgArchivo.Filename = Replace(dlgArchivo.Filename, dlgArchivo.FileTitle, _
                Trim(txtCodigoConvenio.Text) & Year(gdFecSis) & Month(gdFecSis) & Day(gdFecSis)) & "." & cboTipoArchivo.Text
                
            If oSys.FileExists(dlgArchivo.Filename) = False Then
                oSys.CreateTextFile (dlgArchivo.Filename)
            End If
            
            'Ahora oText tiene el fichero abierto con permisos de lectura.
            Set oText = oSys.OpenTextFile(dlgArchivo.Filename, ForReading)
        
            'En lugar de lo de arriba puedo cambiar el último parametro para modificarlo.
            
            Set oText = oSys.OpenTextFile(dlgArchivo.Filename, ForWriting)
        
            'y ahora por ejemplo si haces esto...
            Dim nI As Integer
            Dim cTipoConv As String
            
            cTipoConv = Mid(Trim(txtCodigoConvenio.Text), 14, 2)
            oText.WriteLine Trim(txtCodigoConvenio.Text) & "|" & IIf(cTipoConv = "VC", "C", IIf(cTipoConv = "VI", "I", _
                                                                 IIf(cTipoConv = "SV", "S", "P"))) & "|" & _
                                                                 Trim(txtTotalRegistros.Text) & "|" & _
                                                                 Year(gdFecSis) & Month(gdFecSis) & Day(gdFecSis)
                
            oConecta.AbreConexion
            oConecta.BeginTrans
    
            For nI = 1 To grdDetallePago.Rows - 1
                oText.WriteLine getCadenaTipoConvenio(cTipoConv, nI)
                
                If Not (ClsServicioRecaudo.actualizaXEnvioReporteCobros(grdDetallePago.TextMatrix(nI, 9), _
                                                                         gdFecSis, oConecta)) Then
                    oConecta.RollbackTrans
                    MsgBox "No se pudo terminar el proceso con exito", vbInformation, "Aviso"
                    Exit For
                End If
            Next
            
            'oText.WriteLine "esta cadena la estoy agregando al fichero"
            'ya has grabado datos en el, falta cerrarlo.
            
            oText.Close
            oConecta.CommitTrans
            oConecta.CierraConexion
            MsgBox "Se culminó el proceso con exito", vbInformation, "Aviso"
            limpiarFormulario
    
    ElseIf cboTipoArchivo.Text = "xls" Or cboTipoArchivo.Text = "xlsx" Then
            
            Dim fs As Scripting.FileSystemObject
            Dim xlsAplicacion As Excel.Application
            Dim lsArchivo As String
            Dim lsArchivo1 As String
            Dim lsNomHoja As String
            Dim xlsLibro As Excel.Workbook
            Dim xlHoja1 As Excel.Worksheet
            Dim lbExisteHoja As Boolean
            Dim nFila, i As Double
            
            lsArchivo = "FormatoArchivoRetorno"
            lsNomHoja = "FormatoArchivoRetorno"
            nFila = 20
            
            Set fs = New Scripting.FileSystemObject
            Set xlsAplicacion = New Excel.Application
                   
            lsArchivo1 = dlgArchivo.Filename

            If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
                Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
            Else
                MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
                Exit Sub
            End If
            
            For Each xlHoja1 In xlsLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                    Exit For
                End If
            Next
            
            xlHoja1.Cells(7, 5) = txtNombreEmpresa.Text ' Cliente
            xlHoja1.Cells(9, 5) = txtCodigoConvenio.Text ' Cod Convenio
            xlHoja1.Cells(11, 5) = sTipoValidacion ' Tipo Validacion
            xlHoja1.Cells(13, 5) = nRegistros ' Nro Registros
            xlHoja1.Cells(15, 5) = dFechaGeneracion ' Fecha Generacion
            xlHoja1.Cells(7, 10) = sRucEmpresa ' RUC
            
            Dim cTipoC As String
            cTipoC = Mid(Trim(txtCodigoConvenio.Text), 14, 2)
            
            oConecta.AbreConexion
            oConecta.BeginTrans
            For nI = 1 To grdDetallePago.Rows - 1
               nFila = nFila + 1
               
               Select Case cTipoC
               
                    Case "VC"
                        xlHoja1.Cells(nFila, 4) = Format(grdDetallePago.TextMatrix(nI, 8), "yyyy/mm/dd") 'FECHA
                        xlHoja1.Cells(nFila, 5) = grdDetallePago.TextMatrix(nI, 1)
                    
                    Case "VI"
                        xlHoja1.Cells(nFila, 4) = Format(grdDetallePago.TextMatrix(nI, 8), "yyyy/mm/dd") 'FECHA
                        xlHoja1.Cells(nFila, 5) = grdDetallePago.TextMatrix(nI, 1)
                        xlHoja1.Cells(nFila, 11) = grdDetallePago.TextMatrix(nI, 7)
                    
                    Case "SV"
                        xlHoja1.Cells(nFila, 4) = Format(grdDetallePago.TextMatrix(nI, 8), "yyyy/mm/dd") 'FECHA
                        xlHoja1.Cells(nFila, 6) = grdDetallePago.TextMatrix(nI, 2) ' CODIGO
                        xlHoja1.Cells(nFila, 7) = IIf(grdDetallePago.TextMatrix(nI, 3) = "DNI", 1, 2) ' TIPODOI *
                        xlHoja1.Cells(nFila, 8) = grdDetallePago.TextMatrix(nI, 4) ' DOI
                        xlHoja1.Cells(nFila, 9) = grdDetallePago.TextMatrix(nI, 5) ' NOMBRE
                        xlHoja1.Cells(nFila, 10) = grdDetallePago.TextMatrix(nI, 6) ' SERV/CONCEPTO
                        xlHoja1.Cells(nFila, 11) = grdDetallePago.TextMatrix(nI, 7) ' IMPORTE
                        
                    Case "VP"
                        xlHoja1.Cells(nFila, 4) = Format(grdDetallePago.TextMatrix(nI, 8), "yyyy/mm/dd") 'FECHA *
                        xlHoja1.Cells(nFila, 5) = grdDetallePago.TextMatrix(nI, 1) ' ID *
                        xlHoja1.Cells(nFila, 7) = IIf(grdDetallePago.TextMatrix(nI, 3) = "DNI", 1, 2) ' TIPODOI *
                        xlHoja1.Cells(nFila, 8) = grdDetallePago.TextMatrix(nI, 4) ' DOI *
                        xlHoja1.Cells(nFila, 9) = grdDetallePago.TextMatrix(nI, 5) ' NOMBRE *
                        xlHoja1.Cells(nFila, 10) = grdDetallePago.TextMatrix(nI, 6) ' SERV/CONCEPTO *
                        
               End Select
               
                If Not (ClsServicioRecaudo.actualizaXEnvioReporteCobros(grdDetallePago.TextMatrix(nI, 9), _
                                                                         gdFecSis, oConecta)) Then
                    oConecta.RollbackTrans
                    MsgBox "No se pudo terminar el proceso con exito", vbInformation, "Aviso"
                    Exit For
                End If
               
            Next
            
            xlHoja1.SaveAs lsArchivo1
            xlsAplicacion.Visible = True
            xlsAplicacion.Windows(1).Visible = True
            oConecta.CommitTrans
            oConecta.CierraConexion
            Set xlsAplicacion = Nothing
            Set xlsLibro = Nothing
            Set xlHoja1 = Nothing
            MsgBox "Se culminó el proceso con exito", vbInformation, "Aviso"
            limpiarFormulario
            
    End If
  
Exit Sub
    
error_handler:
    
    If err.Number = 32755 Then
        MsgBox "Se ha cancelado formulario", vbInformation, "Aviso"
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        oConecta.RollbackTrans
        oConecta.CierraConexion
        Set oConecta = Nothing
        Set ClsServicioRecaudo = Nothing
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Error al momento de generar archivo de recaudo", vbCritical, "Aviso"
    End If
    
End Sub

Private Function getCadenaTipoConvenio(ByVal cTipo As String, ByVal nFila As Integer) As String
    
    Dim cFecha As String
    
    cFecha = grdDetallePago.TextMatrix(nFila, 8)
    
    If cTipo = "VC" Then
        getCadenaTipoConvenio = grdDetallePago.TextMatrix(nFila, 1) & "|" & _
                                Year(cFecha) & Month(cFecha) & Day(cFecha)
    ElseIf cTipo = "VI" Then
        getCadenaTipoConvenio = grdDetallePago.TextMatrix(nFila, 1) & "|" & _
                                grdDetallePago.TextMatrix(nFila, 7) & "|" & _
                                Year(cFecha) & Month(cFecha) & Day(cFecha)
    ElseIf cTipo = "SV" Then
        getCadenaTipoConvenio = grdDetallePago.TextMatrix(nFila, 2) & "|" & _
                                 IIf(grdDetallePago.TextMatrix(nFila, 3) = "DNI", 1, 2) & "|" & _
                                 grdDetallePago.TextMatrix(nFila, 4) & "|" & _
                                 grdDetallePago.TextMatrix(nFila, 5) & "|" & _
                                 grdDetallePago.TextMatrix(nFila, 6) & "|" & _
                                 grdDetallePago.TextMatrix(nFila, 7) & "|" & _
                                 Year(cFecha) & Month(cFecha) & Day(cFecha)
    ElseIf cTipo = "VP" Then
            getCadenaTipoConvenio = grdDetallePago.TextMatrix(nFila, 1) & "|" & _
                                 IIf(grdDetallePago.TextMatrix(nFila, 3) = "DNI", 1, 2) & "|" & _
                                 grdDetallePago.TextMatrix(nFila, 4) & "|" & _
                                 grdDetallePago.TextMatrix(nFila, 5) & "|" & _
                                 grdDetallePago.TextMatrix(nFila, 6) & "|" & _
                                 Year(cFecha) & Month(cFecha) & Day(cFecha)
    End If
    
End Function

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
cboTipoArchivo.ListIndex = 0
dDesde.value = CDate(gdFecSis)
dHasta.value = CDate(gdFecSis)
End Sub

Private Sub limpiarFormulario()
    
    txtCodigoConvenio.Text = ""
    txtDescripcionConvenio.Text = ""
    txtCodigoEmpresa.Text = ""
    txtNombreEmpresa.Text = ""
    txtTotalRegistros.Text = ""
    grdDetallePago.Clear
    grdDetallePago.FormaCabecera
    grdDetallePago.Rows = 2
    txtCodigoConvenio.SetFocus
    
End Sub

Private Function validarFormulario() As Boolean
    
    If txtCodigoConvenio.Text = "" Then
        validarFormulario = False
        Exit Function
    ElseIf txtDescripcionConvenio = "" Then
        validarFormulario = False
        Exit Function
    ElseIf txtCodigoEmpresa.Text = "" Then
        validarFormulario = False
        Exit Function
    ElseIf txtNombreEmpresa.Text = "" Then
        validarFormulario = False
        Exit Function
    ElseIf grdDetallePago.Rows <= 2 Then
        If grdDetallePago.Rows = 2 Then
            If grdDetallePago.TextMatrix(1, 0) = "" Then
                validarFormulario = False
                Exit Function
            End If
        End If
      
    End If
    validarFormulario = True
    
End Function



Private Sub txtCodigoConvenio_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub
