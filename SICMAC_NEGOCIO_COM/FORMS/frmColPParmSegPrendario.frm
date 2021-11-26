VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmColPParmSegPrendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de Segmentación Prendario"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "frmColPParmSegPrendario.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7858
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cliente Recurrente"
      TabPicture(0)   =   "frmColPParmSegPrendario.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(2)=   "feClientesRecurrentes"
      Tab(0).Control(3)=   "cmdEditarCR"
      Tab(0).Control(4)=   "cmdGuardarCR"
      Tab(0).Control(5)=   "cmdCancelarCR"
      Tab(0).Control(6)=   "cmdAgregaCR"
      Tab(0).Control(7)=   "cmdQuitarCR"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Cliente Nuevo"
      TabPicture(1)   =   "frmColPParmSegPrendario.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdQuitarCN"
      Tab(1).Control(1)=   "cmdAgregaCN"
      Tab(1).Control(2)=   "cmdCancelarCN"
      Tab(1).Control(3)=   "cmdGuardarCN"
      Tab(1).Control(4)=   "cmdEditarCN"
      Tab(1).Control(5)=   "feClienteNuevo"
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(7)=   "Label1"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Monto Préstamo"
      TabPicture(2)   =   "frmColPParmSegPrendario.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "feSegExtMontoPrestamo"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "feMontoPrestamo"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdEditarMP"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdGuardarMP"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdCancelarMP"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdAgregaMP"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdQuitarMP"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.CommandButton cmdQuitarMP 
         Caption         =   "-"
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
         Left            =   4440
         TabIndex        =   22
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdAgregaMP 
         Caption         =   "+"
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
         Left            =   3840
         TabIndex        =   21
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdQuitarCN 
         Caption         =   "-"
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
         Left            =   -66840
         TabIndex        =   20
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdAgregaCN 
         Caption         =   "+"
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
         Left            =   -67440
         TabIndex        =   19
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdQuitarCR 
         Caption         =   "-"
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
         Left            =   -74280
         TabIndex        =   18
         Top             =   4005
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdAgregaCR 
         Caption         =   "+"
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
         Left            =   -74880
         TabIndex        =   17
         Top             =   4005
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdCancelarMP 
         Caption         =   "Cancelar"
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
         Left            =   9240
         TabIndex        =   9
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardarMP 
         Caption         =   "Guardar"
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
         Left            =   8160
         TabIndex        =   8
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdEditarMP 
         Caption         =   "Editar"
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
         TabIndex        =   7
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarCN 
         Caption         =   "Cancelar"
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
         Left            =   -67440
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardarCN 
         Caption         =   "Guardar"
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
         Left            =   -67440
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdEditarCN 
         Caption         =   "Editar"
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
         Left            =   -67440
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarCR 
         Caption         =   "Cancelar"
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
         Left            =   -65640
         TabIndex        =   3
         Top             =   4005
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardarCR 
         Caption         =   "Guardar"
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
         Left            =   -66645
         TabIndex        =   2
         Top             =   4005
         Width           =   975
      End
      Begin VB.CommandButton cmdEditarCR 
         Caption         =   "Editar"
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
         Left            =   -67650
         TabIndex        =   1
         Top             =   4005
         Width           =   975
      End
      Begin SICMACT.FlexEdit feClienteNuevo 
         Height          =   3300
         Left            =   -74040
         TabIndex        =   10
         Top             =   840
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   5821
         Cols0           =   7
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Segmentación-Sub Segmentación-Monto Tasa Desde-Monto Tasa Hasta-nTpCliente-nId"
         EncabezadosAnchos=   "400-1500-1500-1500-1500-0-0"
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
         ColumnasAEditar =   "X-1-2-3-4-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-R-R-R-C"
         FormatosEdit    =   "0-1-0-2-2-3-0"
         CantEntero      =   10
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit feMontoPrestamo 
         Height          =   3180
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   6244
         Cols0           =   4
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Sub Segmentación-Monto %-nId"
         EncabezadosAnchos=   "400-1500-1200-0"
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
         ColumnasAEditar =   "X-1-2-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R"
         FormatosEdit    =   "0-0-2-3"
         CantEntero      =   10
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit feClientesRecurrentes 
         Height          =   3120
         Left            =   -74925
         TabIndex        =   16
         Top             =   840
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   5503
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmColPParmSegPrendario.frx":035E
         EncabezadosAnchos=   "400-1100-1500-1500-1500-930-900-700-1700-970-0-00"
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
         ColumnasAEditar =   "X-1-2-3-4-5-6-7-8-9-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-3-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-R-R-C-R-L-C-C-C-C"
         FormatosEdit    =   "0-1-0-2-2-3-3-0-3-3-2-0"
         CantEntero      =   10
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit feSegExtMontoPrestamo 
         Height          =   3180
         Left            =   3840
         TabIndex        =   23
         Top             =   720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   5609
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   ""
         FormatosEdit    =   ""
         CantEntero      =   10
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -74910
         TabIndex        =   15
         Top             =   375
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condición"
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
         Left            =   -71880
         TabIndex        =   14
         Top             =   375
         Width           =   7215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -74040
         TabIndex        =   13
         Top             =   600
         Width           =   3450
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condición"
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
         Left            =   -70600
         TabIndex        =   12
         Top             =   600
         Width           =   3100
      End
   End
End
Attribute VB_Name = "frmColPParmSegPrendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'Archivo:  frmColPParmSegPrendario.frm
'JOEP   :  05/12/2017 ---SUBIDO DESDE LA 60
'Resumen:  Registra la Configuracion de Pignoraticios
'******************************************************

Option Explicit
Dim oDPig As COMDColocPig.DCOMColPContrato
Dim pgcMovNro As String 'JOEP20210816 Segmentacion Externa

Public Sub Inicio()
    Call HabilitaContrles(False)
    Call CargaFlexClienteRecurrente(0)
    Call CargaFlexClienteNuevo(1)
    Call CargaFlexMontoPrestamo(2)
    Call CargaDatosMatrizSegExterna 'JOEP20210929 Prendario Externo
    CentraForm Me
    SSTab1.Tab = 0
    Me.Show
End Sub

Public Function HabilitaContrles(ByVal nValor As Boolean, Optional ByVal nTpTab As Integer)
    If nTpTab = 1 Then
        feClientesRecurrentes.Enabled = nValor
        cmdGuardarCR.Enabled = nValor
        cmdCancelarCR.Enabled = nValor
        cmdAgregaCR.Enabled = nValor
        cmdQuitarCR.Enabled = nValor
    ElseIf nTpTab = 2 Then
        feClienteNuevo.Enabled = nValor
        cmdGuardarCN.Enabled = nValor
        cmdCancelarCN.Enabled = nValor
        cmdAgregaCN.Enabled = nValor
        cmdQuitarCN.Enabled = nValor
    ElseIf nTpTab = 3 Then
        feMontoPrestamo.Enabled = nValor
        feSegExtMontoPrestamo.Enabled = nValor  'JOEP20210812 Segmentacion Externa
        cmdGuardarMP.Enabled = nValor
        cmdCancelarMP.Enabled = nValor
        cmdAgregaMP.Enabled = nValor
        cmdQuitarMP.Enabled = nValor
    Else
        feClientesRecurrentes.Enabled = nValor
        feClienteNuevo.Enabled = nValor
        feMontoPrestamo.Enabled = nValor
        feSegExtMontoPrestamo.Enabled = nValor  'JOEP20210812 Segmentacion Externa
        
        cmdGuardarCR.Enabled = nValor
        cmdCancelarCR.Enabled = nValor
        cmdAgregaCR.Enabled = nValor
        cmdQuitarCR.Enabled = nValor
        
        cmdGuardarCN.Enabled = nValor
        cmdCancelarCN.Enabled = nValor
        cmdAgregaCN.Enabled = nValor
        cmdQuitarCN.Enabled = nValor
        
        cmdGuardarMP.Enabled = nValor
        cmdCancelarMP.Enabled = nValor
        cmdAgregaMP.Enabled = nValor
        cmdQuitarMP.Enabled = nValor
    End If
End Function

Public Sub CargaFlexClienteRecurrente(ByVal nValor As Integer)
    Dim rsCR As ADODB.Recordset
    Dim i As Integer
    Set oDPig = New COMDColocPig.DCOMColPContrato
    Set rsCR = oDPig.ObtieneDatosFlexEdit(nValor)
        
        feClientesRecurrentes.Clear
        feClientesRecurrentes.FormaCabecera
        Call LimpiaFlex(feClientesRecurrentes)
        If Not (rsCR.EOF And rsCR.BOF) Then
            For i = 1 To rsCR.RecordCount
                feClientesRecurrentes.AdicionaFila
                    feClientesRecurrentes.TextMatrix(i, 1) = rsCR!cSegmento
                    feClientesRecurrentes.TextMatrix(i, 2) = rsCR!cSubSegmento
                    feClientesRecurrentes.TextMatrix(i, 3) = Format(rsCR!nMontoTasaDesde, "#,##0.00")
                    feClientesRecurrentes.TextMatrix(i, 4) = Format(rsCR!nMontoTasaHasta, "#,##0.00")
                    feClientesRecurrentes.TextMatrix(i, 5) = rsCR!nDiasDesde
                    feClientesRecurrentes.TextMatrix(i, 6) = rsCR!nDiasHasta
                    feClientesRecurrentes.TextMatrix(i, 7) = rsCR!Condicion
                    feClientesRecurrentes.TextMatrix(i, 8) = rsCR!nCantAdjudicado
                    
                    feClientesRecurrentes.TextMatrix(i, 9) = rsCR!nDiasAdj 'Agrego JOEP20180410 Mejora de Segmentacion Pig
                    
                    feClientesRecurrentes.TextMatrix(i, 10) = 2 'Clientes recurrentes
                    feClientesRecurrentes.TextMatrix(i, 11) = rsCR!nId 'JOEP20210812 Segmentacion Externa
                rsCR.MoveNext
            Next i
        End If
        
    Set oDPig = Nothing
    RSClose rsCR
End Sub

Public Sub CargaFlexClienteNuevo(ByVal nValor As Integer)
    Dim rsCN As ADODB.Recordset
    Dim i As Integer
    Set oDPig = New COMDColocPig.DCOMColPContrato
    Set rsCN = oDPig.ObtieneDatosFlexEdit(nValor)
    
        feClienteNuevo.Clear
        feClienteNuevo.FormaCabecera
        Call LimpiaFlex(feClienteNuevo)
        If Not (rsCN.EOF And rsCN.BOF) Then
            For i = 1 To rsCN.RecordCount
                feClienteNuevo.AdicionaFila
                    feClienteNuevo.TextMatrix(i, 1) = rsCN!cSegmento
                    feClienteNuevo.TextMatrix(i, 2) = rsCN!cSubSegmento
                    feClienteNuevo.TextMatrix(i, 3) = Format(rsCN!nMontoTasaDesde, "#,##0.00")
                    feClienteNuevo.TextMatrix(i, 4) = Format(rsCN!nMontoTasaHasta, "#,##0.00")
                    feClienteNuevo.TextMatrix(i, 5) = 1 'Clientes Nuevos
                    feClienteNuevo.TextMatrix(i, 6) = rsCN!nId 'JOEP20210812 Segmentacion Externa
                rsCN.MoveNext
            Next i
        End If
    Set oDPig = Nothing
End Sub

Public Sub CargaFlexMontoPrestamo(ByVal nValor As Integer)
    Dim rsMP As ADODB.Recordset
    Dim i As Integer
    Set oDPig = New COMDColocPig.DCOMColPContrato
    Set rsMP = oDPig.ObtieneDatosFlexEdit(nValor)
        
        feMontoPrestamo.Clear
        feMontoPrestamo.FormaCabecera
        Call LimpiaFlex(feMontoPrestamo)
        If Not (rsMP.EOF And rsMP.BOF) Then
            For i = 1 To rsMP.RecordCount
                feMontoPrestamo.AdicionaFila
                   feMontoPrestamo.TextMatrix(i, 1) = rsMP!cSegmento
                   feMontoPrestamo.TextMatrix(i, 2) = Format(rsMP!nMontoTasa, "#,##0.00")
                   feMontoPrestamo.TextMatrix(i, 3) = rsMP!nId 'JOEP20210812 Segmentacion Externa
                rsMP.MoveNext
            Next i
        End If
    Set oDPig = Nothing
End Sub

Private Sub cmdAgregaCN_Click()
    feClienteNuevo.lbEditarFlex = True
    feClienteNuevo.AdicionaFila
End Sub

Private Sub cmdAgregaCR_Click()
    feClientesRecurrentes.lbEditarFlex = True
    feClientesRecurrentes.AdicionaFila
End Sub

Private Sub cmdAgregaMP_Click()
    feMontoPrestamo.lbEditarFlex = True
    feMontoPrestamo.AdicionaFila
End Sub

Private Sub cmdCancelarCN_Click()
    CargaFlexClienteNuevo (1)
    Call HabilitaContrles(False, 2)
End Sub

Private Sub cmdCancelarCR_Click()
    Call CargaFlexClienteRecurrente(0)
    Call HabilitaContrles(False, 1)
End Sub

Private Sub cmdCancelarMP_Click()
    Call CargaFlexMontoPrestamo(2)
    Call CargaDatosMatrizSegExterna 'JOEP20210929 Prendario Externo
    Call HabilitaContrles(False, 3)
    'JOEP20210812 Segmentacion Externa
        If feSegExtMontoPrestamo.Col >= 2 Then
            feSegExtMontoPrestamo.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
        End If
    'JOEP20210812 Segmentacion Externa
End Sub

Private Sub cmdEditarCN_Click()
    Call HabilitaContrles(True, 2) 'SSTAB 2 Cliente Nuevo
End Sub

Private Sub cmdEditarCR_Click()
    Call HabilitaContrles(True, 1) 'SSTAB 1 Cliente Recurrente
End Sub

Private Sub cmdEditarMP_Click()
    Call HabilitaContrles(True, 3) 'SSTAB 3 Monto Prestamo
    'JOEP20210708 Segmentacion Externa
        If feSegExtMontoPrestamo.Col >= 2 Then
            feSegExtMontoPrestamo.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-9-10-11-12-13-14-15-16"
        End If
    'JOEP20210708 Segmentacion Externa
    
     'para que se posiciones en la primera columna
    feMontoPrestamo.Col = 1
    feMontoPrestamo.row = 1
    feMontoPrestamo.SetFocus
End Sub

Private Sub cmdGuardarMP_Click()
    Dim Guardar As Boolean
    Dim oNPig As COMNColoCPig.NCOMColPContrato
    Dim nMatMP As Variant
    Dim nMatMPSegExt As Variant
    Dim j As Integer
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim objMovNro As COMNContabilidad.NCOMContFunciones
    
    Set objMovNro = New COMNContabilidad.NCOMContFunciones
        pgcMovNro = objMovNro.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set objMovNro = Nothing
    Set oNPig = New COMNColoCPig.NCOMColPContrato
    
    If ValidaDatos(3) Then
        Set nMatMP = Nothing
        ReDim nMatMP(feMontoPrestamo.rows - 1, 10)
        For i = 1 To feMontoPrestamo.rows - 1
                nMatMP(i, 1) = feMontoPrestamo.TextMatrix(i, 3) 'i
                nMatMP(i, 2) = feMontoPrestamo.TextMatrix(i, 1)
                nMatMP(i, 3) = feMontoPrestamo.TextMatrix(i, 2)
        Next i
        
        Set nMatMPSegExt = Nothing
         
        'JOEP20210812 Segmentacion Externa
        ReDim nMatMPSegExt(576, 60)
        X = 1
        
        For i = 1 To feSegExtMontoPrestamo.rows - 1
            For j = 2 To feSegExtMontoPrestamo.cols - 1
                nMatMPSegExt(X, 1) = i                                       'no se usa en el script.
                X = X + 1
                nMatMPSegExt(X, 2) = feSegExtMontoPrestamo.TextMatrix(i, 0)  'id Segmento Externo
                X = X + 1
                nMatMPSegExt(X, 3) = feSegExtMontoPrestamo.TextMatrix(0, j) 'Segmento
                X = X + 1
                nMatMPSegExt(X, 4) = feSegExtMontoPrestamo.TextMatrix(i, j)  'Monto
                X = X + 1
            Next j
        Next i
        'JOEP20210812 Segmentacion Externa
    
        If MsgBox("Los datos serán grabados. ¿Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        'Guardar = oNPig.RegistroDatosSegmentacion(3, nMatMP)
        Guardar = oNPig.RegistroDatosSegmentacion(3, nMatMP, nMatMPSegExt, pgcMovNro) 'JOEP20210812 Segmentacion Externa
        
        If Guardar Then
            MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            Call CargaFlexMontoPrestamo(2) 'JOEP20210812 Segmentacion Externa
			Call CargaDatosMatrizSegExterna 'JOEP20210929 Prendario Externo
            Call HabilitaContrles(False, 3) 'SSTAB 3 Monto Prestamo
        Else
            MsgBox "Hubo errores al grabar la información", vbInformation, "Aviso"
        End If
        
    End If
End Sub

Private Sub cmdGuardarCN_Click()
    Dim Guardar As Boolean
    Dim oNPig As COMNColoCPig.NCOMColPContrato
    Dim nMatCliNuevo As Variant
    Dim i As Integer
    Dim objMovNro As COMNContabilidad.NCOMContFunciones
    
    Set oNPig = New COMNColoCPig.NCOMColPContrato
    Set objMovNro = New COMNContabilidad.NCOMContFunciones
        pgcMovNro = objMovNro.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set objMovNro = Nothing
    
    If ValidaDatos(2) Then
    
        Set nMatCliNuevo = Nothing
        ReDim nMatCliNuevo(feClienteNuevo.rows - 1, 10)
        For i = 1 To feClienteNuevo.rows - 1
                nMatCliNuevo(i, 1) = feClienteNuevo.TextMatrix(i, 6) 'i
                nMatCliNuevo(i, 2) = feClienteNuevo.TextMatrix(i, 1)
                nMatCliNuevo(i, 3) = feClienteNuevo.TextMatrix(i, 2)
                nMatCliNuevo(i, 4) = feClienteNuevo.TextMatrix(i, 3)
                nMatCliNuevo(i, 5) = feClienteNuevo.TextMatrix(i, 4)
                nMatCliNuevo(i, 6) = 1
        Next i
    
        If MsgBox("Los datos serán grabados. ¿Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        'Guardar = oNPig.RegistroDatosSegmentacion(2, nMatCliNuevo)
        Guardar = oNPig.RegistroDatosSegmentacion(2, nMatCliNuevo, , pgcMovNro)
        
        If Guardar Then
            MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            Call CargaFlexClienteNuevo(1) 'JOEP20210812 Segmentacion Externa
            Call HabilitaContrles(False, 2) 'SSTAB 2 Cliente Nuevo
        Else
            MsgBox "Hubo errores al grabar la información", vbInformation, "Aviso"
        End If
        
    End If
End Sub

Private Sub cmdGuardarCR_Click()
    Dim Guardar As Boolean
    Dim oNPig As COMNColoCPig.NCOMColPContrato
    Dim nMatCliRecurrentes As Variant
    Dim i As Integer
    Dim objMovNro As COMNContabilidad.NCOMContFunciones
    
    Set oNPig = New COMNColoCPig.NCOMColPContrato
    Set objMovNro = New COMNContabilidad.NCOMContFunciones
        pgcMovNro = objMovNro.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set objMovNro = Nothing
    
    If ValidaDatos(1) Then
        
        Set nMatCliRecurrentes = Nothing
        ReDim nMatCliRecurrentes(feClientesRecurrentes.rows - 1, 11)
        For i = 1 To feClientesRecurrentes.rows - 1
                nMatCliRecurrentes(i, 1) = feClientesRecurrentes.TextMatrix(i, 11) 'i
                nMatCliRecurrentes(i, 2) = feClientesRecurrentes.TextMatrix(i, 1)
                nMatCliRecurrentes(i, 3) = feClientesRecurrentes.TextMatrix(i, 2)
                nMatCliRecurrentes(i, 4) = feClientesRecurrentes.TextMatrix(i, 3)
                nMatCliRecurrentes(i, 5) = feClientesRecurrentes.TextMatrix(i, 4)
                nMatCliRecurrentes(i, 6) = feClientesRecurrentes.TextMatrix(i, 5)
                nMatCliRecurrentes(i, 7) = feClientesRecurrentes.TextMatrix(i, 6)
                nMatCliRecurrentes(i, 8) = Right(feClientesRecurrentes.TextMatrix(i, 7), 1)
                nMatCliRecurrentes(i, 9) = feClientesRecurrentes.TextMatrix(i, 8)
                
                nMatCliRecurrentes(i, 10) = feClientesRecurrentes.TextMatrix(i, 9)
                
                nMatCliRecurrentes(i, 11) = 2
        Next i
    
        If MsgBox("Los datos serán grabados. ¿Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        'Guardar = oNPig.RegistroDatosSegmentacion(1, nMatCliRecurrentes)
        Guardar = oNPig.RegistroDatosSegmentacion(1, nMatCliRecurrentes, , pgcMovNro)
        
        If Guardar Then
            MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            Call CargaFlexClienteRecurrente(0) 'JOEP20210812 Segmentacion Externa
            Call HabilitaContrles(False, 1) 'SSTAB 1 Cliente Recurrente
        Else
            MsgBox "Hubo errores al grabar la información", vbInformation, "Aviso"
        End If
        
    End If
    
    Set oNPig = Nothing
    
End Sub

Public Function ValidaDatos(ByVal nValor As Integer) As Boolean
    Dim i As Integer
    Dim j As Integer
    ValidaDatos = True
    
    If nValor = 1 Then
        For i = 1 To feClientesRecurrentes.rows - 1
            If feClientesRecurrentes.TextMatrix(i, 1) = "" Or feClientesRecurrentes.TextMatrix(i, 2) = "" Or feClientesRecurrentes.TextMatrix(i, 3) = "" Or feClientesRecurrentes.TextMatrix(i, 4) = "" Or feClientesRecurrentes.TextMatrix(i, 5) = "" Or feClientesRecurrentes.TextMatrix(i, 6) = "" Or feClientesRecurrentes.TextMatrix(i, 7) = "" Or feClientesRecurrentes.TextMatrix(i, 8) = "" Then
                MsgBox "Uno de los datos esta vacío, verificar los datos", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Next i
        For i = 1 To feClientesRecurrentes.rows - 1
            If feClientesRecurrentes.TextMatrix(i, 1) = feClientesRecurrentes.TextMatrix(i, 2) Then
                MsgBox "El campo Segmentación no puede ser igual al campo Sub Segmentación", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Next i
        For i = 1 To feClientesRecurrentes.rows - 1
            If feClientesRecurrentes.TextMatrix(i, 1) <> Left(feClientesRecurrentes.TextMatrix(i, 2), 1) Then
                MsgBox "La Sub Segmentación " & feClientesRecurrentes.TextMatrix(i, 2) & " no pertenece al grupo de la Segmentación " & feClientesRecurrentes.TextMatrix(i, 1), vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Next i
        For i = 1 To feClientesRecurrentes.rows - 1
            For j = 1 To feClientesRecurrentes.rows - 1
                If i <> j Then
                    If feClientesRecurrentes.TextMatrix(i, 2) = feClientesRecurrentes.TextMatrix(j, 2) Then
                        MsgBox "Existe datos duplicados: " & feClientesRecurrentes.TextMatrix(i, 2), vbInformation, "Aviso"
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            Next j
        Next i
    ElseIf nValor = 2 Then
        For i = 1 To feClienteNuevo.rows - 1
            If feClienteNuevo.TextMatrix(i, 1) = "" Or feClienteNuevo.TextMatrix(i, 2) = "" Or feClienteNuevo.TextMatrix(i, 3) = "" Then
                MsgBox "Uno de los datos esta vacío, verificar los datos", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Next i
        For i = 1 To feClienteNuevo.rows - 1
            If feClienteNuevo.TextMatrix(i, 1) = feClienteNuevo.TextMatrix(i, 2) Then
                MsgBox "El campo Segmentación no puede ser igual al campo Sub Segmentación", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Next i
        For i = 1 To feClienteNuevo.rows - 1
            If feClienteNuevo.TextMatrix(i, 1) <> Left(feClienteNuevo.TextMatrix(i, 2), 1) Then
                MsgBox "La Sub Segmentación " & feClienteNuevo.TextMatrix(i, 2) & " no pertenece al grupo de la Segmentación " & feClienteNuevo.TextMatrix(i, 1), vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Next i
        For i = 1 To feClienteNuevo.rows - 1
            For j = 1 To feClienteNuevo.rows - 1
                If i <> j Then
                    If feClienteNuevo.TextMatrix(i, 2) = feClienteNuevo.TextMatrix(j, 2) Then
                        MsgBox "Existe datos duplicados: " & feClienteNuevo.TextMatrix(i, 2), vbInformation, "Aviso"
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            Next j
        Next i
    ElseIf nValor = 3 Then
        For i = 1 To feMontoPrestamo.rows - 1
            If feMontoPrestamo.TextMatrix(i, 1) = "" Or feMontoPrestamo.TextMatrix(i, 2) = "" Then
                MsgBox "Uno de los datos esta vacío, verificar los datos", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
        Next i
        For i = 1 To feMontoPrestamo.rows - 1
            For j = 1 To feMontoPrestamo.rows - 1
                If i <> j Then
                    If feMontoPrestamo.TextMatrix(i, 1) = feMontoPrestamo.TextMatrix(j, 1) Then
                        MsgBox "Existe datos duplicados: " & feMontoPrestamo.TextMatrix(i, 1), vbInformation, "Aviso"
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            Next j
        Next i
    End If
    
End Function

Private Sub cmdQuitarCN_Click()
    If MsgBox("Está seguro de eliminar registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feClienteNuevo.EliminaFila (feClienteNuevo.row)
    End If
End Sub

Private Sub cmdQuitarCR_Click()
    If MsgBox("Está seguro de eliminar el Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feClientesRecurrentes.EliminaFila (feClientesRecurrentes.row)
    End If
End Sub

Private Sub cmdQuitarMP_Click()
    If MsgBox("Está seguro de eliminar el registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feMontoPrestamo.EliminaFila (feMontoPrestamo.row)
    End If
End Sub

Private Sub feClientesRecurrentes_Click()
    Dim rsCargaCondicion As ADODB.Recordset
    Set oDPig = New COMDColocPig.DCOMColPContrato
    
    Set rsCargaCondicion = oDPig.ObtieneDatosCondicionFlex()
    
    If Not (rsCargaCondicion.EOF And rsCargaCondicion.BOF) Then
        If feClientesRecurrentes.Col = 7 Then
          feClientesRecurrentes.CargaCombo rsCargaCondicion
        End If
    End If
   
    Set oDPig = Nothing
End Sub

Private Sub feClienteNuevo_OnCellChange(pnRow As Long, pnCol As Long)
    Dim cString As String
    Dim cNumero As String
    cString = "QqWwEeRrTtYyUuIiOoPpAaSsDdFfGgHhJjKkLlÑñZzXxCcVvBbNnMm"
    cNumero = "1234567890"
    
    If feClienteNuevo.Col = 1 Then
        If Len(feClienteNuevo.TextMatrix(feClienteNuevo.row, 1)) > 1 Then
            MsgBox "Solo se acepta un carácter. Ej. [A, . . . ,Z ]", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 1) = ""
        Else
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 1) = UCase(feClienteNuevo.TextMatrix(feClienteNuevo.row, 1))
        End If
    End If
    If feClienteNuevo.Col = 2 Then
        If Len(feClienteNuevo.TextMatrix(feClienteNuevo.row, 2)) > 2 Then
            MsgBox "Solo se acepta dos caracteres. Ej. [A1, . . . ,Z9 ]", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 2) = ""
        Else
            If InStr(cString, Left(feClienteNuevo.TextMatrix(feClienteNuevo.row, 2), 1)) > 0 And InStr(cNumero, Right(feClienteNuevo.TextMatrix(feClienteNuevo.row, 2), 1)) > 0 Then
                feClienteNuevo.TextMatrix(feClienteNuevo.row, 2) = UCase(Left(feClienteNuevo.TextMatrix(feClienteNuevo.row, 2), 1)) & Right(feClienteNuevo.TextMatrix(feClienteNuevo.row, 2), 1)
            Else
                MsgBox "Solo se acepta dos caracteres. Ej. [A1, . . . ,Z9 ]", vbInformation, "Aviso"
                feClienteNuevo.TextMatrix(feClienteNuevo.row, 2) = ""
            End If
        End If
    End If
    If feClienteNuevo.Col = 3 Then
        If Len((Replace(feClienteNuevo.TextMatrix(feClienteNuevo.row, 3), ",", ""))) > 10 Then
            MsgBox "Solo se acepta 10 dígitos", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 3) = 0#
        End If
        If val(Replace(feClienteNuevo.TextMatrix(feClienteNuevo.row, 3), ",", "")) < 0 Then
            MsgBox "Solo se acepta números positivos", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 3) = 0#
        End If
        If InStr(1, feClienteNuevo.TextMatrix(feClienteNuevo.row, 3), "-") > 0 Then
            MsgBox "Solo se acepta números", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 3) = 0#
        Else
            If Len(feClienteNuevo.TextMatrix(feClienteNuevo.row, 3)) = 1 Then
                If feClienteNuevo.TextMatrix(feClienteNuevo.row, 3) = "." Then
                    MsgBox "Solo se acepta números", vbInformation, "Aviso"
                    feClienteNuevo.TextMatrix(feClienteNuevo.row, 3) = 0#
                End If
            End If
        End If
        If val(Replace(feClienteNuevo.TextMatrix(feClienteNuevo.row, 3), ",", "")) > val(Replace(feClienteNuevo.TextMatrix(feClienteNuevo.row, 4), ",", "")) Then
            MsgBox "El [Monto Tasa Desde] no debe ser mayor al [Monto Tasa Hasta]", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 3) = 0#
        End If
    End If
    If feClienteNuevo.Col = 4 Then
        If Len((Replace(feClienteNuevo.TextMatrix(feClienteNuevo.row, 4), ",", ""))) > 10 Then
            MsgBox "Solo se acepta 10 dígitos", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 4) = 0#
        End If
        If val(Replace(feClienteNuevo.TextMatrix(feClienteNuevo.row, 4), ",", "")) < 0 Then
            MsgBox "Solo se acepta números positivos", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 4) = 0#
        End If
        If InStr(1, feClienteNuevo.TextMatrix(feClienteNuevo.row, 4), "-") > 0 Then
            MsgBox "Solo se acepta números", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 4) = 0#
        Else
            If Len(feClienteNuevo.TextMatrix(feClienteNuevo.row, 4)) = 1 Then
                If feClienteNuevo.TextMatrix(feClienteNuevo.row, 4) = "." Then
                    MsgBox "Solo se acepta números", vbInformation, "Aviso"
                    feClienteNuevo.TextMatrix(feClienteNuevo.row, 4) = 0#
                End If
            End If
        End If
        If val(Replace(feClienteNuevo.TextMatrix(feClienteNuevo.row, 4), ",", "")) < val(Replace(feClienteNuevo.TextMatrix(feClienteNuevo.row, 3), ",", "")) Then
            MsgBox "El [Monto Tasa Hasta] no debe ser menor al [Monto Tasa Desde]", vbInformation, "Aviso"
            feClienteNuevo.TextMatrix(feClienteNuevo.row, 4) = 0#
        End If
    End If
End Sub

Private Sub feClientesRecurrentes_DblClick()
Dim rsCargaCondicion As ADODB.Recordset
    Set oDPig = New COMDColocPig.DCOMColPContrato
    
    Set rsCargaCondicion = oDPig.ObtieneDatosCondicionFlex()
    
    If Not (rsCargaCondicion.EOF And rsCargaCondicion.BOF) Then
        If feClientesRecurrentes.Col = 7 Then
          feClientesRecurrentes.CargaCombo rsCargaCondicion
        End If
    End If
   
    Set oDPig = Nothing
End Sub

Private Sub feClientesRecurrentes_EnterCell()
Dim rsCargaCondicion As ADODB.Recordset
    Set oDPig = New COMDColocPig.DCOMColPContrato
    
    Set rsCargaCondicion = oDPig.ObtieneDatosCondicionFlex()
    
    If Not (rsCargaCondicion.EOF And rsCargaCondicion.BOF) Then
        If feClientesRecurrentes.Col = 7 Then
          feClientesRecurrentes.CargaCombo rsCargaCondicion
        End If
    End If
   
    Set oDPig = Nothing
End Sub

Private Sub feClientesRecurrentes_KeyPress(KeyAscii As Integer)
Dim rsCargaCondicion As ADODB.Recordset
    Set oDPig = New COMDColocPig.DCOMColPContrato
    
    Set rsCargaCondicion = oDPig.ObtieneDatosCondicionFlex()
    
    If Not (rsCargaCondicion.EOF And rsCargaCondicion.BOF) Then
        If feClientesRecurrentes.Col = 7 Then
          feClientesRecurrentes.CargaCombo rsCargaCondicion
        End If
    End If
   
    Set oDPig = Nothing
End Sub

Private Sub feClientesRecurrentes_OnCellChange(pnRow As Long, pnCol As Long)
Dim cString As String
Dim cNumero As String
Dim rsCargaCondicion As ADODB.Recordset
    cString = "QqWwEeRrTtYyUuIiOoPpAaSsDdFfGgHhJjKkLlÑñZzXxCcVvBbNnMm"
    cNumero = "1234567890"
    
    Set oDPig = New COMDColocPig.DCOMColPContrato
    Set rsCargaCondicion = oDPig.ObtieneDatosCondicionFlex()
    
    If feClientesRecurrentes.Col = 1 Then
        If Len(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 1)) > 1 Then
            MsgBox "Solo se acepta un carácter. Ej. [A, . . . ,Z ]", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 1) = ""
        Else
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 1) = UCase(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 1))
        End If
    End If
    If feClientesRecurrentes.Col = 2 Then
        If Len(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 2)) > 2 Then
            MsgBox "Solo se acepta dos caracteres. Ej. [A1, . . . ,Z9 ]", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 2) = ""
        Else
            If InStr(cString, Left(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 2), 1)) > 0 And InStr(cNumero, Right(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 2), 1)) > 0 Then
                feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 2) = UCase(Left(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 2), 1)) & Right(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 2), 1)
            Else
                MsgBox "Solo se acepta dos caracteres. Ej. [A1, . . . ,Z9 ]", vbInformation, "Aviso"
                feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 2) = ""
            End If
        End If
    End If
    If feClientesRecurrentes.Col = 3 Then
        If Len((Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3), ",", ""))) > 13 Then
            MsgBox "Solo se acepta 10 dígitos", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3) = 0#
        End If
        If val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3), ",", "")) < 0 Then
            MsgBox "Solo se acepta números positivos", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3) = 0#
        End If
        If InStr(1, feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3), "-") > 0 Then
            MsgBox "Solo se acepta números", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3) = 0#
        Else
            If Len(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3)) = 1 Then
                If feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3) = "." Then
                    MsgBox "Solo se acepta números", vbInformation, "Aviso"
                    feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3) = 0#
                End If
            End If
        End If
        If val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3), ",", "")) > val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4), ",", "")) Then
            MsgBox "El [Monto Tasa Desde] no debe ser mayor al [Monto Tasa Hasta].", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3) = 0#
        End If
    End If
    If feClientesRecurrentes.Col = 4 Then
        If Len((Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4), ",", ""))) > 13 Then
            MsgBox "Solo se acepta 10 dígitos", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4) = 0#
        End If
        If val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4), ",", "")) < 0 Then
            MsgBox "Solo se acepta números positivos", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4) = 0#
        End If
        If InStr(1, feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4), "-") > 0 Then
            MsgBox "Solo se acepta números", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4) = 0#
        Else
            If Len(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4)) = 1 Then
                If feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4) = "." Then
                    MsgBox "Solo se acepta números", vbInformation, "Aviso"
                    feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4) = 0#
                End If
            End If
        End If
        If val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4), ",", "")) < val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 3), ",", "")) Then
            MsgBox "El [Monto Tasa Hasta] no debe ser menor al [Monto Tasa Desde].", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 4) = 0#
        End If
    End If
    If feClientesRecurrentes.Col = 5 Then
        If Len(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 5), ",", "")) > 6 Then
            MsgBox "Solo se acepta 6 dígitos", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 5) = 0#
        End If
        If feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 6) <> "" Then
            If val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 5), ",", "")) > val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 6), ",", "")) Then
                MsgBox "Los [Diás Desde] no puede ser mayor al [Diás Hasta]", vbInformation, "Aviso"
                feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 5) = 0#
            End If
        End If
    End If
    If feClientesRecurrentes.Col = 6 Then
        If Len(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 6), ",", "")) > 6 Then
            MsgBox "Solo se acepta 6 dígitos", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 6) = 0#
        End If
        If feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 5) <> "" Then
            If val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 6), ",", "")) < val(Replace(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 5), ",", "")) Then
                MsgBox "Los [Diás Hasta] no puede ser menor al [Diás Desde]", vbInformation, "Aviso"
                feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 6) = 0#
            End If
        End If
    End If
    If feClientesRecurrentes.Col = 8 Then
        If Len(feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 8)) > 1 Then
            MsgBox "Solo se acepta un carácter. Ej. [0, . . . ,9 ]", vbInformation, "Aviso"
            feClientesRecurrentes.TextMatrix(feClientesRecurrentes.row, 8) = 0#
        End If
    End If
    
     If Not (rsCargaCondicion.EOF And rsCargaCondicion.BOF) Then
        If feClientesRecurrentes.Col = 7 Then
          feClientesRecurrentes.CargaCombo rsCargaCondicion
        End If
    End If
    Set oDPig = Nothing
End Sub

Private Sub feMontoPrestamo_OnCellChange(pnRow As Long, pnCol As Long)
Dim cString As String
Dim cNumero As String
Dim cCaracter As String
    cString = "QqWwEeRrTtYyUuIiOoPpAaSsDdFfGgHhJjKkLlÑñZzXxCcVvBbNnMm"
    cNumero = "1234567890"
    cCaracter = "-"
    
    If feMontoPrestamo.Col = 1 Then
        If Len(feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 1)) > 2 Then
            MsgBox "Solo se acepta dos caracteres. Ej. [A1, . . . ,Z9 ]", vbInformation, "Aviso"
            feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 1) = ""
        Else
            If InStr(cString, Left(feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 1), 1)) > 0 And InStr(cNumero, Right(feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 1), 1)) > 0 Then
                feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 1) = UCase(feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 1))
            Else
                MsgBox "Solo se acepta dos caracteres. Ej. [A1, . . . ,Z9 ]", vbInformation, "Aviso"
                feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 1) = ""
            End If
        End If
    End If
    If feMontoPrestamo.Col = 2 Then
        If InStr(1, feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 2), "-") > 0 Then
            MsgBox "Solo se acepta números", vbInformation, "Aviso"
            feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 2) = 0#
        Else
            If Len(feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 2)) = 1 Then
                If feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 2) = "." Then
                    MsgBox "Solo se acepta números", vbInformation, "Aviso"
                    feMontoPrestamo.TextMatrix(feMontoPrestamo.row, 2) = 0#
                End If
            End If
        End If
    End If
End Sub

'JOEP20210812 Segmentacion Externa
Private Sub CargaDatosMatrizSegExterna()
    Dim lsEncabezadosNombres As String
    Dim lsColumnasAEditar As String
    Dim lsEncabezadosAlineacion As String
    Dim lsEncabezadosAnchos As String
    Dim lsFormatosEdit As String
    Dim lsListaControles As String
    Dim i As Integer, j As Integer
    Dim ixFlex As Integer
    Dim nRSE As Integer

    Dim obTituloSegPreInter As COMDColocPig.DCOMColPContrato
    Set obTituloSegPreInter = New COMDColocPig.DCOMColPContrato

    Dim rsSegPrenInterna As ADODB.Recordset
    Dim rsSegPrenExterna As ADODB.Recordset
    Dim rsSegPrenDatos As ADODB.Recordset

    Set rsSegPrenInterna = obTituloSegPreInter.ConfiSegTitulo(1)
    Set rsSegPrenExterna = obTituloSegPreInter.ConfiSegTitulo(2)

    'seteamos Variables
        lsEncabezadosNombres = ""
        lsColumnasAEditar = ""
        lsEncabezadosAlineacion = ""
        lsEncabezadosAnchos = ""
        lsFormatosEdit = ""
        lsListaControles = ""

    If Not (rsSegPrenInterna.BOF And rsSegPrenInterna.EOF) And Not (rsSegPrenExterna.BOF And rsSegPrenExterna.EOF) Then

        FormateaFlex feSegExtMontoPrestamo
        feSegExtMontoPrestamo.cols = 2

        lsEncabezadosNombres = "#-Seg. Externa"
        lsColumnasAEditar = "X-X"
        lsEncabezadosAlineacion = "C-L"
        lsEncabezadosAnchos = "0-1500"
        lsFormatosEdit = "0-0"
        lsListaControles = "0-0"

        For j = 1 To rsSegPrenInterna.RecordCount
            lsEncabezadosNombres = lsEncabezadosNombres & "-" & rsSegPrenInterna!cSegmento
            lsColumnasAEditar = lsColumnasAEditar & "-X"
            lsEncabezadosAlineacion = lsEncabezadosAlineacion & "-C"
            lsEncabezadosAnchos = lsEncabezadosAnchos & "-700"
            lsFormatosEdit = lsFormatosEdit & "-0"
            lsListaControles = lsListaControles & IIf(i = 1, "-0", "-0")
            rsSegPrenInterna.MoveNext
        Next

        feSegExtMontoPrestamo.EncabezadosNombres = lsEncabezadosNombres
        feSegExtMontoPrestamo.ColumnasAEditar = lsColumnasAEditar
        feSegExtMontoPrestamo.EncabezadosAlineacion = lsEncabezadosAlineacion
        feSegExtMontoPrestamo.EncabezadosAnchos = lsEncabezadosAnchos
        feSegExtMontoPrestamo.FormatosEdit = lsFormatosEdit
        feSegExtMontoPrestamo.ListaControles = lsListaControles

        For i = 1 To rsSegPrenExterna.RecordCount
            feSegExtMontoPrestamo.AdicionaFila , , True
            ixFlex = feSegExtMontoPrestamo.rows - 1

            feSegExtMontoPrestamo.row = i

            feSegExtMontoPrestamo.TextMatrix(ixFlex, 0) = rsSegPrenExterna!nSegmentoExterno
            feSegExtMontoPrestamo.TextMatrix(ixFlex, 1) = rsSegPrenExterna!cSegmentoExterno

            'Se empieza de la columna 2 porque el flexedit al empezar de la columna 1 y si se daba click en la cabecera quitaba la primera columna
            Set rsSegPrenDatos = obTituloSegPreInter.ConfiSegExtDatos(feSegExtMontoPrestamo.TextMatrix(ixFlex, 0))
            For j = 2 To rsSegPrenInterna.RecordCount + 1
                feSegExtMontoPrestamo.TextMatrix(ixFlex, j) = Format(rsSegPrenDatos!nMontoTasa, "#0.00")
                rsSegPrenDatos.MoveNext
            Next
            rsSegPrenExterna.MoveNext
        Next

        feSegExtMontoPrestamo.FixedCols = IIf(feSegExtMontoPrestamo.cols >= 3, 2, 0)
    End If

RSClose rsSegPrenInterna
RSClose rsSegPrenExterna
Set obTituloSegPreInter = Nothing
Set obTituloSegPreInter = Nothing
End Sub

Private Sub feSegExtMontoPrestamo_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

    If IsNumeric(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol)) = False Then
        MsgBox "Dato Incorrecto", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    If IsNumeric(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol)) = True Then
        If CDbl(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol)) < 0 Then
            MsgBox "Dato Incorrecto", vbInformation, "Aviso"
            Cancel = False
            feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol) = Format(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol), "#,0.00")
            SendKeys "{TAB}"
            Exit Sub
        End If
        
        If feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol) = "0." Or feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol) = ".0" Then
            MsgBox "Ingrese el dato correcto", vbInformation, "Aviso"
            Cancel = False
            feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol) = Format(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol), "#,0.00")
            SendKeys "{TAB}"
            Exit Sub
        End If
        
        If Len(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol)) > "8" Then
            MsgBox "Solo es permitido 5 digito", vbInformation, "Aviso"
            Cancel = False
            feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol) = Format(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol), "#,0.00")
            SendKeys "{TAB}"
            Exit Sub
        End If
        
    End If
    
    If IsNumeric(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol)) = True Then
        If CDbl(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol)) > 0 Then
            feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol) = Format(feSegExtMontoPrestamo.TextMatrix(pnRow, pnCol), "#,0.00")
        End If
    End If
End Sub
'JOEP20210812 Segmentacion Externa
