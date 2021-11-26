VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmColPTarifarioNotific 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarifario de Carta Notariales"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmColPTarifarioNotific.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Agencias"
      TabPicture(0)   =   "frmColPTarifarioNotific.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdTarifario"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdHistorico"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   310
         Left            =   4320
         TabIndex        =   5
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdHistorico 
         Caption         =   "Historico"
         Height          =   310
         Left            =   1340
         TabIndex        =   4
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdTarifario 
         Caption         =   "Editar"
         Height          =   310
         Left            =   120
         TabIndex        =   3
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5535
         Begin SICMACT.FlexEdit feTarifario 
            Height          =   4215
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   7435
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Agencia-Costo-Fecha inicio-nTarifarioID-nDatoOriginal"
            EncabezadosAnchos=   "400-2500-800-1200-0-0"
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
            ColumnasAEditar =   "X-X-2-3-X-X"
            ListaControles  =   "0-0-0-2-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-C-C-C"
            FormatosEdit    =   "0-0-2-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
End
Attribute VB_Name = "frmColPTarifarioNotific"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPTarifarioNotific
'** Descripción : Formulario que administra el tarifario de costo por carta notarial en las agencias.
'** Creación    : RECO, 20160229 - ERS056-2015
'**********************************************************************************************

Option Explicit

Dim nTipoOpe As Integer

Private Sub cmdCerrar_Click()
    Call CargarDatos
    Unload Me
End Sub

Private Sub CargarDatos()
    Dim oColP As New COMDColocPig.DCOMColPActualizaBD
    Dim drDatosTarif As New ADODB.Recordset
    Dim i As Integer
    
    Set drDatosTarif = oColP.PignoObtieneTarifarioNotific(PigTipoTarifarioNotif.gPigTarifarioAge, "")
    
    
    If Not (drDatosTarif.EOF And drDatosTarif.BOF) Then
        feTarifario.Clear
        FormateaFlex feTarifario
        For i = 1 To drDatosTarif.RecordCount
            feTarifario.AdicionaFila
            feTarifario.TextMatrix(i, 1) = drDatosTarif!cAgeDescripcion
            feTarifario.TextMatrix(i, 2) = Format(drDatosTarif!nValor, gcFormView)
            feTarifario.TextMatrix(i, 3) = Format(drDatosTarif!dFecIni, "dd/MM/YYYY")
            feTarifario.TextMatrix(i, 4) = drDatosTarif!nTarifarioID
            feTarifario.TextMatrix(i, 5) = "0"
            drDatosTarif.MoveNext
        Next
        feTarifario.Top = 1
        feTarifario.TopRow = 1
    End If
End Sub

Private Sub cmdHistorico_Click()
    If nTipoOpe = 1 Then
        frmColPTarifarioNotificHis.Inicia feTarifario.TextMatrix(feTarifario.row, 4), feTarifario.TextMatrix(feTarifario.row, 1)
    Else
        nTipoOpe = 1
        feTarifario.lbEditarFlex = False
        cmdTarifario.Caption = "Editar"
        cmdHistorico.Caption = "Historico"
        Call CargarDatos
    End If
End Sub

Private Sub cmdTarifario_Click()
    If nTipoOpe = 1 Then
        nTipoOpe = 2
        feTarifario.lbEditarFlex = True
        cmdTarifario.Caption = "Guardar"
        cmdHistorico.Caption = "Cancelar"
    Else
        nTipoOpe = 1
        feTarifario.lbEditarFlex = False
        Call GuardarTarifario
        cmdTarifario.Caption = "Editar"
        cmdHistorico.Caption = "Historico"
    End If
End Sub

Private Sub feTarifario_OnCellChange(pnRow As Long, pnCol As Long)
    Select Case pnCol
        Case 2, 3
            feTarifario.TextMatrix(pnRow, 5) = "1"
    End Select
     If ValidaMontos = False Then MsgBox "Valor no pueden ser negativo", vbInformation, "Alerta SICMAC"
End Sub

Private Sub Form_Load()
    Call CargarDatos
    feTarifario.lbEditarFlex = False
    nTipoOpe = 1
End Sub

Private Sub GuardarTarifario()
    Dim oTarj As New COMDColocPig.DCOMColPActualizaBD
    Dim nIndex As Integer
    
    For nIndex = 1 To feTarifario.Rows - 1
        If feTarifario.TextMatrix(nIndex, 5) <> "0" Then
            Call oTarj.ActualizaTarifarioPigno(feTarifario.TextMatrix(nIndex, 4), feTarifario.TextMatrix(nIndex, 3), feTarifario.TextMatrix(nIndex, 2), gsCodUser)
        End If
    Next
    Call CargarDatos
End Sub

Private Function ValidaMontos() As Boolean
    Dim nIndice As Integer
    
    ValidaMontos = True
    
    For nIndice = 1 To feTarifario.Rows - 1
        On Local Error Resume Next
        
        If feTarifario.TextMatrix(nIndice, 2) < 0 Then
        feTarifario.TextMatrix(nIndice, 2) = "0.00"
            ValidaMontos = False
        End If
        
        If Err <> 0 Then
            ValidaMontos = ValidaMontos = False
        End If
    Next
    
End Function
