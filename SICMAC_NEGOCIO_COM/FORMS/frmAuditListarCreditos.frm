VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAuditListarCreditos 
   Caption         =   "Listado de Créditos"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAuditListarCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "LISTADO DE CREDITOS"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cCtaCod"
            Caption         =   "Código Cuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cAgeDescripcion"
            Caption         =   "Agencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cTipoCred"
            Caption         =   "T. Crédito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            Size            =   182
            BeginProperty Column00 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1995.024
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAuditListarCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsmensaje As String
Dim lsCodCta As String

Private Sub Command1_Click()
    If ValidarCredito(dgBuscar.Columns(0).Text, Format(frmRevisionRegistrar.mskPeriodo1Del.Text, "yyyymmdd")) = False Then
        lsCodCta = dgBuscar.Columns(0).Text
        If frmRevisionRegistrar.lsmensaje = "" Then
            Unload Me
        End If
        frmRevisionRegistrar.txtCodCta.Text = lsCodCta
        frmRevisionRegistrar.BuscarValores
    Else
        MsgBox ("La Revisión ya esta Registrada"), vbCritical, Me.Caption
        frmRevisionRegistrar.txtCodigo.Text = ""
        Unload Me
    End If
End Sub

Public Function ValidarCredito(ByVal lsCodCta As String, lsFCierre As String) As Boolean
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs As ADODB.Recordset
    Dim Mensaje As String
    Dim Valor As Boolean
    ValidarCredito = True
    lsmensaje = ""
    Set rs = objCOMNAuditoria.ValRevision(lsCodCta, lsFCierre, lsmensaje)
    If lsmensaje = "" Then
        ValidarCredito = True
        Exit Function
    Else
       ValidarCredito = False
       Exit Function
    End If
    Set rs = Nothing
    Set objCOMNAuditoria = Nothing
    
End Function

Public Sub CargarCreditos(ByRef sBoll As Boolean, ByVal lsCodPers As String)
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs As ADODB.Recordset
    
    lsmensaje = ""
    Set rs = objCOMNAuditoria.DarCreditoXPersona(lsCodPers, Format(frmRevisionRegistrar.mskPeriodo1Del.Text, "yyyymmdd"), lsmensaje)
    If lsmensaje = "" Then
        Set dgBuscar.DataSource = rs
        dgBuscar.Refresh
        Screen.MousePointer = 0
        sBoll = True
    Else
        frmRevisionRegistrar.LimpiarControles frmRevisionRegistrar, True, False, False
        frmRevisionRegistrar.txtCodigo.Text = ""
        frmRevisionRegistrar.txtFRegistro.Text = "__/__/____"
        frmRevisionRegistrar.txtFSDCmac.Text = "__/__/____"
        frmRevisionRegistrar.txtFSDSF.Text = "__/__/____"
        frmRevisionRegistrar.CargarDatosLoad 2
        sBoll = False
        MsgBox ("Usuario No tiene Créditos"), vbCritical, "SICMACT"
    End If
    
    Set rs = Nothing
    Set objCOMNAuditoria = Nothing
End Sub
