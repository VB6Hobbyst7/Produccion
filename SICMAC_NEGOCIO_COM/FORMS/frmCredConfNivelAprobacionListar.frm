VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmCredConfNivelAprobacionListar 
   Caption         =   "Listar Niveles de Aprobación"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "cNivAprCod"
            Caption         =   "cNivAprCod"
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
            DataField       =   "cNivelPor"
            Caption         =   "Nivel Por"
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
            DataField       =   "cNivel"
            Caption         =   "Nivel"
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
         BeginProperty Column03 
            DataField       =   "nMontoMax"
            Caption         =   "Monto Apr"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cRiesgo"
            Caption         =   "Tipo Riesgo"
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
         BeginProperty Column05 
            DataField       =   ""
            Caption         =   "Motivo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   ""
            Caption         =   "Estado"
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
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMensaje 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   840
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmCredConfNivelAprobacionListar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsmensaje As String

Public Sub BuscarDatos()
    Dim objNivAprob As COMDCredito.DCOMNivelAprobacion
    Set objNivAprob = New COMDCredito.DCOMNivelAprobacion
    Dim rs As Recordset
    lsmensaje = ""
    
    Set rs = objNivAprob.ListarNivelesAprobacion(lsmensaje)
    
    If lsmensaje = "" Then
        lblMensaje.Visible = False
        dgBuscar.Visible = True
        Set dgBuscar.DataSource = rs
        dgBuscar.Refresh
        Screen.MousePointer = 0
    Else
        Set dgBuscar.DataSource = Nothing
        dgBuscar.Refresh
        lblMensaje.Visible = True
        lblMensaje.Caption = "No Existen Datos"
        dgBuscar.Visible = False
    End If
    
    Set rs = Nothing
    Set objNivAprob = Nothing

End Sub

Private Sub Form_Load()
    CentraForm Me
    BuscarDatos
End Sub
