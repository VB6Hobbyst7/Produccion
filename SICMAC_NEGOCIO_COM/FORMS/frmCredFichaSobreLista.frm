VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCredFichaSobreLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Clientes Sobreendeudados"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmCredFichaSobreLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin MSDataGridLib.DataGrid DGFcihaSobre 
         Height          =   2475
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   4366
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "cPersCod"
            Caption         =   "CODIGO"
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
            DataField       =   "cPersNombre"
            Caption         =   "NOMBRE"
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
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               DividerStyle    =   1
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4289.953
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCredFichaSobreLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGRFicha As ADODB.Recordset
Dim vsSelecPers As String

Public Function Inicio() As String
   Dim oCredFicha As COMDCredito.DCOMCredito
  ' Dim rsGRFicha As ADODB.Recordset

   Set oCredFicha = New COMDCredito.DCOMCredito
    Set rsGRFicha = oCredFicha.ListarFichaSobreEnd()
   
   Set DGFcihaSobre.DataSource = rsGRFicha
   DGFcihaSobre.Refresh
   
   Me.Show 1
   Set rsGRFicha = Nothing
    
   Inicio = vsSelecPers
   
End Function

Private Sub cmdAceptar_Click()
    
    If rsGRFicha.RecordCount > 0 Then
        vsSelecPers = rsGRFicha.Fields(0)
    End If
    Unload Me
    
End Sub

Private Sub cmdsalir_Click()
    vsSelecPers = ""
    Unload Me
End Sub

Private Sub DGFcihaSobre_KeyPress(KeyAscii As Integer)
Dim rs As ADODB.Recordset
Dim nPos As Integer

    Set rs = rsGRFicha.Clone

    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
        rsGRFicha.MoveFirst
    End If

    nPos = 0

    Do Until rs.EOF
        nPos = nPos + 1
        If Mid(rs!cPersNombre, 1, 1) = UCase(Chr(KeyAscii)) Then
            rsGRFicha.Move nPos - 1
            Exit Do
        End If
        rs.MoveNext
    Loop

    Set rs = Nothing
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub
