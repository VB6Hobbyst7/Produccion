VERSION 5.00
Begin VB.Form frmDJSujetosObligados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DJ Sujetos Obligados"
   ClientHeight    =   6495
   ClientLeft      =   3420
   ClientTop       =   2445
   ClientWidth     =   9585
   Icon            =   "frmDJSujetosObligados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9585
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   8160
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sectores CIIU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.OptionButton optCIIU 
         Caption         =   "Ninguno"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optCIIU 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin SICMACT.FlexEdit fgCIIU 
         Height          =   5445
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   9604
         Cols0           =   4
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-CodCIIU-Sectores CIIU-OK"
         EncabezadosAnchos=   "500-0-5500-500"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmDJSujetosObligados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()
  limpiarflex
  CargarSectores
End Sub
Private Sub limpiarflex()
    Dim i As Integer
    For i = 1 To fgCIIU.Rows - 1
         Me.fgCIIU.TextMatrix(i, 3) = 0
    Next i
End Sub
Private Sub cmdGuardar_Click()
    Dim i As Integer
    Dim nContChk As Integer
    nContChk = 0
    For i = 1 To fgCIIU.Rows - 1
        If Me.fgCIIU.TextMatrix(i, 3) = "." Then
            nContChk = 1
            Exit For
        End If
    Next i
    If nContChk = 0 Then
        MsgBox "Debe Seleccionar al menos un sector", vbInformation, "AVISO"
    End If
    
    If MsgBox("Se van a guardar los Datos...", vbYesNo, "Aviso") = vbYes Then
         Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
         Dim clsMov As COMNContabilidad.NCOMContFunciones
         Dim sMovNro As String
         
         Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
         Set clsMov = New COMNContabilidad.NCOMContFunciones
         
         sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
         clsServ.modificarSectoresDJSujetosObligados
         For i = 1 To fgCIIU.Rows - 1
            If Me.fgCIIU.TextMatrix(i, 3) = "." Then
                clsServ.guardarSectoresDJSujetosObligados Me.fgCIIU.TextMatrix(i, 1), sMovNro
            End If
         Next i
         MsgBox "Se han guardado los datos", vbInformation, "AVISO"
         Me.cmdSalir.SetFocus
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LlenarFlexSectoresCIIU
    CargarSectores
End Sub
Private Sub LlenarFlexSectoresCIIU()
   
    Dim oPersona As COMDPersona.DCOMPersonas
    Dim rsCIIU As Recordset
    Dim i As Integer
    Set oPersona = New COMDPersona.DCOMPersonas
    Set rsCIIU = New Recordset
    
    Set rsCIIU = oPersona.Cargar_CIIU("0")
    If Not (rsCIIU.EOF And rsCIIU.BOF) Then
        fgCIIU.Rows = rsCIIU.RecordCount + 1
        For i = 1 To rsCIIU.RecordCount
            Me.fgCIIU.TextMatrix(i, 0) = i
            Me.fgCIIU.TextMatrix(i, 1) = rsCIIU!cCIIUcod
            Me.fgCIIU.TextMatrix(i, 2) = rsCIIU!cCIIUdescripcion
            Me.fgCIIU.TextMatrix(i, 3) = 0
            rsCIIU.MoveNext
        Next i
    End If
End Sub
Private Sub CargarSectores()
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rsCIIU As Recordset
    Dim rsCIIU2 As Recordset
   
    Dim i As Integer
    Dim j As Integer
    
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsCIIU = New Recordset
    Set rsCIIU2 = New Recordset
    
    Set rsCIIU = clsServ.obtenerSectoresDJSujetosObligados
    If Not (rsCIIU.EOF And rsCIIU.BOF) Then
        'Set rsCIIU2 = Me.fgCIIU.rsFlex.
        For i = 1 To fgCIIU.Rows - 1
           For j = 0 To rsCIIU.RecordCount - 1
             If Me.fgCIIU.TextMatrix(i, 1) = rsCIIU!cCIIUcod Then
                    Me.fgCIIU.TextMatrix(i, 3) = 1
                    Exit For
             End If
             rsCIIU.MoveNext
           Next j
           rsCIIU.MoveFirst
        Next i
    End If
    
    'Set rsBuscarCIIU = Nothing
    Set rsCIIU = Nothing
    
End Sub
Private Sub optCIIU_Click(Index As Integer)
    Dim i As Integer
    If Index = 0 Then
        
        For i = 1 To fgCIIU.Rows - 1
            fgCIIU.TextMatrix(i, 3) = 1
        Next i
    Else
        For i = 1 To fgCIIU.Rows - 1
            fgCIIU.TextMatrix(i, 3) = 0
        Next i
    
    End If
End Sub
