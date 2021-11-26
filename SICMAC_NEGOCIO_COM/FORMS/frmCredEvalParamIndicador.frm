VERSION 5.00
Begin VB.Form frmCredEvalParamIndicador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de Indicadores de Evaluación"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   Icon            =   "frmCredEvalParamIndicador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   1215
         Begin VB.Label Label2 
            Caption         =   "Formato"
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
            Left            =   240
            TabIndex        =   12
            Top             =   90
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   3255
         Begin VB.Label Label1 
            Caption         =   "Indicador"
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
            Left            =   1170
            TabIndex        =   10
            Top             =   90
            Width           =   975
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   3760
         TabIndex        =   6
         Text            =   "Formato"
         Top             =   450
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   450
         TabIndex        =   5
         Text            =   "Indicador"
         Top             =   480
         Width           =   3270
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   150
         TabIndex        =   8
         Top             =   480
         Width           =   280
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   9170
         TabIndex        =   7
         Top             =   310
         Width           =   290
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8280
         TabIndex        =   2
         Top             =   3240
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7080
         TabIndex        =   1
         Top             =   3240
         Width           =   1170
      End
      Begin SICMACT.FlexEdit fgIndicador 
         Height          =   2535
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4471
         Cols0           =   9
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "-Indicador-Formato-Condicion-% Min-% Max-% Min-% Max-Aux"
         EncabezadosAnchos=   "300-3300-1400-0-1000-1000-1000-1000-0"
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
         ColumnasAEditar =   "X-X-X-X-4-5-6-7-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-R-R-R-R-L"
         FormatosEdit    =   "0-0-0-0-2-2-2-2-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit FlexEdit1 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1931
         Cols0           =   5
         ScrollBars      =   0
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "---Aceptable-Crítico"
         EncabezadosAnchos=   "300-3300-1400-2010-2010"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCredEvalParamIndicador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredEvalParamIndicador
'** Descripción : Administración de parametros de los indicadores de evaluación de créditos creado
'**               segun RFC090-2012
'** Creación : JUEZ, 20120903 09:00:00 AM
'**********************************************************************************************

Option Explicit

Private Sub CmdGrabar_Click()
    Dim oNCred As COMDCredito.DCOMCredActBD
    Dim i As Integer
    Dim nId As String
    Set oNCred = New COMDCredito.DCOMCredActBD
    
    If ValidaGrilla Then
        If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            
            Call oNCred.dEliminaCredEvalParametrosIndicador
            For i = 1 To fgIndicador.Rows - 1
                If i < 10 Then
                    nId = "00" & CStr(i)
                ElseIf i < 100 Then
                    nId = "0" & CStr(i)
                Else
                    nId = CStr(i)
                End If
                nId = "IND" & nId
                Call oNCred.dInsertaCredEvalParametrosIndicador(nId, Trim(fgIndicador.TextMatrix(i, 1)), CInt(Trim(Right(fgIndicador.TextMatrix(i, 2), 2))), CInt(fgIndicador.TextMatrix(i, 3)), CDbl(Format(fgIndicador.TextMatrix(i, 4), "#,##0.00")), CDbl(Format(fgIndicador.TextMatrix(i, 5), "#,##0.00")), CDbl(Format(fgIndicador.TextMatrix(i, 6), "#,##0.00")), CDbl(Format(fgIndicador.TextMatrix(i, 7), "#,##0.00")))
            Next i
            MsgBox "Los parámetros se grabaron con exito", vbInformation, "Aviso"
            'Call Form_Load
    Else
        MsgBox "Faltan datos en la lista de parametros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim oDCred As COMDCredito.DCOMCredito
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Set oDCred = New COMDCredito.DCOMCredito
    
    Set rs = oDCred.ListaCredEvalParamIndicadores
    
    If Not (rs.BOF And rs.EOF) Then
        fgIndicador.lbEditarFlex = True
        Call LimpiaFlex(fgIndicador)
            For i = 0 To rs.RecordCount - 1
                fgIndicador.AdicionaFila
                fgIndicador.TextMatrix(i + 1, 0) = i + 1
                fgIndicador.TextMatrix(i + 1, 1) = rs!cIndicadorDesc
                fgIndicador.TextMatrix(i + 1, 2) = rs!cIndicadorFormato
                fgIndicador.TextMatrix(i + 1, 3) = rs!cIndicadorCondicion
                fgIndicador.TextMatrix(i + 1, 4) = Format(rs!cIndicadorAcepMin, "#,##0.00")
                fgIndicador.TextMatrix(i + 1, 5) = Format(rs!cIndicadorAcepMax, "#,##0.00")
                fgIndicador.TextMatrix(i + 1, 6) = Format(rs!cIndicadorCritMin, "#,##0.00")
                fgIndicador.TextMatrix(i + 1, 7) = Format(rs!cIndicadorCritMax, "#,##0.00")
                rs.MoveNext
            Next i
    End If
End Sub

Private Function ValidaGrilla() As Boolean
    Dim i As Integer
    
    ValidaGrilla = False
    For i = 1 To fgIndicador.Rows - 1
        If fgIndicador.TextMatrix(i, 0) <> "" Then
            If Trim(fgIndicador.TextMatrix(i, 1)) = "" Or Trim(fgIndicador.TextMatrix(i, 2)) = "" Or _
               Trim(fgIndicador.TextMatrix(i, 3)) = "" Or Trim(fgIndicador.TextMatrix(i, 4)) = "" Or _
               Trim(fgIndicador.TextMatrix(i, 5)) = "" Or Trim(fgIndicador.TextMatrix(i, 6)) = "" Or _
               Trim(fgIndicador.TextMatrix(i, 7)) = "" Then
                ValidaGrilla = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrilla = True
End Function
