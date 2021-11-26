VERSION 5.00
Begin VB.Form frmCredSeguimientoDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle Seguimiento"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   ControlBox      =   0   'False
   Icon            =   "frmCredSeguimientoDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   8400
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   " Lista de Procesos "
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin SICMACT.FlexEdit feDetalle 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7011
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Ubicación-Fecha/Hora Inicio-Fecha/Hora Fin-Tiempo Transcurrido-% Permanencia"
         EncabezadosAnchos=   "2100-2000-2000-2000-1200"
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
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "Ubicación"
         Enabled         =   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   2100
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL :"
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lbTiempo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label lblPatrimonioActualDesc 
      AutoSize        =   -1  'True
      Caption         =   "Tiempo Trasncuirrido:"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   4680
      Width           =   1530
   End
End
Attribute VB_Name = "frmCredSeguimientoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCredSeguimiento
'*** Descripción : Formulario para realizar el seguimiento de los creditos
'*** Creación : RECO el 20161020, según ERS060-2016
'********************************************************************
Option Explicit

Public Sub Inicia(ByVal psCtaCod As String)
    Call CargarDatos(psCtaCod)
    Me.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CargarDatos(ByVal psCtaCod As String)
    Dim oEvalN As New COMNCredito.NCOMColocEval
    Dim RS, rs1, Rs2 As New ADODB.Recordset 'ARLO 201700717 ARGREGO rs1,rs2
    Dim dFecha, dFecha2 As Date         'ARLO 20170717
    Dim nTotal, nTotalTotal As Integer              'ARLO 20170717
    Dim nPorcentaje As Double           'ARLO 20170717
    Dim nDias As Integer                'ARLO 20170717
    Dim nContador As Integer
    Dim nDifeHoras, nDifeMinutos, nComparar, nHoras, nHorasTotal As Integer

    
    'ARLO 20170717
    Set rs1 = oEvalN.CredDetalleSeguimiento(psCtaCod)
    If Not (rs1.EOF And rs1.BOF) Then
        Dim nIndice As Integer
        
        dFecha = rs1!HoraIni
       
        
        feDetalle.Clear
        FormateaFlex feDetalle
        For nIndice = 1 To rs1.RecordCount
            nTotal = CInt(DateDiff("N", dFecha, rs1!HoraIni)) + nTotal
            dFecha = rs1!HoraIni
            rs1.MoveNext
        Next
    End If
    'ARLO 20170717
    
    
    'Set rs = oEvalN.CredDetalleSeguimiento(psCtaCod) 'COMENTADO POR ARLO 20170717
    

    Set RS = rs1.Clone 'ARLO 20170717
    Set Rs2 = rs1.Clone 'ARLO 20170717
    
    If Not (RS.EOF And RS.BOF) Then
        'Dim nIndice As Integer              'COMENTADO POR ARLO 20170717
        
        nComparar = 0
        For nIndice = 1 To Rs2.RecordCount
        nHoras = CInt(DateDiff("N", Rs2!HoraIni, Rs2!HoraFin))
        If (nHoras > nHorasTotal) Then
        nHorasTotal = nHoras
        End If
        Rs2.MoveNext
        Next
        
        
        feDetalle.Clear
        FormateaFlex feDetalle
            For nIndice = 1 To RS.RecordCount
                'ARLO 20170717
                nPorcentaje = CInt(DateDiff("N", RS!HoraIni, RS!HoraFin))
                nContador = CInt(DateDiff("N", RS!HoraIni, RS!HoraFin))
                If (nTotal = 0 Or nPorcentaje = 0) Then
                nPorcentaje = 0
                Else
                nPorcentaje = Round((CInt(DateDiff("N", RS!HoraIni, RS!HoraFin)) / nTotal) * 100, 2)
                End If
                'ARLO 20170717
                nDifeMinutos = CInt(DateDiff("N", RS!HoraIni, RS!HoraFin))
                nDifeHoras = (-Int(nDifeMinutos / 60) * (-1))
                nDias = (-Int(nDifeHoras / 24) * (-1))
                feDetalle.AdicionaFila
                feDetalle.TextMatrix(nIndice, 0) = RS!ubicacion
                feDetalle.TextMatrix(nIndice, 1) = RS!HoraIni
                feDetalle.TextMatrix(nIndice, 2) = RS!HoraFin
    '            feDetalle.TextMatrix(nIndice, 3) = rs!TiempoTra    'COMENTADO POR ARLO 20170717
    '            feDetalle.TextMatrix(nIndice, 4) = rs!Porc         'COMENTADO POR ARLO 20170717
                feDetalle.TextMatrix(nIndice, 3) = IIf(nContador >= 1440, _
                                                CStr(CStr(nDias) + "    Días") _
                                                + " " + IIf((nContador - (1440 * nDias) >= 60), CStr(CStr(nContador - (1440 * nDias)) + "    Horas"), CStr(CStr(nContador - (1440 * nDias)) + "    Minutos")), _
                                                IIf(nContador >= 60, CStr(CStr(nDifeHoras) + "    Horas") + " " + CStr(CStr(nDifeMinutos - (nDifeHoras * 60)) + "    Minutos"), _
                                                CStr(CStr(nDifeMinutos) + "    Minutos")))
                feDetalle.TextMatrix(nIndice, 4) = nPorcentaje
                If (nHorasTotal = nDifeMinutos) Then
                feDetalle.BackColorRow (&HC0C0FF)
                End If
                RS.MoveNext
            Next
    End If
    
    
    nTotalTotal = nTotal / 60 '(-Int(nTotal / 60) * (-1))
    nDias = (-Int(nTotalTotal / 24) * (-1))
    'Me.lblCumpliento.Caption = "100 %"
    
    
    Me.lbTiempo.Caption = IIf(nTotalTotal >= 24, CStr(CStr(nDias) + "    Días") _
                                            + " " + CStr(CStr(nTotalTotal - (24 * nDias)) + "    Horas"), _
                                            CStr(CStr(nTotalTotal) + "    Horas"))

 
End Sub

