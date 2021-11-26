VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogBieSerMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Bienes/Servicios"
   ClientHeight    =   6255
   ClientLeft      =   960
   ClientTop       =   1320
   ClientWidth     =   8550
   Icon            =   "frmLogBieSerMant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBS 
      Caption         =   "Imprimir"
      Height          =   360
      Index           =   3
      Left            =   6885
      TabIndex        =   7
      Top             =   5235
      Width           =   1350
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "Eliminar"
      Height          =   360
      Index           =   2
      Left            =   6870
      TabIndex        =   6
      Top             =   4845
      Width           =   1350
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "Modificar"
      Height          =   360
      Index           =   1
      Left            =   6870
      TabIndex        =   4
      Top             =   4455
      Width           =   1350
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "Agregar"
      Height          =   360
      Index           =   0
      Left            =   6870
      TabIndex        =   3
      Top             =   4065
      Width           =   1350
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   6870
      TabIndex        =   0
      Top             =   5640
      Width           =   1350
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBS 
      Height          =   3555
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6271
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBSDet 
      Height          =   1980
      Left            =   180
      TabIndex        =   5
      Top             =   4050
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   3493
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Bienes/Servicios"
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
      Height          =   240
      Index           =   0
      Left            =   285
      TabIndex        =   2
      Top             =   105
      Width           =   1560
   End
End
Attribute VB_Name = "frmLogBieSerMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDBS As DLogBieSer

Private Sub cmdBS_Click(Index As Integer)
    Dim rsBSDet As ADODB.Recordset
    Dim sBSCod As String, sBSCodDet As String, sBSNomDet As String
    Dim nResult As Integer
    Select Case Index
        Case 0:
            'Agregar
            sBSCod = fgBS.TextMatrix(fgBS.Row, 0)
            'nResult = frmLogBieSerMantIngreso.Inicio("1", sBSCod)
            nResult = frmLogMantOpc.Inicio("1", "1", sBSCod)
            If nResult = 0 Then
                Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
                If rsBSDet.RecordCount > 0 Then
                    Set fgBSDet.Recordset = rsBSDet
                Else
                    fgBSDet.Rows = 2
                    fgBSDet.Clear
                End If
            End If
        Case 1:
            'Modificar
            sBSCod = fgBS.TextMatrix(fgBS.Row, 0)
            sBSCodDet = fgBSDet.TextMatrix(fgBSDet.Row, 0)
            sBSNomDet = fgBSDet.TextMatrix(fgBSDet.Row, 2)
            If sBSCodDet = "" Then
                MsgBox "No existe bien/servicio", vbInformation, " Aviso"
                Exit Sub
            End If
            'nResult = º.Inicio("2", sBSCodDet)
            nResult = frmLogMantOpc.Inicio("1", "2", sBSCodDet)
            If nResult = 0 Then
                Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
                If rsBSDet.RecordCount > 0 Then
                    Set fgBSDet.Recordset = rsBSDet
                Else
                    fgBSDet.Rows = 2
                    fgBSDet.Clear
                End If
            End If
        Case 2:
            'Eliminar
            sBSCod = fgBS.TextMatrix(fgBS.Row, 0)
            sBSCodDet = fgBSDet.TextMatrix(fgBSDet.Row, 0)
            sBSNomDet = fgBSDet.TextMatrix(fgBSDet.Row, 2)
            
            If sBSCodDet = "" Then
                MsgBox "No existe bien/servicio", vbInformation, " Aviso"
                Exit Sub
            End If
            
            If MsgBox("¿ Estás seguro de eliminar " & sBSNomDet & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                nResult = clsDBS.EliminaBS(sBSCodDet)
                If nResult = 0 Then
                    Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
                    If rsBSDet.RecordCount > 0 Then
                        Set fgBSDet.Recordset = rsBSDet
                    Else
                        fgBSDet.Rows = 2
                        fgBSDet.Clear
                    End If
                Else
                    MsgBox "No se terminó la operación", vbInformation, " Aviso "
                End If
            End If
        Case 3:
            'IMPRIMIR
            Dim clsNImp As NLogImpre
            Dim clsPrevio As clsPrevio
            Dim sImpre As String
            MousePointer = 11
            
            Set clsNImp = New NLogImpre
            sImpre = clsNImp.ImpBS(gsNomAge, gdFecSis)
            Set clsNImp = Nothing
            
            MousePointer = 0
            Set clsPrevio = New clsPrevio
            clsPrevio.Show sImpre, Me.Caption, True
            Set clsPrevio = Nothing
        
        Case Else
            MsgBox "Indice de comand de bien/servicio no reconocido", vbInformation, " Aviso "
    End Select
End Sub

Private Sub cmdSalir_Click()
    Set clsDBS = Nothing
    Unload Me
End Sub


Private Sub fgBS_Click()
    Dim rsBSDet As ADODB.Recordset
    Dim sBSCod As String
    sBSCod = fgBS.TextMatrix(fgBS.Row, 0)
    
    Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
    If rsBSDet.RecordCount > 0 Then
        Set fgBSDet.Recordset = rsBSDet
    Else
        fgBSDet.Rows = 2
        fgBSDet.Clear
    End If
End Sub

Private Sub Form_Load()
    Dim rsBSDet As ADODB.Recordset
    Dim sBSCod As String
    Set clsDBS = New DLogBieSer
    Call CentraForm(Me)
    fgBS.ColWidth(0) = 0
    fgBS.ColWidth(1) = 2000
    fgBS.ColWidth(2) = 5500
    fgBS.ColWidth(3) = 0
    fgBSDet.ColWidth(0) = 0
    fgBSDet.ColWidth(1) = 2000
    fgBSDet.ColWidth(2) = 4000
    
    Set fgBS.Recordset = clsDBS.CargaBS(BsTodosFlex)
    sBSCod = fgBS.TextMatrix(fgBS.Row, 0)
    
    Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
    If rsBSDet.RecordCount > 0 Then
        Set fgBSDet.Recordset = rsBSDet
    Else
        fgBSDet.Rows = 2
        fgBSDet.Clear
    End If
    
    frmMdiMain.staMain.Panels(2).Text = fgBS.Rows & " registros."
End Sub
