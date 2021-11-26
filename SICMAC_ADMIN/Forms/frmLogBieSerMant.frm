VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogBieSerMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Bienes/Servicios"
   ClientHeight    =   6765
   ClientLeft      =   990
   ClientTop       =   1200
   ClientWidth     =   8550
   Icon            =   "frmLogBieSerMant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   360
      Left            =   90
      TabIndex        =   6
      Top             =   6330
      Width           =   1080
   End
   Begin VB.CommandButton cmdBuscarSig 
      Caption         =   "Buscar S&iguiente >>>"
      Height          =   360
      Left            =   1215
      TabIndex        =   5
      Top             =   6330
      Width           =   2085
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Bienes / Servicios "
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
      Height          =   6225
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   8460
      Begin TabDlg.SSTab SSTab1 
         Height          =   2010
         Left            =   105
         TabIndex        =   7
         Top             =   4125
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   3545
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         ForeColor       =   8388608
         TabCaption(0)   =   "Detalle"
         TabPicture(0)   =   "frmLogBieSerMant.frx":08CA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdBS(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdBS(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fgBSDet"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdCancelar"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdAceptar"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdBS(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Proveedores Mismo Producto"
         TabPicture(1)   =   "frmLogBieSerMant.frx":08E6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fgBSProv"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Productos Similares"
         TabPicture(2)   =   "frmLogBieSerMant.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fgBSProvNS"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.CommandButton cmdBS 
            Caption         =   "Eliminar"
            Height          =   330
            Index           =   2
            Left            =   -68130
            TabIndex        =   10
            Top             =   1275
            Width           =   1230
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   330
            Left            =   -68130
            TabIndex        =   9
            Top             =   540
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Modificar"
            Height          =   330
            Left            =   -68130
            TabIndex        =   8
            Top             =   915
            Visible         =   0   'False
            Width           =   1230
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBSDet 
            Height          =   1515
            Left            =   -74910
            TabIndex        =   11
            Top             =   405
            Width           =   6720
            _ExtentX        =   11853
            _ExtentY        =   2672
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            TextStyleFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
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
         Begin VB.CommandButton cmdBS 
            Caption         =   "Agregar"
            Height          =   330
            Index           =   0
            Left            =   -68130
            TabIndex        =   12
            Top             =   540
            Width           =   1230
         End
         Begin VB.CommandButton cmdBS 
            Caption         =   "Modificar"
            Height          =   330
            Index           =   1
            Left            =   -68130
            TabIndex        =   13
            Top             =   900
            Width           =   1230
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBSProv 
            Height          =   1515
            Left            =   90
            TabIndex        =   14
            Top             =   405
            Width           =   8070
            _ExtentX        =   14235
            _ExtentY        =   2672
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            TextStyleFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBSProvNS 
            Height          =   1515
            Left            =   -74910
            TabIndex        =   15
            Top             =   405
            Width           =   8070
            _ExtentX        =   14235
            _ExtentY        =   2672
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            TextStyleFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
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
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBS 
         Height          =   3555
         Left            =   75
         TabIndex        =   3
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6271
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         Index           =   1
         X1              =   8340
         X2              =   165
         Y1              =   3855
         Y2              =   3855
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   8325
         X2              =   150
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Detalle de Item seleccionado"
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
         Left            =   120
         TabIndex        =   4
         Top             =   3885
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "Imprimir"
      Height          =   360
      Index           =   3
      Left            =   6045
      TabIndex        =   1
      Top             =   6330
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   7305
      TabIndex        =   0
      Top             =   6330
      Width           =   1200
   End
End
Attribute VB_Name = "frmLogBieSerMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDBS As DLogBieSer
Dim clsProveedor As DLogProveedor
Dim lsCodigo As String
Dim lsTipo As String

Dim lsCad As String
Dim lnI As Integer

Public Function Inicio(pbBienes As Boolean) As String
    Dim I As Integer
    
    If pbBienes Then
        lsTipo = "1"
    Else
        lsTipo = "2"
    End If
    
    Me.cmdAceptar.Enabled = True
    Me.cmdAceptar.Visible = True
    Me.cmdCancelar.Enabled = True
    Me.cmdCancelar.Visible = True
    
    For I = 0 To 3
        Me.cmdBS(I).Enabled = False
        Me.cmdBS(I).Visible = False
    Next I
    
    Me.Show 1
    Inicio = lsCodigo
End Function

Private Sub cmdAceptar_Click()
    Dim oBS As DLogBieSer
    Set oBS = New DLogBieSer
    
    If oBS.EsUltimoNivel(Trim(Me.fgBSDet.TextMatrix(Me.fgBSDet.Row, 1))) Then
        lsCodigo = Trim(Me.fgBSDet.TextMatrix(Me.fgBSDet.Row, 1))
        Set oBS = Nothing
        Unload Me
    Else
        MsgBox "Debe elegir un Bien o Servicio de último Nivel.", vbInformation, "Aviso"
        Set oBS = Nothing
    End If
End Sub

Private Sub cmdBS_Click(Index As Integer)
    Dim rsBSDet As ADODB.Recordset
    Dim sBSCod As String, sBSCodDet As String, sBSNomDet As String
    Dim nResult As Integer
    Select Case Index
        Case 0:
            'Agregar
            sBSCod = Trim(fgBS.TextMatrix(fgBS.Row, 0))
            'nResult = frmLogBieSerMantIngreso.Inicio("1", sBSCod)
            nResult = frmLogMantOpc.Inicio("1", "1", sBSCod)
            If nResult = 0 Then
                Set fgBS.Recordset = clsDBS.CargaBS(BsTodosFlex)
    
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
                Set fgBS.Recordset = clsDBS.CargaBS(BsTodosFlex)
                
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
                    Set fgBS.Recordset = clsDBS.CargaBS(BsTodosFlex)
                    
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
            clsPrevio.Show sImpre, Me.Caption, True, , gImpresora
            Set clsPrevio = Nothing
        
        Case Else
            MsgBox "Indice de comand de bien/servicio no reconocido", vbInformation, " Aviso "
    End Select
End Sub

Private Sub cmdBuscar_Click()
    lsCad = UCase(Trim(InputBox("Ingrese Cadena a Buscar.", "Aviso")))
    
    If lsCad = "" Then Exit Sub
    
    For lnI = 1 To Me.fgBS.Rows - 1
        If InStr(1, UCase(fgBS.TextMatrix(lnI, 2)), lsCad) <> 0 Then
            fgBS.TopRow = lnI
            fgBS.Row = lnI
            Exit Sub
        End If
    Next lnI
    
    MsgBox "No se encontro el bien", vbInformation, "Aviso"
End Sub

Private Sub cmdBuscarSig_Click()
    If lsCad = "" Then Exit Sub
    
    For lnI = lnI + 1 To Me.fgBS.Rows - 1
        If InStr(1, UCase(fgBS.TextMatrix(lnI, 2)), lsCad) <> 0 Then
            fgBS.TopRow = lnI
            fgBS.Row = lnI
            Exit Sub
        End If
    Next lnI
    MsgBox "No se encontro el bien", vbInformation, "Aviso"
End Sub

Private Sub cmdCancelar_Click()
    lsCodigo = ""
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Set clsDBS = Nothing
    Set clsProveedor = Nothing
    Unload Me
End Sub

Private Sub fgBS_Click()
    Dim rsBSDet As ADODB.Recordset
    Dim rsProv As ADODB.Recordset
    Dim rsProvNS As ADODB.Recordset
    Dim sBSCod As String
    sBSCod = Trim(fgBS.TextMatrix(fgBS.Row, 0))
    
    Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
    If rsBSDet.RecordCount > 0 Then
        Set fgBSDet.Recordset = rsBSDet
        
        Set rsProv = clsProveedor.CargaProveedorBS(ProBSProveedor, sBSCod, False)
        Set rsProvNS = clsProveedor.CargaProveedorBS(ProBSProveedor, sBSCod, True)
        
        If rsProv.EOF And rsProv.BOF Then
            fgBSProv.Clear
            fgBSProv.Rows = 2
        Else
            Set Me.fgBSProv.Recordset = rsProv
        End If
        
        If rsProvNS.EOF And rsProvNS.BOF Then
            fgBSProvNS.Clear
            fgBSProvNS.Rows = 2
        Else
            Set Me.fgBSProvNS.Recordset = rsProvNS
        End If
        
        fgBSProv.ColWidth(0) = 1500
        fgBSProv.ColWidth(1) = 3000
        fgBSProv.ColWidth(2) = 3000
        fgBSProv.ColWidth(3) = 1000
        fgBSProv.ColWidth(4) = 1000
        fgBSProv.ColWidth(5) = 3000
        
        fgBSProvNS.ColWidth(0) = 1500
        fgBSProvNS.ColWidth(1) = 3000
        fgBSProvNS.ColWidth(2) = 3000
        fgBSProvNS.ColWidth(3) = 1000
        fgBSProvNS.ColWidth(4) = 1000
        fgBSProvNS.ColWidth(5) = 3000
    Else
        fgBSDet.Rows = 2
        fgBSDet.Clear
    End If
End Sub

Private Sub Form_Load()
    Dim rsBS As ADODB.Recordset, rsBSDet As ADODB.Recordset
    Dim sBSCod As String
    Dim I As Integer
    Set clsDBS = New DLogBieSer
    Set clsProveedor = New DLogProveedor
    
    Me.cmdAceptar.Enabled = False
    Me.cmdAceptar.Visible = False
    Me.cmdCancelar.Enabled = False
    Me.cmdCancelar.Visible = False
    
    For I = 0 To 3
        Me.cmdBS(I).Enabled = True
        Me.cmdBS(I).Visible = True
    Next I
    
    fgBS.ColWidth(0) = 0
    fgBS.ColWidth(1) = 2000
    fgBS.ColWidth(2) = 5500
    fgBS.ColWidth(3) = 0
    fgBSDet.ColWidth(0) = 0
    fgBSDet.ColWidth(1) = 2000
    fgBSDet.ColWidth(2) = 4000
    
    Set rsBS = clsDBS.CargaBS(BsTodosFlex, lsTipo)
    If rsBS.RecordCount > 0 Then
        Set fgBS.Recordset = rsBS
    Else
        cmdBS(0).Enabled = False
        cmdBS(1).Enabled = False
        cmdBS(2).Enabled = False
        cmdBS(3).Enabled = False
        Exit Sub
    End If
    sBSCod = fgBS.TextMatrix(fgBS.Row, 0)
    
    Set rsBSDet = clsDBS.CargaBS(BsSuperiorFlex, sBSCod)
    If rsBSDet.RecordCount > 0 Then
        Set fgBSDet.Recordset = rsBSDet
    Else
        fgBSDet.Rows = 2
        fgBSDet.Clear
    End If
    
    MDISicmact.staMain.Panels(2).Text = fgBS.Rows & " registros."
End Sub

