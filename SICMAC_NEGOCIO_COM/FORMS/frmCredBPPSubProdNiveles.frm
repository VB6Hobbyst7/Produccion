VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPSubProdNiveles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - SubProductos X Niveles"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "frmCredBPPSubProdNiveles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTNivelesCartera 
      Height          =   6450
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11377
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Niveles de Cartera"
      TabPicture(0)   =   "frmCredBPPSubProdNiveles.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraEditarNivel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGuardar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   5880
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   2
         Top             =   5880
         Width           =   1170
      End
      Begin VB.Frame fraEditarNivel 
         Caption         =   "Editar Nivel"
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
         Height          =   5415
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmCredBPPSubProdNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************************
''***     Rutina         :   frmCredBPPSubProdNiveles
''***     Descripcion    :   Configurar SubProductos x Niveles del BPP
''***     Creado por     :   WIOR
''***     Maquina        :   TIF-1-19
''***     Fecha-Creación :   24/05/2013 08:20:00 AM
''*****************************************************************************************
'Option Explicit
'
'Private Sub cmdCancelar_Click()
'Unload Me
'End Sub
'
'Private Sub cmdGuardar_Click()
'If MsgBox("Estas seguro de Guardar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    Dim i As Integer
'    Dim oBPP As COMDCredito.DCOMBPPR
'    Set oBPP = New COMDCredito.DCOMBPPR
'
'    For i = 0 To feSubProductos.Rows - 2
'        If CInt(Trim(Right(feSubProductos.TextMatrix(i + 1, 2), 4))) <> CInt(Trim(feSubProductos.TextMatrix(i + 1, 5))) Then
'            Call oBPP.OpeSubProductosCredNiveles(CInt(feSubProductos.TextMatrix(i + 1, 4)), feSubProductos.TextMatrix(i + 1, 3), CInt(Trim(Right(feSubProductos.TextMatrix(i + 1, 2), 4))), 1)
'        End If
'    Next i
'    Call CargaControles
'    MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'End If
'End Sub
'
'Private Sub Form_Load()
'Call CargaControles
'End Sub
'
'Private Sub CargaControles()
'Dim oCredito As COMDCredito.DCOMCredito
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim oBPP As COMDCredito.DCOMBPPR
'
'Dim rsSubProd As ADODB.Recordset
'Dim rsBPP As ADODB.Recordset
'Dim rsConst As ADODB.Recordset
'
'Dim i As Integer
'Dim sNivelDefault As String
'Call LimpiaFlex(feSubProductos)
'
'
'Set oBPP = New COMDCredito.DCOMBPPR
'
''CARGA NIVELES
'Set oConst = New COMDConstantes.DCOMConstantes
'Set rsConst = oConst.RecuperaConstantes(7064)
'sNivelDefault = Trim(rsConst!cConsDescripcion) & Space(75) & Trim(rsConst!nConsValor)
'feSubProductos.CargaCombo rsConst
'
''CARGA LOS SUB PRODUCTOS
'Set oCredito = New COMDCredito.DCOMCredito
'Set rsSubProd = oCredito.RecuperaSubProductosCrediticios
'
'If Not (rsSubProd.EOF And rsSubProd.BOF) Then
'    For i = 0 To rsSubProd.RecordCount - 1
'        feSubProductos.AdicionaFila
'        feSubProductos.TextMatrix(i + 1, 0) = i + 1
'        feSubProductos.TextMatrix(i + 1, 1) = UCase(Trim(rsSubProd!cConsDescripcion))
'        feSubProductos.TextMatrix(i + 1, 3) = UCase(Trim(rsSubProd!nConsValor))
'
'        Set rsBPP = oBPP.ObtenerNivelesXSubProdCred(UCase(Trim(rsSubProd!nConsValor)), 1)
'        If Not (rsBPP.BOF And rsBPP.EOF) Then
'            feSubProductos.TextMatrix(i + 1, 2) = Trim(rsBPP!NivelDesc) & Space(75) & Trim(rsBPP!nNivel)
'            feSubProductos.TextMatrix(i + 1, 4) = "2"
'            feSubProductos.TextMatrix(i + 1, 5) = Trim(rsBPP!nNivel)
'        Else
'            feSubProductos.TextMatrix(i + 1, 2) = sNivelDefault
'            feSubProductos.TextMatrix(i + 1, 4) = "1"
'            feSubProductos.TextMatrix(i + 1, 5) = "0"
'        End If
'
'        rsSubProd.MoveNext
'    Next i
'Else
'    MsgBox "No Hay Datos.", vbInformation, "Aviso"
'End If
'feSubProductos.TopRow = 1
'
'End Sub
'
'
