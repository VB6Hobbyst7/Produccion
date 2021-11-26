VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCampanas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Campañas"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "FrmCampanas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   11145
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Campaña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   -30
      TabIndex        =   10
      Top             =   0
      Width           =   11085
      Begin TabDlg.SSTab Stab 
         Height          =   4305
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   7594
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Campañas"
         TabPicture(0)   =   "FrmCampanas.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Asignación de Campañas"
         TabPicture(1)   =   "FrmCampanas.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(1)=   "LblMensaje"
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame4 
            Height          =   3315
            Left            =   -74880
            TabIndex        =   19
            Top             =   570
            Width           =   10695
            Begin VB.CommandButton CmdAll 
               Caption         =   "<="
               Height          =   285
               Left            =   7560
               TabIndex        =   27
               Top             =   1590
               Width           =   375
            End
            Begin VB.Frame Frame5 
               Height          =   675
               Left            =   180
               TabIndex        =   25
               Top             =   2520
               Width           =   10275
               Begin VB.CommandButton CmdSalirAsignacion 
                  Caption         =   "Salir"
                  Height          =   315
                  Left            =   6120
                  TabIndex        =   9
                  Top             =   240
                  Width           =   945
               End
               Begin VB.CommandButton CmdEliminarAsignacion 
                  Caption         =   "Eliminar"
                  Height          =   315
                  Left            =   4980
                  TabIndex        =   8
                  Top             =   240
                  Width           =   945
               End
               Begin VB.CommandButton CmdNuevoAsiignacion 
                  Caption         =   "Nuevo"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   7
                  Top             =   240
                  Width           =   945
               End
               Begin VB.CommandButton CmdGrabarAsignacion 
                  Caption         =   "Grabar"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   28
                  Top             =   240
                  Width           =   945
               End
            End
            Begin VB.CommandButton CmdAgencia 
               Caption         =   "<--"
               Height          =   285
               Left            =   7560
               TabIndex        =   24
               Top             =   1200
               Width           =   375
            End
            Begin VB.ListBox LstAgencias 
               Height          =   2010
               Left            =   8040
               TabIndex        =   6
               Top             =   480
               Width           =   2475
            End
            Begin VB.CommandButton CmdCampana 
               Caption         =   "-->"
               Height          =   285
               Left            =   3240
               TabIndex        =   21
               Top             =   1230
               Width           =   375
            End
            Begin VB.ListBox LstCampanas 
               Height          =   2010
               Left            =   120
               TabIndex        =   5
               Top             =   480
               Width           =   3075
            End
            Begin MSComctlLib.ListView Lst 
               Height          =   2040
               Left            =   3720
               TabIndex        =   32
               Top             =   480
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   3598
               View            =   3
               Arrange         =   2
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Agencia"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Campaña"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "IdAgencia"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "IdCampanas"
                  Object.Width           =   0
               EndProperty
            End
            Begin VB.Label Agenc 
               AutoSize        =   -1  'True
               Caption         =   "Agencias"
               Height          =   195
               Left            =   8130
               TabIndex        =   23
               Top             =   270
               Width           =   660
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Asignación"
               Height          =   195
               Left            =   3810
               TabIndex        =   22
               Top             =   300
               Width           =   780
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Campañas"
               Height          =   195
               Left            =   240
               TabIndex        =   20
               Top             =   270
               Width           =   750
            End
         End
         Begin VB.Frame Frame2 
            Height          =   3675
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   10575
            Begin MSMask.MaskEdBox txtDesde 
               Height          =   300
               Left            =   1500
               TabIndex        =   33
               Top             =   1200
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Frame Frame3 
               Height          =   705
               Left            =   210
               TabIndex        =   18
               Top             =   2910
               Width           =   10275
               Begin VB.CommandButton CmdActualizar 
                  Caption         =   "Actualizar"
                  Height          =   315
                  Left            =   1170
                  TabIndex        =   30
                  Top             =   240
                  Width           =   945
               End
               Begin VB.CommandButton CmdGrabarCampanas 
                  Caption         =   "Grabar"
                  Height          =   315
                  Left            =   90
                  TabIndex        =   26
                  Top             =   240
                  Width           =   945
               End
               Begin VB.CommandButton CmdSalirCampanas 
                  Caption         =   "Salir"
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   4
                  Top             =   240
                  Width           =   945
               End
               Begin VB.CommandButton CmdEliminar 
                  Caption         =   "Eliminar"
                  Height          =   315
                  Left            =   4740
                  TabIndex        =   3
                  Top             =   240
                  Width           =   945
               End
               Begin VB.CommandButton CmdNuevoCampanas 
                  Caption         =   "Nuevo"
                  Height          =   315
                  Left            =   90
                  TabIndex        =   2
                  Top             =   240
                  Width           =   945
               End
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshCampanas 
               Height          =   1185
               Left            =   180
               TabIndex        =   17
               Top             =   1680
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   2090
               _Version        =   393216
               Cols            =   4
               _NumberOfBands  =   1
               _Band(0).Cols   =   4
            End
            Begin VB.TextBox txtDescripcionCampanas 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1500
               MaxLength       =   50
               TabIndex        =   1
               Top             =   720
               Width           =   3945
            End
            Begin VB.TextBox txtCodigo 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   0
               Text            =   "0"
               Top             =   300
               Visible         =   0   'False
               Width           =   1005
            End
            Begin MSMask.MaskEdBox txtHasta 
               Height          =   300
               Left            =   3950
               TabIndex        =   35
               Top             =   1200
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Hasta:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3240
               TabIndex        =   34
               Top             =   1200
               Width           =   585
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Desde:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   735
               TabIndex        =   16
               Top             =   1200
               Width           =   660
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Descripción:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   300
               TabIndex        =   14
               Top             =   780
               Width           =   1125
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Codigo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   720
               TabIndex        =   13
               Top             =   360
               Visible         =   0   'False
               Width           =   705
            End
         End
         Begin VB.Label LblMensaje 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   345
            Left            =   -74850
            TabIndex        =   29
            Top             =   3870
            Width           =   10605
         End
      End
   End
   Begin MSComctlLib.ListView Lstxxx 
      Height          =   720
      Left            =   5880
      TabIndex        =   31
      Top             =   5280
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   1270
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Campaña"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Agencia"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IdAgencia"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IdCampanas"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "FrmCampanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCampanas As Boolean
Dim bAsignacion As Boolean
Dim nEstado As Integer 'ARLO 20171013

Public Enum OpeMant
    lRegistro = 1
    lMantenimiento = 2
    lConsultas = 3
End Enum
Dim nOperacion As OpeMant
Public Sub Registro()
    bCampanas = True
    bAsignacion = True
    ConfiguraControles (lRegistro)
    '** COMENTADO POR ARLO 20171013
    'If cmbEstado.ListCount > 0 Then
    '   cmbEstado.ListIndex = 0
    'End If
    Me.txtDesde = gdFecSis  'BY ARLO 20171013
    Me.txtHasta = gdFecSis  'BY ARLO 20171013
    nOperacion = lRegistro
    Me.Show vbModal
End Sub
Public Sub Mantenimiento()
    bCampanas = True
    bAsignacion = True
    ConfiguraControles (lMantenimiento)
    '** COMENTADO POR ARLO 20171013
    'If cmbEstado.ListCount > 0 Then
    '   cmbEstado.ListIndex = 0
    'End If
    nOperacion = lMantenimiento
    Me.Show vbModal
End Sub

Public Sub Consultas()
    bCampanas = True
    bAsignacion = True
    ConfiguraControles (lConsultas)
    '** COMENTADO POR ARLO 20171013
    'If cmbEstado.ListCount > 0 Then
    '   cmbEstado.ListIndex = 0
    'End If
    nOperacion = lConsultas
    Me.Show vbModal
End Sub

Sub ConfiguraControles(ByVal pOptMant As OpeMant)
    If pOptMant = lRegistro Then
        CmdEliminar.Enabled = False
        CmdNuevoCampanas.Enabled = True
        CmdEliminarAsignacion.Enabled = False
        Call ConfigurarMShCampanas
        Call CargarCampanas
        CargarCampanasAsignacion
        CargarAgencias
        Stab.Tab = 0
        CmdNuevoCampanas.Visible = True
        CmdGrabarCampanas.Visible = False
        CmdActualizar.Enabled = False
        Me.txtCodigo.Enabled = False                'BY ARLO 20171013
        Me.txtDescripcionCampanas.Enabled = False   'BY ARLO 20171013
        Me.txtDesde.Enabled = False                 'BY ARLO 20171013
        Me.txtHasta.Enabled = False                 'BY ARLO 20171013
    ElseIf pOptMant = lMantenimiento Then
        CmdEliminar.Enabled = True
        CmdNuevoCampanas.Enabled = False
        CmdEliminarAsignacion.Enabled = True
        Call ConfigurarMShCampanas
        Call CargarCampanas
        CargarCampanasAsignacion
        CargarAgencias
        Stab.Tab = 0
        CmdNuevoCampanas.Visible = True
        CmdGrabarCampanas.Visible = False
        CmdActualizar.Enabled = True
    Else
        CmdEliminar.Enabled = False
        CmdNuevoCampanas.Enabled = False
        CmdEliminarAsignacion.Enabled = False
        Call ConfigurarMShCampanas
        Call CargarCampanas
        CargarCampanasAsignacion
        CargarAgencias
        Stab.Tab = 0
        CmdNuevoCampanas.Visible = False
        CmdActualizar.Enabled = False
        CmdGrabarCampanas.Visible = False
        Me.txtCodigo.Enabled = False                'BY ARLO 20171013
        Me.txtDescripcionCampanas.Enabled = False   'BY ARLO 20171013
        Me.txtDesde.Enabled = False                 'BY ARLO 20171013
        Me.txtHasta.Enabled = False                 'BY ARLO 20171013
    End If
End Sub

Sub ConfigurarMShCampanas()
    MshCampanas.Clear
    MshCampanas.cols = 5
    MshCampanas.rows = 2
    
    With MshCampanas
        .TextMatrix(0, 0) = "Codigo"
        .TextMatrix(0, 1) = "Descripcion"
        .TextMatrix(0, 2) = "Inicio"    'BY ARLO 20171013
        .TextMatrix(0, 3) = "Fin"       'BY ARLO 20171013
        .TextMatrix(0, 4) = "Estado"
        
        .ColWidth(1) = 4500
        .ColWidth(2) = 1000         'BY ARLO 20171013
        .ColWidth(3) = 1000         'BY ARLO 20171013
        .ColWidth(4) = 1000
    End With
     
End Sub

Private Sub CmdActualizar_Click()
    Dim odCamp As COMDCredito.DCOMCampanas
    nEstado = 1 'ARLO 20171013
        If txtCodigo.Text <> "" And txtDescripcionCampanas.Text <> "" Then
            Set odCamp = New COMDCredito.DCOMCampanas
            'If odCamp.ActualizarCampanas(txtCodigo, txtDescripcionCampanas.Text, IIf(cmbEstado.Text = "Activo", 1, 0)) = True Then 'COMENTADO POR ARLO 20171013
            If odCamp.ActualizarCampanas(txtCodigo.Text, txtDescripcionCampanas, nEstado, Format(txtDesde.Text, "YYYYMMDD"), Format(txtHasta.Text, "YYYYMMDD")) = True Then 'BY ARLO 20171013
                MsgBox "Se ha actualizado correctamente", vbInformation, "AVISO"
                CargarCampanas
                CargarCampAgeAsignacion
            Else
                MsgBox "Error al actualizar", vbInformation, "AVISO"
            End If
            Set odCamp = Nothing
        End If
End Sub

Private Sub cmdAgencia_Click()
    Dim i As Integer
    Dim nPos As Integer
    Dim nLen As Integer
    
    If LstCampanas.ListIndex = -1 Then
        MsgBox "Debe seleccionar primero la campaña", vbInformation, "AVISO"
        Exit Sub
    End If
    
    If LstAgencias.ListIndex <> -1 Then
        For i = 1 To Lst.ListItems.count
            If Lst.ListItems(i).SubItems(1) = "" Then
                nPos = i
                Exit For
            End If
        Next i
        
        If nPos > 0 Then 'marg20170504
            'Asignado el valor de la agencia asinada
            nLen = Len(CStr(LstAgencias.ItemData(LstAgencias.ListIndex)))
            Lst.ListItems(nPos).SubItems(1) = LstAgencias.List(LstAgencias.ListIndex)
            Lst.ListItems(nPos).SubItems(2) = Right("00", 2 - nLen) & LstAgencias.ItemData(LstAgencias.ListIndex)
            
        'marg20170504*********
        Else
           MsgBox "Debe agregar una campaña", vbInformation, "Aviso"
        End If
        'end marg20170504*****
    Else
        MsgBox "No ha seleccionado una agencia", vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdAll_Click()
Dim i As Integer
Dim j As Integer
Dim iTem As ListItem
Dim nPos As Integer

    If LstCampanas.ListIndex = -1 Then
        MsgBox "Debe seleccionar una campaña", vbInformation, "Aviso"
        Exit Sub
    End If


     For i = 1 To Lst.ListItems.count
            If Lst.ListItems(i).SubItems(1) = "" Then
                nPos = i
                Exit For
            End If
     Next i
      
    If nPos > 0 Then 'marg20170504
        'recorriendo todas laa agencias
        For j = 0 To LstAgencias.ListCount - 1
           
           If j = 0 Then
               Lst.ListItems(nPos).SubItems(1) = LstAgencias.List(j)
               Lst.ListItems(nPos).SubItems(2) = Right("00", 2 - Len(CStr(LstAgencias.ItemData(j)))) & CStr(LstAgencias.ItemData(j))
           Else
               Set iTem = Lst.ListItems.Add(, , LstCampanas.List(LstCampanas.ListIndex))
                   iTem.SubItems(1) = LstAgencias.List(j)
                   iTem.SubItems(3) = LstCampanas.ItemData(LstCampanas.ListIndex)
                   iTem.SubItems(2) = Right("00", 2 - Len(CStr(LstAgencias.ItemData(j)))) & CStr(LstAgencias.ItemData(j))
           End If
        Next j
     
     'marg20170504*********
     Else
        MsgBox "Debe agregar una campaña", vbInformation, "Aviso"
     End If
     'end marg20170504*****

End Sub

Private Sub CmdCampana_Click()
    Dim iTem As ListItem
     
     
        If LstCampanas.ListIndex <> -1 Then
                Set iTem = Lst.ListItems.Add(, , LstCampanas.List(LstCampanas.ListIndex))
                iTem.SubItems(3) = LstCampanas.ItemData(LstCampanas.ListIndex)
        Else
           MsgBox "Debe seleccionar una Campaña", vbInformation, "AVISO"
      End If
     
End Sub

Private Sub cmdEliminar_Click()
    Dim oCamp As COMDCredito.DCOMCampanas
    If txtCodigo <> "" And txtDescripcionCampanas <> "" Then
        Set oCamp = New COMDCredito.DCOMCampanas
        If oCamp.VerificacionEliminacion(CInt(txtCodigo)) = True Then
            If oCamp.EliminacionCampanas(CInt(txtCodigo)) = True Then
                MsgBox "Se eliminino correctamente " & Err.Description, vbInformation, "AVISO"
                Set oCamp = Nothing
                CargarCampanas
            Else
                MsgBox "Existe un error en la elimininacion " & Err.Description, vbInformation, "AVISO"
            End If
        Else
            MsgBox "No se puede eliminar existe registros dependiente", vbInformation, "AVISO"
        End If
        Set oCamp = Nothing
    Else
        MsgBox "No ha seleccionado ningun registro", vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdEliminarAsignacion_Click()
    Dim oCamp As COMDCredito.DCOMCampanas
    
    If Not Lst.SelectedItem Is Nothing Then
        Set oCamp = New COMDCredito.DCOMCampanas
        If oCamp.EliminacionCampanasAgencias(Lst.ListItems(Lst.SelectedItem.Index).SubItems(3), Lst.ListItems(Lst.SelectedItem.Index).SubItems(2)) = True Then
            MsgBox "Se elimino correctamente", vbInformation, "AVISO"
            CargarCampAgeAsignacion
        Else
            MsgBox "Error al momento de eliminar el registro", vbInformation, "AVISO"
        End If
        Set oCamp = Nothing
    Else
        MsgBox "Debe seleccionar un registro", vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdGrabarAsignacion_Click()
    'Dim nAsignacion As Integer
    Dim oCamp As COMDCredito.DCOMCampanas
    Dim bResultado As Boolean
    Dim i As Integer
    'Dim dFecha As String
    Dim MatListaInsercion As Variant
    Dim sMensaje As String
    
    ReDim MatListaInsercion(Lst.ListItems.count, 2)
    
    For i = 0 To Lst.ListItems.count - 1
        MatListaInsercion(i, 0) = Lst.ListItems(i + 1).SubItems(2)
        MatListaInsercion(i, 1) = Lst.ListItems(i + 1).SubItems(3)
    Next i
    
    Set oCamp = New COMDCredito.DCOMCampanas
    bResultado = oCamp.GrabarAsignacionCampanas(gdFecSis, MatListaInsercion, sMensaje)
    Set oCamp = Nothing
    
    'dFecha = Format(gdFecSis, "MM/dd/yyyy")
    'nAsignacion = fn_VerificarPreviewInsercion
    'If nAsignacion = -1 Then
    '    MsgBox "Usted no ha asignado ninguna campaña a una agencia", vbInformation, "AVISO"
    'Else
    '    Set oCamp = New COMDCredito.DCOMCampanas
'       nAsignacion = fn_VerificarPreviewInsercion ' Posicion en que se  empieza la insercion
    '    For i = nAsignacion To Lst.ListItems.Count
    '        bResultado = oCamp.InsertarAsignacion(Lst.ListItems(i).SubItems(2), Lst.ListItems(i).SubItems(3), dFecha)
    '        If bResultado = False Then
    '            MsgBox "Existe error en la insercción", vbInformation, "AVISO"
    '        End If
    '    Next i
    '    Set oCamp = Nothing
    'End If
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
    End If
    If bResultado = True Then
        MsgBox "Se guardo de forma correcta", vbInformation, "AVISO"
        CargarCampAgeAsignacion
    End If
End Sub

Private Sub CmdGrabarCampanas_Click()
    Dim sMensaje As String
    Dim oDCampana As COMDCredito.DCOMCampanas
    nEstado = 1 'ARLO 20171013
    
    If ValidaDatos = "" Then
        Set oDCampana = New COMDCredito.DCOMCampanas
        'If oDCampana.InsertarCampanas(txtCodigo.Text, txtDescripcionCampanas, IIf(CmbEstado.Text = "Activo", 1, 0)) = True Then 'COMENTADO POR ARLO20171013
        If oDCampana.InsertarCampanas(txtCodigo.Text, txtDescripcionCampanas, nEstado, Format(txtDesde.Text, "YYYYMMDD"), Format(txtHasta.Text, "YYYYMMDD")) = True Then 'ARLO20171013
            MsgBox "Se registro correctamente", vbInformation, "AVISO"
            CargarCampanas
            CargarCampanasAsignacion
            CargarAgencias
            txtCodigo.Text = "0"
            txtDescripcionCampanas.Text = ""
            'CmbEstado.ListIndex = 0    'BY ARLO 20171013
            CmdNuevoCampanas.Visible = True
            CmdGrabarCampanas.Visible = False
        Else
            MsgBox "Error al registro", vbInformation, "AVISO"
            Exit Sub
        End If
        Set oDCampana = Nothing
    Else
        sMensaje = ValidaDatos
        MsgBox sMensaje, vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdNuevoAsiignacion_Click()
    CmdCampana.Enabled = True
    CmdAgencia.Enabled = True
    CmdAll.Enabled = True
    CmdGrabarAsignacion.Visible = True
    CmdNuevoAsiignacion.Visible = False
End Sub

Private Sub CmdNuevoCampanas_Click()
    Dim oNCampana As COMDCredito.DCOMCampanas
    txtCodigo.Text = "0"
    txtDescripcionCampanas = ""
    '**COMENTADO POR ARLO 20171013
    'If cmbEstado.ListCount > 0 Then
    '   cmbEstado.ListIndex = 0
    'End If
    
    CmdNuevoCampanas.Visible = False
    CmdGrabarCampanas.Visible = True
    Me.txtCodigo.Enabled = True                 'BY ARLO 20171013
    Me.txtDescripcionCampanas.Enabled = True    'BY ARLO 20171013
    Me.txtDesde.Enabled = True                  'BY ARLO 20171013
    Me.txtHasta.Enabled = True                  'BY ARLO 20171013
    
    txtDescripcionCampanas.SetFocus
    
    '** Comentado por DAOR 20071102
    'Set oNCampana = New COMDCredito.DCOMCampanas
    'txtCodigo.Text = oNCampana.GeneraIdCampanas
    'Set oNCampana = Nothing
End Sub

Private Sub CmdSalirAsignacion_Click()
    If bCampanas = True And bAsignacion = True Then
            Unload Me
    Else
            MsgBox "Guarde los cambios hechos", vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdSalirCampanas_Click()
    If bCampanas = True And bAsignacion = True Then
        Unload Me
    Else
        MsgBox "Guarde los cambios hechos", vbInformation, "AVISO"
    End If
End Sub

Private Sub Form_Load()
    ConfigurarMShCampanas
End Sub

Sub CargarCampanas()
    Dim rs As ADODB.Recordset
    Dim oCampanas As COMDCredito.DCOMCampanas
    
    ConfigurarMShCampanas
    
    txtDescripcionCampanas.Text = "" 'BY ARLO 20171013
    txtDesde = gdFecSis  'BY ARLO 20171013
    txtHasta = gdFecSis  'BY ARLO 20171013
    
    
    Set oCampanas = New COMDCredito.DCOMCampanas
    Set rs = oCampanas.CargarCampanas
    Set oCampanas = Nothing
    
    Do Until rs.EOF
        With MshCampanas
            .rows = .rows + 1
            .TextMatrix(.rows - 2, 0) = rs!idCampana
            .TextMatrix(.rows - 2, 1) = rs!cDescripcion
            .TextMatrix(.rows - 2, 2) = rs!dfechaini    'ARLO20171013 ERS062-2017
            .TextMatrix(.rows - 2, 3) = rs!dfechafin    'ARLO20171013 ERS062-2017
            If rs!bEstado = True Then
                .TextMatrix(.rows - 2, 4) = "Activo"
            Else
                .TextMatrix(.rows - 2, 4) = "Inactivo"
            End If
        End With
        rs.MoveNext
    Loop
End Sub

Function ValidaDatos() As String
    If txtDescripcionCampanas.Text = "" Then
        ValidaDatos = "No existe descripción para la campaña"
    End If
    
    '**COMENTADO POR ARLO 20171013
    'If cmbEstado.ListIndex = -1 Then
    '    ValidaDatos = "No ha seleccionado algun estado"
    'End If
    
    'ARLO20171013 ERS062-2017 ***INICIO
    
    If CDate(txtDesde.Text) > CDate(txtHasta.Text) Then
        'MsgBox "Fecha de inicio mayor que fecha final, corregir", vbInformation, "MENSAJE DEL SISTEMA"
        ValidaDatos = "Fecha de inicio mayor que fecha final, por favor de corregir"
        txtHasta.SetFocus
    End If
    'ARLO20171013 ERS062-2017 ***FIN
    
End Function

Private Sub LstAgencias_DblClick()
    Dim odCamp As COMDCredito.DCOMCampanas
    Dim rs As ADODB.Recordset
    Dim iTem As ListItem
    Dim sAgenciaCod As String
    
    sAgenciaCod = Right("00", 2 - Len(CStr(LstAgencias.ItemData(LstAgencias.ListIndex)))) & CStr(LstAgencias.ItemData(LstAgencias.ListIndex))
    
    If LstCampanas.ListIndex <> -1 Then
        Set odCamp = New COMDCredito.DCOMCampanas
        Set rs = odCamp.ListaCampanasAgenciacXAgencia(sAgenciaCod)
        Set odCamp = Nothing
        
        Lst.ListItems.Clear
        Do Until rs.EOF
            Set iTem = Lst.ListItems.Add(, , rs!CAgencia)
            iTem.SubItems(1) = rs!cCampanas
            iTem.SubItems(2) = Right("00", 2 - Len(rs!cAgeCod)) & rs!cAgeCod
            iTem.SubItems(3) = rs!idCampana
            rs.MoveNext
        Loop
        Set rs = Nothing
    Else
        MsgBox "Debe seleccionar una campaña", vbInformation, "AVISO"
    End If
End Sub

Private Sub LstAgencias_GotFocus()
    LblMensaje.Caption = "Puedes hacer doble click para observar las agencias relacionadas con la campaña"
End Sub

Private Sub LstAgencias_LostFocus()
    LblMensaje.Caption = ""
End Sub

Private Sub LstCampanas_DblClick()
    Dim odCamp As COMDCredito.DCOMCampanas
    Dim rs As ADODB.Recordset
    Dim iTem As ListItem
    
    If LstCampanas.ListIndex <> -1 Then
        Set odCamp = New COMDCredito.DCOMCampanas
        Set rs = odCamp.ListaCampanasAgenciacXCampanas(LstCampanas.ItemData(LstCampanas.ListIndex))
        Set odCamp = Nothing
        
        Lst.ListItems.Clear
        Do Until rs.EOF
            Set iTem = Lst.ListItems.Add(, , rs!CAgencia)
            iTem.SubItems(1) = rs!cCampanas
            iTem.SubItems(2) = Right("00", 2 - Len(rs!cAgeCod)) & rs!cAgeCod
            iTem.SubItems(3) = rs!idCampana
            rs.MoveNext
        Loop
        Set rs = Nothing
    Else
        MsgBox "Debe seleccionar una campaña", vbInformation, "AVISO"
    End If
End Sub

Private Sub LstCampanas_GotFocus()
    LblMensaje.Caption = "Puedes hacer doble click para observar las agencias relacionadas con la campaña"
End Sub

Private Sub LstCampanas_LostFocus()
    LblMensaje.Caption = ""
End Sub

Private Sub MshCampanas_Click()
    On Error GoTo ErrHandler
    
    If nOperacion = lMantenimiento Then
        With MshCampanas
            txtCodigo.Text = .TextMatrix(.row, 0)
            txtDescripcionCampanas.Text = .TextMatrix(.row, 1)
            Me.txtDesde = Format(.TextMatrix(.row, 2), "DD/MM/YYYY") 'BY ARLO 20171013
            Me.txtHasta = Format(.TextMatrix(.row, 3), "DD/MM/YYYY") 'BY ARLO 20171013
            'cmbEstado.ListIndex = IIf(.TextMatrix(.row, 2) = "Activo", 0, 1) 'COMENTADO BY ARLO20171013 ERS062-2017
        End With
   End If
    Exit Sub
ErrHandler:
    MsgBox "Se ha producido un error al seleccionar el registro", vbInformation, "AVISO"
End Sub

Private Sub Stab_Click(PreviousTab As Integer)
    If Stab.Tab = 1 Then
        If nOperacion = lRegistro Then
            CmdNuevoAsiignacion.Visible = True
            CmdGrabarAsignacion.Visible = False
            CmdEliminarAsignacion.Enabled = False
            CmdCampana.Enabled = False
            CmdAgencia.Enabled = False
            CmdAll.Enabled = False
        ElseIf nOperacion = lMantenimiento Then
            CmdNuevoAsiignacion.Visible = True
            CmdGrabarAsignacion.Visible = False
            CmdEliminarAsignacion.Enabled = True
            CmdNuevoAsiignacion.Enabled = False
            CmdCampana.Enabled = False
            CmdAgencia.Enabled = False
            CmdAll.Enabled = False
        Else
            CmdNuevoAsiignacion.Visible = True
            CmdGrabarAsignacion.Visible = False
            CmdEliminarAsignacion.Enabled = False
            CmdNuevoAsiignacion.Enabled = False
            CmdCampana.Enabled = False
            CmdAgencia.Enabled = False
            CmdAll.Enabled = False
        End If
        CargarCampAgeAsignacion
    End If

End Sub

Private Sub Stab_DblClick()
    If Stab.Tab = 1 Then
        If nOperacion = lRegistro Then
            CmdNuevoAsiignacion.Visible = True
            CmdGrabarAsignacion.Visible = False
           ' CmdModificarAsignacion.Enabled = True
            CmdEliminarAsignacion.Enabled = True
            CmdCampana.Enabled = False
            CmdAgencia.Enabled = False
            CmdAll.Enabled = False
        End If
        CargarCampAgeAsignacion
    End If
End Sub

Private Sub txtDescripcionCampanas_KeyPress(KeyAscii As Integer)
    Dim c As String
    
    c = Chr(KeyAscii)
    c = UCase(c)
    KeyAscii = Asc(c)
    '**BY ARLO 20171013
    If KeyAscii = 13 Then
    Me.txtDesde.SetFocus
    End If
    
End Sub

Sub CargarCampanasAsignacion()
    Dim odCamp As COMDCredito.DCOMCampanas
    Dim rs As ADODB.Recordset
    
    Set odCamp = New COMDCredito.DCOMCampanas
    Set rs = odCamp.CargarCampanas
    Set odCamp = Nothing
    
    LstCampanas.Clear
    
    Do Until rs.EOF
        If rs!bEstado = True Then
            LstCampanas.AddItem rs!cDescripcion
            LstCampanas.ItemData(LstCampanas.NewIndex) = rs!idCampana
        End If
        rs.MoveNext
    Loop
End Sub

Sub CargarAgencias()
    Dim odCamp As COMDCredito.DCOMCampanas
    Dim rs As ADODB.Recordset
    
    LstAgencias.Clear
    
    Set odCamp = New COMDCredito.DCOMCampanas
    Set rs = odCamp.CargarAgencias
    Set odCamp = Nothing
    
    Do Until rs.EOF
        LstAgencias.AddItem rs!cAgeDescripcion
        LstAgencias.ItemData(LstAgencias.NewIndex) = CInt(rs!cAgeCod)
        rs.MoveNext
    Loop
    
End Sub

Sub CargarCampAgeAsignacion()
    Dim odCamp As COMDCredito.DCOMCampanas
    Dim rs As ADODB.Recordset
    Dim iTem As ListItem
    
    Lst.ListItems.Clear
    
    Set odCamp = New COMDCredito.DCOMCampanas
    Set rs = odCamp.AsignacionAgenCamp
    Set odCamp = Nothing
    Do Until rs.EOF
        Set iTem = Lst.ListItems.Add(, , rs!CAgencia)
        iTem.SubItems(1) = rs!cCampanas
        iTem.SubItems(2) = Right("00", 2 - Len(rs!cAgeCod)) & rs!cAgeCod
        iTem.SubItems(3) = rs!idCampana
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

'Function fn_VerificarPreviewInsercion() As Integer
'
' Dim i As Integer
' Dim nValor As Integer
' Dim odCamp As COMDCredito.DCOMCampanas
'
' nValor = -1
' Set odCamp = New COMDCredito.DCOMCampanas
'
' For i = 1 To Lst.ListItems.Count
'    If odCamp.VerificarRegistro(Lst.ListItems(i).SubItems(2), Lst.ListItems(i).SubItems(3)) = False Then
'        'significa que el registro de aca empieza
'        nValor = i
'        Exit For
'    End If
' Next i
' Set odCamp = Nothing
'
' fn_VerificarPreviewInsercion = nValor
'End Function
Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsDate(txtDesde.Text) Then
            txtHasta.SetFocus
            Me.txtHasta.SetFocus
        Else
            MsgBox "Fecha Invalida", vbInformation, "MENSAJE DEL SISTEMA"
            txtDesde.SetFocus
        End If
    End If
End Sub
Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsDate(txtHasta.Text) Then
            If (nOperacion) = 1 Then
            Me.CmdGrabarCampanas.SetFocus
            ElseIf (nOperacion) = 2 Then
            Me.CmdActualizar.SetFocus
            End If
        Else
            MsgBox "Fecha Invalida", vbInformation, "MENSAJE DEL SISTEMA"
            txtHasta.SetFocus
        End If
        
    End If
End Sub
