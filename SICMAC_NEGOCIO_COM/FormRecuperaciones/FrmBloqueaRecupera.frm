VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmBloqueaRecupera 
   Caption         =   "Bloqueo Pago Recuperaciones"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   Icon            =   "FrmBloqueaRecupera.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.CheckBox Check1 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Por Nombre"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "&Archivo (Reporte 108335 - Vencidos)"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   7
         Top             =   120
         Width           =   4185
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   5640
         Width           =   1080
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   5640
         Width           =   1080
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   5640
         Width           =   1080
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "+"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   2
         ToolTipText     =   "Agrega Persona a Lista"
         Top             =   840
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   3840
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6773
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Reference Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Credito"
            Object.Width           =   4304
         EndProperty
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   2520
         TabIndex        =   12
         Top             =   480
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1980
         _ExtentX        =   3281
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingresar Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblIngChqDescIF1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   4515
      End
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   330
      Left            =   3600
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   582
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmBloqueaRecupera.frx":030A
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Archivo de Excel (*.xls)|*.xls"
      FilterIndex     =   1
   End
End
Attribute VB_Name = "FrmBloqueaRecupera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPersona As UPersona_Cli   ' COMDPersona.DCOMPersona
Dim sFilename As String

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Dim PerAut As COMDPersona.DCOMPersonas
    Set PerAut = New COMDPersona.DCOMPersonas

            If Not PerAut.ValidaPersBloqueaCredRecupera(Trim(TxtBCodPers.Text), ActxCta.NroCuenta) Then
                MsgBox "No se pudo encontrar el Crédito, o el Credito No esta Vigente", vbInformation, "Aviso"
                cmdAgregar.Enabled = False
            Else
                cmdAgregar.Enabled = True
                cmdAgregar.SetFocus
           End If
            Set PerAut = Nothing
    End If
End Sub

Private Sub Check1_Click()
Dim i As Integer
Dim c As Integer
c = 0
    If lvwNiveles.ListItems.Count >= 1 Then
        If Me.Check1.value = 1 Then
            For i = 1 To lvwNiveles.ListItems.Count
                lvwNiveles.ListItems.iTem(i).Checked = True
            Next
        Else
            For i = 1 To lvwNiveles.ListItems.Count
               lvwNiveles.ListItems.iTem(i).Checked = False
            Next
        End If
    End If
End Sub

Private Sub CmdActualizar_Click()
     definir_solicitud
End Sub

Private Sub CmdAgregar_Click()
     If Me.OptBusqueda(0).value = True Then
           If Len(TxtBCodPers.Text) > 0 Then
            Dim PerAut As COMDPersona.DCOMPersonas
            Dim rs As New ADODB.Recordset
            Set rs = New ADODB.Recordset
            Set PerAut = New COMDPersona.DCOMPersonas

            Set rs = PerAut.DevuelvePersBloqueaRecupera(Trim(TxtBCodPers.Text), ActxCta.NroCuenta)
            If rs.EOF Or rs.BOF Then
                PerAut.InsertaPersBloqueaRecupera Trim(TxtBCodPers.Text), Trim(lblIngChqDescIF1.Caption), Trim(ActxCta.NroCuenta), gdFecSis, True
            End If
            Call CargaDatos
            rs.Close
            Set rs = Nothing
            Set PerAut = Nothing
           Else
             MsgBox "Complete los campos antes de Agregar", vbInformation, "Aviso"
             Exit Sub
           End If
    Else
        CargaArchivoTransferir sFilename
    End If
    limpia_datos
    CargaDatos 1
End Sub

Sub CargaArchivoTransferir(psNomArchivo As String)
    Dim xlApp As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Dim varMatriz As Variant
    Dim cNombreHoja As String
    Dim i As Long, n As Long, ind As Long
    Dim lista As ListItem
    Dim PerAut As COMDPersona.DCOMPersonas
    Set PerAut = New COMDPersona.DCOMPersonas
    i = 0
    ind = 1
    Set xlApp = New Excel.Application

    If Trim(psNomArchivo) = "" Then
        MsgBox "Debe indicar la ruta del Archivo Excel", vbCritical, "Mensaje"
        Me.OptBusqueda(3).value = False
        psNomArchivo = ""
        Exit Sub
   Else
        Set xlLibro = xlApp.Workbooks.Open(psNomArchivo, True, True, , "")
        cNombreHoja = "Hoja1"
        Set xlHoja = xlApp.Worksheets(cNombreHoja)
        varMatriz = xlHoja.Range("A2:AA2000").value
        xlLibro.Close SaveChanges:=False
        xlApp.Quit
        Set xlHoja = Nothing
        Set xlLibro = Nothing
        Set xlApp = Nothing
        n = UBound(varMatriz)
        lvwNiveles.ListItems.Clear
        For i = 8 To n
            If varMatriz(i, 1) = "" Then
                If i = 8 Then
                    MsgBox "Archivo No tiene Estructura Correcta, la informacion debe estar en la Celda (A9)", vbCritical, "Mensaje"
                End If
                Me.OptBusqueda(3).value = False
                psNomArchivo = ""
                Exit For
            Else
                    If varMatriz(i, 2) = "" Or varMatriz(i, 3) = "" Then
                       MsgBox "Archivo No tiene Estructura Correcta", vbCritical, "Mensaje"
                       Exit For
                    Else
                         If Not PerAut.DevuelvePersBloqueaRecuperaBool(Trim(varMatriz(i, 3)), Trim(varMatriz(i, 2))) Then
                            Set lista = lvwNiveles.ListItems.Add(, , gdFecSis)
                            lvwNiveles.ListItems.iTem(ind).Checked = True
                            lista.SubItems(1) = IIf(varMatriz(i, 2) = "", "", varMatriz(i, 2))
                            lista.SubItems(2) = IIf(varMatriz(i, 4) = "", "", varMatriz(i, 4))
                            lista.SubItems(3) = IIf(varMatriz(i, 3) = "", "", varMatriz(i, 3))
                            PerAut.InsertaPersBloqueaRecupera Trim(varMatriz(i, 3)), Trim(varMatriz(i, 4)), Trim(varMatriz(i, 2)), gdFecSis, True
                            ind = ind + 1
                          End If
                    End If
            End If
        Next i
    End If
    Set PerAut = Nothing
End Sub

Public Sub CargaDatos(Optional ByVal ind As Integer = 0)
    Dim PerAut As COMDPersona.DCOMPersonas
    Dim lista As ListItem
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Set PerAut = New COMDPersona.DCOMPersonas
    Set rs = New ADODB.Recordset
    i = 1
    If ind = 1 Then
        Set rs = PerAut.DevuelvePersBloqueaRecuperaTodos
    Else
        Set rs = PerAut.DevuelvePersBloqueaRecupera(Trim(TxtBCodPers.Text), ActxCta.NroCuenta)
    End If
    lvwNiveles.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
       lvwNiveles.ListItems.Clear
       Do Until rs.EOF
         Set lista = lvwNiveles.ListItems.Add(, , rs!dRegistro)
         lvwNiveles.ListItems.iTem(i).Checked = IIf(rs!dVigente, True, False)
         lista.SubItems(1) = IIf(rs!cPersCod = "", "", rs!cPersCod)
         lista.SubItems(2) = rs!cPersNombre
         lista.SubItems(3) = rs!cCtaCod
         i = i + 1
         rs.MoveNext
       Loop
    Else
       MsgBox "No Existen Datos", vbInformation, "Aviso"
    End If
    rs.Close
    Set rs = Nothing
    Set PerAut = Nothing
End Sub

Private Sub cmdEliminar_Click()
    Call eliminar_solicitud
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
    Call limpia_datos
    Call CargaDatos(1)
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 3 Then
        TxtBCodPers.Enabled = False
        lblIngChqDescIF1.Enabled = False
        ActxCta.Enabled = False
        CommonDialog1.ShowOpen
        sFilename = CommonDialog1.FileName
        MsgBox "Seleccione el Archivo y Presione el Boton + ", vbInformation, "Aviso"
        cmdAgregar.Enabled = True
    Else
        TxtBCodPers.Enabled = True
        lblIngChqDescIF1.Enabled = True
        ActxCta.Enabled = True
        cmdAgregar.Enabled = False
    End If
    'cmdAgregar.Visible = True
End Sub

Private Sub TxtBCodPers_EmiteDatos()
 If Trim(TxtBCodPers.Text) = "" Then
        Exit Sub
    End If

If Cargar_Datos_Persona(Trim(TxtBCodPers.Text)) = False Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
End If
End Sub

Private Sub TxtBCodPers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(TxtBCodPers) > 0 Then
            ActxCta.SetFocus
    End If
End If
End Sub

Function Cargar_Datos_Persona(pcPersCod As String) As Boolean

    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona

    Cargar_Datos_Persona = True

    Call oPersona.RecuperaPersona(pcPersCod, , gsCodUser)

    If oPersona.PersCodigo = "" Then
        Cargar_Datos_Persona = False
        Exit Function
    Else
        lblIngChqDescIF1 = oPersona.NombreCompleto
    End If
End Function

Public Function ValidarNroDatos() As Boolean
    Dim i As Integer
    Dim c As Integer
    ValidarNroDatos = True
    c = 0
    For i = 1 To lvwNiveles.ListItems.Count
        If lvwNiveles.ListItems.iTem(i).Checked = True Then
            c = c + 1
        End If
    Next
'    If c > 1 Then
'        MsgBox "Solo debe seleccionar una sola Fila de Datos", vbInformation, "Aviso"
'        ValidarNroDatos = False
'    ElseIf c = 0 Then
'        MsgBox "Seleccione una sola Fila de Datos", vbInformation, "Aviso"
'        ValidarNroDatos = False
'    End If
End Function

Sub definir_solicitud()
Dim ocapaut As COMDPersona.DCOMPersonas
Dim lnNumApro As Integer
Dim lnNumDApro As Integer
Dim fechasol As Date
Dim lbAprobado As Boolean
Dim pCodUser As String
Dim pCtaCod As String
Dim i As Integer
   If ValidarNroDatos = False Then Exit Sub
   lnNumApro = 0
   lnNumDApro = 0
   For i = 1 To lvwNiveles.ListItems.Count
        fechasol = CDate(lvwNiveles.ListItems.iTem(i).Text)
        pCtaCod = lvwNiveles.ListItems.iTem(i).SubItems(3)
        pCodUser = lvwNiveles.ListItems.iTem(i).SubItems(1)
        lbAprobado = IIf(lvwNiveles.ListItems.iTem(i).Checked, True, False)
        Set ocapaut = New COMDPersona.DCOMPersonas
        If lbAprobado = True Then
            ocapaut.ActualizarPersBloqueaRecupera pCodUser, pCtaCod, gdFecSis, True
            Set ocapaut = Nothing
            lnNumApro = lnNumApro + 1
       Else
            ocapaut.ActualizarPersBloqueaRecupera pCodUser, pCtaCod, gdFecSis, False
            Set ocapaut = Nothing
            lnNumDApro = lnNumDApro + 1
       End If
   Next
'   If lnNumApro = 0 Or lnNumDApro = 0 Then
'      MsgBox "No se pudo Actualizar el Registro", vbInformation, "Aviso"
'   Else
      CargaDatos 1
'   End If
End Sub

Sub eliminar_solicitud()
Dim ocapaut As COMDPersona.DCOMPersonas
Dim lnNumApro As Integer
Dim lbAprobado As Boolean
Dim pCodUser As String
Dim pCtaCod As String
Dim i As Integer
   If ValidarNroDatos = False Then Exit Sub
   lnNumApro = 0
   For i = 1 To lvwNiveles.ListItems.Count
        pCtaCod = lvwNiveles.ListItems.iTem(i).SubItems(3)
        pCodUser = lvwNiveles.ListItems.iTem(i).SubItems(1)
        lbAprobado = IIf(lvwNiveles.ListItems.iTem(i).Checked, True, False)
        Set ocapaut = New COMDPersona.DCOMPersonas
        If lbAprobado = True Then
            ocapaut.EliminaPersBloqueaRecupera pCodUser, pCtaCod
            Set ocapaut = Nothing
            lnNumApro = lnNumApro + 1
       End If
   Next i
   If lnNumApro = 0 Then
      MsgBox "Seleccione un registro para ser eliminado", vbInformation, "Aviso"
   Else
      CargaDatos 1
   End If
End Sub

Sub limpia_datos()
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    lblIngChqDescIF1.Caption = ""
    TxtBCodPers.Text = ""
End Sub


