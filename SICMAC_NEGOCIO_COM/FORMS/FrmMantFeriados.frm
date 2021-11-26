VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmMantFeriados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Feriados"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   2655
      Left            =   45
      TabIndex        =   22
      Top             =   1185
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   3
      TextStyleFixed  =   1
      GridLinesUnpopulated=   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Frame Frame4 
      Height          =   1965
      Left            =   6195
      TabIndex        =   4
      Top             =   90
      Width           =   1365
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   135
         TabIndex        =   8
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   135
         TabIndex        =   7
         Top             =   615
         Width           =   1095
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   135
         TabIndex        =   6
         Top             =   1035
         Width           =   1095
      End
      Begin VB.CommandButton CmdDomingos 
         Caption         =   "&Domingos"
         Height          =   375
         Left            =   135
         TabIndex        =   5
         Top             =   1455
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1545
      Left            =   6195
      TabIndex        =   0
      Top             =   2295
      Width           =   1365
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   165
         TabIndex        =   3
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   165
         TabIndex        =   2
         Top             =   630
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   165
         TabIndex        =   1
         Top             =   1050
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha"
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   1095
      Left            =   75
      TabIndex        =   17
      Top             =   0
      Width           =   5985
      Begin VB.CommandButton CmdMuestraAgencias 
         Caption         =   "..."
         Height          =   405
         Left            =   5490
         TabIndex        =   23
         ToolTipText     =   "Mostrar Agencias"
         Top             =   465
         Width           =   375
      End
      Begin VB.TextBox TxtDesFer 
         Height          =   405
         Left            =   1785
         MaxLength       =   25
         TabIndex        =   19
         Top             =   480
         Width           =   3615
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   1785
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período"
      ForeColor       =   &H00404040&
      Height          =   1095
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin MSMask.MaskEdBox txtFechaI 
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaF 
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Inicio:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Fin:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "FrmMantFeriados.frx":0000
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Buscar Feriado"
      ForeColor       =   &H00404040&
      Height          =   1095
      Left            =   90
      TabIndex        =   9
      Top             =   0
      Width           =   5985
      Begin MSMask.MaskEdBox txtFechaB 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   360
         Picture         =   "FrmMantFeriados.frx":0442
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         Caption         =   "Ingrese Fecha a Buscar, pulse enter para realizar la búsqueda..........."
         Height          =   495
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "FrmMantFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OpcB As Integer
Dim mItem As ListItem
Dim Opc As Integer ' Control de la opción seleccionada
Dim DFer As Date
Dim DescFer As String
Dim OpcDom As Boolean
Dim RegVerFec As New ADODB.Recordset
Dim RFeriadoAge As New ADODB.Recordset
Dim MatAgencias() As String

Private Sub cmdBuscar_Click()

'realiza la búsqueda de una determinada fecha
Frame3.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Inicializa
txtFechaB.SetFocus

End Sub

Private Sub CmdEliminar_Click()
'Elimina un registro
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
CmdGrabar.Enabled = False
If MSH.Col = 1 Then
    If MsgBox("Desea eliminar  " & Me.MSH.TextMatrix(MSH.row, 1) & "  de la Tabla Feriados?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
        If CDate(Me.MSH.TextMatrix(MSH.row, 1)) < gdFecSis Then
            MsgBox "No puede eliminar un feriado menor a la fecha actual", vbInformation, "Aviso"
        Else
            Call EliminaRegistro(Me.MSH.TextMatrix(MSH.row, 1))
            LlenaFeriados
       End If
        Inicializa
    End If
End If
End Sub

Private Sub CmdMuestraAgencias_Click()

'NRLO 20180502 SATI INC1804260002 CORE
If Len(ValidaFecha(txtFecha)) > 0 Then
    MsgBox ValidaFecha(txtFecha), vbInformation, "Aviso"
    txtFecha.SetFocus
    Exit Sub
End If
'FIN NRLO 20180502 SATI INC1804260002 CORE

Dim oDFeriado As COMDCredito.DCOMFeriado
Dim i As Integer
    Set oDFeriado = New COMDCredito.DCOMFeriado
    Set RFeriadoAge = oDFeriado.RecuperaFeriadoAgencias(CDate(txtFecha.Text))
    Set oDFeriado = Nothing
    ReDim MatAgencias(RFeriadoAge.RecordCount, 3)
    i = 0
    Do While Not RFeriadoAge.EOF
        MatAgencias(i, 0) = RFeriadoAge!cAgeCod
        MatAgencias(i, 1) = RFeriadoAge!cAgeDescripcion
        MatAgencias(i, 2) = RFeriadoAge!valor
        i = i + 1
        RFeriadoAge.MoveNext
    Loop
    RFeriadoAge.Close
    Set RFeriadoAge = Nothing
    
    If CmdGrabar.Enabled Then
        MatAgencias = frmMntFeriadoAge.CargaFlex(MatAgencias, True)
    Else
        MatAgencias = frmMntFeriadoAge.CargaFlex(MatAgencias, False)
    End If
    
End Sub

Private Sub Form_Load()
Inicializa
LlenaFeriados
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub cmdNuevo_Click()
Opc = 1
Frame1.Visible = True
Frame1.Enabled = True
Frame2.Visible = False
Frame3.Visible = False
CmdEliminar.Enabled = False
CmdBuscar.Enabled = False
CmdDomingos.Enabled = False
CmdGrabar.Enabled = True
CmdCancelar.Enabled = True
txtFecha.SetFocus
CmdNuevo.Enabled = False
ReDim MatAgencias(0)
CmdMuestraAgencias.Enabled = True

End Sub

Sub Marco()
With MSH
    .Clear
    .ColWidth(0) = 100
    .ColWidth(1) = 1500
    .ColWidth(2) = 4000
    .Clear
    .Rows = 2
    .TextMatrix(0, 1) = "Fecha"
    .TextMatrix(0, 2) = "Descripcion"
End With
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Sub Inicializa()
Opc = 0
OpcDom = False
txtFecha.Mask = ""
txtFecha = ""
txtFecha.Mask = "##/##/####"
TxtDesFer = ""
txtFechaB.Mask = ""
txtFechaB = ""
txtFechaB.Mask = "##/##/####"
txtFechaI.Mask = ""
txtFechaI = ""
txtFechaI.Mask = "##/##/####"
txtFechaF.Mask = ""
txtFechaF = ""
txtFechaF.Mask = "##/##/####"
CmdEliminar.Enabled = True
CmdBuscar.Enabled = True
CmdDomingos.Enabled = True
CmdGrabar.Enabled = False
CmdNuevo.Enabled = True
End Sub

Sub LlenaFeriados()
'LLenado del ListView con los feriados desde 5 años atrás
Dim RangoF As Variant
Dim RegFeriado As New ADODB.Recordset
Dim DFe As COMDCredito.DCOMFeriado
Marco
RangoF = Format(DateAdd("yyyy", -5, gdFecSis), "mm/dd/yyyy")
Set DFe = New COMDCredito.DCOMFeriado
Set RegFeriado = DFe.RecuperaDias(RangoF)

If RegFeriado.EOF And RegFeriado.BOF Then
Else
    Screen.MousePointer = 11
    While Not RegFeriado.EOF
        MSH.TextMatrix(MSH.Rows - 1, 1) = RegFeriado!DFeriado
        MSH.TextMatrix(MSH.Rows - 1, 2) = RegFeriado!cdescrip
        MSH.Rows = MSH.Rows + 1
        RegFeriado.MoveNext
    Wend
    Screen.MousePointer = 0
End If
RegFeriado.Close
Set RegFeriado = Nothing
Set DFe = Nothing
End Sub


Private Sub MSH_DblClick()
Dim oDFeriado As COMDCredito.DCOMFeriado
Dim i As Integer
    Set oDFeriado = New COMDCredito.DCOMFeriado
    Set RFeriadoAge = oDFeriado.RecuperaFeriadoAgencias(CDate(MSH.TextMatrix(MSH.row, 1)))
    Set oDFeriado = Nothing
    ReDim MatAgencias(RFeriadoAge.RecordCount, 3)
    i = 0
    Do While Not RFeriadoAge.EOF
        MatAgencias(i, 0) = RFeriadoAge!cAgeCod
        MatAgencias(i, 1) = RFeriadoAge!cAgeDescripcion
        MatAgencias(i, 2) = RFeriadoAge!valor
        i = i + 1
        RFeriadoAge.MoveNext
    Loop
    RFeriadoAge.Close
    Set RFeriadoAge = Nothing
    
    If CmdGrabar.Enabled Then
        MatAgencias = frmMntFeriadoAge.CargaFlex(MatAgencias, True)
    Else
        MatAgencias = frmMntFeriadoAge.CargaFlex(MatAgencias, False)
    End If
    
End Sub

Private Sub TxtDesFer_KeyPress(KeyAscii As Integer)
Dim Car As String * 1
Car = Chr$(KeyAscii)
If (KeyAscii = 13 Or KeyAscii = 9) Then
        CmdGrabar.Enabled = True
        CmdGrabar.SetFocus
End If
If (Car > "0" And Car < "9") Then
    Beep
    KeyAscii = 0
Else
    KeyAscii = Letras(KeyAscii)
End If
End Sub



Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Or KeyAscii = 39 Then
   If ValidaFecha(txtFecha) = "" Then
      CmdGrabar.Enabled = True
      TxtDesFer.SetFocus
   Else
      MsgBox ValidaFecha(txtFecha), vbInformation, "Aviso"
      txtFecha.SetFocus
      Exit Sub
   End If
End If
End Sub

Private Sub TxtFechaB_KeyPress(KeyAscii As Integer)
Dim d As Date
Dim tmp As String
If KeyAscii = 13 Or KeyAscii = 9 Or KeyAscii = 39 Then
   If ValidaFecha(txtFechaB) = "" Then
   Else
      txtFechaB.Mask = ""
      txtFechaB = ""
      txtFechaB.Mask = "##/##/####"
      MsgBox ValidaFecha(txtFechaB), vbInformation, "Aviso"
      txtFechaB.SetFocus
      Exit Sub
   End If
   If Not VerSiExisFer(Format(txtFechaB, "dd/mm/yyyy")) Then
      MsgBox "La fecha seleccionada no se encuentra en la Base de Datos", vbInformation, "Aviso"
      txtFechaB.Mask = ""
      txtFechaB = ""
      txtFechaB.Mask = "##/##/####"
      txtFechaB.SetFocus
      Exit Sub
   Else
        d = MSH.TextMatrix(MSH.row, MSH.Col)
        tmp = MSH.TextMatrix(MSH.row, MSH.Col + 1)
        Marco
        MSH.TextMatrix(1, 1) = d
        MSH.TextMatrix(1, 2) = tmp
        MsgBox "Fecha encontrada", vbInformation, "Aviso"
        cmdCancelar_Click
   End If
End If
End Sub

Function VerSiExisFer(FecVer As Variant) As Boolean
Dim RegVerFec As New ADODB.Recordset
Dim lnFecVer As String
Dim DF As COMDCredito.DCOMFeriado
Dim i As Integer
Set DF = New COMDCredito.DCOMFeriado

Set RegVerFec = DF.DetallaFeriado(FecVer, gsCodAge)

If RegVerFec.EOF And RegVerFec.BOF Then
    VerSiExisFer = False
Else
    For i = 1 To MSH.Rows - 1
        If MSH.TextMatrix(i, 1) = Format(FecVer, "dd/mm/yyyy") Then
            MSH.Col = 1
            MSH.row = i
            Exit For
        End If
    Next i
    VerSiExisFer = True
End If

RegVerFec.Close
Set DF = Nothing
Set RegVerFec = Nothing
  
End Function

Sub EliminaRegistro(Fecha As Variant)
Dim DelFer As Variant
Dim DF As COMDCredito.DCOMFeriado
Set DF = New COMDCredito.DCOMFeriado
DelFer = Format(Fecha, "mm/dd/yyyy")
DF.EliminaFeriado (CDate(Fecha))
Set DF = Nothing
End Sub

Private Sub GetDomingos()
Dim Dias As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim i As Integer

fecha1 = CDate(txtFechaI)
fecha2 = CDate(txtFechaF)

Dias = DateDiff("d", fecha1, fecha2)

For i = 0 To Dias
    If Weekday(fecha1) = 1 Then
        If VerSiExisFer(Format(fecha1, "mm/dd/yyyy")) Then
        Else
            Call LlenaDomingos(fecha1)
        End If
    End If
    fecha1 = DateAdd("d", 1, fecha1)
Next i
LlenaFeriados
End Sub

Private Sub LlenaDomingos(Domingo As Variant)
Dim lsDomingo
Dim DF As COMDCredito.DCOMFeriado
Set DF = New COMDCredito.DCOMFeriado
OpcDom = True
lsDomingo = Format$(Domingo, "dd/mm/yyyy")
Call DF.LlenaDomingo(lsDomingo, gdFecSis, gsCodAge, gsCodUser)
End Sub

Private Sub cmdCancelar_Click()
Inicializa
LlenaFeriados
CmdCancelar.Enabled = False
CmdMuestraAgencias.Enabled = False
End Sub

Private Sub CmdDomingos_Click()
Inicializa
Opc = 2
txtFechaI.Mask = ""
txtFechaI = ""
txtFechaI.Mask = "##/##/####"
txtFechaI.Format = "dd/mm/yyyy"

OpcDom = False
Frame1.Visible = False
Frame3.Visible = False
Frame2.Visible = True
CmdNuevo.Enabled = False
CmdEliminar.Enabled = False
CmdBuscar.Enabled = False
CmdGrabar.Enabled = True
CmdCancelar.Enabled = True
txtFechaI.SetFocus
End Sub

Private Sub CmdGrabar_Click()
Dim FecFeriado As String
Dim FecMod As String
Dim DF As COMDCredito.DCOMFeriado
Set DF = New COMDCredito.DCOMFeriado
Dim i As Integer

'COMENTADO POR LARI 20210104
'If UBound(MatAgencias) = 0 Then
'    MsgBox "No se verifico agencias", vbInformation, "Aviso"
'    Exit Sub
'End If

'20210104LARI******************************************************
If Opc <> 2 Then
    If UBound(MatAgencias) = 0 Then
        MsgBox "No se verifico agencias", vbInformation, "Aviso"
        Exit Sub
    End If
End If
'******************************************************************

CmdGrabar.Enabled = False
CmdCancelar.Enabled = False
Select Case Opc
Case 1 'Nuevo feriado
    If ValidaFecha(txtFecha) = "" Then
    Else
       MsgBox ValidaFecha(txtFecha), vbInformation, "Aviso"
       txtFecha.SetFocus
       CmdGrabar.Enabled = True
       CmdCancelar.Enabled = True
       Exit Sub
    End If
    'by capi 06122007 por el formato de fecha
    'FecFeriado = Format(txtFecha, "mm/dd/yyyy")
    FecFeriado = Format(txtFecha, "yyyy/mm/dd")
    
    If Not VerSiExisFer(FecFeriado) Then
        If MsgBox("Desea Grabar la Información?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            'By Capi 06122007 por el formato fecha
            'Call DF.InsertaFecha(CDate(txtFecha.Text), Me.TxtDesFer, gdFecSis, gsCodAge, gsCodUser)
            Call DF.InsertaFecha(FecFeriado, Me.TxtDesFer, gdFecSis, gsCodAge, gsCodUser)
            
            For i = 0 To UBound(MatAgencias) - 1
                If MatAgencias(i, 2) = "." Then
                    Call DF.dInsertFeriadoAge(CDate(txtFecha.Text), MatAgencias(i, 0))
                End If
            Next i
            
        Else
            CmdGrabar.Enabled = True
            CmdCancelar.Enabled = True
            'dbCmact.RollbackTrans
            Exit Sub
        End If
    Else
        MsgBox "Dato ya ingresado", vbInformation, "Aviso"
        CmdGrabar.Enabled = False
        CmdCancelar.Enabled = False
        CmdNuevo.Enabled = True
        CmdEliminar.Enabled = True
        CmdBuscar.Enabled = True
        CmdDomingos.Enabled = True
        Exit Sub
    End If
Case 2 'Agregar Domingos
    If ValidaFecha(txtFechaI) = "" Then
    Else
        MsgBox ValidaFecha(txtFechaI), vbInformation, "Aviso"
        txtFechaI.SetFocus
        Exit Sub
    End If
    If ValidaFecha(txtFechaF) = "" Then
         GetDomingos
         If OpcDom = False Then
            MsgBox "No existen Domingos dentro de este período", vbInformation, "Aviso"
            CmdNuevo.Enabled = True
            CmdEliminar.Enabled = True
            CmdBuscar.Enabled = True
            Exit Sub
         End If
    Else
        MsgBox ValidaFecha(txtFechaF), vbInformation, "Aviso"
        txtFechaF.SetFocus
        Exit Sub
    End If
End Select

Inicializa
LlenaFeriados
OpcDom = False
CmdMuestraAgencias.Enabled = False
End Sub


Private Sub txtFechaF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Or KeyAscii = 39 Then
   If ValidaFecha(txtFechaF) = "" Then
      If CDate(txtFechaI) > CDate(txtFechaF) Then
        MsgBox "Fecha Fin menor que Fecha Inicio, vuelva a Ingresar.", vbInformation, "Aviso"
        txtFechaF.SetFocus
      Else
          CmdGrabar.Enabled = True
          CmdGrabar.SetFocus
      End If
   Else
      MsgBox ValidaFecha(txtFechaF), vbInformation, "Aviso"
      txtFechaF.SetFocus
      Exit Sub
   End If
End If
End Sub

Private Sub txtFechaI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Or KeyAscii = 39 Then
    If ValidaFecha(txtFechaI) = "" Then
       txtFechaF.Enabled = True
       txtFechaF.SetFocus
    Else
       MsgBox ValidaFecha(txtFechaI), vbInformation, "Aviso"
       txtFechaI.SetFocus
       Exit Sub
    End If
End If
End Sub
