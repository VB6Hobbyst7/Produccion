VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmContCtasxAge 
   Caption         =   "Mantenimiento de cuentas contables por agencias"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   Icon            =   "frmContCtasxAge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraContCtasxAge 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Top             =   1130
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6960
         Picture         =   "frmContCtasxAge.frx":030A
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid grdCtas 
         Height          =   2715
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   4789
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   17
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "cCtaContCod"
            Caption         =   "Código"
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
            DataField       =   "cCtaContDesc"
            Caption         =   "Descripción"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Locked          =   -1  'True
            Size            =   2
            BeginProperty Column00 
               DividerStyle    =   6
               ColumnAllowSizing=   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1755.213
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
               ColumnAllowSizing=   -1  'True
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   6224.882
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Frame fraMoneda 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   4200
         TabIndex        =   5
         Top             =   120
         Width           =   2595
         Begin VB.CheckBox chkMoneda 
            Caption         =   "1 - Moneda Nacional"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   8
            Top             =   570
            Width           =   1845
         End
         Begin VB.CheckBox chkMoneda 
            Caption         =   " 0 - Integrador"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   7
            Top             =   270
            Width           =   1935
         End
         Begin VB.CheckBox chkMoneda 
            Caption         =   "2 - Moneda Extranjera"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   6
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.TextBox txtDigRest 
         BackColor       =   &H00F0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   1920
         MaxLength       =   23
         TabIndex        =   3
         Top             =   240
         Width           =   2115
      End
      Begin VB.TextBox txtDigInt 
         Alignment       =   2  'Center
         BackColor       =   &H00F0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   "Indica la moneda para la cuenta contable"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtDigDos 
         BackColor       =   &H00F0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView lvAgencia 
      Height          =   2265
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3995
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Agencia"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.CommandButton cmdAgencia 
      Caption         =   "&Generar Agencias >>>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   2355
   End
End
Attribute VB_Name = "frmContCtasxAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTabla As String
Dim lOk  As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim nPost As Integer
Dim clsCtaCont As DCtaCont
Dim nOrdenCta As Integer
Dim E As Boolean
Dim rsCtaObj As ADODB.Recordset
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub cmdAceptar_Click()
Dim SQLctas As String
Dim sCta      As String
Dim lAgencias As Boolean
Dim m As Integer
Dim n As Integer
Set rs = New ADODB.Recordset
Dim bAge As Integer
Dim nEst As Integer
Dim nEstGen As Integer
Dim sCodHisto As String
Dim sCodDescHisto As String

On Error GoTo errAcepta
If chkMoneda(1).value = 0 And chkMoneda(2).value = 0 Then
    MsgBox " No se definió moneda...! ", vbInformation, "Aviso"
   Exit Sub
End If
If txtDigDos = "" And txtDigInt = "" And txtDigRest = "" Then
   MsgBox " No se definió Cuenta Contable...! ", vbInformation, "Aviso"
   txtDigDos.SetFocus
   Exit Sub
End If
If sTabla = "CtaCont" Then
   sTabla = gsCentralCom & "CtaCont"
Else
   sTabla = "CtaCont"
End If
sCta = txtDigDos & "0" & txtDigRest
If sTabla = gsCentralCom & "CtaCont" Then
   glAceptar = True
   Set clsCtaCont = New DCtaCont
   Set rs = clsCtaCont.CargaCtaCont("substring('" & sCta & "',1,LEN(cCtaContCod)) = cCtaContCod", "CtaContBase")
   If rs.EOF Then
      MsgBox "Cuenta Contable no permitida. Verificar Plan de Cuentas de SBS", vbInformation, "Aviso"
      glAceptar = False
      Exit Sub
   End If
   rs.MoveLast
   sSql = "cCtaContCod LIKE '" & rs!cCtaContCod & "_%' "
   Set rs = clsCtaCont.CargaCtaCont(sSql, "CtaContBase")
   If rs.RecordCount > 0 Then
      If MsgBox("Cuenta Contable no permitida segun Plan de Cuentas de SBS. ¿ Desea Continuar ?", vbQuestion + vbYesNo, "!Confirmación¡") = vbNo Then
         glAceptar = False
      End If
   End If
   RSClose rs
   If Not glAceptar Then
      If txtDigDos.Enabled Then
         txtDigDos.SetFocus
      Exit Sub
   End If
End If

If lvAgencia.ListItems.Count > 0 Then
   Set rs = clsCtaCont.CargaCtaCont(" cCtaContCod LIKE '" & sCta & "__' ")
   If rs.RecordCount > 0 Then
      lAgencias = True
   End If
   RSClose rs
End If

   If MsgBox(" ¿ Seguro de grabar Datos ? ", vbOKCancel + vbQuestion, "Confirma grabación") = vbOk Then
        bAge = "0"
        nEst = "0"
        nEstGen = 0
   End If
   If lvAgencia.ListItems.Count > 0 Then
      If lAgencias Then
         n = MsgBox(" Cuenta ya posee divisionarias de Agencias. ¿ Desea Continuar ? ", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Confirmación")
         If n = vbCancel Then
            Exit Sub
         End If
      Else
         n = vbYes
      End If
      If n = vbYes Then
         For n = 1 To lvAgencia.ListItems.Count
            sCta = txtDigDos & "_" & txtDigRest & Right(lvAgencia.ListItems(n).Text, 2)
            If lvAgencia.ListItems(n).Checked Then
                If chkMoneda(1).value = 1 Then
                    clsCtaCont.InsertaCtaContAge Replace(sCta, "_", "1"), Right(lvAgencia.ListItems(n).Text, 2)
                End If
                If chkMoneda(2).value = 1 Then
                    clsCtaCont.InsertaCtaContAge Replace(sCta, "_", "2"), Right(lvAgencia.ListItems(n).Text, 2)
                End If
            End If
         Next
      End If
   End If
End If
   lOk = True
   Call DatosCtasAgencias
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantCctaCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " |Se Agrego la Cuenta Contable  " & sCta
            Set objPista = Nothing
            '*******
Exit Sub
errAcepta:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdEliminar_Click()
Dim sCta As String
Set clsCtaCont = New DCtaCont
sCta = txtDigDos & "_" & txtDigRest
If chkMoneda(1).value = 1 Then
    clsCtaCont.EliminarCtaContAge Replace(sCta, "_", "1"), ""
End If
If chkMoneda(2).value = 1 Then
    clsCtaCont.EliminarCtaContAge Replace(sCta, "_", "2"), ""
End If
MsgBox "Proceso realizado satisfactoriamente"
Call DatosCtasAgencias
Call cmdAgencia_Click
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantCctaCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " |Se Elimino la Cuenta Contable  " & sCta
            Set objPista = Nothing
            '*******
End Sub

Private Sub grdCtas_GotFocus()
grdCtas.MarqueeStyle = dbgHighlightRow
End Sub
Private Sub grdCtas_HeadClick(ByVal ColIndex As Integer)
If Not rs Is Nothing Then
   If Not rs.EOF Then
      rs.Sort = grdCtas.Columns(ColIndex).DataField
   End If
End If
End Sub
Private Sub grdCtas_LostFocus()
grdCtas.MarqueeStyle = dbgNoMarquee
End Sub
Private Sub grdCtas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If E = False Then
If Not LastRow = "" Then
   If Not rs.EOF Then
'      If LastRow <> rs.Bookmark Then
'         CargaCtaObjs
'      End If
   End If
End If
End If
End Sub
Private Sub cmdAgencia_Click()
Dim rsAge As ADODB.Recordset
Dim lvItem As ListItem
Dim oAge As New DActualizaDatosArea
If cmdAgencia.Caption = "&Generar Agencias >>>" Then
    Set rsAge = New ADODB.Recordset
    Set rsAge = oAge.GetAgencias(, False)
   If rsAge.EOF Then
      RSClose rsAge
      MsgBox "No se definieron Agencias en el Sistema...Consultar con Sistemas", vbInformation, "Aviso"
      Exit Sub
   End If
   cmdAgencia.Caption = "No &Generar Agencias <<<"
   'Me.Height = Me.Height + 2500
   Do While Not rsAge.EOF
      Set lvItem = lvAgencia.ListItems.Add(, , rsAge!Codigo)
      lvItem.SubItems(1) = rsAge!Descripcion
      lvItem.Checked = True
      rsAge.MoveNext
   Loop
   RSClose rsAge
End If
End Sub

Private Sub Form_Load()
    Set clsCtaCont = New DCtaCont
    Call DatosCtasAgencias
    Call cmdAgencia_Click
End Sub
Private Sub DatosCtasAgencias()
    Dim i, j As Integer
    Set rs = New ADODB.Recordset
    Set clsCtaCont = New DCtaCont
    Set rs = clsCtaCont.ListarCtaContAge
    Set clsCtaCont = Nothing
    Set grdCtas.DataSource = Nothing
    Set grdCtas.DataSource = rs
End Sub
Private Sub RefrescaGrid(npMoneda As Integer)
Set clsCtaCont = New DCtaCont
Set rs = clsCtaCont.ListarCtaContAge
Set grdCtas.DataSource = rs
End Sub
'Private Sub CargaCtaObjs()
'If Not rs.EOF And Not rs.BOF Then
'   Set rsCtaObj = clsCtaCont.CargaCtaObj(rs!cCtaContCod, , True)
'Else
'   Set rsCtaObj = Nothing
'End If
'End Sub
Private Sub ManejaBoton(plOpcion As Boolean)
cmdBuscar.Enabled = plOpcion
grdCtas.Enabled = plOpcion
End Sub
Private Sub cmdBuscar_Click()
Dim clsBuscar As New ClassDescObjeto
On Error GoTo ErrMsg
ManejaBoton False
clsBuscar.BuscarDato rs, nOrdenCta, "Cuenta Contable"
nOrdenCta = clsBuscar.gnOrdenBusca
Set clsBuscar = Nothing
ManejaBoton True
grdCtas.SetFocus
Exit Sub
ErrMsg:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
