VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmPromoRegistro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Promociones : Registro de Promotoras"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "frmPromoRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   165
      TabIndex        =   8
      Top             =   6540
      Width           =   870
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   3960
      TabIndex        =   4
      Top             =   6540
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   4920
      TabIndex        =   5
      Top             =   6540
      Width           =   870
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   3015
      TabIndex        =   3
      Top             =   6540
      Width           =   870
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   2070
      TabIndex        =   2
      Top             =   6540
      Width           =   870
   End
   Begin VB.Frame fraPromotor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Promotor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Width           =   5895
      Begin VB.TextBox txtPromotor 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   135
         TabIndex        =   0
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lblProNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1215
         TabIndex        =   6
         Top             =   315
         Width           =   4515
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   345
      Left            =   3840
      TabIndex        =   10
      Top             =   6660
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   609
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmPromoRegistro.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Editar"
      Height          =   390
      Left            =   1125
      TabIndex        =   9
      Top             =   6540
      Width           =   870
   End
   Begin VB.Frame frmDetalle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   30
      TabIndex        =   7
      Top             =   585
      Width           =   5895
      Begin VB.Frame fraContactos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   5655
         Begin MSComctlLib.ListView lvwContactos 
            Height          =   1710
            Left            =   90
            TabIndex        =   19
            Top             =   150
            Width           =   5475
            _ExtentX        =   9657
            _ExtentY        =   3016
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cliente"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Codigo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Comenta"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Capta Soles"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Capta Dolares"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Coloc Soles"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Coloc Dolares"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraProductos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3660
         Left            =   120
         TabIndex        =   11
         Top             =   2055
         Width           =   5655
         Begin Sicmact.TxtBuscar txtBuscar 
            Height          =   315
            Left            =   150
            TabIndex        =   32
            Top             =   240
            Width           =   1455
            _ExtentX        =   2143
            _ExtentY        =   582
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            TipoBusqueda    =   3
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "B4 - HIPECARIO"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   2940
            TabIndex        =   31
            Tag             =   "B4"
            Top             =   1320
            Width           =   2085
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "B5 - MI VIVIENDA"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   2940
            TabIndex        =   30
            Tag             =   "B5"
            Top             =   1545
            Width           =   2085
         End
         Begin Sicmact.EditMoney txtCapSol 
            Height          =   360
            Left            =   825
            TabIndex        =   21
            Top             =   1815
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtComenta 
            Appearance      =   0  'Flat
            Height          =   645
            Left            =   75
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   2910
            Width           =   5475
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "B3 - PIGNORATICIO"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   2940
            TabIndex        =   17
            Tag             =   "B3"
            Top             =   1095
            Width           =   2085
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "B2 - CONSUMO"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   2940
            TabIndex        =   16
            Tag             =   "B2"
            Top             =   870
            Width           =   2085
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "B1 - PYME"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   2940
            TabIndex        =   15
            Tag             =   "B1"
            Top             =   615
            Width           =   2085
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "A3 - CTS"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Tag             =   "A3"
            Top             =   1095
            Width           =   1680
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "A2 - PLAZO FIJO"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Tag             =   "A2"
            Top             =   855
            Width           =   1635
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "A1 - AHORROS"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Tag             =   "A1"
            Top             =   615
            Width           =   1605
         End
         Begin Sicmact.EditMoney txtColSol 
            Height          =   360
            Left            =   3630
            TabIndex        =   23
            Top             =   1830
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin Sicmact.EditMoney txtCapDol 
            Height          =   360
            Left            =   825
            TabIndex        =   25
            Top             =   2205
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            BackColor       =   65280
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin Sicmact.EditMoney txtColDol 
            Height          =   360
            Left            =   3630
            TabIndex        =   27
            Top             =   2220
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            BackColor       =   65280
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblNomPers 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1605
            TabIndex        =   33
            Top             =   240
            Width           =   3960
         End
         Begin VB.Label lblComentario 
            Caption         =   "Comentario"
            Height          =   240
            Left            =   90
            TabIndex        =   29
            Top             =   2685
            Width           =   1050
         End
         Begin VB.Label lblColDol 
            Caption         =   "US$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3030
            TabIndex        =   28
            Top             =   2295
            Width           =   420
         End
         Begin VB.Label lblCaoDol 
            Caption         =   "US$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   225
            TabIndex        =   26
            Top             =   2280
            Width           =   420
         End
         Begin VB.Label lblColSol 
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3030
            TabIndex        =   24
            Top             =   1905
            Width           =   420
         End
         Begin VB.Label lblCapSol 
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   225
            TabIndex        =   22
            Top             =   1890
            Width           =   420
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            DrawMode        =   2  'Blackness
            X1              =   2805
            X2              =   2820
            Y1              =   540
            Y2              =   2175
         End
      End
   End
End
Attribute VB_Name = "frmPromoRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Registro de Promocion de Productos
'Fecha : 17/07/2002
'LAYG
Dim lsNueMod As String * 1

Dim lbConexion As Boolean
Dim fConexCentral As New ADODB.Connection

Private Sub cmdCancelar_Click()

    Me.fraPromotor.Enabled = True
    Me.fraContactos.Enabled = True
    Me.fraProductos.Enabled = False
    
    Call HabilitarComandos(True, True, False, False, False, True)

End Sub

Private Sub cmdGrabar_Click()
Dim sSQL1 As String
Dim rs As New ADODB.Recordset
Dim lnProd As Integer
Dim lsCodNivApr As String
Dim lsNroMaximo As String
Dim oCon As DConecta
Set oCon = New DConecta
                

If MsgBox("Desea Grabar el Registro de la Promocion ? ", vbQuestion + vbYesNo, "Aviso ") = 6 Then
   
   oCon.AbreConexion
   
   'Carga fecha y hora de grabación
    gdHoraGrab = Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm")
    If lsNueMod = "N" Then   ' Nuevo
         If ValPromocion(Me.txtbuscar.Text) Then
            MsgBox "La persona ya fue promocionada.", vbInformation, "Aviso"
            Exit Sub
         End If
         
         If Len(Me.txtbuscar.Text) = 0 Then
            MsgBox "No se han registrado los datos necesarios", vbInformation, "Aviso"
            Exit Sub
         End If
         sSQL1 = "Select Max(cPromocNro) Maximo From PersPromocion "
         Set rs = oCon.CargaRecordSet(sSQL1)
         If rs.BOF And rs.EOF Then
             lsNroMaximo = "000001"
         Else
             lsNroMaximo = FillNum(rs!Maximo + 1, 6, "0")
         End If
         rs.Close
        
         sSQL1 = "INSERT INTO PersPromocion (cPromocNro, cPersCod, cCodPromotor, dFecPromo, cComenta, nCapMonSoles, nCapMonDolares, nColMonSoles, nColMonDolares,cCodAge)" _
              & " VALUES ('" & lsNroMaximo & "','" & Me.txtbuscar.Text & "','" & Me.txtPromotor & _
              "','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm") & "','" & Trim(Me.txtComenta.Text) & "'," & Me.txtCapSol.value & "," & Me.txtCapDol.value & "," & Me.txtColSol.value & "," & Me.txtColDol.value & ",'" & Right(gsCodAge, 2) & "')"
         oCon.Ejecutar sSQL1
    Else   ' Modificacion
         lsNroMaximo = Me.lvwContactos.ListItems.Item(lvwContactos.SelectedItem.Index)
    
         sSQL1 = " UPDATE PersPromocion" _
               & " Set cComenta = '" & Trim(txtComenta.Text) & "'," _
               & " nCapMonSoles = " & Me.txtCapSol.value & ", nCapMonDolares = " & Me.txtCapDol.value & ", nColMonSoles = " & Me.txtColSol.value & ", nColMonDolares = " & Me.txtColDol.value & " " _
               & " Where cPromocNro = '" & lsNroMaximo & "' "
         oCon.Ejecutar sSQL1
    
         sSQL1 = "DELETE PersPromocionProd Where cPromocNro = '" & lsNroMaximo & "' "
         oCon.Ejecutar sSQL1
    End If
   
    For lnProd = 1 To 8
        If chkProd(lnProd).value = 1 Then
            lsProducto = chkProd(lnProd).Tag
            
            sSQL1 = "INSERT INTO PersPromocionProd (cPromocNro, cProducto, cComenta )" _
                 & " VALUES ('" & lsNroMaximo & "','" & lsProducto & "','' )"
            oCon.Ejecutar sSQL1
            
        End If
    Next lnProd
    
    Call HabilitarComandos(True, True, False, False, False, True)
    
    Me.fraPromotor.Enabled = True
    Me.fraContactos.Enabled = True
    Me.fraProductos.Enabled = False
    
    Limpiar
    LlenaListaContactos (txtPromotor)
End If

End Sub

Private Sub CmdImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    rtf.Text = ""
    ImprimeReporte
    oPrevio.Show rtf.Text, "REPORTE DE METAS POR ANALISTA", True, 66
End Sub

Private Sub cmdModificar_Click()
    lsNueMod = "M"
    Call HabilitarComandos(False, False, False, True, True, False)
    Me.fraPromotor.Enabled = False
    Me.fraContactos.Enabled = False
    Me.fraProductos.Enabled = True
End Sub

Private Sub CmdNuevo_Click()
    If txtPromotor.Text = "" Then
        MsgBox "Debe Ingresar un Promotor.", vbInformation, "Aviso"
        Me.txtPromotor.SetFocus
        Exit Sub
    End If
    
    lsNueMod = "N"
    Limpiar
    Call HabilitarComandos(False, False, False, True, True, False)
    Me.fraPromotor.Enabled = False
    Me.fraContactos.Enabled = False
    Me.fraProductos.Enabled = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call HabilitarComandos(True, True, False, False, False, True)
    Me.fraPromotor.Enabled = True
    Me.fraContactos.Enabled = True
    Me.fraProductos.Enabled = False
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'CierraConexion
End Sub

Private Sub Limpiar()
Dim lnCont As Integer
Me.txtbuscar.Text = ""
LblNomPers.Caption = ""
For lnCont = 1 To 8
    Me.chkProd(lnCont).value = 0
Next lnCont
txtComenta.Text = ""
End Sub

Private Sub HabilitarComandos(pcmdNuevo As Boolean, pcmdModificar As Boolean, _
                              pcmdeliminar As Boolean, pcmdGrabar As Boolean, pcmdCancelar As Boolean, pcmdImprimir As Boolean)
    Me.cmdNuevo.Enabled = pcmdNuevo
    Me.cmdModificar.Enabled = pcmdModificar
    Me.cmdGrabar.Enabled = pcmdGrabar
    Me.CmdCancelar.Enabled = pcmdCancelar
    Me.CmdImprimir.Enabled = pcmdImprimir
End Sub

Private Sub LlenaListaContactos(ByVal psPromotor As String)
Dim L As ListItem
Dim sSQLLista As String
Dim rsLista As New ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

  lvwContactos.ListItems.Clear
  If Len(psPromotor) = 0 Then
    Exit Sub
  End If
  
  sSQLLista = " SELECT Pro.cPromocNro, Per.cPersCod cCodPers, Per.cPersNombre cNomPers, Pro.cComenta, nCapMonSoles, nCapMonDolares, nColMonSoles, nColMonDolares   " _
            & " FROM PersPromocion Pro INNER JOIN Persona Per ON Pro.cPersCod = Per.cPersCod " _
            & " WHERE Pro.cCodPromotor = '" & psPromotor & "' Order By Pro.cPromocNro "  'and Datediff(d, dFecPromo, '" & Format(gdFecSis, gsformatofecha) & "' ) = 0
    Set rsLista = oCon.CargaRecordSet(sSQLLista)
    
    If rsLista.BOF And rsLista.EOF Then
       Exit Sub
    Else
       Do While Not rsLista.EOF
            Set L = lvwContactos.ListItems.Add(, , Trim(rsLista!cPromocNro))
                L.SubItems(1) = rsLista!cNomPers
                L.SubItems(2) = rsLista!cCodPers
                L.SubItems(3) = rsLista!cComenta
                L.SubItems(4) = rsLista!nCapMonSoles
                L.SubItems(5) = rsLista!nCapMonDolares
                L.SubItems(6) = rsLista!nColMonSoles
                L.SubItems(7) = rsLista!nColMonDolares
                
            rsLista.MoveNext
       Loop
    End If
    rsLista.Close
    Set rsLista = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub

Private Sub CabeceraReporte()
    rtf.Text = rtf.Text & ImpreFormat(gsNomCmac, 35) & Space(50) & "  Fecha  : " & Format(gdFecSis + Time, "dd/mm/yyyy hh:mm") & oImpresora.gPrnSaltoLinea
    rtf.Text = rtf.Text & ImpreFormat(gsNomAge, 35) & Space(50) & "Usuario: " & gsCodUser & "  " & oImpresora.gPrnSaltoLinea
    rtf.Text = rtf.Text & CentrarCadena("REGISTRO DE PROMOCION CON PROMOTORAS", 125) & oImpresora.gPrnSaltoLinea
    rtf.Text = rtf.Text & Space(2) & "Promotor : " & Trim(Me.lblProNombre.Caption) & oImpresora.gPrnSaltoLinea
    rtf.Text = rtf.Text & String(110, "-") & oImpresora.gPrnSaltoLinea
    rtf.Text = rtf.Text & "  Item  Promocion  Cliente " & Space(38) & " CAP.Sol.     CAP.Dol      COL.Sol     COL.Dol" & oImpresora.gPrnSaltoLinea
    rtf.Text = rtf.Text & String(110, "-") & oImpresora.gPrnSaltoLinea
End Sub

Private Sub ImprimeReporte()
    Dim lnLineas As Integer
    Dim lnPaginas As Integer
    Dim lnCont As Integer
    Dim lsNombre As String * 36
    Dim lsCapSol As String * 13
    Dim lsCapDol As String * 13
    Dim lsColSol As String * 13
    Dim lsColDol As String * 13
    Dim lsCad As String
    Dim lsCabecera As String
    Dim lnItem As Long
          lnLineas = 6
          
          CabeceraReporte
          
          lnItem = 1
          lsCabecera = rtf.Text
          
          lsCad = lsCad & lsCabecera
          For lnCont = 1 To Me.lvwContactos.ListItems.Count
      
            lnLineas = lnLineas + 1
            lsCad = lsCad & Space(3) & FillNum(Str(lnCont), 4, "0") & Space(3) & Me.lvwContactos.ListItems.Item(lnCont) & Space(3)
            lnItem = lnItem + 1
            
            lsNombre = lvwContactos.ListItems.Item(lnCont).SubItems(1)
            lsCad = lsCad & lsNombre & Space(3)
            
            RSet lsCapSol = Format(lvwContactos.ListItems.Item(lnCont).SubItems(4), "#,##0.00")
            RSet lsCapDol = Format(lvwContactos.ListItems.Item(lnCont).SubItems(5), "#,##0.00")
            RSet lsColSol = Format(lvwContactos.ListItems.Item(lnCont).SubItems(6), "#,##0.00")
            RSet lsColDol = Format(lvwContactos.ListItems.Item(lnCont).SubItems(7), "#,##0.00")
            
            lsCad = lsCad & lsCapSol
            lsCad = lsCad & lsCapDol
            lsCad = lsCad & lsColSol
            lsCad = lsCad & lsColDol
            
            lsCad = lsCad & oImpresora.gPrnSaltoLinea
            
            If lnItem = 60 Then
                lsCad = lsCad & oImpresora.gPrnSaltoPagina
                lsCad = lsCad & lsCabecera
                lsCad = lsCad & oImpresora.gPrnSaltoLinea
                lnItem = 0
            End If
            
          Next
          lsCad = lsCad & String(110, "-") & oImpresora.gPrnSaltoLinea
          rtf.Text = lsCad
End Sub

Private Sub lvwContactos_Click()
    If lvwContactos.ListItems.Count > 0 Then
        Limpiar
        MostrarDatos (lvwContactos.ListItems.Item(lvwContactos.SelectedItem.Index))
        Me.txtbuscar.Text = lvwContactos.SelectedItem.ListSubItems(2)
        Me.LblNomPers = lvwContactos.SelectedItem.ListSubItems(1)
        Me.txtComenta = lvwContactos.SelectedItem.ListSubItems(3)
        Me.txtCapSol.value = lvwContactos.SelectedItem.ListSubItems(4)
        Me.txtCapDol.value = lvwContactos.SelectedItem.ListSubItems(5)
        Me.txtColSol.value = lvwContactos.SelectedItem.ListSubItems(6)
        Me.txtColDol.value = lvwContactos.SelectedItem.ListSubItems(7)
    End If
End Sub

Private Sub MostrarDatos(ByVal psPromocNro As String)
    Dim L As ListItem
    Dim sSQLMet As String
    Dim rsMet As New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
      
      sSQLMet = "SELECT cProducto FROM PersPromocionProd  where cPromocNro ='" & psPromocNro & "' "
      Set rsMet = oCon.CargaRecordSet(sSQLMet)
      If rsMet.BOF And rsMet.EOF Then
          Exit Sub
      Else
          Do While Not rsMet.EOF
            Select Case rsMet!cProducto
                Case "A1"
                    chkProd(1).value = 1
                Case "A2"
                    chkProd(2).value = 1
                Case "A3"
                    chkProd(3).value = 1
                Case "B1"
                    chkProd(4).value = 1
                Case "B2"
                    chkProd(5).value = 1
                Case "B3"
                    chkProd(6).value = 1
                Case "B4"
                    chkProd(7).value = 1
                Case "B5"
                    chkProd(8).value = 1
            End Select
            
             rsMet.MoveNext
          Loop
      End If
       
      Set rsMet = Nothing
End Sub

Private Function ValPromocion(psCodPers As String) As Boolean
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Sql = " Select cPersCod cCodPers From PersPromocion Where cPersCod = '" & psCodPers & "'" _
        & " And '" & Format(gdFecSis, gsFormatoFecha) & "' Between  convert(varchar(10),dFecPromo,101) and convert(varchar(10),dateadd(day, 31 , dFecPromo) ,101)"
    
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(Sql)
    
    If rs.EOF And rs.BOF Then
        ValPromocion = False
    Else
        ValPromocion = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub txtBuscar_EmiteDatos()
    LblNomPers.Caption = txtbuscar.psDescripcion
End Sub

Private Sub txtCapDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtColSol.SetFocus
    End If
End Sub

Private Sub txtCapSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCapDol.SetFocus
    End If
End Sub

Private Sub txtColDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComenta.SetFocus
    End If
End Sub

Private Sub txtColSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtColDol.SetFocus
    End If
End Sub

Private Sub txtComenta_GotFocus()
    txtComenta.SelStart = 0
    txtComenta.SelLength = 250
End Sub

Private Sub txtComenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtPromotor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdNuevo.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtPromotor_LostFocus()
    Dim oAcceso As UAcceso
    Set oAcceso = New UAcceso
    
    Me.lblProNombre.Caption = oAcceso.MostarNombre(gsDominio, txtPromotor.Text)
    
    If Me.lblProNombre.Caption = "" Then
        MsgBox "Usuario no valido, ingrese un usuario correctamente.", vbInformation, "Aviso"
        txtPromotor = ""
        txtPromotor.SetFocus
    Else

        Limpiar
        ' Llena Lista con registros del Día
        LlenaListaContactos (txtPromotor.Text)
        'MostrarDatos (txtPromotor)
    End If
End Sub
