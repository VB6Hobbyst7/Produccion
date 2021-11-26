VERSION 5.00
Begin VB.Form frmLogSelActaBuenaPro 
   Caption         =   "Acta de Buena Pro"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   Icon            =   "frmLogSelActaBuenaPro.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2100
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtnombreplantilla 
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdgenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox cmbplantilla 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtdescripcionproveedor 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   600
      Width           =   3495
   End
   Begin Sicmact.TxtBuscar txtproveedor 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
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
      sTitulo         =   ""
   End
   Begin VB.TextBox txtnumeroproceso 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Acta de Buena Pro"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   645
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numero Proceso"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   1185
   End
End
Attribute VB_Name = "frmLogSelActaBuenaPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDGAdqui As DLogAdquisi
Dim rs As New ADODB.Recordset




Private Sub cmbplantilla_Click()
txtnombreplantilla.Text = clsDGAdqui.CargalogSelDescPlantilla(1, Right(cmbplantilla.Text, 1))

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdGenerar_Click()
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Dim Row As Integer
Dim Col As Integer
Dim q, a As Integer
If cmbplantilla.Text = "" Then Exit Sub
If txtProveedor.Text = "" Then Exit Sub
If txtnombreplantilla.Text = "" Then Exit Sub


 Set rs = clsDGAdqui.CargaLogSelProvBienes(frmLogSelEvalTecResumen.txtSeleccionA.Text, txtProveedor.Text)
If rs.EOF = True Then
    MsgBox "Este Proveedor no gano en ningun codigode bien ", vbInformation, "No gano en Ningun  codigo de bien "
    Exit Sub
End If


Set rs = clsDGAdqui.CargaLogSelComiteRep(frmLogSelEvalTecResumen.txtSeleccionA.Text)
If rs.EOF = True Then
    MsgBox "No se ha definido miebros para el comite  ", vbInformation, "No Existen Miembros del Comite"
    Exit Sub
End If
'Create an instance of Word
Set oWord = CreateObject("Word.Application")
'Show Word to the user
oWord.Visible = True
'Add a new, blank document
Set oDoc = oWord.Documents.Open(App.path & "\" + txtnombreplantilla.Text + ".doc")
'Get the current document's range object
'Store FlexGrid items to a two dimensional array
'Miembros del Comite
q = 1
'cPersNombre,cDescripcion
Do While Not rs.EOF
    With oWord.Selection.Find
            .Text = "CampNombre" + Str(q)
            .Replacement.Text = rs!cPersNombre
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    With oWord.Selection.Find
            .Text = "CampCargo" + Str(q)
            .Replacement.Text = rs!cDescripcion + " del Comite Especial"
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    rs.MoveNext
    q = q + 1
Loop
 
 For q = q To 7
    With oWord.Selection.Find
            .Text = "CampNombre" + Str(q)
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll

    With oWord.Selection.Find
            .Text = "CampCargo" + Str(q)
            .Replacement.Text = ""
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
 Next
 'Cargar
 Set rs = clsDGAdqui.CargaLogSelProvBienes(frmLogSelEvalTecResumen.txtSeleccionA.Text, txtProveedor.Text)
 a = 1
  Do While Not rs.EOF
    With oWord.Selection.Find
            If a > 9 Then
                .Text = "CodBienes" + Trim(Str(a))
            Else
                .Text = "CodBien" + Trim(Str(a))
            End If
            .Replacement.Text = rs!cBSCod
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    With oWord.Selection.Find
             If a > 9 Then
                .Text = "Descr" + Trim(Str(a))
             Else
                .Text = "Desc" + Trim(Str(a))
             End If
            .Replacement.Text = rs!cBSDescripcion
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    With oWord.Selection.Find
            If a > 9 Then
                .Text = "Canti" + Trim(Str(a))
            Else
                .Text = "Cant" + Trim(Str(a))
            End If
            .Replacement.Text = rs!nLogSelCotDetCantidad
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    With oWord.Selection.Find
             If a > 9 Then
            .Text = "Monto" + Trim(Str(a))
            Else
            .Text = "Mont" + Trim(Str(a))
            End If
            .Replacement.Text = Format(rs!Subtotal, "######.#0")
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    With oWord.Selection.Find
            If a > 9 Then
            .Text = "Tecn" + Trim(Str(a))
            Else
            .Text = "Tec" + Trim(Str(a))
            End If
            .Replacement.Text = Format(rs!Tecnico, "######.##")
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    With oWord.Selection.Find
            If a > 9 Then
            .Text = "Econ" + Trim(Str(a))
            Else
            .Text = "Eco" + Trim(Str(a))
            End If
            .Replacement.Text = Format(rs!Economico, "######.##")
            
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    With oWord.Selection.Find
            If a > 9 Then
            .Text = "Tota" + Trim(Str(a))
            Else
            .Text = "Tot" + Trim(Str(a))
            End If
            .Replacement.Text = Format(rs!PuntajeTotal, "######.##")
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    rs.MoveNext
    a = a + 1
Loop
 
For a = a To 15
            With oWord.Selection.Find
            If a > 9 Then
                .Text = "CodBienes" + Trim(Str(a))
            Else
                .Text = "CodBien" + Trim(Str(a))
            End If
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            End With
            oWord.Selection.Find.Execute Replace:=wdReplaceAll
            
            With oWord.Selection.Find
            If a > 9 Then
                .Text = "Descr" + Trim(Str(a))
            Else
                .Text = "Desc" + Trim(Str(a))
            End If
            .Replacement.Text = ""
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            End With
            
            oWord.Selection.Find.Execute Replace:=wdReplaceAll
            With oWord.Selection.Find
            If a > 9 Then
                .Text = "Canti" + Trim(Str(a))
            Else
                .Text = "Cant" + Trim(Str(a))
            End If
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            End With
            
            oWord.Selection.Find.Execute Replace:=wdReplaceAll
            With oWord.Selection.Find
            If a > 9 Then
            .Text = "Monto" + Trim(Str(a))
            Else
            .Text = "Mont" + Trim(Str(a))
            End If
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            End With
            
            oWord.Selection.Find.Execute Replace:=wdReplaceAll
            With oWord.Selection.Find
            If a > 9 Then
            .Text = "Tecn" + Trim(Str(a))
            Else
            .Text = "Tec" + Trim(Str(a))
            End If
            .Replacement.Text = ""
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            End With
            
            oWord.Selection.Find.Execute Replace:=wdReplaceAll
            With oWord.Selection.Find
            If a > 9 Then
            .Text = "Econ" + Trim(Str(a))
            Else
            .Text = "Eco" + Trim(Str(a))
            End If
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            End With
            
            oWord.Selection.Find.Execute Replace:=wdReplaceAll
            With oWord.Selection.Find
            If a > 9 Then
            .Text = "Tota" + Trim(Str(a))
            Else
            .Text = "Tot" + Trim(Str(a))
            End If
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            End With
            oWord.Selection.Find.Execute Replace:=wdReplaceAll
Next
 
 With oWord.Selection.Find
            .Text = "CampNumProceso"
            .Replacement.Text = frmLogSelEvalTecResumen.txtSeleccionA.Text
             .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
 With oWord.Selection.Find
            .Text = "CampDescripcion" 'Mid(string, start[, length])
            .Replacement.Text = Mid(frmLogSelEvalTecResumen.txttipo.Text, InStr(1, frmLogSelEvalTecResumen.txttipo.Text, "-"))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
 With oWord.Selection.Find
            .Text = "CampTipProceso" 'Mid(string, start[, length])
            .Replacement.Text = Left(frmLogSelEvalTecResumen.txttipo.Text, InStr(1, frmLogSelEvalTecResumen.txttipo.Text, "-") - 1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
 
Set rs = clsDGAdqui.CargaLogSelDescripcionProceso(frmLogSelEvalTecResumen.txtSeleccionA.Text)
 With oWord.Selection.Find
            .Text = "CampCotizacion"
            .Replacement.Text = rs!nLogSelNumeroCot
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
 With oWord.Selection.Find
            .Text = "CampProveedor"
            .Replacement.Text = txtDescripcionProveedor.Text
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
 oWord.Selection.Find.Execute Replace:=wdReplaceAll
 

End Sub

Private Sub Form_Load()
Me.Width = 7335
Me.Height = 2610

Set rs = New ADODB.Recordset
Set clsDGAdqui = New DLogAdquisi
txtnumeroproceso.Text = frmLogSelEvalTecResumen.txtSeleccionA.Text
txtnumeroproceso.Enabled = False
Me.txtProveedor.rs = clsDGAdqui.LogSeleccionListaProveedores(frmLogSelEvalTecResumen.txtSeleccionA.Text)
Set rs = clsDGAdqui.CargalogSelPlantilla(1)
Call CargaCombo(rs, cmbplantilla)
End Sub

Private Sub txtProveedor_EmiteDatos()
txtDescripcionProveedor.Text = txtProveedor.psDescripcion
End Sub
