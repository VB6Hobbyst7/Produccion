VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPasePersona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pase Persona"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   Icon            =   "frmPasePersona.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9705
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   6000
      Width           =   135
   End
   Begin VB.OptionButton optruc 
      Alignment       =   1  'Right Justify
      Caption         =   "Razon Social"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1815
   End
   Begin VB.OptionButton optruc 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero de Ruc"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtrazonsocial 
      Height          =   285
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   11
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtNumdoc 
      Height          =   285
      Left            =   2280
      MaxLength       =   12
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdgrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdizquierda 
      Caption         =   "Copiar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFSicmacI 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2990
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFSunat 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblgrabar 
      AutoSize        =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   5880
      Width           =   5085
   End
   Begin VB.Label lblcopiar 
      AutoSize        =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   3480
      Width           =   5205
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                                                                                     PERSONAS EN EL  SICMAC I"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   9495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                                                                                     PERSONAS EN LA SUNAT"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   9495
   End
End
Attribute VB_Name = "frmPasePersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oBuscaPersona As UPersona
Dim ClsPersona As DPersonas

Private Sub cmdBuscar_Click()
Dim R As ADODB.Recordset
Set R = New ADODB.Recordset

MSHFSunat.Clear
MSHFSunat.Rows = 2
MSHFSicmacI.Clear
Formato_grilla MSHFSunat
Formato_grilla MSHFSicmacI


If optruc(0).value = True Then

                txtNumdoc.Text = Trim(txtNumdoc.Text)
                If txtNumdoc.Text = "" Then
                    MsgBox "El Numero de documento no puede estar Vacio", vbInformation, "El documento No puede estar Vacio"
                    txtNumdoc.SetFocus
                    Exit Sub
                End If
                
                If IsNumeric(txtNumdoc.Text) = False Then
                    MsgBox "El Numero de documento No es Numerico", vbInformation, "Numero de documento No es Numerico"
                    txtNumdoc.SetFocus
                    Exit Sub
                End If
                
                
                If Len(txtNumdoc.Text) < 8 Then
                    MsgBox "El Numero de documento Ingresado es menor Que 8 Digitos", vbInformation, "Numero de documento es Menor que 8 Digitos"
                    txtNumdoc.SetFocus
                    Exit Sub
                End If
                
                If Len(txtNumdoc.Text) > 11 Then
                    MsgBox "El Numero de documento Ingresado es mayor Que 11 Digitos", vbInformation, "El Numero de documento es mayor que 11 digitos"
                    txtNumdoc.SetFocus
                    Exit Sub
                End If
                
                lblcopiar.Caption = " Se esta buscando persona por numero de  documento Numero " + txtNumdoc.Text
                MousePointer = vbHourglass
                Set R = ClsPersona.BuscaProveedor(txtNumdoc.Text, BusquedaDocumento)
 End If
 
 If optruc(1).value = True Then
 
                txtrazonsocial.Text = Trim(txtrazonsocial.Text)
                If txtrazonsocial.Text = "" Then
                    MsgBox "La Razon Social no puede estar vacio ", vbInformation, "La Razon Social no puede estar en blanco"
                    txtrazonsocial.SetFocus
                    Exit Sub
                End If
                If Len(txtrazonsocial.Text) < 4 Then
                    MsgBox "Ingrese mas de 3 digitos para la Razon Social", vbInformation, "Ingrese mas de tres digitos para la razon Social"
                    txtrazonsocial.SetFocus
                    Exit Sub
                End If
               lblcopiar.Caption = " Se esta buscando persona por numero de  documento Numero " + txtNumdoc.Text
                MousePointer = vbHourglass
                Set R = ClsPersona.BuscaProveedor(txtrazonsocial.Text, BusquedaNombre)
 
 End If
 
   
   
   
   
   If R.EOF = True Then
   MsgBox "El Cliente No  Se encuentra En la Base de Datos", vbInformation, "Cliente no encontrado"
   MSHFSunat.Clear
   Formato_grilla MSHFSunat
   MousePointer = vbArrow
   lblcopiar.Caption = ""
   Exit Sub
   Else
   Set MSHFSunat.DataSource = R
   Formato_grilla MSHFSunat
End If
MousePointer = vbArrow
lblcopiar.Caption = ""

End Sub

Private Sub cmdGrabar_Click()

Dim CodPersona As String



If MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, MSHFSicmacI.Col) = "" Then
   MsgBox "No Hay Codigo Seleccionado de Persona para  copiar", vbInformation, "Seleccione Codigo"
   Exit Sub
End If
MousePointer = vbHourglass
'Valida Por Codigo Persona
CodPersona = MSHFSicmacI.TextMatrix(1, 0)

If ClsPersona.ValidaCodPersona(CodPersona, 1) > 0 Then
    MsgBox "No se Puede Copiar el Codigo de Persona  " + MSHFSicmacI.TextMatrix(1, 0) + " ya existe  ", vbInformation, "No se Puede Copiar"
    MousePointer = vbArrow
    Exit Sub
End If
'Por NroDNI
If MSHFSicmacI.TextMatrix(1, 3) <> "" Or Len(MSHFSicmacI.TextMatrix(1, 3)) <> 0 Then
    If ClsPersona.ValidaCodPersona(MSHFSicmacI.TextMatrix(1, 3), 2) > 0 Then
        MsgBox "No se Puede Copiar el DNI  de la persona ya Existe en el SICMACI  " + MSHFSicmacI.TextMatrix(1, 3) + " Ya Existe  ", vbInformation, "Ya existe numero de DNI"
        MousePointer = vbArrow
        Exit Sub
    End If
End If
'Por NroRuc

If MSHFSicmacI.TextMatrix(1, 4) <> "" Or Len(MSHFSicmacI.TextMatrix(1, 4)) <> 0 Then
    If ClsPersona.ValidaCodPersona(MSHFSicmacI.TextMatrix(1, 4), 2) > 0 Then
        MsgBox "No se Puede Copiar El Numero de Ruc de la Persona ya existe en el SICMACI  " + MSHFSicmacI.TextMatrix(1, 3) + "  ", vbInformation, "Ya existe Numero de RUC"
        MousePointer = vbArrow
        Exit Sub
    End If
End If

lblgrabar.Caption = "Se Esta Grabando la Persona " + MSHFSicmacI.TextMatrix(1, 1)
'Grabar
If ClsPersona.InsertaProveedor(CodPersona) = 1 Then
    MsgBox "Se grabo Correctamente el Proveedor " + MSHFSicmacI.TextMatrix(1, 1), vbInformation, "Se copio Correctamente"
    MousePointer = vbArrow
    lblgrabar.Caption = ""
    Exit Sub
End If

lblgrabar.Caption = ""



End Sub

Private Sub cmdizquierda_Click()

If MSHFSunat.TextMatrix(MSHFSunat.Row, MSHFSunat.Col) = "" Then
   MsgBox "No Hay Codigo Seleccionado de Persona para  copiar", vbInformation, "Seleccione Codigo"
   Exit Sub
End If

MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 0) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 0))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 1) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 1))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 2) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 2))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 3) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 3))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 4) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 4))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 5) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 5))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 6) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 6))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 7) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 7))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 8) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 8))
MSHFSicmacI.TextMatrix(MSHFSicmacI.Row, 9) = Trim(MSHFSunat.TextMatrix(MSHFSunat.Row, 9))

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
frmTipoCambio.Show 1
End Sub

Private Sub Form_Activate()
optruc(0).value = True
End Sub

Private Sub Form_Load()
Me.Width = 9825
Me.Height = 6840
Set oBuscaPersona = New UPersona
Set ClsPersona = New DPersonas
MSHFSunat.Cols = 10
MSHFSicmacI.Cols = 10
Formato_grilla MSHFSunat
Formato_grilla MSHFSicmacI
'optruc(0).value = True
End Sub


Sub Formato_grilla(grilla As MSHFlexGrid)
grilla.ColWidth(0) = 1300 'Cod Persona
grilla.ColWidth(1) = 3000 'Nombre o razon Social
grilla.ColWidth(2) = 800 'Tipo Doc
grilla.ColWidth(3) = 1200 'NumeroDNI
grilla.ColWidth(4) = 1200 'NumeroRUC
grilla.ColWidth(5) = 2000 'Fecha Creacion
grilla.ColWidth(6) = 2000 'Direccion
grilla.ColWidth(7) = 1000 'Telefono
grilla.ColWidth(8) = 800 'Zona
grilla.ColWidth(9) = 800 'Pers.Nat.Sexo





grilla.TextMatrix(0, 0) = "Cod Persona"
grilla.TextMatrix(0, 1) = "Nombre o razon Social"
grilla.TextMatrix(0, 2) = "Tipo Doc"
grilla.TextMatrix(0, 3) = "NumeroDNI"
grilla.TextMatrix(0, 4) = "NumeroRUC"
grilla.TextMatrix(0, 5) = "Fecha Creacion"
grilla.TextMatrix(0, 6) = "Direccion"
grilla.TextMatrix(0, 7) = "Telefono"
grilla.TextMatrix(0, 8) = "Zona"
grilla.TextMatrix(0, 9) = "sexo"






End Sub

Private Sub optruc_Click(Index As Integer)
Select Case Index

Case 0 'RUC
        If optruc(0).value = True Then
            txtNumdoc.Enabled = True
            txtNumdoc.SetFocus
            txtrazonsocial.Enabled = False
            txtrazonsocial.Text = ""
            
        Else
            txtNumdoc.Enabled = False
            txtNumdoc.Text = ""
            txtrazonsocial.Enabled = True
        End If

Case 1 'RAZON  SOCIAL
        If optruc(1).value = True Then
            txtrazonsocial.Enabled = True
            txtrazonsocial.SetFocus
            txtNumdoc.Enabled = False
            txtNumdoc.Text = ""
        Else
            txtrazonsocial.Enabled = False
            txtrazonsocial.Text = ""
            
            txtNumdoc.Enabled = True
        End If
End Select

End Sub
