VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPCatAnalista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Categorías de Analistas"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   Icon            =   "frmCredBPPCatAnalista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTCategorias 
      Height          =   5010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8837
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Categorías"
      TabPicture(0)   =   "frmCredBPPCatAnalista.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraEdicionCategorias"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCerrar"
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
         Left            =   7800
         TabIndex        =   10
         Top             =   4440
         Width           =   1170
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
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
         Left            =   9120
         TabIndex        =   9
         Top             =   4440
         Width           =   1170
      End
      Begin VB.Frame fraEdicionCategorias 
         Caption         =   "Edicción de Categorías"
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
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10335
         Begin VB.Frame fraCategoria 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   125
            TabIndex        =   5
            Top             =   910
            Width           =   1530
            Begin VB.Label lblCategoria 
               AutoSize        =   -1  'True
               Caption         =   "Categoría"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   360
               TabIndex        =   6
               Top             =   320
               Width           =   825
            End
         End
         Begin VB.ComboBox cmbNiveles 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   480
            Width           =   2175
         End
         Begin SICMACT.FlexEdit feCategorias 
            Height          =   2415
            Left            =   120
            TabIndex        =   2
            Top             =   1320
            Width           =   9975
            _extentx        =   17595
            _extenty        =   4260
            cols0           =   7
            highlight       =   1
            encabezadosnombres=   "#-Categoria-Min-Max-Min-Max-Aux"
            encabezadosanchos=   "0-1500-2000-2000-2000-2000-0"
            font            =   "frmCredBPPCatAnalista.frx":0326
            font            =   "frmCredBPPCatAnalista.frx":034E
            font            =   "frmCredBPPCatAnalista.frx":0376
            font            =   "frmCredBPPCatAnalista.frx":039E
            font            =   "frmCredBPPCatAnalista.frx":03C6
            fontfixed       =   "frmCredBPPCatAnalista.frx":03EE
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-2-3-4-5-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-3-3-0"
            cantentero      =   15
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label lblNumClientes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Número de Clientes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5630
            TabIndex        =   8
            Top             =   1000
            Width           =   4005
         End
         Begin VB.Label lblSaldoCartera 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo Cartera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1640
            TabIndex        =   7
            Top             =   1000
            Width           =   4000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   520
            Width           =   405
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPCatAnalista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************************
''***     Rutina         :   frmCredBPPCatAnalista
''***     Descripcion    :   Configurar Categoría de Analista
''***     Creado por     :   WIOR
''***     Maquina        :   TIF-1-19
''***     Fecha-Creación :   24/05/2013 08:20:00 AM
''*****************************************************************************************
'Option Explicit
'Dim fbSalCartera As Boolean
'Dim fbNumClientes As Boolean
'
'Private Sub cmbNiveles_Click()
'Call CargaDatos(CInt(Trim(Right(cmbNiveles.Text, 4))))
'End Sub
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub cmdGuardar_Click()
'On Error GoTo Error
'If ValidaDatos Then
'    If MsgBox("Estas seguro de Guardar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim i As Integer
'        Dim nNivel As Integer
'
'        Dim oBPP As COMDCredito.DCOMBPPR
'        Set oBPP = New COMDCredito.DCOMBPPR
'        nNivel = CInt(Trim(Right(cmbNiveles.Text, 5)))
'
'        Call oBPP.OpeCatAnalistaXNivel(2, nNivel)
'
'         For i = 0 To feCategorias.Rows - 2
'            Call oBPP.OpeCatAnalistaXNivel(1, nNivel, Trim(feCategorias.TextMatrix(i + 1, 1)), _
'            CDbl(IIf(Trim(feCategorias.TextMatrix(i + 1, 2)) = "", "-1", Trim(feCategorias.TextMatrix(i + 1, 2)))), _
'            CDbl(IIf(Trim(feCategorias.TextMatrix(i + 1, 3)) = "", "-1", Trim(feCategorias.TextMatrix(i + 1, 3)))), _
'            CDbl(IIf(Trim(feCategorias.TextMatrix(i + 1, 4)) = "", "-1", Trim(feCategorias.TextMatrix(i + 1, 4)))), _
'            CDbl(IIf(Trim(feCategorias.TextMatrix(i + 1, 5)) = "", "-1", Trim(feCategorias.TextMatrix(i + 1, 5)))))
'         Next i
'
'        MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'        CargaDatos (nNivel)
'    End If
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'
'
'Private Sub Form_Load()
'Call CargaControles
'End Sub
'
'Private Sub CargaControles()
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim rsConst As ADODB.Recordset
'
''CARGA NIVELES
'Set oConst = New COMDConstantes.DCOMConstantes
'Set rsConst = oConst.RecuperaConstantes(7064)
'
''CARGA COMBO DE NIVELES
'Call Llenar_Combo_con_Recordset(rsConst, cmbNiveles)
'cmbNiveles.ListIndex = IndiceListaCombo(cmbNiveles, "1")
'End Sub
'
'Private Sub CargaDatos(ByVal pnNivel As Integer)
'Dim oDBPP As COMDCredito.DCOMBPPR
'Dim oNBPP As COMNCredito.NCOMBPPR
'Dim oDBPPDatos As COMDCredito.DCOMBPPR
'
'Dim rsDBPP As ADODB.Recordset
'Dim rsDBPPDatos As ADODB.Recordset
'
'Dim i As Integer
'
'fbSalCartera = False
'fbNumClientes = False
'
'Set oNBPP = New COMNCredito.NCOMBPPR
'Call oNBPP.ObtenerParamCabXNivel(pnNivel, fbSalCartera, fbNumClientes)
'
'If fbSalCartera Then
'    If fbNumClientes Then
'        'SALDO CARTERA Y NUMERO DE CLIENTES HABILITADOS
'        feCategorias.ColumnasAEditar = "X-X-2-3-4-5-X"
'    Else
'        'SALDO CARTERA HABILITADOS
'        feCategorias.ColumnasAEditar = "X-X-2-3-X-X-X"
'    End If
'Else
'    If fbNumClientes Then
'        'NUMERO DE CLIENTES HABILITADOS
'        feCategorias.ColumnasAEditar = "X-X-X-X-4-5-X"
'    Else
'        'TODOS DESHABILITADOS
'        feCategorias.ColumnasAEditar = "X-X-X-X-X-X-X"
'    End If
'End If
'
'Call LimpiaFlex(feCategorias)
'Set oDBPP = New COMDCredito.DCOMBPPR
'Set rsDBPP = oDBPP.ObtenerCatOParamCabXNivel(1, pnNivel)
'
'
'Set oDBPPDatos = New COMDCredito.DCOMBPPR
'
'If Not (rsDBPP.BOF And rsDBPP.EOF) Then
'    For i = 0 To rsDBPP.RecordCount - 1
'        feCategorias.AdicionaFila
'        feCategorias.TextMatrix(i + 1, 0) = i + 1
'        feCategorias.TextMatrix(i + 1, 1) = Trim(rsDBPP!cCategoria)
'
'        Set rsDBPPDatos = oDBPPDatos.ObtenerCatAnalistaXNivel(pnNivel, Trim(rsDBPP!cCategoria))
'        If Not (rsDBPPDatos.BOF And rsDBPPDatos.EOF) Then
'
'            feCategorias.TextMatrix(i + 1, 2) = Format(CDbl(rsDBPPDatos!nSalCarteraMin), "###," & String(15, "#") & "#0." & String(2, "0"))
'            feCategorias.TextMatrix(i + 1, 3) = Format(CDbl(rsDBPPDatos!nSalCarteraMax), "###," & String(15, "#") & "#0." & String(2, "0"))
'
'            If CDbl(Trim(rsDBPPDatos!nNumClientesMin)) = -1 Then
'                feCategorias.TextMatrix(i + 1, 4) = ""
'            Else
'                feCategorias.TextMatrix(i + 1, 4) = Format(CDbl(rsDBPPDatos!nNumClientesMin), "#," & String(15, "#") & "#0")
'            End If
'
'            If CDbl(Trim(rsDBPPDatos!nNumClientesMax)) = -1 Then
'                feCategorias.TextMatrix(i + 1, 5) = ""
'            Else
'                feCategorias.TextMatrix(i + 1, 5) = Format(CDbl(rsDBPPDatos!nNumClientesMax), "#," & String(15, "#") & "#0")
'            End If
'
'        Else
'            feCategorias.TextMatrix(i + 1, 2) = ""
'            feCategorias.TextMatrix(i + 1, 3) = ""
'            feCategorias.TextMatrix(i + 1, 4) = ""
'            feCategorias.TextMatrix(i + 1, 5) = ""
'        End If
'
'        rsDBPP.MoveNext
'    Next i
'Else
'
'
'    MsgBox "No Hay Datos.", vbInformation, "Aviso"
'End If
'
'End Sub
'
'Private Function ValidaDatos() As Boolean
'Dim i As Integer
'
'If Trim(feCategorias.TextMatrix(1, 1)) <> "" Then
'    For i = 0 To feCategorias.Rows - 2
'        If fbSalCartera Then
'            If Trim(feCategorias.TextMatrix(i + 1, 2)) = "" Or Trim(feCategorias.TextMatrix(i + 1, 3)) = "" Then
'                MsgBox "Ingrese Completamente los parámetros para Saldo de Cartera en la Categoría ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "''.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If CDbl(Trim(feCategorias.TextMatrix(i + 1, 2))) > CDbl(Trim(feCategorias.TextMatrix(i + 1, 3))) Then
'                MsgBox "Saldo Minimo de Cartera no puede ser Mayor que el Saldo Maximo de Cartera  en la Categoría ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "''.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'        Else
'            feCategorias.TextMatrix(i + 1, 2) = ""
'            feCategorias.TextMatrix(i + 1, 3) = ""
'        End If
'
'        If fbNumClientes Then
'            If Trim(feCategorias.TextMatrix(i + 1, 4)) = "" Or Trim(feCategorias.TextMatrix(i + 1, 5)) = "" Then
'                MsgBox "Ingrese Completamente los parámetros para Número de Clientes en la Categoría ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "''.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If CDbl(Trim(feCategorias.TextMatrix(i + 1, 4))) > CDbl(Trim(feCategorias.TextMatrix(i + 1, 5))) Then
'                MsgBox "Números Mínimo de Clientes no puede ser Mayor que el Números Máximo de Clientes  en la Categoría ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "''.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If CDbl(Trim(feCategorias.TextMatrix(i + 1, 4))) > 99999999 Then
'                MsgBox "El Valor del Número Minimo de Clientes Supera el Valor Máximo Permitido en la Categoría ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "''.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If CDbl(Trim(feCategorias.TextMatrix(i + 1, 5))) > 99999999 Then
'                MsgBox "El Valor del Número Máximo de Clientes Supera el Valor Máximo Permitido en la Categoría''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "''.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'        Else
'            feCategorias.TextMatrix(i + 1, 4) = ""
'            feCategorias.TextMatrix(i + 1, 5) = ""
'        End If
'    Next i
'
'    For i = 0 To feCategorias.Rows - 3
'        If fbSalCartera Then
'            If CCur(CDbl(Trim(feCategorias.TextMatrix(i + 2, 2))) - CDbl(Trim(feCategorias.TextMatrix(i + 1, 3)))) <> 0.01 Then
'                MsgBox "Ingrese Correctamente los parámetros para Saldo de Cartera en las Categorías ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "'' y ''" & Trim(feCategorias.TextMatrix(i + 2, 1)) & "''.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'        End If
'
'        If fbNumClientes Then
'            If CCur(CDbl(Trim(feCategorias.TextMatrix(i + 2, 4))) - CDbl(Trim(feCategorias.TextMatrix(i + 1, 5)))) <> 1 Then
'                MsgBox "Ingrese Correctamente los parámetros para Número de Clientes en las Categorías ''" & Trim(feCategorias.TextMatrix(i + 1, 1)) & "'' y ''" & Trim(feCategorias.TextMatrix(i + 2, 1)) & "''.", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'        End If
'    Next i
'Else
'    MsgBox "No hay Datos a Guardar", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'
'ValidaDatos = True
'End Function
