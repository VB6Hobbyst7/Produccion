VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPigProyeccionGuia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proyección de la Guía"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmPigProyeccionGuia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7335
      TabIndex        =   22
      Top             =   4080
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   5910
      TabIndex        =   21
      Top             =   4080
      Width           =   1245
   End
   Begin VB.CommandButton cmdProyectar 
      Caption         =   "&Proyectar"
      Height          =   390
      Left            =   4500
      TabIndex        =   20
      Top             =   4080
      Width           =   1230
   End
   Begin VB.Frame fraSeleccion 
      Enabled         =   0   'False
      Height          =   2850
      Left            =   60
      TabIndex        =   5
      Top             =   1095
      Width           =   8595
      Begin VB.Frame Frame4 
         Caption         =   "Datos de la Pieza"
         Height          =   2595
         Left            =   4380
         TabIndex        =   7
         Top             =   165
         Width           =   4110
         Begin VB.Label lblValorAdj 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   19
            Top             =   2130
            Width           =   1425
         End
         Begin VB.Label lblRemate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   18
            Top             =   1755
            Width           =   1425
         End
         Begin VB.Label lblPesoNeto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1320
            TabIndex        =   17
            Top             =   1410
            Width           =   1425
         End
         Begin VB.Label lblPesoBruto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   16
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label lblMaterial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1305
            TabIndex        =   15
            Top             =   675
            Width           =   2715
         End
         Begin VB.Label lblDescripcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1305
            TabIndex        =   14
            Top             =   300
            Width           =   2715
         End
         Begin VB.Label Label11 
            Caption         =   "Valor Adj."
            Height          =   240
            Left            =   270
            TabIndex        =   13
            Top             =   2160
            Width           =   990
         End
         Begin VB.Label Label10 
            Caption         =   "Remate"
            Height          =   240
            Left            =   255
            TabIndex        =   12
            Top             =   1785
            Width           =   990
         End
         Begin VB.Label Label9 
            Caption         =   "Peso Neto"
            Height          =   240
            Left            =   270
            TabIndex        =   11
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label Label8 
            Caption         =   "Peso Bruto"
            Height          =   240
            Left            =   270
            TabIndex        =   10
            Top             =   1095
            Width           =   990
         End
         Begin VB.Label Label7 
            Caption         =   "Material"
            Height          =   240
            Left            =   285
            TabIndex        =   9
            Top             =   720
            Width           =   990
         End
         Begin VB.Label Label6 
            Caption         =   "Descripción"
            Height          =   240
            Left            =   255
            TabIndex        =   8
            Top             =   360
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Selección"
         Height          =   2595
         Left            =   105
         TabIndex        =   6
         Top             =   165
         Width           =   4170
         Begin VB.CommandButton cmdSeleccionar 
            Caption         =   "Seleccionar"
            Height          =   345
            Left            =   2865
            TabIndex        =   39
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtCuenta 
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
            Height          =   285
            Left            =   2460
            MaxLength       =   10
            TabIndex        =   38
            Top             =   1770
            Width           =   1170
         End
         Begin VB.TextBox txtAge 
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
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   37
            Top             =   1770
            Width           =   345
         End
         Begin VB.TextBox txtCmac 
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
            Height          =   285
            Left            =   1245
            TabIndex        =   36
            Top             =   1770
            Width           =   435
         End
         Begin VB.TextBox txtPieza 
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
            Height          =   285
            Left            =   3615
            MaxLength       =   2
            TabIndex        =   35
            Top             =   1770
            Width           =   345
         End
         Begin VB.TextBox txtProducto 
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
            Height          =   285
            Left            =   2025
            TabIndex        =   34
            Top             =   1770
            Width           =   435
         End
         Begin VB.TextBox txtRemate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1275
            TabIndex        =   32
            Top             =   480
            Width           =   1410
         End
         Begin VB.CheckBox chkPieza 
            Appearance      =   0  'Flat
            Caption         =   "Pieza"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   210
            TabIndex        =   30
            Top             =   1800
            Width           =   1050
         End
         Begin VB.CheckBox chkMaterial 
            Appearance      =   0  'Flat
            Caption         =   "Material"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   195
            TabIndex        =   29
            Top             =   1365
            Width           =   945
         End
         Begin VB.CheckBox chkRemate 
            Appearance      =   0  'Flat
            Caption         =   "Remate"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   225
            TabIndex        =   28
            Top             =   510
            Width           =   960
         End
         Begin VB.CheckBox chkAgencia 
            Appearance      =   0  'Flat
            Caption         =   "Agencia"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   210
            TabIndex        =   27
            Top             =   930
            Width           =   915
         End
         Begin MSDataListLib.DataCombo cboAgencia 
            Height          =   315
            Left            =   1260
            TabIndex        =   31
            Top             =   885
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboMaterial 
            Height          =   315
            Left            =   1260
            TabIndex        =   33
            Top             =   1320
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      Begin MSDataListLib.DataCombo cboMotivo 
         Height          =   315
         Left            =   765
         TabIndex        =   23
         Top             =   270
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboOrigen 
         Height          =   315
         Left            =   765
         TabIndex        =   24
         Top             =   660
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboClase 
         Height          =   315
         Left            =   5550
         TabIndex        =   25
         Top             =   285
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboDestino 
         Height          =   315
         Left            =   5550
         TabIndex        =   26
         Top             =   690
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label4 
         Caption         =   "Destino"
         Height          =   255
         Left            =   4905
         TabIndex        =   4
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Clase"
         Height          =   255
         Left            =   4905
         TabIndex        =   3
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Origen"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   705
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmPigProyeccionGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Dim lbProyectar As Boolean

Private Sub cboClase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboOrigen.SetFocus
    End If
End Sub

Private Sub cboDestino_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        If cboOrigen.BoundText = cboDestino.BoundText Then
            MsgBox "El Destino no puede ser el mismo que el Origen", vbInformation, "Aviso"
        End If
        
        If cboClase.BoundText = 1 Then
            fraSeleccion.Enabled = False
            If chkRemate.Enabled Then chkRemate.SetFocus
        ElseIf cboClase.BoundText = 2 Then
            fraSeleccion.Enabled = True
            If cmdProyectar.Enabled Then cmdProyectar.SetFocus
        End If
        
     End If
     
End Sub

Private Sub cboMotivo_Click(Area As Integer)
    
    cboClase.Enabled = True
    cboDestino.Enabled = True
    cboOrigen.Enabled = True
    
    cboClase.BoundText = -1
    cboDestino.BoundText = -1
    cboOrigen.BoundText = -1
    
    If cboMotivo.BoundText <> "" Then
        Select Case cboMotivo.BoundText
        Case 1 'Contratos nuevos
            cboClase.BoundText = 1
            cboDestino.BoundText = 99
            cboOrigen.BoundText = CInt(gsCodAge)
            cboClase.Enabled = False
            cboDestino.Enabled = False
            cboOrigen.Enabled = False
            fraSeleccion.Enabled = False
        Case 2
            cboClase.BoundText = 1
            cboDestino.BoundText = 99
            cboClase.Enabled = False
            cboDestino.Enabled = False
            fraSeleccion.Enabled = False
        Case 3
            cboClase.BoundText = 1
            cboOrigen.BoundText = 99
            If Not ValidaAgBovedaValores Then
                MsgBox "Usuario no se encuentra en la Agencia de la Boveda de Valores", vbInformation, "Aviso"
                cboOrigen.BoundText = 1
                Exit Sub
            End If
            cboClase.Enabled = False
            cboOrigen.Enabled = False
            fraSeleccion.Enabled = False
        Case 4 'Remate
            cboClase.BoundText = 2
            cboOrigen.BoundText = 99
            If Not ValidaAgBovedaValores Then
                MsgBox "Usuario no se encuentra en la Agencia de la Boveda de Valores", vbInformation, "Aviso"
                cboOrigen.BoundText = 1
                Exit Sub
            End If
            cboClase.Enabled = False
            cboOrigen.Enabled = False
            fraSeleccion.Enabled = False
        Case 5
            cboClase.BoundText = 2
            cboDestino.BoundText = 99
            cboClase.Enabled = False
            cboDestino.Enabled = False
            fraSeleccion.Enabled = False
        Case 6
            cboClase.BoundText = 2
            cboOrigen.BoundText = 99
            If Not ValidaAgBovedaValores Then
                MsgBox "Usuario no se encuentra en la Agencia de la Boveda de Valores", vbInformation, "Aviso"
                cboOrigen.BoundText = 1
                Exit Sub
            End If
            cboClase.Enabled = False
            cboOrigen.Enabled = False
            fraSeleccion.Enabled = True
        Case 7
            cboClase.BoundText = 2
            cboDestino.BoundText = 99
            cboClase.Enabled = False
            cboDestino.Enabled = False
            fraSeleccion.Enabled = True
        Case 8
            cboClase.BoundText = 2
            cboOrigen.BoundText = 99
            If Not ValidaAgBovedaValores Then
                MsgBox "Usuario no se encuentra en la Agencia de la Boveda de Valores", vbInformation, "Aviso"
                cboOrigen.BoundText = 1
                Exit Sub
            End If
            cboClase.Enabled = False
            cboOrigen.Enabled = False
            fraSeleccion.Enabled = True
        Case 9
            cboClase.BoundText = 2
            cboDestino.BoundText = 99
            cboClase.Enabled = False
            cboDestino.Enabled = False
            fraSeleccion.Enabled = True
        End Select
    End If

End Sub

Private Sub cboMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        cboClase.Enabled = True
        cboDestino.Enabled = True
        cboOrigen.Enabled = True
        
        cboClase.BoundText = -1
        cboDestino.BoundText = -1
        cboOrigen.BoundText = -1
        
        
        If cboMotivo.BoundText <> "" Then
            Select Case cboMotivo.BoundText
            Case 1
                cboClase.BoundText = 1
                cboDestino.BoundText = 99
                cboOrigen.BoundText = CInt(gsCodAge)
                cboClase.Enabled = False
                cboDestino.Enabled = False
                cboOrigen.Enabled = False
            Case 2
                cboClase.BoundText = 1
                cboDestino.BoundText = 99
                cboClase.Enabled = False
                cboDestino.Enabled = False
            Case 3
                cboClase.BoundText = 1
                cboOrigen.BoundText = 99
                cboClase.Enabled = False
                cboOrigen.Enabled = False
            Case 4
                cboClase.BoundText = 2
                cboOrigen.BoundText = 99
                cboClase.Enabled = False
                cboOrigen.Enabled = False
                fraSeleccion.Enabled = True
            Case 5
                cboClase.BoundText = 2
                cboDestino.BoundText = 99
                cboClase.Enabled = False
                cboDestino.Enabled = False
                fraSeleccion.Enabled = True
            Case 6
                cboClase.BoundText = 2
                cboOrigen.BoundText = 99
                cboClase.Enabled = False
                cboOrigen.Enabled = False
                fraSeleccion.Enabled = True
            Case 7
                cboClase.BoundText = 2
                cboDestino.BoundText = 99
                cboClase.Enabled = False
                cboDestino.Enabled = False
                fraSeleccion.Enabled = True
            End Select
        End If
        cboClase.SetFocus
    End If
End Sub

Private Sub cboOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboDestino.Enabled Then cboDestino.SetFocus
    End If
End Sub

Private Sub chkAgencia_Click()
    If chkAgencia.value = 1 Then
        Call CargaCombo(cboAgencia, gColocPigUbicacion)
        cboAgencia.Visible = True
        cboAgencia.SetFocus
    End If
End Sub

Private Sub chkMaterial_Click()
    If chkMaterial.value = 1 Then
        Call CargaCombo(cboMaterial, gColocPigMaterial)
        cboMaterial.Visible = True
        cboMaterial.SetFocus
    End If
End Sub

Private Sub chkPieza_Click()
    If chkPieza.value = 1 Then
        txtCmac = gsCodCMAC
        txtProducto = "305"
        txtAge.Enabled = True
        txtCuenta.Enabled = True
        txtpieza.Enabled = True
        txtAge.SetFocus
    End If
End Sub

Private Sub chkRemate_Click()
    If chkRemate.value = 1 Then
        txtRemate.Visible = True
        txtRemate.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
End Sub

Private Sub cmdProyectar_Click()

Dim oPigFunc As DPigFunciones
Dim oPigDatos As dPigContrato
Dim oPigGrabar As DPigActualizaBD
Dim lsNumGuia As String
Dim rs As Recordset
Dim lsEstados As String
Dim lsMovNro As String
Dim lnDiasCustodia As Integer
Dim oContFunc As NContFunciones
Dim I As Integer
Dim lsEstado As String
Dim lnRemate As Long
Dim lnAgencia As Integer
Dim lnMaterial As Integer
Dim lspieza As String

Set oPigFunc = New DPigFunciones

If ValidaProyectar Then

    Select Case cboMotivo.BoundText
    
    Case 1      'Lote   Contratos Nuevos  **** Agencia X - Boveda
        lsEstado = "2802"
        lnDiasCustodia = -1
    Case 2      'Lote   Pendientes de Rescate > 30 dias  **** Agencia X - Boveda
        lsEstado = "2805"
        lnDiasCustodia = oPigFunc.GetParamValor(8020)
    Case 3      'Lote
        lsEstado = "2805, 2815"     'Pendientes de Rescate   **** Boveda -  Agencia
        lnDiasCustodia = -1
    Case 4      'Pieza          'Remate (Vencidos y Pend de Rescate > 30 dias)   **** Boveda - Remate
        lsEstado = "2807, 2808"
    Case 5      'Pieza          'Devolucion de Piezas (Rem/Cancel/Adju) **** Remate - Boveda
        lsEstado = "2803, 2804, 2805, 2807, 2808, 2809"
    Case 6      'Pieza      'Joyas para Venta   **** Boveda - Tienda
        lsEstado = "2812"
    Case 7, 8, 9     'Pieza      'Devolucion de Joyas a Boveda   **** Tienda - Boveda
        lsEstado = "2812"
    End Select
       
    If cboClase.BoundText = 1 Then  'LOTE - VALOR DE ADJUDICACION = 0.00
    
        Set oPigDatos = New dPigContrato
        
        Set rs = oPigDatos.dObtieneContratosColocPigGuia(cboOrigen.BoundText, lsEstado, gdFecSis, lnDiasCustodia, cboMotivo.BoundText, cboDestino.BoundText)
                
        If Not rs.EOF And Not rs.BOF Then
                    
            If rs!nItem > 0 Then
                
                lsNumGuia = oPigFunc.GetNumGuia(cboOrigen.BoundText)
                Set oPigFunc = Nothing
                
                Set oContFunc = New NContFunciones
                lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Set oContFunc = Nothing
                
                Set oPigGrabar = New DPigActualizaBD
                
                oPigGrabar.dBeginTrans
                
                Call oPigGrabar.dInsertColocPigGuia(lsNumGuia, cboMotivo.BoundText, cboOrigen.BoundText, cboDestino.BoundText, _
                        1, rs!nItem, rs!nPBruto, rs!nPNeto, rs!nTasacion, 0, lsMovNro, False)
                    
                Set rs = Nothing
                
                Set rs = oPigDatos.dObtieneContratosColocPigGuiaDet(cboOrigen.BoundText, lsEstado, gdFecSis, lnDiasCustodia, cboMotivo.BoundText, cboDestino.BoundText)
                I = 1
                
                Do While Not rs.EOF
                    Call oPigGrabar.dInsertColocPigGuiaDet(lsNumGuia, I, rs!cCtaCod, 0, rs!npiezas, False)
                    I = I + 1
                    rs.MoveNext
                Loop
                
                Call oPigGrabar.dInsertColocPigGuiaEtapa(lsNumGuia, gPigEtapaRemProyec, lsMovNro, False)
                
                oPigGrabar.dCommitTrans
                
                Set oPigGrabar = Nothing
                
                MsgBox "Proyección de la Guía finalizó satisfactoriamente", vbInformation, "Aviso"
                Limpiar
            Else
                MsgBox "No existen Contratos para la Remesa", vbInformation, "Aviso"
            End If
        Else
            MsgBox "No existen Contratos para la Remesa", vbInformation, "Aviso"
        End If
        
        Set rs = Nothing
        Set oPigDatos = Nothing
        
    ElseIf cboClase.BoundText = 2 Then 'PIEZA
    
        Select Case cboMotivo.BoundText
        
        Case 4, 5 'SELECCION AUTOMATICA - VALOR DE ADJUDICACION = 0.00
    
            Set oPigDatos = New dPigContrato
            Set rs = oPigDatos.dObtienePiezasColocPigGuia(cboOrigen.BoundText, lsEstado, gdFecSis, cboMotivo.BoundText)
            
            
            If Not rs.EOF And Not rs.BOF Then
                If rs!nItem > 0 Then
                
                    lsNumGuia = oPigFunc.GetNumGuia(cboOrigen.BoundText)
                    Set oPigFunc = Nothing
                    
                    Set oContFunc = New NContFunciones
                    lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Set oContFunc = Nothing
                    
                    Set oPigGrabar = New DPigActualizaBD
                    
                    oPigGrabar.dBeginTrans
                    
                    Call oPigGrabar.dInsertColocPigGuia(lsNumGuia, cboMotivo.BoundText, cboOrigen.BoundText, cboDestino.BoundText, _
                            2, rs!nItem, rs!nPBruto, rs!nPNeto, rs!nTasacion, 0, lsMovNro, False)
                        
                    Set rs = Nothing
                    
                    Set rs = oPigDatos.dObtienePiezasColocPigGuiaDet(cboOrigen.BoundText, lsEstado, gdFecSis)
                    I = 1
                    
                    Do While Not rs.EOF
                        Call oPigGrabar.dInsertColocPigGuiaDet(lsNumGuia, I, rs!cCtaCod, rs!nItemPieza, 0, False)
                        I = I + 1
                        rs.MoveNext
                    Loop
                    
                    Call oPigGrabar.dInsertColocPigGuiaEtapa(lsNumGuia, gPigEtapaRemProyec, lsMovNro, False)
                    
                    oPigGrabar.dCommitTrans
                    
                    Set oPigGrabar = Nothing
                    
                    MsgBox "Proyección de la Guía finalizó satisfactoriamente", vbInformation, "Aviso"
                    Limpiar
                Else
                    MsgBox "No existen Piezas para la Remesa", vbInformation, "Aviso"
                End If
            Else
                MsgBox "No existen Piezas para la Remesa", vbInformation, "Aviso"
            End If
            
            Set rs = Nothing
            Set oPigDatos = Nothing
        
        Case 6, 7, 8, 9  'PIEZA - PARA TIENDAS
                 
            lsNumGuia = oPigFunc.GetNumGuia(cboOrigen.BoundText)
            Set oPigFunc = Nothing
            
            Set oContFunc = New NContFunciones
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set oContFunc = Nothing
            
            Set oPigDatos = New dPigContrato
            Set rs = oPigDatos.dObtienePiezasAdjColocPigGuia(cboOrigen.BoundText, cboDestino.BoundText)
                            
            If rs!nTotItem > 0 Then
            
                Set oPigGrabar = New DPigActualizaBD
                
                oPigGrabar.dBeginTrans
                                    
                Call oPigGrabar.dInsertColocPigGuia(lsNumGuia, cboMotivo.BoundText, cboOrigen.BoundText, cboDestino.BoundText, _
                        2, rs!nTotItem, rs!nPBruto, rs!nPNeto, rs!nTasacion, rs!nValorAdjud, lsMovNro, False)
                  
                Set rs = Nothing
                
                Set rs = oPigDatos.dObtienePiezasAdjColocPigGuiaDet(cboOrigen.BoundText, cboDestino.BoundText)
                I = 1
                
                Do While Not rs.EOF
                
                    Call oPigGrabar.dInsertColocPigGuiaDet(lsNumGuia, I, rs!cCodCta, rs!nItemPieza, 0, False)
                    I = I + 1
                
                    rs.MoveNext
                Loop
                
                Call oPigGrabar.dInsertColocPigGuiaEtapa(lsNumGuia, gPigEtapaRemProyec, lsMovNro, False)
                
                oPigGrabar.dCommitTrans
                
                Call oPigGrabar.dDeletePiezasAdjColocPigGuiaDet(cboOrigen.BoundText, cboDestino.BoundText)
                
                Set oPigGrabar = Nothing
                Set rs = Nothing
                
            Else
                MsgBox "No existen datos para la Remesa", vbInformation, "Aviso"
                Set oPigDatos = Nothing
                Set rs = Nothing
                Exit Sub
                
            End If
                           
            MsgBox "Proyección de la Guía finalizó satisfactoriamente", vbInformation, "Aviso"
            Limpiar

        End Select
        
    End If
    
End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSeleccionar_Click()
Dim oPigDatos As dPigContrato
Dim lnRemate As Long
Dim lnAgencia As Integer
Dim lnMaterial As Integer
Dim lspieza As String
Dim lsEstado As String

    If txtRemate = "" Then
        lnRemate = -1
    Else
        lnRemate = Val(txtRemate)
    End If
    
    If cboMaterial.Text <> "" Then
        lnMaterial = cboMaterial.BoundText
    Else
        lnMaterial = -1
    End If
    
    If cboAgencia.Text <> "" Then
        lnAgencia = cboAgencia.BoundText
    Else
        lnAgencia = -1
    End If
                
    If chkPieza.value = 1 Then
        lspieza = txtCmac & Trim(txtAge) & txtProducto & Trim(txtCuenta) & Trim(txtpieza.Text)
    Else
        lspieza = "@"
    End If

    lsEstado = "2812"

    Set oPigDatos = New dPigContrato
    Call oPigDatos.dInsertaPiezasAdjColocPigGuia(cboOrigen.BoundText, cboDestino.BoundText, lsEstado, lnAgencia, lnMaterial, lnRemate, lspieza)
    Set oPigDatos = Nothing

End Sub

Private Sub Form_Load()

Call CargaCombo(cboMotivo, gColocPigMotivoRem)
Call CargaCombo(cboOrigen, gColocPigUbicacion)
Call CargaCombo(cboDestino, gColocPigUbicacion)
Call CargaCombo(cboClase, gColocPigTipoGuia)
cboOrigen.BoundText = gsCodAge

lbProyectar = False

End Sub

Private Sub CargaCombo(Combo As DataCombo, ByVal psConsCod As String)
Dim oPigFunc As DPigFunciones
Dim rs As Recordset

Set oPigFunc = New DPigFunciones
    
    Set rs = oPigFunc.GetConstante(psConsCod)
    Set Combo.RowSource = rs
    Combo.ListField = "cConsDescripcion"
    Combo.BoundColumn = "nConsValor"

    Set rs = Nothing
    
    Set oPigFunc = Nothing

End Sub

Private Function ValidaProyectar() As Boolean

    ValidaProyectar = True
    
    If cboMotivo.Text = "" Then
        MsgBox "Debe seleccionar un motivo para la Remesa", vbInformation, "Aviso"
        ValidaProyectar = False
        Exit Function
    End If
    
    If cboClase.Text = "" Then
        MsgBox "Debe seleccionar la Clase de Remesa", vbInformation, "Aviso"
        ValidaProyectar = False
        Exit Function
    End If

    If cboOrigen.Text = "" Then
        MsgBox "Debe seleccionar el Origen", vbInformation, "Aviso"
        ValidaProyectar = False
        Exit Function
    End If

    If cboDestino.Text = "" Then
        MsgBox "Debe seleccionar el Destino", vbInformation, "Aviso"
        ValidaProyectar = False
        Exit Function
    End If

End Function

Private Sub Limpiar()

cboMotivo.BoundText = ""
cboClase.BoundText = ""
cboOrigen.BoundText = ""
cboDestino.BoundText = ""
chkRemate.value = 0
chkAgencia.value = 0
chkMaterial.value = 0
chkPieza.value = 0
txtAge = ""
txtCuenta = ""
txtpieza = ""
cboAgencia.Text = ""
cboMaterial.Text = ""

End Sub

Private Sub txtAge_Change()
If Len(Trim(txtAge)) = 2 Then
    If txtCuenta.Enabled Then
        txtCuenta.SetFocus
    End If
End If
End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(txtAge) = 2 Then
    txtCuenta.SetFocus
End If
End Sub

Private Sub txtCuenta_Change()
If Len(Trim(txtCuenta)) = 10 Then
    If txtpieza.Enabled Then
        txtpieza.SetFocus
    End If
End If
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyTab
         If Len(Trim(txtCuenta)) = 0 Then
            txtCuenta.SetFocus
         End If
    Case vbKeyBack
         If Len(Trim(txtCuenta)) = 0 Then
            If txtAge.Enabled Then
                txtAge.SetFocus
            End If
         End If
         Exit Sub
    Case Else
         If Len(Trim(txtCuenta)) = 9 Then
            txtpieza.SetFocus
         End If
    End Select
    If NumerosEnteros(KeyAscii) = 0 Or Len(Trim(txtCuenta)) = 10 Then
       If txtpieza.Enabled Then
            txtpieza.SetFocus
       End If
    End If
End Sub

Private Sub txtPieza_KeyPress(KeyAscii As Integer)
Dim oPigCont As dPigContrato
Dim rs As Recordset
Dim lsCuenta As String

    Select Case KeyAscii
    Case vbKeyTab
         If Len(Trim(txtpieza)) = 0 Then
            txtpieza.SetFocus
         End If
    Case vbKeyBack
         If Len(Trim(txtpieza)) = 0 Then
            If txtCuenta.Enabled Then
                txtCuenta.SetFocus
            End If
         End If
         Exit Sub
    Case Else
        Set oPigCont = New dPigContrato
        If KeyAscii = 13 Then
            If txtpieza <> "" Then
                lsCuenta = txtCmac & txtAge & txtProducto & txtCuenta
                If Len(lsCuenta) <> 18 Then
                    MsgBox "Número de Contrato incompleto", vbInformation, "Aviso"
                    Exit Sub
                Else
                    Set rs = oPigCont.dObtieneDatosPieza(lsCuenta, CInt(txtpieza))
                End If
            Else
                MsgBox "Ingrese el número de Pieza del Contrato", vbInformation, "Aviso"
                Exit Sub
            End If
               
            If Not rs.EOF And Not rs.BOF Then
                lblDescripcion = rs!cDescripcion
                lblMaterial = rs!cConsDescripcion
                lblPesoBruto = rs!nPesoBruto
                lblPesoNeto = rs!nPesoNeto
                lblRemate = rs!nRemate
                lblValorAdj = rs!nValorProceso
            Else
                MsgBox "Pieza no existe", vbInformation, "Aviso"
            End If
            
            Set oPigCont = Nothing
            Set rs = Nothing
        End If
    End Select
End Sub

Private Function ValidaAgBovedaValores() As Boolean
Dim oParam As DPigFunciones
Dim lsAgeBovVal As String

    Set oParam = New DPigFunciones
        lsAgeBovVal = CStr(oParam.GetParamValor(8040))
        If Right(lsAgeBovVal, 2) <> Right(gsCodAge, 2) Then
            ValidaAgBovedaValores = False
        Else
            ValidaAgBovedaValores = True
        End If
    Set oParam = Nothing

End Function
